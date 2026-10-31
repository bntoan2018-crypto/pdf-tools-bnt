from flask import Flask, request, send_file, jsonify, send_from_directory, after_this_request
from functools import wraps
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docxcompose.composer import Composer
import fitz
import tempfile
import traceback
import os
import webbrowser
import threading
import time
from urllib.parse import quote

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
app = Flask(__name__, static_folder=str(STATIC_DIR))
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 100 * 1024 * 1024))
limiter = Limiter(get_remote_address, app=app, default_limits=["30 per hour", "8 per minute"],
                  storage_uri=os.environ.get("RATELIMIT_STORAGE_URI", "memory://"),
                  headers_enabled=True)
TURNSTILE_SITE_KEY = os.environ.get("TURNSTILE_SITE_KEY", "")
TURNSTILE_SECRET_KEY = os.environ.get("TURNSTILE_SECRET_KEY", "")
REQUIRE_CAPTCHA = os.environ.get("REQUIRE_CAPTCHA", "true").lower() == "true"
TEMP_ROOT = Path(os.environ.get("PDF_TOOLS_TEMP", "/tmp/pdf-tools"))
TEMP_ROOT.mkdir(parents=True, exist_ok=True)


def verify_turnstile(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not TURNSTILE_SECRET_KEY:
            if REQUIRE_CAPTCHA:
                return jsonify({"error": "Máy chủ chưa cấu hình CAPTCHA. Vui lòng thử lại sau."}), 503
            return view(*args, **kwargs)
        token = request.form.get("cf-turnstile-response", "")
        if not token:
            return jsonify({"error": "Vui lòng hoàn tất xác minh CAPTCHA."}), 400
        try:
            import urllib.parse, urllib.request, json
            payload = urllib.parse.urlencode({
                "secret": TURNSTILE_SECRET_KEY, "response": token,
                "remoteip": request.headers.get("CF-Connecting-IP", request.remote_addr or "")
            }).encode()
            req = urllib.request.Request("https://challenges.cloudflare.com/turnstile/v0/siteverify",
                data=payload, headers={"Content-Type":"application/x-www-form-urlencoded"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                result = json.loads(resp.read().decode("utf-8"))
            if not result.get("success"):
                return jsonify({"error": "Xác minh CAPTCHA không thành công. Vui lòng thử lại."}), 403
        except Exception:
            return jsonify({"error": "Không thể xác minh CAPTCHA lúc này. Vui lòng thử lại."}), 503
        return view(*args, **kwargs)
    return wrapped


def cleanup_stale_temp_dirs():
    import shutil
    cutoff = time.time() - 3600
    for item in TEMP_ROOT.glob("pdf_tools_*"):
        try:
            if item.is_dir() and item.stat().st_mtime < cutoff:
                shutil.rmtree(item, ignore_errors=True)
        except OSError:
            pass


@app.before_request
def periodic_temp_cleanup():
    if int(time.time()) % 300 < 2:
        cleanup_stale_temp_dirs()

ALLOWED_EXTENSIONS = {".pdf"}


def valid_pdf(file_storage):
    return (
        file_storage is not None
        and bool(file_storage.filename)
        and Path(file_storage.filename).suffix.lower() in ALLOWED_EXTENSIONS
    )


def save_upload(file_storage, folder: Path, fallback="input.pdf") -> Path:
    safe_name = secure_filename(file_storage.filename or fallback) or fallback
    path = folder / safe_name
    file_storage.save(path)
    return path


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")



def _ocr_available():
    import shutil
    return shutil.which("tesseract") is not None


def _cleanup_after_response(temp_dir: Path):
    """Remove uploaded source and generated files after response transmission."""
    @after_this_request
    def cleanup(response):
        import shutil
        # Flask/Werkzeug closes the response iterable after transmission. Defer cleanup
        # briefly so send_file can finish reading the generated artifact.
        response.call_on_close(lambda: shutil.rmtree(temp_dir, ignore_errors=True))
        return response


@app.get("/api/config")
def public_config():
    return jsonify({"turnstile_site_key": TURNSTILE_SITE_KEY})


@app.get("/api/health")
def health():
    try:
        import pdf2docx
        converter_version = getattr(pdf2docx, "__version__", "installed")
    except Exception:
        converter_version = "unknown"

    try:
        with fitz.open(stream=b"%PDF-1.4\n%%EOF", filetype="pdf"):
            pymupdf_ok = True
    except Exception:
        pymupdf_ok = True  # library import itself is the meaningful health check

    return jsonify({
        "ok": True,
        "engine": "pdf2docx + PyMuPDF + pypdf",
        "local_only": False,
        "ocr_available": _ocr_available(),
        "captcha_configured": bool(TURNSTILE_SITE_KEY and TURNSTILE_SECRET_KEY),
        "note": "OCR availability depends on Tesseract and Vietnamese/English language data installed on the host.",
        "version": converter_version,
    })



def page_has_unmapped_symbols(page) -> bool:
    """Detect common Symbol/Wingdings private-use glyphs that Word converters mis-map."""
    text = page.get_text("text") or ""
    private_use = sum(0xE000 <= ord(ch) <= 0xF8FF for ch in text)
    replacement = text.count("\ufffd")
    return private_use >= 2 or replacement >= 2


def page_needs_visual_fallback(page) -> bool:
    """Conservatively preserve complex designed pages as images instead of corrupting layout.

    pdf2docx often scatters text on slides, brochures and training handouts that use
    many vector shapes, layered text, logos or diagrams. Those pages are more faithful
    as page images; ordinary text pages can still be converted to editable Word.
    """
    if page_has_unmapped_symbols(page):
        return True
    drawings = len(page.get_drawings())
    images = len(page.get_images(full=True))
    blocks = len(page.get_text("blocks") or [])
    # A high vector-object count is a strong signal of a designed page/diagram.
    if drawings >= 35:
        return True
    # Pages combining a logo/photo with vector artwork often break in editable conversion.
    if images >= 1 and drawings >= 5:
        return True
    # Dense layouts with many independent text blocks are prone to overlap.
    if blocks >= 32:
        return True
    return False


def make_page_image_docx(page, output_path: Path, image_path: Path):
    """Create one-page DOCX for a PDF page whose symbols cannot be mapped reliably."""
    rect = page.rect
    doc = Document()
    section = doc.sections[0]
    section.page_width = Pt(rect.width)
    section.page_height = Pt(rect.height)
    section.top_margin = Pt(0)
    section.bottom_margin = Pt(0)
    section.left_margin = Pt(0)
    section.right_margin = Pt(0)
    section.header_distance = Pt(0)
    section.footer_distance = Pt(0)
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.add_run().add_picture(str(image_path), width=Inches(rect.width / 72), height=Inches(rect.height / 72))
    doc.save(str(output_path))


def inspect_pdf(pdf_path: Path):
    """Return simple diagnostics useful for explaining likely conversion limitations."""
    info = {"pages": 0, "text_chars": 0, "image_count": 0, "has_text": False,
            "replacement_chars": 0, "encrypted": False}
    with fitz.open(str(pdf_path)) as pdf:
        info["pages"] = len(pdf)
        info["encrypted"] = bool(pdf.needs_pass)
        for page in pdf:
            text = page.get_text("text") or ""
            info["text_chars"] += len(text)
            info["replacement_chars"] += text.count("\ufffd")
            info["image_count"] += len(page.get_images(full=True))
    info["has_text"] = info["text_chars"] >= max(20, info["pages"] * 10)
    return info



@app.post("/api/convert")
@verify_turnstile
def convert_pdf():
    """Editable conversion. Text, tables and images are reconstructed where possible."""
    uploaded = request.files.get("file")
    if not valid_pdf(uploaded):
        return jsonify({"error": "Vui lòng chọn một file PDF hợp lệ."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    src = save_upload(uploaded, temp_dir)
    out = temp_dir / f"{src.stem}.docx"

    try:
        with fitz.open(str(src)) as pdf:
            if pdf.needs_pass:
                return jsonify({"error": "PDF được bảo vệ bằng mật khẩu. Vui lòng mở khóa trước khi chuyển đổi."}), 400
            if len(pdf) == 0:
                return jsonify({"error": "PDF không có trang."}), 400

        diagnostics = inspect_pdf(src)
        converter = Converter(str(src))
        try:
            converter.convert(str(out), start=0, end=None)
        finally:
            converter.close()

        if not out.exists() or out.stat().st_size == 0:
            raise RuntimeError("Bộ chuyển đổi không tạo được file Word.")

        warning = None
        if not diagnostics["has_text"]:
            warning = (
                "PDF có ít hoặc không có lớp văn bản. File Word có thể thiếu chữ; "
                "hãy dùng OCR nếu đã cài Tesseract, hoặc dùng chế độ giữ nguyên hình ảnh."
            )
        elif diagnostics["replacement_chars"]:
            warning = (
                "PDF có ký tự bị ánh xạ lỗi trong lớp văn bản. Một số ký hiệu đặc biệt "
                "hoặc công thức có thể cần OCR/đối chiếu; chỉ đổi font thường không đủ."
            )

        response = send_file(
            str(out),
            as_attachment=True,
            download_name=out.name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if warning:
            response.headers["X-PDF-Tools-Warning"] = quote(warning, safe="")
        _cleanup_after_response(temp_dir)
        return response
    except Exception as exc:
        app.logger.exception("Editable PDF conversion failed")
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({
            "error": "Không chuyển đổi được PDF sang Word.",
            "detail": str(exc)[:500],
            "suggestion": "Thử mở khóa PDF, kiểm tra file không hỏng, hoặc dùng chế độ giữ nguyên hình ảnh."
        }), 500



@app.post("/api/convert_hybrid")
@verify_turnstile
def convert_hybrid():
    """Hybrid conversion: editable on simple pages, visual fallback on complex designed pages."""
    uploaded = request.files.get("file")
    if not valid_pdf(uploaded):
        return jsonify({"error": "Vui lòng chọn file PDF hợp lệ."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    src = save_upload(uploaded, temp_dir)
    out = temp_dir / f"{src.stem}_hybrid.docx"
    try:
        with fitz.open(str(src)) as pdf:
            if pdf.needs_pass:
                return jsonify({"error": "PDF được bảo vệ bằng mật khẩu. Vui lòng mở khóa trước."}), 400
            if len(pdf) == 0:
                return jsonify({"error": "PDF không có trang."}), 400

            page_docs = []
            image_fallback_pages = []
            for index in range(len(pdf)):
                page_pdf = temp_dir / f"source_page_{index + 1}.pdf"
                page_writer = PdfWriter()
                page_writer.add_page(PdfReader(str(src)).pages[index])
                with page_pdf.open("wb") as stream:
                    page_writer.write(stream)

                page_docx = temp_dir / f"converted_page_{index + 1}.docx"
                with fitz.open(str(page_pdf)) as single_page:
                    page = single_page[0]
                    use_image_fallback = page_needs_visual_fallback(page)
                    if use_image_fallback:
                        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                        image_path = temp_dir / f"formula_page_{index + 1}.png"
                        pix.save(str(image_path))
                        make_page_image_docx(page, page_docx, image_path)
                        image_fallback_pages.append(index + 1)
                    else:
                        converter = Converter(str(page_pdf))
                        try:
                            converter.convert(str(page_docx), start=0, end=None)
                        finally:
                            converter.close()
                if not page_docx.exists() or page_docx.stat().st_size == 0:
                    raise RuntimeError(f"Không tạo được Word cho trang {index + 1}.")
                page_docs.append(page_docx)

        if not page_docs:
            raise RuntimeError("Không có trang nào để chuyển đổi.")
        master = Document(str(page_docs[0]))
        composer = Composer(master)
        for page_docx in page_docs[1:]:
            composer.append(Document(str(page_docx)))
        composer.save(str(out))
        if not out.exists() or out.stat().st_size == 0:
            raise RuntimeError("Không tạo được file Word kết quả.")

        response = send_file(
            str(out), as_attachment=True, download_name=out.name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if image_fallback_pages:
            pages = ", ".join(map(str, image_fallback_pages))
            response.headers["X-PDF-Tools-Warning"] = quote((
                f"Các trang {pages} có bố cục phức tạp, nhiều hình/vector hoặc ký hiệu khó giải mã nên được giữ nguyên dưới dạng ảnh để tránh chữ chồng lấn và sai công thức. "
                "Chỉ các trang còn lại mới là nội dung Word có thể chỉnh sửa. Nếu hầu hết trang đều phức tạp, file kết quả sẽ ưu tiên giữ bố cục và có thể ít nội dung chỉnh sửa được."
            ), safe="")
        else:
            response.headers["X-PDF-Tools-Warning"] = quote((
                "Không phát hiện trang cần giữ nguyên ảnh theo heuristic. PDF có bố cục phức tạp vẫn có thể bị lệch; hãy kiểm tra kết quả."
            ), safe="")
        return response
    except Exception as exc:
        app.logger.exception("Hybrid PDF conversion failed")
        return jsonify({
            "error": "Không thể chuyển đổi PDF ở chế độ kết hợp.",
            "detail": str(exc)[:500],
            "suggestion": "Kiểm tra thư viện đã cài đủ trong requirements.txt; thử Python 3.11 hoặc 3.12 nếu gặp lỗi tương thích."
        }), 500




@app.post("/api/convert_hybrid_editable")
@verify_turnstile
def convert_hybrid_editable():
    """Editable-first hybrid: preserve text editability; use page images only when PDF text contains unmapped special symbols."""
    uploaded = request.files.get("file")
    if not valid_pdf(uploaded):
        return jsonify({"error": "Vui lòng chọn file PDF hợp lệ."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    src = save_upload(uploaded, temp_dir)
    out = temp_dir / f"{src.stem}_editable_first.docx"
    try:
        with fitz.open(str(src)) as pdf:
            if pdf.needs_pass:
                return jsonify({"error": "PDF được bảo vệ bằng mật khẩu. Vui lòng mở khóa trước."}), 400
            if len(pdf) == 0:
                return jsonify({"error": "PDF không có trang."}), 400

            page_docs = []
            image_fallback_pages = []
            for index in range(len(pdf)):
                page_pdf = temp_dir / f"source_page_{index + 1}.pdf"
                page_writer = PdfWriter()
                page_writer.add_page(PdfReader(str(src)).pages[index])
                with page_pdf.open("wb") as stream:
                    page_writer.write(stream)

                page_docx = temp_dir / f"converted_page_{index + 1}.docx"
                with fitz.open(str(page_pdf)) as single_page:
                    page = single_page[0]
                    use_image_fallback = page_has_unmapped_symbols(page)
                    if use_image_fallback:
                        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                        image_path = temp_dir / f"formula_page_{index + 1}.png"
                        pix.save(str(image_path))
                        make_page_image_docx(page, page_docx, image_path)
                        image_fallback_pages.append(index + 1)
                    else:
                        converter = Converter(str(page_pdf))
                        try:
                            converter.convert(str(page_docx), start=0, end=None)
                        finally:
                            converter.close()
                if not page_docx.exists() or page_docx.stat().st_size == 0:
                    raise RuntimeError(f"Không tạo được Word cho trang {index + 1}.")
                page_docs.append(page_docx)

        if not page_docs:
            raise RuntimeError("Không có trang nào để chuyển đổi.")
        master = Document(str(page_docs[0]))
        composer = Composer(master)
        for page_docx in page_docs[1:]:
            composer.append(Document(str(page_docx)))
        composer.save(str(out))
        if not out.exists() or out.stat().st_size == 0:
            raise RuntimeError("Không tạo được file Word kết quả.")

        response = send_file(
            str(out), as_attachment=True, download_name=out.name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if image_fallback_pages:
            pages = ", ".join(map(str, image_fallback_pages))
            response.headers["X-PDF-Tools-Warning"] = quote((
                f"Các trang {pages} có ký hiệu font đặc biệt không giải mã tin cậy được giữ dạng ảnh để tránh làm sai công thức. Các trang còn lại được chuyển thành nội dung Word có thể chỉnh sửa; bố cục nhiều cột, sơ đồ và bảng có thể cần căn chỉnh lại."
            ), safe="")
        else:
            response.headers["X-PDF-Tools-Warning"] = quote((
                "Chế độ ưu tiên chỉnh sửa đã chuyển văn bản thành nội dung Word. Bố cục, bảng, biểu đồ và hình có thể bị lệch; kiểm tra lại sau chuyển đổi."
            ), safe="")
        return response
    except Exception as exc:
        app.logger.exception("Hybrid PDF conversion failed")
        return jsonify({
            "error": "Không thể chuyển đổi PDF ở chế độ kết hợp.",
            "detail": str(exc)[:500],
            "suggestion": "Kiểm tra thư viện đã cài đủ trong requirements.txt; thử Python 3.11 hoặc 3.12 nếu gặp lỗi tương thích."
        }), 500



@app.post("/api/convert_visual")
@verify_turnstile
def convert_visual():
    """Create a Word file with each PDF page as a high-resolution image."""
    uploaded = request.files.get("file")
    if not valid_pdf(uploaded):
        return jsonify({"error": "Vui lòng chọn file PDF hợp lệ."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    src = save_upload(uploaded, temp_dir)
    out = temp_dir / f"{src.stem}_giu_nguyen_hinh_anh.docx"

    try:
        docx = Document()
        with fitz.open(str(src)) as pdf:
            if pdf.needs_pass:
                return jsonify({"error": "PDF có mật khẩu. Vui lòng mở khóa trước."}), 400
            if not len(pdf):
                return jsonify({"error": "PDF không có trang."}), 400

            for index, page in enumerate(pdf):
                rect = page.rect
                section = docx.sections[0]
                section.page_width = Pt(rect.width)
                section.page_height = Pt(rect.height)
                section.top_margin = Pt(0)
                section.bottom_margin = Pt(0)
                section.left_margin = Pt(0)
                section.right_margin = Pt(0)
                section.header_distance = Pt(0)
                section.footer_distance = Pt(0)

                # 2.5x rasterization is a practical balance of quality and file size.
                pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
                image_path = temp_dir / f"page_{index + 1}.png"
                pix.save(str(image_path))

                if index:
                    docx.add_page_break()
                paragraph = docx.add_paragraph()
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1
                paragraph.add_run().add_picture(
                    str(image_path),
                    width=Inches(rect.width / 72),
                    height=Inches(rect.height / 72)
                )

        docx.save(str(out))
        response = send_file(
            str(out), as_attachment=True, download_name=out.name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        _cleanup_after_response(temp_dir)
        return response
    except Exception as exc:
        app.logger.exception("Visual conversion failed")
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Không thể tạo Word giữ nguyên hình ảnh.", "detail": str(exc)[:300]}), 500



@app.post("/api/convert_ocr")
@verify_turnstile
def convert_ocr():
    """
    OCR conversion for scanned PDFs.
    Requires Tesseract OCR installed on Windows and Vietnamese/English language packs.
    Text is editable, but complex page layout may not be retained exactly.
    """
    uploaded = request.files.get("file")
    if not valid_pdf(uploaded):
        return jsonify({"error": "Vui lòng chọn file PDF hợp lệ."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    src = save_upload(uploaded, temp_dir)
    out = temp_dir / f"{src.stem}_OCR.docx"

    try:
        docx = Document()
        with fitz.open(str(src)) as pdf:
            if pdf.needs_pass:
                return jsonify({"error": "PDF có mật khẩu. Vui lòng mở khóa trước."}), 400
            if not len(pdf):
                return jsonify({"error": "PDF không có trang."}), 400

            for page_index, page in enumerate(pdf):
                try:
                    text_page = page.get_textpage_ocr(language="vie+eng", dpi=300, full=True)
                    text = page.get_text("text", textpage=text_page)
                except Exception as exc:
                    raise RuntimeError(
                        "Không chạy được OCR. Hãy cài Tesseract OCR và dữ liệu ngôn ngữ vie, eng. "
                        f"Chi tiết: {exc}"
                    ) from exc

                if page_index:
                    docx.add_page_break()
                heading = docx.add_paragraph()
                heading.paragraph_format.space_after = Pt(4)
                heading.add_run(f"Trang {page_index + 1}").bold = True
                for line in text.splitlines():
                    if line.strip():
                        p = docx.add_paragraph(line.strip())
                        p.paragraph_format.space_after = Pt(0)

        docx.save(str(out))
        response = send_file(
            str(out), as_attachment=True, download_name=out.name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        _cleanup_after_response(temp_dir)
        return response
    except Exception as exc:
        app.logger.exception("OCR conversion failed")
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "OCR không hoàn tất.", "detail": str(exc)[:500]}), 500



@app.post("/api/merge")
@verify_turnstile
def merge_pdf():
    files = request.files.getlist("files")
    if len(files) < 2:
        return jsonify({"error": "Vui lòng chọn ít nhất 2 file PDF."}), 400

    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_tools_request_", dir=TEMP_ROOT))
    writer = PdfWriter()
    try:
        for index, uploaded in enumerate(files):
            if not valid_pdf(uploaded):
                return jsonify({"error": f"File thứ {index + 1} không phải PDF hợp lệ."}), 400
            path = save_upload(uploaded, temp_dir, f"file_{index + 1}.pdf")
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception as exc:
                    raise RuntimeError(f"File {uploaded.filename} được bảo vệ bằng mật khẩu.") from exc
            for page in reader.pages:
                writer.add_page(page)

        out = temp_dir / "PDF_merged.pdf"
        with out.open("wb") as stream:
            writer.write(stream)
        response = send_file(str(out), as_attachment=True, download_name=out.name, mimetype="application/pdf")
        _cleanup_after_response(temp_dir)
        return response
    except Exception as exc:
        app.logger.exception("PDF merge failed")
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Không nối được các file PDF.", "detail": str(exc)[:500]}), 500


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "Tổng dung lượng file vượt quá giới hạn cho phép (100 MB mặc định)."}), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8765"))
    host = "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1"
    if not os.environ.get("RENDER"):
        def open_browser():
            time.sleep(1.2)
            webbrowser.open(f"http://127.0.0.1:{port}/")
        threading.Thread(target=open_browser, daemon=True).start()
    app.run(host=host, port=port, debug=False)
