# modules/pdf_editor.py
import base64
import io
import logging
import os
from datetime import datetime
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except:
    PYPDF_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except:
    PIL_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except:
    PYMUPDF_AVAILABLE = False

logger = logging.getLogger(__name__)


def open_pdf(path):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed. Run: pip install pymupdf"
    if not os.path.exists(path):
        return None, f"File not found: {path}"
    try:
        doc   = fitz.open(path)
        pages = _extract_all_pages(doc)
        doc.close()
        return {"total_pages": len(pages), "pages": pages, "file_path": path}, None
    except Exception as e:
        return None, str(e)


def _extract_all_pages(doc):
    pages = []
    for page_num in range(len(doc)):
        page   = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        text_blocks = []
        for block in blocks:
            if block.get("type") == 0:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "").strip()
                        if not text:
                            continue
                        bbox = span.get("bbox", [0, 0, 0, 0])
                        text_blocks.append({
                            "id":    f"p{page_num}_b{len(text_blocks)}",
                            "text":  text,
                            "x":     bbox[0],
                            "y":     bbox[1],
                            "x1":    bbox[2],
                            "y1":    bbox[3],
                            "font":  span.get("font", "helv"),
                            "size":  round(span.get("size", 12), 1),
                            "color": _int_to_hex(span.get("color", 0)),
                            "flags": span.get("flags", 0),
                        })
        pages.append({"text_blocks": text_blocks})
    return pages


def _int_to_hex(color_int):
    try:
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8)  & 0xFF
        b =  color_int        & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except:
        return "#000000"


def render_page(path, page_num, zoom=2.0):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed"
    try:
        doc  = fitz.open(path)
        page = doc[page_num]
        mat  = fitz.Matrix(zoom, zoom)
        pix  = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        doc.close()
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        return {
            "image":  f"data:image/png;base64,{b64}",
            "width":  pix.width,
            "height": pix.height,
            "zoom":   zoom,
        }, None
    except Exception as e:
        return None, str(e)


def save_with_edits(path, edits):
    if not PYMUPDF_AVAILABLE:
        return None, "PyMuPDF not installed"
    try:
        doc = fitz.open(path)
        for edit in edits:
            page      = doc[edit["page"]]
            bbox      = fitz.Rect(
                edit["bbox"]["x"],  edit["bbox"]["y"],
                edit["bbox"]["x1"], edit["bbox"]["y1"]
            )
            page.add_redact_annot(bbox)
            page.apply_redactions()

            style     = edit.get("style", {})
            font_name = style.get("font", "helv")
            font_size = float(style.get("size", 12))
            color_hex = style.get("color", "#000000").lstrip("#")
            r = int(color_hex[0:2], 16) / 255
            g = int(color_hex[2:4], 16) / 255
            b = int(color_hex[4:6], 16) / 255

            page.insert_text(
                fitz.Point(bbox.x0, bbox.y1),
                edit["new_text"],
                fontname=font_name,
                fontsize=font_size,
                color=(r, g, b),
            )

        out_dir  = os.path.dirname(path)
        stem     = Path(path).stem
        out_path = os.path.join(out_dir, f"{stem}_edited.pdf")
        doc.save(out_path)
        doc.close()
        return out_path, None
    except Exception as e:
        return None, str(e)


def detect_form_fields(pdf_path):
    if not PYPDF_AVAILABLE:
        return {}
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        return {name: field.get("/FT", "Unknown") for name, field in fields.items()}
    except Exception as e:
        logger.error(f"Form field detection error: {e}")
        return {}


def get_form_field_options(pdf_path, field_name):
    if not PYPDF_AVAILABLE:
        return []
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        field  = fields.get(field_name)
        if field and "/Opt" in field:
            return [str(o) for o in field["/Opt"]]
        return []
    except Exception as e:
        logger.error(f"Get options error: {e}")
        return []


def fill_form(pdf_path, form_data):
    if not PYPDF_AVAILABLE:
        return False
    try:
        from modules import pdf_utils
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        writer.append(pages_from=reader)
        writer.update_page_form_field_values(writer.pages[0], form_data)
        out_path = pdf_utils.ask(
            kind="savefile",
            defaultname="filled_form.pdf",
            title="Save Filled Form As",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if not out_path:
            return False
        with open(out_path, "wb") as f:
            writer.write(f)
        return out_path
    except Exception as e:
        logger.error(f"Fill form error: {e}")
        return False


# Backward-compatible API used by current server routes.
def extract_pdf_text(path):
    data, err = open_pdf(path)
    if err:
        return {"status": "fail", "message": err}
    return {
        "status": "success",
        "total_pages": data.get("total_pages", 0),
        "pages": data.get("pages", []),
        "file_path": data.get("file_path", path),
    }


def render_page_as_image(path, page_num=0):
    try:
        page_num = int(page_num or 0)
    except Exception:
        page_num = 0
    data, err = render_page(path, page_num, zoom=1.0)
    if err:
        return {"status": "fail", "message": err}
    payload = {"status": "success"}
    payload.update(data or {})
    return payload


def save_edited_pdf(path, edits):
    out_path, err = save_with_edits(path, edits or [])
    if err:
        return {"status": "fail", "message": err}
    return {
        "status": "success",
        "message": f"Saved edited PDF: {out_path}",
        "output_path": out_path,
    }
