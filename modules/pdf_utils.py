# modules/pdf_utils.py
import os
import re
from fpdf import FPDF
from pypdf import PdfReader, PdfWriter
import queue as q

dialog_request_queue = q.Queue()
dialog_result_queue  = q.Queue()


def run_dialog_listener():
    """MUST be called on the main thread in server.py."""
    import tkinter as tk
    from tkinter import filedialog
    print("Dialog listener ready on main thread")
    while True:
        try:
            req = dialog_request_queue.get(timeout=1)
        except:
            continue

        kind        = req.get("kind")
        defaultname = req.get("defaultname", "file")
        title       = req.get("title", "Save As")
        filetypes   = req.get("filetypes", [("All Files", "*.*")])

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        result = None
        try:
            if kind == "savefile":
                result = filedialog.asksaveasfilename(
                    title=title,
                    initialfile=defaultname,
                    defaultextension=filetypes[0][1].replace("*", ""),
                    filetypes=filetypes
                ) or None
            elif kind == "choosefolder":
                result = filedialog.askdirectory(title=title) or None
            elif kind == "openmultiple":
                files = filedialog.askopenfilenames(
                    title=title, filetypes=filetypes
                )
                result = list(files) if files else []
            elif kind == "splitfull":
                total_pages = req.get("totalpages", 0)
                options = _show_split_dialog(root, total_pages)
                if options:
                    root2 = tk.Tk()
                    root2.withdraw()
                    root2.attributes("-topmost", True)
                    folder = filedialog.askdirectory(
                        title="Select Folder to Save Split Files"
                    )
                    root2.destroy()
                    result = {"options": options, "folder": folder} if folder else None
                else:
                    result = None
            else:
                result = None
        except Exception as e:
            print(f"Dialog error: {e}")
        finally:
            try:
                root.destroy()
            except:
                pass
        dialog_result_queue.put(result)


def _show_split_dialog(parent, total_pages):
    import tkinter as tk
    result = {"value": None}

    win = tk.Toplevel(parent)
    win.title("Split PDF")
    win.geometry("580x500")
    win.attributes("-topmost", True)
    win.configure(bg="#f5f7fa")
    win.grab_set()

    tk.Label(win, text="PDF SPLITTER", font=("Helvetica", 13, "bold"),
             bg="#007bff", fg="white").pack(fill="x", padx=0, pady=0)
    tk.Label(win, text=f"Total pages: {total_pages}", font=("Helvetica", 10),
             bg="#f5f7fa", fg="#495057").pack(anchor="w", padx=20, pady=(10, 0))

    tk.Label(win, text="Enter page ranges (each range becomes one PDF):",
             font=("Helvetica", 10), bg="#f5f7fa", fg="#495057").pack(
        anchor="w", padx=20, pady=(10, 0))

    entry_var = tk.StringVar(value=f"1-{total_pages}")
    entry = tk.Entry(win, textvariable=entry_var, font=("Helvetica", 11),
                     bd=1, relief="solid", bg="#ffffff", fg="#212529")
    entry.pack(fill="x", padx=20, pady=8)

    err_label = tk.Label(win, text="", font=("Helvetica", 9),
                         bg="#f5f7fa", fg="#dc3545")
    err_label.pack(anchor="w", padx=20)

    def on_confirm():
        raw = entry_var.get().strip()
        groups = _parse_range_groups(raw, total_pages)
        if groups is None:
            err_label.config(
                text=f"Invalid range. Use format: 1-25, 26-80 (pages 1 to {total_pages})")
            return
        if not groups:
            err_label.config(text="Please enter at least one range.")
            return
        result["value"] = {"mode": "ranges", "groups": groups}
        win.destroy()

    def on_cancel():
        win.destroy()

    btn_frame = tk.Frame(win, bg="#f5f7fa")
    btn_frame.pack(side="bottom", fill="x", padx=20, pady=12)
    tk.Button(btn_frame, text="Cancel", command=on_cancel,
              font=("Helvetica", 10), bg="#f1f3f5", fg="#495057",
              relief="flat", padx=12, pady=7).pack(side="right", padx=8)
    tk.Button(btn_frame, text="Split PDF", command=on_confirm,
              font=("Helvetica", 10, "bold"), bg="#007bff", fg="white",
              relief="flat", padx=12, pady=7).pack(side="right")

    win.bind("<Return>", lambda e: on_confirm())
    win.bind("<Escape>", lambda e: on_cancel())
    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.wait_window()
    return result["value"]


def _parse_range_groups(raw, total_pages):
    if not raw.strip():
        return []
    groups = []
    try:
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                a, b = int(a.strip()), int(b.strip())
                if a < 1 or b > total_pages or a > b:
                    return None
                groups.append(list(range(a - 1, b)))
            else:
                n = int(part.strip())
                if n < 1 or n > total_pages:
                    return None
                groups.append([n - 1])
    except:
        return None
    return groups


def ask(kind, defaultname="file", title="Save As", filetypes=None, **kwargs):
    if filetypes is None:
        filetypes = [("All Files", "*.*")]
    req = {"kind": kind, "defaultname": defaultname,
           "title": title, "filetypes": filetypes}
    req.update(kwargs)
    dialog_request_queue.put(req)
    try:
        return dialog_result_queue.get(timeout=120)
    except:
        return None


def _safe_line(text):
    return (text.encode("latin-1", errors="replace")
                .decode("latin-1")
                .replace("\x00", ""))


class ReportPDF(FPDF):
    def __init__(self, doc_title):
        super().__init__()
        self.doc_title = doc_title

    def header(self):
        self.set_font("Helvetica", "B", 13)
        self.cell(0, 10, self.doc_title, new_x="LMARGIN", new_y="NEXT", align="C")
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def create_report(text, title="Report", output_path=None):
    if output_path is None:
        safe_name = re.sub(r'[\\/*?:"<>|]', "", title.strip()) or "Report"
        output_path = ask(
            kind="savefile",
            defaultname=f"{safe_name}.pdf",
            title="Save PDF As",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
    if not output_path:
        return None

    pdf = ReportPDF(doc_title=_safe_line(title))
    pdf.set_margins(15, 20, 15)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)
    pdf.set_text_color(0, 0, 0)

    clean_text = _safe_line(text)
    for para in clean_text.split("\n"):
        if para.strip():
            pdf.write(5, para.strip())
        else:
            pdf.ln(2)

    pdf.output(output_path)
    print(f"✅ PDF Created: {output_path}")
    return output_path


def merge_pdfs(input_paths):
    output_path = ask(
        kind="savefile",
        defaultname="merged.pdf",
        title="Save Merged PDF As",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    if not output_path:
        return None

    writer = PdfWriter()
    for path in input_paths:
        if os.path.exists(path):
            writer.append(PdfReader(path))
    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"✅ Merged: {output_path}")
    return output_path


def split_pdf(input_path):
    reader     = PdfReader(input_path)
    total_pages = len(reader.pages)
    response   = ask(kind="splitfull", title="Split PDF", totalpages=total_pages)
    if not response:
        return None

    options       = response["options"]
    output_folder = response["folder"]
    groups        = options["groups"]
    os.makedirs(output_folder, exist_ok=True)

    output_files = []
    for page_indices in groups:
        writer = PdfWriter()
        for i in page_indices:
            writer.add_page(reader.pages[i])
        first = page_indices[0] + 1
        last  = page_indices[-1] + 1
        fname = f"pages{first}-{last}.pdf" if first != last else f"page{first}.pdf"
        out_path = os.path.join(output_folder, fname)
        with open(out_path, "wb") as f:
            writer.write(f)
        output_files.append(out_path)

    print(f"✅ Done: {len(output_files)} files → {output_folder}")
    return output_files


def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)
