import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk  # เพิ่มสำหรับ Combobox
from PIL import Image, ImageTk, ImageOps
import requests
import os
# เพิ่ม import สำหรับ excel
import openpyxl
import fitz  # เพิ่มบรรทัดนี้
from openpyxl.utils import get_column_letter
import tkinter.font as tkfont
from functools import partial
import json  # เพิ่มสำหรับอ่านไฟล์ json

API_KEY = 'ZlA0wiVWdf8AvPodV4MCrXIxs3rcIDAx'
OCR_URL = "https://api.iapp.co.th/ocr/v3/receipt/file"

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Invoice Extractor")
        self.root.geometry("1200x800")
        self.root.configure(bg="#e9ecef")

        # --- Modern style ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("TH Niramit AS", 16), padding=6)
        style.configure("TEntry", font=("TH Niramit AS", 16))
        style.configure("TCombobox", font=("TH Niramit AS", 16))
        style.configure("TCheckbutton", font=("TH Niramit AS", 16), background="#f5f5f5")
        style.configure("TLabel", font=("TH Niramit AS", 16), background="#f5f5f5")
        style.map("TButton", background=[("active", "#1ca21c")])

        # --- Main window layout (no scroll) ---
        self.main_frame = tk.Frame(root, background="#e9ecef")
        self.main_frame.pack(fill="both", expand=True)

        # --- Responsive grid config ---
        self.main_frame.grid_columnconfigure(0, weight=0)  # left panel fixed
        self.main_frame.grid_columnconfigure(1, weight=0)
        self.main_frame.grid_columnconfigure(2, weight=1)  # right panel flex

        self.file_path = None
        self.ocr_data = {}

        # Left: Image preview (flexible size, white bg)
        self.preview_width = 350
        self.preview_height = 500
        self.img_canvas = tk.Canvas(
            self.main_frame, bg="#fff", bd=0, relief=tk.RIDGE,
            highlightthickness=1, highlightbackground="#bbb",
            width=self.preview_width, height=self.preview_height
        )
        self.img_canvas.grid(row=0, column=0, rowspan=1, padx=(20,10), pady=(20,5), sticky="nw")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)  # เปลี่ยนจาก 0 เป็น 1 เพื่อให้ canvas ขยาย flex ได้
        self.tk_img = None  # Keep reference

        # --- ปุ่ม Zoom/Reset ใต้กรอบภาพ ---
        zoom_frame = tk.Frame(self.main_frame, bg="#e9ecef")
        zoom_frame.grid(row=2, column=0, padx=(20,10), pady=(0,0), sticky="nw")
        ttk.Button(zoom_frame, text="Zoom In", command=self.zoom_in, width=10).pack(side="left", padx=(0,6))
        ttk.Button(zoom_frame, text="Zoom Out", command=self.zoom_out, width=10).pack(side="left", padx=(0,6))
        ttk.Button(zoom_frame, text="Reset", command=self.reset_zoom, width=10).pack(side="left")

        # --- ปุ่มอยู่ใต้กรอบภาพ ---
        btns_frame_left = tk.Frame(self.main_frame, bg="#e9ecef")
        btns_frame_left.grid(row=1, column=0, padx=(20,10), pady=(0,10), sticky="new")
        button_width = 22  # ปรับขนาดความยาวปุ่มที่นี่
        ttk.Button(btns_frame_left, text="Upload File", command=self.select_file, width=button_width).pack(fill="x", pady=(0,6))
        ttk.Button(btns_frame_left, text="ประมวลผล OCR", command=self.process_ocr, width=button_width).pack(fill="x", pady=(0,6))
        ttk.Button(btns_frame_left, text="Export เป็น Excel", command=self.export_excel, width=button_width).pack(fill="x")

        # --- สำหรับ Zoom ---
        self.original_img = None
        self.current_img = None
        self.zoom_ratio = 1.0

        # --- โหลด config (fields และ product_codes) จาก JSON เดียวกัน ---
        config = self.load_config()
        self.fields = config.get("fields", [])
        self.product_codes = config.get("product_codes", [])
        self.items_headers = config.get("product_headers", ["#", "รหัสสินค้า", "ชื่อรายการ", "จำนวน", "ราคาต่อชิ้น", "จำนวนรวม"])
        self.supplier_names = config.get("supplierName", [])  # <-- เพิ่มบรรทัดนี้
        self.entries = {}
        self.field_vars = {}

        # Right: Fields (all possible keys from processed)เ
        self.entries = {}
        self.field_vars = {}  # เพิ่ม dict สำหรับเก็บตัวแปร checkbox

        doc_frame = tk.LabelFrame(
            self.main_frame,
            text="Document Info",
            font=("TH Niramit AS", 16, "bold"),
            bg="#f5f5f5", bd=2, relief=tk.GROOVE, labelanchor="nw", padx=12, pady=10
        )
        doc_frame.grid(row=0, column=1, columnspan=2, padx=(10,20), pady=(20,10), sticky="new")
        doc_frame.grid_columnconfigure(2, weight=1)  # ปรับให้ช่องผลลัพธ์ขยาย

        for idx, (label, key) in enumerate(self.fields):
            var = tk.BooleanVar(value=True)
            self.field_vars[key] = var
            cb = tk.Checkbutton(
                doc_frame, variable=var, bg="#f5f5f5", activebackground="#f5f5f5",
                selectcolor="#f5f5f5", indicatoron=True, width=2, padx=8, pady=8,
                font=("TH Niramit AS", 16)
            )
            cb.grid(row=idx, column=0, sticky="w", padx=(0,10), pady=8)
            ttk.Label(doc_frame, text=label, anchor="w", width=18).grid(
                row=idx, column=1, sticky="w", padx=(0,10), pady=8
            )
            # ช่องผลลัพธ์อยู่ติดกับ label
            if key == "supplierName" and self.supplier_names:
                entry = ttk.Combobox(doc_frame, values=self.supplier_names, width=44, font=("TH Niramit AS", 16))
            else:
                entry = ttk.Entry(doc_frame, width=46, font=("TH Niramit AS", 16))
            entry.grid(row=idx, column=2, padx=(0,10), pady=8, sticky="ew")
            self.entries[key] = entry

        # สำหรับซ่อน/แสดงตารางสินค้า
        self.items_table_visible = True

        # Items section (table-like)
        table_frame = tk.LabelFrame(self.main_frame, text="Product List", font=("Segoe UI", 13, "bold"), bg="#f5f5f5", bd=2, relief=tk.GROOVE, labelanchor="nw", padx=10, pady=8)
        table_frame.grid(row=1, column=1, columnspan=2, padx=(10,20), pady=(0,10), sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)

        # --- Canvas + Scrollbar for items table ---
        self.items_table_canvas = tk.Canvas(table_frame, bg="#f5f5f5", highlightthickness=0, height=220, width=700)
        self.items_table_canvas.grid(row=0, column=0, sticky="nsew")
        self.items_table_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.items_table_canvas.xview)
        self.items_table_scroll_x.grid(row=1, column=0, sticky="ew")
        self.items_table_canvas.configure(xscrollcommand=self.items_table_scroll_x.set)
        self.items_table_frame = tk.Frame(self.items_table_canvas, bg="#f5f5f5")
        self.items_table_window = self.items_table_canvas.create_window((0, 0), window=self.items_table_frame, anchor="nw")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # ปุ่ม Add row และ Add column ใต้ตาราง
        btns_frame = tk.Frame(table_frame, bg="#f5f5f5")
        btns_frame.grid(row=2, column=0, sticky="e", pady=(4,0))
        ttk.Button(btns_frame, text="+ Add row", command=self.add_item_row).pack(side="right", padx=(0,8))
        ttk.Button(btns_frame, text="+ Add column", command=self.add_item_column).pack(side="right", padx=(0,8))

        # ปุ่ม Submit ใต้ Add row
        ttk.Button(table_frame, text="Submit", command=self.submit).grid(row=3, column=0, sticky="e", pady=(10,2), padx=(0,8))

        # สำหรับเก็บข้อมูลตารางสินค้า
        self.items_data = []  # list of list
        self.items_editors = []  # list of Entry widgets

        # Draw initial empty preview
        self.show_empty_preview()
        self.show_items_table([])

    def show_empty_preview(self):
        self.img_canvas.delete("all")
        self.img_canvas.config(width=self.preview_width, height=self.preview_height)
        self.img_canvas.create_rectangle(0, 0, self.preview_width, self.preview_height, fill="white", outline="")

    def select_file(self):
        filetypes = [
            ("PDF files", "*.pdf"),
            ("Image files", "*.png;*.jpg;*.jpeg;*.bmp"),
            ("All supported", "*.png;*.jpg;*.jpeg;*.bmp;*.pdf"),
            ("All files", "*.*")
        ]
        path = filedialog.askopenfilename(title="เลือกไฟล์ภาพหรือ PDF", filetypes=filetypes)
        if path:
            self.file_path = path
            ext = os.path.splitext(path)[1].lower()
            if ext == ".pdf":
                try:
                    img = self.pdf_page_to_image(path, page_num=0)
                    img = ImageOps.exif_transpose(img)
                    self.original_img = img.copy()
                    self.zoom_ratio = 1.0
                    self.display_image(img)
                except Exception as e:
                    messagebox.showerror("Error", f"PDF preview failed: {e}")
                    self.show_empty_preview()
            else:
                try:
                    img = Image.open(path)
                    img = ImageOps.exif_transpose(img)
                    self.original_img = img.copy()
                    self.zoom_ratio = 1.0
                    self.display_image(img)
                except Exception as e:
                    messagebox.showerror("Error", f"Open image failed: {e}")
                    self.show_empty_preview()
        else:
            self.show_empty_preview()

    def display_image(self, img):
        # ปรับขนาด canvas ให้เท่ากับขนาดภาพ
        w, h = img.size
        self.img_canvas.config(width=w, height=h)
        self.img_canvas.delete("all")
        self.tk_img = ImageTk.PhotoImage(img)
        self.img_canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        self.current_img = img
        # อัปเดตขนาด preview_width/preview_height ด้วย (ถ้าต้องการให้ zoom ใช้ขนาดนี้)
        self.preview_width = w
        self.preview_height = h

    def zoom_in(self):
        if self.current_img is not None:
            self.zoom_ratio *= 1.2
            self._zoom_display()

    def zoom_out(self):
        if self.current_img is not None:
            self.zoom_ratio /= 1.2
            self._zoom_display()

    def reset_zoom(self):
        if self.original_img is not None:
            self.zoom_ratio = 1.0
            self.display_image(self.original_img)

    def _zoom_display(self):
        if self.original_img is not None:
            w, h = self.original_img.size
            new_size = (max(1, int(w * self.zoom_ratio)), max(1, int(h * self.zoom_ratio)))
            img = self.original_img.resize(new_size, Image.LANCZOS)
            self.display_image(img)

    def toggle_items_table(self):
        self.items_table_visible = not self.items_table_visible
        if self.items_table_visible:
            self.items_table_canvas.grid()
            self.items_table_scroll_x.grid()
            self.add_row_btn.grid()
            self.submit_btn.grid()
            self.hide_btn.config(text="Hide ▲")
        else:
            self.items_table_canvas.grid_remove()
            self.items_table_scroll_x.grid_remove()
            self.add_row_btn.grid_remove()
            self.submit_btn.grid_remove()
            self.hide_btn.config(text="Show ▼")

    def show_items_table(self, items):
        # Clear previous table
        for widget in self.items_table_frame.winfo_children():
            widget.destroy()
        self.items_editors = []
        headers = self.items_headers
        col_weights = [1] + [2]*(len(headers)-1)
        # --- วาดหัวตารางแบบแก้ไขชื่อคอลัมน์ได้ ---
        self.header_editors = []
        for col, (h, w) in enumerate(zip(headers, col_weights)):
            header_entry = tk.Entry(
                self.items_table_frame, borderwidth=1, relief="solid",
                font=("Arial", 11, "bold"), justify="center", width=15
            )
            header_entry.insert(0, h)
            header_entry.grid(row=0, column=col, sticky="nsew", padx=0, pady=0, ipady=3)
            # เมื่อแก้ไขชื่อคอลัมน์แล้วอัปเดต self.items_headers
            def save_header(event, idx=col, entry=header_entry):
                self.items_headers[idx] = entry.get()
            header_entry.bind("<FocusOut>", save_header)
            header_entry.bind("<Return>", save_header)
            self.header_editors.append(header_entry)
            self.items_table_frame.grid_columnconfigure(col, weight=w, minsize=40)
        tk.Label(self.items_table_frame, text="", bg="#f5f5f5").grid(row=0, column=len(headers), sticky="nsew")
        # เตรียมข้อมูล
        # ถ้า items เป็น list of dict (จาก OCR) ให้แปลงเป็น list of list
        if items and isinstance(items[0], dict):
            self.items_data = []
            for row, item in enumerate(items, start=1):
                row_data = [
                    str(item.get('itemNo', row)),
                    item.get('itemCode', ''), 
                    item.get('itemName', ''),
                    item.get('itemUnit', ''),
                    item.get('itemUnitCost', ''),
                    item.get('itemTotalCost', '')
                ]
                while len(row_data) < len(headers):
                    row_data.append("")
                self.items_data.append(row_data)
        # ถ้า items เป็น list of list (จากการเพิ่มแถว/คอลัมน์/แก้ไข)
        elif items and isinstance(items[0], list):
            # ป้องกันกรณีคอลัมน์เพิ่ม/ลด
            self.items_data = [row[:len(headers)] + [""]*(len(headers)-len(row)) for row in items]
        # ถ้าไม่มีข้อมูลเลย ให้แสดงแถวเปล่า 1 แถว
        if not self.items_data:
            self.items_data = [[""] * len(headers)]
        # วาดตาราง
        for row_idx, row_data in enumerate(self.items_data, start=1):
            row_editors = []
            for col_idx, value in enumerate(row_data):
                # --- ถ้าเป็นคอลัมน์รหัสสินค้า ให้ใช้ Combobox ---
                if self.items_headers[col_idx] == "รหัสสินค้า":
                    cb = ttk.Combobox(
                    self.items_table_frame, values=self.product_codes,
                    font=("Arial", 11), width=13
                    )
                    cb.set(str(value) if value is not None else "")
                    cb.grid(row=row_idx, column=col_idx, sticky="nsew", padx=0, pady=0, ipady=3)
                    row_editors.append(cb)
                else:
                    e = tk.Entry(
                        self.items_table_frame, borderwidth=1, relief="solid",
                        font=("Arial", 11), justify="left", width=15
                    )
                    e.grid(row=row_idx, column=col_idx, sticky="nsew", padx=0, pady=0, ipady=3)
                    e.insert(0, value)
                    row_editors.append(e)
            # ปุ่มลบแถว (ถังขยะ)
            del_btn = tk.Button(self.items_table_frame, text="🗑", fg="#c00", relief="flat", command=partial(self.delete_item_row, row_idx-1), cursor="hand2")
            del_btn.grid(row=row_idx, column=len(row_data), sticky="nsew", padx=(2,0))
            row_editors.append(del_btn)
            self.items_editors.append(row_editors)
        # --- ปรับ scrollregion และขนาด frame ---
        self.items_table_frame.update_idletasks()
        self.items_table_canvas.config(scrollregion=self.items_table_canvas.bbox("all"))
        # กำหนดความกว้าง canvas คงที่ ไม่ขยายตามตาราง
        # (ถ้าต้องการปรับอัตโนมัติตามขนาดหน้าต่างหลัก ให้ใช้ self.items_table_canvas.config(width=self.items_table_canvas.winfo_width()))
        # ไม่ต้องปรับ width ของ canvas ตาม frame_width

    def delete_item_row(self, idx):
        if 0 <= idx < len(self.items_data):
            del self.items_data[idx]
            self.show_items_table(self.items_data)

    def add_item_row(self):
        # เพิ่มแถวใหม่ (ค่าว่าง)
        if not self.items_data:
            self.items_data = [[""] * len(self.items_headers)]
        else:
            self.items_data.append([""] * len(self.items_headers))
        self.show_items_table(self.items_data)

    def add_item_column(self):
        # เพิ่มคอลัมน์ใหม่ (ชื่อคอลัมน์อัตโนมัติ)
        new_col_name = f"คอลัมน์{len(self.items_headers)}"
        self.items_headers.append(new_col_name)
        for row in self.items_data:
            row.append("")
        self.show_items_table(self.items_data)

    def submit(self):
        messagebox.showinfo("Submit", "Submit clicked!\n(คุณสามารถนำโค้ดนี้ไปต่อยอดบันทึก/ส่งข้อมูลได้)")

    def process_ocr(self):
        if not self.file_path:
            messagebox.showwarning("No file", "กรุณาเลือกไฟล์ภาพก่อน")
            return
        try:
            # ถ้าเป็น PDF ให้แปลงเป็นรูปภาพก่อน
            ext = os.path.splitext(self.file_path)[1].lower()
            if ext == '.pdf':
                # แปลง PDF หน้าแรกเป็นรูปภาพ
                image = self.pdf_page_to_image(self.file_path)
                # บันทึกเป็นไฟล์ชั่วคราว
                temp_path = os.path.join(os.path.dirname(self.file_path), "temp_ocr.png")
                image.save(temp_path, "PNG")
                file_to_send = temp_path
            else:
                file_to_send = self.file_path

            # ส่งไฟล์ไปยัง API
            with open(file_to_send, "rb") as f:
                files = {'file': (os.path.basename(file_to_send), f, 'image/png')}
                headers = {'apikey': API_KEY}
                data = {'return_image': 'false', 'return_ocr': 'false'}
                resp = requests.post(OCR_URL, headers=headers, files=files, data=data)
                resp.raise_for_status()
                result = resp.json()

            # ลบไฟล์ชั่วคราวถ้าเป็น PDF
            if ext == '.pdf' and os.path.exists(temp_path):
                os.remove(temp_path)

            if result.get("message") == "success":
                processed = result.get("processed", {})
                # Fill all fields
                for label, key in self.fields:
                    value = processed.get(key, "")
                    self.entries[key].delete(0, tk.END)
                    self.entries[key].insert(0, str(value) if value is not None else "")
                # Show items as table
                items = processed.get("items", [])
                self.show_items_table(items)
            else:
                messagebox.showerror("OCR Failed", f"OCR API error: {result.get('message')}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_config(self):
        json_path = os.path.join(os.path.dirname(__file__), "config_fields.json")
        if not os.path.exists(json_path):
            # สร้างไฟล์ตัวอย่างถ้ายังไม่มี
            sample = {
                "fields": [
                    ["วันที่เอกสาร", "invoiceDate"],
                    ["ผู้จัดจำหน่าย", "supplierName"],
                    ["คำอธิบาย", "description"]
                ],
                "product_codes": ["P001", "P002", "P003", "P004"]
            }
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(sample, f, ensure_ascii=False, indent=2)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            # รองรับไฟล์เก่า (list of list) เพื่อความเข้ากันได้
            if isinstance(config, list):
                config = {"fields": config, "product_codes": []}
            return config
        except Exception as e:
            messagebox.showerror("Error", f"โหลด config ไม่สำเร็จ: {e}")
            return {"fields": [], "product_codes": []}

    def export_excel(self):
        # เตรียมข้อมูลเฉพาะ field ที่ถูกเลือก
        data = []
        field_names = []
        field_values = []
        for label, key in self.fields:
            if self.field_vars[key].get():
                field_names.append(label)
                field_values.append(self.entries[key].get())
        # เตรียมข้อมูล items จาก Entry (self.items_editors)
        items = []
        for row_editors in self.items_editors:
            row_values = []
            for idx, e in enumerate(row_editors):
                if idx >= len(self.items_headers):
                    # ข้าม widget ที่เกินจำนวน header เช่น ปุ่มลบแถว
                    continue
                # ถ้าเป็น Combobox (รหัสสินค้า) ให้ใช้ get()
                if self.items_headers[idx] == "รหัสสินค้า" and isinstance(e, ttk.Combobox):
                    row_values.append(e.get())
                elif isinstance(e, tk.Entry):
                    row_values.append(e.get())
            if any(row_values):
                items.append(row_values)
        # --- สร้าง header และ value row สำหรับ export ---
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "OCR Data"
        # รวม field headers + items headers
        export_headers = list(field_names)
        export_values = list(field_values)
        if items:
            # ต่อ header ของสินค้าทั้งหมด (ถ้ามีหลายแถว)
            for idx, row in enumerate(items, start=1):
                for col_name in self.items_headers:
                    export_headers.append(f"{col_name}")
            # ต่อค่าของสินค้าทั้งหมด
            for row in items:
                for value in row:
                    export_values.append(value)
        ws.append(export_headers)
        ws.append(export_values)
        # ปรับความกว้างคอลัมน์
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกไฟล์ Excel"
        )
        if file_path:
            try:
                wb.save(file_path)
                messagebox.showinfo("Export สำเร็จ", f"บันทึกไฟล์ Excel ที่\n{file_path}")
            except Exception as e:
                messagebox.showerror("Export ผิดพลาด", str(e))

    def pdf_page_to_image(self, path, page_num=0):
        doc = fitz.open(path)
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        return img

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
