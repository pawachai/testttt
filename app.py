import customtkinter as ctk
from tkinter import filedialog, messagebox, Canvas
import os
import platform
import threading
import re
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import pandas as pd
import qrcode
import io

# ==========================================
# บังคับใช้ธีม Light (พื้นขาว ตัวหนังสือดำ)
# ==========================================
ctk.set_appearance_mode("Light") 
ctk.set_default_color_theme("blue")

# ==========================================
# ระบบค้นหาฟอนต์ (.ttf) ภายในเครื่อง (อัปเดตรองรับ Mac)
# ==========================================
def get_system_fonts():
    font_dict = {}
    sys_name = platform.system()
    font_dirs = []
    
    if sys_name == "Windows":
        font_dirs = [os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')]
    elif sys_name == "Darwin": # สำหรับ macOS
        font_dirs = [
            '/Library/Fonts', 
            '/System/Library/Fonts', 
            '/System/Library/Fonts/Supplemental', 
            os.path.expanduser('~/Library/Fonts')
        ]
    
    for font_dir in font_dirs:
        if os.path.exists(font_dir):
            try:
                for file in os.listdir(font_dir):
                    if file.lower().endswith('.ttf'):
                        name = os.path.splitext(file)[0]
                        font_dict[name] = os.path.join(font_dir, file)
            except Exception:
                pass
                
    if not font_dict:
        font_dict["Helvetica"] = None
        
    sorted_fonts = {k: font_dict[k] for k in sorted(font_dict.keys(), key=lambda x: x.lower())}
    return sorted_fonts

SYSTEM_FONTS = get_system_fonts()
FONT_NAMES_LIST = list(SYSTEM_FONTS.keys())

# ค้นหาฟอนต์ไทยแบบไม่สนใจตัวพิมพ์เล็ก/ใหญ่
DEFAULT_THAI_FONT = "Helvetica"
thai_font_candidates = ["tahoma", "leelawui", "thsarabunnew", "thsarabun", "arial", "krungthep", "ayuthaya", "silom"]

for sys_font in FONT_NAMES_LIST:
    if sys_font.lower() in thai_font_candidates:
        DEFAULT_THAI_FONT = sys_font
        break

PAGE_SIZES_MM = {
    "A3": (297.0, 420.0), "A4": (210.0, 297.0),
    "A5": (148.0, 210.0), "A6": (105.0, 148.0),
    "B4": (250.0, 353.0), "B5": (176.0, 250.0),
    "Letter": (215.9, 279.4), "Legal": (215.9, 355.6),
}

def get_image_aspect_ratio(img_path):
    try:
        with Image.open(img_path) as img:
            return img.height / img.width
    except Exception:
        return 1.0

def generate_qr_pil(data):
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=0)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").get_image()


class ImageGroupFrame(ctk.CTkFrame):
    def __init__(self, master, group_index, update_preview_callback, delete_callback, **kwargs):
        super().__init__(master, **kwargs)
        self.group_index = group_index
        self.folder_path = None
        self.image_files = []
        self.update_preview_callback = update_preview_callback
        self.delete_callback = delete_callback

        self.grid_columnconfigure(7, weight=1)

        self.lbl_title = ctk.CTkLabel(self, text=f"📂 กลุ่มที่ {self.group_index} (รูปภาพ)", font=("Helvetica", 14, "bold"))
        self.lbl_title.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.btn_browse = ctk.CTkButton(self, text="เลือกโฟลเดอร์", command=self.browse_folder, width=100)
        self.btn_browse.grid(row=0, column=1, padx=10, pady=5)

        self.lbl_path = ctk.CTkLabel(self, text="ยังไม่ได้เลือก", text_color="gray")
        self.lbl_path.grid(row=0, column=2, columnspan=4, padx=10, pady=5, sticky="w")

        self.btn_delete = ctk.CTkButton(self, text="❌ ลบ", width=60, fg_color="#e74c3c", hover_color="#c0392b", command=lambda: self.delete_callback(self))
        self.btn_delete.grid(row=0, column=7, padx=10, pady=5, sticky="e")

        self.setup_coordinates_ui()

    def setup_coordinates_ui(self):
        self.lbl_x = ctk.CTkLabel(self, text="X(mm):")
        self.lbl_x.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_x = ctk.CTkEntry(self, width=60)
        self.entry_x.insert(0, "10")
        self.entry_x.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_x.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.lbl_y = ctk.CTkLabel(self, text="Y(mm):")
        self.lbl_y.grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_y = ctk.CTkEntry(self, width=60)
        self.entry_y.insert(0, "10")
        self.entry_y.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_y.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        self.lbl_w = ctk.CTkLabel(self, text="กว้าง(mm):")
        self.lbl_w.grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.entry_w = ctk.CTkEntry(self, width=60)
        self.entry_w.insert(0, "40")
        self.entry_w.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_w.grid(row=1, column=5, padx=5, pady=5, sticky="w")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            valid_exts = ('.png', '.jpg', '.jpeg')
            files = [f for f in os.listdir(folder) if f.lower().endswith(valid_exts)]
            
            # ใช้ Natural Sort เพื่อเรียงตัวเลขในชื่อไฟล์ให้ถูกต้อง
            def natural_keys(text):
                return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)]
            
            files.sort(key=natural_keys)
            
            self.image_files = [os.path.join(folder, f) for f in files]
            self.lbl_path.configure(text=f"พบ {len(self.image_files)} รูป", text_color="green")
            self.update_preview_callback()

    def get_config(self):
        try:
            return {
                "type": "image", "items": self.image_files,
                "x_mm": float(self.entry_x.get() or 0), 
                "y_mm": float(self.entry_y.get() or 0), 
                "width_mm": float(self.entry_w.get() or 10)
            }
        except ValueError:
            return None


class ExcelQrGroupFrame(ctk.CTkFrame):
    def __init__(self, master, group_index, update_preview_callback, delete_callback, **kwargs):
        super().__init__(master, **kwargs)
        self.group_index = group_index
        self.df = None
        self.qr_data_list = []
        self.col_mapping = {}
        self.update_preview_callback = update_preview_callback
        self.delete_callback = delete_callback

        self.grid_columnconfigure(7, weight=1)

        self.lbl_title = ctk.CTkLabel(self, text=f"📊 กลุ่มที่ {self.group_index} (QR Code)", font=("Helvetica", 14, "bold"))
        self.lbl_title.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.btn_browse = ctk.CTkButton(self, text="เลือกไฟล์ Excel", command=self.browse_excel, width=100, fg_color="#27ae60", hover_color="#219a52")
        self.btn_browse.grid(row=0, column=1, padx=10, pady=5)

        self.lbl_path = ctk.CTkLabel(self, text="ยังไม่ได้เลือก", text_color="gray")
        self.lbl_path.grid(row=0, column=2, columnspan=4, padx=10, pady=5, sticky="w")
        
        self.btn_delete = ctk.CTkButton(self, text="❌ ลบ", width=60, fg_color="#e74c3c", hover_color="#c0392b", command=lambda: self.delete_callback(self))
        self.btn_delete.grid(row=0, column=7, padx=10, pady=5, sticky="e")

        self.lbl_col = ctk.CTkLabel(self, text="คอลัมน์:")
        self.lbl_col.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.col_var = ctk.StringVar(value="-")
        self.opt_col = ctk.CTkOptionMenu(self, values=["-"], variable=self.col_var, command=self.extract_data, width=70)
        self.opt_col.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.lbl_row = ctk.CTkLabel(self, text="เริ่มแถว:")
        self.lbl_row.grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_start_row = ctk.CTkEntry(self, width=60)
        self.entry_start_row.insert(0, "1")
        self.entry_start_row.bind("<KeyRelease>", lambda e: self.extract_data())
        self.entry_start_row.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        self.lbl_x = ctk.CTkLabel(self, text="X(mm):")
        self.lbl_x.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_x = ctk.CTkEntry(self, width=60)
        self.entry_x.insert(0, "10")
        self.entry_x.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_x.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        self.lbl_y = ctk.CTkLabel(self, text="Y(mm):")
        self.lbl_y.grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.entry_y = ctk.CTkEntry(self, width=60)
        self.entry_y.insert(0, "10")
        self.entry_y.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_y.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        self.lbl_w = ctk.CTkLabel(self, text="กว้าง(mm):")
        self.lbl_w.grid(row=2, column=4, padx=5, pady=5, sticky="e")
        self.entry_w = ctk.CTkEntry(self, width=60)
        self.entry_w.insert(0, "30") 
        self.entry_w.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_w.grid(row=2, column=5, padx=5, pady=5, sticky="w")

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path, header=None)
                col_options = [str(i + 1) for i in range(len(self.df.columns))]
                self.col_mapping = {str(i + 1): i for i in range(len(self.df.columns))}
                self.opt_col.configure(values=col_options)
                self.col_var.set(col_options[0]) 
                self.extract_data()
            except Exception as e:
                messagebox.showerror("ข้อผิดพลาด", f"อ่านไฟล์ Excel ไม่ได้:\n{e}")

    def extract_data(self, *args):
        if self.df is None: return
        try:
            start_row = max(1, int(self.entry_start_row.get() or 1))
            col_idx = self.col_mapping.get(self.col_var.get(), 0)
            raw_data = self.df.iloc[(start_row - 1):, col_idx].astype(str).tolist()
            self.qr_data_list = [d for d in raw_data if d.strip() != '' and d.strip().lower() != 'nan']
            self.lbl_path.configure(text=f"พบข้อมูล {len(self.qr_data_list)} รายการ", text_color="green")
        except Exception:
            self.qr_data_list = []
            self.lbl_path.configure(text="ดึงข้อมูลล้มเหลว", text_color="red")
        self.update_preview_callback()

    def get_config(self):
        try:
            return {
                "type": "qrcode", "items": self.qr_data_list,
                "x_mm": float(self.entry_x.get() or 0), 
                "y_mm": float(self.entry_y.get() or 0), 
                "width_mm": float(self.entry_w.get() or 10)
            }
        except ValueError: return None


class ExcelTextGroupFrame(ctk.CTkFrame):
    def __init__(self, master, group_index, update_preview_callback, delete_callback, **kwargs):
        super().__init__(master, **kwargs)
        self.group_index = group_index
        self.df = None
        self.text_data_list = []
        self.col_mapping = {}
        self.update_preview_callback = update_preview_callback
        self.delete_callback = delete_callback

        self.grid_columnconfigure(7, weight=1)

        self.lbl_title = ctk.CTkLabel(self, text=f"📝 กลุ่มที่ {self.group_index} (ข้อความ)", font=("Helvetica", 14, "bold"))
        self.lbl_title.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.btn_browse = ctk.CTkButton(self, text="เลือกไฟล์ Excel", command=self.browse_excel, width=100, fg_color="#8e44ad", hover_color="#732d91")
        self.btn_browse.grid(row=0, column=1, padx=10, pady=5)

        self.lbl_path = ctk.CTkLabel(self, text="ยังไม่ได้เลือก", text_color="gray")
        self.lbl_path.grid(row=0, column=2, columnspan=4, padx=10, pady=5, sticky="w")
        
        self.btn_delete = ctk.CTkButton(self, text="❌ ลบ", width=60, fg_color="#e74c3c", hover_color="#c0392b", command=lambda: self.delete_callback(self))
        self.btn_delete.grid(row=0, column=7, padx=10, pady=5, sticky="e")

        # แถว 1: ตั้งค่า Excel + ฟอนต์ + ขนาด
        self.lbl_col = ctk.CTkLabel(self, text="คอลัมน์:")
        self.lbl_col.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.col_var = ctk.StringVar(value="-")
        self.opt_col = ctk.CTkOptionMenu(self, values=["-"], variable=self.col_var, command=self.extract_data, width=60)
        self.opt_col.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.lbl_row = ctk.CTkLabel(self, text="เริ่มแถว:")
        self.lbl_row.grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_start_row = ctk.CTkEntry(self, width=50)
        self.entry_start_row.insert(0, "1")
        self.entry_start_row.bind("<KeyRelease>", lambda e: self.extract_data())
        self.entry_start_row.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        self.lbl_font = ctk.CTkLabel(self, text="ฟอนต์:")
        self.lbl_font.grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.font_var = ctk.StringVar(value=DEFAULT_THAI_FONT)
        self.opt_font = ctk.CTkComboBox(self, values=FONT_NAMES_LIST, variable=self.font_var, command=lambda v: self.update_preview_callback(), width=160)
        self.opt_font.grid(row=1, column=5, padx=5, pady=5, sticky="w")

        self.lbl_size = ctk.CTkLabel(self, text="ขนาดอักษร:")
        self.lbl_size.grid(row=1, column=6, padx=5, pady=5, sticky="e")
        self.entry_size = ctk.CTkEntry(self, width=50)
        self.entry_size.insert(0, "14")
        self.entry_size.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_size.grid(row=1, column=7, padx=5, pady=5, sticky="w")

        # แถว 2: จัดหน้า + พิกัด
        self.lbl_align = ctk.CTkLabel(self, text="จัดหน้า:")
        self.lbl_align.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.align_var = ctk.StringVar(value="ซ้าย (Left)")
        self.opt_align = ctk.CTkOptionMenu(self, values=["ซ้าย (Left)", "กลาง (Center)", "ขวา (Right)"], variable=self.align_var, command=lambda v: self.update_preview_callback(), width=100)
        self.opt_align.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        self.lbl_x = ctk.CTkLabel(self, text="X(mm):")
        self.lbl_x.grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.entry_x = ctk.CTkEntry(self, width=60)
        self.entry_x.insert(0, "10")
        self.entry_x.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_x.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        self.lbl_y = ctk.CTkLabel(self, text="Y(mm):")
        self.lbl_y.grid(row=2, column=4, padx=5, pady=5, sticky="e")
        self.entry_y = ctk.CTkEntry(self, width=60)
        self.entry_y.insert(0, "45") 
        self.entry_y.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_y.grid(row=2, column=5, padx=5, pady=5, sticky="w")

        self.lbl_w = ctk.CTkLabel(self, text="กว้างเขต(mm):")
        self.lbl_w.grid(row=2, column=6, padx=5, pady=5, sticky="e")
        self.entry_w = ctk.CTkEntry(self, width=60)
        self.entry_w.insert(0, "40") 
        self.entry_w.bind("<KeyRelease>", lambda e: self.update_preview_callback())
        self.entry_w.grid(row=2, column=7, padx=5, pady=5, sticky="w")

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path, header=None)
                col_options = [str(i + 1) for i in range(len(self.df.columns))]
                self.col_mapping = {str(i + 1): i for i in range(len(self.df.columns))}
                self.opt_col.configure(values=col_options)
                self.col_var.set(col_options[0]) 
                self.extract_data()
            except Exception as e:
                messagebox.showerror("ข้อผิดพลาด", f"อ่านไฟล์ Excel ไม่ได้:\n{e}")

    def extract_data(self, *args):
        if self.df is None: return
        try:
            start_row = max(1, int(self.entry_start_row.get() or 1))
            col_idx = self.col_mapping.get(self.col_var.get(), 0)
            raw_data = self.df.iloc[(start_row - 1):, col_idx].astype(str).tolist()
            self.text_data_list = [d for d in raw_data if d.strip() != '' and d.strip().lower() != 'nan']
            self.lbl_path.configure(text=f"พบข้อมูล {len(self.text_data_list)} บรรทัด", text_color="green")
        except Exception:
            self.text_data_list = []
            self.lbl_path.configure(text="ดึงข้อมูลล้มเหลว", text_color="red")
        self.update_preview_callback()

    def get_config(self):
        try: fs = int(self.entry_size.get())
        except: fs = 14
        
        try:
            return {
                "type": "text", "items": self.text_data_list,
                "x_mm": float(self.entry_x.get() or 0), 
                "y_mm": float(self.entry_y.get() or 0), 
                "width_mm": float(self.entry_w.get() or 10),
                "font_size": fs, 
                "align": self.align_var.get(),
                "font_name": self.font_var.get()
            }
        except ValueError: return None


class PDFGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Layout Generator to PDF")
        
        w, h = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+0+0")
        try: self.state('zoomed') 
        except: pass 

        self.groups = []
        self.group_counter = 1
        self.preview_images = []

        self.lbl_header = ctk.CTkLabel(self, text="🔲 จัดเลย์เอาต์ รูปภาพ, QR Code และ ข้อความ ลง PDF", font=("Helvetica", 24, "bold"))
        self.lbl_header.pack(pady=10)

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.left_panel = ctk.CTkFrame(self.main_frame, width=800)
        self.left_panel.pack(side="left", fill="y", padx=10)

        self.right_panel = ctk.CTkFrame(self.main_frame)
        self.right_panel.pack(side="right", fill="both", expand=True, padx=10)

        self.frame_page = ctk.CTkFrame(self.left_panel)
        self.frame_page.pack(pady=10, padx=10, fill="x")
        
        self.lbl_page_title = ctk.CTkLabel(self.frame_page, text="📄 ตั้งค่าหน้ากระดาษ", font=("Helvetica", 16, "bold"))
        self.lbl_page_title.grid(row=0, column=0, columnspan=4, pady=5, padx=10, sticky="w")

        self.lbl_size = ctk.CTkLabel(self.frame_page, text="ขนาด:")
        self.lbl_size.grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.page_size_var = ctk.StringVar(value="A4")
        self.opt_size = ctk.CTkOptionMenu(self.frame_page, values=list(PAGE_SIZES_MM.keys()) + ["กำหนดเอง (Custom)"], variable=self.page_size_var, command=self.on_page_setting_change)
        self.opt_size.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.lbl_orient = ctk.CTkLabel(self.frame_page, text="แนว:")
        self.lbl_orient.grid(row=1, column=2, padx=10, pady=5, sticky="e")
        self.orient_var = ctk.StringVar(value="Portrait (แนวตั้ง)")
        self.opt_orient = ctk.CTkOptionMenu(self.frame_page, values=["Portrait (แนวตั้ง)", "Landscape (แนวนอน)"], variable=self.orient_var, command=self.on_page_setting_change)
        self.opt_orient.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        self.frame_custom = ctk.CTkFrame(self.frame_page, fg_color="transparent")
        self.frame_custom.grid(row=2, column=0, columnspan=4, pady=0, sticky="w")
        self.lbl_custom_w = ctk.CTkLabel(self.frame_custom, text="กว้าง(mm):")
        self.entry_custom_w = ctk.CTkEntry(self.frame_custom, width=70)
        self.entry_custom_w.insert(0, "330") 
        self.entry_custom_w.bind("<KeyRelease>", lambda e: self.update_preview())
        self.lbl_custom_h = ctk.CTkLabel(self.frame_custom, text="สูง(mm):")
        self.entry_custom_h = ctk.CTkEntry(self.frame_custom, width=70)
        self.entry_custom_h.insert(0, "480") 
        self.entry_custom_h.bind("<KeyRelease>", lambda e: self.update_preview())

        self.scrollable_frame = ctk.CTkScrollableFrame(self.left_panel, width=800)
        self.scrollable_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.btn_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.btn_frame.pack(pady=5)

        self.btn_add_img_group = ctk.CTkButton(self.btn_frame, text="➕ เพิ่มกลุ่มรูปภาพ", command=self.add_image_group)
        self.btn_add_img_group.grid(row=0, column=0, padx=5)

        self.btn_add_qr_group = ctk.CTkButton(self.btn_frame, text="➕ เพิ่มกลุ่ม QR Code", command=self.add_qr_group, fg_color="#27ae60", hover_color="#219a52")
        self.btn_add_qr_group.grid(row=0, column=1, padx=5)

        self.btn_add_text_group = ctk.CTkButton(self.btn_frame, text="➕ เพิ่มกลุ่มข้อความ", command=self.add_text_group, fg_color="#8e44ad", hover_color="#732d91")
        self.btn_add_text_group.grid(row=0, column=2, padx=5)

        self.lbl_status = ctk.CTkLabel(self.left_panel, text="พร้อมทำงาน", font=("Helvetica", 14), text_color="gray")
        self.lbl_status.pack(pady=(15, 0))

        self.progress = ctk.CTkProgressBar(self.left_panel, width=400)
        self.progress.pack(pady=5)
        self.progress.set(0)

        self.btn_generate = ctk.CTkButton(self.left_panel, text="🖨️ สร้างไฟล์ PDF และ Excel", fg_color="#2980b9", hover_color="#1f618d", command=self.start_generate_thread, font=("Helvetica", 16, "bold"), height=40)
        self.btn_generate.pack(pady=10)

        self.lbl_preview = ctk.CTkLabel(self.right_panel, text="👁️ จำลองหน้ากระดาษ (Preview)", font=("Helvetica", 16, "bold"))
        self.lbl_preview.pack(pady=10)
        
        self.preview_canvas = Canvas(self.right_panel, bg="#ececec", highlightthickness=0)
        self.preview_canvas.pack(pady=10, fill="both", expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self.update_preview())

        self.on_page_setting_change() 
        self.add_qr_group()
        self.add_text_group()

    def get_current_page_size_mm(self):
        size_name = self.page_size_var.get()
        if size_name == "กำหนดเอง (Custom)":
            try: w_mm, h_mm = float(self.entry_custom_w.get()), float(self.entry_custom_h.get())
            except ValueError: w_mm, h_mm = 210.0, 297.0 
        else:
            w_mm, h_mm = PAGE_SIZES_MM[size_name]

        if "Landscape" in self.orient_var.get(): return max(w_mm, h_mm), min(w_mm, h_mm)
        else: return min(w_mm, h_mm), max(w_mm, h_mm)

    def on_page_setting_change(self, *args):
        if self.page_size_var.get() == "กำหนดเอง (Custom)":
            self.lbl_custom_w.grid(row=0, column=0, padx=10, pady=5)
            self.entry_custom_w.grid(row=0, column=1, padx=5, pady=5)
            self.lbl_custom_h.grid(row=0, column=2, padx=10, pady=5)
            self.entry_custom_h.grid(row=0, column=3, padx=5, pady=5)
        else:
            self.lbl_custom_w.grid_forget()
            self.entry_custom_w.grid_forget()
            self.lbl_custom_h.grid_forget()
            self.entry_custom_h.grid_forget()
        self.update_preview()

    def add_image_group(self):
        new_group = ImageGroupFrame(self.scrollable_frame, self.group_counter, self.update_preview, self.delete_group)
        new_group.pack(fill="x", pady=5, padx=5)
        self.groups.append(new_group)
        self.group_counter += 1
        self.update_preview()

    def add_qr_group(self):
        new_group = ExcelQrGroupFrame(self.scrollable_frame, self.group_counter, self.update_preview, self.delete_group)
        new_group.pack(fill="x", pady=5, padx=5)
        self.groups.append(new_group)
        self.group_counter += 1
        self.update_preview()

    def add_text_group(self):
        new_group = ExcelTextGroupFrame(self.scrollable_frame, self.group_counter, self.update_preview, self.delete_group)
        new_group.pack(fill="x", pady=5, padx=5)
        self.groups.append(new_group)
        self.group_counter += 1
        self.update_preview()

    def delete_group(self, group_frame):
        if group_frame in self.groups: self.groups.remove(group_frame)
        group_frame.destroy()
        self.update_preview()

    def update_preview(self, *args):
        w_mm, h_mm = self.get_current_page_size_mm()
        if w_mm <= 0 or h_mm <= 0: return

        canvas_width = max(500, self.preview_canvas.winfo_width())
        canvas_height = max(600, self.preview_canvas.winfo_height())

        scale_x, scale_y = (canvas_width - 40) / max(1, w_mm), (canvas_height - 40) / max(1, h_mm)
        scale = min(scale_x, scale_y) 

        paper_w_px, paper_h_px = w_mm * scale, h_mm * scale
        offset_x, offset_y = (canvas_width - paper_w_px) / 2, (canvas_height - paper_h_px) / 2

        self.preview_canvas.delete("all")
        self.preview_images.clear() 
        self.preview_canvas.create_rectangle(offset_x, offset_y, offset_x + paper_w_px, offset_y + paper_h_px, fill="white", outline="#999", width=2)
        
        colors = ["#3498db", "#e74c3c", "#2ecc71", "#f39c12", "#9b59b6", "#1abc9c", "#34495e", "#d35400"]

        for idx, group in enumerate(self.groups):
            cfg = group.get_config()
            if not cfg: continue
            
            color = colors[idx % len(colors)]
            x_mm, y_mm, img_w_mm = cfg["x_mm"], cfg["y_mm"], cfg["width_mm"]
            x_px, y_px = offset_x + (x_mm * scale), offset_y + (y_mm * scale)
            img_w_px = max(1, img_w_mm * scale)

            if cfg["type"] == "text":
                align_val = cfg.get("align", "ซ้าย")
                if "ซ้าย" in align_val: 
                    anchor, draw_x, justify_mode = "nw", x_px, "left"
                elif "กลาง" in align_val: 
                    anchor, draw_x, justify_mode = "n", x_px + (img_w_px / 2), "center"
                else: 
                    anchor, draw_x, justify_mode = "ne", x_px + img_w_px, "right"

                text_val = str(cfg["items"][0]) if cfg["items"] else f"ทดสอบข้อความ กลุ่ม {group.group_index}"
                font_size = max(6, int(cfg.get("font_size", 12) * scale * 0.35)) 
                selected_font = cfg.get("font_name", "Helvetica")
                
                self.preview_canvas.create_text(
                    draw_x, y_px, text=text_val, anchor=anchor, justify=justify_mode, 
                    width=img_w_px, font=(selected_font, font_size), fill=color
                )
                
                self.preview_canvas.create_line(x_px, y_px, x_px, y_px + 20, fill=color, dash=(2, 2))
                self.preview_canvas.create_line(x_px + img_w_px, y_px, x_px + img_w_px, y_px + 20, fill=color, dash=(2, 2))
                self.preview_canvas.create_line(x_px, y_px, x_px + img_w_px, y_px, fill=color, dash=(2, 2))
                continue

            ratio = get_image_aspect_ratio(cfg["items"][0]) if cfg["items"] and cfg["type"] == "image" else 1.0
            img_h_px = int(max(1, (img_w_mm * ratio) * scale))
            img_w_px = int(img_w_px)

            if cfg["items"]:
                try:
                    img = Image.open(cfg["items"][0]) if cfg["type"] == "image" else generate_qr_pil(str(cfg["items"][0])).convert("RGB")
                    img = img.resize((img_w_px, img_h_px), Image.LANCZOS)
                    tk_img = ImageTk.PhotoImage(img)
                    self.preview_images.append(tk_img)
                    self.preview_canvas.create_image(x_px, y_px, anchor="nw", image=tk_img)
                    self.preview_canvas.create_rectangle(x_px, y_px, x_px + img_w_px, y_px + img_h_px, outline=color, width=3)
                except Exception as e: print(e)
            else:
                self.preview_canvas.create_rectangle(x_px, y_px, x_px + img_w_px, y_px + img_h_px, fill=color, stipple="gray50", outline=color, width=2)
                self.preview_canvas.create_text(x_px + (img_w_px/2), y_px + (img_h_px/2), text=f"กลุ่ม {group.group_index}", fill="black", font=("Helvetica", 10, "bold"))

    def start_generate_thread(self):
        if not self.groups:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเพิ่มกลุ่มอย่างน้อย 1 กลุ่มก่อนสร้าง PDF")
            return

        configs = []
        max_pages = 0
        for group in self.groups:
            cfg = group.get_config()
            if cfg is None:
                messagebox.showerror("ข้อผิดพลาด", f"กรุณากรอกตัวเลขพิกัดให้ถูกต้องในกลุ่มที่ {group.group_index}")
                return
            if not cfg["items"]:
                messagebox.showwarning("แจ้งเตือน", f"กลุ่มที่ {group.group_index} ยังไม่ได้เลือกข้อมูล (โฟลเดอร์ หรือ Excel)")
                return
            configs.append(cfg)
            if len(cfg["items"]) > max_pages: max_pages = len(cfg["items"])

        if max_pages == 0: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not save_path: return

        page_w_mm, page_h_mm = self.get_current_page_size_mm()
        self.btn_generate.configure(state="disabled", text="กำลังประมวลผล...")
        self.progress.set(0)
        self.lbl_status.configure(text="เตรียมการสร้างไฟล์...", text_color="black")
        threading.Thread(target=self.generate_pdf_worker, args=(configs, max_pages, save_path, page_w_mm, page_h_mm), daemon=True).start()

    def generate_pdf_worker(self, configs, max_pages, save_path, page_w_mm, page_h_mm):
        try:
            page_w_pt, page_h_pt = page_w_mm * mm, page_h_mm * mm
            c = pdf_canvas.Canvas(save_path, pagesize=(page_w_pt, page_h_pt))

            registered_fonts = set()
            
            # เก็บข้อมูลเพื่อนำไปสร้าง Excel
            excel_data = []

            for page_idx in range(max_pages):
                if page_idx > 0: c.showPage()
                
                current_page_text = "" # เก็บข้อความเฉพาะกลุ่ม "text" ของแต่ละหน้า

                for cfg in configs:
                    if page_idx >= len(cfg["items"]): continue
                    
                    item_data = cfg["items"][page_idx]
                    x_pt, w_pt = cfg["x_mm"] * mm, cfg["width_mm"] * mm
                    
                    if cfg["type"] == "text":
                        # บันทึกข้อความลงตัวแปร ตรงนี้แหละที่เราจะได้ภาษาไทยเป๊ะๆ 100%
                        current_page_text = str(item_data).replace('\n', ' ')
                        
                        font_key = cfg.get("font_name", "Helvetica")
                        font_path = SYSTEM_FONTS.get(font_key)
                        used_font = "Helvetica"
                        
                        if font_path and os.path.exists(font_path):
                            if font_key not in registered_fonts:
                                try:
                                    pdfmetrics.registerFont(TTFont(font_key, font_path))
                                    registered_fonts.add(font_key)
                                    used_font = font_key
                                except Exception:
                                    pass 
                            else:
                                used_font = font_key

                        align_val = cfg["align"]
                        if "ซ้าย" in align_val: align_code = TA_LEFT
                        elif "กลาง" in align_val: align_code = TA_CENTER
                        else: align_code = TA_RIGHT

                        text_str = str(item_data).replace('\n', '<br/>')
                        
                        style = ParagraphStyle(
                            name='CustomText',
                            fontName=used_font,
                            fontSize=cfg["font_size"],
                            leading=cfg["font_size"] * 1.3, 
                            alignment=align_code,
                            wordWrap='CJK' 
                        )
                        
                        p = Paragraph(text_str, style)
                        p_w, p_h = p.wrap(w_pt, page_h_pt)
                        draw_y = page_h_pt - (cfg["y_mm"] * mm) - p_h + (cfg["font_size"] * 0.2)
                        p.drawOn(c, x_pt, draw_y)
                        
                    elif cfg["type"] == "image":
                        h_pt = w_pt * get_image_aspect_ratio(item_data)
                        y_pt = page_h_pt - (cfg["y_mm"] * mm) - h_pt 
                        c.drawImage(item_data, x_pt, y_pt, width=w_pt, height=h_pt)
                        
                    elif cfg["type"] == "qrcode":
                        y_pt = page_h_pt - (cfg["y_mm"] * mm) - w_pt 
                        qr_img = generate_qr_pil(str(item_data))
                        c.drawImage(ImageReader(qr_img), x_pt, y_pt, width=w_pt, height=w_pt)
                
                # เก็บข้อมูลแถวใหม่สำหรับหน้าปัจจุบันลง List
                excel_data.append([page_idx + 1, current_page_text])
                self.after(0, self.update_ui, (page_idx + 1) / max_pages, f"กำลังสร้างหน้าที่ {page_idx + 1} จาก {max_pages}")
            
            c.save()
            
            # --- สร้างไฟล์ Excel สรุปข้อมูล อัตโนมัติ ---
            if excel_data:
                excel_save_path = save_path.replace('.pdf', '_List.xlsx')
                df_export = pd.DataFrame(excel_data, columns=["หน้ากระดาษ (Page)", "ข้อความที่พิมพ์บนดวง"])
                df_export.to_excel(excel_save_path, index=False)

            self.after(0, self.finish_generation, max_pages, save_path)
        except Exception as e:
            self.after(0, self.show_error, str(e))

    def update_ui(self, pct, status_text):
        self.progress.set(pct)
        self.lbl_status.configure(text=status_text)

    def finish_generation(self, max_pages, save_path):
        self.progress.set(1.0)
        self.lbl_status.configure(text="✅ เสร็จสิ้น!", text_color="green")
        messagebox.showinfo("สำเร็จ", f"สร้าง PDF จำนวน {max_pages} หน้า สำเร็จ!\nและสร้างไฟล์ Excel สรุปรายชื่อให้แล้วที่:\n{save_path.replace('.pdf', '_List.xlsx')}")
        self.btn_generate.configure(state="normal", text="🖨️ สร้างไฟล์ PDF และ Excel")

    def show_error(self, error_msg):
        self.lbl_status.configure(text="❌ เกิดข้อผิดพลาด", text_color="red")
        messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{error_msg}")
        self.btn_generate.configure(state="normal", text="🖨️ สร้างไฟล์ PDF และ Excel")

if __name__ == "__main__":
    app = PDFGeneratorApp()
    app.mainloop()