import os
import pyzipper
import win32com.client
import shutil
import time
from datetime import datetime
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk
import urllib.request
import threading
import json
from tkinter import messagebox

# =============================================
#  SwiftFolder Pro  — Premium Dark UI
# تصميم محسّن بألوان داكنة احترافية مع ألوان صيغ الملفات
# =============================================

# إعدادات المظهر العام
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# --- ألوان التصميم المحسّن ---
COLORS = {
    "bg_main": "#0d1117",         # خلفية رئيسية - أسود مزرق عميق (GitHub Dark)
    "bg_header": "#161b22",       # خلفية الشريط العلوي
    "bg_card": "#21262d",         # خلفية البطاقات
    "bg_footer": "#161b22",       # خلفية الشريط السفلي
    "bg_sidebar": "#0d1117",      # خلفية الشريط الجانبي
    "accent_primary": "#58a6ff",  # أزرق مميز - للأزرار الرئيسية
    "accent_hover": "#388bfd",    # أزرق أغمق - عند التمرير
    "accent_green": "#3fb950",    # أخضر - للنجاح
    "accent_purple": "#bc8cff",   # بنفسجي - للتمييز
    "accent_orange": "#d29922",   # برتقالي - للتحذير
    "btn_secondary": "#21262d",   # أزرار ثانوية
    "btn_sec_hover": "#30363d",   # أزرار ثانوية عند التمرير
    "danger": "#f85149",          # أحمر - للحذف
    "danger_hover": "#da3633",    # أحمر أغمق
    "text_primary": "#f0f6fc",    # نص رئيسي - أبيض ناعم
    "text_secondary": "#8b949e",  # نص ثانوي - رمادي
    "text_muted": "#484f58",      # نص خافت
    "border": "#30363d",          # حدود
    "border_light": "#21262d",    # حدود خفيفة
    "row_even": "#0d1117",        # صف زوجي
    "row_odd": "#161b22",         # صف فردي
    "row_selected": "#1f3a5f",    # صف محدد
    "tree_header": "#1c2128",     # رأس الجدول
    "logo_gradient_1": "#58a6ff", # لون اللوقو 1
    "logo_gradient_2": "#bc8cff", # لون اللوقو 2
    "input_bg": "#0d1117",        # خلفية حقول الإدخال
    "input_border": "#30363d",    # حدود حقول الإدخال
    "input_focus": "#58a6ff",     # حدود التركيز
}

# --- ألوان صيغ الملفات المشهورة ---
FILE_TYPE_COLORS = {
    # مستندات
   
     
}

# لون افتراضي للصيغ غير المعروفة
DEFAULT_FILE_COLOR = {"bg": "#484f58", "fg": "#ffffff", "icon": "📄"}


def get_file_type_info(file_type):
    """الحصول على معلومات لون الصيغة"""
    return FILE_TYPE_COLORS.get(file_type.upper(), DEFAULT_FILE_COLOR)

class CustomMessageBox(ctk.CTkToplevel):
    """نافذة رسائل مخصصة مع لوقو البرنامج"""
    
    def __init__(self, parent, title="", message="", msg_type="info", show_cancel=False):
        super().__init__(parent)
        
        self.result = None
        self.title(title)
        self.configure(fg_color=COLORS["bg_main"])
        self.resizable(False, False)
        self.grab_set()
        self.attributes("-topmost", True)
        
        # حجم النافذة
        width = 420
        height = 260 if show_cancel else 230
        self.geometry(f"{width}x{height}")
        
        # توسيط النافذة
        self.update_idletasks()
        if parent:
            px = parent.winfo_x() + (parent.winfo_width() // 2) - (width // 2)
            py = parent.winfo_y() + (parent.winfo_height() // 2) - (height // 2)
            self.geometry(f"+{px}+{py}")
        
        # الإطار الرئيسي
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(expand=True, fill="both", padx=25, pady=20)
        
        # شريط اللوقو العلوي
        logo_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        logo_frame.pack(fill="x", pady=(0, 15))
        
        
        # اللوقو
        logo_text = ""
        ctk.CTkLabel(
            logo_frame, text=logo_text,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=COLORS["accent_primary"]
        ).pack(side="left")
        
        # أيقونة نوع الرسالة
        type_icons = {
            "info": ("ℹ️", COLORS["accent_primary"]),
            "success": ("✅", COLORS["accent_green"]),
            "warning": ("⚠️", COLORS["accent_orange"]),
            "error": ("❌", COLORS["danger"]),
            "question": ("❓", COLORS["accent_purple"]),
        }
        icon, icon_color = type_icons.get(msg_type, ("ℹ️", COLORS["accent_primary"]))
        
        # أيقونة كبيرة
        ctk.CTkLabel(
            main_frame, text=icon,
            font=ctk.CTkFont(size=36),
        ).pack(pady=(0, 8))
        
        # عنوان الرسالة
        ctk.CTkLabel(
            main_frame, text=title,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=COLORS["text_primary"]
        ).pack(pady=(0, 5))
        
        # نص الرسالة
        ctk.CTkLabel(
            main_frame, text=message,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            wraplength=360
        ).pack(pady=(0, 15))
        
        # إطار الأزرار
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        if show_cancel:
            # زر إلغاء 🚫
            ctk.CTkButton(
                btn_frame, text=" إلغاء" if title != "Confirm Delete" else " Cancel",
                width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color="#2d333b",
                hover_color="#444c56",
                border_width=1, border_color="#444c56",
                corner_radius=12,
                text_color="#adbac7",
                command=self._on_cancel
            ).pack(side="left", expand=True, padx=5)
            
            # زر تأكيد ✅ أو حذف 🗑️
            if msg_type in ("error", "question"):
                confirm_color = COLORS["danger"]
                confirm_hover = COLORS["danger_hover"]
                confirm_text = " حذف" if title != "Confirm Delete" else " Delete"
            else:
                confirm_color = COLORS["accent_primary"]
                confirm_hover = COLORS["accent_hover"]
                confirm_text = "✅ تأكيد" if title != "Confirm Delete" else "✅ Confirm"
            
            ctk.CTkButton(
                btn_frame, text=confirm_text,
                width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color=confirm_color,
                hover_color=confirm_hover,
                corner_radius=12,
                text_color="#ffffff",
                command=self._on_confirm
            ).pack(side="right", expand=True, padx=5)
        else:
            # زر موافق فقط
            if msg_type == "success":
                btn_color = COLORS["accent_green"]
                btn_hover = "#2ea043"
                ok_text = "حسناً" if title not in ("Success", "Error") else "OK"
            elif msg_type == "error":
                btn_color = COLORS["danger"]
                btn_hover = COLORS["danger_hover"]
                ok_text = "😔 حسناً" if title not in ("Success", "Error") else "😔 OK"
            elif msg_type == "warning":
                btn_color = COLORS["accent_orange"]
                btn_hover = "#b88a1e"
                ok_text = "⚠️ فهمت" if title not in ("Success", "Error") else "⚠️ Got it"
            else:
                btn_color = COLORS["accent_primary"]
                btn_hover = COLORS["accent_hover"]
                ok_text = "👌 حسناً" if title not in ("Success", "Error") else "👌 OK"
            
            ctk.CTkButton(
                btn_frame, text=ok_text,
                width=170, height=38,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color=btn_color,
                hover_color=btn_hover,
                corner_radius=12,
                text_color="#ffffff",
                command=self._on_confirm
            ).pack(expand=True)
        
        # ربط مفتاح Enter و Escape
        self.bind("<Return>", lambda e: self._on_confirm())
        self.bind("<Escape>", lambda e: self._on_cancel())
       
        
        self.wait_window()
    
    def _on_confirm(self):
        self.result = True
        self.destroy()
    
    def _on_cancel(self):
        self.result = False
        self.destroy()


class SwiftFolderPro(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # --- إعداد المجلدات والقواعد ---
        self.archive_dir = os.path.join(os.getcwd(), "SwiftArchive")
        if not os.path.exists(self.archive_dir):
            os.makedirs(self.archive_dir)

        self.password_file = "config.txt"
        if not os.path.exists(self.password_file):
            with open(self.password_file, "w") as f: f.write("1234")

        self.customers_file = "customers.txt"
        if not os.path.exists(self.customers_file):
            with open(self.customers_file, "w", encoding="utf-8") as f:
                f.write("كل الزبائن")

        # --- قاموس اللغات ---
        self.languages = {
            "العربية": {
                "title": "SwiftFolder Pro v1.0.0",
                "settings": "⚙️ الإعدادات",
                "backup": "☁️النسخة الاحتياطية",
                "import_btn": "📥 الاستيراد",
                "scan": "🖨️الماسح الضوئي",
                "add": "➕إضافة ملف",
                "search": "🔍 هنا إبحث...",
                "all_years": "السنوات كل",
                "all_types": "الأنواع كل",
                "all_customer":" الزبائن كل ",
                "col_name": "اسم الملف",
                "col_date": "تاريخ الإضافة",
                "col_size": "الحجم",
                "col_type": "النوع",
                "btn_open": "فتح الملف",
                "btn_edit": "تعديل الاسم",
                "btn_delete": "حذف الملف",
                "copyright": "© 2026 SwiftFolder Pro — Developed by Zohir Mk",
                "logo": "⚡ SwiftFolder Pro",
                "sett_win_title": "الإعدادات",
                "sett_lang_tab": "🌐 اللغة",
                "sett_theme_tab": "المظهر",
                "sett_pwd_tab": "🔐 كلمة المرور",
                "sett_lang_head": "🌐 لغة البرنامج",
                "sett_pwd_head": "🔐 حماية البرنامج",
                "sett_old_pwd": "الحالية السر كلمة",
                "sett_new_pwd": "الجديدة السر كلمة",
                "sett_save": "حفظ",
                "file_count": "الملفات إجمالي:",
                "sett_close": "إغلاق",
                "msg_confirm_del": "هل أنت متأكد من حذف الملفات المحددة؟",
                "msg_del_title": "تأكيد الحذف",
                "msg_error": "خطأ",
                "scan_name_title": "تسمية الملف",
                "scan_name_msg": "التلقائي بالاسم للحفظ إلغاء اضغط أو الملف اسم أدخل:",
                "scanning_title": "ماسح المستندات",
                "scan_start": "جاري بدء المسح...",
                "scanned_count": ": الأوراق الممسوحة",
                "scan_done": "تم المسح بنجاح",
                "success_title": "",
                "msg_success": "تم استيراد النسخة بنجاح!",
                "backup_success": "تم إنشاء النسخة الاحتياطية بنجاح!",
                "confirm_btn": "تأكيد",
                "cancel_btn": "إلغاء",
                "ok_btn": "حسناً",
                "login_title": "تسجيل الدخول",
                "login_msg": ": الرجاء إدخال كلمة المرور ",
                "msg_error": "خطأ",
                "wrong_pass": ": المحاولة إعادة الرجاء  ,خاطئة المرور كلمة  ⚠️ ",
                "enter_customer_name": "ادخل اسم الزبون",
                "all_customer": "كل الزبائن",
            },
            "English": {
                "title": "SwiftFolder Pro v1.0.0",
                "settings": "⚙️ Settings",
                "backup": "☁️ Backup",
                "import_btn": "📥 Import",
                "scan": "🖨️ Scan",
                "add": "➕ Add Files",
                "search": "🔍 Search here...",
                "all_years": "All Years",
                "all_types": "All Types",
                "all_customer": "all customer",
                "col_name": "File Name",
                "col_date": "Add Date",
                "col_size": "Size",
                "col_type": "Type",
                "btn_open": "Open File",
                "btn_edit": "Edit Name",
                "btn_delete": "Delete File",
                "copyright": "© 2026 SwiftFolder Pro — Developed by Zohir Mk",
                "logo": "⚡ SwiftFolder Pro",
                "sett_win_title": "Settings",
                "sett_lang_tab": "🌐 Language",
                "sett_pwd_tab": "🔐 Password",
                "sett_lang_head": "🌐 App Language",
                "sett_pwd_head": "🔐 App Protection",
                "sett_old_pwd": "Current Password",
                "sett_new_pwd": "New Password",
                "sett_save": "Save",
                "file_count": "Total Files: ",
                "sett_close": "Close",
                "msg_confirm_del": "Are you sure you want to delete the selected files?",
                "msg_del_title": "Confirm Delete",
                "msg_error": "Error",
                "scan_name_title": "File Naming",
                "scan_name_msg": "Enter file name or press Cancel to use default:",
                "scanning_title": "Document Scanner",
                "scan_start": "Starting scan...",
                "scanned_count": "Scanned pages: ",
                "scan_done": "Scan Complete",
                "success_title": "Success",
                "msg_success": "Backup imported successfully",
                "backup_success": "Backup created successfully",
                "confirm_btn": "Confirm",
                "cancel_btn": "Cancel",
                "ok_btn": "OK",
                "login_title": "App Protection",
                "login_msg": "Enter password to access:",
                "msg_error": "Error",
                "wrong_pass": "Wrong password, please try again.",
                "enter_customer_name": "Enter Customer Name",
                "all_customer": "All Customers",
            }
        }
        
        
        
        self.current_lang = "العربية"
        self.withdraw() 
        self.after(100, self.authenticate)
        
        self.title(self.languages[self.current_lang]["title"])
        self.geometry("1100x720")
        self.configure(fg_color=COLORS["bg_main"])
        self.all_files_data = []
        self.settings_window = None

        self.setup_styles()
        self.setup_ui()
        self.update_clock()
        self.load_initial_archive()
        
        # إعدادات التحديث التلقائي
        self.CURRENT_VERSION = "1.0.0"
        # استبدل USERNAME باسم حسابك على GitHub بدقة
        self.VERSION_URL = "https://raw.githubusercontent.com/zohir94/SwiftFolderPro-/refs/heads/main/version.txt"
        
        
    
    #-----------------------------------------------------------------------------------------------------------------------------------

            
    def authenticate(self):
        lang = self.languages[self.current_lang]
        try:
            with open(self.password_file, "r") as f:
                correct_pass = f.read().strip()
        except:
            correct_pass = "1234"

        # نافذة إدخال كلمة المرور (بطاقة مركزية احترافية)
        login_dialog = ctk.CTkToplevel(self)
        login_dialog.title(lang["login_title"])
        login_dialog.geometry("420x500")
        login_dialog.resizable(False, False)
        login_dialog.grab_set()

        # مركز النافذة في شاشات الكمبيوتر
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 420, 500
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        login_dialog.geometry(f"{w}x{h}+{x}+{y}")

        # --- القسم الأول: الهوية البصرية والترحيب ---
        # شعار البرنامج النصي الكبير
        lbl_logo = ctk.CTkLabel(login_dialog, text="SFP", font=("Impact", 75), text_color="#1f538d")
        lbl_logo.pack(pady=(25, 0))

        # اسم البرنامج الرئيسي
        lbl_title = ctk.CTkLabel(login_dialog, text="SwiftFolder Pro", font=("Segoe UI", 24, "bold"))
        lbl_title.pack(pady=(0, 2))

        # العنوان الفرعي أو الرسالة بلغة المستخدم
        lbl_subtitle = ctk.CTkLabel(login_dialog, text=lang["login_msg"], font=("Segoe UI", 13), text_color="gray")
        lbl_subtitle.pack(pady=(0, 25))

        # --- القسم الثاني: حقول الإدخال والتحكم ---
        # إطار لحقل كلمة المرور مائل للعصرية
        entry_frame = ctk.CTkFrame(login_dialog, corner_radius=10, fg_color="#2d333b", height=45)
        entry_frame.pack(pady=10, padx=35, fill="x")

        # حقل كلمة المرور
        entry = ctk.CTkEntry(entry_frame, show="*", border_width=0, fg_color="#2d333b", 
                             font=("Segoe UI", 14), justify="center")
        entry.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=5)
        entry.focus()

        # دالة تبديل أيقونة القفل وإظهار/إخفاء الباسوورد
        def toggle_password():
            if entry.cget("show") == "*":
                entry.configure(show="")
                lock_label.configure(text="🔓")
            else:
                entry.configure(show="*")
                lock_label.configure(text="🔒")

        # أيقونة القفل التفاعلية بالكامل
        lock_label = ctk.CTkLabel(entry_frame, text="🔒", width=30, height=30, font=("Segoe UI", 16), fg_color=None, cursor="hand2")
        lock_label.pack(side="right", padx=12)
        lock_label.bind("<Button-1>", lambda e: toggle_password())

        # --- القسم الثالث: أزرار الأكشن والتنبيهات المدمجة ---
        # مساحة التنبيهات الذكية (مخفية في البداية)
        lbl_error = ctk.CTkLabel(login_dialog, text="", font=("Segoe UI", 12))
        lbl_error.pack(pady=(5, 5))

        # دالة التحقق الذكية والتوجيه لنظام التحديثات
        def check_password(event=None):
            entered = entry.get()
            if entered == correct_pass:
                lbl_error.configure(text="✔ جاري التحقق والدخول...", text_color="#2ecc71")
                login_dialog.update_idletasks()
                
                # إغلاق النافذة والانتقال للبرنامج الرئيسي وفحص التحديثات
                self.deiconify()
                login_dialog.destroy()
                self.start_update_check()
            else:
                # عرض الخطأ بلون أحمر جذاب داخل نافذة البطاقة دون تشويه العناوين القديمة
                lbl_error.configure(text="⚠️ كلمة السر خاطئة، الرجاء إعادة المحاولة", text_color="#e74c3c")
                entry.delete(0, "end")
                entry.focus()

        # زر الدخول البارز بالأزرق الاحترافي الماتش مع الهوية
        btn = ctk.CTkButton(login_dialog, text="تسجيل الدخول", command=check_password,
                            height=45, corner_radius=10, font=("Segoe UI", 15, "bold"),
                            fg_color="#1f538d", hover_color="#14375e")
        btn.pack(pady=15, padx=35, fill="x")

        # ربط زر Enter لتسهيل تجربة المستخدم
        entry.bind("<Return>", check_password)
        
    #-----------------------------------------------------------------------------------------------------------------------------------    
    def start_update_check(self):
        # تأخير إطلاق خيط الفحص لمدة 5000 ملي ثانية (5 ثوانٍ) بعد ظهور الواجهة
        # هذا يضمن أن يفتح البرنامج فوراً في ثانية واحدة، ثم يبحث عن التحديث في الخلفية دون أي تأثير
        self.after(5000, lambda: threading.Thread(target=self.check_for_updates, daemon=True).start())

    def check_for_updates(self):
        try:
            # 1. الاتصال بالرابط وقراءة رقم الإصدار من موقع GitHub
            req = urllib.request.Request(self.VERSION_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                latest_version = response.read().decode('utf-8').strip()
            
            # 2. مقارنة الإصدار الحالي بالإصدار الموجود على الإنترنت
            if latest_version != self.CURRENT_VERSION:
                if messagebox.askyesno("تحديث جديد متوفر", f"يوجد إصدار جديد للبرنامج ({latest_version}).\nهل تريد تحميل وتثبيت التحديث الآن؟"):
                    
                    exe_url = "https://github.com/zohir94/SwiftFolderPro-/releases/download/1.0.1/SwiftFolderPro.zip"
                              
                    
                    # اسم الملف المؤقت أثناء التحميل بجانب البرنامج الحالي
                    output_path = "SwiftFolderPro_New.exe" 
                    
                    # --- إنشاء نافذة شريط التحميل المرئية ---
                    from tkinter import ttk
                    import tkinter as tk
                    
                    progress_window = tk.Toplevel(self)
                    progress_window.title("جاري تحميل التحديث...")
                    progress_window.geometry("400x120")
                    progress_window.resizable(False, False)
                    
                    # مركز النافذة
                    sw, sh = progress_window.winfo_screenwidth(), progress_window.winfo_screenheight()
                    x = (sw // 2) - (400 // 2)
                    y = (sh // 2) - (120 // 2)
                    progress_window.geometry(f"400x120+{x}+{y}")
                    
                    lbl_status = tk.Label(progress_window, text="جاري الاتصال بالسيرفر...", font=("Arial", 11))
                    lbl_status.pack(pady=10)
                    
                    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=350, mode="determinate")
                    progress_bar.pack(pady=5)
                    
                    # دالة فرعية لتحديث شريط التقدم أثناء التحميل
                    def reporthook(block_num, block_size, total_size):
                        if total_size > 0:
                            downloaded = block_num * block_size
                            percent = min(int(downloaded * 100 / total_size), 100)
                            progress_bar['value'] = percent
                            lbl_status.config(text=f"تم تحميل {percent}% من ملف التحديث...")
                            progress_window.update_idletasks()
                    
                    # دالة تشغيل التحميل والتحديث التلقائي
                    def download_thread():
                        try:
                            # تحميل الملف الجديد باسم مؤقت أولاً
                            urllib.request.urlretrieve(exe_url, output_path, reporthook=reporthook)
                            progress_window.destroy() # غلق نافذة شريط التقدم
                            
                            # معرفة اسم ومسار ملف البرنامج الحالي المشغل الآن
                            import sys
                            import os
                            import subprocess
                            
                            current_exe = sys.executable
                            
                            # كود سحري لـ CMD يقوم بـ: الانتظار ثانية -> استبدال الملف الحالي بالجديد -> إعادة التشغيل
                            # تم وضع timeout 1 ثانية ليعطي فرصة لبرنامج البايثون الحالي كي يغلق تماماً ويتحرر الملف من الذاكرة
                            cmd_command = f'timeout /t 1 && move /y "{output_path}" "{current_exe}" && start "" "{current_exe}"'
                            
                            # تشغيل الأمر في الخلفية صامتاً
                            subprocess.Popen(cmd_command, shell=True)
                            
                            # إغلاق البرنامج الحالي فوراً لكي ينجح أمر الاستبدال في الـ CMD
                            os._exit(0)
                            
                        except Exception as download_error:
                            if 'progress_window' in locals() and progress_window.winfo_exists():
                                progress_window.destroy()
                            messagebox.showerror("خطأ في التحميل", f"فشل تحميل وتثبيت الملف: {download_error}")
                    
                    # تشغيل خيط التحميل في الخلفية لكي لا تتجمد الواجهة
                    import threading
                    threading.Thread(target=download_thread, daemon=True).start()
                    
        except Exception as e:
            print(f"Error checking for updates: {e}")
    #-----------------------------------------------------------------------------------------------------------------------------------
                 
    def setup_styles(self):
        """إعداد أنماط Treeview المخصصة بألوان التصميم المحسّن"""
        style = ttk.Style()
        style.theme_use("clam")
        
        # --- تنسيق الجدول ---
        style.configure("Custom.Treeview",
            background=COLORS["bg_main"],
            foreground=COLORS["text_primary"],
            fieldbackground=COLORS["bg_main"],
            borderwidth=0,
            rowheight=40,
            font=("Segoe UI", 11)
        )
        
        # --- تنسيق رأس الجدول ---
        style.configure("Custom.Treeview.Heading",
            background=COLORS["tree_header"],
            foreground=COLORS["text_primary"],
            borderwidth=0,
            font=("Segoe UI", 11, "bold"),
            padding=(12, 10)
        )
        
        # --- تأثير التمرير على رأس الجدول ---
        style.map("Custom.Treeview.Heading",
            background=[("active", COLORS["accent_primary"])]
        )
        
        # --- تأثير التحديد في الجدول ---
        style.map("Custom.Treeview",
            background=[("selected", COLORS["row_selected"])],
            foreground=[("selected", "#ffffff")]
        )

    def setup_ui(self):
        lang = self.languages[self.current_lang]
        
        # ==========================================
        # 🔝 الشريط العلوي (Header) - محسّن
        # ==========================================
        self.header_frame = ctk.CTkFrame(
            self, height=64, corner_radius=0, 
            fg_color=COLORS["bg_header"],
            border_width=0
        )
        self.header_frame.pack(side="top", fill="x")
        self.header_frame.pack_propagate(False)

        # خط فاصل ملون أسفل الهيدر
        self.header_accent_line = ctk.CTkFrame(
            self, height=2, corner_radius=0,
            fg_color=COLORS["accent_primary"]
        )
        self.header_accent_line.pack(side="top", fill="x")

        # إطار اللوقو
        logo_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        logo_frame.pack(side="left", padx=16)
        
        # أيقونة اللوقو
        self.logo_icon = ctk.CTkLabel(
            logo_frame,
            text="⚡",
            font=ctk.CTkFont(size=24),
        )
        self.logo_icon.pack(side="left", padx=(0, 6))
        
        # نص اللوقو
        self.logo_label = ctk.CTkLabel(
            logo_frame, 
            text="SwiftFolder Pro", 
            font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"),
            text_color=COLORS["accent_primary"]
        )
        self.logo_label.pack(side="left")
        
        # نسخة البرنامج
        self.version_label = ctk.CTkLabel(
            logo_frame,
            text="v1.0.0",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_muted"]
        )
        self.version_label.pack(side="left", padx=(6, 0), pady=(6, 0))
        
        
        
        # --- أزرار الشريط العلوي (من اليمين لليسار) - تصميم محسّن ---
        
        # زر الإعدادات ⚙️ - رمادي أنيق
        self.btn_settings = ctk.CTkButton(
            self.header_frame, text=lang["settings"], width=100, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#2d333b",
            hover_color="#445653",
            border_width=1, border_color="#445653",
            corner_radius=12,
            text_color="#adbac7",
            command=self.open_settings
        )
        self.btn_settings.pack(side="right", padx=4, pady=13)
        
        # زر الاستيراد 📥 - برتقالي دافئ
        self.btn_import = ctk.CTkButton(
            self.header_frame, text=lang["import_btn"], width=100, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#4a2f0a",
            hover_color="#5c3a0d",
            border_width=1, border_color=COLORS["accent_orange"],
            corner_radius=12,
            text_color=COLORS["accent_orange"],
            command=self.import_backup_folder
        )
        self.btn_import.pack(side="right", padx=4, pady=13)

        # زر النسخ الاحتياطي ☁️ - سماوي
        self.btn_backup = ctk.CTkButton(
            self.header_frame, text=lang["backup"], width=100, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0c2d48",
            hover_color="#0f3a5e",
            border_width=1, border_color=COLORS["accent_primary"],
            corner_radius=12,
            text_color=COLORS["accent_primary"],
            command=self.create_backup
        )
        self.btn_backup.pack(side="right", padx=4, pady=13)
        
        # زر الماسح الضوئي 🖨️ - بنفسجي
        self.btn_scan = ctk.CTkButton(
            self.header_frame, text=lang["scan"], width=100, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#2a1a4e",
            hover_color="#3b2566",
            border_width=1, border_color=COLORS["accent_purple"],
            corner_radius=12,
            text_color=COLORS["accent_purple"],
            command=self.scan_document
        )
        self.btn_scan.pack(side="right", padx=4, pady=13)
        
        # زر إضافة ملف ➕ - أخضر مميز (الزر الرئيسي)
        self.btn_add = ctk.CTkButton(
            self.header_frame, text=lang["add"], width=100, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=COLORS["accent_green"],
            hover_color="#362ea0",
            text_color="#ffffff",
            corner_radius=12,
            command=self.add_file
        )
        self.btn_add.pack(side="right", padx=(4, 12), pady=13)

        # ==========================================
        # 🔍 شريط الفلترة (Filter Bar) - محسّن
        # ==========================================
        self.filter_frame = ctk.CTkFrame(
            self, height=54, fg_color=COLORS["bg_header"],
            corner_radius=12, border_width=1, border_color=COLORS["border"]
        )
        self.filter_frame.pack(side="top", fill="x", padx=20, pady=(12, 8))

        # حقل البحث
        self.search_entry = ctk.CTkEntry(
            self.filter_frame, 
            placeholder_text=lang["search"], 
            width=420, height=38,
            font=ctk.CTkFont(size=12),
            fg_color=COLORS["input_bg"],
            border_color=COLORS["input_border"],
            text_color=COLORS["text_primary"],
            placeholder_text_color=COLORS["text_secondary"],
            corner_radius=10
        )
        self.search_entry.pack(side="left", padx=10, pady=8)
        self.search_entry.bind("<KeyRelease>", lambda e: self.apply_filter())
        
        # تفعيل اختصار التنقل بالأسهم للأسفل
        self.search_entry.bind("<Down>", self.move_focus_to_list)
        

        # فلتر السنة
        self.year_filter = ctk.CTkComboBox(
            self.filter_frame, 
            values=[lang["all_years"], "2030", "2029", "2028", "2027", "2026", "2025", "2024", "2023", "2022", "2021", "2020"],
            width=155, height=38,
            font=ctk.CTkFont(size=12),
            fg_color=COLORS["input_bg"],
            border_color=COLORS["input_border"],
            button_color=COLORS["btn_secondary"],
            button_hover_color=COLORS["accent_primary"],
            dropdown_fg_color=COLORS["bg_header"],
            dropdown_hover_color=COLORS["bg_card"],
            text_color=COLORS["text_primary"],
            corner_radius=10,
            command=lambda x: self.apply_filter()
        )
        self.year_filter.pack(side="left", padx=8, pady=8)

        # فلتر النوع
        self.type_filter = ctk.CTkComboBox(
            self.filter_frame, 
            values=[lang["all_types"], "PDF", "JPG", "PNG", "XLSX", "DOCX"],
            width=155, height=38,
            font=ctk.CTkFont(size=12),
            fg_color=COLORS["input_bg"],
            border_color=COLORS["input_border"],
            button_color=COLORS["btn_secondary"],
            button_hover_color=COLORS["accent_primary"],
            dropdown_fg_color=COLORS["bg_header"],
            dropdown_hover_color=COLORS["bg_card"],
            text_color=COLORS["text_primary"],
            corner_radius=10,
            command=lambda x: self.apply_filter()
        )
        self.type_filter.pack(side="left", padx=8, pady=8)
        
        # فلتر الزبائن
        
        # --- فلتر الزبائن بتنسيق مطابق ---
        self.customer_filter = ctk.CTkComboBox(
            self.filter_frame, 
            values=self.load_customers_from_file(),
            command=lambda x: self.apply_filter(),
            width=155, height=38,              # عرض مناسب
            corner_radius=10,       # حواف دائرية مثل الصورة
            border_width=2,
            fg_color="#0f111a",     # لون الخلفية الداكن (تأكد أنه مطابق لكودك)
            border_color="#24262f", # لون الحدود
            button_color="#24262f", # لون السهم الصغير
            button_hover_color="#33363f",
            dropdown_fg_color="#0f111a",
            dropdown_hover_color="#24262f",
            font=("Cairo", 12)      # نفس الخط المستخدم
        )
        self.customer_filter.pack(side="left", padx=5)
        # 🔥 ضع هذه الأسطر الثلاثة مباشرة بعد إنشاء الفلاتر في دالة setup_ui
        self.year_filter.configure(state="readonly")
        self.type_filter.configure(state="readonly")
        self.customer_filter.configure(state="readonly")

        # زر الإضافة (+) بنفس التنسيق
        self.add_customer_btn = ctk.CTkButton(
            self.filter_frame, 
            text="+", 
            width=35,
            height=35,
            corner_radius=10,
            fg_color="#2eb85c",     # اللون الأخضر الموجود في الصورة 
            hover_color="#1e7e34",
            text_color="white",
            font=("Cairo", 15, "bold"),
            command=self.add_new_customer_manually
        )
        self.add_customer_btn.pack(side="left", padx=2)
        
        # زر الحذف (-)
        self.del_customer_btn = ctk.CTkButton(
            self.filter_frame, 
            text="-", 
            width=35,
            height=35,
            corner_radius=10,
            fg_color="#c0392b",     # لون أحمر للتنبيه
            hover_color="#e74c3c",
            text_color="white",
            font=("Cairo", 15, "bold"),
            command=self.delete_selected_customer
        )
        self.del_customer_btn.pack(side="left", padx=2)
        #--------------------------------------------------------

        # ==========================================
        # 📋 إطار الجدول مع حدود مستديرة - محسّن
        # ==========================================
        self.table_container = ctk.CTkFrame(
            self, 
            fg_color=COLORS["bg_main"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=12
        )
        self.table_container.pack(expand=True, fill="both", padx=20, pady=(5, 10))

        # الجدول
        self.tree = ttk.Treeview(
            self.table_container, 
            columns=("name", "date", "size", "type"), 
            show='headings', 
            selectmode='extended',
            style="Custom.Treeview"
        )
        
        # 🟢 ربط عناوين الأعمدة بدالة الترتيب الذكي (كل سطر منفصل على حدة)
        self.tree.heading("name", text=lang["col_name"], command=lambda: self.sort_treeview_column("name", False))
        self.tree.heading("date", text=lang["col_date"], command=lambda: self.sort_treeview_column("date", False))
        self.tree.heading("size", text=lang["col_size"], command=lambda: self.sort_treeview_column("size", False))
        self.tree.heading("type", text=lang["col_type"], command=lambda: self.sort_treeview_column("type", False))
        
        # 🔥 ضع هذا السطر السحري هنا مباشرة تحت إنشاء الجدول:
        self.tree.bind("<Double-1>", self.open_file_on_double_click)
        
        # شريط التمرير
        scrollbar = ctk.CTkScrollbar(
            self.table_container, command=self.tree.yview,
            button_color=COLORS["btn_secondary"],
            button_hover_color=COLORS["accent_primary"]
        )
        scrollbar.pack(side="right", fill="y", padx=(0, 3), pady=3)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(expand=True, fill="both", padx=3, pady=3)

        # تعيين عرض الأعمدة
        self.tree.column("name", width=420, minwidth=200)
        self.tree.column("date", width=140, minwidth=100)
        self.tree.column("size", width=100, minwidth=80)
        self.tree.column("type", width=90, minwidth=70)
        
        # ربط لوحة المفاتيح
        # الحل النهائي والقاطع لتحديد الكل باللغتين (العربية والإنجليزية) بناءً على قراءة جهازك
        self.tree.bind("<Key>", lambda e: self.select_all(e) if (e.keysym.lower() == 'a' or e.keycode in [65, 81] or e.char == '\x11') and (e.state & 0x0004) else None)
        self.tree.bind("<Control-A>", self.select_all)
        self.tree.bind("<Return>", lambda e: self.open_file())
        self.tree.bind("<Delete>", lambda e: self.delete_file())
        self.tree.bind("<F9>", lambda e: self.rename_file())
        self.bind("<F1>", lambda event: self.add_file())
        self.bind("<F2>", lambda event: self.scan_document())
        self.bind("<F3>", lambda event: self.create_backup())
        self.bind("<F4>", lambda event: self.import_backup_folder())
        self.bind("<F5>", lambda event: self.open_settings())
        # 🔥 ربط زر Escape (Echap) بإعادة ضبط البرنامج بالكامل إلى الوضع الافتراضي
        self.bind("<Escape>", lambda event: self.reset_program_to_default())
        
        
    #------------------------------------------------------------------
        # متغيرات تتبع الملف المحدد حالياً لتشغيل الاختصارات
        self.selected_file_item = None
        self.selected_card_widget = None
                
        # ربط اختصار لوحة التحكم Ctrl+C على مستوى النافذة الرئيسية للبرنامج
        self.bind("<Control-c>", self.on_ctrl_c_pressed)
        self.bind("<Control-C>", self.on_ctrl_c_pressed)
        
    #------------------------------------------------------------------    
        
        
        
        self.update_tree_headings()
        
        # دالة ذكية لإلغاء التحديد عند الضغط في أي فراغ (بما في ذلك فراغ الجدول نفسه)
        def clear_selection_on_empty_click(event):
            # إذا كانت الضغطة داخل جدول الملفات
            if "treeview" in str(event.widget).lower():
                # تحديد المنطقة المضيئة تحت الماوس (cell تعني سطر ملف فعلي)
                region = event.widget.identify_region(event.x, event.y)
                if region not in ["cell", "tree"]:
                    self.tree.selection_remove(self.tree.selection())
            else:
                # إذا كانت الضغطة في أي مكان فارغ آخر خارج الجدول بالبرنامج
                self.tree.selection_remove(self.tree.selection())

        # ربط الحدث بـ يسار الفأرة في البرنامج بالكامل ليعمل بسلاسة مطلقة
        self.bind_all("<Button-1>", clear_selection_on_empty_click, add="+")
        
        # ==========================================
        # 📋 قائمة السياق (Context Menu) بتصميم محسّن
        # ==========================================
        self.context_menu = tk.Menu(
            self, tearoff=0,
            bg=COLORS["bg_header"],
            fg=COLORS["text_primary"],
            activebackground=COLORS["accent_primary"],
            activeforeground="#ffffff",
            font=("Segoe UI", 11),
            borderwidth=1,
            relief="flat"
        )
        # الكود الصحيح والمغلق تماماً بدون أي أخطاء سنتكس:
        self.context_menu.add_command(label="📂 " + lang.get("open", "فتح"), command=lambda: (self.context_menu.unpost(), self.update(), self.open_file()))
        self.context_menu.add_command(label="📝" + lang["btn_edit"], command=self.rename_file)
        
        # خيار الجيميل الجديد مضاف ومغلق بشكل سليم 100%
        # إضافة خيار الإرسال عبر Gmail في القائمة الجانبية
        self.context_menu.add_command(label="📧 " + lang.get("send_gmail", "ارسال عبر الجيميل"), command=lambda: (self.context_menu.unpost(), self.update(), self.share_file_via_gmail()))
        
        self.context_menu.add_command(label="📋 نسخ الملف (Ctrl+C)", command=self.on_ctrl_c_pressed)
      
        
        self.context_menu.add_command(label="🗑" + lang["btn_delete"], command=self.delete_file)
        
        # ربط كليك يمين للجدول بالدالة الذكية
        self.tree.bind("<Button-3>", self.show_smart_context_menu)
        
        # ربط زر App في الكيبورد بالدالة الذكية مباشرة
        self.tree.bind("<App>", self.show_smart_context_menu)

        # ==========================================
        # 🔻 الشريط السفلي (Footer) - محسّن
        # ==========================================
        # خط فاصل ملون فوق الفوتر
        self.footer_accent_line = ctk.CTkFrame(
            self, height=1, corner_radius=0,
            fg_color=COLORS["border"]
        )
        self.footer_accent_line.pack(side="bottom", fill="x")
        
        self.footer_frame = ctk.CTkFrame(
            self, height=38, corner_radius=0, 
            fg_color=COLORS["bg_footer"],
            border_width=0
        )
        self.footer_frame.pack(side="bottom", fill="x")
        self.footer_frame.pack_propagate(False)
        
        # لوقو صغير في الفوتر
        self.footer_logo = ctk.CTkLabel(
            self.footer_frame, text="⚡",
            font=ctk.CTkFont(size=13),
        )
        self.footer_logo.pack(side="left", padx=(16, 4))
        
        self.copy_label = ctk.CTkLabel(
            self.footer_frame, text=lang["copyright"],
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_muted"]
        )
        self.copy_label.pack(side="left", padx=0)
        
        self.clock_label = ctk.CTkLabel(
            self.footer_frame, text="",
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color=COLORS["text_secondary"]
        )
        self.clock_label.pack(side="right", padx=16)

        # فاصل نقطي
        ctk.CTkLabel(
            self.footer_frame, text="•",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["text_muted"]
        ).pack(side="right", padx=4)

        self.stats_label = ctk.CTkLabel(
            self.footer_frame, text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=COLORS["accent_primary"]
        )
        self.stats_label.pack(side="right", padx=4)

    def select_all(self, event=None):
        self.tree.selection_set(self.tree.get_children())
        return "break"

    def show_context_menu(self, event):
        # 1. جلب الملفات المحددة حالياً في الجدول
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        count = len(selected_items)
        
        # 2. تحديد حالة الخيارات والألوان بصرياً (إطفاء الخيارات في التحديد المتعدد)
        if count > 1:
            target_state = "disabled"
            text_color = "#555555"          # لون رمادي باهت
        else:
            target_state = "normal"
            text_color = COLORS["text_primary"] # اللون الأبيض الأصلي

        # 3. تطبيق الحالة واللون على خيار الفتح وتعديل الاسم
        try:
            self.context_menu.entryconfigure(0, state=target_state, foreground=text_color)
            self.context_menu.entryconfigure(1, state=target_state, foreground=text_color)
        except Exception as e:
            print(f"Error updating menu states: {e}")

        # 4. التمييز الذكي بين ضغطة الماوس وضغطة زر App في الكيبورد
        # إذا كانت الإحداثيات (event.x و event.y) تساوي 0 أو 1، فهذا يعني أن الضغطة قادمة من زر App وليس الماوس
        if hasattr(event, 'x') and event.x <= 1 and event.y <= 1:
            # تم الضغط على زر App -> الحساب بمحاذاة آخر حرف من الاسم
            item = selected_items[0]
            bbox = self.tree.bbox(item)  # جلب أبعاد السطر المحدد
            
            if bbox:
                # جلب اسم الملف المحدد وقياس طوله التقريبي بالبكسل
                item_text = str(self.tree.item(item)['values'][0]).strip()
                text_width_pixels = len(item_text) * 8  
                
                # جلب العرض الكامل لعمود الاسم
                column_width = self.tree.column("#1", width=None)
                
                # حساب مكان آخر حرف (العرض الكلي للعمود - طول النص) مع ترك مسافة 30 بكسل للتناسق
                x_offset = column_width - text_width_pixels - 30
                if x_offset < 10: 
                    x_offset = 10
                
                x_root = self.tree.winfo_rootx() + bbox[0] + x_offset
                y_root = self.tree.winfo_rooty() + bbox[1] + (bbox[3] // 2)
            else:
                x_root = self.tree.winfo_rootx() + 50
                y_root = self.tree.winfo_rooty() + 50
        else:
            # تم الضغط كليك يمين بالماوس -> تظهر عند رأس السهم طبيعياً
            x_root = event.x_root
            y_root = event.y_root

        # 5. إظهار القائمة على الشاشة
        try:
            self.context_menu.post(x_root, y_root)
        except Exception:
            pass
            
        # دالة حساب حجم الملف بشكل مقروء واحترافي
        def get_file_size_formatted(paths):
            if not paths:
                return "0 KB"
            total_bytes = sum(os.path.getsize(p) for p in paths if os.path.exists(p))
            if total_bytes < 1024 * 1024:
                return f"{total_bytes / 1024:.2f} KB"
            return f"{total_bytes / (1024 * 1024):.2f} MB"
            
               
            
            
    #------------------------------------------------------------
    
    def on_ctrl_c_pressed(self, event=None):
        """التقاط حدث النسخ وفحص المجلد الحقيقي لنسخ الملف بأمان وبالمسار الصحيح"""
        try:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("تنبيه", "يرجى تحديد ملف من الجدول أولاً لنسخه!", parent=self)
                return

            # 1. جلب القيم من الجدول
            vals = self.tree.item(selected_items[0])['values']
            raw_name = str(vals[0]).strip()
            file_type = str(vals[3]).strip().lower().replace('.', '')

            # 2. تنظيف الرموز التعبيرية من الاسم
            clean_name = raw_name
            if len(raw_name) > 0 and ord(raw_name[0]) > 10000:
                clean_name = raw_name[1:].strip()

            # 3. حل مشكلة المسار: البحث الذكي داخل مجلد الأرشيف الفعلي
            file_path = None
            filename_to_copy = f"{clean_name}.{file_type}"
            
            if os.path.exists(self.archive_dir):
                for f in os.listdir(self.archive_dir):
                    if f.lower() == f"{clean_name}.{file_type}".lower() or f.lower() == f"{raw_name}.{file_type}".lower():
                        file_path = os.path.join(self.archive_dir, f)
                        filename_to_copy = f
                        break
                    elif clean_name.lower() in f.lower() and f.lower().endswith(file_type):
                        file_path = os.path.join(self.archive_dir, f)
                        filename_to_copy = f
                        break

            if not file_path:
                file_path = os.path.normpath(os.path.join(self.archive_dir, f"{clean_name}.{file_type}"))

            # 4. التحقق النهائي من وجود الملف على القرص
            if not os.path.exists(file_path):
                messagebox.showerror(
                    "الملف غير موجود", 
                    f"تعذر العثور على الملف الحقيقي في مجلد الأرشيف!\n\n"
                    f"الاسم المبحوث عنه: {filename_to_copy}\n"
                    f"المسار الحالي للبرنامج:\n{file_path}", 
                    parent=self
                )
                return
                
            # 5. النسخ إلى حافظة ويندوز
            import win32clipboard
            import ctypes
            from ctypes import wintypes
            
            class DROPFILES(ctypes.Structure):
                _fields_ = [
                    ("pFiles", wintypes.DWORD),
                    ("pt", wintypes.POINT),
                    ("fNC", wintypes.BOOL),
                    ("fWide", wintypes.BOOL),
                ]
            
            files = file_path + "\0\0"
            allocated_mem = ctypes.create_string_buffer(files.encode('utf-16-le'))
            
            dropfiles = DROPFILES()
            dropfiles.pFiles = ctypes.sizeof(DROPFILES)
            dropfiles.fWide = True
            
            final_clipboard_data = bytes(dropfiles) + allocated_mem.raw
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, final_clipboard_data)
            finally:
                win32clipboard.CloseClipboard()
                
            messagebox.showinfo("حافظة ويندوز", f"تم نسخ الملف '{filename_to_copy}' بنجاح!\nيمكنك الآن لصقه (Ctrl+V) في أي مكان على جهازك.", parent=self)
            
        except Exception as err:
            messagebox.showerror("خطأ في النسخ", f"تعذر إرسال الملف إلى الحافظة بسبب:\n{err}", parent=self)
            
    #-------------------------------------------------------------------------------------  

    def select_card_visual(self, file_item, card_widget):
        """تمكين تحديد الكرت بصرياً وتجهيز مراجع الملف لعملية النسخ الفوري بـ Ctrl+C"""
        # إعادة الكرت السابق إلى وضعه وحدوده الطبيعية
        if self.selected_card_widget and self.selected_card_widget.winfo_exists():
            self.selected_card_widget.configure(border_color=COLORS["border_color"], border_width=1)
            
        self.selected_file_item = file_item
        self.selected_card_widget = card_widget
        
        # تلوين حدود الكرت الحالي بلون الأكسنت لتأكيد جهوزية النسخ
        card_widget.configure(border_color=COLORS["accent_primary"], border_width=2)

    #-------------------------------------------------------------------------------------     

        # دالة الإرسال الفعلي الخلفية (المرحلة 3)
    def share_file_via_gmail(self):
        """
        دالة مطورة تتيح إرسال عدة ملفات محددة معاً عبر الجيميل، 
        وتظهر النافذة متمركزة في وسط البرنامج، مع شريط تمرير للمعلومات المنظمة من اليسار لليومين،
        وشريط تقدم حقيقي ينمو تدريجياً مع عداد نسبة مئوية %.
        """
        try:
            # 1. جلب كافة الملفات المحدّدة حالياً من الجدول (يدعم اختيار متعدد)
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("تنبيه", "يرجى تحديد ملف أو عدة ملفات من الجدول أولاً لإرسالها!")
                return
            
            file_paths = []
            invalid_files = []

            for sel in selected_items:
                vals = self.tree.item(sel)['values']
                raw_name = str(vals[0])
                clean_name = raw_name
                
                # تنظيف اسم الملف من أي أيقونات مضافة بناءً على منطق برنامجك
                for info in FILE_TYPE_COLORS.values():
                    if clean_name.startswith(info["icon"]):
                        clean_name = clean_name[len(info["icon"]):].strip()
                        break
                if clean_name.startswith(DEFAULT_FILE_COLOR["icon"]):
                    clean_name = clean_name[len(DEFAULT_FILE_COLOR["icon"]):].strip()
                
                file_type = str(vals[3]).strip().lower()
                filename = f"{clean_name}.{file_type}"
                full_path = os.path.join(self.archive_dir, filename)
                
                # التأكد من وجود الملف الفعلي في الأرشيف
                if os.path.exists(full_path):
                    file_paths.append(full_path)
                else:
                    invalid_files.append(filename)

            if invalid_files:
                messagebox.showerror("خطأ", f"بعض الملفات المحددة غير موجودة في الأرشيف فعلياً:\n" + "\n".join(invalid_files))
                if not file_paths:
                    return

            # 2. إنشاء النافذة الفرعية المخصصة للإرسال بنمط CustomTkinter المظلم
            email_window = ctk.CTkToplevel(self)
            email_window.geometry("460x360")
            email_window.resizable(False, False)
            email_window.configure(fg_color=COLORS["bg_main"])
            email_window.transient(self)
            email_window.grab_set()

            # 🌟 أولاً: حسابات متطورة لتوسيط النافذة في منتصف البرنامج تماماً
            email_window.update_idletasks()
            px = self.winfo_x() + (self.winfo_width() // 2) - 230
            py = self.winfo_y() + (self.winfo_height() // 2) - 180
            email_window.geometry(f"+{px}+{py}")

            # دالة مساعدة لحساب حجم الملفات الإجمالي وتنسيقه
            def get_local_file_size(paths):
                try:
                    total_bytes = sum(os.path.getsize(p) for p in paths if os.path.exists(p))
                    if total_bytes < 1024 * 1024:
                        return f"{total_bytes / 1024:.1f} KB"
                    return f"{total_bytes / (1024 * 1024):.1f} MB"
                except:
                    return "غير معروف"

            # 3. دالة معالجة الإرسال الفعلي تدريجياً في الخلفية مع النسبة المئوية
            def send_email_process(receiver_email):
                # تنظيف النافذة تماماً لإظهار واجهة التحميل التدريجي
                for widget in email_window.winfo_children():
                    widget.pack_forget()
                
                # نص حالة الإرسال الحالي
                lbl_status = ctk.CTkLabel(email_window, text="بالسيرفر الاتصال جاري...", text_color=COLORS["text_primary"], font=ctk.CTkFont(size=14, weight="bold"))
                lbl_status.pack(pady=(40, 10))
                
                # 🌟 شريط تحميل تدريجي ينمو شيئاً فشيئاً ولا يتكرر بشكل وهمي
                progress_bar = ctk.CTkProgressBar(email_window, width=360, progress_color="#2ed573", fg_color=COLORS["bg_card"])
                progress_bar.pack(pady=10)
                progress_bar.set(0.0)
                
                # 🌟 عداد النسبة المئوية في الوسط أسفل شريط التحميل
                lbl_percentage = ctk.CTkLabel(email_window, text="0%", text_color=COLORS["accent_primary"], font=ctk.CTkFont(size=15, weight="bold"))
                lbl_percentage.pack(pady=5)

                def update_progress(val, status_text):
                    progress_bar.set(val)
                    lbl_percentage.configure(text=f"{int(val * 100)}%")
                    lbl_status.configure(text=status_text)
                    email_window.update()

                import smtplib
                from email.mime.multipart import MIMEMultipart
                from email.mime.text import MIMEText
                from email.mime.base import MIMEBase
                from email import encoders
                from email.header import Header

                try:
                    # الخطوة 1: تجهيز الرسالة (15%)
                    update_progress(0.15, "جاري إعداد الرسالة والملفات...")
                    
                    sender_email = "abcdmkk920@gmail.com"
                    sender_password = "sjupfdaeltmcabjg"

                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = receiver_email
                    msg['Subject'] = f"ملفات مؤرشفة مرسلة من برنامج SwiftFolderPro ({len(file_paths)} ملف)"
                    
                    body = f"مرحباً،\n\nبرفقته الرسالة تجدون الملفات المؤرشفة المرسلة من برنامج SwiftFolderPro بحجمها الكامل.\nعدد الملفات: {len(file_paths)}\n\nتحياتي."
                    msg.attach(MIMEText(body, 'plain', 'utf-8'))

                    # الخطوة 2: إرفاق الملفات تدريجياً (من 15% إلى 50%)
                    start_perc = 0.15
                    end_perc = 0.50
                    for idx, path in enumerate(file_paths):
                        filename_base = os.path.basename(path)
                        update_progress(start_perc + ((idx / len(file_paths)) * (end_perc - start_perc)), f"جاري إرفاق: {filename_base}")
                        try:
                            with open(path, "rb") as attachment:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(attachment.read())
                                encoders.encode_base64(part)
                                part.add_header("Content-Disposition", f"attachment; filename={Header(filename_base, 'utf-8').encode()}")
                                msg.attach(part)
                        except Exception as file_err:
                            print(f"فشل إرفاق الملف {path}: {file_err}")

                    # الخطوة 3: فتح اتصال آمن (70%)
                    update_progress(0.70, "جاري فتح اتصال آمن مع SMTP...")
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    
                    # الخطوة 4: تسجيل الدخول الآمن (85%)
                    update_progress(0.85, "جاري تسجيل الدخول إلى الحساب...")
                    server.login(sender_email, sender_password)
                    
                    # الخطوة 5: تسليم البريد النهائي (95%)
                    update_progress(0.95, "جاري دفع البيانات وإرسال البريد...")
                    text = msg.as_string()
                    server.sendmail(sender_email, receiver_email, text)
                    server.quit()
                    
                    # 🌟 اكتمال الإرسال 100% بنجاح تام ويغلق التحميل فوراً للذهاب لنافذة النجاح
                    update_progress(1.0, "بنجاح العملية تمت")
                    time.sleep(0.4)

                    # تنظيف الواجهة لإظهار نجاح الإرسال المبهج كما طلبت بمجرد الوصول لـ 100%
                    for widget in email_window.winfo_children():
                        widget.pack_forget()
                    
                    email_window.unbind("<Return>")
                    
                    lbl_icon = ctk.CTkLabel(email_window, text="✅", font=ctk.CTkFont(size=45), text_color="#2ed573")
                    lbl_icon.pack(pady=(40, 10))

                    lbl_success = ctk.CTkLabel(email_window, text="بنجاح الإرسال تم", text_color=COLORS["text_primary"], font=ctk.CTkFont(size=16, weight="bold"))
                    lbl_success.pack(pady=10)

                    btn_ok = ctk.CTkButton(email_window, text="موافق", width=120, height=35, fg_color=COLORS["accent_primary"], hover_color=COLORS["accent_hover"], command=email_window.destroy)
                    btn_ok.pack(pady=20)
                    
                    email_window.bind("<Return>", lambda event: email_window.destroy())
                    email_window.bind("<Escape>", lambda event: email_window.destroy())
                    
                except Exception as e:
                    messagebox.showerror("خطأ في الإرسال", f"فشل الاتصال أو الإرسال آلياً:\n{e}", parent=email_window)
                    show_review_stage(receiver_email)

            # 4. دالة واجهة مراجعة البيانات (المرحلة 2) مع صندوق التمرير والترتيب من اليسار لليمين
            def show_review_stage(receiver_email):
                for widget in email_window.winfo_children():
                    widget.pack_forget()
                    
                email_window.title("مراجعة بيانات الإرسال")
                
                lbl_rev_title = ctk.CTkLabel(email_window, text="📋 البيانات و تأكيد مراجعة", text_color=COLORS["text_primary"], font=ctk.CTkFont(size=16, weight="bold"))
                lbl_rev_title.pack(pady=(10, 5))
                
                # 🌟 إضافة شريط تمرير من الأعلى إلى الأسفل في حال كانت معلومات الملفات كثيرة لمنع كراش الواجهة
                scroll_frame = ctk.CTkScrollableFrame(email_window, width=410, height=200, fg_color=COLORS["bg_card"], corner_radius=10)
                scroll_frame.pack(pady=10, padx=15, fill="both", expand=True)
                
                # تجهيز قائمة أسماء الملفات (اسم تحت اسم) مصفوفة عمودياً
                formatted_file_names = "\n".join([f"   • {os.path.basename(p)}" for p in file_paths])
                total_size = get_local_file_size(file_paths)
                
                # 🌟 صياغة نص المعلومات مرتبة من اليسار إلى اليمين سطر تحت سطر بالكامل منظم ومحاذى لليسار
                review_text = (
                    f"To (Receiver Gmail):\n{receiver_email}\n\n"
                    f"Attached Files:\n{formatted_file_names}\n\n"
                    f"Total Files Size:\n{total_size}"
                )
                
                lbl_info = ctk.CTkLabel(
                    scroll_frame, 
                    text=review_text,
                    justify="left",          # 🌟 محاذاة النص من اليسار إلى اليمين بامتياز
                    anchor="w",
                    text_color=COLORS["text_primary"], 
                    font=ctk.CTkFont(family="Consolas", size=12) # خط ثابت يعطي تنظيم مريح للعين
                )
                lbl_info.pack(fill="x", padx=10, pady=10)
                
                # إطار سفلي لترتيب أزرار التنقل
                btn_frame = ctk.CTkFrame(email_window, fg_color="transparent")
                btn_frame.pack(side="bottom", pady=15, fill="x")
                
                btn_back = ctk.CTkButton(btn_frame, text="السابق", width=110, height=32, fg_color="#718093", hover_color="#57606f", command=lambda: show_input_stage(receiver_email))
                btn_back.pack(side="left", padx=25)
                
                btn_final_send = ctk.CTkButton(
                    btn_frame, text="إرسال الآن", width=110, height=32,
                    fg_color=COLORS["accent_primary"], hover_color=COLORS["accent_hover"],
                    command=lambda: threading.Thread(target=send_email_process, args=(receiver_email,), daemon=True).start()
                )
                btn_final_send.pack(side="right", padx=25)
                
                # 🌟 [تعديل التنقل المطور الجديد] 🌟
                btn_final_send.focus_set()  # جعل زر الإرسال هو المحدّد افتراضياً عند فتح الشاشة
                
                # دالة التحكم بالأسهم لتبديل التركيز (Focus)
                def navigate_buttons(event):
                    if event.keysym == "Right":
                        btn_final_send.focus_set()
                    elif event.keysym == "Left":
                        btn_back.focus_set()

                # ربط أزرار الأسهم بالنافذة
                email_window.bind("<Right>", navigate_buttons)
                email_window.bind("<Left>", navigate_buttons)
                email_window.bind("<Escape>", lambda event: email_window.destroy())

                # جعل مفتاح Enter ذكي يضغط على الزر النشط الذي يقف عليه المستخدم حالياً بالأسهم
                email_window.unbind("<Return>")
                email_window.bind("<Return>", lambda event: email_window.focus_get().invoke() if hasattr(email_window.focus_get(), 'invoke') else None)

            # 5. دالة واجهة إدخال البريد الإلكتروني الأولى (المرحلة 1)
            def show_input_stage(default_text=""):
                for widget in email_window.winfo_children():
                    widget.pack_forget()
                    
                email_window.title("Gmail")
                
                lbl_title_input = ctk.CTkLabel(email_window, text="إلكتروني البريد إدخال الرجاء", text_color=COLORS["text_primary"], font=ctk.CTkFont(size=16, weight="bold"))
                lbl_title_input.pack(pady=(40, 15))
                
                entry_receiver_input = ctk.CTkEntry(email_window, placeholder_text="للمستلم الإلكتروني البريد إدخال الرجاء...", width=340, height=38, corner_radius=10)
                entry_receiver_input.pack(pady=10)
                
                if default_text:
                    entry_receiver_input.insert(0, default_text)
                
                def go_to_review():
                    receiver_email = entry_receiver_input.get().strip()
                    if not receiver_email or "@" not in receiver_email:
                        messagebox.showerror("خطأ", "ومكتمل! صحيح إلكتروني بريد إدخال الرجاء", parent=email_window)
                        return
                    show_review_stage(receiver_email)
                
                btn_next = ctk.CTkButton(
                    email_window, text="التالي", width=200, height=35,
                    fg_color=COLORS["accent_primary"], hover_color=COLORS["accent_hover"],
                    corner_radius=10,
                    command=go_to_review
                )
                btn_next.pack(pady=20)
                
                email_window.after(100, lambda: entry_receiver_input.focus_force())
                email_window.bind("<Return>", lambda event: go_to_review())
                email_window.bind("<Escape>", lambda event: email_window.destroy())

            # تشغيل شاشة الإدخال الأولى فوراً عند استدعاء الدالة
            show_input_stage()
            
        except Exception as e:
            print(f"Error in share_file_via_gmail: {e}")

    

    
    

        # دالة الإرسال الفعلي الخلفية (المرحلة 3)
        

    def update_tree_headings(self):
        lang = self.languages[self.current_lang]
        for col in ("name", "date", "size", "type"):
            self.tree.heading(col, text=lang[f"col_{col}"])

    def update_file_stats(self):
        lang = self.languages[self.current_lang]
        count = len(self.tree.get_children())
        self.stats_label.configure(text=f"{lang['file_count']} {count}")

    def load_initial_archive(self):
        self.all_files_data = []
        if os.path.exists(self.archive_dir):
            for filename in os.listdir(self.archive_dir):
                path = os.path.join(self.archive_dir, filename)
                if os.path.isfile(path):
                    stats = os.stat(path)
                    name_only, extension = os.path.splitext(filename)
                    
                    file_info = {
                        "path": path,
                        "name": name_only,
                        "date": datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d"),
                        "size": f"{round(stats.st_size / (1024 * 1024), 2)} MB",
                        "type": extension.replace('.', '').upper()
                    }
                    self.all_files_data.append(file_info)
        
        self.all_files_data.sort(key=lambda x: os.path.getctime(x["path"]), reverse=True)
        self.apply_filter()
        self.update_customer_list()
        
    def update_sidebar_buttons(self, event=None):
        # 1. جلب قائمة الملفات المحددة حالياً
        selected_items = self.tree.selection()
        count = len(selected_items)
        
        # 2. إذا كان التحديد يحتوي على أكثر من ملف (ملفات متعددة)
        if count > 1:
            # جعل زر الفتح والتعديل باهت (مطفي) وغير قابل للضغط - مع الحماية
            try:
                self.btn_open.configure(state="disabled")
                self.btn_edit.configure(state="disabled")
            except Exception:
                pass
        else:
            # إعادة الأزرار لطبيعتها الناصعة والنشطة إذا كان ملف واحد أو صفر - مع الحماية
            try:
                self.btn_open.configure(state="normal")
                self.btn_edit.configure(state="normal") 
            except Exception:
                pass
        
    def trigger_menu_from_keyboard(self, event):
        # 1. التأكد من أن المستخدم حدد ملفاً في الجدول
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        # 2. الحصول على معرف الملف المحدد وحساب موقعه الفعلي داخل الجدول
        item = selected_items[0]
        bbox = self.tree.bbox(item)
        
        try:
            if bbox:
                # إذا نجح في حساب موقع الملف، نولد الضغطة في منتصف السطر المحدّد تماماً
                x = bbox[0] + 50  # إزاحة بسيطة لداخل السطر
                y = bbox[1] + (bbox[3] // 2)
            else:
                # إحداثيات افتراضية آمنة داخل الجدول في حال كان الجدول فارغاً أو تحت التحديث
                x = 100
                y = 30
                
            # توليد حدث كليك يمين حقيقي عند إحداثيات الملف المحدّد (وليس مكان الماوس!)
            self.tree.event_generate("<Button-3>", x=x, y=y)
        except Exception as e:
            print(f"Error triggering menu: {e}")
            
    #------------------------------------------------------------------------------------------        
            
    def show_smart_context_menu(self, event):
        """دالة تجعل القائمة تظهر بجانب الملف المحدد بالأسهم بالضبط وتتحرك معه أينما ذهب"""
        try:
            # 1. تحديث أبعاد الواجهة فوراً
            self.update_idletasks()

            # 2. إذا كان كليك يمين بالماوس
            if event.type == '4' or hasattr(event, 'x'):
                x_root, y_root = event.x_root, event.y_root
                clicked_item = self.tree.identify_row(event.y)
            
            # 3. 🌟 إذا ضغطت زر App (التحرك بالأسهم)
            else:
                selection = self.tree.selection()
                clicked_item = selection[0] if selection else self.tree.focus()
                
                if clicked_item:
                    # جلب إحداثيات السطر الحالي
                    bbox = self.tree.bbox(clicked_item)
                    if bbox:
                        # حساب موقع الجدول الفعلي على الشاشة
                        tree_x = self.tree.winfo_rootx()
                        tree_y = self.tree.winfo_rooty()
                        
                        # وضع القائمة بجانب الملف المحدد بالأسهم مباشرة
                        x_root = tree_x + bbox[0] + 60
                        y_root = tree_y + bbox[1] + int(bbox[3] / 2)
                    else:
                        # حماية في حال عدم قراءة الأبعاد
                        x_root = self.tree.winfo_rootx() + 100
                        y_root = self.tree.winfo_rooty() + 100
                else:
                    x_root = self.winfo_rootx() + int(self.winfo_width() / 2)
                    y_root = self.winfo_rooty() + int(self.winfo_height() / 2)

            # 4. إظهار القائمة
            if clicked_item and clicked_item in self.tree.selection():
                self.context_menu.post(x_root, y_root)
            else:
                if hasattr(self, 'empty_space_menu'):
                    self.empty_space_menu.post(x_root, y_root)
                elif hasattr(self, 'empty_menu'):
                    self.empty_menu.post(x_root, y_root)
                else:
                    self.context_menu.post(x_root, y_root)
                
        except Exception as e:
            print(f"Error showing context menu: {e}")
            
            
    #--------------------------------------------------------------------------------------
           
            
    def reset_program_to_default(self, event=None):
        """إعادة شريط البحث، الفلاتر الثلاثة، وتحديدات الجدول إلى الوضع الافتراضي"""
        lang = self.languages[self.current_lang]
        
        # 1. تفريغ حقل البحث الأصلي تماماً
        self.search_entry.delete(0, 'end')
        
        # 2. إعادة فلتر السنة إلى الخيار الأول الافتراضي (كل السنوات)
        self.year_filter.set(lang["all_years"])
        
        # 3. إعادة فلتر النوع إلى الخيار الأول الافتراضي (كل الأنواع)
        self.type_filter.set(lang["all_types"])
        
        # 4. إعادة فلتر الزبائن إلى الخيار الأول الافتراضي
        customer_values = self.customer_filter.cget("values")
        if customer_values:
            self.customer_filter.set(customer_values[0]) # يختار أول زبون في القائمة (غالباً "كل الزبائن")
        
        # 5. إلغاء أي تحديدات زرقاء حالية على الملفات في الجدول
        self.tree.selection_remove(self.tree.selection())
        
        # 6. تحديث الجدول فوراً وإعادة عرض كافة الملفات بناءً على إزالة الفلاتر
        self.apply_filter()
        
        # 7. إزالة مؤشر الكتابة (Focus) من حقل البحث والجدول لإراحة العين
        self.focus_set()

    def add_file(self):
        file_paths = filedialog.askopenfilenames()
        if not file_paths:
            return

        try:
            lang = self.languages[self.current_lang]
            added_count = 0

            for file_path in file_paths:
                filename = os.path.basename(file_path)
                dest = os.path.join(self.archive_dir, filename)

                if os.path.exists(dest):
                    base, ext = os.path.splitext(filename)
                    i = 1
                    while os.path.exists(dest):
                        dest = os.path.join(self.archive_dir, f"{base} ({i}){ext}")
                        i += 1

                shutil.copy2(file_path, dest)
                added_count += 1

            self.load_initial_archive()
            self.update_file_stats()
            self.update_customer_list()

            s_title = lang.get("success_title", "")
            s_msg = "تم إضافة الملفات بنجاح!" if self.current_lang == "العربية" else "Files added successfully!"
            CustomMessageBox(self, title=s_title, message=f"{s_msg} ({added_count})", msg_type="success")

        except Exception as e:
            err_title = "خطأ" if self.current_lang == "العربية" else "Error"
            CustomMessageBox(self, title=err_title, message=str(e), msg_type="error")
            
    # 🔥 ضع هذه الدالة داخل الكلاس مع بقية دوال الأزرار
    def download_selected_files(self):
        import os
        import shutil
        from tkinter import filedialog, messagebox

        # 1. جلب كافة السطور المحددة من الجدول
        selected_items = self.tree.selection()
        if not selected_items:
            return

        # 2. 🔥 [التعديل الذكي]: فتح النافذة لتحديد مسار المجلد الأساسي مباشرة
        # استخدام الخيارات الإضافية يضمن للويندوز قبول "سطح المكتب" كموقع حفظ مباشر
        target_dir = filedialog.askdirectory(
            title="اختر مكان حفظ الملفات (يمكنك اختيار سطح المكتب مباشرة)",
            mustexist=True # يضمن قبول المسارات الرئيسية للنظام مثل سطح المكتب والبارتشنز
        )
        
        # إذا ألغى المستخدم العملية
        if not target_dir:
            return

        # تصحيح المسار ليفهمه الويندوز بشكل صحيح وسلس
        target_dir = os.path.normpath(target_dir)

        success_count = 0
        fail_count = 0

        # 3. المرور على الملفات ونسخها
        for item in selected_items:
            try:
                raw_file_name = self.tree.item(item, "values")[0]
                cleaned_name = raw_file_name.replace("📄", "").strip()
                
                actual_filename = None
                if os.path.exists(os.path.join(self.archive_dir, cleaned_name)):
                    actual_filename = cleaned_name
                else:
                    for f in os.listdir(self.archive_dir):
                        if f.startswith(cleaned_name) or cleaned_name in f:
                            actual_filename = f
                            break
                
                if actual_filename:
                    source_path = os.path.join(self.archive_dir, actual_filename)
                    destination_path = os.path.join(target_dir, actual_filename)
                    
                    # نسخ الملف الحقيقي
                    shutil.copy2(source_path, destination_path)
                    success_count += 1
                else:
                    fail_count += 1
                    
            except Exception as e:
                print(f"خطأ أثناء نسخ الملف: {e}")
                fail_count += 1

        # 4. رسالة النجاح
        if success_count > 0:
            msg = f"تم إنزال {success_count} ملف بنجاح! 👌"
            if fail_count > 0:
                msg += f"\n⚠️ فشل إنزال {fail_count} ملف."
            messagebox.showinfo("تم الإنزال والنسخ", msg, parent=self)    

    def open_file_on_double_click(self, event):
        import os
        try:
            # 1. تحديد السطر المحدّد في الجدول
            selected_item = self.tree.selection()
            if selected_item:
                # 2. جلب اسم الملف من العمود الأول
                raw_file_name = self.tree.item(selected_item, "values")[0]
                
                # 3. 🔥 [تنظيف الاسم]: إزالة إيموجي الورقة 📄 وأي مسافات مخفية
                cleaned_name = raw_file_name.replace("📄", "").strip()
                
                # 4. محاولة الفتح المباشر إذا كان الاسم يحتوي على الامتداد أصلاً
                full_path = os.path.join(self.archive_dir, cleaned_name)
                
                if os.path.exists(full_path):
                    os.startfile(full_path) # 🚀 فتح الملف فوراً
                else:
                    # 5. 💡 [ذكاء اصطناعي احتياطي]: إذا لم يجد الملف (بسبب غياب الامتداد .pdf مثلاً)
                    # يقوم بالبحث داخل مجلد الأرشيف عن أي ملف يبدأ بنفس الاسم المنظف
                    found = False
                    for actual_file in os.listdir(self.archive_dir):
                        if actual_file.startswith(cleaned_name) or cleaned_name in actual_file:
                            alt_path = os.path.join(self.archive_dir, actual_file)
                            os.startfile(alt_path) # 🚀 فتح الملف المطابق
                            found = True
                            break
                    
                    if not found:
                        print(f"الملف غير موجود في المسار: {full_path}")
                        
        except Exception as e:
            print(f"تعذر فتح الملف أثناء النقر المزدوج: {e}")

    def apply_filter(self, event=None):
        import re # تأكد من وجود هذه المكتبة في أعلى ملفك
        
        lang = self.languages[self.current_lang]
        
        # 1. القيم المختارة
        s_year = self.year_filter.get()
        s_type = self.type_filter.get().upper().replace('.', '')
        s_cust = self.customer_filter.get().strip()
        s_query = self.search_entry.get().lower().strip()
        all_opt = self.customer_filter.cget("values")[0]
        
        # مسح الجدول
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 2. حلقة الفلترة
        for file in self.all_files_data:
            f_name = file["name"].lower() # اسم الملف بالكامل
            f_type = file["type"].upper()
            
            # --- استخراج السنة من اسم الملف ---
            # نبحث عن أي 4 أرقام متتالية في اسم الملف (مثل 2026)
            year_match_in_name = re.search(r'\d{4}', f_name)
            file_year = year_match_in_name.group(0) if year_match_in_name else ""
            
            # أ. فلتر الزبون
            match_cust = (s_cust == all_opt or s_cust.lower() in f_name)
            
            # ب. فلتر السنة (الآن يعتمد على السنة المستخرجة من الاسم)
            if s_year == lang["all_years"] or not s_year:
                match_year = True
            else:
                match_year = (s_year.strip() == file_year)
            
            # ج. فلتر النوع
            match_type = (s_type == lang["all_types"].upper() or not s_type or s_type == f_type)
            
            # د. فلتر البحث
            match_search = (s_query in f_name)
            
            # 3. التحقق التراكمي
            if match_cust and match_year and match_type and match_search:
                type_info = get_file_type_info(file["type"])
                self.tree.insert("", "end", values=(
                    f"{type_info['icon']}  {file['name']}", 
                    file["date"], 
                    file["size"], 
                    file["type"]
                ))
        
        self.update_file_stats()
        
        
    # دالة ذكية لتحويل النصوص والأرقام إلى صيغة قابلة للترتيب الطبيعي (Natural Sort)
    def natural_sort_key(self, text):
        import re
        # تحويل النص إلى قائمة تحتوي على نصوص وأرقام منفصلة، مع تحويل الأرقام إلى نوع int لترتيبها بدقة
        return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

    # الدالة الأساسية لترتيب الجدول عند الضغط على العنوان
    def sort_treeview_column(self, col, reverse):
        # جلب جميع العناصر الحالية المعروضة في الجدول
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        
        # 🔥 تطبيق الترتيب الذكي (Natural Sort) بناءً على نوع العمود
        # إذا كان عمود الاسم، نطبق الفرز الطبيعي للأرقام والحروف والنسخ
        if col == "name":
            l.sort(key=lambda t: self.natural_sort_key(t[0]), reverse=reverse)
        else:
            # للأعمدة الأخرى مثل التاريخ أو النوع
            l.sort(key=lambda t: t[0].lower(), reverse=reverse)

        # إعادة ترتيب العناصر في واجهة الجدول بناءً على الترتيب الجديد
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # عكس اتجاه الترتيب للضغطة القادمة (من تصاعدي لتنازلي أو العكس)
        self.tree.heading(col, command=lambda: self.sort_treeview_column(col, not reverse))
        
    #-------------------------------------------------------------------------------------

    def apply_natural_sort_from_button(self, reverse_order=False):
        """دالة مستوحاة من كودك الذكي لترتيب جدول الأرشيف عبر الزر مباشرة"""
        try:
            from tkinter import messagebox

            # 1. جلب جميع عناصر الجدول بناءً على عمود اسم الملف ("name")
            col = "name"
            l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
            
            # 2. تنظيف خفيف للاسم من الإيموجي والرموز قبل إدخاله في الفرز الطبيعي الخاص بك
            def clean_name(item_tuple):
                raw_name = item_tuple[0]
                return raw_name.replace("📄", "").replace("📁", "").replace("📋", "").strip()

            # 3. 🔥 تطبيق الترتيب الذكي (Natural Sort) الخاص بك
            l.sort(key=lambda t: self.natural_sort_key(clean_name(t)), reverse=reverse_order)

            # 4. إعادة ترتيب العناصر في واجهة الجدول بناءً على الترتيب الجديد
            for index, (val, k) in enumerate(l):
                self.tree.move(k, '', index)

        except Exception as e:
            messagebox.showerror("خطأ في الترتيب", f"تعذر ترتيب الملفات:\n{e}")

    #---------------------------------------------------------------------------------            

    
    def move_focus_to_list(self, event):
        """عند الضغط على السهم السفلي في البحث، يتم الانتقال تلقائياً لأول عنصر في الجدول"""
        if hasattr(self, 'tree'):
            children = self.tree.get_children()
            if children:
                # 1. تحديد السطر الأول في الجدول برمجياً
                first_item = children[0]
                self.tree.selection_set(first_item)
                self.tree.focus(first_item)
                
                # 2. نقل التحكم الفعلي للوحة المفاتيح إلى الجدول ليتحرك بالأسهم
                self.tree.focus_set()
                
                # 3. منع حقل الإدخال من إزاحة مؤشر الكتابة الداخلي
                return "break"
                
    
        
    #-------------------------------------------
    def load_customers_from_file(self):
        """تحميل الزبائن بدون تكرار كلمة 'كل الزبائن'"""
        # نحدد الكلمة بناءً على اللغة
        default_option = "كل الزبائن" if self.current_lang == "العربية" else "All Customers"
        
        try:
            if not os.path.exists(self.customers_file):
                return [default_option]
                
            with open(self.customers_file, "r", encoding="utf-8") as f:
                # نجلب الأسماء ونزيل المسافات الزائدة
                lines = [line.strip() for line in f.readlines() if line.strip()]
            
            # نفلتر أي اسم قد يكون "كل الزبائن" موجوداً بداخل الملف عن طريق الخطأ
            filtered_lines = [name for name in lines if name != "كل الزبائن" and name != "All Customers"]
            
            # نضع الكلمة الافتراضية في المقدمة فقط
            return [default_option] + filtered_lines
        except Exception as e:
            return [default_option]
        
    def add_new_customer_manually(self):
        # 1. إنشاء نافذة منبثقة
        dialog = ctk.CTkToplevel(self)
        dialog.title("") 
        dialog.grab_set() # تجعل النافذة "مودال" (لا يمكنك الضغط على الخلفية حتى تغلقها)

        # 2. التمركز في المنتصف
        w, h = 300, 180
        x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        # 3. النص الديناميكي حسب اللغة
        lang_data = self.languages[self.current_lang]
        label_text = lang_data.get("enter_customer_name", "ادخل اسم الزبون")

        label = ctk.CTkLabel(dialog, text=label_text)
        label.pack(pady=20)

        entry = ctk.CTkEntry(dialog)
        entry.pack(pady=10)

        # زر الحفظ
        save_btn = ctk.CTkButton(dialog, text=lang_data.get("save", "حفظ"), command=lambda: self.save_customer(entry.get(), dialog))
        save_btn.pack(pady=10)

    def save_customer(self, name, dialog):
        # منطق الحفظ في الملف
        if name:
            with open(self.customers_file, "a", encoding="utf-8") as f:
                f.write(f"\n{name}")
            # تحديث القائمة المنسدلة في النافذة الرئيسية
            self.customer_filter.configure(values=self.load_customers_from_file())
            dialog.destroy() # إغلاق النافذة
        
        # زر الحفظ... 
    
    def delete_selected_customer(self):
        """حذف الزبون المختار من ملف customers.txt"""
        selected_customer = self.customer_filter.get()
        
        # لا نحذف "كل الزبائن"
        if selected_customer == "كل الزبائن":
            messagebox.showwarning("تنبيه", "لا يمكنك حذف هذا الخيار!")
            return
            
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف الزبون: {selected_customer}؟"):
            try:
                # 1. قراءة الملف
                with open(self.customers_file, "r", encoding="utf-8") as f:
                    lines = [line.strip() for line in f.readlines() if line.strip()]
                
                # 2. حذف الاسم المختار
                if selected_customer in lines:
                    lines.remove(selected_customer)
                    
                # 3. إعادة كتابة الملف
                with open(self.customers_file, "w", encoding="utf-8") as f:
                    f.write("\n".join(lines))
                
                # 4. تحديث القائمة في البرنامج
                self.customer_filter.configure(values=lines if lines else ["كل الزبائن"])
                self.customer_filter.set("كل الزبائن") # العودة للخيار الافتراضي
                
                messagebox.showinfo("نجاح", "تم حذف الزبون بنجاح!")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {e}")
    
                  
               
    def update_customer_list(self):
        """تحديث قائمة الفلتر بناءً على الملف النصي"""
        if hasattr(self, 'customer_filter'):
            new_list = self.load_customers_from_file()
            self.customer_filter.configure(values=new_list)        

    # ... تأتي بعدها دالة add_file أو غيرها ...-----------------    

    def delete_file(self, event=None):
        selections = self.tree.selection()
        if not selections:
            return

        lang = self.languages[self.current_lang]
        
        # استخدام نافذة الرسائل المخصصة
        dialog = CustomMessageBox(
            self, 
            title=lang["msg_del_title"], 
            message=lang["msg_confirm_del"],
            msg_type="question",
            show_cancel=True
        )
        
        if dialog.result:
            try:
                for item in selections:
                    vals = self.tree.item(item)['values']
                    # إزالة الأيقونة من اسم الملف
                    raw_name = str(vals[0])
                    # إزالة أيقونة الإيموجي والمسافات الزائدة
                    clean_name = raw_name
                    for info in FILE_TYPE_COLORS.values():
                        if clean_name.startswith(info["icon"]):
                            clean_name = clean_name[len(info["icon"]):].strip()
                            break
                    if clean_name.startswith(DEFAULT_FILE_COLOR["icon"]):
                        clean_name = clean_name[len(DEFAULT_FILE_COLOR["icon"]):].strip()
                    
                    file_type = str(vals[3]).strip().lower()
                    filename = f"{clean_name}.{file_type}"
                    path = os.path.join(self.archive_dir, filename)
                    
                    if os.path.exists(path):
                        os.remove(path)
                
                self.load_initial_archive()
                self.update_file_stats()
                
            except Exception as e:
                CustomMessageBox(self, title=lang["msg_error"], message=str(e), msg_type="error")

    def open_file(self, event=None):
        # 🌟 الحل السحري لزر App: إغلاق القائمة فوراً عند الضغط على خيار "فتح" وقبل تنفيذ أي شيء آخر
        try:
            if hasattr(self, 'context_menu'):
                self.context_menu.unpost()
        except Exception as e:
            print(f"Error hiding menu: {e}")

        def execute_open():
            try:
                sel = self.tree.selection()
                if sel:
                    vals = self.tree.item(sel[0])['values']
                    raw_name = str(vals[0])
                    clean_name = raw_name
                    for info in FILE_TYPE_COLORS.values():
                        if clean_name.startswith(info["icon"]):
                            clean_name = clean_name[len(info["icon"]):].strip()
                            break
                    if clean_name.startswith(DEFAULT_FILE_COLOR["icon"]):
                        clean_name = clean_name[len(DEFAULT_FILE_COLOR["icon"]):].strip()
                    
                    file_type = str(vals[3]).strip().lower()
                    filename = f"{clean_name}.{file_type}"
                    full_path = os.path.join(self.archive_dir, filename)
                    
                    # التأكد من وجود الملف الحقيقي لمنع الكراش
                    if os.path.exists(full_path):
                        os.startfile(full_path)
                    else:
                        from tkinter import messagebox
                        messagebox.showwarning("تنبيه", f"الملف الحقيقي غير موجود في الأرشيف:\n{filename}")
            except Exception as e:
                print(f"Error opening file: {e}")

        # الحفاظ على التأخير الآمن بـ 100 مللي ثانية لمنع كراش الـ GIL في بايثون 3.14
        self.after(100, execute_open)

    def rename_file(self, event=None):
        sel = self.tree.selection()
        if not sel: 
            return
            
        item_values = self.tree.item(sel[0])['values']
        raw_name = str(item_values[0])
        clean_name = raw_name
        for info in FILE_TYPE_COLORS.values():
            if clean_name.startswith(info["icon"]):
                clean_name = clean_name[len(info["icon"]):].strip()
                break
        if clean_name.startswith(DEFAULT_FILE_COLOR["icon"]):
            clean_name = clean_name[len(DEFAULT_FILE_COLOR["icon"]):].strip()
        
        old_name_only = clean_name
        extension = str(item_values[3]).strip().lower().replace('.', '')
        
        lang = self.languages[self.current_lang]
        
        # نافذة إعادة التسمية المخصصة
        rename_win = ctk.CTkToplevel(self)
        rename_win.title(lang.get("btn_edit", "Rename"))
        rename_win.geometry("400x220")
        rename_win.configure(fg_color=COLORS["bg_main"])
        rename_win.grab_set()
        rename_win.attributes("-topmost", True)
        rename_win.resizable(False, False)
        
        # توسيط
        rename_win.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() // 2) - 200
        py = self.winfo_y() + (self.winfo_height() // 2) - 110
        rename_win.geometry(f"+{px}+{py}")
        
        # اللوقو
        ctk.CTkLabel(
            rename_win, text="",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=COLORS["accent_primary"]
        ).pack(pady=(15, 5))
        
        # عنوان
        edit_label_text = ": الجديد الاسم أدخل" if self.current_lang == "العربية" else "Enter new name:"
        ctk.CTkLabel(
            rename_win, text=edit_label_text,
            font=ctk.CTkFont(size=13),
            text_color=COLORS["text_primary"]
        ).pack(pady=(5, 8))
        
        # حقل الإدخال
        name_entry = ctk.CTkEntry(
            rename_win, width=320, height=40,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["input_bg"],
            border_color=COLORS["input_border"],
            text_color=COLORS["text_primary"],
            corner_radius=10
        )
        name_entry.pack(pady=5)
        name_entry.insert(0, old_name_only)
        name_entry.select_range(0, 'end')
        name_entry.focus()
        
        # إطار الأزرار
        btn_frame = ctk.CTkFrame(rename_win, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        def do_rename():
            new_name_only = name_entry.get()
            if new_name_only and new_name_only.strip() != "":
                try:
                    new_name_cleaned = new_name_only.replace(f".{extension}", "").strip()
                    old_full_path = os.path.join(self.archive_dir, f"{old_name_only}.{extension}")
                    new_full_path = os.path.join(self.archive_dir, f"{new_name_cleaned}.{extension}")
                    
                    if os.path.exists(new_full_path) and old_full_path != new_full_path:
                        msg = "الاسم موجود مسبقاً!" if self.current_lang == "العربية" else "Name already exists!"
                        CustomMessageBox(rename_win, title="⚠️", message=msg, msg_type="warning")
                        return

                    os.rename(old_full_path, new_full_path)
                    rename_win.destroy()
                    self.load_initial_archive()
                    
                except Exception as e:
                    CustomMessageBox(rename_win, title="Error", message=f"Could not rename: {e}", msg_type="error")
        
        # زر إلغاء 🚫
        cancel_text = "إلغاء" if self.current_lang == "العربية" else "Cancel"
        ctk.CTkButton(
            btn_frame, text=f" {cancel_text}", width=140, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#2d333b",
            hover_color="#444c56",
            border_width=1, border_color="#444c56",
            corner_radius=12,
            text_color="#adbac7",
            command=rename_win.destroy
        ).pack(side="left", padx=8)
        
        # زر حفظ 💾
        save_text = "حفظ" if self.current_lang == "العربية" else "Save"
        ctk.CTkButton(
            btn_frame, text=f" {save_text}", width=140, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=COLORS["accent_green"],
            hover_color="#2ea043",
            corner_radius=12,
            text_color="#ffffff",
            command=do_rename
        ).pack(side="left", padx=8)
        
        # ربط Enter
        name_entry.bind("<Return>", lambda e: do_rename())
        rename_win.bind("<Escape>", lambda e: rename_win.destroy())

    def change_language(self, choice):
        self.current_lang = choice
        lang = self.languages[choice]
        self.title(lang["title"])
        self.logo_label.configure(text="SwiftFolder Pro")
        self.btn_settings.configure(text=lang["settings"])
        self.btn_backup.configure(text=lang["backup"])
        self.btn_import.configure(text=lang["import_btn"])
        self.btn_scan.configure(text=lang["scan"])
        self.btn_add.configure(text=lang["add"])
        self.search_entry.configure(placeholder_text=lang["search"])
        self.copy_label.configure(text=lang["copyright"])
        self.update_tree_headings()
        self.update_file_stats()
        
        # تحديث الفلاتر
        self.year_filter.configure(values=[lang["all_years"], "2030", "2029", "2028", "2027", "2026", "2025", "2024", "2023", "2022", "2021", "2020"])
        self.year_filter.set(lang["all_years"])
        self.type_filter.configure(values=[lang["all_types"], "PDF", "JPG", "PNG", "XLSX", "DOCX"])
        self.type_filter.set(lang["all_types"])
        self.customer_filter.configure(values=self.load_customers_from_file())
        self.customer_filter.set(self.load_customers_from_file()[0])
        

        # تحديث قائمة السياق
        self.context_menu.delete(0, "end")
        self.context_menu.add_command(label="🔓" + lang["btn_open"], command=self.open_file)
        self.context_menu.add_command(label="📝" + lang["btn_edit"], command=self.rename_file)
        self.context_menu.add_command(label="📥 إنزال / نسخ إلى الجهاز", command=self.download_selected_files)
        
        self.context_menu.add_command(label="🗑️" + lang["btn_delete"], command=self.delete_file)
        
        # تشغيل دالة التحديث الذكي للأزرار فوراً عند تغيير التحديد في الجدول
        self.tree.bind("<<TreeviewSelect>>", self.update_sidebar_buttons)

        if self.settings_window: self.settings_window.destroy()

    def update_clock(self):
        self.clock_label.configure(text=datetime.now().strftime("%Y-%m-%d  |  %H:%M:%S"))
        self.after(1000, self.update_clock)

    def create_backup(self):
        try:
            import pyzipper
            import os
            from tkinter import filedialog
            from datetime import datetime

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            folder_inside_zip = f"Backup_Files_{timestamp}"
            default_name = f"Backup_{timestamp}.zip"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                filetypes=[("Zip files", "*.zip")],
                initialfile=default_name
            )

            if not save_path:
                return

            password = b"12345"

            with pyzipper.AESZipFile(save_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as backup_zip:
                backup_zip.setpassword(password)
                
                if os.path.exists(self.archive_dir):
                    for file_name in os.listdir(self.archive_dir):
                        full_path = os.path.join(self.archive_dir, file_name)
                        
                        if os.path.isfile(full_path) and not file_name.endswith('.zip'):
                            arcname_with_folder = os.path.join(folder_inside_zip, file_name)
                            backup_zip.write(full_path, arcname=arcname_with_folder)

            lang = self.languages[self.current_lang]
            CustomMessageBox(
                self, 
                title=lang.get("success_title", "نجاح"), 
                message=lang.get("backup_success", "تم إنشاء النسخة الاحتياطية بنجاح!"),
                msg_type="success"
            )

        except Exception as e:
            err_title = "خطأ" if self.current_lang == "العربية" else "Error"
            CustomMessageBox(self, title=err_title, message=f"حدثت مشكلة: {str(e)}", msg_type="error")
    
    def import_backup_folder(self):
        try:
            import pyzipper, os, shutil
            import customtkinter as ctk
            from tkinter import filedialog

            lang = self.languages[self.current_lang]
            is_ar = self.current_lang == "العربية"

            zip_path = filedialog.askopenfilename(filetypes=[("Zip files", "*.zip")])
            if not zip_path: return

            # نافذة إدخال كلمة السر المخصصة
            pwd_win = ctk.CTkToplevel(self)
            pwd_title = "تشفير" if is_ar else "Security"
            pwd_win.title(pwd_title)
            pwd_win.geometry("400x230")
            pwd_win.configure(fg_color=COLORS["bg_main"])
            pwd_win.grab_set()
            pwd_win.attributes("-topmost", True)
            pwd_win.resizable(False, False)
            
            pwd_win.update_idletasks()
            px = self.winfo_x() + (self.winfo_width() // 2) - 200
            py = self.winfo_y() + (self.winfo_height() // 2) - 115
            pwd_win.geometry(f"+{px}+{py}")
            
            # اللوقو
            ctk.CTkLabel(
                pwd_win, text="",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color=COLORS["accent_primary"]
            ).pack(pady=(15, 5))
            
            pwd_text = ":النسخة سر كلمة أدخل" if is_ar else "Enter backup password:"
            ctk.CTkLabel(
                pwd_win, text=pwd_text,
                font=ctk.CTkFont(size=13),
                text_color=COLORS["text_primary"]
            ).pack(pady=(5, 8))
            
            pwd_entry = ctk.CTkEntry(
                pwd_win, width=300, height=40, show="*",
                font=ctk.CTkFont(size=13),
                fg_color=COLORS["input_bg"],
                border_color=COLORS["input_border"],
                text_color=COLORS["text_primary"],
                corner_radius=10
            )
            pwd_entry.pack(pady=5)
            pwd_entry.focus()
            
            password_result = [None]
            
            def on_submit():
                password_result[0] = pwd_entry.get()
                pwd_win.destroy()
            
            btn_frame = ctk.CTkFrame(pwd_win, fg_color="transparent")
            btn_frame.pack(pady=15)
            
            cancel_text = "إلغاء" if is_ar else "Cancel"
            ctk.CTkButton(
                btn_frame, text=f"{cancel_text}", width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color="#2d333b",
                hover_color="#444c56",
                border_width=1, border_color="#444c56",
                corner_radius=12,
                text_color="#adbac7",
                command=pwd_win.destroy
            ).pack(side="left", padx=8)
            
            confirm_text = "تأكيد" if is_ar else "Confirm"
            ctk.CTkButton(
                btn_frame, text=f"{confirm_text}", width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color=COLORS["accent_primary"],
                hover_color=COLORS["accent_hover"],
                corner_radius=12,
                text_color="#ffffff",
                command=on_submit
            ).pack(side="left", padx=8)
            
            pwd_entry.bind("<Return>", lambda e: on_submit())
            pwd_win.bind("<Escape>", lambda e: pwd_win.destroy())
            
            self.wait_window(pwd_win)
            
            password = password_result[0]
            if not password: return

            added_count = 0
            temp_dir = os.path.join(self.archive_dir, "temp_extract")
            os.makedirs(temp_dir, exist_ok=True)

            with pyzipper.AESZipFile(zip_path, 'r') as zip_ref:
                zip_ref.setpassword(password.encode('utf-8'))
                zip_ref.extractall(temp_dir)
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        shutil.move(os.path.join(root, file), os.path.join(self.archive_dir, file))
                        added_count += 1
                shutil.rmtree(temp_dir)

            self.load_initial_archive()

            s_title = lang.get("success_title", "نجاح")
            s_msg = lang.get("msg_success", "تم استيراد النسخة بنجاح")
            CustomMessageBox(self, title=s_title, message=f"{s_msg} ({added_count})", msg_type="success")

        except Exception as e:
            err_title = "خطأ" if self.current_lang == "العربية" else "Error"
            CustomMessageBox(self, title=err_title, message=str(e), msg_type="error")
            
            
    #-----------------------------------------------------------------------------------------------------------

    def scan_document(self):
        """دالة محسنة تظهر نافذة اختيار الأجهزة المرتبطة أولاً قبل بدء السكانر"""
        try:
            import win32com.client
            from PIL import Image
            import customtkinter as ctk
            import time
            import os
            from datetime import datetime
            
            lang_code = self.current_lang
            lang = self.languages[lang_code]
            
            # 1. استدعاء أداة مسح الصور والوثائق لويندوز (WIA)
            common_dialog = win32com.client.Dispatch("WIA.CommonDialog")
            
            try:
                # 🌟 التعديل الذهبي: إجبار الويندوز على إظهار نافذة اختيار أجهزة السكانر والطابعات المرتبطة
                # المعامل 1 يعني أجهزة مسح الصور (Scanners)، والمعامل True يضمن إظهار النافذة دائماً للمستخدم
                scanner_device = common_dialog.ShowSelectDevice(1, True, False)
            except Exception as device_err:
                # في حال قام المستخدم بإغلاق نافذة اختيار الجهاز أو الضغط على إلغاء (Cancel)
                print(f"User canceled device selection or no device found: {device_err}")
                return

            # إذا لم يقم المستخدم باختيار أي جهاز أو أغلق النافذة، اخرج فوراً دون المتابعة
            if not scanner_device: 
                return
            
            scanned_pages_paths = []
            
            # 2. بناء نافذة التقدم المحسّنة (تظهر بعد اختيار الجهاز بنجاح)
            progress_win = ctk.CTkToplevel(self)
            progress_win.configure(fg_color=COLORS["bg_main"])
           
            title_text = "🖨️ جاري المسح..." if lang_code == "العربية" else "🖨️ Scanning..."
            progress_win.title(lang.get("scanning_title", title_text))
           
            progress_win.geometry("380x220")
            progress_win.attributes("-topmost", True)
            progress_win.grab_set()
            progress_win.resizable(False, False)
           
            progress_win.update_idletasks()
            px = self.winfo_x() + (self.winfo_width() // 2) - 190
            py = self.winfo_y() + (self.winfo_height() // 2) - 110
            progress_win.geometry(f"+{px}+{py}")
            
            # اللوقو
            ctk.CTkLabel(
                progress_win, text="",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color=COLORS["accent_primary"]
            ).pack(pady=(15, 8))
            
            start_msg = "🖨️ بدء عملية المسح..." if lang_code == "العربية" else "🖨️ Starting scan..."
            status_label = ctk.CTkLabel(
                progress_win, text=lang.get("scan_start", start_msg),
                font=("Segoe UI", 13),
                text_color=COLORS["text_primary"]
            )
            status_label.pack(pady=8)
            
            # شريط التقدم
            prog_bar = ctk.CTkProgressBar(
                progress_win, width=280, height=10,
                progress_color="#2ea043",          
                fg_color=COLORS["bg_card"],
                corner_radius=5
            )
            prog_bar.pack(pady=12)
            prog_bar.set(0)
            
            count_msg = "" if lang_code == "العربية" else "Scanned pages: 0"
            count_label = ctk.CTkLabel(
                progress_win, text=lang.get("scanned_count", count_msg),
                font=("Segoe UI", 15, "bold"),
                text_color=COLORS["accent_green"]
            )
            count_label.pack(pady=10)
            
            progress_win.update()
            
            # تقدم وهمي بسيط لتهيئة الجهاز
            for i in range(1, 41):
                prog_bar.set(i / 100)
                progress_win.update()
                time.sleep(0.015)
            
            # 3. محاولة سحب الصورة من الجهاز المختار
            try:
                item = scanner_device.Items(1)
                status_label.configure(text="جاري المسح والنقل..." if lang_code == "العربية" else "Scanning & transferring...")
                progress_win.update()
                
                image = item.Transfer("{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}")
               
                if image:
                    temp_path = os.path.join(self.archive_dir, f"temp_p_{int(time.time())}.jpg")
                    image.SaveFile(temp_path)
                    scanned_pages_paths.append(temp_path)
                    
                    # تقدم بعد النجاح
                    for i in range(41, 101):
                        prog_bar.set(i / 100)
                        progress_win.update()
                        time.sleep(0.012)
                    
                    done_msg = "✅ تم المسح!" if lang_code == "العربية" else "✅ Done!"
                    count_label.configure(text=lang.get("scan_done", done_msg))
                    status_label.configure(text=done_msg)
                    progress_win.update()
                    time.sleep(0.9)

            except Exception as scan_err:
                print(f"Internal Scan Error: {scan_err}")
                status_label.configure(text="⚠ خطأ أثناء المسح" if lang_code == "العربية" else "⚠ Scan error")
                prog_bar.configure(progress_color="red")
                progress_win.update()
                time.sleep(2.5)

            progress_win.destroy()
            
            if not scanned_pages_paths:
                return
            
            # ────── باقي الكود الأصلي لتسمية وحفظ الملف ──────
            default_name = f"Scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
           
            name_win = ctk.CTkToplevel(self)
            name_title = lang.get("scan_name_title", "حفظ")
            name_win.title(name_title)
            name_win.geometry("400x220")
            name_win.configure(fg_color=COLORS["bg_main"])
            name_win.grab_set()
            name_win.attributes("-topmost", True)
            name_win.resizable(False, False)
            
            name_win.update_idletasks()
            dx = self.winfo_x() + (self.winfo_width() // 2) - 200
            dy = self.winfo_y() + (self.winfo_height() // 2) - 110
            name_win.geometry(f"+{dx}+{dy}")
            
            ctk.CTkLabel(
                name_win, text="",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color=COLORS["accent_primary"]
            ).pack(pady=(15, 5))
            
            q_msg = lang.get("scan_name_msg", "أدخل اسم الملف:")
            ctk.CTkLabel(
                name_win, text=q_msg,
                font=ctk.CTkFont(size=13),
                text_color=COLORS["text_primary"]
            ).pack(pady=(5, 8))
            
            name_entry = ctk.CTkEntry(
                name_win, width=320, height=40,
                font=ctk.CTkFont(size=13),
                fg_color=COLORS["input_bg"],
                border_color=COLORS["input_border"],
                text_color=COLORS["text_primary"],
                corner_radius=10
            )
            name_entry.pack(pady=5)
            name_entry.insert(0, default_name)
            name_entry.select_range(0, 'end')
            name_entry.focus()
            
            name_result = [default_name]
            
            def on_name_submit():
                val = name_entry.get().strip()
                name_result[0] = val if val else default_name
                name_win.destroy()
            
            btn_frame = ctk.CTkFrame(name_win, fg_color="transparent")
            btn_frame.pack(pady=12)
            
            cancel_text = "إلغاء" if self.current_lang == "العربية" else "Cancel"
            ctk.CTkButton(
                btn_frame, text=f" {cancel_text}", width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color="#2d333b",
                hover_color="#444c56",
                border_width=1, border_color="#444c56",
                corner_radius=12,
                text_color="#adbac7",
                command=name_win.destroy
            ).pack(side="left", padx=8)
            
            save_text = "حفظ" if self.current_lang == "العربية" else "Save"
            ctk.CTkButton(
                btn_frame, text=f" {save_text}", width=140, height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color=COLORS["accent_green"],
                hover_color="#2ea043",
                corner_radius=12,
                text_color="#ffffff",
                command=on_name_submit
            ).pack(side="left", padx=8)
            
            name_entry.bind("<Return>", lambda e: on_name_submit())
            name_win.bind("<Escape>", lambda e: name_win.destroy())
            
            self.wait_window(name_win)
            
            final_name = name_result[0]
            pdf_path = os.path.join(self.archive_dir, f"{final_name}.pdf")
            
            if os.path.exists(scanned_pages_paths[0]):
                img = Image.open(scanned_pages_paths[0])
                img.convert('RGB').save(pdf_path, quality=95)
                img.close()
                os.remove(scanned_pages_paths[0])

            self.load_initial_archive()
            if hasattr(self, 'all_files_data'):
                self.all_files_data.sort(key=lambda x: os.path.getctime(x['path']), reverse=True)
                if hasattr(self, 'apply_filter'):
                    self.apply_filter()
            
            self.update_file_stats()

        except Exception as e:
            print(f"Main Error: {e}")
                
    #-----------------------------------------------------------------------------------------

    def open_settings(self):
        lang = self.languages[self.current_lang]
        self.settings_window = ctk.CTkToplevel(self)
        self.settings_window.title(lang["sett_win_title"])
        self.settings_window.geometry("620x420")
        self.settings_window.configure(fg_color=COLORS["bg_main"])
        self.settings_window.grab_set()
        self.settings_window.lift()
        self.settings_window.resizable(False, False)
        
        # توسيط النافذة
        self.settings_window.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() // 2) - 310
        py = self.winfo_y() + (self.winfo_height() // 2) - 210
        self.settings_window.geometry(f"+{px}+{py}")

        # 🌟 وضع السطر الجديد هنا: جعل زر Echap (Escape) يغلق نافذة الإعدادات فوراً
        self.settings_window.bind("<Escape>", lambda e: self.settings_window.destroy())

        # الشريط الجانبي
        sidebar = ctk.CTkFrame(
            self.settings_window, width=170, corner_radius=0, 
            fg_color=COLORS["bg_header"],
            border_width=0
        )
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        
        # لوقو في الشريط الجانبي
        logo_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=12, pady=(18, 15))
        
        ctk.CTkLabel(
            logo_frame, text="",
            font=ctk.CTkFont(size=18),
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkLabel(
            logo_frame, text="",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=COLORS["accent_primary"]
        ).pack(side="left")
        
        # خط فاصل
        ctk.CTkFrame(
            sidebar, height=1, fg_color=COLORS["border"]
        ).pack(fill="x", padx=12, pady=(0, 10))
        
        # منطقة المحتوى
        content_area = ctk.CTkFrame(self.settings_window, fg_color="transparent")
        content_area.pack(side="right", expand=True, fill="both", padx=25, pady=25)

        # متغير لتتبع الزر النشط
        self.active_tab_btn = None

        def clear_content():
            for widget in content_area.winfo_children(): widget.destroy()

        def set_active_btn(btn):
            if self.active_tab_btn:
                self.active_tab_btn.configure(fg_color="transparent")
            btn.configure(fg_color=COLORS["bg_card"])
            self.active_tab_btn = btn

        def show_lang_page():
            clear_content()
            set_active_btn(lang_btn)
            
            ctk.CTkLabel(
                content_area, text=lang["sett_lang_head"], 
                font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
                text_color="#ffffff"
            ).pack(pady=(10, 25))
            
            # إطار الأزرار
            btn_frame = ctk.CTkFrame(content_area, fg_color="transparent")
            btn_frame.pack(pady=10)
            
            # زر العربية 🇸🇦
            ar_selected = self.current_lang == "العربية"
            ar_btn = ctk.CTkButton(
                btn_frame, text="العربية", width=140, height=36,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                fg_color=COLORS["accent_primary"] if ar_selected else "#2d333b",
                hover_color=COLORS["accent_hover"] if ar_selected else "#444c56",
                border_width=2 if ar_selected else 1,
                border_color=COLORS["accent_primary"] if ar_selected else "#444c56",
                corner_radius=12,
                text_color="#ffffff" if ar_selected else "#adbac7",
                command=lambda: self.change_language("العربية")
            )
            ar_btn.pack(side="left", padx=10)
            
            # زر الإنجليزية 🇺🇸
            en_selected = self.current_lang == "English"
            en_btn = ctk.CTkButton(
                btn_frame, text="English", width=140, height=36,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                fg_color=COLORS["accent_primary"] if en_selected else "#2d333b",
                hover_color=COLORS["accent_hover"] if en_selected else "#444c56",
                border_width=2 if en_selected else 1,
                border_color=COLORS["accent_primary"] if en_selected else "#444c56",
                corner_radius=12,
                text_color="#ffffff" if en_selected else "#adbac7",
                command=lambda: self.change_language("English")
            )
            en_btn.pack(side="left", padx=10)

        def show_pwd_page():
            clear_content()
            set_active_btn(pwd_btn)
            
            ctk.CTkLabel(
                content_area, text=lang["sett_pwd_head"], 
                font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
                text_color="#ffffff"
            ).pack(pady=(10, 25))
            
            old_e = ctk.CTkEntry(
                content_area, placeholder_text=lang["sett_old_pwd"], show="*",
                height=42, font=ctk.CTkFont(size=12),
                fg_color=COLORS["input_bg"], 
                border_color=COLORS["input_border"],
                text_color=COLORS["text_primary"],
                corner_radius=10
            )
            old_e.pack(pady=8, fill="x")
            
            new_e = ctk.CTkEntry(
                content_area, placeholder_text=lang["sett_new_pwd"], show="*",
                height=42, font=ctk.CTkFont(size=12),
                fg_color=COLORS["input_bg"], 
                border_color=COLORS["input_border"],
                text_color=COLORS["text_primary"],
                corner_radius=10
            )
            new_e.pack(pady=8, fill="x")
            
            def save_new_pwd():
                with open(self.password_file, "r") as f: cur = f.read().strip()
                if old_e.get() == cur:
                    with open(self.password_file, "w") as f: f.write(new_e.get())
                    success_msg = "✅ تم الحفظ بنجاح!" if self.current_lang == "العربية" else "✅ Saved successfully!"
                    CustomMessageBox(self.settings_window, title=lang.get("success_title", "نجاح"), message=success_msg, msg_type="success")
                else:
                    err_msg = "❌ كلمة السر الحالية غير صحيحة!" if self.current_lang == "العربية" else "❌ Current password is incorrect!"
                    CustomMessageBox(self.settings_window, title=lang.get("msg_error", "خطأ"), message=err_msg, msg_type="error")
            
            ctk.CTkButton(
                content_area, text=lang["sett_save"],
                height=42, font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                fg_color=COLORS["accent_green"], 
                hover_color="#2ea043",
                corner_radius=12,
                text_color="#ffffff",
                command=save_new_pwd
            ).pack(pady=18, fill="x")

        # أزرار الشريط الجانبي - محسّنة
        lang_btn = ctk.CTkButton(
            sidebar, text=lang["sett_lang_tab"], 
            fg_color="transparent", 
            hover_color=COLORS["bg_card"],
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            anchor="w", height=44,
            corner_radius=10,
            command=show_lang_page
        )
        lang_btn.pack(fill="x", padx=10, pady=(5, 3))
        
        pwd_btn = ctk.CTkButton(
            sidebar, text=lang["sett_pwd_tab"], 
            fg_color="transparent", 
            hover_color=COLORS["bg_card"],
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            anchor="w", height=44,
            corner_radius=10,
            command=show_pwd_page
        )
        pwd_btn.pack(fill="x", padx=10, pady=3)
        
        # مساحة فارغة
        ctk.CTkFrame(sidebar, fg_color="transparent").pack(expand=True, fill="both")
        
        # عرض صفحة اللغة افتراضياً
        show_lang_page()

if __name__ == "__main__":
    app = SwiftFolderPro()
    app.mainloop()
