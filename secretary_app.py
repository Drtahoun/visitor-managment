import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from tkcalendar import Calendar
from datetime import datetime, date, timedelta
import sqlite3
import pygame
import os
import time
import pandas as pd
from PIL import Image, ImageTk, ImageDraw, ImageFont
import hashlib
import shutil
from threading import Thread
import sys

class SecretaryApp:
    def __init__(self, root, username):
        self.root = root
        self.username = username
        self.root.title("برنامج إدارة الزائرين - جهاز مدير المكتب")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f5f5')
        
        # Flash message frame
        self.flash_frame = tk.Frame(self.root, bg='#FFD700', height=40, bd=1, relief='solid')
        self.flash_label = tk.Label(self.flash_frame, text="", bg='#FFD700', font=('Arial', 12, 'bold'), fg='#333')
        self.flash_label.pack(pady=5)
        self.flash_frame.pack_propagate(False)
        
        # Approved visitor frame
        self.approved_visitor_frame = tk.Frame(self.root, bg='#d4edda', height=40, bd=1, relief='solid')
        self.approved_visitor_label = tk.Label(self.approved_visitor_frame, text="", bg='#d4edda', 
                                             font=('Arial', 12, 'bold'), fg='#155724')
        self.approved_visitor_label.pack(pady=5)
        self.approved_visitor_frame.pack_propagate(False)

        self.animation_colors = ['#FFD700', '#FFA500', '#FF8C00']
        self.current_animation_index = 0
        self.load_icons()
        self.selected_visitor = None
        self.current_date = date.today().strftime("%Y-%m-%d")
        self.auto_refresh_enabled = True
        self.show_hidden_flag = False
        self.current_filters = {'name': None, 'date': None}
        self.logo_image = None
        self.notification_played = False
        self.current_approved_visitor = None
        self._last_notification = 0
        
        self.position_options = ["رئيس قطاع", "وكيل وزارة", "مدير عام", "مدير إدارة", 
                               "مهندس", "موظف", "عامل", "رئيس مجلس نواب", "مقاول", "ضيف",
                               "محافظ", "السيد الوزير", "عضو مجلس نواب"]
        
        # Initialize database with faster connection
        self.db_conn = sqlite3.connect('visitors_management.db', check_same_thread=False, timeout=10)
        self.db_cursor = self.db_conn.cursor()
        self.db_conn.execute("PRAGMA journal_mode=WAL")  # Enable WAL mode for better concurrency
        self.init_database()
        
        # Initialize sound with proper resource handling
        pygame.mixer.init()
        self.notification_sound = None
        self.load_notification_sound()
        
        self.setup_ui()
        self.load_data()
        
        self.start_auto_backup()
        self.root.after(3000, self.auto_refresh)
        self.check_new_visitors()

    def check_new_visitors(self):
        self.db_cursor.execute("SELECT COUNT(*) FROM visitors WHERE new_flag=1")
        count = self.db_cursor.fetchone()[0]
        
        if count > 0 and not self.notification_played:
            self.show_flash_message(f"لديك {count} زائر(ين) جديد(ين) يحتاج(ون) إلى مراجعة")
            self.play_notification_sound()
            self.notification_played = True
        
        self.root.after(10000, self.check_new_visitors)

    def play_notification_sound(self):
        current_time = time.time()
        if current_time - self._last_notification > 10:
            if self.notification_sound:
                try:
                    self.notification_sound.play()
                except:
                    pass
            self._last_notification = current_time

    def load_notification_sound(self):
        try:
            sound_path = os.path.join(os.path.dirname(sys.executable), 'notification.wav') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'notification.wav')
            if os.path.exists(sound_path):
                self.notification_sound = pygame.mixer.Sound(sound_path)
            else:
                print("ملف الصوت notification.wav غير موجود")
        except Exception as e:
                print(f"تعذر تحميل ملف الصوت: {str(e)}")

    def init_database(self):
        # Create tables if not exists with optimized structure
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS visitors (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              daily_id INTEGER DEFAULT 1,
                              name TEXT NOT NULL,
                              phone TEXT,
                              job TEXT,
                              position TEXT,
                              nationality TEXT,
                              visit_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                              status TEXT DEFAULT 'انتظار',
                              entry_time TEXT,
                              exit_time TEXT,
                              new_flag INTEGER DEFAULT 1,
                              hidden INTEGER DEFAULT 0,
                              appointment_time TEXT)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS daily_counter (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              last_date TEXT,
                              counter INTEGER DEFAULT 1)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS positions (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              position_name TEXT UNIQUE)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              username TEXT UNIQUE,
                              password TEXT)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS appointments (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              visitor_name TEXT NOT NULL,
                              phone TEXT,
                              job TEXT,
                              position TEXT,
                              nationality TEXT,
                              appointment_date TEXT NOT NULL,
                              appointment_time TEXT NOT NULL,
                              status TEXT DEFAULT 'معلقة',
                              created_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS commitments (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              place TEXT NOT NULL,
                              purpose TEXT NOT NULL,
                              department TEXT NOT NULL,
                              commitment_date TEXT NOT NULL,
                              commitment_time TEXT NOT NULL,
                              status TEXT DEFAULT 'معلقة',
                              created_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS visit_details (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              visitor_id INTEGER NOT NULL,
                              visit_results TEXT,
                              recommendations TEXT,
                              required_date TEXT,
                              execution_status TEXT DEFAULT 'تحت التنفيذ',
                              FOREIGN KEY(visitor_id) REFERENCES visitors(id))''')
        
        # Insert default users if not exists
        default_users = [
            ("sec", self.hash_password("000000")),
            ("head", self.hash_password("admin")),
            ("it", self.hash_password("admin"))
        ]
        
        for username, password in default_users:
            try:
                self.db_cursor.execute("INSERT OR IGNORE INTO users (username, password) VALUES (?, ?)", 
                                      (username, password))
            except sqlite3.IntegrityError:
                pass
        
        # Load positions from database
        self.db_cursor.execute("SELECT position_name FROM positions")
        db_positions = [row[0] for row in self.db_cursor.fetchall()]
        self.position_options = list(set(self.position_options + db_positions))
        
        self.db_conn.commit()

    @staticmethod
    def hash_password(password):
        return hashlib.sha256(password.encode()).hexdigest()

    def setup_ui(self):
        # Header Frame
        header_frame = tk.Frame(self.root, bg='#2c3e50')
        header_frame.pack(fill=tk.X, pady=5)
        
        # Load Logo with proper resource handling
        try:
            logo_path = os.path.join(os.path.dirname(sys.executable), 'logo.png') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'logo.png')
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((80, 80), Image.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(logo_img)
                logo_label = tk.Label(header_frame, image=self.logo_image, bg='#2c3e50')
                logo_label.pack(side=tk.RIGHT, padx=20)
            else:
                raise FileNotFoundError("ملف اللوجو غير موجود")
        except Exception as e:
            print(f"Error loading logo: {e}")
            logo_label = tk.Label(header_frame, text="شعار المؤسسة", 
                                font=('Arial', 14), bg='#2c3e50', fg='white')
            logo_label.pack(side=tk.RIGHT, padx=20)
        
        # Titles Frame
        titles_frame = tk.Frame(header_frame, bg='#2c3e50')
        titles_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)
        
        tk.Label(titles_frame, 
                text="وزارة الإسكان والمرافق والمجتمعات العمرانية", 
                font=('Arial', 16, 'bold'), 
                bg='#2c3e50', 
                fg='white').pack()
        
        tk.Label(titles_frame, 
                text="الجهاز المركزي للتعمير", 
                font=('Arial', 14), 
                bg='#2c3e50', 
                fg='white').pack()
        
        tk.Label(titles_frame, 
                text="برنامج إدارة الزائرين - رئيس الإدارة المركزية لشئون مكتب رئيس الجهاز", 
                font=('Arial', 12), 
                bg='#2c3e50', 
                fg='white').pack()

        # User Frame
        user_frame = tk.Frame(header_frame, bg='#2c3e50')
        user_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        tk.Label(user_frame, 
               text=f"المستخدم: {self.username}", 
               font=('Arial', 12), 
               bg='#2c3e50', 
               fg='white').pack(pady=5)
        
        change_pass_btn = tk.Button(user_frame, 
                                  text="تغيير كلمة المرور", 
                                  command=self.change_password,
                                  font=('Arial', 10), 
                                  bg='#27ae60', 
                                  fg='white',
                                  bd=0)
        change_pass_btn.pack(pady=5)

        # Input Form
        self.create_input_form()
        
        # Buttons Frame
        self.create_buttons_frame()
        
        # Visitors Table
        self.create_visitors_table()
        
        # Finish Visit Button
        finish_frame = tk.Frame(self.root, bg='#f5f5f5')
        finish_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        self.finish_btn = tk.Button(finish_frame, text="انتهاء الزيارة",
                                  command=self.finish_visit,
                                  font=('Arial', 14, 'bold'), 
                                  bg='#17a2b8', fg='white',
                                  width=15, height=1, bd=0)
        self.finish_btn.pack(pady=5)
        
        # Context Menu
        self.create_context_menu()
        
        # Footer
        footer_frame = tk.Frame(self.root, bg='#f5f5f5')
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))
        
        footer = tk.Label(footer_frame, 
                         text="By Dr. Mohammad Tahoun", 
                         font=('Arial', 12, 'bold'), 
                         fg='#0066cc', 
                         bg='#f5f5f5')
        footer.pack(pady=5)

    def create_input_form(self):
        input_frame = tk.LabelFrame(self.root, text="قائمة الإدخال", padx=10, pady=10, 
                                  bg='#f5f5f5', fg='#333', font=('Arial', 12, 'bold'),
                                  bd=2, relief='groove')
        input_frame.pack(fill=tk.X, padx=15, pady=10)
        
        font_style = ('Arial', 12)
        
        # Name
        tk.Label(input_frame, text="اسم الزائر:", font=font_style, bg='#f5f5f5').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.name_entry = tk.Entry(input_frame, width=30, font=font_style, bg='#ffffff', fg='#333333', bd=1, relief='solid')
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Phone
        tk.Label(input_frame, text="رقم الهاتف:", font=font_style, bg='#f5f5f5').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.phone_entry = tk.Entry(input_frame, width=30, font=font_style, bg='#ffffff', fg='#333333', bd=1, relief='solid')
        self.phone_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Job
        tk.Label(input_frame, text="الوظيفة:", font=font_style, bg='#f5f5f5').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.job_entry = tk.Entry(input_frame, width=30, font=font_style, bg='#ffffff', fg='#333333', bd=1, relief='solid')
        self.job_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Position
        tk.Label(input_frame, text="الدرجة الوظيفية:", font=font_style, bg='#f5f5f5').grid(row=3, column=0, sticky='e', padx=5, pady=5)
        self.create_position_widgets(input_frame, row=3, column=1)
        
        # Nationality
        tk.Label(input_frame, text="الجنسية:", font=font_style, bg='#f5f5f5').grid(row=4, column=0, sticky='e', padx=5, pady=5)
        self.nationality_entry = tk.Entry(input_frame, width=30, font=font_style, bg='#ffffff', fg='#333333', bd=1, relief='solid')
        self.nationality_entry.grid(row=4, column=1, padx=5, pady=5)

    def create_position_widgets(self, parent, row, column):
        position_frame = tk.Frame(parent, bg='#f5f5f5')
        position_frame.grid(row=row, column=column, padx=5, pady=5, sticky='ew')
        
        self.position_var = tk.StringVar()
        self.position_menu = ttk.Combobox(position_frame, textvariable=self.position_var, 
                                        values=self.position_options, width=25, font=('Arial', 12), state='readonly')
        self.position_menu.pack(side=tk.LEFT, padx=(0, 5))
        
        position_btns_frame = tk.Frame(position_frame, bg='#f5f5f5')
        position_btns_frame.pack(side=tk.LEFT)
        
        add_position_btn = tk.Button(position_btns_frame, text="+", font=('Arial', 10),
                                   command=self.add_new_position, bg='#3498db', fg='white',
                                   width=2, height=1, bd=0)
        add_position_btn.grid(row=0, column=0, padx=2)
        
        edit_position_btn = tk.Button(position_btns_frame, text="تعديل", font=('Arial', 8),
                                    command=self.edit_position, bg='#f39c12', fg='white',
                                    width=4, height=1, bd=0)
        edit_position_btn.grid(row=0, column=1, padx=2)
        
        delete_position_btn = tk.Button(position_btns_frame, text="حذف", font=('Arial', 8),
                                      command=self.delete_position, bg='#e74c3c', fg='white',
                                      width=4, height=1, bd=0)
        delete_position_btn.grid(row=0, column=2, padx=2)

    def create_buttons_frame(self):
        button_frame = tk.Frame(self.root, bg='#f5f5f5')
        button_frame.pack(fill=tk.X, padx=15, pady=10)
        
        btn_style = {
            'font': ('Arial', 12), 
            'width': 15,
            'height': 1, 
            'padx': 10, 
            'pady': 8, 
            'bd': 0, 
            'relief': 'raised'
        }
        
        # Left Side Buttons
        self.send_btn = tk.Button(button_frame, text="إرسال البيانات", 
                                 command=self.send_data, bg='#3498db', fg='white', **btn_style)
        self.send_btn.pack(side=tk.LEFT, padx=5)

        self.search_name_btn = tk.Button(button_frame, text="بحث بالاسم", 
                                       command=self.search_by_name,
                                       bg='#2ecc71', fg='white', **btn_style)
        self.search_name_btn.pack(side=tk.LEFT, padx=5)

        self.search_date_btn = tk.Button(button_frame, text="بحث بالتاريخ", 
                                       command=self.search_by_date,
                                       bg='#e67e22', fg='white', **btn_style)
        self.search_date_btn.pack(side=tk.LEFT, padx=5)

        self.clear_filter_btn = tk.Button(button_frame, text="إلغاء الفرز", 
                                        command=self.clear_filters,
                                        bg='#95a5a6', fg='white', **btn_style)
        self.clear_filter_btn.pack(side=tk.LEFT, padx=5)

        self.stats_btn = tk.Button(button_frame, text="الإحصائيات", 
                                 command=self.show_comprehensive_stats,
                                 bg='#9b59b6', fg='white', **btn_style)
        self.stats_btn.pack(side=tk.LEFT, padx=5)

        self.calendar_btn = tk.Button(button_frame, text="حجز موعد",
                                    command=self.show_appointment_dialog,
                                    bg='#8e44ad', fg='white', **btn_style)
        self.calendar_btn.pack(side=tk.LEFT, padx=5)
        
        # Right Side Buttons
        self.db_btn = tk.Button(button_frame, text="قاعدة البيانات", 
                              command=self.show_database_options,
                              bg='#34495e', fg='white', **btn_style)
        self.db_btn.pack(side=tk.RIGHT, padx=5)
        
        self.commitment_btn = tk.Button(button_frame, text="الالتزامات",
                                      command=self.show_commitments_dialog,
                                      bg='#16a085', fg='white', **btn_style)
        self.commitment_btn.pack(side=tk.RIGHT, padx=5)
        
        self.appointments_btn = tk.Button(button_frame, text="عرض المواعيد",
                                        command=self.show_appointments_list,
                                        bg='#8e44ad', fg='white', **btn_style)
        self.appointments_btn.pack(side=tk.RIGHT, padx=5)
        
        self.tasks_btn = tk.Button(button_frame, text="عرض المهام",
                                 command=self.show_tasks,
                                 bg='#3498db', fg='white', **btn_style)
        self.tasks_btn.pack(side=tk.RIGHT, padx=5)

    def create_visitors_table(self):
        table_frame = tk.Frame(self.root, bg='#f5f5f5')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("Treeview", 
                      font=('Arial', 12), 
                      rowheight=30, 
                      background='#ffffff',
                      fieldbackground='#ffffff',
                      foreground='#333333',
                      bordercolor='#dddddd',
                      borderwidth=1)
        
        style.configure("Treeview.Heading", 
                      font=('Arial', 12, 'bold'), 
                      background='#34495e', 
                      foreground='white',
                      relief='flat')
        
        style.map("Treeview", 
                 background=[('selected', '#3498db')],
                 foreground=[('selected', 'white')])
        
        columns = ("id", "daily_id", "name", "phone", "job", "position", "nationality", "visit_date", "status", "entry_time", "exit_time")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        self.tree.heading("id", text="ID", anchor='center')
        self.tree.heading("daily_id", text="الرقم اليومي", anchor='center')
        self.tree.heading("name", text="اسم الزائر", anchor='center')
        self.tree.heading("phone", text="رقم الهاتف", anchor='center')
        self.tree.heading("job", text="الوظيفة", anchor='center')
        self.tree.heading("position", text="الدرجة الوظيفية", anchor='center')
        self.tree.heading("nationality", text="الجنسية", anchor='center')
        self.tree.heading("visit_date", text="تاريخ الزيارة", anchor='center')
        self.tree.heading("status", text="حالة الزيارة", anchor='center')
        self.tree.heading("entry_time", text="وقت الدخول", anchor='center')
        self.tree.heading("exit_time", text="وقت الخروج", anchor='center')
        
        self.tree.column("id", width=50, anchor='center')
        self.tree.column("daily_id", width=80, anchor='center')
        self.tree.column("name", width=150, anchor='center')
        self.tree.column("phone", width=120, anchor='center')
        self.tree.column("job", width=120, anchor='center')
        self.tree.column("position", width=150, anchor='center')
        self.tree.column("nationality", width=100, anchor='center')
        self.tree.column("visit_date", width=150, anchor='center')
        self.tree.column("status", width=120, anchor='center')
        self.tree.column("entry_time", width=150, anchor='center')
        self.tree.column("exit_time", width=150, anchor='center')
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.tag_configure('approved', background='#d4edda')
        self.tree.tag_configure('rejected', background='#f8d7da')
        self.tree.tag_configure('finished', background='#d1ecf1')
        self.tree.tag_configure('waiting', background='#ffffff')
        self.tree.tag_configure('highlight', background='#fffacd')
        self.tree.tag_configure('new_visitor', background='#e6ffe6')
        self.tree.tag_configure('hidden', background='#f8f9fa')
        self.tree.tag_configure('appointment', background='#e6f3ff')
        
        self.tree.bind("<Button-1>", self.on_tree_click)

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.selected_visitor = self.tree.item(item)['values']
            # Highlight the approved visitor
            if self.selected_visitor[8] == 'موافق':
                self.show_approved_visitor(self.selected_visitor[2])
                self.tree.item(item, tags=('highlight',))
                self.root.after(5000, lambda: self.reset_row_color(item))

    def reset_row_color(self, item):
        tags = list(self.tree.item(item)['tags'])
        if 'highlight' in tags:
            tags.remove('highlight')
        self.tree.item(item, tags=tuple(tags))

    def create_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0, bg='#f8f9fa', fg='#333', 
                                  font=('Arial', 10), activebackground='#3498db',
                                  activeforeground='white', bd=1, relief='solid')
        
        self.context_menu.add_command(label="تعديل", command=self.edit_visitor)
        self.context_menu.add_command(label="حذف", command=self.delete_visitor)
        self.context_menu.add_command(label="إخفاء", command=self.hide_visitor)
        self.context_menu.add_command(label="إظهار المحذوفين", command=self.toggle_hidden_visitors)
        
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.selected_visitor = self.tree.item(item)['values']
            self.context_menu.post(event.x_root, event.y_root)

    def load_icons(self):
        try:
            self.edit_icon = None
            self.delete_icon = None
        except Exception as e:
            print(f"Error loading icons: {e}")

    def start_auto_backup(self):
        def backup_task():
            while True:
                now = datetime.now()
                if now.hour == 0 and now.minute == 0:
                    self.create_backup()
                time.sleep(60)
        
        backup_thread = Thread(target=backup_task, daemon=True)
        backup_thread.start()

    def create_backup(self):
        backup_dir = "backup"
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        today = date.today().strftime("%Y-%m-%d")
        backup_path = os.path.join(backup_dir, f"visitors_backup_{today}.db")
        
        try:
            shutil.copy2('visitors_management.db', backup_path)
            print(f"تم إنشاء النسخة الاحتياطية بنجاح: {backup_path}")
        except Exception as e:
            print(f"فشل في إنشاء النسخة الاحتياطية: {str(e)}")

    def change_password(self):
        current_password = simpledialog.askstring("تغيير كلمة المرور", "أدخل كلمة المرور الحالية:", show='*')
        if not current_password:
            return
            
        new_password = simpledialog.askstring("تغيير كلمة المرور", "أدخل كلمة المرور الجديدة:", show='*')
        if not new_password:
            return
            
        confirm_password = simpledialog.askstring("تغيير كلمة المرور", "أعد إدخال كلمة المرور الجديدة:", show='*')
        if not confirm_password:
            return
            
        if new_password != confirm_password:
            messagebox.showerror("خطأ", "كلمات المرور الجديدة غير متطابقة")
            return
            
        self.db_cursor.execute("SELECT password FROM users WHERE username=?", (self.username,))
        result = self.db_cursor.fetchone()
        
        if not result or result[0] != self.hash_password(current_password):
            messagebox.showerror("خطأ", "كلمة المرور الحالية غير صحيحة")
            return
            
        try:
            self.db_cursor.execute("UPDATE users SET password=? WHERE username=?", 
                                 (self.hash_password(new_password), self.username))
            self.db_conn.commit()
            messagebox.showinfo("نجاح", "تم تغيير كلمة المرور بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تغيير كلمة المرور: {str(e)}")

    def get_daily_id(self):
        today = date.today().strftime("%Y-%m-%d")
        self.db_cursor.execute("SELECT last_date, counter FROM daily_counter LIMIT 1")
        result = self.db_cursor.fetchone()
        
        if result:
            last_date, counter = result
            if last_date == today:
                counter += 1
            else:
                counter = 1
            self.db_cursor.execute("UPDATE daily_counter SET last_date=?, counter=?", (today, counter))
        else:
            counter = 1
            self.db_cursor.execute("INSERT INTO daily_counter (last_date, counter) VALUES (?, ?)", (today, counter))
        
        self.db_conn.commit()
        return counter

    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        query = '''SELECT id, daily_id, name, phone, job, position, nationality, 
                  strftime('%Y-%m-%d %H:%M', visit_date) as visit_date, 
                  status, entry_time, exit_time, new_flag, hidden, appointment_time FROM visitors'''
        
        params = []
        
        if not self.show_hidden_flag:
            query += " WHERE hidden=0 AND status != 'انتهاء'"
        else:
            query += " WHERE 1=1"  # Always true condition to allow adding AND clauses
        
        # Filter for appointments - only show if appointment date is today or past
        if not self.show_hidden_flag:
            query += " AND (appointment_time IS NULL OR date(visit_date) <= date('now'))"
        
        if self.current_filters['name']:
            query += " AND name LIKE ?"
            params.append(f"%{self.current_filters['name']}%")
            
        if self.current_filters['date']:
            query += " AND date(visit_date)=?"
            params.append(self.current_filters['date'])
        
        # Sort by status (approved first), then by visit_date (newest first)
        query += " ORDER BY CASE WHEN status='موافق' THEN 0 WHEN status='انتظار' THEN 1 ELSE 2 END, visit_date DESC"
        
        self.db_cursor.execute(query, params)
        visitors = self.db_cursor.fetchall()
        
        for visitor in visitors:
            tags = self.get_visitor_tags(visitor)
            self.tree.insert("", tk.END, values=visitor, tags=tags)
            
            # Show approved visitor name at the top with flash
            if visitor[8] == 'موافق' and visitor[0] != getattr(self.current_approved_visitor, 'id', None):
                self.current_approved_visitor = type('', (), {'id': visitor[0], 'name': visitor[2]})()
                self.show_approved_visitor(visitor[2])
                self.show_flash_message(f"تمت الموافقة على الزائر: {visitor[2]}")
                self.play_notification_sound()

    def get_visitor_tags(self, visitor):
        tags = []
        if visitor[8] == 'موافق':
            tags.append('approved')
        elif visitor[8] == 'رفض':
            tags.append('rejected')
        elif visitor[8] == 'انتهاء':
            tags.append('finished')
        elif visitor[11] == 1:  # new_flag
            tags.append('new_visitor')
            self.show_flash_message(f"زائر جديد: {visitor[2]}")
            self.play_notification_sound()
        elif visitor[12] == 1:  # hidden
            tags.append('hidden')
        elif visitor[13]:  # appointment_time
            tags.append('appointment')
        return tuple(tags)

    def show_approved_visitor(self, visitor_name):
        self.approved_visitor_label.config(text=f"الزائر الحالي: {visitor_name}")
        self.approved_visitor_frame.pack(fill=tk.X, padx=10, pady=5)
        self.root.after(10000, self.hide_approved_visitor)

    def hide_approved_visitor(self):
        self.approved_visitor_frame.pack_forget()

    def show_flash_message(self, message):
        self.flash_label.config(text=message)
        self.flash_frame.pack(fill=tk.X, padx=10, pady=5)
        self.animate_flash_message()
        self.root.after(5000, self.hide_flash_message)

    def animate_flash_message(self):
        if self.flash_frame.winfo_ismapped():
            self.flash_frame.config(bg=self.animation_colors[self.current_animation_index])
            self.current_animation_index = (self.current_animation_index + 1) % len(self.animation_colors)
            self.root.after(500, self.animate_flash_message)

    def hide_flash_message(self):
        self.flash_frame.pack_forget()

    def send_data(self):
        name = self.name_entry.get().strip()
        phone = self.phone_entry.get().strip()
        job = self.job_entry.get().strip()
        position = self.position_var.get().strip()
        nationality = self.nationality_entry.get().strip()
        
        if not name:
            messagebox.showerror("خطأ", "يجب إدخال اسم الزائر")
            return
            
        daily_id = self.get_daily_id()
        
        try:
            self.db_cursor.execute('''INSERT INTO visitors 
                                 (daily_id, name, phone, job, position, nationality)
                                 VALUES (?, ?, ?, ?, ?, ?)''',
                                 (daily_id, name, phone, job, position, nationality))
            
            visitor_id = self.db_cursor.lastrowid
            
            self.db_cursor.execute('''INSERT INTO notifications (visitor_id, notification_type)
                                 VALUES (?, 'new_visitor')''', (visitor_id,))
            
            self.db_conn.commit()
            
            self.play_notification_sound()
            self.show_flash_message(f"تم إضافة الزائر: {name}")
            
            self.name_entry.delete(0, tk.END)
            self.phone_entry.delete(0, tk.END)
            self.job_entry.delete(0, tk.END)
            self.position_var.set('')
            self.nationality_entry.delete(0, tk.END)
            
            self.load_data()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ البيانات: {str(e)}")

    def search_by_name(self):
        name = simpledialog.askstring("البحث بالاسم", "أدخل اسم الزائر للبحث:")
        if name:
            self.current_filters = {'name': name, 'date': None}
            self.load_data()
            self.show_flash_message(f"عرض نتائج البحث عن: {name}")

    def search_by_date(self):
        date_window = tk.Toplevel(self.root)
        date_window.title("اختر تاريخ البحث")
        date_window.geometry("300x300")
        date_window.resizable(False, False)
        
        cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)
        
        def set_date():
            selected_date = cal.get_date()
            self.current_filters = {'name': None, 'date': selected_date}
            self.load_data()
            date_window.destroy()
            self.show_flash_message(f"عرض الزوار لتاريخ: {selected_date}")
        
        tk.Button(date_window, text="تأكيد", command=set_date,
                font=('Arial', 12), bg='#3498db', fg='white').pack(pady=10)

    def clear_filters(self):
        self.current_filters = {'name': None, 'date': None}
        self.show_hidden_flag = False
        self.load_data()
        self.show_flash_message("تم مسح عوامل التصفية")

    def show_comprehensive_stats(self):
        stats_window = tk.Toplevel(self.root)
        stats_window.title("الإحصائيات الشاملة")
        stats_window.geometry("1000x700")
        
        stats = self.load_comprehensive_stats()
        
        frame = tk.Frame(stats_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tree = ttk.Treeview(frame, columns=("stat", "value"), show="headings")
        tree.heading("stat", text="الإحصائية")
        tree.heading("value", text="القيمة")
        tree.column("stat", width=400)
        tree.column("value", width=400)
        
        tree.insert("", tk.END, values=("إجمالي عدد الزوار", stats["إجمالي عدد الزوار"]))
        tree.insert("", tk.END, values=("عدد الزوار اليوم", stats["عدد الزوار اليوم"]))
        
        statuses = ['انتظار', 'موافق', 'رفض', 'انتهاء']
        for status in statuses:
            tree.insert("", tk.END, values=(f"عدد الزوار ({status})", stats[f"عدد الزوار ({status})"]))
        
        tree.insert("", tk.END, values=("عدد المواعيد المسبقة", stats["عدد المواعيد المسبقة"]))
        tree.insert("", tk.END, values=("مواعيد معلقة", stats["مواعيد معلقة"]))
        tree.insert("", tk.END, values=("مواعيد مكتملة", stats["مواعيد مكتملة"]))
        
        tree.insert("", tk.END, values=("عدد الالتزامات", stats["عدد الالتزامات"]))
        tree.insert("", tk.END, values=("التزامات معلقة", stats["التزامات معلقة"]))
        tree.insert("", tk.END, values=("التزامات مكتملة", stats["التزامات مكتملة"]))
        
        for month, count in stats["إحصائيات شهرية"]:
            tree.insert("", tk.END, values=(f"عدد الزيارات في {month}", count))
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        export_btn = tk.Button(stats_window, text="تصدير إلى Excel", 
                              command=lambda: self.export_comprehensive_stats(stats),
                              font=('Arial', 12), bg='#3498db', fg='white')
        export_btn.pack(pady=10)

    def load_comprehensive_stats(self):
        stats = {}
        
        # إحصائيات الزوار
        self.db_cursor.execute("SELECT COUNT(*) FROM visitors")
        stats["إجمالي عدد الزوار"] = self.db_cursor.fetchone()[0]
        
        today = date.today().strftime("%Y-%m-%d")
        self.db_cursor.execute("SELECT COUNT(*) FROM visitors WHERE date(visit_date)=?", (today,))
        stats["عدد الزوار اليوم"] = self.db_cursor.fetchone()[0]
        
        statuses = ['انتظار', 'موافق', 'رفض', 'انتهاء']
        for status in statuses:
            self.db_cursor.execute("SELECT COUNT(*) FROM visitors WHERE status=?", (status,))
            stats[f"عدد الزوار ({status})"] = self.db_cursor.fetchone()[0]
        
        # إحصائيات المواعيد
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments")
        stats["عدد المواعيد المسبقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments WHERE status='معلقة'")
        stats["مواعيد معلقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments WHERE status='مكتملة'")
        stats["مواعيد مكتملة"] = self.db_cursor.fetchone()[0]
        
        # إحصائيات الالتزامات
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments")
        stats["عدد الالتزامات"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments WHERE status='معلقة'")
        stats["التزامات معلقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments WHERE status='مكتملة'")
        stats["التزامات مكتملة"] = self.db_cursor.fetchone()[0]
        
        # إحصائيات شهرية
        self.db_cursor.execute('''SELECT strftime('%Y-%m', visit_date) as month, COUNT(*) 
                               FROM visitors GROUP BY month ORDER BY month DESC''')
        stats["إحصائيات شهرية"] = self.db_cursor.fetchall()
        
        return stats

    def export_comprehensive_stats(self, stats):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="حفظ ملف الإحصائيات")
        if not file_path:
            return
        
        try:
            data = []
            data.append(["إحصائيات الزوار", ""])
            data.append(["إجمالي عدد الزوار", stats["إجمالي عدد الزوار"]])
            data.append(["عدد الزوار اليوم", stats["عدد الزوار اليوم"]])
            
            statuses = ['انتظار', 'موافق', 'رفض', 'انتهاء']
            for status in statuses:
                data.append([f"عدد الزوار ({status})", stats[f"عدد الزوار ({status})"]])
            
            data.append(["", ""])
            data.append(["إحصائيات المواعيد المسبقة", ""])
            data.append(["عدد المواعيد المسبقة", stats["عدد المواعيد المسبقة"]])
            data.append(["مواعيد معلقة", stats["مواعيد معلقة"]])
            data.append(["مواعيد مكتملة", stats["مواعيد مكتملة"]])
            
            data.append(["", ""])
            data.append(["إحصائيات الالتزامات", ""])
            data.append(["عدد الالتزامات", stats["عدد الالتزامات"]])
            data.append(["التزامات معلقة", stats["التزامات معلقة"]])
            data.append(["التزامات مكتملة", stats["التزامات مكتملة"]])
            
            data.append(["", ""])
            data.append(["الإحصائيات الشهرية", ""])
            for month, count in stats["إحصائيات شهرية"]:
                data.append([f"عدد الزيارات في {month}", count])
            
            df = pd.DataFrame(data, columns=["الإحصائية", "القيمة"])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("نجاح", f"تم تصدير البيانات إلى {file_path}")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء التصدير: {str(e)}")

    def show_database_options(self):
        options_window = tk.Toplevel(self.root)
        options_window.title("خيارات قاعدة البيانات")
        options_window.geometry("400x300")
        
        tk.Label(options_window, text="خيارات قاعدة البيانات", font=('Arial', 16, 'bold')).pack(pady=10)
        
        btn_style = {'font': ('Arial', 12), 'width': 20, 'pady': 10}
        
        tk.Button(options_window, text="عرض جميع الزوار", command=self.show_all_visitors, **btn_style).pack(pady=5)
        tk.Button(options_window, text="تصدير إلى Excel", command=self.export_to_excel, **btn_style).pack(pady=5)
        tk.Button(options_window, text="بحث بالاسم", command=self.show_db_name_search, **btn_style).pack(pady=5)
        tk.Button(options_window, text="بحث بالتاريخ", command=self.show_db_date_search, **btn_style).pack(pady=5)
        
        if self.username == 'it':
            tk.Button(options_window, text="حذف جميع الزوار", command=self.delete_all_visitors, **btn_style).pack(pady=5)

    def show_all_visitors(self):
        self.show_hidden_flag = True
        self.load_data()
        messagebox.showinfo("عرض الكل", "يتم الآن عرض جميع الزوار بما فيهم المخفيين")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="حفظ ملف الزوار")
        if not file_path:
            return
        
        try:
            query = "SELECT * FROM visitors"
            df = pd.read_sql_query(query, self.db_conn)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("نجاح", f"تم تصدير البيانات إلى {file_path}")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء التصدير: {str(e)}")

    def show_db_name_search(self):
        name = simpledialog.askstring("بحث بالاسم", "أدخل اسم الزائر:")
        if name:
            self.current_filters['name'] = name
            self.load_data()

    def show_db_date_search(self):
        date_window = tk.Toplevel(self.root)
        date_window.title("اختر تاريخ")
        
        cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)
        
        def set_date():
            selected_date = cal.get_date()
            self.current_filters['date'] = selected_date
            self.load_data()
            date_window.destroy()
        
        tk.Button(date_window, text="تأكيد", command=set_date).pack(pady=5)

    def delete_all_visitors(self):
        if self.username != 'it':
            messagebox.showerror("خطأ", "ليس لديك صلاحية حذف جميع الزوار")
            return
            
        confirm = messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف جميع سجلات الزوار؟")
        if confirm:
            try:
                self.db_cursor.execute("DELETE FROM visitors")
                self.db_conn.commit()
                self.load_data()
                messagebox.showinfo("نجاح", "تم حذف جميع سجلات الزوار بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")

    def hide_visitor(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار زائر لإخفائه")
            return
            
        item = self.tree.item(selected_item)
        visitor_id = item['values'][0]
        visitor_name = item['values'][2]
        
        try:
            self.db_cursor.execute("UPDATE visitors SET hidden=1 WHERE id=?", (visitor_id,))
            self.db_conn.commit()
            
            self.play_notification_sound()
            self.show_flash_message(f"تم إخفاء الزائر: {visitor_name}")
            self.load_data()
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء إخفاء الزائر: {str(e)}")

    def finish_visit(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار زائر لإنهاء الزيارة")
            return
        
        item = self.tree.item(selected_item)
        visitor_id = item['values'][0]
        visitor_name = item['values'][2]
        
        if item['values'][8] != 'موافق':
            messagebox.showerror("خطأ", "يجب الموافقة على الزائر أولاً قبل إنهاء الزيارة")
            return
        
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            self.db_cursor.execute('''UPDATE visitors 
                         SET status=?, exit_time=?, new_flag=0, hidden=1
                         WHERE id=?''', ('انتهاء', now, visitor_id))
            
            self.db_cursor.execute('''INSERT INTO notifications (visitor_id, notification_type)
                         VALUES (?, 'status_change')''', (visitor_id,))
            
            self.db_conn.commit()
            
            self.play_notification_sound()
            self.show_flash_message(f"تم إنهاء زيارة الزائر: {visitor_name}")
            self.load_data()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء إنهاء الزيارة: {str(e)}")

    def add_new_position(self):
        new_position = simpledialog.askstring("إضافة درجة وظيفية", "أدخل اسم الدرجة الوظيفية الجديدة:")
        if new_position and new_position.strip():
            new_position = new_position.strip()
            if new_position not in self.position_options:
                try:
                    self.db_cursor.execute("INSERT INTO positions (position_name) VALUES (?)", (new_position,))
                    self.db_conn.commit()
                    
                    self.position_options.append(new_position)
                    self.position_menu['values'] = self.position_options
                    self.position_var.set(new_position)
                    
                    messagebox.showinfo("نجاح", "تمت إضافة الدرجة الوظيفية بنجاح")
                except sqlite3.IntegrityError:
                    messagebox.showerror("خطأ", "هذه الدرجة الوظيفية موجودة بالفعل")
            else:
                messagebox.showwarning("تحذير", "هذه الدرجة الوظيفية موجودة بالفعل")

    def edit_position(self):
        current_position = self.position_var.get()
        if not current_position:
            messagebox.showwarning("تحذير", "يجب اختيار درجة وظيفية لتعديلها")
            return
            
        new_position = simpledialog.askstring("تعديل درجة وظيفية", "أدخل الاسم الجديد للدرجة الوظيفية:", initialvalue=current_position)
        if new_position and new_position.strip() and new_position != current_position:
            new_position = new_position.strip()
            try:
                self.db_cursor.execute("UPDATE positions SET position_name=? WHERE position_name=?", 
                                     (new_position, current_position))
                
                self.db_cursor.execute("UPDATE visitors SET position=? WHERE position=?", 
                                     (new_position, current_position))
                
                self.db_cursor.execute("UPDATE appointments SET position=? WHERE position=?", 
                                     (new_position, current_position))
                
                self.db_conn.commit()
                
                index = self.position_options.index(current_position)
                self.position_options[index] = new_position
                self.position_menu['values'] = self.position_options
                self.position_var.set(new_position)
                
                messagebox.showinfo("نجاح", "تم تعديل الدرجة الوظيفية بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء التعديل: {str(e)}")

    def delete_position(self):
        position = self.position_var.get()
        if not position:
            messagebox.showwarning("تحذير", "يجب اختيار درجة وظيفية لحذفها")
            return
            
        if messagebox.askyesno("تأكيد الحذف", f"هل أنت متأكد من حذف الدرجة الوظيفية '{position}'؟"):
            try:
                self.db_cursor.execute("DELETE FROM positions WHERE position_name=?", (position,))
                
                self.db_conn.commit()
                
                self.position_options.remove(position)
                self.position_menu['values'] = self.position_options
                self.position_var.set('')
                
                messagebox.showinfo("نجاح", "تم حذف الدرجة الوظيفية بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")

    def toggle_hidden_visitors(self):
        self.show_hidden_flag = not self.show_hidden_flag
        self.load_data()
        
        if self.show_hidden_flag:
            self.show_flash_message("عرض الزوار المخفيين")
        else:
            self.show_flash_message("إخفاء الزوار المخفيين")

    def edit_visitor(self):
        if not self.selected_visitor:
            messagebox.showerror("خطأ", "يجب اختيار زائر لتعديل بياناته")
            return
            
        visitor_id = self.selected_visitor[0]
        
        edit_window = tk.Toplevel(self.root)
        edit_window.title("تعديل بيانات الزائر")
        edit_window.geometry("400x400")
        edit_window.resizable(False, False)
        
        tk.Label(edit_window, text="اسم الزائر:", font=('Arial', 12)).pack(pady=5)
        name_entry = tk.Entry(edit_window, font=('Arial', 12))
        name_entry.insert(0, self.selected_visitor[2])
        name_entry.pack(pady=5)
        
        tk.Label(edit_window, text="رقم الهاتف:", font=('Arial', 12)).pack(pady=5)
        phone_entry = tk.Entry(edit_window, font=('Arial', 12))
        phone_entry.insert(0, self.selected_visitor[3])
        phone_entry.pack(pady=5)
        
        tk.Label(edit_window, text="الوظيفة:", font=('Arial', 12)).pack(pady=5)
        job_entry = tk.Entry(edit_window, font=('Arial', 12))
        job_entry.insert(0, self.selected_visitor[4])
        job_entry.pack(pady=5)
        
        tk.Label(edit_window, text="الدرجة الوظيفية:", font=('Arial', 12)).pack(pady=5)
        position_var = tk.StringVar(value=self.selected_visitor[5])
        position_menu = ttk.Combobox(edit_window, textvariable=position_var, 
                                    values=self.position_options, font=('Arial', 12))
        position_menu.pack(pady=5)
        
        tk.Label(edit_window, text="الجنسية:", font=('Arial', 12)).pack(pady=5)
        nationality_entry = tk.Entry(edit_window, font=('Arial', 12))
        nationality_entry.insert(0, self.selected_visitor[6])
        nationality_entry.pack(pady=5)
        
        def save_changes():
            name = name_entry.get().strip()
            phone = phone_entry.get().strip()
            job = job_entry.get().strip()
            position = position_var.get().strip()
            nationality = nationality_entry.get().strip()
            
            if not name:
                messagebox.showerror("خطأ", "يجب إدخال اسم الزائر")
                return
                
            try:
                self.db_cursor.execute('''UPDATE visitors 
                                     SET name=?, phone=?, job=?, position=?, nationality=?
                                     WHERE id=?''',
                                     (name, phone, job, position, nationality, visitor_id))
                
                self.db_conn.commit()
                
                messagebox.showinfo("نجاح", "تم تعديل بيانات الزائر بنجاح")
                edit_window.destroy()
                self.load_data()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء التعديل: {str(e)}")
        
        tk.Button(edit_window, text="حفظ التعديلات", command=save_changes,
                font=('Arial', 12), bg='#3498db', fg='white').pack(pady=10)

    def delete_visitor(self):
        if not self.selected_visitor:
            messagebox.showerror("خطأ", "يجب اختيار زائر لحذفه")
            return
            
        visitor_id = self.selected_visitor[0]
        visitor_name = self.selected_visitor[2]
        
        if messagebox.askyesno("تأكيد الحذف", f"هل أنت متأكد من حذف الزائر '{visitor_name}'؟"):
            try:
                self.db_cursor.execute("DELETE FROM visitors WHERE id=?", (visitor_id,))
                self.db_conn.commit()
                
                self.play_notification_sound()
                self.show_flash_message(f"تم حذف الزائر: {visitor_name}")
                self.load_data()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")

    def show_appointment_dialog(self):
        appt_window = tk.Toplevel(self.root)
        appt_window.title("حجز موعد مسبق")
        appt_window.geometry("600x600")
        appt_window.resizable(False, False)
        appt_window.configure(bg='#f5f5f5')
        
        info_frame = tk.LabelFrame(appt_window, text="بيانات الزائر", bg='#f5f5f5', padx=10, pady=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(info_frame, text="اسم الزائر:", font=('Arial', 12), bg='#f5f5f5').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.appt_name_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.appt_name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="رقم الهاتف:", font=('Arial', 12), bg='#f5f5f5').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.appt_phone_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.appt_phone_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="الوظيفة:", font=('Arial', 12), bg='#f5f5f5').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.appt_job_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.appt_job_entry.grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="الدرجة الوظيفية:", font=('Arial', 12), bg='#f5f5f5').grid(row=3, column=0, sticky='e', padx=5, pady=5)
        self.appt_position_var = tk.StringVar()
        self.appt_position_menu = ttk.Combobox(info_frame, textvariable=self.appt_position_var, 
                                             values=self.position_options, width=27, font=('Arial', 12))
        self.appt_position_menu.grid(row=3, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="الجنسية:", font=('Arial', 12), bg='#f5f5f5').grid(row=4, column=0, sticky='e', padx=5, pady=5)
        self.appt_nationality_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.appt_nationality_entry.grid(row=4, column=1, padx=5, pady=5)
        
        datetime_frame = tk.LabelFrame(appt_window, text="التاريخ والوقت", bg='#f5f5f5', padx=10, pady=10)
        datetime_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(datetime_frame, text="تاريخ الموعد:", font=('Arial', 12), bg='#f5f5f5').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.appt_cal = Calendar(datetime_frame, selectmode='day', date_pattern='yyyy-mm-dd', font=('Arial', 12))
        self.appt_cal.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(datetime_frame, text="وقت الموعد:", font=('Arial', 12), bg='#f5f5f5').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        
        time_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
        time_frame.grid(row=1, column=1, padx=5, pady=5)
        
        self.appt_hour_var = tk.StringVar(value="08")
        self.appt_minute_var = tk.StringVar(value="00")
        self.appt_ampm_var = tk.StringVar(value="AM")
        
        tk.Label(time_frame, text="ساعة:", font=('Arial', 10), bg='#f5f5f5').pack(side=tk.LEFT)
        self.appt_hour_menu = ttk.Combobox(time_frame, textvariable=self.appt_hour_var, 
                                         values=[f"{i:02d}" for i in range(1, 13)], width=3)
        self.appt_hour_menu.pack(side=tk.LEFT, padx=5)
        
        tk.Label(time_frame, text="دقيقة:", font=('Arial', 10), bg='#f5f5f5').pack(side=tk.LEFT)
        self.appt_minute_menu = ttk.Combobox(time_frame, textvariable=self.appt_minute_var, 
                                           values=[f"{i:02d}" for i in range(0, 60, 5)], width=3)
        self.appt_minute_menu.pack(side=tk.LEFT, padx=5)
        
        self.appt_ampm_menu = ttk.Combobox(time_frame, textvariable=self.appt_ampm_var, 
                                         values=["AM", "PM"], width=3)
        self.appt_ampm_menu.pack(side=tk.LEFT, padx=5)
        
        btn_frame = tk.Frame(appt_window, bg='#f5f5f5')
        btn_frame.pack(pady=10)
        
        save_btn = tk.Button(btn_frame, text="حفظ الموعد", command=self.save_appointment,
                font=('Arial', 12), bg='#2ecc71', fg='white', padx=15, pady=5)
        save_btn.pack(side=tk.LEFT, padx=10)
        
        edit_btn = tk.Button(btn_frame, text="تعديل", command=self.edit_appointment,
                font=('Arial', 12), bg='#f39c12', fg='white', padx=15, pady=5)
        edit_btn.pack(side=tk.LEFT, padx=10)
        
        delete_btn = tk.Button(btn_frame, text="حذف", command=self.delete_appointment,
                font=('Arial', 12), bg='#e74c3c', fg='white', padx=15, pady=5)
        delete_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = tk.Button(btn_frame, text="إلغاء", command=appt_window.destroy,
                font=('Arial', 12), bg='#95a5a6', fg='white', padx=15, pady=5)
        cancel_btn.pack(side=tk.LEFT, padx=10)

    def save_appointment(self):
        name = self.appt_name_entry.get().strip()
        phone = self.appt_phone_entry.get().strip()
        job = self.appt_job_entry.get().strip()
        position = self.appt_position_var.get().strip()
        nationality = self.appt_nationality_entry.get().strip()
        date = self.appt_cal.get_date()
        
        # Convert time to 24-hour format
        hour = int(self.appt_hour_var.get())
        if self.appt_ampm_var.get() == "PM" and hour < 12:
            hour += 12
        elif self.appt_ampm_var.get() == "AM" and hour == 12:
            hour = 0
        
        time = f"{hour:02d}:{self.appt_minute_var.get()}"
        
        if not name:
            messagebox.showerror("خطأ", "يجب إدخال اسم الزائر")
            return
            
        try:
            self.db_cursor.execute('''INSERT INTO appointments 
                                 (visitor_name, phone, job, position, nationality, 
                                  appointment_date, appointment_time)
                                 VALUES (?, ?, ?, ?, ?, ?, ?)''',
                                 (name, phone, job, position, nationality, date, time))
            
            self.db_conn.commit()
            
            daily_id = self.get_daily_id()
            visit_datetime = f"{date} {time}:00"
            
            self.db_cursor.execute('''INSERT INTO visitors 
                                 (daily_id, name, phone, job, position, nationality, 
                                  visit_date, status, new_flag, appointment_time, hidden)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                 (daily_id, name, phone, job, position, nationality, 
                                  visit_datetime, 'انتظار', 1, time, 1))  # Hidden until appointment date
            
            visitor_id = self.db_cursor.lastrowid
            
            self.db_cursor.execute('''INSERT INTO notifications (visitor_id, notification_type)
                                 VALUES (?, 'new_appointment')''', (visitor_id,))
            
            self.db_conn.commit()
            
            messagebox.showinfo("نجاح", "تم حجز الموعد بنجاح")
            self.load_data()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ الموعد: {str(e)}")

    def edit_appointment(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار موعد لتعديله")
            return
            
        item = self.tree.item(selected_item)
        visitor_id = item['values'][0]
        
        self.db_cursor.execute('''SELECT name, phone, job, position, nationality, 
                              strftime('%Y-%m-%d', visit_date) as date, 
                              appointment_time FROM visitors WHERE id=?''', (visitor_id,))
        visitor_data = self.db_cursor.fetchone()
        
        if not visitor_data:
            messagebox.showerror("خطأ", "لم يتم العثور على بيانات الموعد")
            return
            
        name, phone, job, position, nationality, date, time = visitor_data
        
        if not time:
            messagebox.showerror("خطأ", "هذا الزائر ليس لديه موعد مسبق")
            return
            
        # Fill the form with existing data
        self.appt_name_entry.delete(0, tk.END)
        self.appt_name_entry.insert(0, name)
        
        self.appt_phone_entry.delete(0, tk.END)
        self.appt_phone_entry.insert(0, phone)
        
        self.appt_job_entry.delete(0, tk.END)
        self.appt_job_entry.insert(0, job)
        
        self.appt_position_var.set(position)
        
        self.appt_nationality_entry.delete(0, tk.END)
        self.appt_nationality_entry.insert(0, nationality)
        
        self.appt_cal.set_date(date)
        
        # Convert time to 12-hour format
        hour, minute = map(int, time.split(':'))
        ampm = "AM" if hour < 12 else "PM"
        if hour > 12:
            hour -= 12
        elif hour == 0:
            hour = 12
            
        self.appt_hour_var.set(f"{hour:02d}")
        self.appt_minute_var.set(f"{minute:02d}")
        self.appt_ampm_var.set(ampm)
        
        def update_appointment():
            new_name = self.appt_name_entry.get().strip()
            new_phone = self.appt_phone_entry.get().strip()
            new_job = self.appt_job_entry.get().strip()
            new_position = self.appt_position_var.get().strip()
            new_nationality = self.appt_nationality_entry.get().strip()
            new_date = self.appt_cal.get_date()
            
            # Convert time to 24-hour format
            new_hour = int(self.appt_hour_var.get())
            if self.appt_ampm_var.get() == "PM" and new_hour < 12:
                new_hour += 12
            elif self.appt_ampm_var.get() == "AM" and new_hour == 12:
                new_hour = 0
            
            new_time = f"{new_hour:02d}:{self.appt_minute_var.get()}"
            
            if not new_name:
                messagebox.showerror("خطأ", "يجب إدخال اسم الزائر")
                return
                
            try:
                # Update in visitors table
                self.db_cursor.execute('''UPDATE visitors 
                                     SET name=?, phone=?, job=?, position=?, nationality=?,
                                     visit_date=?, appointment_time=?
                                     WHERE id=?''',
                                     (new_name, new_phone, new_job, new_position, new_nationality,
                                      f"{new_date} {new_time}:00", new_time, visitor_id))
                
                # Update in appointments table
                self.db_cursor.execute('''UPDATE appointments 
                                     SET visitor_name=?, phone=?, job=?, position=?, nationality=?,
                                     appointment_date=?, appointment_time=?
                                     WHERE visitor_name=? AND appointment_date=? AND appointment_time=?''',
                                     (new_name, new_phone, new_job, new_position, new_nationality,
                                      new_date, new_time, name, date, time))
                
                self.db_conn.commit()
                
                messagebox.showinfo("نجاح", "تم تعديل الموعد بنجاح")
                self.load_data()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء التعديل: {str(e)}")
        
        # Change save button to update
        for widget in appt_window.winfo_children():
            if isinstance(widget, tk.Button) and widget['text'] == "حفظ الموعد":
                widget.config(command=update_appointment, text="تحديث الموعد")
                break

    def delete_appointment(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار موعد لحذفه")
            return
            
        item = self.tree.item(selected_item)
        visitor_id = item['values'][0]
        visitor_name = item['values'][2]
        
        if messagebox.askyesno("تأكيد الحذف", f"هل أنت متأكد من حذف موعد الزائر '{visitor_name}'؟"):
            try:
                # Get appointment details
                self.db_cursor.execute('''SELECT name, strftime('%Y-%m-%d', visit_date) as date, 
                                      appointment_time FROM visitors WHERE id=?''', (visitor_id,))
                appointment_data = self.db_cursor.fetchone()
                
                if not appointment_data or not appointment_data[2]:
                    messagebox.showerror("خطأ", "لم يتم العثور على بيانات الموعد")
                    return
                    
                name, date, time = appointment_data
                
                # Delete from appointments table
                self.db_cursor.execute('''DELETE FROM appointments 
                                      WHERE visitor_name=? AND appointment_date=? AND appointment_time=?''',
                                      (name, date, time))
                
                # Remove appointment info from visitors table
                self.db_cursor.execute('''UPDATE visitors SET appointment_time=NULL 
                                      WHERE id=?''', (visitor_id,))
                
                self.db_conn.commit()
                
                messagebox.showinfo("نجاح", "تم حذف الموعد بنجاح")
                self.load_data()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")

    def show_appointments_list(self):
        appointments_window = tk.Toplevel(self.root)
        appointments_window.title("قائمة المواعيد المسبقة")
        appointments_window.geometry("800x600")
        appointments_window.configure(bg='#f5f5f5')
        
        frame = tk.Frame(appointments_window, bg='#f5f5f5')
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("id", "visitor_name", "phone", "job", "position", "nationality", "appointment_date", "appointment_time", "status")
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=20)
        
        tree.heading("id", text="ID", anchor='center')
        tree.heading("visitor_name", text="اسم الزائر", anchor='center')
        tree.heading("phone", text="رقم الهاتف", anchor='center')
        tree.heading("job", text="الوظيفة", anchor='center')
        tree.heading("position", text="الدرجة الوظيفية", anchor='center')
        tree.heading("nationality", text="الجنسية", anchor='center')
        tree.heading("appointment_date", text="تاريخ الموعد", anchor='center')
        tree.heading("appointment_time", text="وقت الموعد", anchor='center')
        tree.heading("status", text="الحالة", anchor='center')
        
        tree.column("id", width=50, anchor='center')
        tree.column("visitor_name", width=150, anchor='center')
        tree.column("phone", width=120, anchor='center')
        tree.column("job", width=120, anchor='center')
        tree.column("position", width=150, anchor='center')
        tree.column("nationality", width=100, anchor='center')
        tree.column("appointment_date", width=120, anchor='center')
        tree.column("appointment_time", width=100, anchor='center')
        tree.column("status", width=100, anchor='center')
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.db_cursor.execute('''SELECT id, visitor_name, phone, job, position, 
                              nationality, appointment_date, appointment_time, status
                              FROM appointments ORDER BY appointment_date, appointment_time''')
        appointments = self.db_cursor.fetchall()
        
        for appt in appointments:
            tree.insert("", tk.END, values=appt)
        
        btn_frame = tk.Frame(appointments_window, bg='#f5f5f5')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="إغلاق", command=appointments_window.destroy,
                font=('Arial', 12), bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=10)

    def show_commitments_dialog(self):
        commitments_window = tk.Toplevel(self.root)
        commitments_window.title("إدارة الالتزامات")
        commitments_window.geometry("600x500")
        commitments_window.configure(bg='#f5f5f5')
        
        info_frame = tk.LabelFrame(commitments_window, text="بيانات الالتزام", bg='#f5f5f5', padx=10, pady=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(info_frame, text="المكان:", font=('Arial', 12), bg='#f5f5f5').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.commit_place_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.commit_place_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="الغرض من الزيارة:", font=('Arial', 12), bg='#f5f5f5').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.commit_purpose_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.commit_purpose_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="الجهة:", font=('Arial', 12), bg='#f5f5f5').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.commit_department_entry = tk.Entry(info_frame, width=30, font=('Arial', 12))
        self.commit_department_entry.grid(row=2, column=1, padx=5, pady=5)
        
        datetime_frame = tk.LabelFrame(commitments_window, text="التاريخ والوقت", bg='#f5f5f5', padx=10, pady=10)
        datetime_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(datetime_frame, text="تاريخ الالتزام:", font=('Arial', 12), bg='#f5f5f5').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.commit_cal = Calendar(datetime_frame, selectmode='day', date_pattern='yyyy-mm-dd', font=('Arial', 12))
        self.commit_cal.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(datetime_frame, text="وقت الالتزام:", font=('Arial', 12), bg='#f5f5f5').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        
        time_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
        time_frame.grid(row=1, column=1, padx=5, pady=5)
        
        self.commit_hour_var = tk.StringVar(value="08")
        self.commit_minute_var = tk.StringVar(value="00")
        self.commit_ampm_var = tk.StringVar(value="AM")
        
        tk.Label(time_frame, text="ساعة:", font=('Arial', 10), bg='#f5f5f5').pack(side=tk.LEFT)
        self.commit_hour_menu = ttk.Combobox(time_frame, textvariable=self.commit_hour_var, 
                                           values=[f"{i:02d}" for i in range(1, 13)], width=3)
        self.commit_hour_menu.pack(side=tk.LEFT, padx=5)
        
        tk.Label(time_frame, text="دقيقة:", font=('Arial', 10), bg='#f5f5f5').pack(side=tk.LEFT)
        self.commit_minute_menu = ttk.Combobox(time_frame, textvariable=self.commit_minute_var, 
                                             values=[f"{i:02d}" for i in range(0, 60, 5)], width=3)
        self.commit_minute_menu.pack(side=tk.LEFT, padx=5)
        
        self.commit_ampm_menu = ttk.Combobox(time_frame, textvariable=self.commit_ampm_var, 
                                           values=["AM", "PM"], width=3)
        self.commit_ampm_menu.pack(side=tk.LEFT, padx=5)
        
        btn_frame = tk.Frame(commitments_window, bg='#f5f5f5')
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="حفظ الالتزام", command=self.save_commitment,
                font=('Arial', 12), bg='#2ecc71', fg='white', padx=15, pady=5).pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="إلغاء", command=commitments_window.destroy,
                font=('Arial', 12), bg='#e74c3c', fg='white', padx=15, pady=5).pack(side=tk.LEFT, padx=10)
        
        table_frame = tk.Frame(commitments_window, bg='#f5f5f5')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("id", "place", "purpose", "department", "date", "time", "status")
        self.commit_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=5)
        
        self.commit_tree.heading("id", text="ID", anchor='center')
        self.commit_tree.heading("place", text="المكان", anchor='center')
        self.commit_tree.heading("purpose", text="الغرض", anchor='center')
        self.commit_tree.heading("department", text="الجهة", anchor='center')
        self.commit_tree.heading("date", text="التاريخ", anchor='center')
        self.commit_tree.heading("time", text="الوقت", anchor='center')
        self.commit_tree.heading("status", text="الحالة", anchor='center')
        
        self.commit_tree.column("id", width=50, anchor='center')
        self.commit_tree.column("place", width=100, anchor='center')
        self.commit_tree.column("purpose", width=120, anchor='center')
        self.commit_tree.column("department", width=100, anchor='center')
        self.commit_tree.column("date", width=100, anchor='center')
        self.commit_tree.column("time", width=80, anchor='center')
        self.commit_tree.column("status", width=80, anchor='center')
        
        self.commit_tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.commit_tree.yview)
        self.commit_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.load_commitments()

    def load_commitments(self):
        for item in self.commit_tree.get_children():
            self.commit_tree.delete(item)
            
        self.db_cursor.execute('''SELECT id, place, purpose, department, 
                               commitment_date, commitment_time, status
                               FROM commitments ORDER BY commitment_date, commitment_time''')
        commitments = self.db_cursor.fetchall()
        
        for commit in commitments:
            self.commit_tree.insert("", tk.END, values=commit)

    def save_commitment(self):
        place = self.commit_place_entry.get().strip()
        purpose = self.commit_purpose_entry.get().strip()
        department = self.commit_department_entry.get().strip()
        date = self.commit_cal.get_date()
        
        # Convert time to 24-hour format
        hour = int(self.commit_hour_var.get())
        if self.commit_ampm_var.get() == "PM" and hour < 12:
            hour += 12
        elif self.commit_ampm_var.get() == "AM" and hour == 12:
            hour = 0
        
        time = f"{hour:02d}:{self.commit_minute_var.get()}"
        
        if not place or not purpose or not department:
            messagebox.showerror("خطأ", "يجب إدخال جميع بيانات الالتزام")
            return
            
        try:
            self.db_cursor.execute('''INSERT INTO commitments 
                                 (place, purpose, department, 
                                  commitment_date, commitment_time)
                                 VALUES (?, ?, ?, ?, ?)''',
                                 (place, purpose, department, date, time))
            
            self.db_conn.commit()
            
            messagebox.showinfo("نجاح", "تم حفظ الالتزام بنجاح")
            self.load_commitments()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ الالتزام: {str(e)}")

    def show_commitments_list(self):
        commitments_window = tk.Toplevel(self.root)
        commitments_window.title("قائمة الالتزامات")
        commitments_window.geometry("800x600")
        commitments_window.configure(bg='#f5f5f5')
        
        frame = tk.Frame(commitments_window, bg='#f5f5f5')
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("id", "place", "purpose", "department", "commitment_date", "commitment_time", "status")
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=20)
        
        tree.heading("id", text="ID", anchor='center')
        tree.heading("place", text="المكان", anchor='center')
        tree.heading("purpose", text="الغرض", anchor='center')
        tree.heading("department", text="الجهة", anchor='center')
        tree.heading("commitment_date", text="تاريخ الالتزام", anchor='center')
        tree.heading("commitment_time", text="وقت الالتزام", anchor='center')
        tree.heading("status", text="الحالة", anchor='center')
        
        tree.column("id", width=50, anchor='center')
        tree.column("place", width=150, anchor='center')
        tree.column("purpose", width=200, anchor='center')
        tree.column("department", width=150, anchor='center')
        tree.column("commitment_date", width=120, anchor='center')
        tree.column("commitment_time", width=100, anchor='center')
        tree.column("status", width=100, anchor='center')
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.db_cursor.execute('''SELECT id, place, purpose, department, 
                              commitment_date, commitment_time, status
                              FROM commitments ORDER BY commitment_date, commitment_time''')
        commitments = self.db_cursor.fetchall()
        
        for commit in commitments:
            tree.insert("", tk.END, values=commit)
        
        btn_frame = tk.Frame(commitments_window, bg='#f5f5f5')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="إغلاق", command=commitments_window.destroy,
                font=('Arial', 12), bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=10)

    def show_tasks(self):
        tasks_window = tk.Toplevel(self.root)
        tasks_window.title("المهام")
        tasks_window.geometry("800x600")
        
        self.load_tasks_data(tasks_window)
    
    def load_tasks_data(self, window):
        for widget in window.winfo_children():
            widget.destroy()
        
        tree = ttk.Treeview(window, columns=("id", "visitor", "results", "status", "required_date"), show="headings")
        tree.heading("id", text="ID")
        tree.heading("visitor", text="الزائر")
        tree.heading("results", text="نتائج الزيارة")
        tree.heading("status", text="حالة التنفيذ")
        tree.heading("required_date", text="تاريخ التنفيذ")
        
        tree.column("id", width=50)
        tree.column("visitor", width=150)
        tree.column("results", width=200)
        tree.column("status", width=100)
        tree.column("required_date", width=120)
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.db_cursor.execute('''SELECT vd.id, v.name, vd.visit_results, vd.execution_status, vd.required_date
                              FROM visit_details vd
                              JOIN visitors v ON vd.visitor_id = v.id''')
        tasks = self.db_cursor.fetchall()
        
        for task in tasks:
            tree.insert("", tk.END, values=task)
        
        btn_frame = tk.Frame(window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="عرض التفاصيل", command=lambda: self.show_task_details(tree),
                font=('Arial', 12), bg='#3498db', fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="إغلاق", command=window.destroy,
                font=('Arial', 12), bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)
    
    def show_task_details(self, tree):
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار مهمة لعرض تفاصيلها")
            return
        
        task_id = tree.item(selected_item)['values'][0]
        
        self.db_cursor.execute('''SELECT vd.id, vd.visit_results, vd.recommendations, 
                              vd.required_date, vd.execution_status, v.name
                              FROM visit_details vd
                              JOIN visitors v ON vd.visitor_id = v.id
                              WHERE vd.id=?''', (task_id,))
        task_details = self.db_cursor.fetchone()
        
        if not task_details:
            messagebox.showerror("خطأ", "تعذر العثور على تفاصيل المهمة")
            return
        
        details_window = tk.Toplevel(self.root)
        details_window.title("تفاصيل المهمة")
        details_window.geometry("600x500")
        
        tk.Label(details_window, text=f"الزائر: {task_details[5]}", font=('Arial', 12, 'bold')).pack(pady=5)
        
        tk.Label(details_window, text="نتائج الزيارة:", font=('Arial', 12, 'bold')).pack(pady=5)
        results_text = tk.Text(details_window, height=5, width=60, font=('Arial', 12))
        results_text.insert(tk.END, task_details[1])
        results_text.pack(pady=5)
        
        tk.Label(details_window, text="التوصيات:", font=('Arial', 12, 'bold')).pack(pady=5)
        recommendations_text = tk.Text(details_window, height=5, width=60, font=('Arial', 12))
        recommendations_text.insert(tk.END, task_details[2])
        recommendations_text.pack(pady=5)
        
        tk.Label(details_window, text="تاريخ التنفيذ المطلوب:", font=('Arial', 12)).pack(pady=5)
        required_date_entry = tk.Entry(details_window, font=('Arial', 12))
        required_date_entry.insert(0, task_details[3])
        required_date_entry.pack(pady=5)
        
        tk.Label(details_window, text="حالة التنفيذ:", font=('Arial', 12)).pack(pady=5)
        status_var = tk.StringVar(value=task_details[4])
        status_menu = ttk.Combobox(details_window, textvariable=status_var, 
                                 values=["تحت التنفيذ", "انتهت", "متأخر"], font=('Arial', 12))
        status_menu.pack(pady=5)
        
        if task_details[4] == "تحت التنفيذ" and task_details[3]:
            required_date = datetime.strptime(task_details[3], "%Y-%m-%d").date()
            if date.today() > required_date:
                tk.Label(details_window, text="حالة المهمة: متأخرة", font=('Arial', 12, 'bold'), fg='red').pack(pady=5)
        
        def update_task():
            new_results = results_text.get("1.0", tk.END).strip()
            new_recommendations = recommendations_text.get("1.0", tk.END).strip()
            new_required_date = required_date_entry.get()
            new_status = status_var.get()
            
            try:
                self.db_cursor.execute('''UPDATE visit_details 
                                      SET visit_results=?, recommendations=?, 
                                      required_date=?, execution_status=?
                                      WHERE id=?''',
                                      (new_results, new_recommendations, new_required_date, new_status, task_id))
                self.db_conn.commit()
                messagebox.showinfo("نجاح", "تم تحديث المهمة بنجاح")
                details_window.destroy()
                self.load_tasks_data(details_window.master)
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء التحديث: {str(e)}")
        
        def delete_task():
            if not task_id:
                messagebox.showinfo("معلومة", "لا توجد مهمة لحذفها")
                return
            
            confirm = messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذه المهمة؟")
            if confirm:
                try:
                    self.db_cursor.execute("DELETE FROM visit_details WHERE id=?", (task_id,))
                    self.db_conn.commit()
                    messagebox.showinfo("نجاح", "تم حذف المهمة بنجاح")
                    details_window.destroy()
                except Exception as e:
                    messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف المهمة: {str(e)}")
        
        btn_frame = tk.Frame(details_window)
        btn_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(btn_frame, text="حفظ التعديلات", command=update_task,
                font=('Arial', 12), bg='#28a745', fg='white').pack(side=tk.LEFT, padx=10)
        
        if task_id:
            tk.Button(btn_frame, text="حذف", command=delete_task,
                    font=('Arial', 12), bg='#dc3545', fg='white').pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="إغلاق", command=details_window.destroy,
                font=('Arial', 12), bg='#6c757d', fg='white').pack(side=tk.RIGHT, padx=10)

    def auto_refresh(self):
        if self.auto_refresh_enabled:
            self.load_data()
        self.root.after(30000, self.auto_refresh)


if __name__ == "__main__":
    def login_gui():
        db_conn = sqlite3.connect('visitors_management.db', check_same_thread=False, timeout=10)
        db_cursor = db_conn.cursor()
        
        db_cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                          id INTEGER PRIMARY KEY AUTOINCREMENT,
                          username TEXT UNIQUE,
                          password TEXT)''')
        
        db_cursor.execute("SELECT COUNT(*) FROM users")
        if db_cursor.fetchone()[0] == 0:
            default_users = [
                ("sec", SecretaryApp.hash_password("000000")),
                ("head", SecretaryApp.hash_password("admin")),
                ("it", SecretaryApp.hash_password("admin"))
            ]
            db_cursor.executemany("INSERT INTO users (username, password) VALUES (?, ?)", default_users)
            db_conn.commit()
        
        root = tk.Tk()
        root.title("نظام إدارة الزائرين - تسجيل الدخول")
        root.geometry("600x400")
        root.configure(bg='#f0f5f5')
        
        login_result = {'username': None, 'success': False}
        
        title_frame = tk.Frame(root, bg='#2c3e50', padx=10, pady=10)
        title_frame.pack(fill=tk.X)
        
        try:
            logo_path = os.path.join(os.path.dirname(sys.executable), 'logo.png') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'logo.png')
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((100, 100), Image.LANCZOS)
                logo_image = ImageTk.PhotoImage(logo_img)
                logo_label = tk.Label(title_frame, image=logo_image, bg='#2c3e50')
                logo_label.image = logo_image
                logo_label.pack(side=tk.RIGHT, padx=10)
            else:
                raise FileNotFoundError("ملف اللوجو غير موجود")
        except Exception as e:
            print(f"خطأ في تحميل الشعار: {e}")
            logo_label = tk.Label(title_frame, text="شعار المؤسسة", 
                                font=('Arial', 14), bg='#2c3e50', fg='white')
            logo_label.pack(side=tk.RIGHT, padx=10)
        
        tk.Label(title_frame, text="مرحباً بكم في برنامج إدارة الزائرين", 
                font=('Arial', 18, 'bold'), bg='#2c3e50', fg='white').pack(pady=(20,5))
        tk.Label(title_frame, text="وزارة الإسكان والمرافق والمجتمعات العمرانية", 
                font=('Arial', 14), bg='#2c3e50', fg='white').pack()
        
        input_frame = tk.Frame(root, bg='#f0f5f5', padx=20, pady=20)
        input_frame.pack(expand=True, fill=tk.BOTH)
        
        tk.Label(input_frame, text="اسم المستخدم:", font=('Arial', 14), bg='#f0f5f5').pack(pady=10)
        username_entry = tk.Entry(input_frame, font=('Arial', 14), bd=2, relief='groove')
        username_entry.pack(pady=5, ipady=5)
        
        tk.Label(input_frame, text="كلمة المرور:", font=('Arial', 14), bg='#f0f5f5').pack(pady=10)
        password_entry = tk.Entry(input_frame, show="*", font=('Arial', 14), bd=2, relief='groove')
        password_entry.pack(pady=5, ipady=5)
        
        button_frame = tk.Frame(root, bg='#f0f5f5', padx=10, pady=10)
        button_frame.pack(fill=tk.X)
        
        def try_login():
            username = username_entry.get()
            password = password_entry.get()
            
            if not username or not password:
                messagebox.showerror("خطأ", "يجب إدخال اسم المستخدم وكلمة المرور")
                return
                
            db_cursor.execute("SELECT password FROM users WHERE username=?", (username,))
            result = db_cursor.fetchone()
            
            if result and result[0] == SecretaryApp.hash_password(password):
                login_result['username'] = username
                login_result['success'] = True
                messagebox.showinfo("نجاح", "تم تسجيل الدخول بنجاح")
                root.destroy()
            else:
                messagebox.showerror("خطأ", "اسم المستخدم أو كلمة المرور غير صحيحة")
                username_entry.delete(0, tk.END)
                password_entry.delete(0, tk.END)
                username_entry.focus()
        
        login_btn = tk.Button(button_frame, text="تسجيل الدخول", command=try_login,
                            font=('Arial', 14), bg='#3498db', fg='white',
                            width=15, padx=20, pady=10, bd=0)
        login_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = tk.Button(button_frame, text="إلغاء", command=root.destroy,
                            font=('Arial', 14), bg='#95a5a6', fg='white',
                            width=15, padx=20, pady=10, bd=0)
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        
        root.mainloop()
        db_conn.close()
        return login_result['username'] if login_result['success'] else None

    username = login_gui()
    
    if username:
        root = tk.Tk()
        app = SecretaryApp(root, username)
        root.mainloop()