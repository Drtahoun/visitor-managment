import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from tkcalendar import Calendar
from datetime import datetime, date, timedelta
import sqlite3
import pygame
import os
import time
import pandas as pd
from PIL import Image, ImageTk
import sys
import shutil
from threading import Thread

def login_gui(allowed_users):
    users = {
        "sec": "000000",
        "head": "admin",
        "it": "admin"
    }
    
    login_result = None
    
    def change_password():
        username = username_entry.get()
        current_password = password_entry.get()
        
        if not username or not current_password:
            messagebox.showerror("خطأ", "يجب إدخال اسم المستخدم وكلمة السر الحالية")
            return
            
        if username in allowed_users and users.get(username) == current_password:
            new_password = simpledialog.askstring("تغيير كلمة السر", "أدخل كلمة السر الجديدة:", show='*')
            if new_password:
                confirm_password = simpledialog.askstring("تغيير كلمة السر", "أكد كلمة السر الجديدة:", show='*')
                if new_password == confirm_password:
                    users[username] = new_password
                    messagebox.showinfo("نجاح", "تم تغيير كلمة السر بنجاح")
                else:
                    messagebox.showerror("خطأ", "كلمة السر غير متطابقة")
        else:
            messagebox.showerror("خطأ", "اسم المستخدم أو كلمة السر غير صحيحة")
    
    def perform_login():
        nonlocal login_result
        username = username_entry.get()
        password = password_entry.get()
        
        if not username or not password:
            messagebox.showerror("خطأ", "يجب إدخال اسم المستخدم وكلمة السر")
            return
            
        if username in allowed_users and users.get(username) == password:
            login_result = username
            root.destroy()
        else:
            messagebox.showerror("خطأ", "اسم المستخدم أو كلمة السر غير صحيحة")
            username_entry.delete(0, tk.END)
            password_entry.delete(0, tk.END)
            username_entry.focus()
    
    root = tk.Tk()
    root.title("نظام إدارة الزائرين")
    root.geometry("800x750")
    root.configure(bg='#f0f8ff')
    
    logo_frame = tk.Frame(root, bg='#f0f8ff')
    logo_frame.pack(pady=20)
    
    try:
        logo_path = os.path.join(os.path.dirname(sys.executable), 'logo.png') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'logo.png')
        if os.path.exists(logo_path):
            logo_img = Image.open(logo_path)
            logo_img = logo_img.resize((150, 150), Image.LANCZOS)
            logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(logo_frame, image=logo, bg='#f0f8ff')
            logo_label.image = logo
            logo_label.pack()
        else:
            raise FileNotFoundError("ملف اللوجو غير موجود")
    except Exception as e:
        print(f"Error loading logo: {e}")
        tk.Label(logo_frame, text="شعار المؤسسة", font=('Arial', 16), bg='#f0f8ff').pack()
    
    title_frame = tk.Frame(root, bg='#2c3e50', height=150)
    title_frame.pack(fill=tk.X)
    
    tk.Label(title_frame, text="مرحباً بكم في برنامج إدارة الزائرين", 
            font=('Arial', 24, 'bold'), bg='#2c3e50', fg='white').pack(pady=(20,5))
    tk.Label(title_frame, text="وزارة الإسكان والمرافق والمجتمعات العمرانية", 
            font=('Arial', 18), bg='#2c3e50', fg='white').pack()
    tk.Label(title_frame, text="الجهاز المركزي للتعمير", 
            font=('Arial', 16), bg='#2c3e50', fg='white').pack(pady=(0,20))
    
    login_frame = tk.Frame(root, bg='#f0f8ff', padx=20, pady=20)
    login_frame.pack(pady=20)
    
    tk.Label(login_frame, text="اسم المستخدم:", font=('Arial', 16), bg='#f0f8ff').grid(row=0, column=0, pady=10, sticky='e')
    username_entry = tk.Entry(login_frame, font=('Arial', 16), width=25, bd=2, relief='groove')
    username_entry.grid(row=0, column=1, pady=10, padx=10)
    
    tk.Label(login_frame, text="كلمة السر:", font=('Arial', 16), bg='#f0f8ff').grid(row=1, column=0, pady=10, sticky='e')
    password_entry = tk.Entry(login_frame, font=('Arial', 16), width=25, show="*", bd=2, relief='groove')
    password_entry.grid(row=1, column=1, pady=10, padx=10)
    
    button_frame = tk.Frame(root, bg='#f0f8ff')
    button_frame.pack(pady=20)
    
    login_btn = tk.Button(button_frame, text="دخول", command=perform_login,
                         font=('Arial', 14, 'bold'), bg='#2ecc71', fg='white',
                         width=15, height=1, bd=0)
    login_btn.pack(side=tk.LEFT, padx=10)
    
    change_pass_btn = tk.Button(button_frame, text="تغيير كلمة السر", command=change_password,
                              font=('Arial', 14, 'bold'), bg='#3498db', fg='white',
                              width=15, height=1, bd=0)
    change_pass_btn.pack(side=tk.LEFT, padx=10)
    
    cancel_btn = tk.Button(button_frame, text="إلغاء", command=root.destroy,
                          font=('Arial', 14, 'bold'), bg='#e74c3c', fg='white',
                          width=15, height=1, bd=0)
    cancel_btn.pack(side=tk.LEFT, padx=10)
    
    password_entry.bind('<Return>', lambda event: perform_login())
    root.eval('tk::PlaceWindow . center')
    root.mainloop()
    
    return login_result

user = login_gui(["head", "it"])
if not user:
    sys.exit()

class ManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("برنامج إدارة الزائرين - جهاز المدير")
        self.root.geometry("1200x850")
        self.root.configure(bg='#f5f5f5')
        
        self.setup_header()
        
        # تحسين أداء الاتصال بقاعدة البيانات
        self.db_conn = sqlite3.connect('visitors_management.db', check_same_thread=False)
        self.db_conn.execute("PRAGMA journal_mode = WAL")  # تحسين الأداء للقراءة والكتابة المتزامنة
        self.db_conn.execute("PRAGMA synchronous = NORMAL")  # تحسين الأداء مع الحفاظ على سلامة البيانات
        self.db_conn.execute("PRAGMA cache_size = -10000")  # زيادة حجم الكاش
        self.db_cursor = self.db_conn.cursor()
        self.create_tables()
        
        pygame.mixer.init()
        self.notification_sound = None
        self._last_notification = 0
        self.notification_played = False
        try:
            sound_path = os.path.join(os.path.dirname(sys.executable), 'notification.wav') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'notification.wav')
            if os.path.exists(sound_path):
                self.notification_sound = pygame.mixer.Sound(sound_path)
            else:
                messagebox.showwarning("تحذير", "ملف الصوت notification.wav غير موجود")
        except Exception as e:
            messagebox.showwarning("تحذير", f"تعذر تحميل ملف الصوت: {str(e)}")
        
        self.auto_refresh_enabled = True
        self.selected_visitor_id = None
        self.current_filters = {'name': None, 'date': None, 'status': None}
        self.show_hidden_flag = False
        
        self.flash_frame = tk.Frame(self.root, bg='#FFD700', height=40, bd=1, relief='solid')
        self.flash_label = tk.Label(self.flash_frame, text="", bg='#FFD700', font=('Arial', 12, 'bold'), fg='#333')
        self.flash_label.pack(pady=5)
        self.flash_frame.pack_propagate(False)
        
        self.animation_colors = ['#FFD700', '#FFA500', '#FF8C00']
        self.current_animation_index = 0
        
        self.start_auto_backup()
        
        self.setup_ui()
        self.load_data()
        self.check_new_visitors()
        self.root.after(3000, self.auto_refresh)
        self.update_overdue_tasks()
    
    def update_overdue_tasks(self):
        """تحديث المهام المتأخرة تلقائياً"""
        today = date.today().strftime("%Y-%m-%d")
        self.db_cursor.execute('''UPDATE visit_details 
                               SET execution_status='متأخر' 
                               WHERE required_date < ? AND execution_status='جاري التنفيذ' ''', (today,))
        self.db_conn.commit()
        self.root.after(86400000, self.update_overdue_tasks)  # تحديث يومياً
    
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

    def setup_header(self):
        header_frame = tk.Frame(self.root, bg='#2c3e50')
        header_frame.pack(fill=tk.X, pady=5)
        
        try:
            logo_path = os.path.join(os.path.dirname(sys.executable), 'logo.png') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), 'logo.png')
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((80, 80), Image.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                
                logo_label = tk.Label(header_frame, image=self.logo_photo, bg='#2c3e50')
                logo_label.image = self.logo_photo
                logo_label.pack(side=tk.RIGHT, padx=20)
            else:
                raise FileNotFoundError("ملف اللوجو غير موجود")
        except Exception as e:
            print(f"Error loading logo: {e}")
            logo_label = tk.Label(header_frame, text="شعار المؤسسة", 
                                font=('Arial', 14), bg='#2c3e50', fg='white')
            logo_label.pack(side=tk.RIGHT, padx=20)
        
        title_frame = tk.Frame(header_frame, bg='#2c3e50')
        title_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)
        
        tk.Label(title_frame, text="وزارة الإسكان والمرافق والمجتمعات العمرانية", 
                font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack()
        tk.Label(title_frame, text="الجهاز المركزي للتعمير", 
                font=('Arial', 14), bg='#2c3e50', fg='white').pack()
        tk.Label(title_frame, text="برنامج إدارة الزائرين - السيد رئيس الجهاز المركزي للتعمير", 
                font=('Arial', 12), bg='#2c3e50', fg='white').pack(pady=(0,20))
    
    def create_tables(self):
        # التحقق من وجود الأعمدة قبل إنشائها
        self.db_cursor.execute("PRAGMA table_info(visitors)")
        columns = [column[1] for column in self.db_cursor.fetchall()]
        
        if 'original_name' not in columns:
            try:
                self.db_cursor.execute('''ALTER TABLE visitors ADD COLUMN original_name TEXT''')
                self.db_conn.commit()
            except Exception as e:
                print(f"Error adding column: {e}")
        
        if 'daily_id' not in columns:
            try:
                self.db_cursor.execute('''ALTER TABLE visitors ADD COLUMN daily_id INTEGER DEFAULT 1''')
                self.db_conn.commit()
            except Exception as e:
                print(f"Error adding column: {e}")
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS daily_counter (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              last_date TEXT,
                              counter INTEGER DEFAULT 1)''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS visitors (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              daily_id INTEGER DEFAULT 1,
                              name TEXT NOT NULL,
                              phone TEXT,
                              job TEXT,
                              position TEXT,
                              nationality TEXT,
                              visit_date TEXT,
                              status TEXT DEFAULT 'انتظار',
                              entry_time TEXT,
                              exit_time TEXT,
                              new_flag INTEGER DEFAULT 1,
                              hidden INTEGER DEFAULT 0,
                              appointment_time TEXT,
                              original_name TEXT,
                              last_update TEXT DEFAULT CURRENT_TIMESTAMP  -- عمود جديد لتتبع آخر تحديث
                              )''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS visit_details (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              visitor_id INTEGER NOT NULL,
                              visit_results TEXT,
                              recommendations TEXT,
                              required_date TEXT,
                              required_time TEXT,
                              execution_status TEXT DEFAULT 'جاري التنفيذ',
                              FOREIGN KEY(visitor_id) REFERENCES visitors(id)
                              )''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS notifications (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              visitor_id INTEGER,
                              notification_type TEXT,
                              created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                              FOREIGN KEY(visitor_id) REFERENCES visitors(id)
                              )''')
        
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS archived_visitors (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              visitor_data TEXT,
                              archived_at TEXT DEFAULT CURRENT_TIMESTAMP
                              )''')
        
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
        
        # إنشاء فهرس لتحسين أداء الاستعلامات
        self.db_cursor.execute('''CREATE INDEX IF NOT EXISTS idx_visitors_status ON visitors(status)''')
        self.db_cursor.execute('''CREATE INDEX IF NOT EXISTS idx_visitors_date ON visitors(visit_date)''')
        self.db_cursor.execute('''CREATE INDEX IF NOT EXISTS idx_visitors_new ON visitors(new_flag)''')
        self.db_cursor.execute('''CREATE INDEX IF NOT EXISTS idx_visitors_last_update ON visitors(last_update)''')
        
        self.db_conn.commit()
    
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
    
    def setup_ui(self):
        search_frame = tk.Frame(self.root, bg='#f5f5f5')
        search_frame.pack(fill=tk.X, padx=15, pady=10)
        
        btn_style = {'font': ('Arial', 12, 'bold'), 'height': 1, 'padx': 15, 'pady': 5, 'bd': 0}
        
        self.search_name_btn = tk.Button(search_frame, text="بحث بالاسم", command=self.search_by_name,
                                      bg='#2ecc71', fg='white', **btn_style)
        self.search_name_btn.pack(side=tk.LEFT, padx=5)
        
        self.search_date_btn = tk.Button(search_frame, text="بحث بالتاريخ", command=self.search_by_date,
                                      bg='#e67e22', fg='white', **btn_style)
        self.search_date_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_filter_btn = tk.Button(search_frame, text="إلغاء الفرز", command=self.clear_filters,
                                       bg='#95a5a6', fg='white', **btn_style)
        self.clear_filter_btn.pack(side=tk.LEFT, padx=5)
        
        self.stats_btn = tk.Button(search_frame, text="الإحصائيات", command=self.show_comprehensive_stats,
                                 bg='#9b59b6', fg='white', **btn_style)
        self.stats_btn.pack(side=tk.LEFT, padx=5)
        
        self.db_btn = tk.Button(search_frame, text="قاعدة البيانات", command=self.show_database_options,
                              bg='#34495e', fg='white', **btn_style)
        self.db_btn.pack(side=tk.LEFT, padx=5)
        
        self.tasks_btn = tk.Button(search_frame, text="المهام", command=self.show_tasks,
                                 bg='#3498db', fg='white', **btn_style)
        self.tasks_btn.pack(side=tk.LEFT, padx=5)
        
        self.commitments_btn = tk.Button(search_frame, text="الالتزامات", command=self.show_commitments,
                                       bg='#16a085', fg='white', **btn_style)
        self.commitments_btn.pack(side=tk.LEFT, padx=5)
        
        table_frame = tk.Frame(self.root, bg='#f5f5f5')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", font=('Arial', 12), rowheight=30, background='#fff', fieldbackground='#fff')
        style.configure("Treeview.Heading", font=('Arial', 12, 'bold'), background='#34495e', foreground='white')
        style.map("Treeview", background=[('selected', '#3498db')])
        
        columns = ("id", "daily_id", "name", "phone", "job", "position", "nationality", "visit_date", "status", "entry_time", "exit_time")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
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
        self.tree.column("position", width=120, anchor='center')
        self.tree.column("nationality", width=100, anchor='center')
        self.tree.column("visit_date", width=120, anchor='center')
        self.tree.column("status", width=100, anchor='center')
        self.tree.column("entry_time", width=120, anchor='center')
        self.tree.column("exit_time", width=120, anchor='center')
        
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        self.tree.bind('<Double-1>', self.on_name_click)
        self.tree.bind('<ButtonRelease-1>', self.on_tree_select)
        
        scrollbar = ttt.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.tag_configure('new', background='#fff3cd')
        self.tree.tag_configure('approved', background='#d4edda')
        self.tree.tag_configure('rejected', background='#f8d7da')
        self.tree.tag_configure('finished', background='#d1ecf1')
        self.tree.tag_configure('highlight', background='#fffacd')
        self.tree.tag_configure('hidden', background='#e9ecef')
        self.tree.tag_configure('appointment', background='#e6f3ff')
        
        action_frame = tk.Frame(self.root, bg='#f5f5f5')
        action_frame.pack(fill=tk.X, padx=15, pady=10)
        
        btn_style = {'font': ('Arial', 14, 'bold'), 'width': 15, 'height': 1, 'bd': 0}
        
        self.approve_btn = tk.Button(action_frame, text="موافق", command=self.approve_visitor,
                                  bg='#28a745', fg='white', **btn_style)
        self.approve_btn.pack(side=tk.LEFT, padx=5)
        
        self.reject_btn = tk.Button(action_frame, text="رفض", command=self.reject_visitor,
                                  bg='#dc3545', fg='white', **btn_style)
        self.reject_btn.pack(side=tk.LEFT, padx=5)
        
        self.finish_btn = tk.Button(action_frame, text="انتهاء", command=self.finish_visit,
                                  bg='#17a2b8', fg='white', **btn_style)
        self.finish_btn.pack(side=tk.LEFT, padx=5)
        
        footer_frame = tk.Frame(self.root, bg='#f5f5f5')
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))
        
        footer = tk.Label(footer_frame, text="By Dr. Mohammad Tahoun", 
                        font=('Arial', 12, 'bold'), fg='#0066cc', bg='#f5f5f5')
        footer.pack(pady=5)
    
    def on_tree_select(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.selected_visitor_id = self.tree.item(item)['values'][0]
            # إخفاء التنبيه عند اختيار الزائر
            self.hide_flash_message()
    
    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        query = '''SELECT id, daily_id, name, phone, job, position, nationality, 
                  visit_date, status, entry_time, exit_time, new_flag, hidden
                  FROM visitors'''
        
        params = []
        
        if not self.show_hidden_flag:
            query += " WHERE hidden = 0"
        
        if self.current_filters['name']:
            if "WHERE" not in query:
                query += " WHERE name LIKE ?"
            else:
                query += " AND name LIKE ?"
            params.append(f"%{self.current_filters['name']}%")
        
        if self.current_filters['date']:
            if "WHERE" not in query:
                query += " WHERE visit_date = ?"
            else:
                query += " AND visit_date = ?"
            params.append(self.current_filters['date'])
        
        if self.current_filters['status']:
            if "WHERE" not in query:
                query += " WHERE status = ?"
            else:
                query += " AND status = ?"
            params.append(self.current_filters['status'])
        
        # تعديل ترتيب العرض ليكون حسب آخر تحديث ثم التاريخ
        query += " ORDER BY last_update DESC, visit_date DESC, id DESC"
        
        self.db_cursor.execute(query, params)
        visitors = self.db_cursor.fetchall()
        
        for visitor in visitors:
            visitor_id, daily_id, name, phone, job, position, nationality, visit_date, status, entry_time, exit_time, new_flag, hidden = visitor
            
            tags = []
            if new_flag:
                tags.append('new')
            if hidden:
                tags.append('hidden')
            
            if status == 'موافق':
                tags.append('approved')
            elif status == 'رفض':
                tags.append('rejected')
            elif status == 'انتهاء':
                tags.append('finished')
            
            self.tree.insert("", tk.END, values=(
                visitor_id,
                daily_id,
                name,
                phone,
                job,
                position,
                nationality,
                visit_date,
                status,
                entry_time,
                exit_time
            ), tags=tuple(tags))
    
    def on_name_click(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if item and column == "#3":  # Changed from "#2" to "#3" because we added daily_id column
            self.selected_visitor_id = self.tree.item(item)['values'][0]
            self.show_visitor_details()
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.tree.see(item)
            self.tree.item(item, tags=('highlight',))
            self.root.after(3000, lambda: self.reset_row_color(item))
    
    def reset_row_color(self, item):
        tags = list(self.tree.item(item)['tags'])
        if 'highlight' in tags:
            tags.remove('highlight')
        self.tree.item(item, tags=tuple(tags))
    
    def show_visitor_details(self):
        if not self.selected_visitor_id:
            messagebox.showerror("خطأ", "لم يتم تحديد زائر")
            return
        
        details_window = tk.Toplevel(self.root)
        details_window.title("تفاصيل الزائر والمهام")
        details_window.geometry("1000x750")
        details_window.configure(bg='#f5f5f5')
        
        notebook = ttk.Notebook(details_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # تبويب بيانات الزيارة
        visit_frame = ttk.Frame(notebook, padding=10)
        notebook.add(visit_frame, text="بيانات الزيارة")
        
        self.db_cursor.execute('''SELECT name, phone, job, position, nationality, 
                               visit_date, status, entry_time, exit_time, hidden
                               FROM visitors WHERE id=?''', (self.selected_visitor_id,))
        visitor_data = self.db_cursor.fetchone()
        
        if visitor_data:
            name, phone, job, position, nationality, visit_date, status, entry_time, exit_time, hidden = visitor_data
            
            visit_info_frame = tk.Frame(visit_frame, bg='#f5f5f5', padx=10, pady=10)
            visit_info_frame.pack(fill=tk.BOTH, expand=True)
            
            labels = ["الاسم", "الهاتف", "الوظيفة", "الدرجة الوظيفية", "الجنسية", 
                     "تاريخ الزيارة", "الحالة", "وقت الدخول", "وقت الخروج"]
            
            values = [
                f"{name} {'(مخفي)' if hidden else ''}",
                phone,
                job,
                position,
                nationality,
                visit_date,
                status,
                entry_time,
                exit_time
            ]
            
            for i, (label, value) in enumerate(zip(labels, values)):
                label_frame = tk.Frame(visit_info_frame, bg='#f5f5f5')
                label_frame.pack(fill=tk.X, pady=5)
                
                tk.Label(label_frame, text=label+":", font=('Arial', 12, 'bold'), 
                        bg='#f5f5f5', width=15, anchor='e').pack(side=tk.LEFT, padx=5)
                tk.Label(label_frame, text=value, font=('Arial', 12), 
                        bg='#f5f5f5', width=30, anchor='w').pack(side=tk.LEFT, padx=5)
        
        # تبويب التوصيات والمهام
        tasks_frame = ttk.Frame(notebook, padding=10)
        notebook.add(tasks_frame, text="التوصيات والمهام")
        
        self.db_cursor.execute('''SELECT id, visit_results, recommendations, 
                               required_date, required_time, execution_status 
                               FROM visit_details WHERE visitor_id=?''', (self.selected_visitor_id,))
        task_data = self.db_cursor.fetchone()
        
        tasks_content_frame = tk.Frame(tasks_frame, bg='#f5f5f5')
        tasks_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        if task_data:
            task_id, results, recommendations, required_date, required_time, execution_status = task_data
            
            # نتائج الزيارة
            results_frame = tk.LabelFrame(tasks_content_frame, text="نتائج الزيارة", 
                                        font=('Arial', 12, 'bold'), bg='#f5f5f5', padx=10, pady=10)
            results_frame.pack(fill=tk.X, pady=5)
            
            results_text = tk.Text(results_frame, height=5, font=('Arial', 12), wrap=tk.WORD)
            results_text.insert(tk.END, results)
            results_text.pack(fill=tk.BOTH, expand=True)
            
            # التوصيات
            recommendations_frame = tk.LabelFrame(tasks_content_frame, text="التوصيات", 
                                               font=('Arial', 12, 'bold'), bg='#f5f5f5', padx=10, pady=10)
            recommendations_frame.pack(fill=tk.X, pady=5)
            
            recommendations_text = tk.Text(recommendations_frame, height=5, font=('Arial', 12), wrap=tk.WORD)
            recommendations_text.insert(tk.END, recommendations)
            recommendations_text.pack(fill=tk.BOTH, expand=True)
            
            # تاريخ ووقت التنفيذ
            datetime_frame = tk.Frame(tasks_content_frame, bg='#f5f5f5')
            datetime_frame.pack(fill=tk.X, pady=10)
            
            # تاريخ التنفيذ
            date_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
            date_frame.pack(side=tk.LEFT, padx=10)
            
            tk.Label(date_frame, text="تاريخ التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.required_date_entry = ttk.Entry(date_frame, font=('Arial', 12), width=15)
            self.required_date_entry.insert(0, required_date)
            self.required_date_entry.pack(side=tk.LEFT, padx=5)
            
            def pick_date():
                date_window = tk.Toplevel(details_window)
                date_window.title("اختر تاريخ")
                
                cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
                cal.pack(pady=10)
                
                def set_date():
                    selected_date = cal.get_date()
                    self.required_date_entry.delete(0, tk.END)
                    self.required_date_entry.insert(0, selected_date)
                    date_window.destroy()
                
                tk.Button(date_window, text="تأكيد", command=set_date).pack(pady=5)
            
            date_btn = ttk.Button(date_frame, text="...", command=pick_date, width=3)
            date_btn.pack(side=tk.LEFT, padx=5)
            
            # وقت التنفيذ
            time_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
            time_frame.pack(side=tk.LEFT, padx=10)
            
            tk.Label(time_frame, text="وقت التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.time_var = tk.StringVar(value=required_time if required_time else "09:00")
            self.am_pm_var = tk.StringVar(value="ص" if required_time and required_time.endswith("AM") else "م")
            
            time_options = [f"{h:02d}:00" for h in range(1, 13)]
            self.time_menu = ttk.Combobox(time_frame, textvariable=self.time_var, 
                                        values=time_options, font=('Arial', 12), width=8)
            self.time_menu.pack(side=tk.LEFT, padx=5)
            
            am_pm_menu = ttk.Combobox(time_frame, textvariable=self.am_pm_var, 
                                    values=["ص", "م"], font=('Arial', 12), width=3)
            am_pm_menu.pack(side=tk.LEFT, padx=5)
            
            # حالة التنفيذ
            status_frame = tk.Frame(tasks_content_frame, bg='#f5f5f5')
            status_frame.pack(fill=tk.X, pady=10)
            
            tk.Label(status_frame, text="حالة التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.status_var = tk.StringVar(value=execution_status)
            status_options = ["جاري التنفيذ", "تم", "متأخر"]
            status_menu = ttk.Combobox(status_frame, textvariable=self.status_var, 
                                     values=status_options, font=('Arial', 12))
            status_menu.pack(side=tk.LEFT, padx=10)
            
            # التحقق من حالة التأخير
            if execution_status == "جاري التنفيذ" and required_date:
                required_date_obj = datetime.strptime(required_date, "%Y-%m-%d").date()
                if date.today() > required_date_obj:
                    self.status_var.set("متأخر")
                    tk.Label(status_frame, text="حالة المهمة: متأخرة", 
                            font=('Arial', 12, 'bold'), fg='red', bg='#f5f5f5').pack(side=tk.LEFT, padx=20)
        else:
            task_id = None
            tk.Label(tasks_content_frame, text="لا توجد بيانات مهام لهذا الزائر", 
                    font=('Arial', 12, 'bold'), bg='#f5f5f5').pack(pady=20)
            
            # نتائج الزيارة
            results_frame = tk.LabelFrame(tasks_content_frame, text="نتائج الزيارة", 
                                        font=('Arial', 12, 'bold'), bg='#f5f5f5', padx=10, pady=10)
            results_frame.pack(fill=tk.X, pady=5)
            
            results_text = tk.Text(results_frame, height=5, font=('Arial', 12), wrap=tk.WORD)
            results_text.pack(fill=tk.BOTH, expand=True)
            
            # التوصيات
            recommendations_frame = tk.LabelFrame(tasks_content_frame, text="التوصيات", 
                                               font=('Arial', 12, 'bold'), bg='#f5f5f5', padx=10, pady=10)
            recommendations_frame.pack(fill=tk.X, pady=5)
            
            recommendations_text = tk.Text(recommendations_frame, height=5, font=('Arial', 12), wrap=tk.WORD)
            recommendations_text.pack(fill=tk.BOTH, expand=True)
            
            # تاريخ ووقت التنفيذ
            datetime_frame = tk.Frame(tasks_content_frame, bg='#f5f5f5')
            datetime_frame.pack(fill=tk.X, pady=10)
            
            # تاريخ التنفيذ
            date_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
            date_frame.pack(side=tk.LEFT, padx=10)
            
            tk.Label(date_frame, text="تاريخ التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.required_date_entry = ttk.Entry(date_frame, font=('Arial', 12), width=15)
            self.required_date_entry.pack(side=tk.LEFT, padx=5)
            
            def pick_date():
                date_window = tk.Toplevel(details_window)
                date_window.title("اختر تاريخ")
                
                cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
                cal.pack(pady=10)
                
                def set_date():
                    selected_date = cal.get_date()
                    self.required_date_entry.delete(0, tk.END)
                    self.required_date_entry.insert(0, selected_date)
                    date_window.destroy()
                
                tk.Button(date_window, text="تأكيد", command=set_date).pack(pady=5)
            
            date_btn = ttk.Button(date_frame, text="...", command=pick_date, width=3)
            date_btn.pack(side=tk.LEFT, padx=5)
            
            # وقت التنفيذ
            time_frame = tk.Frame(datetime_frame, bg='#f5f5f5')
            time_frame.pack(side=tk.LEFT, padx=10)
            
            tk.Label(time_frame, text="وقت التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.time_var = tk.StringVar(value="09:00")
            self.am_pm_var = tk.StringVar(value="ص")
            
            time_options = [f"{h:02d}:00" for h in range(1, 13)]
            self.time_menu = ttk.Combobox(time_frame, textvariable=self.time_var, 
                                        values=time_options, font=('Arial', 12), width=8)
            self.time_menu.pack(side=tk.LEFT, padx=5)
            
            am_pm_menu = ttk.Combobox(time_frame, textvariable=self.am_pm_var, 
                                    values=["ص", "م"], font=('Arial', 12), width=3)
            am_pm_menu.pack(side=tk.LEFT, padx=5)
            
            # حالة التنفيذ
            status_frame = tk.Frame(tasks_content_frame, bg='#f5f5f5')
            status_frame.pack(fill=tk.X, pady=10)
            
            tk.Label(status_frame, text="حالة التنفيذ:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT)
            
            self.status_var = tk.StringVar(value="جاري التنفيذ")
            status_options = ["جاري التنفيذ", "تم", "متأخر"]
            status_menu = ttk.Combobox(status_frame, textvariable=self.status_var, 
                                     values=status_options, font=('Arial', 12))
            status_menu.pack(side=tk.LEFT, padx=10)
        
        # أزرار التحكم
        control_frame = tk.Frame(details_window, bg='#f5f5f5', pady=10)
        control_frame.pack(fill=tk.X)
        
        btn_style = {'font': ('Arial', 12), 'width': 10, 'padx': 10, 'pady': 5}
        
        save_btn = tk.Button(control_frame, text="حفظ", command=lambda: self.save_task(
            task_id, results_text, recommendations_text, details_window),
            bg='#28a745', fg='white', **btn_style)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        if task_id:
            edit_btn = tk.Button(control_frame, text="تعديل", command=lambda: self.edit_task(
                task_id, results_text, recommendations_text, details_window),
                bg='#17a2b8', fg='white', **btn_style)
            edit_btn.pack(side=tk.LEFT, padx=5)
            
            delete_btn = tk.Button(control_frame, text="حذف", command=lambda: self.delete_task(
                task_id, details_window),
                bg='#dc3545', fg='white', **btn_style)
            delete_btn.pack(side=tk.LEFT, padx=5)
        
        complete_btn = tk.Button(control_frame, text="اكتمال", command=lambda: self.complete_task(
            task_id, results_text, recommendations_text, details_window),
            bg='#2ecc71', fg='white', **btn_style)
        complete_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(control_frame, text="إلغاء", command=details_window.destroy,
                             bg='#e74c3c', fg='white', **btn_style)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def save_task(self, task_id, results_text, recommendations_text, window):
        results = results_text.get("1.0", tk.END).strip()
        recommendations = recommendations_text.get("1.0", tk.END).strip()
        required_date = self.required_date_entry.get()
        required_time = f"{self.time_var.get()} {self.am_pm_var.get()}"
        status = self.status_var.get()
        
        if status == "جاري التنفيذ" and required_date:
            required_date_obj = datetime.strptime(required_date, "%Y-%m-%d").date()
            if date.today() > required_date_obj:
                status = "متأخر"
        
        try:
            if task_id:
                self.db_cursor.execute('''UPDATE visit_details 
                                      SET visit_results=?, recommendations=?, 
                                      required_date=?, required_time=?, execution_status=?
                                      WHERE id=?''',
                                      (results, recommendations, required_date, required_time, status, task_id))
            else:
                self.db_cursor.execute('''INSERT INTO visit_details 
                                      (visitor_id, visit_results, recommendations, 
                                       required_date, required_time, execution_status)
                                      VALUES (?, ?, ?, ?, ?, ?)''',
                                      (self.selected_visitor_id, results, recommendations, 
                                       required_date, required_time, status))
            
            # تحديث وقت التعديل للزائر
            self.db_cursor.execute('''UPDATE visitors SET last_update=CURRENT_TIMESTAMP 
                                   WHERE id=?''', (self.selected_visitor_id,))
            
            self.db_conn.commit()
            messagebox.showinfo("نجاح", "تم حفظ بيانات المهمة بنجاح")
            window.destroy()
            self.load_data()  # تحديث الجدول الرئيسي
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ المهمة: {str(e)}")
    
    def edit_task(self, task_id, results_text, recommendations_text, window):
        self.save_task(task_id, results_text, recommendations_text, window)
    
    def delete_task(self, task_id, window):
        confirm = messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذه المهمة؟")
        if confirm:
            try:
                self.db_cursor.execute("DELETE FROM visit_details WHERE id=?", (task_id,))
                # تحديث وقت التعديل للزائر
                self.db_cursor.execute('''UPDATE visitors SET last_update=CURRENT_TIMESTAMP 
                                       WHERE id=?''', (self.selected_visitor_id,))
                self.db_conn.commit()
                messagebox.showinfo("نجاح", "تم حذف المهمة بنجاح")
                window.destroy()
                self.load_data()  # تحديث الجدول الرئيسي
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف المهمة: {str(e)}")
    
    def complete_task(self, task_id, results_text, recommendations_text, window):
        self.status_var.set("تم")
        self.save_task(task_id, results_text, recommendations_text, window)
    
    def search_by_name(self):
        name = simpledialog.askstring("بحث بالاسم", "أدخل اسم الزائر:")
        if name is not None:
            self.current_filters['name'] = name.strip() if name.strip() else None
            self.load_data()
    
    def search_by_date(self):
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
    
    def clear_filters(self):
        self.current_filters = {'name': None, 'date': None, 'status': None}
        self.show_hidden_flag = False
        self.load_data()
    
    def approve_visitor(self):
        if not self.selected_visitor_id:
            messagebox.showerror("خطأ", "يجب اختيار زائر أولا")
            return
        
        try:
            current_time = datetime.now().strftime("%H:%M")
            self.db_cursor.execute('''UPDATE visitors SET status='موافق', entry_time=?, new_flag=0, last_update=CURRENT_TIMESTAMP
                                   WHERE id=?''', (current_time, self.selected_visitor_id))
            self.db_conn.commit()
            self.notification_played = False
            self.load_data()
            self.play_notification_sound()
            self.highlight_row(self.selected_visitor_id)
            messagebox.showinfo("نجاح", "تمت الموافقة على الزيارة بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء الموافقة: {str(e)}")
    
    def reject_visitor(self):
        if not self.selected_visitor_id:
            messagebox.showerror("خطأ", "يجب اختيار زائر أولا")
            return
        
        try:
            self.db_cursor.execute('''SELECT name FROM visitors WHERE id=?''', (self.selected_visitor_id,))
            visitor_data = self.db_cursor.fetchone()
            
            if not visitor_data:
                messagebox.showerror("خطأ", "الزائر غير موجود")
                return
                
            original_name = visitor_data[0]
            
            self.db_cursor.execute('''UPDATE visitors 
                                  SET status='رفض', new_flag=0, hidden=1, original_name=?, last_update=CURRENT_TIMESTAMP
                                  WHERE id=?''', 
                                  (original_name, self.selected_visitor_id))
            self.db_conn.commit()
            self.notification_played = False
            self.load_data()
            self.play_notification_sound()
            messagebox.showinfo("نجاح", "تم رفض الزيارة وإخفاء الزائر بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء الرفض: {str(e)}")
    
    def finish_visit(self):
        if not self.selected_visitor_id:
            messagebox.showerror("خطأ", "يجب اختيار زائر أولا")
            return
        
        try:
            current_time = datetime.now().strftime("%H:%M")
            
            self.db_cursor.execute('''SELECT name FROM visitors WHERE id=?''', (self.selected_visitor_id,))
            visitor_data = self.db_cursor.fetchone()
            
            if not visitor_data:
                messagebox.showerror("خطأ", "الزائر غير موجود")
                return
                
            original_name = visitor_data[0]
            
            self.db_cursor.execute('''UPDATE visitors 
                                  SET status='انتهاء', exit_time=?, new_flag=0, hidden=1, original_name=?, last_update=CURRENT_TIMESTAMP
                                  WHERE id=?''', 
                                  (current_time, original_name, self.selected_visitor_id))
            self.db_conn.commit()
            self.notification_played = False
            self.load_data()
            self.play_notification_sound()
            messagebox.showinfo("نجاح", "تم إنهاء الزيارة وإخفاء الزائر بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء إنهاء الزيارة: {str(e)}")
    
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
        
        self.db_cursor.execute("SELECT COUNT(*) FROM visitors")
        stats["إجمالي عدد الزوار"] = self.db_cursor.fetchone()[0]
        
        today = date.today().strftime("%Y-%m-%d")
        self.db_cursor.execute("SELECT COUNT(*) FROM visitors WHERE visit_date=?", (today,))
        stats["عدد الزوار اليوم"] = self.db_cursor.fetchone()[0]
        
        statuses = ['انتظار', 'موافق', 'رفض', 'انتهاء']
        for status in statuses:
            self.db_cursor.execute("SELECT COUNT(*) FROM visitors WHERE status=?", (status,))
            stats[f"عدد الزوار ({status})"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments")
        stats["عدد المواعيد المسبقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments WHERE status='معلقة'")
        stats["مواعيد معلقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM appointments WHERE status='مكتملة'")
        stats["مواعيد مكتملة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments")
        stats["عدد الالتزامات"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments WHERE status='معلقة'")
        stats["التزامات معلقة"] = self.db_cursor.fetchone()[0]
        
        self.db_cursor.execute("SELECT COUNT(*) FROM commitments WHERE status='مكتملة'")
        stats["التزامات مكتملة"] = self.db_cursor.fetchone()[0]
        
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
        confirm = messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف جميع سجلات الزوار؟")
        if confirm:
            try:
                self.db_cursor.execute("DELETE FROM visitors")
                self.db_conn.commit()
                self.load_data()
                messagebox.showinfo("نجاح", "تم حذف جميع سجلات الزوار بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")
    
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
            confirm = messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذه المهمة؟")
            if confirm:
                try:
                    self.db_cursor.execute("DELETE FROM visit_details WHERE id=?", (task_id,))
                    self.db_conn.commit()
                    messagebox.showinfo("نجاح", "تم حذف المهمة بنجاح")
                    details_window.destroy()
                    self.load_tasks_data(details_window.master)
                except Exception as e:
                    messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")
        
        btn_frame = tk.Frame(details_window)
        btn_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(btn_frame, text="حفظ التعديلات", command=update_task,
                font=('Arial', 12), bg='#28a745', fg='white').pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="حذف المهمة", command=delete_task,
                font=('Arial', 12), bg='#dc3545', fg='white').pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="إغلاق", command=details_window.destroy,
                font=('Arial', 12), bg='#6c757d', fg='white').pack(side=tk.RIGHT, padx=10)
    
    def show_commitments(self):
        commitments_window = tk.Toplevel(self.root)
        commitments_window.title("الالتزامات")
        commitments_window.geometry("1000x600")
        commitments_window.configure(bg='#f5f5f5')
        
        filter_frame = tk.Frame(commitments_window, bg='#f5f5f5', padx=10, pady=10)
        filter_frame.pack(fill=tk.X)
        
        tk.Label(filter_frame, text="حالة الالتزام:", font=('Arial', 12), bg='#f5f5f5').pack(side=tk.LEFT, padx=5)
        
        self.commit_status_var = tk.StringVar()
        self.commit_status_var.set("الكل")
        
        status_options = ["الكل", "معلقة", "مكتملة"]
        status_menu = tk.OptionMenu(filter_frame, self.commit_status_var, *status_options)
        status_menu.config(font=('Arial', 12), bg='#f8f9fa', bd=1, relief='solid')
        status_menu.pack(side=tk.LEFT, padx=5)
        
        apply_filter_btn = tk.Button(filter_frame, text="تطبيق الفلتر", command=lambda: self.load_commitments_data(self.commit_status_var.get()),
                                   font=('Arial', 12), bg='#3498db', fg='white', padx=10)
        apply_filter_btn.pack(side=tk.LEFT, padx=5)
        
        clear_filter_btn = tk.Button(filter_frame, text="إلغاء الفرز", command=lambda: [self.commit_status_var.set("الكل"), self.load_commitments_data("الكل")],
                                   font=('Arial', 12), bg='#95a5a6', fg='white', padx=10)
        clear_filter_btn.pack(side=tk.LEFT, padx=5)
        
        table_frame = tk.Frame(commitments_window, bg='#f5f5f5')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Commitments.Treeview", font=('Arial', 12), rowheight=30, background='#fff', fieldbackground='#fff')
        style.configure("Commitments.Treeview.Heading", font=('Arial', 12, 'bold'), background='#34495e', foreground='white')
        style.map("Commitments.Treeview", background=[('selected', '#3498db')])
        
        columns = ("id", "place", "purpose", "department", "commitment_date", "commitment_time", "status")
        self.commitments_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, style="Commitments.Treeview")
        
        self.commitments_tree.heading("id", text="ID", anchor='center')
        self.commitments_tree.heading("place", text="المكان", anchor='center')
        self.commitments_tree.heading("purpose", text="الغرض", anchor='center')
        self.commitments_tree.heading("department", text="الجهة", anchor='center')
        self.commitments_tree.heading("commitment_date", text="التاريخ", anchor='center')
        self.commitments_tree.heading("commitment_time", text="الوقت", anchor='center')
        self.commitments_tree.heading("status", text="الحالة", anchor='center')
        
        self.commitments_tree.column("id", width=50, anchor='center')
        self.commitments_tree.column("place", width=150, anchor='center')
        self.commitments_tree.column("purpose", width=200, anchor='center')
        self.commitments_tree.column("department", width=150, anchor='center')
        self.commitments_tree.column("commitment_date", width=120, anchor='center')
        self.commitments_tree.column("commitment_time", width=100, anchor='center')
        self.commitments_tree.column("status", width=100, anchor='center')
        
        self.commitments_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.commitments_tree.yview)
        self.commitments_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.commitments_tree.tag_configure('pending', background='#fff3cd')
        self.commitments_tree.tag_configure('completed', background='#d4edda')
        
        control_frame = tk.Frame(commitments_window, bg='#f5f5f5')
        control_frame.pack(fill=tk.X, padx=15, pady=10)
        
        complete_btn = tk.Button(control_frame, text="تم التنفيذ", command=self.mark_commitment_complete,
                               font=('Arial', 12), bg='#2ecc71', fg='white', padx=15)
        complete_btn.pack(side=tk.LEFT, padx=10)
        
        close_btn = tk.Button(control_frame, text="إغلاق", command=commitments_window.destroy,
                            font=('Arial', 12), bg='#e74c3c', fg='white', padx=15)
        close_btn.pack(side=tk.LEFT, padx=10)
        
        self.load_commitments_data("الكل")
    
    def load_commitments_data(self, status_filter="الكل"):
        for item in self.commitments_tree.get_children():
            self.commitments_tree.delete(item)
        
        query = '''SELECT id, place, purpose, department, 
                  commitment_date, commitment_time, status
                  FROM commitments'''
        
        params = []
        
        if status_filter != "الكل":
            query += " WHERE status = ?"
            params.append("مكتملة" if status_filter == "مكتملة" else "معلقة")
        
        query += " ORDER BY commitment_date, commitment_time"
        
        self.db_cursor.execute(query, params)
        commitments = self.db_cursor.fetchall()
        
        for commit in commitments:
            commit_id, place, purpose, department, commit_date, commit_time, status = commit
            
            tags = []
            if status == "معلقة":
                tags.append('pending')
            else:
                tags.append('completed')
            
            self.commitments_tree.insert("", tk.END, values=(
                commit_id,
                place,
                purpose,
                department,
                commit_date,
                commit_time,
                status
            ), tags=tuple(tags))
    
    def mark_commitment_complete(self):
        selected_item = self.commitments_tree.selection()
        if not selected_item:
            messagebox.showerror("خطأ", "يجب اختيار التزام لتحديده كمكتمل")
            return
            
        commit_id = self.commitments_tree.item(selected_item)['values'][0]
        
        try:
            self.db_cursor.execute('''UPDATE commitments SET status='مكتملة' WHERE id=?''', (commit_id,))
            self.db_conn.commit()
            self.load_commitments_data(self.commit_status_var.get())
            messagebox.showinfo("نجاح", "تم تحديث حالة الالتزام إلى مكتمل")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تحديث الالتزام: {str(e)}")
    
    def show_flash_message(self, message):
        self.flash_label.config(text=message)
        self.flash_frame.pack(fill=tk.X, pady=5)
        self.animate_flash_message()
    
    def animate_flash_message(self):
        if self.flash_frame.winfo_ismapped():
            self.flash_frame.config(bg=self.animation_colors[self.current_animation_index])
            self.current_animation_index = (self.current_animation_index + 1) % len(self.animation_colors)
            self.root.after(500, self.animate_flash_message)
    
    def hide_flash_message(self):
        self.flash_frame.pack_forget()
    
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
    
    def highlight_row(self, visitor_id):
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == visitor_id:
                self.tree.selection_set(item)
                self.tree.focus(item)
                self.tree.see(item)
                self.tree.item(item, tags=('highlight',))
                self.root.after(3000, lambda: self.reset_row_color(item))
                break
    
    def reset_row_color(self, item):
        tags = list(self.tree.item(item)['tags'])
        if 'highlight' in tags:
            tags.remove('highlight')
        self.tree.item(item, tags=tuple(tags))
    
    def auto_refresh(self):
        if self.auto_refresh_enabled:
            self.load_data()
        self.root.after(30000, self.auto_refresh)

if __name__ == "__main__":
    root = tk.Tk()
    app = ManagerApp(root)
    root.mainloop()