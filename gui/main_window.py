import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime
import os
import shutil
import pandas as pd

class MainWindow:
    def __init__(self, root, db_manager, file_manager, export_manager, printer_manager=None):
        self.root = root
        self.db_manager = db_manager
        self.file_manager = file_manager
        self.export_manager = export_manager
        self.printer_manager = printer_manager
        
        self.setup_window()
        self.create_menu()
        self.create_widgets()
        self.load_statistics()
    
    def setup_window(self):
        """إعداد النافذة الرئيسية"""
        self.root.title("نظام إدارة المراسلات - الإصدار 2.0")
        self.root.geometry("1200x700")
        self.root.state('zoomed')
        
        # تحسين مظهر الواجهة
        self.setup_styles()
        
        # إنشاء إطارات رئيسية
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def setup_styles(self):
        """إعداد أنماط الواجهة"""
        style = ttk.Style()
        
        # محاولة استخدام tema حديث
        try:
            style.theme_use('vista')
        except:
            try:
                style.theme_use('clam')
            except:
                pass
        
        # تخصيص الأنماط
        style.configure('Title.TLabel', font=('Arial', 12, 'bold'), foreground='#2C3E50')
        style.configure('Stats.TLabel', font=('Arial', 10, 'bold'))
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'), background='#3498DB')
    
    def create_menu(self):
        """إنشاء شريط القوائم"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # قائمة الملف
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="الملف", menu=file_menu)
        file_menu.add_command(label="تسجيل وارد جديد", command=self.open_incoming_form)
        file_menu.add_command(label="تسجيل صادر جديد", command=self.open_outgoing_form)
        file_menu.add_separator()
        file_menu.add_command(label="تصدير البيانات", command=self.export_all_data)
        file_menu.add_separator()
        file_menu.add_command(label="خروج", command=self.root.quit)
        
        # قائمة التقارير
        reports_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="التقارير", menu=reports_menu)
        reports_menu.add_command(label="تقرير الوارد", command=self.open_incoming_reports)
        reports_menu.add_command(label="تقرير الصادر", command=self.open_outgoing_reports)
        reports_menu.add_separator()
        reports_menu.add_command(label="تقرير الموظفين", command=self.open_employee_reports)
        reports_menu.add_command(label="تقرير شامل", command=self.open_comprehensive_report)
        
        # قائمة الإدارة
        management_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="الإدارة", menu=management_menu)
        management_menu.add_command(label="إدارة الكيانات المرجعية", command=self.open_reference_management)
        management_menu.add_command(label="إدارة الموظفين", command=self.open_employee_management)
        management_menu.add_separator()
        management_menu.add_command(label="نسخ احتياطي", command=self.backup_database)
        management_menu.add_command(label="استعادة نسخة احتياطية", command=self.restore_database)
        
        # قائمة المساعدة
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="المساعدة", menu=help_menu)
        help_menu.add_command(label="دليل المستخدم", command=self.show_user_guide)
        help_menu.add_command(label="حول النظام", command=self.show_about)
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # إنشاء Notebook (تبويبات)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # تبويب لوحة التحكم
        self.dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_frame, text="🏠 لوحة التحكم")
        
        # تبويب سجلات الوارد
        self.incoming_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.incoming_frame, text="📥 سجلات الوارد")
        
        # تبويب سجلات الصادر
        self.outgoing_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.outgoing_frame, text="📤 سجلات الصادر")
        
        # تبويب البحث
        self.search_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.search_frame, text="🔍 بحث متقدم")
        
        self.setup_dashboard()
        self.setup_incoming_tab()
        self.setup_outgoing_tab()
        self.setup_search_tab()
    
    def setup_dashboard(self):
        """إعداد لوحة التحكم"""
        # إطار الإحصائيات
        stats_frame = ttk.LabelFrame(self.dashboard_frame, text="📊 الإحصائيات العامة", padding=15)
        stats_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # شبكة للإحصائيات
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X, padx=10)
        
        # إحصائيات الوارد
        ttk.Label(stats_grid, text="إجمالي سجلات الوارد:", 
                 font=('Arial', 11, 'bold'), foreground='#2E86AB').grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.incoming_count_label = ttk.Label(stats_grid, text="0", 
                                            font=('Arial', 12, 'bold'), foreground='#2E86AB')
        self.incoming_count_label.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(stats_grid, text="سجلات هذا الشهر:", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=10, pady=2)
        self.incoming_month_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.incoming_month_label.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        # إحصائيات الصادر
        ttk.Label(stats_grid, text="إجمالي سجلات الصادر:", 
                 font=('Arial', 11, 'bold'), foreground='#A23B72').grid(row=0, column=2, sticky='w', padx=20, pady=5)
        self.outgoing_count_label = ttk.Label(stats_grid, text="0", 
                                            font=('Arial', 12, 'bold'), foreground='#A23B72')
        self.outgoing_count_label.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(stats_grid, text="سجلات هذا الشهر:", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=20, pady=2)
        self.outgoing_month_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.outgoing_month_label.grid(row=1, column=3, sticky='w', padx=5, pady=2)
        
        # إحصائيات إضافية
        ttk.Label(stats_grid, text="إجمالي الموظفين:", 
                 font=('Arial', 11, 'bold'), foreground='#F18F01').grid(row=0, column=4, sticky='w', padx=20, pady=5)
        self.employees_count_label = ttk.Label(stats_grid, text="0", 
                                             font=('Arial', 12, 'bold'), foreground='#F18F01')
        self.employees_count_label.grid(row=0, column=5, sticky='w', padx=5, pady=5)
        
        ttk.Label(stats_grid, text="إجمالي المرفقات:", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=4, sticky='w', padx=20, pady=2)
        self.attachments_count_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.attachments_count_label.grid(row=1, column=5, sticky='w', padx=5, pady=2)
        
        ttk.Label(stats_grid, text="آخر تحديث:", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=6, sticky='w', padx=20, pady=5)
        self.last_update_label = ttk.Label(stats_grid, text=datetime.now().strftime('%Y-%m-%d %H:%M'), 
                                         font=('Arial', 10))
        self.last_update_label.grid(row=0, column=7, sticky='w', padx=5, pady=5)
        
        # أزرار سريعة
        buttons_frame = ttk.Frame(self.dashboard_frame)
        buttons_frame.pack(fill=tk.X, pady=15, padx=5)
        
        # إنشاء إطار للأزرار مع تخطيط أفضل
        quick_actions_frame = ttk.LabelFrame(buttons_frame, text="⚡ إجراءات سريعة", padding=10)
        quick_actions_frame.pack(fill=tk.X)
        
        # أزرار بالإطار الخاص بها
        action_buttons = [
            ("📥 تسجيل وارد جديد", self.open_incoming_form, "#2E86AB"),
            ("📤 تسجيل صادر جديد", self.open_outgoing_form, "#A23B72"),
            ("🔍 بحث متقدم", self.open_search_window, "#F18F01"),
            ("👥 تقارير الموظفين", self.open_employee_reports, "#C73E1D"),
            ("🔄 تحديث البيانات", self.refresh_data, "#4CAF50"),
            ("⚙️ إدارة النظام", self.open_reference_management, "#6A4C93")
        ]
        
        for i, (text, command, color) in enumerate(action_buttons):
            btn = tk.Button(quick_actions_frame, 
                          text=text,
                          command=command,
                          bg=color,
                          fg='white',
                          font=('Arial', 10, 'bold'),
                          padx=15,
                          pady=8,
                          relief='raised',
                          bd=2,
                          cursor='hand2')
            btn.pack(side=tk.RIGHT, padx=5, pady=2)
        
        # أحدث السجلات
        self.setup_recent_records()
    
    def setup_recent_records(self):
        """إعداد عرض أحدث السجلات"""
        records_frame = ttk.Frame(self.dashboard_frame)
        records_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # أحدث الوارد
        incoming_frame = ttk.LabelFrame(records_frame, text="🆕 أحدث سجلات الوارد", padding=10)
        incoming_frame.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=5)
        
        columns = ('رقم السجل', 'الرقم التسلسلي', 'العنوان', 'التاريخ')
        self.recent_incoming_tree = ttk.Treeview(incoming_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.recent_incoming_tree.heading(col, text=col)
            self.recent_incoming_tree.column(col, width=150)
        
        # إضافة شريط تمرير
        incoming_scrollbar = ttk.Scrollbar(incoming_frame, orient=tk.VERTICAL, command=self.recent_incoming_tree.yview)
        self.recent_incoming_tree.configure(yscrollcommand=incoming_scrollbar.set)
        
        self.recent_incoming_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        incoming_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # أحدث الصادر
        outgoing_frame = ttk.LabelFrame(records_frame, text="🆕 أحدث سجلات الصادر", padding=10)
        outgoing_frame.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT, padx=5)
        
        self.recent_outgoing_tree = ttk.Treeview(outgoing_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.recent_outgoing_tree.heading(col, text=col)
            self.recent_outgoing_tree.column(col, width=150)
        
        # إضافة شريط تمرير
        outgoing_scrollbar = ttk.Scrollbar(outgoing_frame, orient=tk.VERTICAL, command=self.recent_outgoing_tree.yview)
        self.recent_outgoing_tree.configure(yscrollcommand=outgoing_scrollbar.set)
        
        self.recent_outgoing_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        outgoing_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_incoming_tab(self):
        """إعداد تبويب سجلات الوارد"""
        # إطار البحث والتصفية
        filter_frame = ttk.LabelFrame(self.incoming_frame, text="🔍 بحث وتصفية سجلات الوارد", padding=10)
        filter_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # صف البحث
        search_row = ttk.Frame(filter_frame)
        search_row.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_row, text="بحث:", font=('Arial', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        self.incoming_search_entry = ttk.Entry(search_row, width=30, font=('Arial', 10))
        self.incoming_search_entry.pack(side=tk.RIGHT, padx=5)
        self.incoming_search_entry.bind('<KeyRelease>', self.search_incoming)
        
        # صف الأزرار
        buttons_row = ttk.Frame(filter_frame)
        buttons_row.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_row, text="🔄 عرض الكل", 
                  command=self.load_incoming_records,
                  style='Accent.TButton').pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="✏️ تعديل", 
                  command=self.edit_incoming_record).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🗑️ حذف", 
                  command=self.delete_incoming_record).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="📊 تصدير", 
                  command=self.export_incoming).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🖨️ طباعة", 
                  command=self.print_incoming).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🖨️ طباعة محدد", 
                  command=self.print_selected_incoming).pack(side=tk.RIGHT, padx=3)
        
        # جدول سجلات الوارد
        table_frame = ttk.Frame(self.incoming_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        columns = ('ID', 'رقم السجل', 'رقم الوارد', 'الرقم التسلسلي', 'العنوان', 
                  'جهة الوارد', 'النوع', 'الموظف', 'التاريخ')
        
        self.incoming_tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        # إخفاء عمود ID
        self.incoming_tree.column('ID', width=0, stretch=tk.NO)
        
        # تعيين العناوين والأبعاد
        column_config = {
            'رقم السجل': 120,
            'رقم الوارد': 120,
            'الرقم التسلسلي': 100,
            'العنوان': 200,
            'جهة الوارد': 150,
            'النوع': 120,
            'الموظف': 120,
            'التاريخ': 100
        }
        
        for col, width in column_config.items():
            self.incoming_tree.heading(col, text=col)
            self.incoming_tree.column(col, width=width)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.incoming_tree.yview)
        self.incoming_tree.configure(yscrollcommand=scrollbar.set)
        
        self.incoming_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.incoming_tree.bind('<Double-1>', lambda e: self.edit_incoming_record())
        
        self.load_incoming_records()
    
    def setup_outgoing_tab(self):
        """إعداد تبويب سجلات الصادر"""
        # إطار البحث والتصفية
        filter_frame = ttk.LabelFrame(self.outgoing_frame, text="🔍 بحث وتصفية سجلات الصادر", padding=10)
        filter_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # صف البحث
        search_row = ttk.Frame(filter_frame)
        search_row.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_row, text="بحث:", font=('Arial', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        self.outgoing_search_entry = ttk.Entry(search_row, width=30, font=('Arial', 10))
        self.outgoing_search_entry.pack(side=tk.RIGHT, padx=5)
        self.outgoing_search_entry.bind('<KeyRelease>', self.search_outgoing)
        
        # صف الأزرار
        buttons_row = ttk.Frame(filter_frame)
        buttons_row.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_row, text="🔄 عرض الكل", 
                  command=self.load_outgoing_records,
                  style='Accent.TButton').pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="✏️ تعديل", 
                  command=self.edit_outgoing_record).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🗑️ حذف", 
                  command=self.delete_outgoing_record).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="📊 تصدير", 
                  command=self.export_outgoing).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🖨️ طباعة", 
                  command=self.print_outgoing).pack(side=tk.RIGHT, padx=3)
        ttk.Button(buttons_row, text="🖨️ طباعة محدد", 
                  command=self.print_selected_outgoing).pack(side=tk.RIGHT, padx=3)
        
        # جدول سجلات الصادر
        table_frame = ttk.Frame(self.outgoing_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        columns = ('ID', 'رقم السجل', 'رقم الصادر', 'الرقم التسلسلي', 'العنوان', 
                  'جهة الصادر', 'الموظف', 'التاريخ')
        
        self.outgoing_tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        # إخفاء عمود ID
        self.outgoing_tree.column('ID', width=0, stretch=tk.NO)
        
        # تعيين العناوين والأبعاد
        column_config = {
            'رقم السجل': 120,
            'رقم الصادر': 120,
            'الرقم التسلسلي': 100,
            'العنوان': 200,
            'جهة الصادر': 150,
            'الموظف': 120,
            'التاريخ': 100
        }
        
        for col, width in column_config.items():
            self.outgoing_tree.heading(col, text=col)
            self.outgoing_tree.column(col, width=width)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.outgoing_tree.yview)
        self.outgoing_tree.configure(yscrollcommand=scrollbar.set)
        
        self.outgoing_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.outgoing_tree.bind('<Double-1>', lambda e: self.edit_outgoing_record())
        
        self.load_outgoing_records()
    
    def setup_search_tab(self):
        """إعداد تبويب البحث"""
        search_frame = ttk.LabelFrame(self.search_frame, text="🔍 بحث متقدم في النظام", padding=10)
        search_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # إطار معايير البحث
        criteria_frame = ttk.Frame(search_frame)
        criteria_frame.pack(fill=tk.X, pady=10)
        
        # خيارات البحث
        ttk.Label(criteria_frame, text="نوع البحث:", font=('Arial', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        self.search_type = tk.StringVar(value="both")
        ttk.Radiobutton(criteria_frame, text="📥 وارد", variable=self.search_type, value="incoming").pack(side=tk.RIGHT, padx=5)
        ttk.Radiobutton(criteria_frame, text="📤 صادر", variable=self.search_type, value="outgoing").pack(side=tk.RIGHT, padx=5)
        ttk.Radiobutton(criteria_frame, text="📊 الكل", variable=self.search_type, value="both").pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(criteria_frame, text="نص البحث:", font=('Arial', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        self.search_entry = ttk.Entry(criteria_frame, width=40, font=('Arial', 10))
        self.search_entry.pack(side=tk.RIGHT, padx=5)
        self.search_entry.bind('<Return>', lambda e: self.perform_search())
        
        ttk.Button(criteria_frame, text="🔍 بحث", 
                  command=self.perform_search,
                  style='Accent.TButton').pack(side=tk.RIGHT, padx=5)
        ttk.Button(criteria_frame, text="🗑️ مسح", 
                  command=self.clear_search).pack(side=tk.RIGHT, padx=5)
        
        # نتائج البحث
        results_frame = ttk.Frame(search_frame)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ('النوع', 'رقم السجل', 'الرقم', 'الرقم التسلسلي', 'العنوان', 'التاريخ', 'ID', 'RecordType')
        self.search_results_tree = ttk.Treeview(results_frame, columns=columns, show='headings')
        
        # إخفاء الأعمدة الإضافية
        self.search_results_tree.column('ID', width=0, stretch=tk.NO)
        self.search_results_tree.column('RecordType', width=0, stretch=tk.NO)
        
        for col in ('النوع', 'رقم السجل', 'الرقم', 'الرقم التسلسلي', 'العنوان', 'التاريخ'):
            self.search_results_tree.heading(col, text=col)
            self.search_results_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.search_results_tree.yview)
        self.search_results_tree.configure(yscrollcommand=scrollbar.set)
        
        # أزرار التحكم في نتائج البحث
        results_buttons_frame = ttk.Frame(search_frame)
        results_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(results_buttons_frame, text="👁️ عرض التفاصيل", 
                  command=self.show_search_details).pack(side=tk.RIGHT, padx=5)
        ttk.Button(results_buttons_frame, text="📊 تصدير النتائج", 
                  command=self.export_search_results).pack(side=tk.RIGHT, padx=5)
        ttk.Button(results_buttons_frame, text="🖨️ طباعة", 
                  command=self.print_search_results).pack(side=tk.RIGHT, padx=5)
        
        self.search_results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.search_results_tree.bind('<Double-1>', lambda e: self.show_search_details())
    
    def load_statistics(self):
        """تحميل الإحصائيات"""
        try:
            # إحصائيات الوارد
            incoming_total = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM incoming_records"
            )[0][0]
            
            current_month = datetime.now().strftime('%Y-%m')
            incoming_month = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM incoming_records WHERE strftime('%Y-%m', registration_date) = ?",
                (current_month,)
            )[0][0]
            
            # إحصائيات الصادر
            outgoing_total = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM outgoing_records"
            )[0][0]
            
            outgoing_month = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM outgoing_records WHERE strftime('%Y-%m', registration_date) = ?",
                (current_month,)
            )[0][0]
            
            # إحصائيات الموظفين
            employees_total = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM employees WHERE is_active = 1"
            )[0][0]
            
            # إحصائيات المرفقات
            attachments_total = self.db_manager.execute_query(
                "SELECT COUNT(*) FROM attachments"
            )[0][0]
            
            self.incoming_count_label.config(text=str(incoming_total))
            self.incoming_month_label.config(text=str(incoming_month))
            self.outgoing_count_label.config(text=str(outgoing_total))
            self.outgoing_month_label.config(text=str(outgoing_month))
            self.employees_count_label.config(text=str(employees_total))
            self.attachments_count_label.config(text=str(attachments_total))
            self.last_update_label.config(text=datetime.now().strftime('%Y-%m-%d %H:%M'))
            
            # تحميل أحدث السجلات
            self.load_recent_records()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"خطأ في تحميل الإحصائيات: {e}")
    
    def load_recent_records(self):
        """تحميل أحدث السجلات"""
        # أحدث الوارد
        recent_incoming = self.db_manager.execute_query(
            "SELECT record_number, serial_number, title, registration_date "
            "FROM incoming_records ORDER BY id DESC LIMIT 10"
        )
        
        self.recent_incoming_tree.delete(*self.recent_incoming_tree.get_children())
        for record in recent_incoming:
            self.recent_incoming_tree.insert('', tk.END, values=record)
        
        # أحدث الصادر
        recent_outgoing = self.db_manager.execute_query(
            "SELECT record_number, outgoing_number, title, registration_date "
            "FROM outgoing_records ORDER BY id DESC LIMIT 10"
        )
        
        self.recent_outgoing_tree.delete(*self.recent_outgoing_tree.get_children())
        for record in recent_outgoing:
            self.recent_outgoing_tree.insert('', tk.END, values=record)
    
    def load_incoming_records(self):
        """تحميل سجلات الوارد"""
        query = """
        SELECT ir.id, ir.record_number, ir.incoming_number, ir.serial_number, ir.title,
               isrc.name, it.name, e.name, ir.registration_date
        FROM incoming_records ir
        LEFT JOIN incoming_sources isrc ON ir.incoming_source_id = isrc.id
        LEFT JOIN incoming_types it ON ir.incoming_type_id = it.id
        LEFT JOIN employees e ON ir.employee_id = e.id
        ORDER BY ir.id DESC
        """
        
        records = self.db_manager.execute_query(query)
        self.incoming_tree.delete(*self.incoming_tree.get_children())
        
        for record in records:
            self.incoming_tree.insert('', tk.END, values=record)
    
    def load_outgoing_records(self):
        """تحميل سجلات الصادر"""
        query = """
        SELECT orc.id, orc.record_number, orc.outgoing_number, orc.serial_number, orc.title,
               od.name, e.name, orc.registration_date
        FROM outgoing_records orc
        LEFT JOIN outgoing_destinations od ON orc.outgoing_destination_id = od.id
        LEFT JOIN employees e ON orc.employee_id = e.id
        ORDER BY orc.id DESC
        """
        
        records = self.db_manager.execute_query(query)
        self.outgoing_tree.delete(*self.outgoing_tree.get_children())
        
        for record in records:
            self.outgoing_tree.insert('', tk.END, values=record)
    
    def search_incoming(self, event=None):
        """بحث في سجلات الوارد"""
        search_term = self.incoming_search_entry.get().strip()
        if not search_term:
            self.load_incoming_records()
            return
        
        query = """
        SELECT ir.id, ir.record_number, ir.incoming_number, ir.serial_number, ir.title,
               isrc.name, it.name, e.name, ir.registration_date
        FROM incoming_records ir
        LEFT JOIN incoming_sources isrc ON ir.incoming_source_id = isrc.id
        LEFT JOIN incoming_types it ON ir.incoming_type_id = it.id
        LEFT JOIN employees e ON ir.employee_id = e.id
        WHERE ir.record_number LIKE ? OR ir.incoming_number LIKE ? OR ir.serial_number LIKE ? 
           OR ir.title LIKE ? OR isrc.name LIKE ? OR it.name LIKE ? OR e.name LIKE ?
        ORDER BY ir.id DESC
        """
        
        search_pattern = f"%{search_term}%"
        params = [search_pattern] * 7
        
        records = self.db_manager.execute_query(query, params)
        self.incoming_tree.delete(*self.incoming_tree.get_children())
        
        for record in records:
            self.incoming_tree.insert('', tk.END, values=record)
    
    def search_outgoing(self, event=None):
        """بحث في سجلات الصادر"""
        search_term = self.outgoing_search_entry.get().strip()
        if not search_term:
            self.load_outgoing_records()
            return
        
        query = """
        SELECT orc.id, orc.record_number, orc.outgoing_number, orc.serial_number, orc.title,
               od.name, e.name, orc.registration_date
        FROM outgoing_records orc
        LEFT JOIN outgoing_destinations od ON orc.outgoing_destination_id = od.id
        LEFT JOIN employees e ON orc.employee_id = e.id
        WHERE orc.record_number LIKE ? OR orc.outgoing_number LIKE ? OR orc.serial_number LIKE ? 
           OR orc.title LIKE ? OR od.name LIKE ? OR e.name LIKE ?
        ORDER BY orc.id DESC
        """
        
        search_pattern = f"%{search_term}%"
        params = [search_pattern] * 6
        
        records = self.db_manager.execute_query(query, params)
        self.outgoing_tree.delete(*self.outgoing_tree.get_children())
        
        for record in records:
            self.outgoing_tree.insert('', tk.END, values=record)
    
    def perform_search(self):
        """إجراء بحث متقدم"""
        search_term = self.search_entry.get().strip()
        search_type = self.search_type.get()
        
        if not search_term:
            messagebox.showwarning("تحذير", "يرجى إدخال نص للبحث")
            return
        
        self.search_results_tree.delete(*self.search_results_tree.get_children())
        
        if search_type == "incoming" or search_type == "both":
            # بحث في الوارد
            query = """
            SELECT '📥 وارد', record_number, incoming_number, serial_number, title, registration_date, id, 'incoming'
            FROM incoming_records
            WHERE record_number LIKE ? OR incoming_number LIKE ? OR serial_number LIKE ? 
               OR title LIKE ? OR details LIKE ?
            """
            search_pattern = f"%{search_term}%"
            params = [search_pattern] * 5
            
            results = self.db_manager.execute_query(query, params)
            for result in results:
                self.search_results_tree.insert('', tk.END, values=result, tags=('incoming',))
        
        if search_type == "outgoing" or search_type == "both":
            # بحث في الصادر
            query = """
            SELECT '📤 صادر', record_number, outgoing_number, serial_number, title, registration_date, id, 'outgoing'
            FROM outgoing_records
            WHERE record_number LIKE ? OR outgoing_number LIKE ? OR serial_number LIKE ? 
               OR title LIKE ? OR details LIKE ?
            """
            search_pattern = f"%{search_term}%"
            params = [search_pattern] * 5
            
            results = self.db_manager.execute_query(query, params)
            for result in results:
                self.search_results_tree.insert('', tk.END, values=result, tags=('outgoing',))
        
        # تلوين النتائج حسب النوع
        self.search_results_tree.tag_configure('incoming', background='#f0f8ff')
        self.search_results_tree.tag_configure('outgoing', background='#fff8f0')
        
        total_results = len(self.search_results_tree.get_children())
        if total_results > 0:
            messagebox.showinfo("نتائج البحث", f"تم العثور على {total_results} نتيجة")
        else:
            messagebox.showinfo("نتائج البحث", "لم يتم العثور على نتائج")
    
    def clear_search(self):
        """مسح نتائج البحث"""
        self.search_results_tree.delete(*self.search_results_tree.get_children())
        self.search_entry.delete(0, tk.END)
    
    def show_search_details(self):
        """عرض تفاصيل السجل المحدد في البحث"""
        selected = self.search_results_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل لعرض تفاصيله")
            return
        
        item = self.search_results_tree.item(selected[0])
        values = item['values']
        
        record_id = values[6]  # ID
        record_type = values[7]  # RecordType
        
        if record_type == 'incoming':
            self.open_incoming_form(record_id)
        else:
            self.open_outgoing_form(record_id)
    
    def export_search_results(self):
        """تصدير نتائج البحث"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("PDF files", "*.pdf"), ("Word files", "*.docx"), ("All files", "*.*")],
            title="حفظ نتائج البحث"
        )
        
        if file_path:
            try:
                # جمع البيانات من نتائج البحث
                data = []
                for item in self.search_results_tree.get_children():
                    values = self.search_results_tree.item(item)['values']
                    # استبعاد الأعمدة المخفية
                    visible_values = values[:6]
                    data.append(visible_values)
                
                columns = ['النوع', 'رقم السجل', 'الرقم', 'الرقم التسلسلي', 'العنوان', 'التاريخ']
                
                if file_path.endswith('.pdf'):
                    success = self.export_manager.export_to_pdf(data, columns, file_path, "نتائج البحث")
                elif file_path.endswith('.docx'):
                    success = self.export_manager.export_to_word(data, columns, file_path, "نتائج البحث")
                else:
                    success = self.export_manager.export_to_excel(data, columns, file_path, "نتائج البحث")
                
                if success:
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def export_incoming(self):
        """تصدير سجلات الوارد"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("PDF files", "*.pdf"), ("Word files", "*.docx"), ("All files", "*.*")],
            title="حفظ سجلات الوارد"
        )
        
        if file_path:
            try:
                data = []
                for item in self.incoming_tree.get_children():
                    values = self.incoming_tree.item(item)['values']
                    # استبعاد عمود ID
                    visible_values = values[1:]
                    data.append(visible_values)
                
                columns = ['رقم السجل', 'رقم الوارد', 'الرقم التسلسلي', 'العنوان', 
                          'جهة الوارد', 'النوع', 'الموظف', 'التاريخ']
                
                if file_path.endswith('.pdf'):
                    success = self.export_manager.export_to_pdf(data, columns, file_path, "سجلات الوارد")
                elif file_path.endswith('.docx'):
                    success = self.export_manager.export_to_word(data, columns, file_path, "سجلات الوارد")
                else:
                    success = self.export_manager.export_to_excel(data, columns, file_path, "سجلات الوارد")
                
                if success:
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def export_outgoing(self):
        """تصدير سجلات الصادر"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("PDF files", "*.pdf"), ("Word files", "*.docx"), ("All files", "*.*")],
            title="حفظ سجلات الصادر"
        )
        
        if file_path:
            try:
                data = []
                for item in self.outgoing_tree.get_children():
                    values = self.outgoing_tree.item(item)['values']
                    # استبعاد عمود ID
                    visible_values = values[1:]
                    data.append(visible_values)
                
                columns = ['رقم السجل', 'رقم الصادر', 'الرقم التسلسلي', 'العنوان', 
                          'جهة الصادر', 'الموظف', 'التاريخ']
                
                if file_path.endswith('.pdf'):
                    success = self.export_manager.export_to_pdf(data, columns, file_path, "سجلات الصادر")
                elif file_path.endswith('.docx'):
                    success = self.export_manager.export_to_word(data, columns, file_path, "سجلات الصادر")
                else:
                    success = self.export_manager.export_to_excel(data, columns, file_path, "سجلات الصادر")
                
                if success:
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def export_all_data(self):
        """تصدير جميع البيانات"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="حفظ جميع البيانات"
        )
        
        if file_path:
            try:
                # تصدير الوارد
                incoming_data = []
                for item in self.incoming_tree.get_children():
                    values = self.incoming_tree.item(item)['values']
                    incoming_data.append(values[1:])  # استبعاد ID
                
                # تصدير الصادر
                outgoing_data = []
                for item in self.outgoing_tree.get_children():
                    values = self.outgoing_tree.item(item)['values']
                    outgoing_data.append(values[1:])  # استبعاد ID
                
                # استخدام مدير التصدير لإنشاء ملف متعدد الأوراق
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # ورقة الوارد
                    incoming_columns = ['رقم السجل', 'رقم الوارد', 'الرقم التسلسلي', 'العنوان', 
                                      'جهة الوارد', 'النوع', 'الموظف', 'التاريخ']
                    pd.DataFrame(incoming_data, columns=incoming_columns).to_excel(
                        writer, sheet_name='الوارد', index=False
                    )
                    
                    # ورقة الصادر
                    outgoing_columns = ['رقم السجل', 'رقم الصادر', 'الرقم التسلسلي', 'العنوان', 
                                      'جهة الصادر', 'الموظف', 'التاريخ']
                    pd.DataFrame(outgoing_data, columns=outgoing_columns).to_excel(
                        writer, sheet_name='الصادر', index=False
                    )
                
                messagebox.showinfo("نجاح", f"تم تصدير جميع البيانات إلى: {file_path}")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def refresh_data(self):
        """تحديث جميع البيانات"""
        self.load_statistics()
        self.load_incoming_records()
        self.load_outgoing_records()
        messagebox.showinfo("نجاح", "تم تحديث البيانات بنجاح")
    
    def open_incoming_form(self, record_id=None):
        """فتح نموذج تسجيل وارد"""
        from gui.incoming_form import IncomingForm
        form_window = tk.Toplevel(self.root)
        IncomingForm(form_window, self.db_manager, self.file_manager, record_id)
        form_window.transient(self.root)
        form_window.grab_set()
    
    def open_outgoing_form(self, record_id=None):
        """فتح نموذج تسجيل صادر"""
        from gui.outgoing_form import OutgoingForm
        form_window = tk.Toplevel(self.root)
        OutgoingForm(form_window, self.db_manager, self.file_manager, record_id)
        form_window.transient(self.root)
        form_window.grab_set()
    
    def open_search_window(self):
        """فتح نافذة البحث المتقدم"""
        from gui.search_window import SearchWindow
        search_window = tk.Toplevel(self.root)
        SearchWindow(search_window, self.db_manager)
        search_window.transient(self.root)
        search_window.grab_set()
    
    def open_incoming_reports(self):
        """فتح تقارير الوارد"""
        from gui.reports_window import ReportsWindow
        reports_window = tk.Toplevel(self.root)
        ReportsWindow(reports_window, self.db_manager, self.export_manager, 'incoming')
        reports_window.transient(self.root)
        reports_window.grab_set()
    
    def open_outgoing_reports(self):
        """فتح تقارير الصادر"""
        from gui.reports_window import ReportsWindow
        reports_window = tk.Toplevel(self.root)
        ReportsWindow(reports_window, self.db_manager, self.export_manager, 'outgoing')
        reports_window.transient(self.root)
        reports_window.grab_set()
    
    def open_employee_reports(self):
        """فتح تقارير الموظفين"""
        try:
            from gui.employee_reports import EmployeeReportsWindow
            reports_window = tk.Toplevel(self.root)
            EmployeeReportsWindow(reports_window, self.db_manager, self.export_manager)
            reports_window.transient(self.root)
            reports_window.grab_set()
        except ImportError as e:
            messagebox.showerror("خطأ", f"لم يتم العثور على وحدة تقارير الموظفين: {e}")
    
    def open_comprehensive_report(self):
        """فتح التقرير الشامل"""
        try:
            from gui.comprehensive_report import ComprehensiveReportWindow
            report_window = tk.Toplevel(self.root)
            ComprehensiveReportWindow(report_window, self.db_manager, self.export_manager)
            report_window.transient(self.root)
            report_window.grab_set()
        except ImportError as e:
            messagebox.showerror("خطأ", f"لم يتم العثور على وحدة التقرير الشامل: {e}")
    
    def open_reference_management(self):
        """فتح إدارة الكيانات المرجعية"""
        try:
            from gui.reference_management import ReferenceManagement
            management_window = tk.Toplevel(self.root)
            ReferenceManagement(management_window, self.db_manager)
            management_window.transient(self.root)
            management_window.grab_set()
        except ImportError as e:
            messagebox.showerror("خطأ", f"لم يتم العثور على وحدة إدارة الكيانات المرجعية: {e}")
    
    def open_employee_management(self):
        """فتح إدارة الموظفين"""
        try:
            from gui.employee_management import EmployeeManagementWindow
            management_window = tk.Toplevel(self.root)
            EmployeeManagementWindow(management_window, self.db_manager)
            management_window.transient(self.root)
            management_window.grab_set()
        except ImportError as e:
            messagebox.showerror("خطأ", f"لم يتم العثور على وحدة إدارة الموظفين: {e}")
    
    def edit_incoming_record(self):
        """تعديل سجل وارد محدد"""
        selected = self.incoming_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل للتعديل")
            return
        
        record_id = self.incoming_tree.item(selected[0])['values'][0]
        self.open_incoming_form(record_id)
    
    def edit_outgoing_record(self):
        """تعديل سجل صادر محدد"""
        selected = self.outgoing_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل للتعديل")
            return
        
        record_id = self.outgoing_tree.item(selected[0])['values'][0]
        self.open_outgoing_form(record_id)
    
    def delete_incoming_record(self):
        """حذف سجل وارد"""
        selected = self.incoming_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل للحذف")
            return
        
        item = self.incoming_tree.item(selected[0])
        record_number = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف السجل '{record_number}'؟"):
            record_id = item['values'][0]
            
            try:
                # حذف المرفقات أولاً
                self.db_manager.execute_query(
                    "DELETE FROM attachments WHERE record_id = ? AND record_type = 'incoming'",
                    (record_id,)
                )
                
                # حذف السجل
                self.db_manager.execute_query(
                    "DELETE FROM incoming_records WHERE id = ?",
                    (record_id,)
                )
                
                messagebox.showinfo("نجاح", "تم حذف السجل بنجاح")
                self.load_incoming_records()
                self.load_statistics()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في حذف السجل: {e}")
    
    def delete_outgoing_record(self):
        """حذف سجل صادر"""
        selected = self.outgoing_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل للحذف")
            return
        
        item = self.outgoing_tree.item(selected[0])
        record_number = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف السجل '{record_number}'؟"):
            record_id = item['values'][0]
            
            try:
                # حذف المرفقات أولاً
                self.db_manager.execute_query(
                    "DELETE FROM attachments WHERE record_id = ? AND record_type = 'outgoing'",
                    (record_id,)
                )
                
                # حذف السجل
                self.db_manager.execute_query(
                    "DELETE FROM outgoing_records WHERE id = ?",
                    (record_id,)
                )
                
                messagebox.showinfo("نجاح", "تم حذف السجل بنجاح")
                self.load_outgoing_records()
                self.load_statistics()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في حذف السجل: {e}")
    
    def print_incoming(self):
        """طباعة سجلات الوارد"""
        try:
            if self.printer_manager and hasattr(self.printer_manager, 'quick_print_current_view'):
                if self.printer_manager.quick_print_current_view(self.incoming_tree, 'incoming'):
                    messagebox.showinfo("نجاح", "تم إرسال سجلات الوارد للطباعة")
            else:
                messagebox.showwarning("تحذير", "خاصية الطباعة غير متاحة حالياً")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الطباعة: {e}")
    
    def print_outgoing(self):
        """طباعة سجلات الصادر"""
        try:
            if self.printer_manager and hasattr(self.printer_manager, 'quick_print_current_view'):
                if self.printer_manager.quick_print_current_view(self.outgoing_tree, 'outgoing'):
                    messagebox.showinfo("نجاح", "تم إرسال سجلات الصادر للطباعة")
            else:
                messagebox.showwarning("تحذير", "خاصية الطباعة غير متاحة حالياً")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الطباعة: {e}")
    
    def print_search_results(self):
        """طباعة نتائج البحث"""
        try:
            if self.printer_manager and hasattr(self.printer_manager, 'quick_print_current_view'):
                if self.printer_manager.quick_print_current_view(self.search_results_tree, 'search'):
                    messagebox.showinfo("نجاح", "تم إرسال نتائج البحث للطباعة")
            else:
                messagebox.showwarning("تحذير", "خاصية الطباعة غير متاحة حالياً")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة نتائج البحث: {e}")
    
    def print_selected_incoming(self):
        """طباعة سجل وارد محدد"""
        try:
            if self.printer_manager and hasattr(self.printer_manager, 'print_selected_record'):
                if self.printer_manager.print_selected_record(self.incoming_tree, "وارد"):
                    messagebox.showinfo("نجاح", "تم إرسال السجل المحدد للطباعة")
            else:
                messagebox.showwarning("تحذير", "خاصية الطباعة غير متاحة حالياً")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة السجل المحدد: {e}")
    
    def print_selected_outgoing(self):
        """طباعة سجل صادر محدد"""
        try:
            if self.printer_manager and hasattr(self.printer_manager, 'print_selected_record'):
                if self.printer_manager.print_selected_record(self.outgoing_tree, "صادر"):
                    messagebox.showinfo("نجاح", "تم إرسال السجل المحدد للطباعة")
            else:
                messagebox.showwarning("تحذير", "خاصية الطباعة غير متاحة حالياً")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة السجل المحدد: {e}")
    
    def backup_database(self):
        """نسخ احتياطي لقاعدة البيانات"""
        try:
            backup_dir = "backups"
            os.makedirs(backup_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"database_backup_{timestamp}.db")
            
            shutil.copy2("data/database.db", backup_file)
            messagebox.showinfo("نجاح", f"تم إنشاء نسخة احتياطية في: {backup_file}")
        
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في إنشاء النسخة الاحتياطية: {e}")
    
    def restore_database(self):
        """استعادة نسخة احتياطية"""
        from tkinter import filedialog
        
        backup_file = filedialog.askopenfilename(
            title="اختر ملف النسخة الاحتياطية",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")]
        )
        
        if backup_file and os.path.exists(backup_file):
            if messagebox.askyesno("تأكيد", "هل أنت متأكد من استعادة النسخة الاحتياطية؟ سيتم فقدان جميع البيانات الحالية."):
                try:
                    shutil.copy2(backup_file, "data/database.db")
                    messagebox.showinfo("نجاح", "تم استعادة النسخة الاحتياطية بنجاح")
                    # إعادة تحميل البيانات
                    self.refresh_data()
                except Exception as e:
                    messagebox.showerror("خطأ", f"فشل في استعادة النسخة الاحتياطية: {e}")
    
    def show_user_guide(self):
        """عرض دليل المستخدم"""
        guide_text = """
📚 دليل استخدام نظام إدارة المراسلات - الإصدار 2.0

1. 📥 تسجيل السجلات:
   - استخدام قائمة الملف لتسجيل وارد أو صادر جديد
   - تعبئة جميع الحقول الإلزامية
   - إرفاق الملفات إذا لزم الأمر

2. 🔍 البحث والتصفية:
   - استخدام تبويب البحث للبحث المتقدم
   - استخدام حقول البحث في تبويبي الوارد والصادر
   - إمكانية التصدير والطباعة

3. 📊 التقارير:
   - إنشاء تقارير حسب الفترة الزمنية
   - تقارير الموظفين (الفاكسات والإيميلات)
   - التصدير إلى Excel, PDF, Word
   - إحصائيات مفصلة

4. ⚙️ الإدارة:
   - إدارة الكيانات المرجعية (الموظفين، الجهات، etc.)
   - إدارة الموظفين وتقارير الأداء
   - نسخ احتياطي للبيانات
   - استعادة النسخ الاحتياطية

5. 👥 تقارير الموظفين:
   - عرض إحصائيات الفاكسات لكل موظف
   - عرض إحصائيات الإيميلات لكل موظف
   - مقارنة أداء الموظفين
   - تصدير تقارير مفصلة

للحصول على مساعدة إضافية، يرجى التواصل مع الدعم الفني.
        """
        messagebox.showinfo("📚 دليل المستخدم", guide_text)
    
    def show_about(self):
        """عرض معلومات عن النظام"""
        about_text = """
نظام إدارة المراسلات - الإصدار 2.0

🎯 المميزات الرئيسية:
• 📥 تسجيل وحفظ سجلات الوارد والصادر
• 📎 إدارة المرفقات والوثائق
• 🔍 بحث متقدم في السجلات
• 📊 إعداد التقارير والتصدير
• 👥 تقارير أداء الموظفين (الفاكسات والإيميلات)
• 🎨 واجهة مستخدم محسنة وسهلة الاستخدام
• 💾 نسخ احتياطي تلقائي

🛠️ التقنيات المستخدمة:
• Python 3.7+
• Tkinter للواجهة الرسومية
• SQLite لقاعدة البيانات
• مكتبات متقدمة للتصدير والطباعة
• دعم كامل للغة العربية

👨‍💻 المطور: فريق التطوير
📧 البريد الإلكتروني: support@company.com
🌐 الموقع: www.company.com

© 2024 جميع الحقوق محفوظة
        """
        messagebox.showinfo("ℹ️ حول النظام", about_text)