import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os

class EmployeeReportsWindow:
    def __init__(self, parent, db_manager, export_manager):
        self.parent = parent
        self.db_manager = db_manager
        self.export_manager = export_manager
        
        self.setup_window()
        self.create_widgets()
        self.generate_report()
    
    def setup_window(self):
        """إعداد نافذة تقارير الموظفين"""
        self.parent.title("تقرير إحصاءات الموظفين")
        self.parent.geometry("1000x700")
        self.parent.resizable(True, True)
        
        # الإطار الرئيسي
        self.main_frame = ttk.Frame(self.parent, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # إطار معايير التقرير
        criteria_frame = ttk.LabelFrame(self.main_frame, text="معايير التقرير", padding=10)
        criteria_frame.pack(fill=tk.X, pady=5)
        
        self.create_report_criteria(criteria_frame)
        
        # إطار الإحصائيات
        stats_frame = ttk.LabelFrame(self.main_frame, text="إحصائيات عامة", padding=10)
        stats_frame.pack(fill=tk.X, pady=5)
        
        self.create_statistics_section(stats_frame)
        
        # إطار النتائج
        results_frame = ttk.LabelFrame(self.main_frame, text="تفاصيل الموظفين", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.setup_results_section(results_frame)
        
        # أزرار التحكم
        self.create_control_buttons()
    
    def create_report_criteria(self, parent):
        """إنشاء معايير التقرير"""
        criteria_frame = ttk.Frame(parent)
        criteria_frame.pack(fill=tk.X, pady=5)
        
        # نطاق التاريخ
        ttk.Label(criteria_frame, text="إلى:").pack(side=tk.RIGHT, padx=5)
        self.end_date = ttk.Entry(criteria_frame, width=12)
        self.end_date.insert(0, datetime.now().strftime('%Y-%m-%d'))
        self.end_date.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(criteria_frame, text="من:").pack(side=tk.RIGHT, padx=5)
        self.start_date = ttk.Entry(criteria_frame, width=12)
        self.start_date.insert(0, datetime.now().strftime('%Y-%m-01'))
        self.start_date.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(criteria_frame, text="نطاق التاريخ:").pack(side=tk.RIGHT, padx=5)
        
        # قسم التصفية
        filter_frame = ttk.Frame(parent)
        filter_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(filter_frame, text="القسم:").pack(side=tk.RIGHT, padx=5)
        self.department_combo = ttk.Combobox(filter_frame, width=20, state="readonly")
        self.department_combo.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(filter_frame, text="حالة الموظف:").pack(side=tk.RIGHT, padx=5)
        self.status_combo = ttk.Combobox(filter_frame, width=15, state="readonly")
        self.status_combo['values'] = ['الكل', 'نشط فقط', 'غير نشط فقط']
        self.status_combo.set('نشط فقط')
        self.status_combo.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(filter_frame, text="توليد التقرير", 
                  command=self.generate_report).pack(side=tk.RIGHT, padx=5)
        
        # تحميل أقسام الموظفين
        self.load_departments()
    
    def create_statistics_section(self, parent):
        """إنشاء قسم الإحصائيات العامة"""
        stats_grid = ttk.Frame(parent)
        stats_grid.pack(fill=tk.X)
        
        # إحصائيات عامة
        ttk.Label(stats_grid, text="إجمالي الموظفين:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=10)
        self.total_employees_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.total_employees_label.grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Label(stats_grid, text="إجمالي الفاكسات:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=10)
        self.total_faxes_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.total_faxes_label.grid(row=0, column=3, sticky='w', padx=5)
        
        ttk.Label(stats_grid, text="إجمالي الإيميلات:", font=('Arial', 10, 'bold')).grid(row=0, column=4, sticky='w', padx=10)
        self.total_emails_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.total_emails_label.grid(row=0, column=5, sticky='w', padx=5)
        
        ttk.Label(stats_grid, text="متوسط الفاكسات:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=10)
        self.avg_faxes_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.avg_faxes_label.grid(row=1, column=1, sticky='w', padx=5)
        
        ttk.Label(stats_grid, text="متوسط الإيميلات:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=10)
        self.avg_emails_label = ttk.Label(stats_grid, text="0", font=('Arial', 10))
        self.avg_emails_label.grid(row=1, column=3, sticky='w', padx=5)
        
        ttk.Label(stats_grid, text="أعلى موظف في الفاكسات:", font=('Arial', 10, 'bold')).grid(row=1, column=4, sticky='w', padx=10)
        self.top_fax_employee_label = ttk.Label(stats_grid, text="---", font=('Arial', 10))
        self.top_fax_employee_label.grid(row=1, column=5, sticky='w', padx=5)
    
    def setup_results_section(self, parent):
        """إعداد قسم النتائج"""
        # أزرار التحكم بالنتائج
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="تصدير إلى Excel", 
                  command=self.export_to_excel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="تصدير إلى PDF", 
                  command=self.export_to_pdf).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="طباعة التقرير", 
                  command=self.print_report).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="عرض التفاصيل", 
                  command=self.show_employee_details).pack(side=tk.RIGHT, padx=5)
        
        # جدول النتائج
        columns = ('ID', 'اسم الموظف', 'القسم', 'المنصب', 'البريد الإلكتروني', 
                  'عدد الفاكسات', 'عدد الإيميلات', 'المجموع', 'النسبة %', 'الحالة')
        
        self.results_tree = ttk.Treeview(parent, columns=columns, show='headings')
        
        # إخفاء عمود ID
        self.results_tree.column('ID', width=0, stretch=tk.NO)
        
        # تعيين العناوين والأبعاد
        column_widths = {
            'اسم الموظف': 150,
            'القسم': 120,
            'المنصب': 120,
            'البريد الإلكتروني': 150,
            'عدد الفاكسات': 100,
            'عدد الإيميلات': 100,
            'المجموع': 80,
            'النسبة %': 80,
            'الحالة': 80
        }
        
        for col, width in column_widths.items():
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=width)
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.results_tree.bind('<Double-1>', lambda e: self.show_employee_details())
    
    def create_control_buttons(self):
        """إنشاء أزرار التحكم"""
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="إغلاق", 
                  command=self.parent.destroy).pack(side=tk.RIGHT, padx=5)
    
    def load_departments(self):
        """تحميل أقسام الموظفين"""
        departments = self.db_manager.execute_query(
            "SELECT DISTINCT department FROM employees WHERE department IS NOT NULL ORDER BY department"
        )
        department_names = ['الكل'] + [dept[0] for dept in departments]
        self.department_combo['values'] = department_names
        self.department_combo.set('الكل')
    
    def generate_report(self):
        """توليد تقرير الموظفين"""
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        department = self.department_combo.get()
        status_filter = self.status_combo.get()
        
        try:
            # بناء استعلام الموظفين
            employee_query = "SELECT id, name, department, position, email, fax_count, email_count, is_active FROM employees WHERE 1=1"
            params = []
            
            if department != 'الكل':
                employee_query += " AND department = ?"
                params.append(department)
            
            if status_filter == 'نشط فقط':
                employee_query += " AND is_active = 1"
            elif status_filter == 'غير نشط فقط':
                employee_query += " AND is_active = 0"
            
            employee_query += " ORDER BY name"
            
            employees = self.db_manager.execute_query(employee_query, params)
            
            # حساب إحصائيات الفاكسات والإيميلات من سجلات الوارد والصادر
            employee_stats = {}
            total_faxes = 0
            total_emails = 0
            
            for emp in employees:
                emp_id = emp[0]
                
                # حساب الفاكسات (سجلات الصادر)
                fax_count = self.db_manager.execute_query(
                    """SELECT COUNT(*) FROM outgoing_records 
                    WHERE employee_id = ? AND registration_date BETWEEN ? AND ?""",
                    (emp_id, start_date, end_date)
                )[0][0]
                
                # حساب الإيميلات (سجلات الوارد من نوع إيميل)
                email_count = self.db_manager.execute_query(
                    """SELECT COUNT(*) FROM incoming_records 
                    WHERE employee_id = ? AND registration_date BETWEEN ? AND ? 
                    AND incoming_type_id IN (SELECT id FROM incoming_types WHERE name LIKE '%إيميل%' OR name LIKE '%email%')""",
                    (emp_id, start_date, end_date)
                )[0][0]
                
                employee_stats[emp_id] = {
                    'fax_count': fax_count,
                    'email_count': email_count,
                    'total': fax_count + email_count
                }
                
                total_faxes += fax_count
                total_emails += email_count
            
            # تحديث الإحصائيات العامة
            self.update_statistics(employees, employee_stats, total_faxes, total_emails)
            
            # عرض البيانات في الجدول
            self.display_employee_data(employees, employee_stats, total_faxes + total_emails)
            
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في توليد التقرير: {e}")
    
    def update_statistics(self, employees, employee_stats, total_faxes, total_emails):
        """تحديث الإحصائيات العامة"""
        total_employees = len(employees)
        self.total_employees_label.config(text=str(total_employees))
        self.total_faxes_label.config(text=str(total_faxes))
        self.total_emails_label.config(text=str(total_emails))
        
        # حساب المتوسطات
        if total_employees > 0:
            avg_faxes = total_faxes / total_employees
            avg_emails = total_emails / total_employees
            self.avg_faxes_label.config(text=f"{avg_faxes:.1f}")
            self.avg_emails_label.config(text=f"{avg_emails:.1f}")
        else:
            self.avg_faxes_label.config(text="0")
            self.avg_emails_label.config(text="0")
        
        # إيجاد الموظف الأكثر في الفاكسات
        top_fax_employee = "---"
        max_faxes = 0
        for emp in employees:
            emp_id = emp[0]
            fax_count = employee_stats.get(emp_id, {}).get('fax_count', 0)
            if fax_count > max_faxes:
                max_faxes = fax_count
                top_fax_employee = emp[1]
        
        self.top_fax_employee_label.config(text=f"{top_fax_employee} ({max_faxes})")
    
    def display_employee_data(self, employees, employee_stats, total_communications):
        """عرض بيانات الموظفين"""
        self.results_tree.delete(*self.results_tree.get_children())
        
        for emp in employees:
            emp_id = emp[0]
            name = emp[1]
            department = emp[2] or "غير محدد"
            position = emp[3] or "غير محدد"
            email = emp[4] or "غير متوفر"
            is_active = emp[7]
            
            stats = employee_stats.get(emp_id, {'fax_count': 0, 'email_count': 0, 'total': 0})
            fax_count = stats['fax_count']
            email_count = stats['email_count']
            total = stats['total']
            
            # حساب النسبة المئوية
            percentage = (total / total_communications * 100) if total_communications > 0 else 0
            
            status_text = "نشط" if is_active else "غير نشط"
            
            self.results_tree.insert('', tk.END, values=(
                emp_id, name, department, position, email, 
                fax_count, email_count, total, f"{percentage:.1f}%", status_text
            ))
    
    def show_employee_details(self):
        """عرض تفاصيل الموظف المحدد"""
        selected = self.results_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف لعرض تفاصيله")
            return
        
        item = self.results_tree.item(selected[0])
        values = item['values']
        emp_id = values[0]
        emp_name = values[1]
        
        # إنشاء نافذة التفاصيل
        details_window = tk.Toplevel(self.parent)
        details_window.title(f"تفاصيل الموظف - {emp_name}")
        details_window.geometry("600x400")
        
        self.create_employee_details(details_window, emp_id, emp_name)
    
    def create_employee_details(self, parent, emp_id, emp_name):
        """إنشاء تفاصيل الموظف"""
        main_frame = ttk.Frame(parent, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # معلومات أساسية
        info_frame = ttk.LabelFrame(main_frame, text="المعلومات الأساسية", padding=10)
        info_frame.pack(fill=tk.X, pady=5)
        
        # الحصول على معلومات الموظف
        emp_info = self.db_manager.execute_query(
            "SELECT name, department, position, email, phone, fax_count, email_count FROM employees WHERE id = ?",
            (emp_id,)
        )
        
        if emp_info:
            info = emp_info[0]
            ttk.Label(info_frame, text=f"الاسم: {info[0]}", font=('Arial', 10, 'bold')).pack(anchor='w')
            ttk.Label(info_frame, text=f"القسم: {info[1] or 'غير محدد'}").pack(anchor='w')
            ttk.Label(info_frame, text=f"المنصب: {info[2] or 'غير محدد'}").pack(anchor='w')
            ttk.Label(info_frame, text=f"البريد الإلكتروني: {info[3] or 'غير متوفر'}").pack(anchor='w')
            ttk.Label(info_frame, text=f"الهاتف: {info[4] or 'غير متوفر'}").pack(anchor='w')
        
        # الإحصائيات التفصيلية
        stats_frame = ttk.LabelFrame(main_frame, text="الإحصائيات التفصيلية", padding=10)
        stats_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # الحصول على سجلات الموظف
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        
        # سجلات الصادر (الفاكسات)
        outgoing_records = self.db_manager.execute_query(
            """SELECT record_number, title, registration_date 
            FROM outgoing_records 
            WHERE employee_id = ? AND registration_date BETWEEN ? AND ?
            ORDER BY registration_date DESC""",
            (emp_id, start_date, end_date)
        )
        
        # سجلات الوارد (الإيميلات)
        incoming_email_records = self.db_manager.execute_query(
            """SELECT ir.record_number, ir.title, ir.registration_date, it.name as type_name
            FROM incoming_records ir
            LEFT JOIN incoming_types it ON ir.incoming_type_id = it.id
            WHERE ir.employee_id = ? AND ir.registration_date BETWEEN ? AND ?
            AND (it.name LIKE '%إيميل%' OR it.name LIKE '%email%')
            ORDER BY ir.registration_date DESC""",
            (emp_id, start_date, end_date)
        )
        
        # إنشاء تبويبات للتفاصيل
        notebook = ttk.Notebook(stats_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # تبويب الفاكسات
        fax_frame = ttk.Frame(notebook)
        notebook.add(fax_frame, text=f"الفاكسات ({len(outgoing_records)})")
        
        fax_columns = ('رقم السجل', 'العنوان', 'التاريخ')
        fax_tree = ttk.Treeview(fax_frame, columns=fax_columns, show='headings')
        for col in fax_columns:
            fax_tree.heading(col, text=col)
            fax_tree.column(col, width=150)
        
        for record in outgoing_records:
            fax_tree.insert('', tk.END, values=record)
        
        fax_scrollbar = ttk.Scrollbar(fax_frame, orient=tk.VERTICAL, command=fax_tree.yview)
        fax_tree.configure(yscrollcommand=fax_scrollbar.set)
        fax_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fax_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # تبويب الإيميلات
        email_frame = ttk.Frame(notebook)
        notebook.add(email_frame, text=f"الإيميلات ({len(incoming_email_records)})")
        
        email_columns = ('رقم السجل', 'العنوان', 'التاريخ', 'النوع')
        email_tree = ttk.Treeview(email_frame, columns=email_columns, show='headings')
        for col in email_columns:
            email_tree.heading(col, text=col)
            email_tree.column(col, width=120)
        
        for record in incoming_email_records:
            email_tree.insert('', tk.END, values=record)
        
        email_scrollbar = ttk.Scrollbar(email_frame, orient=tk.VERTICAL, command=email_tree.yview)
        email_tree.configure(yscrollcommand=email_scrollbar.set)
        email_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        email_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def export_to_excel(self):
        """تصدير إلى Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="حفظ تقرير الموظفين"
        )
        
        if file_path:
            try:
                # جمع البيانات من الجدول
                data = []
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item)['values']
                    # استبعاد عمود ID
                    visible_values = values[1:]
                    data.append(visible_values)
                
                columns = ['اسم الموظف', 'القسم', 'المنصب', 'البريد الإلكتروني', 
                          'عدد الفاكسات', 'عدد الإيميلات', 'المجموع', 'النسبة %', 'الحالة']
                
                if self.export_manager.export_to_excel(data, columns, file_path, "تقرير الموظفين"):
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير إلى Excel")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def export_to_pdf(self):
        """تصدير إلى PDF"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            title="حفظ تقرير الموظفين"
        )
        
        if file_path:
            try:
                # جمع البيانات من الجدول
                data = []
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item)['values']
                    # استبعاد عمود ID
                    visible_values = values[1:]
                    data.append(visible_values)
                
                columns = ['اسم الموظف', 'القسم', 'المنصب', 'البريد الإلكتروني', 
                          'عدد الفاكسات', 'عدد الإيميلات', 'المجموع', 'النسبة %', 'الحالة']
                
                title = f"تقرير إحصاءات الموظفين - {datetime.now().strftime('%Y-%m-%d')}"
                
                if self.export_manager.export_to_pdf(data, columns, file_path, title):
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير إلى PDF")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def print_report(self):
        """طباعة التقرير"""
        messagebox.showinfo("طباعة", "سيتم تطوير خاصية الطباعة في المستقبل")