import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class SearchWindow:
    def __init__(self, parent, db_manager):
        self.parent = parent
        self.db_manager = db_manager
        
        self.setup_window()
        self.create_widgets()
    
    def setup_window(self):
        """إعداد نافذة البحث"""
        self.parent.title("البحث المتقدم")
        self.parent.geometry("900x600")
        self.parent.resizable(True, True)
        
        # الإطار الرئيسي
        self.main_frame = ttk.Frame(self.parent, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # إطار معايير البحث
        criteria_frame = ttk.LabelFrame(self.main_frame, text="معايير البحث", padding=10)
        criteria_frame.pack(fill=tk.X, pady=5)
        
        self.create_search_criteria(criteria_frame)
        
        # إطار النتائج
        results_frame = ttk.LabelFrame(self.main_frame, text="نتائج البحث", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.setup_results_section(results_frame)
        
        # أزرار التحكم
        self.create_control_buttons()
    
    def create_search_criteria(self, parent):
        """إنشاء معايير البحث"""
        # نوع البحث
        type_frame = ttk.Frame(parent)
        type_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(type_frame, text="نوع السجل:").pack(side=tk.RIGHT, padx=5)
        self.record_type = tk.StringVar(value="both")
        ttk.Radiobutton(type_frame, text="الوارد فقط", variable=self.record_type, value="incoming").pack(side=tk.RIGHT, padx=5)
        ttk.Radiobutton(type_frame, text="الصادر فقط", variable=self.record_type, value="outgoing").pack(side=tk.RIGHT, padx=5)
        ttk.Radiobutton(type_frame, text="الكل", variable=self.record_type, value="both").pack(side=tk.RIGHT, padx=5)
        
        # نص البحث
        text_frame = ttk.Frame(parent)
        text_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(text_frame, text="نص البحث:").pack(side=tk.RIGHT, padx=5)
        self.search_text = ttk.Entry(text_frame, width=50)
        self.search_text.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(text_frame, text="حقل البحث:").pack(side=tk.RIGHT, padx=5)
        self.search_field = ttk.Combobox(text_frame, width=15, state="readonly")
        self.search_field['values'] = ['جميع الحقول', 'رقم السجل', 'الرقم', 'الرقم التسلسلي', 'العنوان', 'التفاصيل']
        self.search_field.set('جميع الحقول')
        self.search_field.pack(side=tk.RIGHT, padx=5)
        
        # نطاق التاريخ
        date_frame = ttk.Frame(parent)
        date_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(date_frame, text="إلى:").pack(side=tk.RIGHT, padx=5)
        self.end_date = ttk.Entry(date_frame, width=12)
        self.end_date.insert(0, datetime.now().strftime('%Y-%m-%d'))
        self.end_date.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(date_frame, text="من:").pack(side=tk.RIGHT, padx=5)
        self.start_date = ttk.Entry(date_frame, width=12)
        self.start_date.insert(0, datetime.now().strftime('%Y-%m-01'))  # أول يوم من الشهر
        self.start_date.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(date_frame, text="نطاق التاريخ:").pack(side=tk.RIGHT, padx=5)
        
        # معايير إضافية
        extra_frame = ttk.Frame(parent)
        extra_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(extra_frame, text="الموظف:").pack(side=tk.RIGHT, padx=5)
        self.employee_combo = ttk.Combobox(extra_frame, width=20, state="readonly")
        self.employee_combo.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(extra_frame, text="الاختصاص:").pack(side=tk.RIGHT, padx=5)
        self.specialization_combo = ttk.Combobox(extra_frame, width=20, state="readonly")
        self.specialization_combo.pack(side=tk.RIGHT, padx=5)
        
        # تحميل البيانات المرجعية
        self.load_reference_data()
    
    def setup_results_section(self, parent):
        """إعداد قسم النتائج"""
        # أزرار التحكم بالنتائج
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="عرض التفاصيل", 
                  command=self.show_details).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="تصدير النتائج", 
                  command=self.export_results).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="طباعة", 
                  command=self.print_results).pack(side=tk.RIGHT, padx=5)
        
        # جدول النتائج
        columns = ('النوع', 'رقم السجل', 'الرقم', 'الرقم التسلسلي', 'العنوان', 
                  'الجهة', 'الموظف', 'التاريخ', 'ID', 'RecordType')
        
        self.results_tree = ttk.Treeview(parent, columns=columns, show='headings')
        
        # إخفاء الأعمدة الإضافية
        self.results_tree.column('ID', width=0, stretch=tk.NO)
        self.results_tree.column('RecordType', width=0, stretch=tk.NO)
        
        # تعيين العناوين والأبعاد
        visible_columns = {
            'النوع': 80, 'رقم السجل': 120, 'الرقم': 120, 'الرقم التسلسلي': 100,
            'العنوان': 200, 'الجهة': 150, 'الموظف': 120, 'التاريخ': 100
        }
        
        for col, width in visible_columns.items():
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=width)
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.results_tree.bind('<Double-1>', self.on_double_click)
    
    def create_control_buttons(self):
        """إنشاء أزرار التحكم"""
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="بحث", 
                  command=self.perform_search, style='Accent.TButton').pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="مسح النتائج", 
                  command=self.clear_results).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="إغلاق", 
                  command=self.parent.destroy).pack(side=tk.RIGHT, padx=5)
    
    def load_reference_data(self):
        """تحميل البيانات المرجعية"""
        # الموظفين
        employees = self.db_manager.execute_query(
            "SELECT id, name FROM employees WHERE is_active = 1 ORDER BY name"
        )
        employee_names = [name for _, name in employees]
        self.employee_combo['values'] = [''] + employee_names
        
        # الاختصاصات
        specializations = self.db_manager.execute_query(
            "SELECT id, name FROM specializations ORDER BY name"
        )
        specialization_names = [name for _, name in specializations]
        self.specialization_combo['values'] = [''] + specialization_names
    
    def build_search_query(self):
        """بناء استعلام البحث"""
        search_type = self.record_type.get()
        search_text = self.search_text.get().strip()
        search_field = self.search_field.get()
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        employee = self.employee_combo.get()
        specialization = self.specialization_combo.get()
        
        queries = []
        params = []
        
        if search_type in ['incoming', 'both']:
            query = """
            SELECT 'وارد' as type, record_number, incoming_number as number, 
                   serial_number, title, 
                   (SELECT name FROM incoming_sources WHERE id = incoming_source_id) as entity,
                   (SELECT name FROM employees WHERE id = employee_id) as employee,
                   registration_date, id, 'incoming' as record_type
            FROM incoming_records 
            WHERE 1=1
            """
            
            if search_text:
                if search_field == 'جميع الحقول':
                    query += " AND (record_number LIKE ? OR incoming_number LIKE ? OR serial_number LIKE ? OR title LIKE ? OR details LIKE ?)"
                    params.extend([f"%{search_text}%"] * 5)
                elif search_field == 'رقم السجل':
                    query += " AND record_number LIKE ?"
                    params.append(f"%{search_text}%")
                elif search_field == 'الرقم':
                    query += " AND incoming_number LIKE ?"
                    params.append(f"%{search_text}%")
                elif search_field == 'الرقم التسلسلي':
                    query += " AND serial_number LIKE ?"
                    params.append(f"%{search_text}%")
                elif search_field == 'العنوان':
                    query += " AND title LIKE ?"
                    params.append(f"%{search_text}%")
                elif search_field == 'التفاصيل':
                    query += " AND details LIKE ?"
                    params.append(f"%{search_text}%")
            
            if start_date and end_date:
                query += " AND registration_date BETWEEN ? AND ?"
                params.extend([start_date, end_date])
            
            if employee:
                query += " AND employee_id IN (SELECT id FROM employees WHERE name = ?)"
                params.append(employee)
            
            if specialization:
                query += " AND specialization_id IN (SELECT id FROM specializations WHERE name = ?)"
                params.append(specialization)
            
            queries.append((query, params))
        
        if search_type in ['outgoing', 'both']:
            query = """
            SELECT 'صادر' as type, record_number, outgoing_number as number, 
                   serial_number, title, 
                   (SELECT name FROM outgoing_destinations WHERE id = outgoing_destination_id) as entity,
                   (SELECT name FROM employees WHERE id = employee_id) as employee,
                   registration_date, id, 'outgoing' as record_type
            FROM outgoing_records 
            WHERE 1=1
            """
            
            params_out = []
            
            if search_text:
                if search_field == 'جميع الحقول':
                    query += " AND (record_number LIKE ? OR outgoing_number LIKE ? OR serial_number LIKE ? OR title LIKE ? OR details LIKE ?)"
                    params_out.extend([f"%{search_text}%"] * 5)
                elif search_field == 'رقم السجل':
                    query += " AND record_number LIKE ?"
                    params_out.append(f"%{search_text}%")
                elif search_field == 'الرقم':
                    query += " AND outgoing_number LIKE ?"
                    params_out.append(f"%{search_text}%")
                elif search_field == 'الرقم التسلسلي':
                    query += " AND serial_number LIKE ?"
                    params_out.append(f"%{search_text}%")
                elif search_field == 'العنوان':
                    query += " AND title LIKE ?"
                    params_out.append(f"%{search_text}%")
                elif search_field == 'التفاصيل':
                    query += " AND details LIKE ?"
                    params_out.append(f"%{search_text}%")
            
            if start_date and end_date:
                query += " AND registration_date BETWEEN ? AND ?"
                params_out.extend([start_date, end_date])
            
            if employee:
                query += " AND employee_id IN (SELECT id FROM employees WHERE name = ?)"
                params_out.append(employee)
            
            if specialization:
                query += " AND specialization_id IN (SELECT id FROM specializations WHERE name = ?)"
                params_out.append(specialization)
            
            queries.append((query, params_out))
        
        return queries
    
    def perform_search(self):
        """إجراء البحث"""
        queries = self.build_search_query()
        
        if not queries:
            messagebox.showwarning("تحذير", "يرجى تحديد معايير البحث")
            return
        
        # مسح النتائج السابقة
        self.results_tree.delete(*self.results_tree.get_children())
        
        total_results = 0
        
        for query, params in queries:
            try:
                results = self.db_manager.execute_query(query, params)
                total_results += len(results)
                
                for result in results:
                    # استخراج البيانات المرئية فقط (بدون ID و RecordType)
                    visible_data = result[:8]
                    full_data = result  # كل البيانات بما فيها ID و RecordType
                    self.results_tree.insert('', tk.END, values=full_data, tags=(result[0],))
                
            except Exception as e:
                messagebox.showerror("خطأ", f"خطأ في البحث: {e}")
                return
        
        # تلوين الصفوف حسب النوع
        self.results_tree.tag_configure('وارد', background='#f0f8ff')
        self.results_tree.tag_configure('صادر', background='#fff8f0')
        
        messagebox.showinfo("نتائج البحث", f"تم العثور على {total_results} نتيجة")
    
    def clear_results(self):
        """مسح نتائج البحث"""
        self.results_tree.delete(*self.results_tree.get_children())
        self.search_text.delete(0, tk.END)
        self.employee_combo.set('')
        self.specialization_combo.set('')
    
    def on_double_click(self, event):
        """معالجة النقر المزدوج على نتيجة"""
        selected = self.results_tree.selection()
        if selected:
            self.show_details()
    
    def show_details(self):
        """عرض تفاصيل السجل المحدد"""
        selected = self.results_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار سجل لعرض تفاصيله")
            return
        
        item = self.results_tree.item(selected[0])
        values = item['values']
        
        record_id = values[8]  # ID
        record_type = values[9]  # RecordType
        
        if record_type == 'incoming':
            from gui.incoming_form import IncomingForm
            form_window = tk.Toplevel(self.parent)
            IncomingForm(form_window, self.db_manager, None, record_id)
        else:
            from gui.outgoing_form import OutgoingForm
            form_window = tk.Toplevel(self.parent)
            OutgoingForm(form_window, self.db_manager, None, record_id)
    
    def export_results(self):
        """تصدير نتائج البحث"""
        selected = self.results_tree.get_children()
        if not selected:
            messagebox.showwarning("تحذير", "لا توجد نتائج للتصدير")
            return
        
        # هنا يمكن إضافة منطق التصدير إلى Excel أو PDF
        messagebox.showinfo("تصدير", "سيتم تطوير خاصية التصدير في المستقبل")
    
    def print_results(self):
        """طباعة نتائج البحث"""
        selected = self.results_tree.get_children()
        if not selected:
            messagebox.showwarning("تحذير", "لا توجد نتائج للطباعة")
            return
        
        messagebox.showinfo("طباعة", "سيتم تطوير خاصية الطباعة في المستقبل")