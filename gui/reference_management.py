import tkinter as tk
from tkinter import ttk, messagebox

class ReferenceManagement:
    def __init__(self, parent, db_manager):
        self.parent = parent
        self.db_manager = db_manager
        
        self.setup_window()
        self.create_widgets()
        self.load_data()
    
    def setup_window(self):
        """إعداد نافذة إدارة الكيانات المرجعية"""
        self.parent.title("إدارة الكيانات المرجعية")
        self.parent.geometry("800x600")
        self.parent.resizable(True, True)
        
        self.main_frame = ttk.Frame(self.parent, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # إنشاء تبويبات للكيانات المختلفة
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # تبويب جهات الوارد
        self.incoming_sources_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.incoming_sources_frame, text="جهات الوارد")
        
        # تبويب جهات الصادر
        self.outgoing_destinations_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.outgoing_destinations_frame, text="جهات الصادر")
        
        # تبويب أنواع الوارد
        self.incoming_types_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.incoming_types_frame, text="أنواع الوارد")
        
        # تبويب الموظفين
        self.employees_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.employees_frame, text="الموظفين")
        
        # تبويب الاختصاصات
        self.specializations_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.specializations_frame, text="الاختصاصات")
        
        # إعداد كل تبويب
        self.setup_incoming_sources_tab()
        self.setup_outgoing_destinations_tab()
        self.setup_incoming_types_tab()
        self.setup_employees_tab()
        self.setup_specializations_tab()
    
    def setup_incoming_sources_tab(self):
        """إعداد تبويب جهات الوارد"""
        # إطار الإضافة
        add_frame = ttk.LabelFrame(self.incoming_sources_frame, text="إضافة جهة وارد جديدة", padding=10)
        add_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_frame, text="اسم جهة الوارد:").pack(side=tk.RIGHT, padx=5)
        self.new_source_name = ttk.Entry(add_frame, width=30)
        self.new_source_name.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(add_frame, text="الوصف:").pack(side=tk.RIGHT, padx=5)
        self.new_source_desc = ttk.Entry(add_frame, width=40)
        self.new_source_desc.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(add_frame, text="إضافة", 
                  command=self.add_incoming_source).pack(side=tk.RIGHT, padx=5)
        
        # جدول جهات الوارد
        columns = ('ID', 'الاسم', 'الوصف')
        self.sources_tree = ttk.Treeview(self.incoming_sources_frame, columns=columns, show='headings')
        self.sources_tree.column('ID', width=0, stretch=tk.NO)
        self.sources_tree.heading('الاسم', text='الاسم')
        self.sources_tree.heading('الوصف', text='الوصف')
        
        # أزرار التحكم
        control_frame = ttk.Frame(self.incoming_sources_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="تعديل", 
                  command=self.edit_incoming_source).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="حذف", 
                  command=self.delete_incoming_source).pack(side=tk.RIGHT, padx=5)
        
        scrollbar = ttk.Scrollbar(self.incoming_sources_frame, orient=tk.VERTICAL, command=self.sources_tree.yview)
        self.sources_tree.configure(yscrollcommand=scrollbar.set)
        
        self.sources_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_outgoing_destinations_tab(self):
        """إعداد تبويب جهات الصادر"""
        # إطار الإضافة
        add_frame = ttk.LabelFrame(self.outgoing_destinations_frame, text="إضافة جهة صادر جديدة", padding=10)
        add_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_frame, text="اسم جهة الصادر:").pack(side=tk.RIGHT, padx=5)
        self.new_destination_name = ttk.Entry(add_frame, width=30)
        self.new_destination_name.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(add_frame, text="الوصف:").pack(side=tk.RIGHT, padx=5)
        self.new_destination_desc = ttk.Entry(add_frame, width=40)
        self.new_destination_desc.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(add_frame, text="إضافة", 
                  command=self.add_outgoing_destination).pack(side=tk.RIGHT, padx=5)
        
        # جدول جهات الصادر
        columns = ('ID', 'الاسم', 'الوصف')
        self.destinations_tree = ttk.Treeview(self.outgoing_destinations_frame, columns=columns, show='headings')
        self.destinations_tree.column('ID', width=0, stretch=tk.NO)
        self.destinations_tree.heading('الاسم', text='الاسم')
        self.destinations_tree.heading('الوصف', text='الوصف')
        
        # أزرار التحكم
        control_frame = ttk.Frame(self.outgoing_destinations_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="تعديل", 
                  command=self.edit_outgoing_destination).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="حذف", 
                  command=self.delete_outgoing_destination).pack(side=tk.RIGHT, padx=5)
        
        scrollbar = ttk.Scrollbar(self.outgoing_destinations_frame, orient=tk.VERTICAL, command=self.destinations_tree.yview)
        self.destinations_tree.configure(yscrollcommand=scrollbar.set)
        
        self.destinations_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_incoming_types_tab(self):
        """إعداد تبويب أنواع الوارد"""
        # إطار الإضافة
        add_frame = ttk.LabelFrame(self.incoming_types_frame, text="إضافة نوع وارد جديد", padding=10)
        add_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_frame, text="اسم النوع:").pack(side=tk.RIGHT, padx=5)
        self.new_type_name = ttk.Entry(add_frame, width=30)
        self.new_type_name.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(add_frame, text="الوصف:").pack(side=tk.RIGHT, padx=5)
        self.new_type_desc = ttk.Entry(add_frame, width=40)
        self.new_type_desc.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(add_frame, text="إضافة", 
                  command=self.add_incoming_type).pack(side=tk.RIGHT, padx=5)
        
        # جدول أنواع الوارد
        columns = ('ID', 'الاسم', 'الوصف')
        self.types_tree = ttk.Treeview(self.incoming_types_frame, columns=columns, show='headings')
        self.types_tree.column('ID', width=0, stretch=tk.NO)
        self.types_tree.heading('الاسم', text='الاسم')
        self.types_tree.heading('الوصف', text='الوصف')
        
        # أزرار التحكم
        control_frame = ttk.Frame(self.incoming_types_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="تعديل", 
                  command=self.edit_incoming_type).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="حذف", 
                  command=self.delete_incoming_type).pack(side=tk.RIGHT, padx=5)
        
        scrollbar = ttk.Scrollbar(self.incoming_types_frame, orient=tk.VERTICAL, command=self.types_tree.yview)
        self.types_tree.configure(yscrollcommand=scrollbar.set)
        
        self.types_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_employees_tab(self):
        """إعداد تبويب الموظفين"""
        # إطار الإضافة
        add_frame = ttk.LabelFrame(self.employees_frame, text="إضافة موظف جديد", padding=10)
        add_frame.pack(fill=tk.X, pady=5)
        
        # الصف الأول
        row1 = ttk.Frame(add_frame)
        row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(row1, text="اسم الموظف:").pack(side=tk.RIGHT, padx=5)
        self.new_employee_name = ttk.Entry(row1, width=25)
        self.new_employee_name.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row1, text="القسم:").pack(side=tk.RIGHT, padx=5)
        self.new_employee_dept = ttk.Entry(row1, width=20)
        self.new_employee_dept.pack(side=tk.RIGHT, padx=5)
        
        # الصف الثاني
        row2 = ttk.Frame(add_frame)
        row2.pack(fill=tk.X, pady=2)
        
        ttk.Label(row2, text="المنصب:").pack(side=tk.RIGHT, padx=5)
        self.new_employee_position = ttk.Entry(row2, width=25)
        self.new_employee_position.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row2, text="الحالة:").pack(side=tk.RIGHT, padx=5)
        self.new_employee_active = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="نشط", variable=self.new_employee_active).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(add_frame, text="إضافة", 
                  command=self.add_employee).pack(side=tk.RIGHT, padx=5)
        
        # جدول الموظفين
        columns = ('ID', 'الاسم', 'القسم', 'المنصب', 'الحالة')
        self.employees_tree = ttk.Treeview(self.employees_frame, columns=columns, show='headings')
        self.employees_tree.column('ID', width=0, stretch=tk.NO)
        self.employees_tree.heading('الاسم', text='الاسم')
        self.employees_tree.heading('القسم', text='القسم')
        self.employees_tree.heading('المنصب', text='المنصب')
        self.employees_tree.heading('الحالة', text='الحالة')
        
        # أزرار التحكم
        control_frame = ttk.Frame(self.employees_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="تعديل", 
                  command=self.edit_employee).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="حذف", 
                  command=self.delete_employee).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="تفعيل/تعطيل", 
                  command=self.toggle_employee_status).pack(side=tk.RIGHT, padx=5)
        
        scrollbar = ttk.Scrollbar(self.employees_frame, orient=tk.VERTICAL, command=self.employees_tree.yview)
        self.employees_tree.configure(yscrollcommand=scrollbar.set)
        
        self.employees_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_specializations_tab(self):
        """إعداد تبويب الاختصاصات"""
        # إطار الإضافة
        add_frame = ttk.LabelFrame(self.specializations_frame, text="إضافة اختصاص جديد", padding=10)
        add_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_frame, text="اسم الاختصاص:").pack(side=tk.RIGHT, padx=5)
        self.new_specialization_name = ttk.Entry(add_frame, width=30)
        self.new_specialization_name.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(add_frame, text="الوصف:").pack(side=tk.RIGHT, padx=5)
        self.new_specialization_desc = ttk.Entry(add_frame, width=40)
        self.new_specialization_desc.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(add_frame, text="إضافة", 
                  command=self.add_specialization).pack(side=tk.RIGHT, padx=5)
        
        # جدول الاختصاصات
        columns = ('ID', 'الاسم', 'الوصف')
        self.specializations_tree = ttk.Treeview(self.specializations_frame, columns=columns, show='headings')
        self.specializations_tree.column('ID', width=0, stretch=tk.NO)
        self.specializations_tree.heading('الاسم', text='الاسم')
        self.specializations_tree.heading('الوصف', text='الوصف')
        
        # أزرار التحكم
        control_frame = ttk.Frame(self.specializations_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="تعديل", 
                  command=self.edit_specialization).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="حذف", 
                  command=self.delete_specialization).pack(side=tk.RIGHT, padx=5)
        
        scrollbar = ttk.Scrollbar(self.specializations_frame, orient=tk.VERTICAL, command=self.specializations_tree.yview)
        self.specializations_tree.configure(yscrollcommand=scrollbar.set)
        
        self.specializations_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def load_data(self):
        """تحميل جميع البيانات"""
        self.load_incoming_sources()
        self.load_outgoing_destinations()
        self.load_incoming_types()
        self.load_employees()
        self.load_specializations()
    
    def load_incoming_sources(self):
        """تحميل جهات الوارد"""
        sources = self.db_manager.execute_query("SELECT id, name, description FROM incoming_sources ORDER BY name")
        self.sources_tree.delete(*self.sources_tree.get_children())
        for source in sources:
            self.sources_tree.insert('', tk.END, values=source)
    
    def load_outgoing_destinations(self):
        """تحميل جهات الصادر"""
        destinations = self.db_manager.execute_query("SELECT id, name, description FROM outgoing_destinations ORDER BY name")
        self.destinations_tree.delete(*self.destinations_tree.get_children())
        for destination in destinations:
            self.destinations_tree.insert('', tk.END, values=destination)
    
    def load_incoming_types(self):
        """تحميل أنواع الوارد"""
        types = self.db_manager.execute_query("SELECT id, name, description FROM incoming_types ORDER BY name")
        self.types_tree.delete(*self.types_tree.get_children())
        for type_item in types:
            self.types_tree.insert('', tk.END, values=type_item)
    
    def load_employees(self):
        """تحميل الموظفين"""
        employees = self.db_manager.execute_query("SELECT id, name, department, position, CASE WHEN is_active THEN 'نشط' ELSE 'غير نشط' END FROM employees ORDER BY name")
        self.employees_tree.delete(*self.employees_tree.get_children())
        for employee in employees:
            self.employees_tree.insert('', tk.END, values=employee)
    
    def load_specializations(self):
        """تحميل الاختصاصات"""
        specializations = self.db_manager.execute_query("SELECT id, name, description FROM specializations ORDER BY name")
        self.specializations_tree.delete(*self.specializations_tree.get_children())
        for specialization in specializations:
            self.specializations_tree.insert('', tk.END, values=specialization)
    
    # دوال الإضافة
    def add_incoming_source(self):
        name = self.new_source_name.get().strip()
        desc = self.new_source_desc.get().strip()
        
        if not name:
            messagebox.showwarning("تحذير", "يرجى إدخال اسم جهة الوارد")
            return
        
        try:
            self.db_manager.execute_query(
                "INSERT INTO incoming_sources (name, description) VALUES (?, ?)",
                (name, desc)
            )
            self.new_source_name.delete(0, tk.END)
            self.new_source_desc.delete(0, tk.END)
            self.load_incoming_sources()
            messagebox.showinfo("نجاح", "تم إضافة جهة الوارد بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الإضافة: {e}")
    
    def add_outgoing_destination(self):
        name = self.new_destination_name.get().strip()
        desc = self.new_destination_desc.get().strip()
        
        if not name:
            messagebox.showwarning("تحذير", "يرجى إدخال اسم جهة الصادر")
            return
        
        try:
            self.db_manager.execute_query(
                "INSERT INTO outgoing_destinations (name, description) VALUES (?, ?)",
                (name, desc)
            )
            self.new_destination_name.delete(0, tk.END)
            self.new_destination_desc.delete(0, tk.END)
            self.load_outgoing_destinations()
            messagebox.showinfo("نجاح", "تم إضافة جهة الصادر بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الإضافة: {e}")
    
    def add_incoming_type(self):
        name = self.new_type_name.get().strip()
        desc = self.new_type_desc.get().strip()
        
        if not name:
            messagebox.showwarning("تحذير", "يرجى إدخال اسم النوع")
            return
        
        try:
            self.db_manager.execute_query(
                "INSERT INTO incoming_types (name, description) VALUES (?, ?)",
                (name, desc)
            )
            self.new_type_name.delete(0, tk.END)
            self.new_type_desc.delete(0, tk.END)
            self.load_incoming_types()
            messagebox.showinfo("نجاح", "تم إضافة نوع الوارد بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الإضافة: {e}")
    
    def add_employee(self):
        name = self.new_employee_name.get().strip()
        dept = self.new_employee_dept.get().strip()
        position = self.new_employee_position.get().strip()
        active = self.new_employee_active.get()
        
        if not name:
            messagebox.showwarning("تحذير", "يرجى إدخال اسم الموظف")
            return
        
        try:
            self.db_manager.execute_query(
                "INSERT INTO employees (name, department, position, is_active) VALUES (?, ?, ?, ?)",
                (name, dept, position, 1 if active else 0)
            )
            self.new_employee_name.delete(0, tk.END)
            self.new_employee_dept.delete(0, tk.END)
            self.new_employee_position.delete(0, tk.END)
            self.load_employees()
            messagebox.showinfo("نجاح", "تم إضافة الموظف بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الإضافة: {e}")
    
    def add_specialization(self):
        name = self.new_specialization_name.get().strip()
        desc = self.new_specialization_desc.get().strip()
        
        if not name:
            messagebox.showwarning("تحذير", "يرجى إدخال اسم الاختصاص")
            return
        
        try:
            self.db_manager.execute_query(
                "INSERT INTO specializations (name, description) VALUES (?, ?)",
                (name, desc)
            )
            self.new_specialization_name.delete(0, tk.END)
            self.new_specialization_desc.delete(0, tk.END)
            self.load_specializations()
            messagebox.showinfo("نجاح", "تم إضافة الاختصاص بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الإضافة: {e}")
    
    # دوال التعديل (سيتم تطويرها بشكل كامل)
    def edit_incoming_source(self):
        selected = self.sources_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار جهة وارد للتعديل")
            return
        # سيتم تطوير نافذة التعديل
        messagebox.showinfo("معلومة", "سيتم تطوير خاصية التعديل قريباً")
    
    def edit_outgoing_destination(self):
        selected = self.destinations_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار جهة صادر للتعديل")
            return
        messagebox.showinfo("معلومة", "سيتم تطوير خاصية التعديل قريباً")
    
    def edit_incoming_type(self):
        selected = self.types_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار نوع وارد للتعديل")
            return
        messagebox.showinfo("معلومة", "سيتم تطوير خاصية التعديل قريباً")
    
    def edit_employee(self):
        selected = self.employees_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف للتعديل")
            return
        messagebox.showinfo("معلومة", "سيتم تطوير خاصية التعديل قريباً")
    
    def edit_specialization(self):
        selected = self.specializations_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار اختصاص للتعديل")
            return
        messagebox.showinfo("معلومة", "سيتم تطوير خاصية التعديل قريباً")
    
    # دوال الحذف
    def delete_incoming_source(self):
        selected = self.sources_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار جهة وارد للحذف")
            return
        
        item = self.sources_tree.item(selected[0])
        source_id = item['values'][0]
        source_name = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف جهة الوارد '{source_name}'؟"):
            try:
                # التحقق من عدم وجود سجلات مرتبطة
                related_records = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM incoming_records WHERE incoming_source_id = ?",
                    (source_id,)
                )[0][0]
                
                if related_records > 0:
                    messagebox.showerror("خطأ", f"لا يمكن حذف جهة الوارد لأنها مرتبطة بـ {related_records} سجل")
                    return
                
                self.db_manager.execute_query(
                    "DELETE FROM incoming_sources WHERE id = ?",
                    (source_id,)
                )
                self.load_incoming_sources()
                messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في الحذف: {e}")
    
    def delete_outgoing_destination(self):
        selected = self.destinations_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار جهة صادر للحذف")
            return
        
        item = self.destinations_tree.item(selected[0])
        destination_id = item['values'][0]
        destination_name = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف جهة الصادر '{destination_name}'؟"):
            try:
                # التحقق من عدم وجود سجلات مرتبطة
                related_records = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM outgoing_records WHERE outgoing_destination_id = ?",
                    (destination_id,)
                )[0][0]
                
                if related_records > 0:
                    messagebox.showerror("خطأ", f"لا يمكن حذف جهة الصادر لأنها مرتبطة بـ {related_records} سجل")
                    return
                
                self.db_manager.execute_query(
                    "DELETE FROM outgoing_destinations WHERE id = ?",
                    (destination_id,)
                )
                self.load_outgoing_destinations()
                messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في الحذف: {e}")
    
    def delete_incoming_type(self):
        selected = self.types_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار نوع وارد للحذف")
            return
        
        item = self.types_tree.item(selected[0])
        type_id = item['values'][0]
        type_name = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف نوع الوارد '{type_name}'؟"):
            try:
                # التحقق من عدم وجود سجلات مرتبطة
                related_records = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM incoming_records WHERE incoming_type_id = ?",
                    (type_id,)
                )[0][0]
                
                if related_records > 0:
                    messagebox.showerror("خطأ", f"لا يمكن حذف نوع الوارد لأنه مرتبط بـ {related_records} سجل")
                    return
                
                self.db_manager.execute_query(
                    "DELETE FROM incoming_types WHERE id = ?",
                    (type_id,)
                )
                self.load_incoming_types()
                messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في الحذف: {e}")
    
    def delete_employee(self):
        selected = self.employees_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف للحذف")
            return
        
        item = self.employees_tree.item(selected[0])
        employee_id = item['values'][0]
        employee_name = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف الموظف '{employee_name}'؟"):
            try:
                # التحقق من عدم وجود سجلات مرتبطة
                related_incoming = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM incoming_records WHERE employee_id = ?",
                    (employee_id,)
                )[0][0]
                
                related_outgoing = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM outgoing_records WHERE employee_id = ?",
                    (employee_id,)
                )[0][0]
                
                if related_incoming > 0 or related_outgoing > 0:
                    messagebox.showerror("خطأ", f"لا يمكن حذف الموظف لأنه مرتبط بـ {related_incoming + related_outgoing} سجل")
                    return
                
                self.db_manager.execute_query(
                    "DELETE FROM employees WHERE id = ?",
                    (employee_id,)
                )
                self.load_employees()
                messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في الحذف: {e}")
    
    def delete_specialization(self):
        selected = self.specializations_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار اختصاص للحذف")
            return
        
        item = self.specializations_tree.item(selected[0])
        specialization_id = item['values'][0]
        specialization_name = item['values'][1]
        
        if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف الاختصاص '{specialization_name}'؟"):
            try:
                # التحقق من عدم وجود سجلات مرتبطة
                related_incoming = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM incoming_records WHERE specialization_id = ?",
                    (specialization_id,)
                )[0][0]
                
                related_outgoing = self.db_manager.execute_query(
                    "SELECT COUNT(*) FROM outgoing_records WHERE specialization_id = ?",
                    (specialization_id,)
                )[0][0]
                
                if related_incoming > 0 or related_outgoing > 0:
                    messagebox.showerror("خطأ", f"لا يمكن حذف الاختصاص لأنه مرتبط بـ {related_incoming + related_outgoing} سجل")
                    return
                
                self.db_manager.execute_query(
                    "DELETE FROM specializations WHERE id = ?",
                    (specialization_id,)
                )
                self.load_specializations()
                messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في الحذف: {e}")
    
    def toggle_employee_status(self):
        """تفعيل/تعطيل حالة الموظف"""
        selected = self.employees_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف")
            return
        
        item = self.employees_tree.item(selected[0])
        employee_id = item['values'][0]
        employee_name = item['values'][1]
        current_status = item['values'][4]
        
        new_status = 0 if current_status == 'نشط' else 1
        
        try:
            self.db_manager.execute_query(
                "UPDATE employees SET is_active = ? WHERE id = ?",
                (new_status, employee_id)
            )
            self.load_employees()
            status_text = "تم التفعيل" if new_status else "تم التعطيل"
            messagebox.showinfo("نجاح", f"{status_text} للموظف {employee_name}")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في تغيير الحالة: {e}")