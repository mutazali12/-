import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os

class OutgoingForm:
    def __init__(self, parent, db_manager, file_manager, record_id=None):
        self.parent = parent
        self.db_manager = db_manager
        self.file_manager = file_manager
        self.record_id = record_id
        self.attachments = []
        
        self.setup_window()
        self.create_widgets()
        self.load_reference_data()
        
        if record_id:
            self.load_record_data()
        else:
            self.generate_record_numbers()
    
    def setup_window(self):
        """إعداد نافذة النموذج"""
        self.parent.title("نموذج تسجيل صادر" if not self.record_id else "تعديل سجل صادر")
        self.parent.geometry("800x600")
        self.parent.resizable(True, True)
        
        # الإطار الرئيسي
        self.main_frame = ttk.Frame(self.parent, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # إطار بيانات السجل
        record_frame = ttk.LabelFrame(self.main_frame, text="بيانات السجل", padding=10)
        record_frame.pack(fill=tk.X, pady=5)
        
        # شبكة الحقول
        self.create_record_fields(record_frame)
        
        # إطار التفاصيل
        details_frame = ttk.LabelFrame(self.main_frame, text="تفاصيل إضافية", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.details_text = tk.Text(details_frame, height=8, width=80)
        scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=scrollbar.set)
        
        self.details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # إطار المرفقات
        attachments_frame = ttk.LabelFrame(self.main_frame, text="المرفقات", padding=10)
        attachments_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.setup_attachments_section(attachments_frame)
        
        # أزرار التحكم
        self.create_control_buttons()
    
    def create_record_fields(self, parent):
        """إنشاء حقول بيانات السجل"""
        # الصف الأول
        row1 = ttk.Frame(parent)
        row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(row1, text="رقم السجل:").pack(side=tk.RIGHT, padx=5)
        self.record_number = ttk.Entry(row1, width=20)
        self.record_number.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row1, text="رقم الصادر:").pack(side=tk.RIGHT, padx=5)
        self.outgoing_number = ttk.Entry(row1, width=20)
        self.outgoing_number.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row1, text="الرقم التسلسلي:").pack(side=tk.RIGHT, padx=5)
        self.serial_number = ttk.Entry(row1, width=20)
        self.serial_number.pack(side=tk.RIGHT, padx=5)
        
        # الصف الثاني
        row2 = ttk.Frame(parent)
        row2.pack(fill=tk.X, pady=2)
        
        ttk.Label(row2, text="العنوان:").pack(side=tk.RIGHT, padx=5)
        self.title_entry = ttk.Entry(row2, width=80)
        self.title_entry.pack(side=tk.RIGHT, padx=5)
        
        # الصف الثالث
        row3 = ttk.Frame(parent)
        row3.pack(fill=tk.X, pady=2)
        
        ttk.Label(row3, text="جهة الصادر:").pack(side=tk.RIGHT, padx=5)
        self.destination_combo = ttk.Combobox(row3, width=25, state="readonly")
        self.destination_combo.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row3, text="الموظف:").pack(side=tk.RIGHT, padx=5)
        self.employee_combo = ttk.Combobox(row3, width=25, state="readonly")
        self.employee_combo.pack(side=tk.RIGHT, padx=5)
        
        # الصف الرابع
        row4 = ttk.Frame(parent)
        row4.pack(fill=tk.X, pady=2)
        
        ttk.Label(row4, text="الاختصاص:").pack(side=tk.RIGHT, padx=5)
        self.specialization_combo = ttk.Combobox(row4, width=25, state="readonly")
        self.specialization_combo.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(row4, text="تاريخ التسجيل:").pack(side=tk.RIGHT, padx=5)
        self.registration_date = ttk.Entry(row4, width=15)
        self.registration_date.insert(0, datetime.now().strftime('%Y-%m-%d'))
        self.registration_date.pack(side=tk.RIGHT, padx=5)
    
    def setup_attachments_section(self, parent):
        """إعداد قسم المرفقات"""
        # أزرار التحكم بالمرفقات
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="إضافة مرفق", 
                  command=self.add_attachment).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="حذف مرفق", 
                  command=self.delete_attachment).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="عرض مرفق", 
                  command=self.view_attachment).pack(side=tk.RIGHT, padx=5)
        
        # قائمة المرفقات
        columns = ('ID', 'اسم الملف', 'الحجم', 'الوصف')
        self.attachments_tree = ttk.Treeview(parent, columns=columns, show='headings', height=4)
        
        self.attachments_tree.column('ID', width=0, stretch=tk.NO)
        self.attachments_tree.heading('اسم الملف', text='اسم الملف')
        self.attachments_tree.heading('الحجم', text='الحجم')
        self.attachments_tree.heading('الوصف', text='الوصف')
        
        self.attachments_tree.column('اسم الملف', width=200)
        self.attachments_tree.column('الحجم', width=100)
        self.attachments_tree.column('الوصف', width=150)
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.attachments_tree.yview)
        self.attachments_tree.configure(yscrollcommand=scrollbar.set)
        
        self.attachments_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_control_buttons(self):
        """إنشاء أزرار التحكم"""
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="حفظ", 
                  command=self.save_record, style='Accent.TButton').pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="إلغاء", 
                  command=self.parent.destroy).pack(side=tk.RIGHT, padx=5)
        
        if self.record_id:
            ttk.Button(buttons_frame, text="حذف السجل", 
                      command=self.delete_record).pack(side=tk.LEFT, padx=5)
    
    def load_reference_data(self):
        """تحميل البيانات المرجعية"""
        # جهات الصادر
        destinations = self.db_manager.execute_query("SELECT id, name FROM outgoing_destinations ORDER BY name")
        self.destination_combo['values'] = [name for _, name in destinations]
        self.destinations_dict = {name: id for id, name in destinations}
        
        # الموظفين
        employees = self.db_manager.execute_query(
            "SELECT id, name FROM employees WHERE is_active = 1 ORDER BY name"
        )
        self.employee_combo['values'] = [name for _, name in employees]
        self.employees_dict = {name: id for id, name in employees}
        
        # الاختصاصات
        specializations = self.db_manager.execute_query(
            "SELECT id, name FROM specializations ORDER BY name"
        )
        self.specialization_combo['values'] = [name for _, name in specializations]
        self.specializations_dict = {name: id for id, name in specializations}
    
    def generate_record_numbers(self):
        """توليد أرقام السجلات تلقائياً"""
        # توليد رقم سجل فريد
        year = datetime.now().strftime('%Y')
        last_record = self.db_manager.execute_query(
            "SELECT record_number FROM outgoing_records WHERE record_number LIKE ? ORDER BY id DESC LIMIT 1",
            (f'OUT-{year}-%',)
        )
        
        if last_record:
            last_number = int(last_record[0][0].split('-')[-1])
            new_number = last_number + 1
        else:
            new_number = 1
        
        self.record_number.delete(0, tk.END)
        self.record_number.insert(0, f'OUT-{year}-{new_number:04d}')
        
        # توليد رقم صادر (يمكن تعديله يدوياً)
        self.outgoing_number.delete(0, tk.END)
        self.outgoing_number.insert(0, f'OUT-{datetime.now().strftime("%Y%m%d")}')
        
        # توليد رقم تسلسلي
        last_serial = self.db_manager.execute_query(
            "SELECT serial_number FROM outgoing_records ORDER BY id DESC LIMIT 1"
        )
        if last_serial:
            last_serial_num = int(last_serial[0][0])
            new_serial = last_serial_num + 1
        else:
            new_serial = 1
        
        self.serial_number.delete(0, tk.END)
        self.serial_number.insert(0, str(new_serial))
    
    def load_record_data(self):
        """تحميل بيانات السجل للتعديل"""
        query = """
        SELECT record_number, outgoing_number, serial_number, title,
               outgoing_destination_id, employee_id, specialization_id,
               registration_date, details
        FROM outgoing_records WHERE id = ?
        """
        
        record = self.db_manager.execute_query(query, (self.record_id,))
        if record:
            data = record[0]
            
            self.record_number.insert(0, data[0])
            self.outgoing_number.insert(0, data[1])
            self.serial_number.insert(0, data[2])
            self.title_entry.insert(0, data[3])
            self.registration_date.delete(0, tk.END)
            self.registration_date.insert(0, data[7])
            self.details_text.insert('1.0', data[8] if data[8] else '')
            
            # تحميل القوائم المنسدلة
            self.load_combo_values(data[4], data[5], data[6])
            
            # تحميل المرفقات
            self.load_attachments()
    
    def load_combo_values(self, destination_id, employee_id, specialization_id):
        """تحميل قيم القوائم المنسدلة بناءً على الـ IDs"""
        # جهة الصادر
        destination_name = self.db_manager.execute_query(
            "SELECT name FROM outgoing_destinations WHERE id = ?", (destination_id,)
        )
        if destination_name:
            self.destination_combo.set(destination_name[0][0])
        
        # الموظف
        employee_name = self.db_manager.execute_query(
            "SELECT name FROM employees WHERE id = ?", (employee_id,)
        )
        if employee_name:
            self.employee_combo.set(employee_name[0][0])
        
        # الاختصاص
        specialization_name = self.db_manager.execute_query(
            "SELECT name FROM specializations WHERE id = ?", (specialization_id,)
        )
        if specialization_name:
            self.specialization_combo.set(specialization_name[0][0])
    
    def load_attachments(self):
        """تحميل المرفقات"""
        attachments = self.db_manager.execute_query(
            "SELECT id, file_name, file_path, description FROM attachments WHERE record_id = ? AND record_type = 'outgoing'",
            (self.record_id,)
        )
        
        for att in attachments:
            size = self.file_manager.get_attachment_size(att[2])
            self.attachments_tree.insert('', tk.END, values=(att[0], att[1], size, att[3]))
            self.attachments.append({
                'id': att[0],
                'file_name': att[1],
                'file_path': att[2],
                'description': att[3]
            })
    
    def add_attachment(self):
        """إضافة مرفق جديد"""
        file_path = self.file_manager.select_file("اختر ملف للمرفق")
        if file_path:
            # طلب وصف للمرفق
            description = tk.simpledialog.askstring("وصف المرفق", "أدخل وصفاً للمرفق (اختياري):")
            
            if self.record_id:
                # حفظ المرفق مباشرة إذا كان سجل موجود
                saved_file = self.file_manager.save_attachment(file_path, self.record_id, 'outgoing')
                if saved_file:
                    # إدخال في قاعدة البيانات
                    query = """
                    INSERT INTO attachments (record_id, record_type, file_name, file_path, description)
                    VALUES (?, 'outgoing', ?, ?, ?)
                    """
                    attachment_id = self.db_manager.execute_query(
                        query, (self.record_id, saved_file['file_name'], saved_file['file_path'], description)
                    )
                    
                    # إضافة للعرض
                    size = self.file_manager.get_attachment_size(saved_file['file_path'])
                    self.attachments_tree.insert('', tk.END, 
                                               values=(attachment_id, saved_file['file_name'], size, description))
                    self.attachments.append({
                        'id': attachment_id,
                        'file_name': saved_file['file_name'],
                        'file_path': saved_file['file_path'],
                        'description': description
                    })
            else:
                # تخزين مؤقت حتى حفظ السجل
                size = self.file_manager.get_attachment_size(file_path)
                self.attachments_tree.insert('', tk.END, 
                                           values=(None, os.path.basename(file_path), size, description))
                self.attachments.append({
                    'file_path': file_path,
                    'description': description
                })
    
    def delete_attachment(self):
        """حذف مرفق محدد"""
        selected = self.attachments_tree.selection()
        if not selected:
            return
        
        item = self.attachments_tree.item(selected[0])
        values = item['values']
        attachment_id = values[0]
        
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذا المرفق؟"):
            if attachment_id:  # مرفق محفوظ في قاعدة البيانات
                # حذف من قاعدة البيانات
                self.db_manager.execute_query("DELETE FROM attachments WHERE id = ?", (attachment_id,))
                # حذف الملف
                for att in self.attachments:
                    if att['id'] == attachment_id:
                        self.file_manager.delete_attachment(att['file_path'])
                        break
            
            self.attachments_tree.delete(selected[0])
            # إزالة من القائمة
            self.attachments = [att for att in self.attachments if att.get('id') != attachment_id]
    
    def view_attachment(self):
        """عرض المرفق المحدد"""
        selected = self.attachments_tree.selection()
        if not selected:
            return
        
        item = self.attachments_tree.item(selected[0])
        values = item['values']
        attachment_id = values[0]
        
        for att in self.attachments:
            if att.get('id') == attachment_id:
                self.file_manager.open_attachment(att['file_path'])
                break
    
    def validate_form(self):
        """التحقق من صحة البيانات"""
        if not self.record_number.get().strip():
            messagebox.showerror("خطأ", "رقم السجل مطلوب")
            return False
        
        if not self.outgoing_number.get().strip():
            messagebox.showerror("خطأ", "رقم الصادر مطلوب")
            return False
        
        if not self.serial_number.get().strip():
            messagebox.showerror("خطأ", "الرقم التسلسلي مطلوب")
            return False
        
        if not self.title_entry.get().strip():
            messagebox.showerror("خطأ", "العنوان مطلوب")
            return False
        
        if not self.destination_combo.get():
            messagebox.showerror("خطأ", "جهة الصادر مطلوبة")
            return False
        
        return True
    
    def save_record(self):
        """حفظ السجل"""
        if not self.validate_form():
            return
        
        try:
            # تجميع البيانات
            data = (
                self.record_number.get().strip(),
                self.outgoing_number.get().strip(),
                self.serial_number.get().strip(),
                self.title_entry.get().strip(),
                self.destinations_dict[self.destination_combo.get()],
                self.employees_dict[self.employee_combo.get()],
                self.specializations_dict[self.specialization_combo.get()],
                self.registration_date.get().strip(),
                self.details_text.get('1.0', tk.END).strip()
            )
            
            if self.record_id:
                # تحديث السجل الموجود
                query = """
                UPDATE outgoing_records 
                SET record_number=?, outgoing_number=?, serial_number=?, title=?,
                    outgoing_destination_id=?, employee_id=?, specialization_id=?,
                    registration_date=?, details=?
                WHERE id=?
                """
                self.db_manager.execute_query(query, data + (self.record_id,))
                messagebox.showinfo("نجاح", "تم تحديث السجل بنجاح")
            else:
                # إدخال سجل جديد
                query = """
                INSERT INTO outgoing_records 
                (record_number, outgoing_number, serial_number, title, outgoing_destination_id,
                 employee_id, specialization_id, registration_date, details)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                new_id = self.db_manager.execute_query(query, data)
                self.record_id = new_id
                
                # حفظ المرفقات المؤقتة
                for att in self.attachments:
                    if 'file_path' in att and not att.get('id'):
                        saved_file = self.file_manager.save_attachment(
                            att['file_path'], self.record_id, 'outgoing'
                        )
                        if saved_file:
                            query = """
                            INSERT INTO attachments (record_id, record_type, file_name, file_path, description)
                            VALUES (?, 'outgoing', ?, ?, ?)
                            """
                            self.db_manager.execute_query(
                                query, (self.record_id, saved_file['file_name'], 
                                      saved_file['file_path'], att['description'])
                            )
                
                messagebox.showinfo("نجاح", "تم حفظ السجل بنجاح")
            
            self.parent.destroy()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في حفظ السجل: {e}")
    
    def delete_record(self):
        """حذف السجل"""
        if not self.record_id:
            return
        
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذا السجل وجميع مرفقاته؟"):
            try:
                # حذف المرفقات أولاً
                for att in self.attachments:
                    if att.get('id'):
                        self.file_manager.delete_attachment(att['file_path'])
                
                self.db_manager.execute_query(
                    "DELETE FROM attachments WHERE record_id = ? AND record_type = 'outgoing'",
                    (self.record_id,)
                )
                
                # حذف السجل
                self.db_manager.execute_query(
                    "DELETE FROM outgoing_records WHERE id = ?",
                    (self.record_id,)
                )
                
                messagebox.showinfo("نجاح", "تم حذف السجل بنجاح")
                self.parent.destroy()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في حذف السجل: {e}")