import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os

class ReportsWindow:
    def __init__(self, parent, db_manager, export_manager, report_type):
        self.parent = parent
        self.db_manager = db_manager
        self.export_manager = export_manager
        self.report_type = report_type  # 'incoming' or 'outgoing'
        
        self.setup_window()
        self.create_widgets()
        self.generate_report()
    
    def setup_window(self):
        """إعداد نافذة التقارير"""
        title = "تقرير الوارد" if self.report_type == 'incoming' else "تقرير الصادر"
        self.parent.title(title)
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
        
        # إطار النتائج
        results_frame = ttk.LabelFrame(self.main_frame, text="نتائج التقرير", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.setup_results_section(results_frame)
        
        # أزرار التحكم
        self.create_control_buttons()
    
    def create_report_criteria(self, parent):
        """إنشاء معايير التقرير"""
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
        
        # معايير التجميع
        group_frame = ttk.Frame(parent)
        group_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(group_frame, text="التجميع حسب:").pack(side=tk.RIGHT, padx=5)
        self.group_by = ttk.Combobox(group_frame, width=15, state="readonly")
        
        if self.report_type == 'incoming':
            self.group_by['values'] = ['لا يوجد', 'جهة الوارد', 'نوع الوارد', 'الموظف', 'الاختصاص', 'الشهر']
        else:
            self.group_by['values'] = ['لا يوجد', 'جهة الصادر', 'الموظف', 'الاختصاص', 'الشهر']
        
        self.group_by.set('لا يوجد')
        self.group_by.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(group_frame, text="توليد التقرير", 
                  command=self.generate_report).pack(side=tk.RIGHT, padx=5)
    
    def setup_results_section(self, parent):
        """إعداد قسم النتائج"""
        # أزرار التحكم بالنتائج
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="تصدير إلى Excel", 
                  command=self.export_to_excel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="تصدير إلى PDF", 
                  command=self.export_to_pdf).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="تصدير إلى Word", 
                  command=self.export_to_word).pack(side=tk.RIGHT, padx=5)
        
        # جدول النتائج
        if self.report_type == 'incoming':
            columns = ['رقم السجل', 'رقم الوارد', 'الرقم التسلسلي', 'العنوان', 
                      'جهة الوارد', 'النوع', 'الموظف', 'الاختصاص', 'التاريخ']
        else:
            columns = ['رقم السجل', 'رقم الصادر', 'الرقم التسلسلي', 'العنوان', 
                      'جهة الصادر', 'الموظف', 'الاختصاص', 'التاريخ']
        
        self.results_tree = ttk.Treeview(parent, columns=columns, show='headings')
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_control_buttons(self):
        """إنشاء أزرار التحكم"""
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="إغلاق", 
                  command=self.parent.destroy).pack(side=tk.RIGHT, padx=5)
    
    def generate_report(self):
        """توليد التقرير"""
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        group_by = self.group_by.get()
        
        try:
            if self.report_type == 'incoming':
                data, columns = self.export_manager.generate_incoming_report(start_date, end_date)
            else:
                data, columns = self.export_manager.generate_outgoing_report(start_date, end_date)
            
            # تطبيق التجميع إذا تم اختياره
            if group_by != 'لا يوجد':
                data = self.apply_grouping(data, columns, group_by)
                columns = self.get_grouped_columns(group_by)
            
            # عرض البيانات
            self.display_data(data, columns)
            
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في توليد التقرير: {e}")
    
    def apply_grouping(self, data, columns, group_by):
        """تطبيق التجميع على البيانات"""
        grouped_data = {}
        
        for row in data:
            if group_by == 'جهة الوارد':
                key = row[4]  # جهة الوارد
            elif group_by == 'نوع الوارد':
                key = row[5]  # نوع الوارد
            elif group_by == 'الموظف':
                key = row[6]  # الموظف
            elif group_by == 'الاختصاص':
                key = row[7]  # الاختصاص
            elif group_by == 'جهة الصادر':
                key = row[4]  # جهة الصادر
            elif group_by == 'الشهر':
                key = row[-1][:7]  # السنة-الشهر
            else:
                key = 'غير محدد'
            
            if key not in grouped_data:
                grouped_data[key] = 0
            grouped_data[key] += 1
        
        # تحويل إلى قائمة من الصفوف
        return [(key, count) for key, count in grouped_data.items()]
    
    def get_grouped_columns(self, group_by):
        """الحصول على أعمدة التقرير المجمع"""
        if group_by == 'الشهر':
            return ['الفترة', 'عدد السجلات']
        else:
            return [group_by, 'عدد السجلات']
    
    def display_data(self, data, columns):
        """عرض البيانات في الجدول"""
        # مسح البيانات القديمة
        self.results_tree.delete(*self.results_tree.get_children())
        
        # تحديث الأعمدة
        self.update_tree_columns(columns)
        
        # إضافة البيانات الجديدة
        for row in data:
            self.results_tree.insert('', tk.END, values=row)
    
    def update_tree_columns(self, columns):
        """تحديث أعمدة الجدول"""
        # مسح الأعمدة الحالية
        for col in self.results_tree['columns']:
            self.results_tree.heading(col, text="")
        
        # تعيين الأعمدة الجديدة
        self.results_tree['columns'] = columns
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=120)
    
    def export_to_excel(self):
        """تصدير إلى Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="حفظ التقرير كملف Excel"
        )
        
        if file_path:
            try:
                # جمع البيانات من الجدول
                data = []
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item)['values']
                    data.append(values)
                
                # الحصول على الأعمدة
                columns = self.results_tree['columns']
                
                if self.export_manager.export_to_excel(data, columns, file_path, "تقرير"):
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
            title="حفظ التقرير كملف PDF"
        )
        
        if file_path:
            try:
                # جمع البيانات من الجدول
                data = []
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item)['values']
                    data.append(values)
                
                # الحصول على الأعمدة
                columns = self.results_tree['columns']
                
                title = "تقرير الوارد" if self.report_type == 'incoming' else "تقرير الصادر"
                
                if self.export_manager.export_to_pdf(data, columns, file_path, title):
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير إلى PDF")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")
    
    def export_to_word(self):
        """تصدير إلى Word"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
            title="حفظ التقرير كملف Word"
        )
        
        if file_path:
            try:
                # جمع البيانات من الجدول
                data = []
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item)['values']
                    data.append(values)
                
                # الحصول على الأعمدة
                columns = self.results_tree['columns']
                
                title = "تقرير الوارد" if self.report_type == 'incoming' else "تقرير الصادر"
                
                if self.export_manager.export_to_word(data, columns, file_path, title):
                    messagebox.showinfo("نجاح", f"تم التصدير إلى: {file_path}")
                else:
                    messagebox.showerror("خطأ", "فشل في التصدير إلى Word")
                    
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في التصدير: {e}")