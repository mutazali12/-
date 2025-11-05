import os
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import sys

class PrinterManager:
    def __init__(self, db_manager, export_manager):
        self.db_manager = db_manager
        self.export_manager = export_manager
    
    def print_treeview_data(self, treeview, title="طباعة البيانات"):
        """طباعة بيانات من Treeview"""
        try:
            # جمع البيانات من Treeview
            data = []
            columns = []
            
            # الحصول على الأعمدة المرئية فقط
            for col in treeview['columns']:
                if treeview.column(col).get('width', 0) > 0:  # استبعاد الأعمدة المخفية
                    columns.append(treeview.heading(col)['text'])
            
            # الحصول على البيانات
            for item in treeview.get_children():
                values = treeview.item(item)['values']
                if values:  # التأكد من وجود قيم
                    data.append(values)
            
            if not data:
                messagebox.showwarning("تحذير", "لا توجد بيانات للطباعة")
                return False
            
            # إنشاء ملف PDF مؤقت وطباعته
            return self.print_to_pdf(data, columns, title)
            
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الطباعة: {e}")
            return False
    
    def print_to_pdf(self, data, columns, title="تقرير"):
        """طباعة البيانات إلى PDF ثم إرسالها للطابعة"""
        try:
            # إنشاء ملف PDF مؤقت
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
                temp_pdf_path = temp_file.name
            
            # تصدير البيانات إلى PDF
            success = self.export_manager.export_to_pdf(data, columns, temp_pdf_path, title)
            
            if success:
                # محاولة فتح PDF للطباعة
                self.open_file_for_printing(temp_pdf_path)
                return True
            else:
                messagebox.showerror("خطأ", "فشل في إنشاء ملف الطباعة")
                return False
                
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الطباعة: {e}")
            return False
    
    def open_file_for_printing(self, file_path):
        """فتح الملف للطباعة باستخدام التطبيق الافتراضي"""
        try:
            import subprocess
            import os
            import sys
            
            if sys.platform == "win32":
                # نظام Windows - محاولة الطباعة المباشرة
                try:
                    os.startfile(file_path, "print")
                    messagebox.showinfo("نجاح", "تم إرسال الملف للطابعة. يرجى التأكد من أن الطابعة جاهزة.")
                    return
                except:
                    # إذا فشلت الطباعة المباشرة، فتح الملف
                    self.open_file_with_message(file_path)
                    
            elif sys.platform == "darwin":
                # نظام macOS
                try:
                    subprocess.run(["lpr", file_path], check=True)
                    messagebox.showinfo("نجاح", "تم إرسال الملف للطابعة.")
                except:
                    self.open_file_with_message(file_path)
            else:
                # نظام Linux
                try:
                    subprocess.run(["lp", file_path], check=True)
                    messagebox.showinfo("نجاح", "تم إرسال الملف للطابعة.")
                except:
                    self.open_file_with_message(file_path)
            
        except Exception as e:
            # إذا فشلت جميع المحاولات، فتح الملف للطباعة يدوياً
            self.open_file_with_message(file_path)
    
    def open_file_with_message(self, file_path):
        """فتح الملف مع رسالة توضيحية"""
        try:
            self.open_file(file_path)
            messagebox.showinfo("طباعة", 
                              f"تم فتح ملف PDF للطباعة.\n"
                              f"يرجى استخدام أمر الطباعة (Ctrl+P) من التطبيق الذي فتح الملف.\n\n"
                              f"مسار الملف: {file_path}")
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر فتح الملف: {e}")
    
    def open_file(self, file_path):
        """فتح الملف باستخدام التطبيق الافتراضي"""
        try:
            import os
            import subprocess
            import sys
            
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", file_path])
            else:
                subprocess.run(["xdg-open", file_path])
                
        except Exception as e:
            raise Exception(f"تعذر فتح الملف: {e}")
    
    def print_incoming_records(self, start_date=None, end_date=None):
        """طباعة سجلات الوارد"""
        try:
            data, columns = self.export_manager.generate_incoming_report(start_date, end_date)
            if not data:
                messagebox.showwarning("تحذير", "لا توجد بيانات للطباعة")
                return False
                
            title = f"سجلات الوارد - {datetime.now().strftime('%Y-%m-%d')}"
            return self.print_to_pdf(data, columns, title)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة سجلات الوارد: {e}")
            return False
    
    def print_outgoing_records(self, start_date=None, end_date=None):
        """طباعة سجلات الصادر"""
        try:
            data, columns = self.export_manager.generate_outgoing_report(start_date, end_date)
            if not data:
                messagebox.showwarning("تحذير", "لا توجد بيانات للطباعة")
                return False
                
            title = f"سجلات الصادر - {datetime.now().strftime('%Y-%m-%d')}"
            return self.print_to_pdf(data, columns, title)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة سجلات الصادر: {e}")
            return False
    
    def print_employee_report(self, start_date=None, end_date=None, department=None):
        """طباعة تقرير الموظفين"""
        try:
            data, columns, _, _ = self.export_manager.generate_employee_report(
                start_date, end_date, department
            )
            if not data:
                messagebox.showwarning("تحذير", "لا توجد بيانات للطباعة")
                return False
                
            title = f"تقرير الموظفين - {datetime.now().strftime('%Y-%m-%d')}"
            return self.print_to_pdf(data, columns, title)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة تقرير الموظفين: {e}")
            return False
    
    def quick_print_current_view(self, treeview, report_type):
        """طباعة سريعة للعرض الحالي"""
        titles = {
            'incoming': 'سجلات الوارد',
            'outgoing': 'سجلات الصادر', 
            'search': 'نتائج البحث',
            'employees': 'تقرير الموظفين'
        }
        
        title = f"{titles.get(report_type, 'تقرير')} - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        return self.print_treeview_data(treeview, title)
    
    def print_selected_record(self, treeview, record_type):
        """طباعة سجل محدد"""
        try:
            selected = treeview.selection()
            if not selected:
                messagebox.showwarning("تحذير", "يرجى اختيار سجل للطباعة")
                return False
            
            # جمع بيانات السجل المحدد
            data = []
            columns = []
            
            # الحصول على الأعمدة المرئية
            for col in treeview['columns']:
                if treeview.column(col).get('width', 0) > 0:
                    columns.append(treeview.heading(col)['text'])
            
            # الحصول على بيانات السجل المحدد
            for item in selected:
                values = treeview.item(item)['values']
                if values:
                    data.append(values)
            
            if not data:
                messagebox.showwarning("تحذير", "لا توجد بيانات للطباعة")
                return False
            
            title = f"{record_type} - سجل محدد - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            return self.print_to_pdf(data, columns, title)
            
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في طباعة السجل المحدد: {e}")
            return False