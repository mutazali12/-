import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

class FileManager:
    def __init__(self, attachments_dir="data/attachments"):
        self.attachments_dir = attachments_dir
        os.makedirs(attachments_dir, exist_ok=True)
    
    def select_file(self, title="اختر ملف"):
        """فتح نافذة لاختيار ملف"""
        root = tk.Tk()
        root.withdraw()  # إخفاء النافذة الرئيسية
        
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[
                ("جميع الملفات", "*.*"),
                ("ملفات PDF", "*.pdf"),
                ("ملفات Word", "*.docx"),
                ("ملفات Excel", "*.xlsx"),
                ("الصور", "*.jpg *.jpeg *.png *.gif")
            ]
        )
        
        root.destroy()
        return file_path
    
    def save_attachment(self, source_path, record_id, record_type):
        """حفظ المرفق في المجلد المخصص"""
        if not source_path or not os.path.exists(source_path):
            return None
        
        # إنشاء اسم فريد للملف
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_name = os.path.basename(source_path)
        name, ext = os.path.splitext(original_name)
        new_filename = f"{record_type}_{record_id}_{timestamp}{ext}"
        
        # المسار الجديد
        new_path = os.path.join(self.attachments_dir, new_filename)
        
        try:
            shutil.copy2(source_path, new_path)
            return {
                'file_name': original_name,
                'file_path': new_path,
                'saved_name': new_filename
            }
        except Exception as e:
            print(f"خطأ في حفظ الملف: {e}")
            return None
    
    def delete_attachment(self, file_path):
        """حذف ملف مرفق"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
        except Exception as e:
            print(f"خطأ في حذف الملف: {e}")
        return False
    
    def open_attachment(self, file_path):
        """فتح المرفق باستخدام التطبيق الافتراضي"""
        try:
            if os.path.exists(file_path):
                os.startfile(file_path)  # يعمل على Windows
                return True
        except Exception as e:
            print(f"خطأ في فتح الملف: {e}")
        return False
    
    def get_attachment_size(self, file_path):
        """الحصول على حجم الملف"""
        try:
            if os.path.exists(file_path):
                size = os.path.getsize(file_path)
                return self._format_size(size)
        except:
            pass
        return "غير معروف"
    
    def _format_size(self, size_bytes):
        """تنسيق حجم الملف"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"