import re
from datetime import datetime

class Validators:
    @staticmethod
    def validate_required(value, field_name):
        """التحقق من الحقول الإلزامية"""
        if not value or str(value).strip() == "":
            return f"حقل {field_name} مطلوب"
        return None
    
    @staticmethod
    def validate_number(value, field_name):
        """التحقق من الأرقام"""
        if value and not str(value).isdigit():
            return f"حقل {field_name} يجب أن يكون رقماً"
        return None
    
    @staticmethod
    def validate_date(value, field_name):
        """التحقق من تاريخ صحيح"""
        try:
            if value:
                datetime.strptime(value, '%Y-%m-%d')
            return None
        except ValueError:
            return f"حقل {field_name} يجب أن يكون تاريخاً صحيحاً (YYYY-MM-DD)"
    
    @staticmethod
    def validate_email(value, field_name):
        """التحقق من البريد الإلكتروني"""
        if value:
            pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            if not re.match(pattern, value):
                return f"حقل {field_name} يجب أن يكون بريداً إلكترونياً صحيحاً"
        return None
    
    @staticmethod
    def validate_phone(value, field_name):
        """التحقق من رقم الهاتف"""
        if value:
            # نمط بسيط لرقم الهاتف
            pattern = r'^[\+]?[0-9\s\-\(\)]{8,}$'
            if not re.match(pattern, value):
                return f"حقل {field_name} يجب أن يكون رقم هاتف صحيح"
        return None