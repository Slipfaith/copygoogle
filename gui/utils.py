from functools import wraps
from PySide6.QtWidgets import QMessageBox


def handle_errors(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as e:
            # Log and show error dialog
            if hasattr(self, 'log_message'):
                self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(self, "Ошибка", str(e))
    return wrapper
