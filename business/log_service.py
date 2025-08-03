from datetime import datetime
from pathlib import Path


class LogService:
    """Service responsible for file based logging."""

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.log_file = None
        self.log_file_path = None

    def open(self, header_lines):
        logs_dir = self.base_dir / "logs"
        logs_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file_path = logs_dir / f"log_{timestamp}.txt"
        self.log_file = open(self.log_file_path, "w", encoding="utf-8")
        for line in header_lines:
            self.log_file.write(line + "\n")
        self.log_file.write("\n")

    def close(self):
        if self.log_file:
            self.log_file.close()
            self.log_file = None

    def log(self, message: str) -> str:
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {message}"
        if self.log_file:
            self.log_file.write(formatted + "\n")
            self.log_file.flush()
        return formatted
