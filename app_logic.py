"""Application business logic separated from GUI."""

from typing import Dict, List, Optional, Callable

from PySide6.QtCore import QThread, Signal

from business.processor import ExcelToGoogleSheets
from config import BASE_DIR


class WorkerThread(QThread):
    """Background worker for processing tasks."""

    progress_update = Signal(int, int, str)
    log_message = Signal(str)
    finished_successfully = Signal()
    error_occurred = Signal(str)

    def __init__(self, mode: str, **kwargs):
        super().__init__()
        self.mode = mode
        self.kwargs = kwargs
        self.processor: Optional[ExcelToGoogleSheets] = None

    def run(self):
        try:
            self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))

            if self.mode == "single":
                excel_path = self.kwargs['excel_path']
                google_sheet_url = self.kwargs['google_sheet_url']
                config = self.kwargs['config']

                self.processor.update_config(
                    sheet_mapping=config['sheet_mapping'],
                    column_mapping=config['column_mapping'],
                    start_row=config['start_row']
                )

                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)

                self.processor.process_excel_file(
                    excel_path,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )

            elif self.mode == "batch":
                file_mappings = self.kwargs['file_mappings']
                google_sheet_url = self.kwargs['google_sheet_url']

                self.processor.process_multiple_excel_files(
                    file_mappings,
                    google_sheet_url,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )

            elif self.mode == "download":
                google_sheet_url = self.kwargs['google_sheet_url']
                save_path = self.kwargs['save_path']
                sheet_names = self.kwargs.get('sheet_names')

                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)

                self.processor.download_google_sheet(
                    save_path,
                    sheet_names=sheet_names,
                    log_callback=self.log_message.emit
                )

            self.finished_successfully.emit()

        except Exception as e:  # pragma: no cover - defensive
            self.error_occurred.emit(str(e))


class AppLogic:
    """Facade for business logic used by GUI."""

    def __init__(self) -> None:
        self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))
        self.worker_thread: Optional[WorkerThread] = None

    # Data retrieval helpers
    def get_excel_sheets(self, excel_path: str) -> List[str]:
        return self.processor.get_excel_sheets(excel_path)

    def get_google_sheets(self, google_url: str) -> List[str]:
        self.processor.connect_to_google_sheets(google_url)
        return self.processor.get_google_sheets()

    def get_google_sheet_title(self) -> str:
        return self.processor.google_sheet.title if self.processor.google_sheet else ""

    # Processing starters
    def start_single_processing(
        self,
        excel_path: str,
        google_url: str,
        config: Dict,
        progress_cb: Callable[[int, int, str], None],
        log_cb: Callable[[str], None],
        finished_cb: Callable[[], None],
        error_cb: Callable[[str], None],
    ) -> None:
        self.worker_thread = WorkerThread(
            mode="single",
            excel_path=excel_path,
            google_sheet_url=google_url,
            config=config,
        )
        self._connect_worker_signals(progress_cb, log_cb, finished_cb, error_cb)
        self.worker_thread.start()

    def start_batch_processing(
        self,
        file_mappings: List[Dict[str, str]],
        google_url: str,
        progress_cb: Callable[[int, int, str], None],
        log_cb: Callable[[str], None],
        finished_cb: Callable[[], None],
        error_cb: Callable[[str], None],
    ) -> None:
        self.worker_thread = WorkerThread(
            mode="batch",
            file_mappings=file_mappings,
            google_sheet_url=google_url,
        )
        self._connect_worker_signals(progress_cb, log_cb, finished_cb, error_cb)
        self.worker_thread.start()

    def start_download(
        self,
        google_url: str,
        save_path: str,
        sheet_names: Optional[List[str]],
        progress_cb: Callable[[int, int, str], None],
        log_cb: Callable[[str], None],
        finished_cb: Callable[[], None],
        error_cb: Callable[[str], None],
    ) -> None:
        self.worker_thread = WorkerThread(
            mode="download",
            google_sheet_url=google_url,
            save_path=save_path,
            sheet_names=sheet_names,
        )
        self._connect_worker_signals(progress_cb, log_cb, finished_cb, error_cb)
        self.worker_thread.start()

    # Internal helper
    def _connect_worker_signals(
        self,
        progress_cb: Callable[[int, int, str], None],
        log_cb: Callable[[str], None],
        finished_cb: Callable[[], None],
        error_cb: Callable[[str], None],
    ) -> None:
        if not self.worker_thread:
            return
        self.worker_thread.progress_update.connect(progress_cb)
        self.worker_thread.log_message.connect(log_cb)
        self.worker_thread.finished_successfully.connect(finished_cb)
        self.worker_thread.error_occurred.connect(error_cb)

