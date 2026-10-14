"""Students Marks Generator - polished CustomTkinter desktop application.

Keep this file in your project and retain the existing modules:
    readers/excel_reader.py
    extractors/student_extractor.py
    processors/data_validator.py
    generators/statement_generator.py

Expected template: templates/statement_template.xlsx
"""

from __future__ import annotations

import logging
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import customtkinter as ctk
from tkinter import Canvas, filedialog

from readers.excel_reader import load_workbook
from extractors.student_extractor import extract_students
from processors.data_validator import validate_students
from generators.statement_generator import generate_student_statement


# ---------------------------------------------------------------------------
# Application configuration
# ---------------------------------------------------------------------------
APP_NAME = "Students Marks Generator"
NAVY = "#10243A"
NAVY_HOVER = "#1B3855"
GREEN = "#16845B"
GREEN_HOVER = "#116746"
YELLOW = "#F2C94C"
BG = "#F3F6F9"
WHITE = "#FFFFFF"
TEXT = "#17212F"
MUTED = "#667085"
BORDER = "#DCE3EA"
RED = "#C0392B"
ORANGE = "#B96A12"


def find_project_root() -> Path:
    """Find the project folder in source and PyInstaller builds."""
    if getattr(sys, "frozen", False):
        candidates = [Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)),
                      Path(sys.executable).resolve().parent]
    else:
        here = Path(__file__).resolve().parent
        candidates = [here, here.parent]

    for candidate in candidates:
        if (candidate / "templates" / "statement_template.xlsx").is_file():
            return candidate
    return candidates[0]


PROJECT_ROOT = find_project_root()
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "statement_template.xlsx"
ICON_PATH = PROJECT_ROOT / "app_icon.ico"


def configure_logging() -> Path:
    """Configure a persistent application log in the user's documents folder."""
    log_dir = Path.home() / "Documents" / "StudentsMarksGenerator" / "logs"
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        log_dir = Path.home() / "StudentsMarksGenerator" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "students_marks_generator.log"
    logging.basicConfig(
        filename=str(log_path),
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(threadName)s | %(message)s",
        encoding="utf-8",
    )
    return log_path


LOG_PATH = configure_logging()


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------
class CircularProgress(Canvas):
    """Lightweight circular progress indicator using Tk's native Canvas."""

    def __init__(self, master: Any, size: int = 112, thickness: int = 9,
                 **kwargs: Any) -> None:
        super().__init__(master, width=size, height=size, bg=WHITE,
                         highlightthickness=0, **kwargs)
        self.size = size
        self.thickness = thickness
        self.progress = 0.0
        self._draw()

    def set(self, value: float) -> None:
        self.progress = max(0.0, min(100.0, float(value)))
        self._draw()

    def _draw(self) -> None:
        self.delete("all")
        pad = self.thickness + 4
        self.create_oval(pad, pad, self.size - pad, self.size - pad,
                         outline="#E5EAF0", width=self.thickness)
        if self.progress > 0:
            self.create_arc(pad, pad, self.size - pad, self.size - pad,
                            start=90, extent=-(self.progress / 100) * 360,
                            style="arc", outline=GREEN, width=self.thickness)
        self.create_text(self.size / 2, self.size / 2,
                         text=f"{int(self.progress)}%", fill=NAVY,
                         font=("Segoe UI", 19, "bold"))


class PremiumDialog(ctk.CTkToplevel):
    """Consistent, scrollable application dialog."""

    def __init__(self, parent: ctk.CTk, title: str, message: str,
                 dialog_type: str = "info",
                 buttons: list[tuple[str, Callable[[], None], str]] | None = None):
        super().__init__(parent)
        self.title(title)
        self.geometry("560x430")
        self.minsize(460, 340)
        self.transient(parent)
        self.configure(fg_color=WHITE)
        if ICON_PATH.is_file():
            try:
                self.iconbitmap(str(ICON_PATH))
            except Exception:
                pass

        colors = {"info": NAVY, "warning": ORANGE, "error": RED, "success": GREEN}
        accent = colors.get(dialog_type, NAVY)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 14))
        ctk.CTkFrame(header, width=5, height=36, fg_color=accent,
                     corner_radius=4).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(header, text=title, text_color=NAVY,
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", anchor="w")

        body = ctk.CTkFrame(self, fg_color="#F7F9FB", corner_radius=10,
                            border_width=1, border_color=BORDER)
        body.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 14))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)
        textbox = ctk.CTkTextbox(body, wrap="word", fg_color="#F7F9FB",
                                 text_color=TEXT, border_width=0,
                                 font=ctk.CTkFont(size=13))
        textbox.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        textbox.insert("1.0", message)
        textbox.configure(state="disabled")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 20))
        if buttons is None:
            buttons = [("OK", self.destroy, accent)]
        for label, command, color in buttons:
            ctk.CTkButton(footer, text=label, command=command, width=112,
                          height=36, fg_color=color,
                          hover_color=self._darken(color), corner_radius=8,
                          font=ctk.CTkFont(size=13, weight="bold")).pack(
                              side="right", padx=(8, 0))

        self.update_idletasks()
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.grab_set()
        self.focus_force()

    @staticmethod
    def _darken(color: str) -> str:
        try:
            color = color.lstrip("#")
            rgb = [int(color[i:i + 2], 16) for i in (0, 2, 4)]
            return "#" + "".join(f"{int(channel * 0.82):02x}" for channel in rgb)
        except (ValueError, IndexError):
            return NAVY_HOVER


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class MarksGeneratorApp(ctk.CTk):
    def __init__(self) -> None:
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("green")
        super().__init__()

        self.title(APP_NAME)
        self.geometry("800x560")
        self.minsize(760, 560)
        self.configure(fg_color=BG)
        if ICON_PATH.is_file():
            try:
                self.iconbitmap(str(ICON_PATH))
            except Exception:
                pass

        # UI variables are read on the main thread only.
        self.source_path = ctk.StringVar(value="")
        self.output_path = ctk.StringVar(
            value=str(Path.home() / "Documents" / "StudentsMarksGenerator" / "Output"))
        self.programme_var = ctk.StringVar(value="")
        self.semester_var = ctk.StringVar(value="")
        self.centre_code_var = ctk.StringVar(value="")
        self.exam_period_var = ctk.StringVar(value="")
        self.publication_date_var = ctk.StringVar(value="")
        self.issued_on_var = ctk.StringVar(value="")

        self.students_processed = 0
        self.statements_generated = 0
        self.errors_count = 0
        self.skipped_count = 0
        self.job_running = False
        self.events: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self._poll_scheduled = False

        self.programmes = [
            "BACHELOR OF ARTS IN EDUCATION WITH ENGLISH & HISTORY",
            "BACHELOR OF ARTS IN EDUCATION WITH ENGLISH & GEOGRAPHY",
            "BACHELOR OF ARTS IN EDUCATION WITH HISTORY & GEOGRAPHY",
            "BACHELOR OF BUSINESS ADMINISTRATION IN ADVERTISING & MARKETING",
            "BACHELOR OF BUSINESS ADMINISTRATION",
            "BACHELOR OF BUSINESS ADMINISTRATION IN FINANCE MANAGEMENT",
            "BACHELOR OF ENGINEERING IN COMPUTER SCIENCE",
            "BACHELOR OF SOCIAL WORK",
            "BACHELOR OF SOCIAL WORK IN PROJECT MANAGEMENT MONITORING & EVALUATION",
            "BACHELOR OF SOCIAL WORK IN COMMUNITY DEVELOPMENT",
            "BACHELOR OF COMMERCE IN ACCOUNTING & FINANCE",
            "BACHELOR OF SCIENCE IN COMPUTER SCIENCE",
            "BACHELOR OF SCIENCE EDUCATION WITH MATHEMATICS & BIOLOGY",
            "BACHELOR OF SCIENCE EDUCATION WITH MATHEMATICS & PHYSICS",
            "BACHELOR OF SCIENCE EDUCATION WITH CHEMISTRY & BIOLOGY",
            "DIPLOMA IN ARTS EDUCATION WITH ENGLISH & GEOGRAPHY",
            "DIPLOMA IN ARTS EDUCATION WITH ENGLISH & HISTORY",
            "DIPLOMA IN BUSINESS ADMINISTRATION",
            "DIPLOMA IN COMMERCE",
            "DIPLOMA IN SOCIAL WORK",
            "DIPLOMA IN COMPUTER SCIENCE",
            "DIPLOMA IN SCIENCE EDUCATION WITH MATHEMATICS & BIOLOGY",
            "DIPLOMA IN SCIENCE EDUCATION WITH MATHEMATICS & PHYSICS",
            "DIPLOMA IN SCIENCE EDUCATION WITH CHEMISTRY & BIOLOGY",
        ]
        roman_map = {1: "I", 2: "II", 3: "III", 4: "IV", 5: "V",
                     6: "VI", 7: "VII", 8: "VIII", 9: "IX", 10: "X"}
        self.semesters = [roman_map[i] for i in range(1, 11)]
        self.centre_codes = ["DMISJBU-41", "DMISJBU-42", "DMISJBU-43"]

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------- UI construction -------------------------
    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.main_frame = ctk.CTkScrollableFrame(self, fg_color=BG,
                                                 corner_radius=0,
                                                 scrollbar_button_color="#B9C5D2")
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_stats()
        self._build_file_section()
        self._build_exam_section()
        self._build_actions()
        self._build_status()
        self._build_progress_overlay()

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self.main_frame, fg_color=NAVY, corner_radius=0, height=112)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 22))
        header.grid_columnconfigure(0, weight=1)
        content = ctk.CTkFrame(header, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=22)

        if ICON_PATH.is_file():
            try:
                from PIL import Image
                image = Image.open(ICON_PATH).convert("RGBA").resize((48, 48))
                self.header_icon = ctk.CTkImage(light_image=image, dark_image=image, size=(48, 48))
                ctk.CTkLabel(content, image=self.header_icon, text="").pack(side="left", padx=(0, 15))
            except Exception:
                self._header_mark(content)
        else:
            self._header_mark(content)

        title_frame = ctk.CTkFrame(content, fg_color="transparent")
        title_frame.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(title_frame, text=APP_NAME, text_color=WHITE,
                     font=ctk.CTkFont(size=23, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Academic records automation  •  Excel to student statements",
                     text_color="#C7D5E4", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(content, text="ACADEMIC SERVICES", text_color=YELLOW,
                     font=ctk.CTkFont(size=10, weight="bold")).pack(side="right", padx=(12, 0))

    @staticmethod
    def _header_mark(parent: Any) -> None:
        mark = ctk.CTkFrame(parent, width=48, height=48, corner_radius=12, fg_color=GREEN)
        mark.pack(side="left", padx=(0, 15))
        mark.pack_propagate(False)
        ctk.CTkLabel(mark, text="SM", text_color=WHITE,
                     font=ctk.CTkFont(size=17, weight="bold")).pack(expand=True)

    def _section_title(self, row: int, title: str, subtitle: str = "") -> None:
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", padx=24, pady=(0, 10))
        ctk.CTkLabel(frame, text=title, text_color=NAVY,
                     font=ctk.CTkFont(size=17, weight="bold")).pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(frame, text=subtitle, text_color=MUTED,
                         font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(3, 0))

    def _card(self, parent: Any, **kwargs: Any) -> ctk.CTkFrame:
        return ctk.CTkFrame(parent, fg_color=WHITE, corner_radius=12,
                            border_width=1, border_color=BORDER, **kwargs)

    def _build_stats(self) -> None:
        self.stats_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.stats_frame.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 24))
        for col in range(4):
            self.stats_frame.grid_columnconfigure(col, weight=1, uniform="stats")
        self.lbl_students = self._stat_card(self.stats_frame, 0, "Students detected", "—", "")
        self.lbl_statements = self._stat_card(self.stats_frame, 1, "Generated", "—", "")
        self.lbl_skipped = self._stat_card(self.stats_frame, 2, "Skipped", "—", "")
        self.lbl_errors = self._stat_card(self.stats_frame, 3, "Failed", "—", "")

    def _stat_card(self, parent: Any, col: int, title: str, value: str, number: str) -> ctk.CTkLabel:
        card = self._card(parent)
        card.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 0 if col == 3 else 6))
        ctk.CTkLabel(card, text=number, text_color=GREEN,
                     font=ctk.CTkFont(size=10, weight="bold")).pack(anchor="w", padx=14, pady=(12, 0))
        ctk.CTkLabel(card, text=title, text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14, pady=(5, 0))
        value_label = ctk.CTkLabel(card, text=value, text_color=NAVY,
                                   font=ctk.CTkFont(size=24, weight="bold"))
        value_label.pack(anchor="w", padx=14, pady=(0, 12))
        return value_label

    def _build_file_section(self) -> None:
        self._section_title(2, "File locations", "Choose the results workbook and where generated statements should be saved.")
        card = self._card(self.main_frame)
        card.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 22))
        card.grid_columnconfigure(1, weight=1)
        self._file_row(card, 0, "Results workbook", "Select an Excel results file (.xlsx or .xlsm)",
                       self.source_path, "Browse workbook", self._browse_source, "XLS")
        ctk.CTkFrame(card, height=1, fg_color=BORDER).grid(row=1, column=0, columnspan=3, sticky="ew", padx=18)
        self._file_row(card, 2, "Output folder", "Choose a folder for the individual statements",
                       self.output_path, "Browse folder", self._browse_output, "DIR")

    def _file_row(self, parent: Any, row: int, title: str, hint: str,
                  variable: ctk.StringVar, button_text: str,
                  command: Callable[[], None], tag: str) -> None:
        ctk.CTkLabel(parent, text=tag, width=48, height=42, corner_radius=8,
                     fg_color="#EAF4EF", text_color=GREEN,
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
                         row=row, column=0, padx=(18, 12), pady=14)
        details = ctk.CTkFrame(parent, fg_color="transparent")
        details.grid(row=row, column=1, sticky="ew", pady=12)
        details.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(details, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w")
        path_label = ctk.CTkLabel(details, text=variable.get() or hint,
                                  text_color=MUTED if not variable.get() else TEXT,
                                  font=ctk.CTkFont(size=11), anchor="w")
        path_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        def update_path(*_: Any) -> None:
            value = variable.get().strip()
            shown = value if len(value) <= 75 else "…" + value[-72:]
            path_label.configure(text=shown or hint, text_color=TEXT if value else MUTED)

        variable.trace_add("write", update_path)
        ctk.CTkButton(parent, text=button_text, command=command, width=132, height=34,
                      fg_color="#EDF2F6", hover_color="#DDE6ED", text_color=NAVY,
                      corner_radius=8, font=ctk.CTkFont(size=12, weight="bold")).grid(
                          row=row, column=2, padx=18, pady=14)

    def _build_exam_section(self) -> None:
        self._section_title(4, "Examination details", "These details are printed on the generated statements.")
        card = self._card(self.main_frame)
        card.grid(row=5, column=0, sticky="ew", padx=24, pady=(0, 22))
        for col in range(4):
            card.grid_columnconfigure(col, weight=1 if col in (1, 3) else 0)
        self._dropdown_field(card, 0, 0, "Programme", self.programme_var, self.programmes)
        self._dropdown_field(card, 0, 2, "Semester", self.semester_var, self.semesters)
        self._dropdown_field(card, 1, 0, "Centre code", self.centre_code_var, self.centre_codes)
        self._entry_field(card, 1, 2, "Exam period", self.exam_period_var, "e.g. June, 2026")
        self._entry_field(card, 2, 0, "Publication date", self.publication_date_var, "DD.MM.YYYY")
        self._entry_field(card, 2, 2, "Issued on", self.issued_on_var, "DD.MM.YYYY")

    def _dropdown_field(self, parent: Any, row: int, col: int, label: str,
                        variable: ctk.StringVar, values: list[str]) -> None:
        ctk.CTkLabel(parent, text=label, text_color=TEXT,
                     font=ctk.CTkFont(size=12, weight="bold")).grid(
                         row=row, column=col, sticky="w", padx=(18, 8), pady=(15, 5))
        menu = ctk.CTkOptionMenu(parent, variable=variable, values=values,
                                 height=36, width=180, fg_color=NAVY,
                                 button_color=GREEN, button_hover_color=GREEN_HOVER,
                                 text_color=WHITE, corner_radius=8,
                                 font=ctk.CTkFont(size=12))
        menu.grid(row=row, column=col + 1, sticky="ew", padx=(0, 18), pady=(15, 5))
        menu.set(variable.get() or values[0])


    def _entry_field(self, parent: Any, row: int, col: int, label: str,
                    variable: ctk.StringVar, placeholder: str) -> None:
        ctk.CTkLabel(
            parent,
            text=label,
            text_color=TEXT,
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(
            row=row, column=col, sticky="w",
            padx=(18, 8), pady=(10, 5)
        )

        # Do not attach textvariable: it prevents the placeholder
        # from displaying correctly in this configuration.
        entry = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            placeholder_text_color="#888888",
            height=36,
            width=180,
            corner_radius=8,
            border_color=BORDER,
            fg_color="#FBFCFD",
            text_color=TEXT
        )

        # Keep the existing StringVar updated for validation and generation.
        def sync_variable(event=None) -> None:
            variable.set(entry.get())

        entry.bind("<KeyRelease>", sync_variable)
        entry.bind("<FocusOut>", sync_variable)

        entry.grid(
            row=row, column=col + 1, sticky="ew",
            padx=(0, 18), pady=(10, 5)
        )


    def _build_actions(self) -> None:
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.grid(row=6, column=0, sticky="ew", padx=24, pady=(0, 18))
        frame.grid_columnconfigure((0, 1), weight=1, uniform="actions")
        self.btn_inspect = ctk.CTkButton(
            frame, text="Inspect workbook", command=self._inspect_workbook,
            height=46, fg_color=NAVY, hover_color=NAVY_HOVER,
            corner_radius=9, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_inspect.grid(row=0, column=0, sticky="ew", padx=(0, 7))
        self.btn_generate = ctk.CTkButton(
            frame, text="Generate statements", command=self._generate_all_statements,
            height=46, fg_color=GREEN, hover_color=GREEN_HOVER,
            corner_radius=9, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_generate.grid(row=0, column=1, sticky="ew", padx=(7, 0))

    def _build_status(self) -> None:
        card = self._card(self.main_frame)
        card.grid(row=7, column=0, sticky="ew", padx=24, pady=(0, 22))
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text="Activity", text_color=NAVY,
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
                         row=0, column=0, sticky="w", padx=16, pady=(12, 0))
        self.status_label = ctk.CTkLabel(card, text="Ready. Select a workbook to begin.",
                                         text_color=MUTED, anchor="w",
                                         font=ctk.CTkFont(size=12))
        self.status_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(5, 12))
        ctk.CTkLabel(self.main_frame,
                     text=f"SANEX SYSTEMS", text_color="#10243A",
                     font=ctk.CTkFont(size=10)).grid(
                         row=8, column=0, sticky="w", padx=26, pady=(0, 18))

    def _build_progress_overlay(self) -> None:
        self.progress_overlay = ctk.CTkFrame(self, fg_color=WHITE,
                                             corner_radius=16, border_width=1,
                                             border_color=BORDER)
        self.progress_overlay.grid_columnconfigure(0, weight=1)
        content = ctk.CTkFrame(self.progress_overlay, fg_color="transparent")
        content.grid(row=0, column=0, padx=36, pady=28)
        self.progress_canvas = CircularProgress(content, size=112, thickness=9)
        self.progress_canvas.pack(pady=(0, 14))
        self.progress_title = ctk.CTkLabel(content, text="Preparing generation…",
                                           text_color=NAVY,
                                           font=ctk.CTkFont(size=16, weight="bold"))
        self.progress_title.pack()
        self.progress_detail = ctk.CTkLabel(content, text="Please keep the application open.",
                                             text_color=MUTED,
                                             font=ctk.CTkFont(size=12))
        self.progress_detail.pack(pady=(5, 0))
        self.progress_bar = ctk.CTkProgressBar(content, width=300, height=8,
                                               progress_color=GREEN,
                                               fg_color="#E5EAF0", corner_radius=5)
        self.progress_bar.pack(pady=(18, 0))
        self.progress_bar.set(0)

    # ------------------------------- Helpers -----------------------------
    def _show_dialog(self, title: str, message: str, kind: str = "info") -> None:
        PremiumDialog(self, title, message, kind)

    def _ask_overwrite(self, count: int, on_yes: Callable[[], None],
                       on_no: Callable[[], None],
                       on_cancel: Callable[[], None]) -> None:
        """Ask how to handle existing files; closing the dialog cancels cleanly."""
        dialog: PremiumDialog
        resolved = False

        def finish(callback: Callable[[], None]) -> None:
            nonlocal resolved
            if resolved:
                return
            resolved = True
            try:
                dialog.grab_release()
            except Exception:
                pass
            try:
                dialog.destroy()
            except Exception:
                pass
            callback()

        dialog = PremiumDialog(
            self, "Existing statements found",
            f"{count} statement file(s) already exist in the selected output folder.\n\n"
            "Choose whether to replace matching files or skip them and generate the rest.\n\n"
            "Close this window to cancel generation.",
            "warning",
            [("Replace files", lambda: finish(on_yes), GREEN),
             ("Skip existing", lambda: finish(on_no), NAVY),
             ("Cancel", lambda: finish(on_cancel), RED)],
        )
        # Handle the title-bar X and Escape exactly like the Cancel button.
        dialog.protocol("WM_DELETE_WINDOW", lambda: finish(on_cancel))
        dialog.bind("<Escape>", lambda _event: finish(on_cancel))

    def _browse_source(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="Select results workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if path:
            self.source_path.set(path)
            self.status_label.configure(text=f"Workbook selected: {Path(path).name}")

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(parent=self, title="Select output folder")
        if path:
            self.output_path.set(path)

    def _validate_inputs(self) -> bool:
        values = {
            "Results workbook": self.source_path.get().strip(),
            "Output folder": self.output_path.get().strip(),
            "Programme": self.programme_var.get().strip(),
            "Semester": self.semester_var.get().strip(),
            "Centre code": self.centre_code_var.get().strip(),
            "Exam period": self.exam_period_var.get().strip(),
            "Publication date": self.publication_date_var.get().strip(),
            "Issued on": self.issued_on_var.get().strip(),
        }
        missing = [name for name, value in values.items() if not value]
        if missing:
            self._show_dialog("Missing information",
                              "Complete the following fields before generating:\n\n" +
                              "\n".join(f"• {item}" for item in missing), "warning")
            return False

        source = Path(values["Results workbook"])
        if not source.is_file():
            self._show_dialog("Workbook not found", f"The selected workbook does not exist:\n\n{source}", "error")
            return False
        if source.suffix.lower() not in {".xlsx", ".xlsm"}:
            self._show_dialog("Unsupported file", "Select an .xlsx or .xlsm workbook.", "warning")
            return False

        for label, var in (("Publication date", self.publication_date_var),
                            ("Issued on", self.issued_on_var)):
            try:
                datetime.strptime(var.get().strip(), "%d.%m.%Y")
            except ValueError:
                self._show_dialog("Invalid date format",
                                  f"{label} must use DD.MM.YYYY.\n\nExample: 20.09.2026", "error")
                return False
        return True

    def _set_busy(self, busy: bool) -> None:
        self.job_running = busy
        state = "disabled" if busy else "normal"
        self.btn_inspect.configure(state=state)
        self.btn_generate.configure(state=state)

    def _update_stats(self, students: int | None = None,
                      generated: int | None = None,
                      skipped: int | None = None,
                      errors: int | None = None) -> None:
        if students is not None:
            self.students_processed = students
            self.lbl_students.configure(text=str(students))
        if generated is not None:
            self.statements_generated = generated
            self.lbl_statements.configure(text=str(generated))
        if skipped is not None:
            self.skipped_count = skipped
            self.lbl_skipped.configure(text=str(skipped))
        if errors is not None:
            self.errors_count = errors
            self.lbl_errors.configure(text=str(errors))

    def _load_records(self, path: str) -> tuple[list[Any], list[Any], list[str], str]:
        """Load and validate workbook data; always close the workbook."""
        workbook = None
        try:
            workbook = load_workbook(path)
            ws = workbook.worksheets[0]
            students, modules = extract_students(ws)
            errors = validate_students(students)
            return students, modules, errors, ws.title
        finally:
            if workbook is not None:
                workbook.close()

    # ------------------------------ Inspection ---------------------------
    def _inspect_workbook(self) -> None:
        if self.job_running:
            return
        path = self.source_path.get().strip()
        if not path:
            self._show_dialog("Workbook required", "Select a results workbook first.", "warning")
            return
        if not Path(path).is_file():
            self._show_dialog("Workbook not found", f"Cannot find this file:\n\n{path}", "error")
            return
        self._set_busy(True)
        self.status_label.configure(text="Inspecting workbook…")
        try:
            students, modules, errors, sheet_name = self._load_records(path)
            self._update_stats(students=len(students), generated=0, skipped=0, errors=len(errors))
            lines = [
                "WORKBOOK INSPECTION REPORT", "=" * 54, "",
                f"File: {Path(path).name}", f"Worksheet: {sheet_name}",
                f"Students detected: {len(students)}", f"Modules detected: {len(modules)}",
                f"Validation issues: {len(errors)}", "",
                "MODULES", "-" * 54,
            ]
            if modules:
                for module in modules:
                    code = module.get("code", "?") if isinstance(module, dict) else getattr(module, "code", "?")
                    title = module.get("title", "") if isinstance(module, dict) else getattr(module, "title", "")
                    lines.append(f"• {code} — {title}" if title else f"• {code}")
            else:
                lines.append("No modules detected.")
            lines.extend(["", "STUDENT PREVIEW (first 10)", "-" * 54])
            for student in students[:10]:
                reg = getattr(student, "registration_number", "")
                name = getattr(student, "name", "")
                student_modules = getattr(student, "modules", [])
                gpa = getattr(student, "semester_gpa", "—")
                lines.append(f"• {reg} | {name} | {len(student_modules)} modules | GPA: {gpa}")
            if len(students) > 10:
                lines.append(f"… and {len(students) - 10} more student(s).")
            lines.extend(["", "VALIDATION", "-" * 54])
            lines.extend([f"• {error}" for error in errors[:40]] if errors else ["Validation passed — no issues found."])
            if len(errors) > 40:
                lines.append(f"… and {len(errors) - 40} more issue(s).")
            self._show_dialog("Workbook inspection", "\n".join(lines), "info")
            self.status_label.configure(text=f"Inspection complete: {len(students)} student(s), {len(errors)} issue(s).")
        except Exception as exc:
            logging.exception("Workbook inspection failed")
            self._show_dialog("Inspection failed", f"Could not inspect the workbook.\n\n{exc}\n\nSee log: {LOG_PATH}", "error")
            self.status_label.configure(text="Workbook inspection failed. See the log for details.")
        finally:
            self._set_busy(False)

    # ------------------------------ Generation ---------------------------
    def _generate_all_statements(self) -> None:
        if self.job_running or not self._validate_inputs():
            return
        if not TEMPLATE_PATH.is_file():
            self._show_dialog("Template missing",
                              f"Statement template not found:\n\n{TEMPLATE_PATH}\n\n"
                              "Place statement_template.xlsx inside the project's templates folder.", "error")
            return

        self._set_busy(True)
        self.status_label.configure(text="Reading workbook and validating student records…")
        try:
            students, _modules, errors, sheet_name = self._load_records(self.source_path.get().strip())
            if not students:
                self._set_busy(False)
                self._show_dialog("No students found", "No student records were detected in the first worksheet.", "warning")
                self.status_label.configure(text="No student records detected.")
                return
            self._update_stats(students=len(students), generated=0, skipped=0, errors=len(errors))
            if errors:
                details = "\n".join(f"• {error}" for error in errors[:25])
                if len(errors) > 25:
                    details += f"\n… and {len(errors) - 25} more issue(s)."
                self._set_busy(False)
                self._show_dialog("Validation failed",
                                  f"Found {len(errors)} validation issue(s). No statements were generated.\n\n"
                                  f"Worksheet: {sheet_name}\n\n{details}", "error")
                self.status_label.configure(text=f"Generation stopped: {len(errors)} validation issue(s).")
                return

            output_folder = Path(self.output_path.get().strip()).expanduser()
            try:
                output_folder.mkdir(parents=True, exist_ok=True)
                # Test that the selected output directory can be written to.
                test_file = output_folder / ".marks_generator_write_test"
                test_file.touch(exist_ok=True)
                test_file.unlink(missing_ok=True)
            except OSError as exc:
                self._set_busy(False)
                self._show_dialog("Output folder unavailable", f"Cannot write to the selected folder:\n\n{exc}", "error")
                return

            existing = [student for student in students
                        if (output_folder / f"{student.registration_number}.xlsx").exists()]
            if existing:
                self._ask_overwrite(
                    len(existing),
                    on_yes=lambda: self._start_generation(students, output_folder, True),
                    on_no=lambda: self._start_generation(students, output_folder, False),
                    on_cancel=self._cancel_generation_prompt,
                )
            else:
                self._start_generation(students, output_folder, False)
        except Exception as exc:
            logging.exception("Could not prepare statement generation")
            self._set_busy(False)
            self._show_dialog("Preparation failed", f"Could not prepare generation:\n\n{exc}\n\nSee log: {LOG_PATH}", "error")

    def _cancel_generation_prompt(self) -> None:
        """Restore the UI when the user cancels the existing-files prompt."""
        self.progress_overlay.place_forget()
        self._set_busy(False)
        self.status_label.configure(text="Generation cancelled. No statements were changed.")
        logging.info("Generation cancelled at existing-files confirmation prompt")

    def _start_generation(self, students: list[Any], output_folder: Path,
                          overwrite_existing: bool) -> None:
        exam_details = {
            "programme": self.programme_var.get().strip(),
            "semester": self.semester_var.get().strip(),
            "centre": self.centre_code_var.get().strip(),
            "examination_period": self.exam_period_var.get().strip(),
            "publication_date": self.publication_date_var.get().strip(),
            "issued_on": self.issued_on_var.get().strip(),
        }
        self.progress_canvas.set(0)
        self.progress_bar.set(0)
        self.progress_title.configure(text="Generating statements…")
        self.progress_detail.configure(text="Starting batch…")
        self.progress_overlay.place(relx=0.5, rely=0.5, anchor="center")
        self.status_label.configure(text="Statement generation is in progress…")
        self._update_stats(generated=0, skipped=0, errors=0)

        worker = threading.Thread(
            target=self._generation_worker,
            args=(students, output_folder, overwrite_existing, exam_details),
            name="StatementGenerationWorker", daemon=True)
        worker.start()
        self._schedule_event_poll()

    def _generation_worker(self, students: list[Any], output_folder: Path,
                           overwrite_existing: bool, details: dict[str, str]) -> None:
        generated = 0
        skipped = 0
        failed: list[str] = []
        total = len(students)
        fatal_error = None
        logging.info("Generation started: %s students; output=%s; overwrite=%s",
                     total, output_folder, overwrite_existing)
        try:
            for index, student in enumerate(students, start=1):
                registration = str(getattr(student, "registration_number", "")).strip()
                output_path = output_folder / f"{registration}.xlsx"
                if output_path.exists() and not overwrite_existing:
                    skipped += 1
                    self.events.put(("progress", index, total, registration,
                                     "Skipped — file already exists"))
                    continue
                try:
                    generate_student_statement(
                        student=student,
                        template_path=TEMPLATE_PATH,
                        output_folder=output_folder,
                        programme=details["programme"],
                        examination_period=details["examination_period"],
                        centre=details["centre"],
                        publication_date=details["publication_date"],
                        issued_on=details["issued_on"],
                        semester=details["semester"],
                    )
                    generated += 1
                    logging.info("Generated statement for %s", registration)
                except Exception as exc:
                    failed.append(f"{registration}: {exc}")
                    logging.exception("Statement generation failed for %s", registration)
                self.events.put(("progress", index, total, registration,
                                 "Processed"))
        except Exception as exc:
            fatal_error = str(exc)
            logging.exception("Unexpected generation worker failure")
        finally:
            result = {
                "total": total,
                "generated": generated,
                "skipped": skipped,
                "failed": failed,
                "fatal_error": fatal_error,
                "output_folder": str(output_folder.resolve()),
            }
            logging.info("Generation finished: %s", result)
            self.events.put(("complete", result))

    def _schedule_event_poll(self) -> None:
        if not self._poll_scheduled:
            self._poll_scheduled = True
            self.after(100, self._poll_events)

    def _poll_events(self) -> None:
        self._poll_scheduled = False
        finished = False
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    _, current, total, registration, message = event
                    percent = current / total if total else 0
                    self.progress_canvas.set(percent * 100)
                    self.progress_bar.set(percent)
                    self.progress_title.configure(text=f"Processing {current} of {total}")
                    self.progress_detail.configure(text=f"{registration}  •  {message}")
                elif kind == "complete":
                    finished = True
                    self._finish_generation(event[1])
        except queue.Empty:
            pass
        if self.job_running and not finished:
            self._schedule_event_poll()

    def _finish_generation(self, result: dict[str, Any]) -> None:
        generated = result["generated"]
        skipped = result["skipped"]
        failed = result["failed"]
        self._update_stats(generated=generated, skipped=skipped, errors=len(failed))
        self.progress_canvas.set(100)
        self.progress_bar.set(1)
        self.progress_title.configure(text="Generation complete")
        self.progress_detail.configure(text="Preparing final summary…")
        self.progress_overlay.place_forget()
        self._set_busy(False)

        lines = [
            "BATCH GENERATION SUMMARY", "=" * 42, "",
            f"Students detected:       {result['total']}",
            f"Statements generated:    {generated}",
            f"Files skipped:            {skipped}",
            f"Failed:                   {len(failed)}", "",
            "Output folder:", result["output_folder"],
        ]
        if result.get("fatal_error"):
            lines.extend(["", "UNEXPECTED WORKER ERROR", "-" * 42,
                          str(result["fatal_error"])])
        if failed:
            lines.extend(["", "ERRORS (first 15)", "-" * 42])
            lines.extend(f"• {error}" for error in failed[:15])
            if len(failed) > 15:
                lines.append(f"… and {len(failed) - 15} more error(s). See the log for full details.")
        lines.extend(["", f"Detailed log: {LOG_PATH}"])
        if failed or result.get("fatal_error"):
            title, kind = "Generation finished with errors", "warning"
        else:
            title, kind = "Generation completed successfully", "success"
        self._show_dialog(title, "\n".join(lines), kind)
        self.status_label.configure(
            text=f"Finished: {generated} generated, {skipped} skipped, {len(failed)} failed.")

    def _on_close(self) -> None:
        if self.job_running:
            self._show_dialog("Generation in progress",
                              "Statement generation is still running. Please wait for it to finish before closing the application.",
                              "warning")
            return
        self.destroy()


if __name__ == "__main__":
    app = MarksGeneratorApp()
    app.mainloop()
