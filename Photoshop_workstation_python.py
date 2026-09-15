import os
import sys
import subprocess
import tempfile
import math
import html
import json   

import win32com.client

from PIL import Image, ImageEnhance

from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QRectF, QUrl, QPointF,
    QPropertyAnimation, QEasingCurve,
)
from PyQt6.QtGui import (
    QFont, QPixmap, QImage, QCursor, QPainter, QColor,
    QLinearGradient, QPen, QConicalGradient, QDesktopServices,
    QRadialGradient, QIcon,
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QComboBox, QCheckBox, QPushButton,
    QTextEdit, QProgressBar, QFrame, QDialog, QTextBrowser,
    QSizePolicy, QFileDialog, QListView, QLayout,
    QListWidget, QListWidgetItem, QAbstractItemView,
)
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from PyQt6.QtMultimediaWidgets import QVideoWidget


def resource_path(relative_path):
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ----------------------------------------------------------------------
# ★★★ TKINTER-ДИАЛОГ ВВОДА НОМЕРА ШКОЛЫ ★★★
# Тёмный фон + закруглённые кнопки в стиле приложения.
# Ввод и вставка работают на любой раскладке.
# Окно строго по центру экрана (Win32 API + DPI aware).
# ----------------------------------------------------------------------
def tkinter_ask_school_number(parent_title="Создание структуры папок"):
    import tkinter as tk

    result = {"value": None}

    # DPI awareness
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title(parent_title)
    root.resizable(False, False)
    root.configure(bg="#161616")

    win_w = 520
    win_h = 260

    # Сначала показываем окно с произвольной позицией,
    # чтобы tkinter успел инициализировать графику.
    root.geometry(f"{win_w}x{win_h}+0+0")
    root.update_idletasks()
    root.update()

    # ★ Получаем физические размеры экрана через Win32 API.
    screen_w = win_w
    screen_h = win_h
    try:
        import ctypes
        user32 = ctypes.windll.user32
        screen_w = user32.GetSystemMetrics(0)
        screen_h = user32.GetSystemMetrics(1)
    except Exception:
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()

    pos_x = max(0, (screen_w - win_w) // 2)
    pos_y = max(0, (screen_h - win_h) // 2)

    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
    root.update_idletasks()
    root.update()

    BG = "#161616"
    FIELD_BG = "#1A1A1A"
    BORDER = "#333333"
    ACCENT = "#88FFCC"
    TEXT = "#FFFFFF"
    SUBTLE = "#AAAAAA"

    container = tk.Frame(root, bg=BG)
    container.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)

    title_lbl = tk.Label(
        container,
        text="Создание структуры папок для НОВОЙ ШКОЛЫ",
        bg=BG, fg=ACCENT,
        font=("Segoe UI", 12, "bold"),
    )
    title_lbl.pack(pady=(0, 8))

    hint_lbl = tk.Label(
        container,
        text="Введите номер школы (минимум 3 символа).\n"
             "Будут созданы: 1520, 2030, QR, образцы, коды.",
        bg=BG, fg=SUBTLE,
        font=("Segoe UI", 9),
        justify="center",
    )
    hint_lbl.pack(pady=(0, 14))

    entry_wrap = tk.Frame(container, bg=BORDER, bd=0, highlightthickness=0)
    entry_wrap.pack(fill=tk.X, pady=(0, 16))

    entry = tk.Entry(
        entry_wrap,
        font=("Consolas", 13, "bold"),
        bg=FIELD_BG, fg=TEXT,
        insertbackground=ACCENT,
        relief=tk.FLAT,
        bd=0,
    )
    entry.pack(fill=tk.X, padx=1, pady=1, ipady=10)
    entry.focus_set()

    def on_focus_in(event=None):
        entry_wrap.configure(bg=ACCENT)

    def on_focus_out(event=None):
        entry_wrap.configure(bg=BORDER)

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

    def select_all(event=None):
        entry.select_range(0, tk.END)
        entry.icursor(tk.END)
        return "break"

    def paste_clipboard(event=None):
        try:
            clipboard_text = root.clipboard_get()
            try:
                entry.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass
            entry.insert(tk.INSERT, clipboard_text)
        except tk.TclError:
            pass
        return "break"

    def copy_clipboard(event=None):
        try:
            selected = entry.selection_get()
            root.clipboard_clear()
            root.clipboard_append(selected)
        except tk.TclError:
            pass
        return "break"

    def cut_clipboard(event=None):
        try:
            selected = entry.selection_get()
            root.clipboard_clear()
            root.clipboard_append(selected)
            entry.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
        return "break"

    def handle_ctrl_hotkeys(event):
        is_ctrl = (
            (event.state & 0x0004) != 0
            or (event.state & 0x0001) != 0
            or (event.state & 4)
        )
        if is_ctrl:
            if event.keycode == 65:
                return select_all()
            elif event.keycode == 86:
                return paste_clipboard()
            elif event.keycode == 67:
                return copy_clipboard()
            elif event.keycode == 88:
                return cut_clipboard()

    context_menu = tk.Menu(
        entry, tearoff=0,
        bg="#1E2229", fg="#FFFFFF",
        activebackground="#2A3F55", activeforeground="#7CC4FF",
        bd=0, relief=tk.FLAT,
    )
    context_menu.add_command(label="Выделить всё", command=select_all)
    context_menu.add_command(label="Вставить", command=paste_clipboard)
    context_menu.add_command(label="Копировать", command=copy_clipboard)
    context_menu.add_command(label="Вырезать", command=cut_clipboard)

    def show_context_menu(event):
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    entry.bind("<Button-3>", show_context_menu)
    entry.bind("<Key>", handle_ctrl_hotkeys)

    def on_ok(event=None):
        text = entry.get().strip()
        if not text:
            return "break"
        result["value"] = text
        root.destroy()
        return "break"

    def on_cancel(event=None):
        result["value"] = None
        root.destroy()
        return "break"

    # ── ЗАКРУГЛЁННЫЕ КНОПКИ ──────────────────────────────────────────
    btn_row = tk.Frame(container, bg=BG)
    btn_row.pack(fill=tk.X)

    def make_rounded_button(parent, text, fill_color, text_color,
                            hover_color, command,
                            width=210, height=44, radius=16):
        """
        Кнопка с настоящими закруглёнными углами.
        Рисуется через 4 дуги (create_arc) + 2 прямоугольника.
        Без обводки, полностью однотонная.
        """
        canvas = tk.Canvas(
            parent, width=width, height=height,
            bg=BG, highlightthickness=0, bd=0,
        )
        canvas.pack(side=tk.LEFT, padx=6, expand=True)

        shape_items = []

        shape_items.append(canvas.create_rectangle(
            radius, 0, width - radius, height,
            fill=fill_color, outline="",
        ))
        shape_items.append(canvas.create_rectangle(
            0, radius, width, height - radius,
            fill=fill_color, outline="",
        ))

        shape_items.append(canvas.create_arc(
            0, 0, radius * 2, radius * 2,
            start=90, extent=90,
            fill=fill_color, outline=fill_color, style="pieslice",
        ))
        shape_items.append(canvas.create_arc(
            width - radius * 2, 0, width, radius * 2,
            start=0, extent=90,
            fill=fill_color, outline=fill_color, style="pieslice",
        ))
        shape_items.append(canvas.create_arc(
            width - radius * 2, height - radius * 2, width, height,
            start=270, extent=90,
            fill=fill_color, outline=fill_color, style="pieslice",
        ))
        shape_items.append(canvas.create_arc(
            0, height - radius * 2, radius * 2, height,
            start=180, extent=90,
            fill=fill_color, outline=fill_color, style="pieslice",
        ))

        canvas.create_text(
            width // 2, height // 2,
            text=text, fill=text_color,
            font=("Segoe UI", 11, "bold"),
        )

        def set_fill(color):
            for item in shape_items:
                canvas.itemconfig(item, fill=color, outline=color)

        def on_enter(event):
            set_fill(hover_color)

        def on_leave(event):
            set_fill(fill_color)

        def on_click(event):
            command()

        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        canvas.bind("<Button-1>", on_click)
        canvas.configure(cursor="hand2")
        return canvas

    make_rounded_button(
        btn_row,
        text="Создать",
        fill_color="#A8F0C0",
        text_color="#000000",
        hover_color="#88E0A0",
        command=on_ok,
    )

    make_rounded_button(
        btn_row,
        text="Отмена",
        fill_color="#FFD27A",
        text_color="#000000",
        hover_color="#F0C060",
        command=on_cancel,
    )

    def recenter():
        try:
            import ctypes
            user32 = ctypes.windll.user32
            sw = user32.GetSystemMetrics(0)
            sh = user32.GetSystemMetrics(1)
        except Exception:
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
        px = max(0, (sw - win_w) // 2)
        py = max(0, (sh - win_h) // 2)
        root.geometry(f"{win_w}x{win_h}+{px}+{py}")

    root.after(10, recenter)

    try:
        root.attributes("-topmost", True)
        root.after(200, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    root.mainloop()
    return result["value"]


# ----------------------------------------------------------------------
# ★★★ ПРЕСЕТЫ ДЛЯ БЫСТРОГО BATCH ACTIONS ★★★
# ----------------------------------------------------------------------
ACTION_PRESETS = [
    ("RAW", "CAMERA RAW", "raw", "raw"),
    ("RAW + Export JPG", "CAMERA RAW ", "Raw + Export JPG", "raw_export"),
    ("Export JPG", "ExportJPG", "ExportJPG", "export_jpg"),
    ("BRIGHT +0.15", "BRIGHT ", "+0.15 bright", "bright_plus"),
    ("BRIGHT -0.15", "BRIGHT ", "-0.15 bright", "bright_minus"),
    ("QR CODE вставка", "QR_add", "QR_fast", "qr_code"),
    ("BLUR фон", "BACKGROUND_BLUR", "select_inverse_blursurface", "blur_bg"),
    # ★★★ ДОБАВЛЕННЫЕ КНОПКИ ЗАМЕНЫ РАМОК ★★★
    ("Заменить рамку 2030", "ЗАМЕНА РАМОК", "ЗАМЕНА РАМКИ 2030", "replace_2030"),
    ("Заменить рамку 1520", "ЗАМЕНА РАМОК", "ЗАМЕНА РАМКИ 1520", "replace_1520"),
]


def get_preset_button_style(color_key):
    palettes = {
        "raw": ("#3A2410", "#FFB74D", "#5A3818", "#FFCC80", "#FFA726"),
        "bright_plus": ("#2A3A10", "#E5FF55", "#3A4A18", "#F0FF88", "#D4E838"),
        "bright_minus": ("#1A1A3A", "#B39DFF", "#2A2A5A", "#C7B7FF", "#9F88E8"),
        "export_jpg": ("#103A3A", "#5CE8E0", "#185A5A", "#88F0E8", "#3DD0C8"),
        "qr_code": ("#3A103A", "#FF70D8", "#5A185A", "#FFA0E8", "#E858C0"),
        "add_frames": ("#103A18", "#88FFA8", "#185A28", "#B0FFC8", "#58E878"),
        "raw_export": ("#10243A", "#7CC4FF", "#183A5A", "#A8D8FF", "#58A8E8"),
        "blur_bg": ("#2A103A", "#D89BFF", "#3A185A", "#E8B7FF", "#B87BE8"),
        # Новые цвета для кнопок замены рамок
        "replace_2030": ("#0A2A2A", "#00E5FF", "#104040", "#80F0FF", "#00B8D4"),
        "replace_1520": ("#3A0A1A", "#FF4081", "#5A1028", "#FF80AB", "#E91E63"),
    }
    bg, fg, border, hover_fg, pressed_bg = palettes.get(
        color_key, ("#2A2A2A", "#CCFF00", "#444444", "#E5FF55", "#1A1A1A")
    )
    return f"""
        QPushButton {{
            background-color: {bg};
            color: {fg};
            border: 2px solid {border};
            border-radius: 12px;
            padding: 8px 6px;
            font-weight: bold;
            font-size: 11px;
        }}
        QPushButton:hover {{
            background-color: {border};
            border: 1px solid {fg};
            color: {hover_fg};
        }}
        QPushButton:pressed {{
            background-color: {pressed_bg};
            border: 1px solid {fg};
        }}
    """


# ----------------------------------------------------------------------
# RGB ANIMATED BORDER
# ----------------------------------------------------------------------
class RGBBorderWidget(QWidget):
    def __init__(self, child_widget, palette=None, speed=4, radius=18, parent=None):
        super().__init__(parent)
        self.child = child_widget
        self.angle = 0
        self.active = False
        self.speed = speed
        self.radius = radius
        self.palette = palette or [
            QColor(200, 80, 130), QColor(200, 80, 80),
            QColor(200, 150, 70), QColor(90, 180, 100),
            QColor(80, 160, 200), QColor(120, 80, 190),
            QColor(200, 80, 130),
        ]
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(0)
        layout.addWidget(self.child)

        self.timer = QTimer(self)
        self.timer.setInterval(30)
        self.timer.timeout.connect(self._rotate)

    def start_border(self):
        self.active = True
        if not self.timer.isActive():
            self.timer.start()
        self.update()

    def stop_border(self):
        self.active = False
        self.timer.stop()
        self.update()

    def _rotate(self):
        self.angle = (self.angle + self.speed) % 360
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.active:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = self.rect().adjusted(1, 1, -1, -1)
        center = rect.center()
        gradient = QConicalGradient(
            QPointF(center.x(), center.y()), float(self.angle)
        )
        n = len(self.palette) - 1
        for i, col in enumerate(self.palette):
            gradient.setColorAt(i / n, col)
        pen = QPen()
        pen.setWidth(2)
        pen.setBrush(gradient)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(rect, self.radius, self.radius)


# ----------------------------------------------------------------------
# КАСТОМНЫЙ COMBOBOX
# ----------------------------------------------------------------------
class PixelComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.setMaxVisibleItems(30)
        self.setFixedHeight(32)
        view = QListView()
        view.setStyleSheet(
            """
            QListView {
                background-color: #1E2229; color: #FFFFFF;
                border: 1px solid #3A4150; border-radius: 18px;
                padding: 6px; outline: none; font-size: 12px;
            }
            QListView::item { padding: 8px 12px; border-radius: 10px; min-height: 20px; }
            QListView::item:hover { background-color: #2A2F3A; }
            QListView::item:selected { background-color: #1A2740; color: #7CC4FF; }
            QScrollBar:vertical { background: #1E2229; width: 8px;
                margin: 6px 2px 6px 0px; border-radius: 4px; }
            QScrollBar::handle:vertical { background: #3A4150;
                min-height: 20px; border-radius: 4px; }
            QScrollBar::handle:vertical:hover { background: #4A5568; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }
            """
        )
        self.setView(view)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = self.rect().adjusted(1, 1, -1, -1)
        radius = 18
        painter.setBrush(QColor("#1E2229"))
        painter.setPen(QPen(QColor("#3A4150"), 1))
        painter.drawRoundedRect(rect, radius, radius)
        painter.setPen(QColor("#FFFFFF"))
        text = self.currentText()
        text_rect = rect.adjusted(16, 0, -40, 0)
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            text,
        )
        arrow_cx = rect.right() - 20
        arrow_cy = rect.center().y()
        painter.setBrush(QColor("#7CFFCB"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawPolygon(
            QPointF(arrow_cx - 5, arrow_cy - 3),
            QPointF(arrow_cx + 5, arrow_cy - 3),
            QPointF(arrow_cx, arrow_cy + 4),
        )

    def showPopup(self):
        super().showPopup()
        popup = self.view().parentWidget()
        if popup:
            view = self.view()
            item_count = view.model().rowCount()
            item_height = 32
            max_height = 600
            calculated_height = min(item_count * item_height + 12, max_height)
            popup.setFixedHeight(calculated_height)
            popup.setStyleSheet(
                "QFrame { background-color: #1E2229; border: 1px solid #3A4150;"
                "border-radius: 18px; }"
            )


# ----------------------------------------------------------------------
# ДИАЛОГ ВЫБОРА PSD-ФАЙЛОВ
# ----------------------------------------------------------------------
class PickPSDFilesDialog(QDialog):
    def __init__(self, parent=None, school_folder="", preselected=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор PSD-файлов для обработки")
        self.setFixedSize(700, 520)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.school_folder = school_folder
        self.selected_files = list(preselected or [])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(12)

        title = QLabel("Выберите PSD-файлы для BRIGHT-обработки")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setStyleSheet("color: #E5FF55; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        hint = QLabel(
            "Отметьте нужные PSD-файлы (Ctrl+A — выбрать все).\n"
            "Обрабатываться будут только выбранные файлы."
        )
        hint.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self.list_widget.setStyleSheet(
            """
            QListWidget {
                background-color: #1A1A1A; color: #FFFFFF;
                border: 1px solid #333333; border-radius: 8px;
                padding: 6px; font-family: Consolas; font-size: 11px;
            }
            QListWidget::item { padding: 6px 8px; border-radius: 6px; }
            QListWidget::item:selected { background-color: #2A3A10; color: #E5FF55; }
            QListWidget::item:hover { background-color: #2A2A2A; }
            """
        )
        layout.addWidget(self.list_widget, 1)
        self._populate_files()

        self.lbl_count = QLabel("Выбрано: 0")
        self.lbl_count.setStyleSheet(
            "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
        )
        self.lbl_count.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_count)
        self.list_widget.itemSelectionChanged.connect(self._update_count)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_all = QPushButton("Выбрать все")
        btn_all.setStyleSheet(
            "QPushButton { background-color: #3A3A3A; color: #FFFFFF;"
            "border: 1px solid #555555; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #4A4A4A;"
            "border-color: #88FFCC; color: #88FFCC; }"
        )
        btn_all.setMinimumHeight(34)
        btn_all.clicked.connect(self.list_widget.selectAll)
        btn_row.addWidget(btn_all, 1)

        btn_ok = QPushButton("Применить")
        btn_ok.setStyleSheet(
            "QPushButton { background-color: #2E7D32; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #1B5E20; }"
        )
        btn_ok.setMinimumHeight(34)
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_ok, 1)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #D32F2F; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #B71C1C; }"
        )
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel, 1)
        layout.addLayout(btn_row)

    def _populate_files(self):
        if not self.school_folder or not os.path.isdir(self.school_folder):
            return
        for dp, dn, filenames in os.walk(self.school_folder):
            for f in filenames:
                if f.lower().endswith(".psd"):
                    full = os.path.join(dp, f)
                    rel = os.path.relpath(full, self.school_folder)
                    item = QListWidgetItem(rel)
                    item.setData(Qt.ItemDataRole.UserRole, full)
                    self.list_widget.addItem(item)
                    if full in self.selected_files:
                        item.setSelected(True)

    def _update_count(self):
        cnt = len(self.list_widget.selectedItems())
        self.lbl_count.setText(f"Выбрано: {cnt}")

    def _on_ok(self):
        items = self.list_widget.selectedItems()
        if not items:
            self.lbl_count.setText("❌ Выберите хотя бы один PSD файл!")
            self.lbl_count.setStyleSheet(
                "color: #FF5555; font-weight: bold; font-size: 11px; border: none;"
            )
            return
        self.selected_files = [it.data(Qt.ItemDataRole.UserRole) for it in items]
        self.accept()


# ----------------------------------------------------------------------
# ДИАЛОГ BATCH ACTIONS (пункт 4)
# ----------------------------------------------------------------------
class BatchActionsDialog(QDialog):
    def __init__(self, parent=None, default_set="", default_actions=""):
        super().__init__(parent)
        self.setWindowTitle("БЫСТРАЯ ПАКЕТНАЯ ОБРАБОТКА")
        self.setFixedSize(720, 760) # Чуть увеличили высоту, чтобы влезли новые кнопки
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.action_set = default_set
        self.actions_text = default_actions
        self.fast_mode = True
        self.auto_alert = True
        self.main_window = parent

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(10)

        title = QLabel("Настройка запуска Action Set")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #00BFFF; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        hint = QLabel(
            "Укажите название Action Set и список Actions через запятую,\n"
            "либо нажмите одну из быстрых кнопок ниже."
        )
        hint.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        lbl_presets = QLabel("Быстрые действия:")
        lbl_presets.setStyleSheet(
            "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
        )
        layout.addWidget(lbl_presets)

        presets_grid = QGridLayout()
        presets_grid.setSpacing(8)
        for i, preset in enumerate(ACTION_PRESETS):
            preset_name = preset[0]
            p_set = preset[1]
            p_actions = preset[2]
            color_key = preset[3] if len(preset) > 3 else "raw"
            p_btn = QPushButton(preset_name)
            p_btn.setStyleSheet(get_preset_button_style(color_key))
            p_btn.setMinimumHeight(36)
            p_btn.clicked.connect(
                lambda checked=False, s=p_set, a=p_actions, k=color_key:
                self._apply_preset(s, a, k)
            )
            row = i // 3
            col = i % 3
            presets_grid.addWidget(p_btn, row, col)
        layout.addLayout(presets_grid)
        layout.addSpacing(6)

        lbl_set = QLabel("Введите Action Set:")
        lbl_set.setStyleSheet(
            "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
        )
        layout.addWidget(lbl_set)

        self.input_set = QTextEdit()
        self.input_set.setFixedHeight(40)
        self.input_set.setFont(QFont("Consolas", 12, QFont.Weight.Bold))
        self.input_set.setStyleSheet(
            "QTextEdit { background-color: #1A1A1A; color: #FFFFFF;"
            "border: 1px solid #333333; border-radius: 6px; padding: 6px; }"
            "QTextEdit:focus { border: 1px solid #CCFF00; }"
        )
        self.input_set.setPlainText(default_set or "")
        layout.addWidget(self.input_set)

        lbl_act = QLabel("Введите ваш Action:")
        lbl_act.setStyleSheet(
            "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
        )
        layout.addWidget(lbl_act)

        self.input_actions = QTextEdit()
        self.input_actions.setFixedHeight(60)
        self.input_actions.setFont(QFont("Consolas", 11))
        self.input_actions.setStyleSheet(
            "QTextEdit { background-color: #1A1A1A; color: #FFFFFF;"
            "border: 1px solid #333333; border-radius: 6px; padding: 6px; }"
            "QTextEdit:focus { border: 1px solid #CCFF00; }"
        )
        self.input_actions.setPlainText(default_actions or "")
        layout.addWidget(self.input_actions)

        chk_style = (
            "QCheckBox { color: white; font-weight: bold; border: none; padding: 4px; }"
            "QCheckBox::indicator { border: 1px solid white; border-radius: 3px;"
            "background: #222222; width: 14px; height: 14px; }"
            "QCheckBox::indicator:checked { background-color: #CCFF00;"
            "border: 1px solid white; }"
        )
        self.chk_fast = QCheckBox("Ускоренный режим")
        self.chk_fast.setStyleSheet(chk_style)
        self.chk_fast.setChecked(True)
        layout.addWidget(self.chk_fast)
        self.chk_alert = QCheckBox("Игнорировать alert()")
        self.chk_alert.setStyleSheet(chk_style)
        self.chk_alert.setChecked(True)
        layout.addWidget(self.chk_alert)
        layout.addStretch()

        self.lbl_status = QLabel()
        self.lbl_status.setStyleSheet(
            "border: none; font-size: 11px; font-weight: bold;"
        )
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_status)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_start = QPushButton("Запустить")
        self.btn_start.setStyleSheet(
            "QPushButton { background-color: #2E7D32; color: white;"
            "border: none; border-radius: 15px; padding: 10px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #1B5E20; }"
            "QPushButton:disabled { background-color: #333333; color: #777777; }"
        )
        self.btn_start.clicked.connect(self._on_start)
        btn_row.addWidget(self.btn_start, 1)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #D32F2F; color: white;"
            "border: none; border-radius: 15px; padding: 10px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #B71C1C; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel, 1)
        layout.addLayout(btn_row)

        self.input_set.textChanged.connect(self._validate)
        self.input_actions.textChanged.connect(self._validate)
        self._validate()

    def _frames_selected(self):
        if self.main_window is None:
            return True
        try:
            return (
                self.main_window.chk_1520.isChecked()
                or self.main_window.chk_2030.isChecked()
            )
        except Exception:
            return True

    def _apply_preset(self, preset_set, preset_actions, color_key):
        if not self._frames_selected():
            try:
                if self.main_window is not None:
                    self.main_window.append_log(
                        "__ERROR__Пресет не применён: не выбраны папки "
                        "1520 или 2030! Поставьте галочку под школой."
                    )
            except Exception:
                pass
            self.lbl_status.setText(
                "❌ Сначала отметьте 1520 или 2030 под школой!"
            )
            self.lbl_status.setStyleSheet(
                "color: #FF5555; border: none; font-size: 11px; font-weight: bold;"
            )
            return

        if color_key in ("bright_plus", "bright_minus"):
            if self.main_window is None:
                return
            school_folder = self.main_window.selected_folder or ""
            if not school_folder or not os.path.isdir(school_folder):
                try:
                    self.main_window.append_log(
                        "__ERROR__Папка школы не найдена для выбора PSD."
                    )
                except Exception:
                    pass
                self.lbl_status.setText("❌ Папка школы не найдена!")
                self.lbl_status.setStyleSheet(
                    "color: #FF5555; border: none; font-size: 11px; font-weight: bold;"
                )
                return

            pick_dlg = PickPSDFilesDialog(
                self,
                school_folder=school_folder,
                preselected=self.main_window.selected_files,
            )
            if pick_dlg.exec() != QDialog.DialogCode.Accepted:
                return

            self.main_window.selected_files = pick_dlg.selected_files
            self.main_window.lbl_count_info.setText(
                f"PSD файлов выбрано: {len(pick_dlg.selected_files)}"
            )
            try:
                self.main_window.append_log(
                    f"__INFO__Выбрано для BRIGHT: "
                    f"{len(pick_dlg.selected_files)} PSD-файлов"
                )
            except Exception:
                pass

        self.input_set.setPlainText(preset_set)
        self.input_actions.setPlainText(preset_actions)

    def _validate(self):
        set_ok = bool(self.input_set.toPlainText().strip())
        act_ok = bool(self.input_actions.toPlainText().strip())
        if set_ok and act_ok:
            self.lbl_status.setText("✅ Готово к запуску")
            self.lbl_status.setStyleSheet(
                "color: #66DD66; border: none; font-size: 11px; font-weight: bold;"
            )
        else:
            self.lbl_status.setText("❌ Заполните Action Set и Actions")
            self.lbl_status.setStyleSheet(
                "color: #FF5555; border: none; font-size: 11px; font-weight: bold;"
            )
        self.btn_start.setEnabled(set_ok and act_ok)

    def _on_start(self):
        self.action_set = self.input_set.toPlainText().strip()
        self.actions_text = self.input_actions.toPlainText().strip()
        self.fast_mode = self.chk_fast.isChecked()
        self.auto_alert = self.chk_alert.isChecked()
        self.accept()


# ----------------------------------------------------------------------
# ПИКСЕЛЬНЫЙ ЗНАЧОК
# ----------------------------------------------------------------------
class PixelIconWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(52, 40)
        self.shape_timer = QTimer(self)
        self.shape_timer.setInterval(2000)
        self.shape_timer.timeout.connect(self.next_pattern)
        self.anim_timer = QTimer(self)
        self.anim_timer.setInterval(30)
        self.anim_timer.timeout.connect(self._anim_step)
        self.current_index = 0
        self.anim_phase = 0.0
        self.states = [
            {"type": "square", "color": QColor("#FF8C00"),
             "pattern": [[1,1,1,0,0],[1,1,1,0,0],[1,1,1,0,0],[0,0,0,0,0],[0,0,0,0,0]]},
            {"type": "diamond", "color": QColor("#4169E1"),
             "pattern": [[0,0,1,0,0],[0,1,1,1,0],[1,1,1,1,1],[0,1,1,1,0],[0,0,1,0,0]]},
            {"type": "diagonal", "color": QColor("#C71585"),
             "pattern": [[1,1,0,0,0],[1,1,1,0,0],[0,1,1,1,0],[0,0,1,1,1],[0,0,0,1,1]]},
            {"type": "circle", "color": QColor("#00C853"),
             "pattern": [[0,0,1,0,0],[0,1,1,1,0],[1,1,1,1,1],[0,1,1,1,0],[0,0,1,0,0]]},
            {"type": "diamond", "color": QColor("#FFD700"),
             "pattern": [[0,0,1,0,0],[0,0,1,0,0],[1,1,1,1,1],[0,0,1,0,0],[0,0,1,0,0]]},
            {"type": "diagonal", "color": QColor("#8A2BE2"),
             "pattern": [[1,1,1,0,0],[1,0,0,0,0],[1,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0]]},
        ]

    def start(self):
        self.shape_timer.start()
        self.anim_timer.start()

    def stop(self):
        self.shape_timer.stop()
        self.anim_timer.stop()
        self.anim_phase = 0.0
        self.update()

    def next_pattern(self):
        self.current_index = (self.current_index + 1) % len(self.states)
        self.anim_phase = 0.0
        self.update()

    def _anim_step(self):
        self.anim_phase = (self.anim_phase + 0.02) % 1.0
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pixel_size = 5
        gap = 1
        icon_w = 5 * pixel_size + 4 * gap
        icon_h = 5 * pixel_size + 4 * gap
        base_offset_x = 6
        base_offset_y = (self.height() - icon_h) // 2

        if not self.shape_timer.isActive():
            color = QColor("#555555")
            pattern = [
                [0,0,1,0,0],[0,0,1,0,0],[1,1,1,1,1],[0,0,1,0,0],[0,0,1,0,0],
            ]
            self._draw_pattern(
                painter, pattern, color, base_offset_x, base_offset_y,
                pixel_size, gap, scale=1.0, alpha=255, rotation=0.0, reveal=1.0,
            )
            return

        state = self.states[self.current_index]
        base_color = state["color"]
        pattern = state["pattern"]
        anim_type = state["type"]
        phase = self.anim_phase
        color = QColor(base_color.red(), base_color.green(), base_color.blue(), 255)

        if anim_type == "square":
            scale = 0.9 + 0.25 * math.sin(phase * 2 * math.pi)
            alpha = 180 + int(75 * (0.5 + 0.5 * math.sin(phase * 2 * math.pi)))
            self._draw_pattern(painter, pattern, color, base_offset_x, base_offset_y,
                               pixel_size, gap, scale=scale, alpha=alpha,
                               rotation=0.0, reveal=1.0)
        elif anim_type == "diamond":
            rotation = phase * 360.0
            alpha = 180 + int(75 * abs(math.cos(phase * 2 * math.pi)))
            self._draw_pattern(painter, pattern, color, base_offset_x, base_offset_y,
                               pixel_size, gap, scale=1.0, alpha=alpha,
                               rotation=rotation, reveal=1.0)
        elif anim_type == "diagonal":
            reveal = phase * 2.0 if phase < 0.5 else (1.0 - phase) * 2.0
            self._draw_pattern(painter, pattern, color, base_offset_x, base_offset_y,
                               pixel_size, gap, scale=1.0, alpha=255,
                               rotation=0.0, reveal=reveal)
        elif anim_type == "circle":
            pulse = 0.5 + 0.5 * math.sin(phase * 2 * math.pi * 2)
            alpha = 150 + int(105 * pulse)
            scale = 0.9 + 0.15 * pulse
            self._draw_pattern(painter, pattern, color, base_offset_x, base_offset_y,
                               pixel_size, gap, scale=scale, alpha=alpha,
                               rotation=0.0, reveal=1.0)

    def _draw_pattern(self, painter, pattern, color, base_offset_x, base_offset_y,
                      pixel_size, gap, scale, alpha, rotation, reveal):
        icon_w = 5 * pixel_size + 4 * gap
        icon_h = 5 * pixel_size + 4 * gap
        painter.save()
        cx = base_offset_x + icon_w / 2
        cy = base_offset_y + icon_h / 2
        painter.translate(cx, cy)
        painter.rotate(rotation)
        painter.scale(scale, scale)
        painter.translate(-cx, -cy)
        draw_color = QColor(color.red(), color.green(), color.blue(), alpha)
        pixels = []
        total = 0
        for r in range(5):
            for c in range(5):
                if pattern[r][c] == 1:
                    pixels.append((r, c))
                    total += 1
        visible_count = int(total * reveal)
        if alpha > 30 and visible_count > 0:
            glow_alpha = int(alpha * 0.5)
            for layer in range(2):
                expand = (layer + 1) * 3
                layer_alpha = int(glow_alpha * (0.5 - layer * 0.2))
                if layer_alpha <= 0:
                    continue
                glow_color = QColor(color.red(), color.green(), color.blue(), layer_alpha)
                painter.setBrush(glow_color)
                painter.setPen(Qt.PenStyle.NoPen)
                for idx, (r, c) in enumerate(pixels):
                    if idx >= visible_count:
                        break
                    painter.drawEllipse(
                        base_offset_x + c * (pixel_size + gap) - expand // 2,
                        base_offset_y + r * (pixel_size + gap) - expand // 2,
                        pixel_size + expand, pixel_size + expand,
                    )
            for idx, (r, c) in enumerate(pixels):
                if idx >= visible_count:
                    break
                px = base_offset_x + c * (pixel_size + gap) + pixel_size / 2
                py = base_offset_y + r * (pixel_size + gap) + pixel_size / 2
                rad = QRadialGradient(QPointF(px, py), pixel_size * 2.0)
                rad.setColorAt(0.0, QColor(
                    min(255, color.red() + 50), min(255, color.green() + 50),
                    min(255, color.blue() + 50), int(alpha * 0.5)))
                rad.setColorAt(0.5, QColor(
                    color.red(), color.green(), color.blue(), int(alpha * 0.25)))
                rad.setColorAt(1.0, QColor(
                    color.red(), color.green(), color.blue(), 0))
                painter.setBrush(rad)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawEllipse(QPointF(px, py), pixel_size * 2.0, pixel_size * 2.0)
        painter.setBrush(draw_color)
        painter.setPen(Qt.PenStyle.NoPen)
        for idx, (r, c) in enumerate(pixels):
            if idx >= visible_count:
                break
            painter.fillRect(
                base_offset_x + c * (pixel_size + gap),
                base_offset_y + r * (pixel_size + gap),
                pixel_size, pixel_size, draw_color,
            )
        painter.restore()


# ----------------------------------------------------------------------
# АНИМИРОВАННЫЙ PROGRESS BAR
# ----------------------------------------------------------------------
class YellowGlowProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(14)
        self.setTextVisible(False)
        self.anim_offset = 0
        self.timer = QTimer(self)
        self.timer.setInterval(40)
        self.timer.timeout.connect(self.update_anim)

    def start_anim(self):
        self.timer.start()

    def stop_anim(self):
        self.timer.stop()
        self.update()

    def update_anim(self):
        self.anim_offset = (self.anim_offset + 3) % 40
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = self.rect()
        painter.setBrush(QColor("#1A1A1A"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(rect, 7, 7)
        if self.maximum() <= 0:
            return
        progress_width = int(rect.width() * (self.value() / self.maximum()))
        if progress_width <= 0:
            return
        fill_rect = QRectF(0, 0, progress_width, rect.height())
        gradient = QLinearGradient(0, 0, rect.width(), 0)
        gradient.setColorAt(0.0, QColor("#00A8FF"))
        gradient.setColorAt(0.5, QColor("#00C9A7"))
        gradient.setColorAt(1.0, QColor("#4DE8B1"))
        painter.setBrush(gradient)
        painter.drawRoundedRect(fill_rect, 7, 7)
        if self.timer.isActive():
            painter.setClipRect(fill_rect)
            glow_x = ((self.anim_offset * 10) % (rect.width() + 100)) - 50
            glow_grad = QLinearGradient(glow_x, 0, glow_x + 40, 0)
            glow_grad.setColorAt(0.0, QColor(255, 255, 255, 0))
            glow_grad.setColorAt(0.5, QColor(255, 255, 255, 180))
            glow_grad.setColorAt(1.0, QColor(255, 255, 255, 0))
            painter.setBrush(glow_grad)
            painter.drawRect(fill_rect)


# ----------------------------------------------------------------------
# ОКНО "О ПРОГРАММЕ"
# ----------------------------------------------------------------------
class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("О программе — Photoshop WorkStation+")
        self.setFixedSize(460, 440)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        title_lbl = QLabel("Photoshop WorkStation+ (v4.0)")
        title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title_lbl.setStyleSheet("color: #00BFFF; border: none;")
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_lbl)
        info_text = """
        <h3>Назначение программы:</h3>
        <p>Профессиональный программный комплекс для автоматизации
        пакетной обработки школьных фотографий, работы с RAW-фильтрами,
        наложения рамок, интеграции QR-кодов и быстрого экспорта
        через Adobe Photoshop.</p>
        <p><b>Версия:</b> v4.0</p>
        <p><b>Год выпуска:</b> 2026</p>
        <p><b>Разработчик:</b> DirtSmoke44 & JohnnySuon</p>
        """
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(info_text)
        browser.setStyleSheet(
            "QTextBrowser { background-color: transparent;"
            "border: 1px solid #333333; border-radius: 8px;"
            "padding: 10px; font-size: 12px; color: #FFFFFF; }"
        )
        layout.addWidget(browser)
        btn_close = QPushButton("Закрыть")
        btn_close.setStyleSheet(
            "QPushButton { background-color: #007ACC; color: white;"
            "border: none; border-radius: 8px; padding: 8px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #005999; }"
        )
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


# ----------------------------------------------------------------------
# ДИАЛОГ QR-СКРИПТА
# ----------------------------------------------------------------------
class QRDialog(QDialog):
    def __init__(self, parent=None, default_frame="", default_qr_folder="",
                 fallback_qr_root=""):
        super().__init__(parent)
        self.setWindowTitle("Параметры импорта QR-кодов")
        self.setFixedSize(760, 400)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.frame_path = default_frame
        self.qr_folder_path = default_qr_folder
        self.fallback_qr_root = fallback_qr_root

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(12)

        title = QLabel("Настройка импорта QR-кодов в рамку 20x30")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setStyleSheet("color: #00BFFF; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        lbl_frame = QLabel("1. Файл рамки 20x30.psd:")
        lbl_frame.setStyleSheet("color: #CCFF00; font-weight: bold; border: none;")
        layout.addWidget(lbl_frame)

        row_frame = QHBoxLayout()
        self.lbl_frame_val = QLabel(self.frame_path or "Файл не выбран")
        self.lbl_frame_val.setWordWrap(True)
        self.lbl_frame_val.setStyleSheet(
            "QLabel { color: #EEEEEE; background-color: #1A1A1A;"
            "border: 1px solid #333333; border-radius: 6px;"
            "padding: 8px; font-size: 11px; }"
        )
        row_frame.addWidget(self.lbl_frame_val, 1)
        btn_pick_frame = QPushButton("Выбрать файл")
        btn_pick_frame.setFixedWidth(120)
        btn_pick_frame.setStyleSheet(
            "QPushButton { background-color: #ff0095; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; }"
            "QPushButton:hover { background-color: #ff9ed7; }"
        )
        btn_pick_frame.clicked.connect(self._pick_frame)
        row_frame.addWidget(btn_pick_frame)
        layout.addLayout(row_frame)

        self.lbl_frame_status = QLabel()
        self.lbl_frame_status.setStyleSheet("border: none; font-size: 11px;")
        layout.addWidget(self.lbl_frame_status)

        lbl_qr = QLabel("2. Папка с QR-кодами (папка 'коды' внутри школы):")
        lbl_qr.setStyleSheet("color: #CCFF00; font-weight: bold; border: none;")
        layout.addWidget(lbl_qr)

        row_qr = QHBoxLayout()
        self.lbl_qr_val = QLabel(self.qr_folder_path or "Папка не выбрана")
        self.lbl_qr_val.setWordWrap(True)
        self.lbl_qr_val.setStyleSheet(
            "QLabel { color: #EEEEEE; background-color: #1A1A1A;"
            "border: 1px solid #333333; border-radius: 6px;"
            "padding: 8px; font-size: 11px; }"
        )
        row_qr.addWidget(self.lbl_qr_val, 1)
        btn_pick_qr = QPushButton("Выбрать папку")
        btn_pick_qr.setFixedWidth(120)
        btn_pick_qr.setStyleSheet(
            "QPushButton { background-color: #007ACC; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; }"
            "QPushButton:hover { background-color: #005999; }"
        )
        btn_pick_qr.clicked.connect(self._pick_qr_folder)
        row_qr.addWidget(btn_pick_qr)
        layout.addLayout(row_qr)

        self.lbl_qr_status = QLabel()
        self.lbl_qr_status.setStyleSheet("border: none; font-size: 11px;")
        layout.addWidget(self.lbl_qr_status)
        layout.addStretch()

        btn_row = QHBoxLayout()
        self.btn_start = QPushButton("Запустить импорт")
        self.btn_start.setStyleSheet(
            "QPushButton { background-color: #7af6ff; color: black;"
            "border: none; border-radius: 15px; padding: 10px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #11848c; }"
            "QPushButton:disabled { background-color: #333333; color: #777777; }"
        )
        self.btn_start.clicked.connect(self.accept)
        btn_row.addWidget(self.btn_start, 1)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #D32F2F; color: white;"
            "border: none; border-radius: 15px; padding: 10px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #B71C1C; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel, 1)
        layout.addLayout(btn_row)
        self._update_statuses()

    def _pick_frame(self):
        start_dir = os.path.dirname(self.frame_path) if self.frame_path else ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл рамки 20x30.psd", start_dir, "Photoshop PSD (*.psd)"
        )
        if path:
            self.frame_path = path
            self.lbl_frame_val.setText(path)
            self._update_statuses()

    def _pick_qr_folder(self):
        if self.qr_folder_path and os.path.isdir(self.qr_folder_path):
            start_dir = self.qr_folder_path
        elif self.fallback_qr_root and os.path.isdir(self.fallback_qr_root):
            start_dir = self.fallback_qr_root
        else:
            start_dir = ""
        path = QFileDialog.getExistingDirectory(
            self, "Выберите папку 'коды' внутри школы", start_dir
        )
        if path:
            self.qr_folder_path = path
            self.lbl_qr_val.setText(path)
            self._update_statuses()

    def _update_statuses(self):
        if self.frame_path and os.path.exists(self.frame_path):
            self.lbl_frame_status.setText("✅ Файл рамки выбран и существует")
            self.lbl_frame_status.setStyleSheet(
                "color: #66DD66; font-weight: bold; font-size: 11px; border: none;"
            )
            frame_ok = True
        else:
            self.lbl_frame_status.setText("❌ Файл рамки не выбран или не существует")
            self.lbl_frame_status.setStyleSheet(
                "color: #FF5555; font-weight: bold; font-size: 11px; border: none;"
            )
            frame_ok = False
        if self.qr_folder_path and os.path.isdir(self.qr_folder_path):
            self.lbl_qr_status.setText("✅ Папка выбрана и существует")
            self.lbl_qr_status.setStyleSheet(
                "color: #66DD66; font-weight: bold; font-size: 11px; border: none;"
            )
            qr_ok = True
        else:
            self.lbl_qr_status.setText("❌ Папка не выбрана или не существует")
            self.lbl_qr_status.setStyleSheet(
                "color: #FF5555; font-weight: bold; font-size: 11px; border: none;"
            )
            qr_ok = False
        self.btn_start.setEnabled(frame_ok and qr_ok)


# ----------------------------------------------------------------------
# ДИАЛОГ ВЫБОРА НАБОРА ПУТЕЙ ПРИ СБРОСЕ
# ----------------------------------------------------------------------
class ResetPathsDialog(QDialog):
    def __init__(self, parent=None, default_paths=None, alternate_paths=None):
        super().__init__(parent)
        self.setWindowTitle("Сброс к стандартным путям")
        self.setFixedSize(900, 480)
        self.setSizeGripEnabled(False)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.default_paths = dict(default_paths or {})
        self.alternate_paths = dict(alternate_paths or {})
        self.chosen_set = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)
        layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)

        title = QLabel("Выберите набор путей для сброса")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #00BFFF; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        hint = QLabel(
            "У вас есть два предустановленных набора путей.\n"
            "Выберите, к какому из них сбросить текущие значения."
        )
        hint.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setWordWrap(True)
        layout.addWidget(hint)
        layout.addSpacing(6)

        card_style = (
            "QFrame { background-color: #1A1A1A; border: 1px solid #333333;"
            "border-radius: 10px; }"
            "QFrame:hover { border: 1px solid #00BFFF; }"
        )
        label_style = "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
        value_style = (
            "QLabel { color: #EEEEEE; background-color: #101010;"
            "border: 1px solid #2A2A2A; border-radius: 6px;"
            "padding: 6px 8px; font-size: 10px; font-family: Consolas; }"
        )

        card_default = QFrame()
        card_default.setStyleSheet(card_style)
        cd_layout = QVBoxLayout(card_default)
        cd_layout.setContentsMargins(15, 12, 15, 12)
        cd_layout.setSpacing(6)
        cd_title = QLabel("DirtSmoke44")
        cd_title.setStyleSheet(
            "color: #7CFFCB; font-weight: bold; font-size: 12px; border: none;"
        )
        cd_layout.addWidget(cd_title)
        cd_sub = QLabel("(Первый набор)")
        cd_sub.setStyleSheet("color: #888888; font-size: 10px; border: none;")
        cd_layout.addWidget(cd_sub)
        for key in ("schools_dir", "frame_dir", "scripts_dir"):
            row = QHBoxLayout()
            row.setSpacing(8)
            key_lbl = QLabel(self._key_to_label(key))
            key_lbl.setStyleSheet(label_style)
            key_lbl.setFixedWidth(90)
            row.addWidget(key_lbl)
            val_lbl = QLabel(self.default_paths.get(key, "—") or "—")
            val_lbl.setStyleSheet(value_style)
            val_lbl.setWordWrap(False)
            val_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            row.addWidget(val_lbl, 1)
            cd_layout.addLayout(row)
        cd_layout.addSpacing(4)
        btn_default = QPushButton("Назначить")
        btn_default.setStyleSheet(
            "QPushButton { background-color: #2E7D32; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #1B5E20; }"
        )
        btn_default.setMinimumHeight(30)
        btn_default.clicked.connect(self._choose_default)
        cd_layout.addWidget(btn_default)
        layout.addWidget(card_default)
        layout.addSpacing(6)

        card_alt = QFrame()
        card_alt.setStyleSheet(card_style)
        ca_layout = QVBoxLayout(card_alt)
        ca_layout.setContentsMargins(15, 12, 15, 12)
        ca_layout.setSpacing(6)
        ca_title = QLabel("JohnnySuon")
        ca_title.setStyleSheet(
            "color: #FFD700; font-weight: bold; font-size: 12px; border: none;"
        )
        ca_layout.addWidget(ca_title)
        ca_sub = QLabel("(Второй набор)")
        ca_sub.setStyleSheet("color: #888888; font-size: 10px; border: none;")
        ca_layout.addWidget(ca_sub)
        for key in ("schools_dir", "frame_dir", "scripts_dir"):
            row = QHBoxLayout()
            row.setSpacing(8)
            key_lbl = QLabel(self._key_to_label(key))
            key_lbl.setStyleSheet(label_style)
            key_lbl.setFixedWidth(90)
            row.addWidget(key_lbl)
            val_lbl = QLabel(self.alternate_paths.get(key, "—") or "—")
            val_lbl.setStyleSheet(value_style)
            val_lbl.setWordWrap(False)
            val_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            row.addWidget(val_lbl, 1)
            ca_layout.addLayout(row)
        ca_layout.addSpacing(4)
        btn_alt = QPushButton("Назначить")
        btn_alt.setStyleSheet(
            "QPushButton { background-color: #0b9dd6; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #8B6508; }"
        )
        btn_alt.setMinimumHeight(30)
        btn_alt.clicked.connect(self._choose_alternate)
        ca_layout.addWidget(btn_alt)
        layout.addWidget(card_alt)
        layout.addSpacing(10)

        btn_cancel = QPushButton("Отмена (оставить текущие пути)")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #D32F2F; color: #FFFFFF;"
            "border: 1px solid #444444; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #444444; }"
        )
        btn_cancel.setMinimumHeight(32)
        btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_cancel)

    @staticmethod
    def _key_to_label(key):
        return {"schools_dir": "Школы:", "frame_dir": "Рамки:",
                "scripts_dir": "Скрипты:"}.get(key, key + ":")

    def _choose_default(self):
        self.chosen_set = "default"
        self.accept()

    def _choose_alternate(self):
        self.chosen_set = "alternate"
        self.accept()


# ----------------------------------------------------------------------
# ДИАЛОГ НАЗНАЧЕНИЯ ПУТЕЙ
# ----------------------------------------------------------------------
class PathSettingsDialog(QDialog):
    def __init__(self, parent=None, current_paths=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setFixedSize(820, 560)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.paths = dict(current_paths or {})

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        title_lbl = QLabel("Настройка рабочих путей")
        title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title_lbl.setStyleSheet("color: #00BFFF; border: none;")
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_lbl)

        hint_lbl = QLabel(
            "Укажите пути к папкам, которые программа использует для работы.\n"
            "Изменения сохраняются автоматически в файл config.json."
        )
        hint_lbl.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        hint_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint_lbl.setWordWrap(True)
        layout.addWidget(hint_lbl)
        layout.addSpacing(8)

        self.fields = [
            ("schools_dir", "1. Папка со школами (рабочая директория 2026):", "dir", None),
            ("frame_dir", "2. Папка с рамками (Рамка с плотными слоями):", "dir", None),
            ("scripts_dir", "3. Папка со скриптами (!MY_SCRIPTS):", "dir", None),
            ("ambient_video", "4. Видео для фона (MP4, необязательно):",
             "file_mp4", "Видеофайл (*.mp4 *.mov *.avi *.mkv)"),
        ]
        self.value_labels = {}

        for key, caption, ftype, _ in self.fields:
            lbl = QLabel(caption)
            lbl.setStyleSheet(
                "color: #CCFF00; font-weight: bold; font-size: 11px; border: none;"
            )
            layout.addWidget(lbl)
            row = QHBoxLayout()
            row.setSpacing(8)
            val_lbl = QLabel(self.paths.get(key, "") or "Путь не задан")
            val_lbl.setWordWrap(True)
            val_lbl.setStyleSheet(
                "QLabel { color: #EEEEEE; background-color: #1A1A1A;"
                "border: 1px solid #333333; border-radius: 6px;"
                "padding: 8px; font-size: 11px; font-family: Consolas; }"
            )
            val_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            row.addWidget(val_lbl, 1)
            btn_pick = QPushButton("Выбрать папку")
            btn_pick.setFixedWidth(130)
            btn_pick.setStyleSheet(
                "QPushButton { background-color: #69d6d3; color: black;"
                "border: none; border-radius: 13px; padding: 8px;"
                "font-weight: bold; font-size: 11px; }"
                "QPushButton:hover { background-color: #009994; }"
            )
            btn_pick.clicked.connect(
                lambda checked=False, k=key, t=ftype: self._pick_path(k, t)
            )
            row.addWidget(btn_pick)
            layout.addLayout(row)
            self.value_labels[key] = val_lbl

        layout.addStretch()
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_defaults = QPushButton("Сменить профиль")
        btn_defaults.setStyleSheet(
            "QPushButton { background-color: #ffa8ef; color: black;"
            "border: none; border-radius: 15px; padding: 9px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #E78AD8; }"
        )
        btn_defaults.clicked.connect(self._reset_to_defaults)
        btn_row.addWidget(btn_defaults, 1)

        btn_save = QPushButton("Сохранить")
        btn_save.setStyleSheet(
            "QPushButton { background-color: #a19eff; color: black;"
            "border: none; border-radius: 15px; padding: 9px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #8582D9; }"
        )
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save, 1)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #ffbb3d; color: black;"
            "border: none; border-radius: 15px; padding: 9px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #E0A52F; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel, 1)
        layout.addLayout(btn_row)

    def _pick_path(self, key, ftype):
        current = self.paths.get(key, "")
        if ftype == "dir":
            start_dir = current if os.path.isdir(current) else ""
            path = QFileDialog.getExistingDirectory(self, f"Выберите папку: {key}", start_dir)
        elif ftype == "file_mp4":
            start_dir = os.path.dirname(current) if current else ""
            path, _ = QFileDialog.getOpenFileName(
                self, "Выберите видеофайл", start_dir,
                "Видео (*.mp4 *.mov *.avi *.mkv);;Все файлы (*.*)"
            )
        else:
            start_dir = os.path.dirname(current) if current else ""
            path, _ = QFileDialog.getOpenFileName(self, f"Выберите файл: {key}",
                                                   start_dir, "Все файлы (*.*)")
        if path:
            self.paths[key] = path
            self.value_labels[key].setText(path)

    def _reset_to_defaults(self):
        default_paths = dict(ModernPhotoshopWorkstation.DEFAULT_PATHS)
        alternate_paths = dict(getattr(ModernPhotoshopWorkstation, "ALTERNATE_PATHS", {}))
        choose_dlg = ResetPathsDialog(self, default_paths=default_paths,
                                      alternate_paths=alternate_paths)
        if choose_dlg.exec() != QDialog.DialogCode.Accepted:
            return
        if choose_dlg.chosen_set == "default":
            target = default_paths
        elif choose_dlg.chosen_set == "alternate":
            target = alternate_paths
        else:
            return
        for key, val in target.items():
            if key in self.paths:
                self.paths[key] = val
                self.value_labels[key].setText(val or "Путь не задан")


# ----------------------------------------------------------------------
# УТИЛИТА РАБОТЫ С КОНФИГОМ
# ----------------------------------------------------------------------
def get_config_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


def load_config():
    path = get_config_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Ошибка чтения конфига: {e}")
        return {}


def save_config(data):
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Ошибка сохранения конфига: {e}")
        return False


# ----------------------------------------------------------------------
# ПОТОК ОБРАБОТКИ
# ----------------------------------------------------------------------
class WorkerThread(QThread):
    progress_signal = pyqtSignal(int, int)
    file_status_signal = pyqtSignal(str, str)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(int, str)

    def __init__(self, mode, files, scripts_dir, script_filename, script_title,
                 school_name, is_visible=False, qr_frame_path=None,
                 qr_folder_path=None, action_set=None, actions_list=None,
                 fast_mode=True, auto_alert=True, frame_1520_path=None,
                 frame_2030_path=None):
        super().__init__()
        self.mode = mode
        self.files = files
        self.scripts_dir = scripts_dir
        self.script_filename = script_filename
        self.script_title = script_title
        self.school_name = school_name
        self.is_visible = False
        self.qr_frame_path = qr_frame_path
        self.qr_folder_path = qr_folder_path
        self.action_set = action_set
        self.actions_list = actions_list
        self.fast_mode = fast_mode
        self.auto_alert = auto_alert
        self.is_cancelled = False
        self.ps_app = None
        self.frame_1520_path = frame_1520_path
        self.frame_2030_path = frame_2030_path

    def _build_jsx_path(self):
        original = os.path.join(self.scripts_dir, self.script_filename)
        if self.mode == "qr" and self.qr_frame_path and self.qr_folder_path:
            if not os.path.exists(original):
                return original
            try:
                with open(original, "r", encoding="utf-8") as f:
                    content = f.read()
                def esc(p):
                    return p.replace("\\", "\\\\")
                content = content.replace("__PSD_FRAME_PATH__", esc(self.qr_frame_path))
                content = content.replace("__QR_FOLDER_PATH__", esc(self.qr_folder_path))
                content = content.replace(
                    "__INITIAL_PATH__", esc(os.path.dirname(self.qr_folder_path))
                )
                tmp_path = os.path.join(
                    tempfile.gettempdir(),
                    f"_qr_temp_{os.getpid()}_{id(self)}.jsx",
                )
                with open(tmp_path, "w", encoding="utf-8") as f:
                    f.write(content)
                return tmp_path
            except Exception as e:
                self.log_signal.emit(f"Ошибка генерации временного JSX: {e}")
                return original
        if self.mode == "smart_frames":
            if not os.path.exists(original):
                self.log_signal.emit(
                    f"__ERROR__Файл {self.script_filename} не найден в {self.scripts_dir}"
                )
            return original
        return original

    def _build_smart_jsx_for_file(self, template_path, psd_path, kind):
        try:
            with open(template_path, "r", encoding="utf-8") as f:
                content = f.read()
            def esc_path(p):
                return str(p).replace("\\", "\\\\").replace('"', '\\"')
            def esc_str(s):
                return (str(s).replace("\\", "\\\\").replace('"', '\\"')
                        .replace("\n", "\\n").replace("\r", ""))
            crop_jsx_1520 = os.path.join(self.scripts_dir, "13_SetCrop1520.jsx")
            crop_jsx_2030 = os.path.join(self.scripts_dir, "12_SetCrop2030.jsx")
            content = content.replace("__PSD_PATH__", esc_path(psd_path))
            content = content.replace("__FRAME_1520__", esc_path(self.frame_1520_path or ""))
            content = content.replace("__FRAME_2030__", esc_path(self.frame_2030_path or ""))
            content = content.replace("__CROP_JSX_1520__", esc_path(crop_jsx_1520))
            content = content.replace("__CROP_JSX_2030__", esc_path(crop_jsx_2030))
            content = content.replace("__TARGET_KIND__", esc_str(kind))
            import hashlib
            h = hashlib.md5((psd_path + kind).encode("utf-8", errors="ignore")).hexdigest()[:8]
            tmp_path = os.path.join(
                tempfile.gettempdir(),
                f"_smart_frames_{os.getpid()}_{id(self)}_{h}.jsx",
            )
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(content)
            return tmp_path
        except Exception as e:
            self.log_signal.emit(f"__ERROR__Не удалось собрать JSX для {psd_path}: {e}")
            return None

    def run(self):
        try:
            try:
                self.file_status_signal.emit("Подключение к Photoshop...", "")
                self.ps_app = win32com.client.Dispatch("Photoshop.Application")
                self.ps_app.DisplayDialogs = 3
                self.ps_app.Visible = False
            except Exception as e:
                self.log_signal.emit(f"__ERROR__Ошибка подключения к Photoshop:\n{e}")
                self.finished_signal.emit(0, self.script_title)
                return

            if self.mode == "batch_actions":
                self._run_batch_actions_com()
                return
            if self.mode == "smart_frames":
                self._run_smart_frames()
                return

            jsx_path = self._build_jsx_path()
            if not os.path.exists(jsx_path):
                self.log_signal.emit(
                    f"__ERROR__Скрипт {self.script_filename} не найден!\n"
                    f"Проверьте папку: {self.scripts_dir}"
                )
                self.finished_signal.emit(0, self.script_title)
                return

            if self.mode == "qr":
                try:
                    self.file_status_signal.emit("Импорт QR-кодов...", "Выполнение...")
                    self.progress_signal.emit(50, 100)
                    self.ps_app.DoJavaScriptFile(jsx_path)
                    self.progress_signal.emit(100, 100)
                    self.finished_signal.emit(1, self.script_title)
                except Exception as e:
                    self.log_signal.emit(f"__ERROR__Ошибка выполнения: {e}")
                    self.finished_signal.emit(0, self.script_title)
                return

            total = len(self.files)
            processed_count = 0
            self.log_signal.emit("--- НАЧАЛО ОБРАБОТКИ ---")

            for idx, psd_path in enumerate(self.files, start=1):
                if self.is_cancelled:
                    break
                file_name = os.path.basename(psd_path)
                parent_dir = os.path.basename(os.path.dirname(psd_path))
                self.file_status_signal.emit(f"Обработка [{idx}/{total}]", file_name)
                try:
                    doc = self.ps_app.Open(psd_path)
                    self.ps_app.DoJavaScriptFile(jsx_path)
                    doc.Close(1)
                    processed_count += 1
                    self.log_signal.emit(
                        f"[{idx}/{total}] Обработан: {file_name}\n  (Папка: {parent_dir})"
                    )
                except Exception as e:
                    self.log_signal.emit(f"[{idx}/{total}] ОШИБКА: {file_name}\n  ({e})")
                self.progress_signal.emit(idx, total)

            if not self.is_cancelled:
                self.finished_signal.emit(processed_count, self.script_title)

        except Exception as fatal:
            self.log_signal.emit(f"__ERROR__КРИТИЧЕСКАЯ ОШИБКА ПОТОКА: {fatal}")
            try:
                self.finished_signal.emit(0, self.script_title)
            except Exception:
                pass

    def _run_smart_frames(self):
        total = len(self.files)
        processed_count = 0
        errors_count = 0
        skipped_count = 0
        self.log_signal.emit("--- НАЧАЛО УМНОЙ ВСТАВКИ РАМОК ---")
        self.log_signal.emit(f"Файлов: {total}")

        for kind, path in (("1520", self.frame_1520_path), ("2030", self.frame_2030_path)):
            if path and os.path.exists(path):
                self.log_signal.emit(f"__INFO__Рамка {kind}: {path}")
            else:
                self.log_signal.emit(f"__WARN__Рамка {kind} не найдена: {path}")

        template_path = os.path.join(self.scripts_dir, self.script_filename)
        if not os.path.exists(template_path):
            self.log_signal.emit(f"__ERROR__Не найден шаблон JSX: {template_path}")
            self.finished_signal.emit(0, self.script_title)
            return

        for idx, psd_path in enumerate(self.files, start=1):
            if self.is_cancelled:
                break
            file_name = os.path.basename(psd_path)
            parent_dir = os.path.basename(os.path.dirname(psd_path)).lower()
            if "1520" in parent_dir:
                kind = "1520"
            elif "2030" in parent_dir:
                kind = "2030"
            else:
                self.log_signal.emit(
                    f"[{idx}/{total}] __WARN__Пропуск (нет 1520/2030): {file_name}"
                )
                skipped_count += 1
                self.progress_signal.emit(idx, total)
                continue
            self.file_status_signal.emit(f"Умная вставка [{idx}/{total}] ({kind})", file_name)
            try:
                _ = self.ps_app.Version
            except Exception:
                self.log_signal.emit("__ERROR__Photoshop перестал отвечать. Прерывание.")
                break
            jsx_for_file = self._build_smart_jsx_for_file(template_path, psd_path, kind)
            if not jsx_for_file:
                errors_count += 1
                self.progress_signal.emit(idx, total)
                continue
            try:
                self.ps_app.DoJavaScriptFile(jsx_for_file)
                processed_count += 1
                self.log_signal.emit(f"[{idx}/{total}] Обработан ({kind}): {file_name}")
            except Exception as e:
                errors_count += 1
                self.log_signal.emit(f"[{idx}/{total}] ОШИБКА ({kind}): {file_name}\n  ({e})")
            finally:
                try:
                    os.remove(jsx_for_file)
                except Exception:
                    pass
            self.progress_signal.emit(idx, total)

        if not self.is_cancelled:
            self.log_signal.emit(
                f"__INFO__Итого: обработано {processed_count}, "
                f"пропущено {skipped_count}, ошибок {errors_count}"
            )
            self.finished_signal.emit(processed_count, self.script_title)

    def _run_batch_actions_com(self):
        total = len(self.files)
        processed_count = 0
        errors_count = 0
        self.log_signal.emit("--- НАЧАЛО BATCH ACTIONS (COM) ---")
        self.log_signal.emit(f"Action Set: {self.action_set}")
        self.log_signal.emit(f"Actions: {self.actions_list}")
        self.log_signal.emit(f"Файлов: {total}")

        actions = [a.strip() for a in (self.actions_list or "").split(",") if a.strip()]
        if not actions:
            self.log_signal.emit("__ERROR__Список Actions пуст.")
            self.finished_signal.emit(0, self.script_title)
            return

        for idx, psd_path in enumerate(self.files, start=1):
            if self.is_cancelled:
                break
            file_name = os.path.basename(psd_path)
            parent_dir = os.path.basename(os.path.dirname(psd_path))
            self.file_status_signal.emit(f"Batch Actions [{idx}/{total}]", file_name)
            try:
                _ = self.ps_app.Version
            except Exception:
                self.log_signal.emit("__ERROR__Photoshop перестал отвечать. Прерывание.")
                break
            doc = None
            try:
                doc = self.ps_app.Open(psd_path)
                try:
                    self.ps_app.ActiveDocument = doc
                except Exception:
                    pass
                self._select_top_layer_after_frames()
                for action_name in actions:
                    try:
                        self.ps_app.DoAction(action_name, self.action_set)
                    except Exception as act_err:
                        self.log_signal.emit(
                            f"  __WARN__action '{action_name}' не выполнен: {act_err}"
                        )
                try:
                    doc.Close(1)
                except Exception:
                    pass
                processed_count += 1
                self.log_signal.emit(
                    f"[{idx}/{total}] Обработан: {file_name}\n  (Папка: {parent_dir})"
                )
            except Exception as e:
                errors_count += 1
                self.log_signal.emit(f"[{idx}/{total}] ОШИБКА: {file_name}\n  ({e})")
                try:
                    if doc is not None:
                        doc.Close(1)
                except Exception:
                    pass
            self.progress_signal.emit(idx, total)

        if not self.is_cancelled:
            self.log_signal.emit(
                f"__INFO__Итого: обработано {processed_count} из {total}, "
                f"ошибок: {errors_count}"
            )
            self.finished_signal.emit(processed_count, self.script_title)

    def _select_top_layer_after_frames(self):
        try:
            doc = self.ps_app.ActiveDocument
        except Exception:
            return False
        FRAME_NAMES = {"рамка 2030", "рамка 1520"}

        def is_frame_group(layer_obj):
            try:
                _ = layer_obj.LayerSets
                name = (layer_obj.Name or "").strip().lower()
                return name in FRAME_NAMES
            except Exception:
                return False

        def get_top_layer_inside_group(group_obj):
            try:
                total = group_obj.Layers.Count
                if total == 0:
                    return None
                top = group_obj.Layers.Item(1)
                try:
                    _ = top.LayerSets
                    inner = get_top_layer_inside_group(top)
                    if inner is not None:
                        return inner
                    return top
                except Exception:
                    return top
            except Exception:
                return None

        try:
            total = doc.Layers.Count
            if total == 0:
                return False
            for i in range(1, total + 1):
                try:
                    layer = doc.Layers.Item(i)
                except Exception:
                    continue
                if is_frame_group(layer):
                    continue
                seen_frame_above = False
                for j in range(1, i):
                    try:
                        prev = doc.Layers.Item(j)
                        if is_frame_group(prev):
                            seen_frame_above = True
                            break
                    except Exception:
                        continue
                if not seen_frame_above:
                    continue
                try:
                    _ = layer.LayerSets
                    candidate = get_top_layer_inside_group(layer)
                    if candidate is not None:
                        doc.ActiveLayer = candidate
                        return True
                except Exception:
                    doc.ActiveLayer = layer
                    return True
        except Exception:
            pass
        return False

    def cancel(self):
        self.is_cancelled = True
        if self.ps_app:
            try:
                self.ps_app.Quit()
            except Exception:
                pass
        try:
            subprocess.run(
                ["taskkill", "/f", "/im", "Photoshop.exe"],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass


# ----------------------------------------------------------------------
# ГЛАВНОЕ ОКНО ПРИЛОЖЕНИЯ
# ----------------------------------------------------------------------
class ModernPhotoshopWorkstation(QMainWindow):

    DEFAULT_PATHS = {
        "schools_dir": (r"C:\Users\ahmad\OneDrive\Desktop"
                        r"\! Работа photoshop 2019 20.0.8\2026"),
        "frame_dir": (r"C:\Users\ahmad\OneDrive\Desktop"
                      r"\! Работа photoshop 2019 20.0.8"
                      r"\Рамка с плотными слоями"),
        "scripts_dir": (r"C:\Users\ahmad\OneDrive\Desktop"
                        r"\! Работа photoshop 2019 20.0.8"
                        r"\Actions, скрипты, camera raw и тд"
                        r"\!MY_SCRIPTS"),
    }
    ALTERNATE_PATHS = {
        "schools_dir": "E:/Работа",
        "frame_dir": "E:/Работа/Рамки",
        "scripts_dir": "E:/Работа/Скрипты",
    }
    FRAME_FILES = {"1520": "15x20.psd", "2030": "20x30.psd"}
    CROP_SCRIPTS = {"1520": "3_SetCrop1520.jsx", "2030": "4_SetCrop2030.jsx"}

    def __init__(self):
        super().__init__()
        self.frame_1520_path = ""
        self.frame_2030_path = ""
        self.ambient_video_path = ""
        self.setWindowTitle("Photoshop WorkStation+ v4.0")
        self.setFixedSize(1480, 760)

        icon_path = resource_path("iconapp_Photoshop Workstation+.png")
        if not os.path.exists(icon_path):
            icon_path = resource_path("iconapp_Photoshop_Workstation+.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        saved = load_config()
        self.default_schools_dir = saved.get("schools_dir", self.DEFAULT_PATHS["schools_dir"])
        self.default_frame_dir = saved.get("frame_dir", self.DEFAULT_PATHS["frame_dir"])
        self.scripts_dir = saved.get("scripts_dir", self.DEFAULT_PATHS["scripts_dir"])
        self.ambient_video_path = saved.get("ambient_video", "")
        self._refresh_frame_paths()
        self.selected_script_mode = None
        self.selected_script_file = None
        self.selected_script_title = None
        self.selected_script_visible = True
        self.selected_script_index = None

        self.selected_school_name = ""
        self.selected_folder = ""
        self.selected_files = []

        self.qr_frame_path = ""
        self.qr_folder_path = ""

        self.batch_action_set = ""
        self.batch_actions_text = ""
        self.batch_fast_mode = True
        self.batch_auto_alert = True

        self.last_script_used = None
        self.last_school_used = None
        self.worker = None

        self.media_player_ambient = None
        self.audio_output_ambient = None
        self.video_widget_ambient = None

        self.script_border_widgets = []
        self.active_script_border = None
        self.script_buttons = []

        self.init_ui()
        self.center_window()
        self.pixel_icon.start()

    def _refresh_frame_paths(self):
        self.frame_1520_path = ""
        self.frame_2030_path = ""
        if not os.path.isdir(self.default_frame_dir):
            return
        for candidate in ("15x20.psd", "15х20.psd", "15X20.psd", "15Х20.psd"):
            p = os.path.join(self.default_frame_dir, candidate)
            if os.path.exists(p):
                self.frame_1520_path = p
                break
        for candidate in ("20x30.psd", "20х30.psd", "20X30.psd", "20Х30.psd"):
            p = os.path.join(self.default_frame_dir, candidate)
            if os.path.exists(p):
                self.frame_2030_path = p
                break

    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def get_school_folders(self):
        if os.path.exists(self.default_schools_dir):
            try:
                folders = [
                    f for f in os.listdir(self.default_schools_dir)
                    if os.path.isdir(os.path.join(self.default_schools_dir, f))
                ]
                return sorted(folders)
            except Exception as e:
                print(f"Ошибка чтения директории 2026: {e}")
        return []

    def open_about_dialog(self):
        AboutDialog(self).exec()

    def open_path_settings(self):
        current = {
            "schools_dir": self.default_schools_dir,
            "frame_dir": self.default_frame_dir,
            "scripts_dir": self.scripts_dir,
            "ambient_video": self.ambient_video_path,
        }
        dlg = PathSettingsDialog(self, current_paths=current)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self.default_schools_dir = dlg.paths.get("schools_dir", "")
        self.default_frame_dir = dlg.paths.get("frame_dir", "")
        self.scripts_dir = dlg.paths.get("scripts_dir", "")
        self.ambient_video_path = dlg.paths.get("ambient_video", "")
        save_config({
            "schools_dir": self.default_schools_dir,
            "frame_dir": self.default_frame_dir,
            "scripts_dir": self.scripts_dir,
            "ambient_video": self.ambient_video_path,
        })
        self._refresh_frame_paths()
        self._load_ambient_video()
        self.refresh_school_list()

    def _load_ambient_video(self):
        if self.media_player_ambient is None:
            return
        try:
            self.media_player_ambient.stop()
        except Exception:
            pass
        ambient_path = None
        if self.ambient_video_path and os.path.exists(self.ambient_video_path):
            ambient_path = self.ambient_video_path
        if ambient_path is None:
            for c in (resource_path("deathnote_light1vid.mp4"),
                      os.path.join(os.getcwd(), "deathnote_light1vid.mp4"),
                      os.path.join(os.path.dirname(sys.executable), "deathnote_light1vid.mp4")):
                if c and os.path.exists(c):
                    ambient_path = c
                    break
        if ambient_path:
            try:
                self.media_player_ambient.setVideoOutput(self.video_widget_ambient)
                self.media_player_ambient.setSource(QUrl.fromLocalFile(ambient_path))
                self.media_player_ambient.setLoops(QMediaPlayer.Loops.Infinite)
                self.media_player_ambient.play()
            except Exception as e:
                print(f"Ошибка запуска видео: {e}")

    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        central_widget.setStyleSheet("background-color: #121212;")

        self.lbl_bg = QLabel(central_widget)
        self.lbl_bg.setGeometry(0, 0, 1480, 760)
        self.lbl_bg.setStyleSheet("background-color: #121212;")
        self.lbl_bg.lower()

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        header_frame = QFrame()
        header_frame.setStyleSheet(
            "QFrame { background-color: rgba(18, 18, 18, 210);"
            "border: 1px solid #2B2B2B; border-radius: 10px; }"
        )
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 10, 20, 10)
        header_layout.setSpacing(15)

        SIDE_WIDTH = 200
        left_side = QWidget()
        left_side.setFixedWidth(SIDE_WIDTH)
        header_layout.addWidget(left_side)

        welcome_container = QWidget()
        welcome_layout = QHBoxLayout(welcome_container)
        welcome_layout.setContentsMargins(0, 0, 0, 0)
        welcome_layout.setSpacing(0)
        welcome_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_welcome_part1 = QLabel("Добро пожаловать в ")
        lbl_welcome_part2 = QLabel("Photoshop WorkStation+")
        font_welcome = QFont("Segoe UI", 13, QFont.Weight.Bold)
        lbl_welcome_part1.setFont(font_welcome)
        lbl_welcome_part1.setStyleSheet(
            "color: #FFFFFF; border: none; background: transparent;"
        )
        lbl_welcome_part2.setFont(font_welcome)
        lbl_welcome_part2.setStyleSheet(
            "color: #00BFFF; border: none; background: transparent;"
        )
        welcome_layout.addWidget(lbl_welcome_part1)
        welcome_layout.addWidget(lbl_welcome_part2)
        header_layout.addWidget(welcome_container, 1)

        right_side = QWidget()
        right_side.setFixedWidth(SIDE_WIDTH + 100)
        right_side_layout = QHBoxLayout(right_side)
        right_side_layout.setContentsMargins(0, 0, 0, 0)
        right_side_layout.setSpacing(8)
        right_side_layout.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )

        btn_paths = QPushButton("Настройки")
        btn_paths.setStyleSheet(
            "QPushButton { background-color: #262626; color: #CCFF00;"
            "border: 1px solid #3A3A3A; border-radius: 13px;"
            "padding: 5px 12px; font-size: 11px; font-weight: bold; }"
            "QPushButton:hover { background-color: #383838;"
            "border-color: #CCFF00; }"
        )
        btn_paths.clicked.connect(self.open_path_settings)
        right_side_layout.addStretch()
        right_side_layout.addWidget(btn_paths)

        btn_about = QPushButton("О программе")
        btn_about.setStyleSheet(
            "QPushButton { background-color: #262626; color: #00BFFF;"
            "border: 1px solid #3A3A3A; border-radius: 13px;"
            "padding: 5px 12px; font-size: 11px; font-weight: bold; }"
            "QPushButton:hover { background-color: #383838;"
            "border-color: #00BFFF; }"
        )
        btn_about.clicked.connect(self.open_about_dialog)
        right_side_layout.addWidget(btn_about)
        header_layout.addWidget(right_side)
        main_layout.addWidget(header_frame)

        columns_layout = QHBoxLayout()
        columns_layout.setSpacing(15)

        card_style = (
            "QFrame { background-color: #121212; border: 1px solid #2B2B2B;"
            "border-radius: 12px; }"
        )
        btn_style_glass = (
            "QPushButton { background-color: rgba(255, 255, 255, 15);"
            "color: #FFFFFF; border: 1px solid rgba(255, 255, 255, 30);"
            "border-radius: 8px; padding: 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: rgba(255, 255, 255, 30);"
            "border: 1px solid rgba(255, 255, 255, 60); }"
            "QPushButton:disabled { background-color: rgba(255, 255, 255, 5);"
            "color: #AAAAAA; border: 1px solid rgba(255, 255, 255, 15); }"
        )
        text_color = "white"
        sub_bg = "#1A1A1A"
        sub_title_color = "#888888"
        chk_style = (
            "QCheckBox { color: white; font-weight: bold; border: none; }"
            "QCheckBox::indicator { border: 1px solid white; border-radius: 3px;"
            "background: #222222; width: 13px; height: 13px; }"
            "QCheckBox::indicator:checked { background-color: #CCFF00;"
            "border: 1px solid white; }"
        )
        log_bg = "#171717"
        log_fg = "#EEEEEE"

        # ==============================================================
        # КОЛОНКА 1 — РАБОЧАЯ ПАПКА (ШКОЛЫ)
        # ==============================================================
        card_school = QFrame()
        card_school.setStyleSheet(card_style)
        layout_school = QVBoxLayout(card_school)
        layout_school.setContentsMargins(15, 15, 15, 15)
        layout_school.setSpacing(10)

        title_school = QLabel("1. Рабочая папка (ШКОЛЫ)")
        title_school.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        title_school.setStyleSheet(f"color: {text_color}; border: none;")
        title_school.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_school.addWidget(title_school)

        video_container = QFrame()
        video_container.setStyleSheet(
            "QFrame { background-color: #000000; border: none; border-radius: 8px; }"
        )
        video_container.setFixedHeight(340)
        video_container.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        video_container_layout = QVBoxLayout(video_container)
        video_container_layout.setContentsMargins(0, 0, 0, 0)
        video_container_layout.setSpacing(0)

        self.video_widget_ambient = QVideoWidget()
        self.video_widget_ambient.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.video_widget_ambient.setStyleSheet(
            "QVideoWidget { background-color: #000000; border: none;"
            "border-radius: 8px; }"
        )
        self.video_widget_ambient.setAspectRatioMode(
            Qt.AspectRatioMode.KeepAspectRatioByExpanding
        )
        video_container_layout.addWidget(self.video_widget_ambient)
        layout_school.addWidget(video_container)
        layout_school.addSpacing(18)

        lbl_select = QLabel("Выберите школу:")
        lbl_select.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl_select.setStyleSheet(f"color: {text_color}; border: none;")
        layout_school.addWidget(lbl_select)

        self.combo_schools = PixelComboBox()
        schools = self.get_school_folders()
        self.combo_schools.addItem("Выберите школу из списка...")
        if schools:
            self.combo_schools.addItems(schools)
        else:
            self.combo_schools.addItem("Папки не найдены")
        self.combo_schools.currentTextChanged.connect(self.on_school_selected)
        self.combo_border = RGBBorderWidget(self.combo_schools, radius=18, speed=4)
        layout_school.addWidget(self.combo_border)

        self.frame_sub = QFrame()
        self.frame_sub.setStyleSheet(
            f"background-color: {sub_bg}; border-radius: 8px; border: none;"
        )
        layout_sub = QVBoxLayout(self.frame_sub)
        layout_sub.setContentsMargins(10, 8, 10, 8)

        lbl_sub = QLabel("Обрабатывать подпапки:")
        lbl_sub.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        lbl_sub.setStyleSheet(f"color: {sub_title_color}; border: none;")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_sub.addWidget(lbl_sub)

        layout_checks = QHBoxLayout()
        layout_checks.setSpacing(30)
        self.chk_1520 = QCheckBox("Папки 1520")
        self.chk_1520.setStyleSheet(chk_style)
        self.chk_1520.stateChanged.connect(self.update_file_list)
        self.chk_2030 = QCheckBox("Папки 2030")
        self.chk_2030.setStyleSheet(chk_style)
        self.chk_2030.stateChanged.connect(self.update_file_list)
        layout_checks.addStretch()
        layout_checks.addWidget(self.chk_1520)
        layout_checks.addWidget(self.chk_2030)
        layout_checks.addStretch()
        layout_sub.addLayout(layout_checks)
        layout_school.addWidget(self.frame_sub)

        self.qr_panel = QFrame()
        self.qr_panel.setStyleSheet(
            "QFrame { background-color: #1A0F2A; border: 1px solid #4A2A7A;"
            "border-radius: 8px; }"
        )
        layout_qr = QVBoxLayout(self.qr_panel)
        layout_qr.setContentsMargins(10, 10, 10, 10)
        layout_qr.setSpacing(6)

        lbl_qr_hint = QLabel("Режим QR-импорта активен")
        lbl_qr_hint.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        lbl_qr_hint.setStyleSheet("color: #CC88FF; border: none;")
        lbl_qr_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_qr.addWidget(lbl_qr_hint)

        self.lbl_qr_frame = QLabel("❌ Рамка: не выбрана")
        self.lbl_qr_frame.setWordWrap(True)
        self.lbl_qr_frame.setStyleSheet(
            "color: #FF5555; font-size: 10px; border: none;"
        )
        layout_qr.addWidget(self.lbl_qr_frame)

        self.lbl_qr_folder = QLabel("❌ Папка QR: не выбрана")
        self.lbl_qr_folder.setWordWrap(True)
        self.lbl_qr_folder.setStyleSheet(
            "color: #FF5555; font-size: 10px; border: none;"
        )
        layout_qr.addWidget(self.lbl_qr_folder)

        btn_qr_settings = QPushButton("Настроить QR-импорт")
        btn_qr_settings.setStyleSheet(
            "QPushButton { background-color: #7A4FCF; color: white;"
            "border: none; border-radius: 6px; padding: 6px;"
            "font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A3A9F; }"
        )
        btn_qr_settings.clicked.connect(self.open_qr_settings)
        layout_qr.addWidget(btn_qr_settings)
        self.qr_panel.setVisible(False)
        layout_school.addWidget(self.qr_panel)
        layout_school.addSpacing(15)

        self.lbl_path_info = QLabel("Школа не выбрана")
        self.lbl_path_info.setStyleSheet(
            "color: #FF4444; border: none; font-size: 12px; font-weight: bold;"
        )
        self.lbl_path_info.setWordWrap(True)
        self.lbl_path_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_school.addWidget(self.lbl_path_info)

        self.lbl_count_info = QLabel("")
        self.lbl_count_info.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.lbl_count_info.setStyleSheet("color: #CCFF00; border: none;")
        self.lbl_count_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_school.addWidget(self.lbl_count_info)
        layout_school.addStretch(1)

        self.media_player_ambient = QMediaPlayer(self)
        self.audio_output_ambient = QAudioOutput(self)
        self.media_player_ambient.setAudioOutput(self.audio_output_ambient)
        self.audio_output_ambient.setVolume(0)
        self._load_ambient_video()

        # ==============================================================
        # КОЛОНКА 2 — ПАНЕЛЬ СКРИПТОВ
        # ==============================================================
        card_scripts = QFrame()
        card_scripts.setStyleSheet(card_style)
        layout_scripts = QVBoxLayout(card_scripts)
        layout_scripts.setContentsMargins(15, 15, 15, 15)
        layout_scripts.setSpacing(6)

        title_scripts = QLabel("2. Панель скриптов")
        title_scripts.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        title_scripts.setStyleSheet(f"color: {text_color}; border: none;")
        title_scripts.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_scripts.addWidget(title_scripts)
        layout_scripts.addStretch(1)

        self.lbl_hint_scripts = QLabel(
            "Сначала выберите школу\nв первой колонке →"
        )
        self.lbl_hint_scripts.setWordWrap(True)
        self.lbl_hint_scripts.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_hint_scripts.setStyleSheet(
            "QLabel { color: #FF9800; background-color: #2A1A00;"
            "border: 1px solid #5A3A00; border-radius: 8px; padding: 6px;"
            "font-size: 11px; font-weight: bold; }"
        )
        layout_scripts.addWidget(self.lbl_hint_scripts)

        buttons_data = [
            ("1. СОЗДАТЬ ПАПКИ ШКОЛЫ",
             "create_folders", "", "1. Создать структуру папок", True),
            ("2. БЫСТРЫЙ ЭКСПОРТ JPG (2_FAST_Export jpg)",
            "folder_fast", "2_FAST_Export jpg.jsx", "2. ЭКСПОРТ JPG", True),
            ("3. ИМПОРТ QR-КОДОВ В РАМКУ 2030",
             "qr", "23_QR_add_FRAME2030.jsx", "3. Импорт QR-кодов", True),
            ("4. БЫСТРАЯ ПАКЕТНАЯ ОБРАБОТКА (25_batch_run_action)",
             "batch_actions", "25_batch_run_action.jsx", "4. BATCH ACTIONS", True),
            ("5. УМНАЯ ВСТАВКА РАМОК (ПРОСТО ДОБАВИТЬ РАМКИ) (24_AddFramesSmart)",
             "smart_frames", "24_AddFramesSmart.jsx", "5. Умная вставка рамок", True),
        ]

        script_palettes = [
            {"colors": [QColor(210, 110, 40), QColor(220, 80, 60), QColor(200, 60, 90), QColor(170, 60, 130), QColor(210, 110, 40)], "speed": 4},
            {"colors": [QColor(60, 170, 200), QColor(50, 140, 210), QColor(70, 110, 200), QColor(100, 90, 190), QColor(60, 170, 200)], "speed": 5},
            {"colors": [QColor(70, 190, 130), QColor(60, 200, 170), QColor(90, 180, 110), QColor(120, 190, 90), QColor(70, 190, 130)], "speed": 6},
            {"colors": [QColor(170, 80, 200), QColor(200, 80, 170), QColor(210, 90, 130), QColor(160, 70, 190), QColor(170, 80, 200)], "speed": 4},
            {"colors": [QColor(220, 190, 70), QColor(210, 160, 60), QColor(230, 200, 90), QColor(190, 150, 50), QColor(220, 190, 70)], "speed": 5},
        ]

        self.active_script_border = None
        self.script_buttons = []
        self.script_border_widgets.clear()

        for idx, (text, mode, jsx_file, title, visible) in enumerate(buttons_data):
            btn = QPushButton(text)
            btn.setStyleSheet(btn_style_glass)
            btn.setMinimumHeight(38)
            if mode == "create_folders" or mode == "qr":
                btn.setEnabled(True)
            else:
                btn.setEnabled(False)
            pal = script_palettes[idx % len(script_palettes)]
            border_wrap = RGBBorderWidget(btn, palette=pal["colors"],
                                          speed=pal["speed"], radius=8)
            self.script_border_widgets.append(border_wrap)
            self.script_buttons.append(btn)
            btn.setProperty("mode", mode)
            btn.setProperty("jsx_file", jsx_file)
            btn.setProperty("title", title)
            btn.setProperty("visible", visible)
            btn.setProperty("btn_index", idx)
            btn.clicked.connect(
                lambda checked=False, b=border_wrap, i=idx:
                self._on_script_clicked(b, i)
            )
            layout_scripts.addWidget(border_wrap)

        layout_scripts.addSpacing(20)
        btn_exit = QPushButton("Выход")
        btn_exit.setStyleSheet(
            "QPushButton { background-color: #eb2654; color: black;"
            "border: none; border-radius: 15px; font-weight: bold;"
            "font-size: 12px; padding: 8px; }"
            "QPushButton:hover { background-color: #474747; color: white; }"
        )
        btn_exit.setMinimumHeight(36)
        btn_exit.clicked.connect(self.close)
        layout_scripts.addWidget(btn_exit)
        layout_scripts.addSpacing(8)

        self.btn_cancel = QPushButton("Отменить выполнение / закрыть PS")
        self.btn_cancel.setStyleSheet(
            "QPushButton { background-color: #8f96ff; color: black;"
            "border: none; border-radius: 15px; font-weight: bold;"
            "font-size: 12px; padding: 8px; }"
            "QPushButton:hover { background-color: #376dad; color: white; }"
        )
        self.btn_cancel.setMinimumHeight(36)
        self.btn_cancel.clicked.connect(self.cancel_process)
        layout_scripts.addWidget(self.btn_cancel)
        layout_scripts.addStretch(1)

        frame_footer = QFrame()
        frame_footer.setStyleSheet("border: none; background: transparent;")
        layout_footer = QHBoxLayout(frame_footer)
        layout_footer.setContentsMargins(0, 5, 0, 0)
        layout_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_footer.setSpacing(0)
        lbl_dev = QLabel("made by ")
        lbl_dev.setStyleSheet(
            f"color: {text_color}; font-weight: bold; font-size: 12px;"
        )
        lbl_author = QLabel(
            '<a href="https://github.com/DirtSmoke44" '
            'style="color:#FF69B4; text-decoration:underline;">'
            '@DirtSmoke44</a>'
            '<span style="color:' + text_color + ';"> and </span>'
            '<a href="https://github.com/JohnnySuon" '
            'style="color:#FF8C00; text-decoration:underline;">'
            '@JohnnySuon</a>'
        )
        lbl_author.setOpenExternalLinks(False)
        lbl_author.setStyleSheet(
            "font-weight: bold; font-size: 12px; border: none;"
        )
        lbl_author.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        lbl_author.linkActivated.connect(self._open_github)
        layout_footer.addWidget(lbl_dev)
        layout_footer.addWidget(lbl_author)
        layout_scripts.addWidget(frame_footer)

        # ==============================================================
        # КОЛОНКА 3 — СТАТУС ВЫПОЛНЕНИЯ
        # ==============================================================
        card_status = QFrame()
        card_status.setStyleSheet(card_style)
        layout_status = QVBoxLayout(card_status)
        layout_status.setContentsMargins(15, 15, 15, 15)

        title_status = QLabel("3. Статус выполнения")
        title_status.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        title_status.setStyleSheet(f"color: {text_color}; border: none;")
        title_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_status.addWidget(title_status)

        self.txt_log = QTextEdit()
        self.txt_log.setFont(QFont("Consolas", 11))
        self.txt_log.setStyleSheet(
            f"QTextEdit {{ background-color: {log_bg}; color: {log_fg};"
            f"border: 1px solid #CCCCCC; border-radius: 8px; padding: 8px; }}"
        )
        self.txt_log.setReadOnly(True)
        layout_status.addWidget(self.txt_log)

        frame_status_anim = QFrame()
        frame_status_anim.setStyleSheet("border: none; background: transparent;")
        layout_status_anim = QHBoxLayout(frame_status_anim)
        layout_status_anim.setContentsMargins(0, 20, 0, 16)
        layout_status_anim.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_status_anim.setSpacing(8)
        self.pixel_icon = PixelIconWidget()
        layout_status_anim.addWidget(self.pixel_icon)
        self.lbl_anim_status = QLabel("В ожидании действий...")
        self.lbl_anim_status.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.lbl_anim_status.setStyleSheet("color: #CCFF00; border: none;")
        layout_status_anim.addWidget(self.lbl_anim_status)
        layout_status.addWidget(frame_status_anim)

        self.progress_bar = YellowGlowProgressBar()
        layout_status.addWidget(self.progress_bar)

        columns_layout.addWidget(card_school, 1)
        columns_layout.addWidget(card_scripts, 1)
        columns_layout.addWidget(card_status, 1)

        main_layout.addLayout(columns_layout)
        self.append_log("__INFO__Ожидание действий...")

    def _open_github(self, url):
        QDesktopServices.openUrl(QUrl(url))

    def create_school_folders(self, school_num):
        school_num = school_num.strip()
        if not school_num or len(school_num) < 3:
            return False, "Номер школы должен содержать минимум 3 символа."

        base_path = self.default_schools_dir
        school_dir = os.path.join(base_path, school_num)

        try:
            os.makedirs(base_path, exist_ok=True)
            os.makedirs(school_dir, exist_ok=True)
            for suffix in [" 1520", " 2030", " QR", " образцы"]:
                os.makedirs(os.path.join(school_dir, school_num + suffix),
                            exist_ok=True)
            os.makedirs(os.path.join(school_dir, "коды"), exist_ok=True)
            for sub in ("1520", "2030"):
                os.makedirs(os.path.join(school_dir, school_num + " образцы", sub),
                            exist_ok=True)
                os.makedirs(os.path.join(school_dir, school_num + " QR", sub),
                            exist_ok=True)
            return True, school_dir
        except Exception as e:
            return False, f"Ошибка создания папок:\n{e}"

    def open_create_folders_dialog(self):
        for wrap in self.script_border_widgets:
            wrap.stop_border()
        if self.script_border_widgets:
            self.script_border_widgets[0].start_border()
            self.active_script_border = self.script_border_widgets[0]

        school_num = tkinter_ask_school_number("Создание структуры папок")
        if school_num is None:
            return

        school_num = school_num.strip()
        if len(school_num) < 3:
            self.clear_log()
            self.append_log("__ERROR__Номер школы слишком короткий (минимум 3 символа).")
            return

        self.clear_log()
        self.append_log("__INFO__--- СОЗДАНИЕ СТРУКТУРЫ ПАПОК ---")
        self.append_log(f"__INFO__Номер школы: {school_num}")
        self.append_log(f"__INFO__Базовый путь: {self.default_schools_dir}")

        success, result = self.create_school_folders(school_num)
        if success:
            self.append_log("__OK__Структура папок успешно создана:")
            self.append_log(f"__OK__{result}")
            self.append_log("__INFO__Созданные разделы:")
            self.append_log(f"__INFO__• {school_num} 1520")
            self.append_log(f"__INFO__• {school_num} 2030")
            self.append_log(f"__INFO__• {school_num} QR (1520/2030)")
            self.append_log(f"__INFO__• {school_num} образцы (1520/2030)")
            self.append_log(f"__INFO__• коды")
            self.refresh_school_list()
        else:
            self.append_log(f"__ERROR__{result}")

    def refresh_school_list(self):
        current = self.combo_schools.currentText()
        self.combo_schools.blockSignals(True)
        self.combo_schools.clear()
        self.combo_schools.addItem("Выберите школу из списка...")
        schools = self.get_school_folders()
        if schools:
            self.combo_schools.addItems(schools)
        else:
            self.combo_schools.addItem("Папки не найдены")
        if current and current != "Выберите школу из списка...":
            idx = self.combo_schools.findText(current)
            if idx >= 0:
                self.combo_schools.setCurrentIndex(idx)
        self.combo_schools.blockSignals(False)

    def _on_script_clicked(self, border_wrap, btn_index):
        btn = self.script_buttons[btn_index]
        mode = btn.property("mode")

        if mode == "create_folders":
            self.open_create_folders_dialog()
            return

        if mode == "qr":
            for wrap in self.script_border_widgets:
                wrap.stop_border()
            border_wrap.start_border()
            self.active_script_border = border_wrap
            self.selected_script_mode = mode
            self.selected_script_file = btn.property("jsx_file")
            self.selected_script_title = btn.property("title")
            self.selected_script_visible = btn.property("visible")
            self.selected_script_index = btn_index
            self.frame_sub.setVisible(False)
            self.qr_panel.setVisible(True)
            auto_frame = self._find_default_frame()
            auto_qr = self._find_default_qr_folder() if self.selected_school_name else ""
            dlg = QRDialog(self, default_frame=auto_frame,
                           default_qr_folder=auto_qr,
                           fallback_qr_root=self.default_schools_dir)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                self._reset_script_selection()
                return
            self.qr_frame_path = dlg.frame_path
            self.qr_folder_path = dlg.qr_folder_path
            self._update_qr_labels()
            self.launch_worker()
            return

        if not self.selected_school_name:
            self.clear_log()
            self.append_log("__ERROR__Сначала выберите школу в первой колонке.")
            return

        for wrap in self.script_border_widgets:
            wrap.stop_border()
        border_wrap.start_border()
        self.active_script_border = border_wrap
        self.selected_script_mode = btn.property("mode")
        self.selected_script_file = btn.property("jsx_file")
        self.selected_script_title = btn.property("title")
        self.selected_script_visible = btn.property("visible")
        self.selected_script_index = btn_index

        if self.selected_script_mode == "batch_actions":
            self.frame_sub.setVisible(True)
            self.qr_panel.setVisible(False)
            dlg = BatchActionsDialog(
                self,
                default_set=self.batch_action_set,
                default_actions=self.batch_actions_text,
            )
            if dlg.exec() != QDialog.DialogCode.Accepted:
                self._reset_script_selection()
                return
            self.batch_action_set = dlg.action_set
            self.batch_actions_text = dlg.actions_text
            self.batch_fast_mode = dlg.fast_mode
            self.batch_auto_alert = dlg.auto_alert
            self.launch_worker()
        elif self.selected_script_mode == "smart_frames":
            self.frame_sub.setVisible(True)
            self.qr_panel.setVisible(False)
            self._refresh_frame_paths()
            self.launch_worker()
        else:
            self.frame_sub.setVisible(True)
            self.qr_panel.setVisible(False)
            self.launch_worker()

    def _reset_script_selection(self):
        for wrap in self.script_border_widgets:
            wrap.stop_border()
        self.active_script_border = None
        self.selected_script_mode = None
        self.qr_panel.setVisible(False)
        self.frame_sub.setVisible(True)

    def _find_default_frame(self):
        if not os.path.isdir(self.default_frame_dir):
            return ""
        for variant in ("20x30.psd", "20х30.psd", "20X30.psd", "20Х30.psd"):
            candidate = os.path.join(self.default_frame_dir, variant)
            if os.path.exists(candidate):
                return candidate
        return ""

    def _find_default_qr_folder(self):
        if not self.selected_folder:
            return ""
        codes_path = os.path.join(self.selected_folder, "коды")
        if os.path.isdir(codes_path):
            return codes_path
        return self.selected_folder

    def _update_qr_labels(self):
        if self.qr_frame_path and os.path.exists(self.qr_frame_path):
            self.lbl_qr_frame.setText(f"✅ Рамка: {os.path.basename(self.qr_frame_path)}")
            self.lbl_qr_frame.setStyleSheet("color: #66DD66; font-size: 10px; border: none;")
        else:
            self.lbl_qr_frame.setText("❌ Рамка: не выбрана")
            self.lbl_qr_frame.setStyleSheet("color: #FF5555; font-size: 10px; border: none;")
        if self.qr_folder_path and os.path.isdir(self.qr_folder_path):
            self.lbl_qr_folder.setText(f"✅ Папка QR: {os.path.basename(self.qr_folder_path)}")
            self.lbl_qr_folder.setStyleSheet("color: #66DD66; font-size: 10px; border: none;")
        else:
            self.lbl_qr_folder.setText("❌ Папка QR: не выбрана")
            self.lbl_qr_folder.setStyleSheet("color: #FF5555; font-size: 10px; border: none;")

    def open_qr_settings(self):
        dlg = QRDialog(
            self,
            default_frame=self._find_default_frame(),
            default_qr_folder=self._find_default_qr_folder() if self.selected_school_name else "",
            fallback_qr_root=self.default_schools_dir,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.qr_frame_path = dlg.frame_path
            self.qr_folder_path = dlg.qr_folder_path
            self._update_qr_labels()

    def append_log(self, text):
        safe_text = text
        if safe_text.startswith("__OK__"):
            inner = safe_text[len("__OK__"):]
            html_text = ("<span style='color:#7CFFCB; font-weight:bold;'>"
                         "✅ " + html.escape(inner) + "</span>")
        elif safe_text.startswith("__ERROR__"):
            inner = safe_text[len("__ERROR__"):]
            html_text = ("<span style='color:#FF8898; font-weight:bold;'>"
                         "❌ " + html.escape(inner) + "</span>")
        elif safe_text.startswith("__INFO__"):
            inner = safe_text[len("__INFO__"):]
            html_text = ("<span style='color:#7CC4FF;'>"
                         "ℹ️ " + html.escape(inner) + "</span>")
        elif safe_text.startswith("__WARN__"):
            inner = safe_text[len("__WARN__"):]
            html_text = ("<span style='color:#FFC074; font-weight:bold;'>"
                         "⚠️ " + html.escape(inner) + "</span>")
        else:
            html_text = ("<span style='color:#C8CEDA;'>"
                         "• " + html.escape(safe_text) + "</span>")
        self.txt_log.append(html_text)

    def clear_log(self):
        self.txt_log.clear()

    def on_school_selected(self, text):
        if (text and text != "Выберите школу из списка..."
                and text != "Папки не найдены"):
            self.selected_school_name = text
            self.selected_folder = os.path.join(self.default_schools_dir, text)
            self.lbl_path_info.setText(self.selected_folder)
            self.lbl_path_info.setStyleSheet(
                "color: #FFFFFF; border: none; font-size: 12px; font-weight: bold;"
            )
            self.update_file_list()
            self.combo_border.start_border()
            for btn in self.script_buttons:
                mode = btn.property("mode")
                if mode != "create_folders" and mode != "qr":
                    btn.setEnabled(True)
        else:
            if hasattr(self, "combo_border"):
                self.combo_border.stop_border()
            for btn in self.script_buttons:
                mode = btn.property("mode")
                if mode != "create_folders" and mode != "qr":
                    btn.setEnabled(False)
            self.selected_school_name = ""
            self.selected_folder = ""
            self.selected_files = []
            self.lbl_count_info.setText("")
            self.lbl_path_info.setText("Школа не выбрана")
            self.lbl_path_info.setStyleSheet(
                "color: #FF4444; border: none; font-size: 12px; font-weight: bold;"
            )

    def update_file_list(self):
        if not self.selected_folder or not os.path.exists(self.selected_folder):
            return
        use_1520 = self.chk_1520.isChecked()
        use_2030 = self.chk_2030.isChecked()
        matched_files = []
        if use_1520 or use_2030:
            for dp, dn, filenames in os.walk(self.selected_folder):
                folder_name = os.path.basename(dp).lower()
                is_1520 = "1520" in folder_name
                is_2030 = "2030" in folder_name
                if (use_1520 and is_1520) or (use_2030 and is_2030):
                    for f in filenames:
                        if f.lower().endswith(".psd"):
                            matched_files.append(os.path.join(dp, f))
        self.selected_files = matched_files
        self.lbl_count_info.setText(f"PSD файлов найдено: {len(self.selected_files)}")

    def check_ready_to_run(self):
        if self.selected_script_mode == "qr":
            if not self.qr_frame_path or not os.path.exists(self.qr_frame_path):
                self.clear_log()
                self.append_log("__ERROR__Не найдена рамка 20x30.psd.")
                return False
            if not self.qr_folder_path or not os.path.isdir(self.qr_folder_path):
                self.clear_log()
                self.append_log("__ERROR__Не найдена папка 'коды'.")
                return False
            return True
        if not self.selected_school_name:
            self.clear_log()
            self.append_log("__WARN__Сначала выберите школу.")
            return False
        if not self.selected_script_mode:
            self.clear_log()
            self.append_log("__WARN__Скрипт не выбран.")
            return False
        if self.selected_script_mode == "batch_actions":
            if not (self.chk_1520.isChecked() or self.chk_2030.isChecked()):
                self.clear_log()
                self.append_log(
                    "__ERROR__Для BATCH ACTIONS нужно отметить "
                    "галочку 'Папки 1520' или 'Папки 2030'!"
                )
                return False
            if not self.batch_action_set or not self.batch_actions_text:
                self.clear_log()
                self.append_log("__ERROR__Не заполнены Action Set или Actions.")
                return False
        if not (self.chk_1520.isChecked() or self.chk_2030.isChecked()):
            self.clear_log()
            self.append_log("__WARN__Не выбраны папки 1520 или 2030.")
            self.progress_bar.setValue(0)
            return False
        if not self.selected_files:
            self.clear_log()
            self.append_log("__WARN__Нет PSD для обработки.")
            self.progress_bar.setValue(0)
            return False
        if self.selected_script_mode == "smart_frames":
            if not self.frame_1520_path or not os.path.exists(self.frame_1520_path):
                self.clear_log()
                self.append_log("__ERROR__Не найден файл рамки 15x20.psd.")
                return False
            if not self.frame_2030_path or not os.path.exists(self.frame_2030_path):
                self.clear_log()
                self.append_log("__ERROR__Не найден файл рамки 20x30.psd.")
                return False
        return True

    def prepare_log_header(self, current_script_name):
        self.clear_log()
        if self.last_script_used and self.last_school_used:
            self.append_log("__INFO__----------------------------------------")
            self.append_log(f"__INFO__В прошлый раз: {self.last_script_used}")
            self.append_log(f"__INFO__для школы: {self.last_school_used}")
            self.append_log("__INFO__----------------------------------------")
        self.append_log(f"__INFO__Запущен скрипт: {current_script_name}")
        if self.selected_school_name:
            self.append_log(f"__INFO__Школа: {self.selected_school_name}")
        else:
            self.append_log("__INFO__Школа: не выбрана (QR-режим)")

    def launch_worker(self):
        if not self.check_ready_to_run():
            return
        self.prepare_log_header(self.selected_script_title)
        self.pixel_icon.start()
        self.progress_bar.start_anim()
        self.progress_bar.setValue(0)
        self.worker = WorkerThread(
            mode=self.selected_script_mode,
            files=self.selected_files,
            scripts_dir=self.scripts_dir,
            script_filename=self.selected_script_file,
            script_title=self.selected_script_title,
            school_name=self.selected_school_name,
            is_visible=False,
            qr_frame_path=self.qr_frame_path if self.selected_script_mode == "qr" else None,
            qr_folder_path=self.qr_folder_path if self.selected_script_mode == "qr" else None,
            action_set=self.batch_action_set if self.selected_script_mode == "batch_actions" else None,
            actions_list=self.batch_actions_text if self.selected_script_mode == "batch_actions" else None,
            fast_mode=self.batch_fast_mode,
            auto_alert=self.batch_auto_alert,
            frame_1520_path=self.frame_1520_path,
            frame_2030_path=self.frame_2030_path,
        )
        self.worker.progress_signal.connect(self.on_worker_progress)
        self.worker.file_status_signal.connect(self.on_file_status_update)
        self.worker.log_signal.connect(self.append_log)
        self.worker.finished_signal.connect(self.on_worker_finished)
        self.worker.start()

    def on_file_status_update(self, action_text, filename):
        if filename:
            self.lbl_anim_status.setText(f"{action_text}: {filename}")
        else:
            self.lbl_anim_status.setText(action_text)

    def on_worker_progress(self, current, total):
        val = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(val)

    def on_worker_finished(self, processed_count, script_title):
        self.pixel_icon.stop()
        self.progress_bar.stop_anim()
        self.lbl_anim_status.setText("Готово!")
        self.progress_bar.setValue(100)
        self.append_log("__INFO__========================================")
        self.append_log(f"__OK__Вы использовали скрипт: {script_title}")
        if self.selected_school_name:
            self.append_log(f"__OK__Выбранная школа: {self.selected_school_name}")
        self.append_log("__INFO__========================================")
        self.last_script_used = script_title
        self.last_school_used = self.selected_school_name
        QTimer.singleShot(3000, self._restart_icon_after_finish)

    def _restart_icon_after_finish(self):
        self.pixel_icon.start()
        if self.worker is None or not self.worker.isRunning():
            self.lbl_anim_status.setText("В ожидании действий...")

    def _kill_photoshop(self):
        try:
            subprocess.run(
                ["taskkill", "/f", "/im", "Photoshop.exe"],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass

    def cancel_process(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
        self._kill_photoshop()
        self.pixel_icon.stop()
        self.progress_bar.stop_anim()
        self.lbl_anim_status.setText("Процесс отменён / PS закрыт")
        self.append_log("__ERROR__ОПЕРАЦИЯ ОТМЕНЕНА / PHOTOSHOP ЗАКРЫТ!")
        self.progress_bar.setValue(0)
        QTimer.singleShot(3000, self._restart_icon_after_finish)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(3000)
        self._kill_photoshop()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    icon_path = resource_path("iconapp_Photoshop Workstation+.png")
    if not os.path.exists(icon_path):
        icon_path = resource_path("iconapp_Photoshop_Workstation+.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = ModernPhotoshopWorkstation()
    window.show()

    app.aboutToQuit.connect(window._kill_photoshop)
    sys.exit(app.exec())