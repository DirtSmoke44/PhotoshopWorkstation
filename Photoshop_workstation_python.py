import os
import sys
import subprocess
import tempfile
import math
import html
import json   
import win32com.client

from PyQt6.QtWidgets import QTextBrowser

from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QRectF, QUrl, QPointF, QSize,
    QPropertyAnimation, QEasingCurve,
)
from PyQt6.QtGui import (
    QFont, QPixmap, QImage, QCursor, QPainter, QColor,
    QLinearGradient, QPen, QConicalGradient, QDesktopServices,
    QRadialGradient, QIcon, QBrush, QPainterPath,
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QComboBox, QCheckBox, QPushButton,
    QTextEdit, QProgressBar, QFrame, QDialog, QTextBrowser,
    QSizePolicy, QFileDialog, QListView, QLayout,
    QListWidget, QListWidgetItem, QAbstractItemView, QScrollArea,
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
# ★★★ КНОПКА УПРАВЛЕНИЯ ОКНОМ ★★★
# ----------------------------------------------------------------------
class CircleWindowButton(QPushButton):
    def __init__(self, icon_kind="min", style_key="min", parent=None):
        super().__init__(parent)
        self.icon_kind = icon_kind
        self.style_key = style_key

        palettes = {
            "min": {
                "bg": "#2A103A", "fg": "#CC88FF",
                "border": "#4A2A7A",
                "hover_bg": "#3A185A", "hover_border": "#CC88FF",
                "hover_fg": "#E0B7FF",
                "pressed_bg": "#A855F7", "pressed_fg": "#1A0A2A",
            },
            "max": {
                "bg": "#0E3A18", "fg": "#5CFF8A",
                "border": "#1A5A28",
                "hover_bg": "#1A5A28", "hover_border": "#5CFF8A",
                "hover_fg": "#9CFFB8",
                "pressed_bg": "#2ECC71", "pressed_fg": "#0A1A0E",
            },
            "close": {
                "bg": "#3A0A1A", "fg": "#FF4081",
                "border": "#5A1028",
                "hover_bg": "#5A1028", "hover_border": "#FF4081",
                "hover_fg": "#FF80AB",
                "pressed_bg": "#E91E63", "pressed_fg": "#FFFFFF",
            },
        }
        self.pal = palettes.get(style_key, palettes["min"])

        self._icon_color = QColor(self.pal["fg"])
        self.setFixedSize(30, 30)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setStyleSheet("QPushButton { background: transparent; border: none; }")
        self.setFlat(True)

    def set_icon_kind(self, kind):
        self.icon_kind = kind
        self.update()

    def enterEvent(self, event):
        self._icon_color = QColor(self.pal["hover_fg"])
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._icon_color = QColor(self.pal["fg"])
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)

        if self.isDown():
            bg = QColor(self.pal["pressed_bg"])
            border = QColor(self.pal["hover_border"])
            icon_color = QColor(self.pal["pressed_fg"])
        elif self.underMouse():
            bg = QColor(self.pal["hover_bg"])
            border = QColor(self.pal["hover_border"])
            icon_color = QColor(self.pal["hover_fg"])
        else:
            bg = QColor(self.pal["bg"])
            border = QColor(self.pal["border"])
            icon_color = QColor(self.pal["fg"])

        painter.setBrush(bg)
        painter.setPen(QPen(border, 1.5))
        painter.drawEllipse(rect)

        pen = QPen(icon_color, 1.8)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)

        cx = rect.center().x()
        cy = rect.center().y()
        s = 4.0

        if self.icon_kind == "min":
            painter.drawLine(QPointF(cx - s, cy + 1),
                             QPointF(cx + s, cy + 1))
        elif self.icon_kind == "max":
            painter.drawRect(QRectF(cx - s, cy - s, s * 2, s * 2))
        elif self.icon_kind == "restore":
            painter.drawRect(QRectF(cx - s + 1, cy - s + 2,
                                    s * 2 - 2, s * 2 - 2))
            painter.drawLine(QPointF(cx - s + 3, cy - s),
                             QPointF(cx + s, cy - s))
            painter.drawLine(QPointF(cx + s, cy - s),
                             QPointF(cx + s, cy + s - 3))
        elif self.icon_kind == "close":
            painter.drawLine(QPointF(cx - s, cy - s),
                             QPointF(cx + s, cy + s))
            painter.drawLine(QPointF(cx + s, cy - s),
                             QPointF(cx - s, cy + s))

        painter.end()


# ----------------------------------------------------------------------
# ★★★ КРУГЛАЯ ИКОНКА-КНОПКА (НАСТРОЙКИ / О ПРОГРАММЕ / ПРОФИЛЬ) ★★★
# ----------------------------------------------------------------------
class CircleIconButton(QPushButton):
    """Круглая кнопка с нарисованной иконкой: settings, about, user."""
    def __init__(self, icon_kind="settings", style_key="settings",
                 tooltip="", parent=None):
        super().__init__(parent)
        self.icon_kind = icon_kind
        self.style_key = style_key

        palettes = {
            "settings": {
                "bg": "#3A2010", "fg": "#FFB74D",
                "border": "#5A3418",
                "hover_bg": "#553018", "hover_border": "#FFB74D",
                "hover_fg": "#FFCC80",
                "pressed_bg": "#FF9800", "pressed_fg": "#1A0A02",
            },
            "about": {
                "bg": "#10243A", "fg": "#7CC4FF",
                "border": "#183A5A",
                "hover_bg": "#183A5A", "hover_border": "#7CC4FF",
                "hover_fg": "#A8D8FF",
                "pressed_bg": "#58A8E8", "pressed_fg": "#10243A",
            },
            "user": {
                "bg": "#0E3A3A", "fg": "#5CE8E0",
                "border": "#185A5A",
                "hover_bg": "#185A5A", "hover_border": "#5CE8E0",
                "hover_fg": "#88F0E8",
                "pressed_bg": "#3DD0C8", "pressed_fg": "#0A1A1A",
            },
        }
        self.pal = palettes.get(style_key, palettes["settings"])

        self.setFixedSize(34, 34)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setStyleSheet("QPushButton { background: transparent; border: none; }")
        self.setFlat(True)
        if tooltip:
            self.setToolTip(tooltip)

    def enterEvent(self, event):
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(3.0, 3.0, -3.0, -3.0)

        if self.isDown():
            bg = QColor(self.pal["pressed_bg"])
            border = QColor(self.pal["hover_border"])
            icon_color = QColor(self.pal["pressed_fg"])
        elif self.underMouse():
            bg = QColor(self.pal["hover_bg"])
            border = QColor(self.pal["hover_border"])
            icon_color = QColor(self.pal["hover_fg"])
        else:
            bg = QColor(self.pal["bg"])
            border = QColor(self.pal["border"])
            icon_color = QColor(self.pal["fg"])

        painter.setBrush(bg)
        painter.setPen(QPen(border, 1.5))
        painter.drawEllipse(rect)

        pen = QPen(icon_color, 1.4)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)

        cx = rect.center().x()
        cy = rect.center().y()

        if self.icon_kind == "settings":
            r_inner = 3.8
            r_outer = 5.4
            painter.drawEllipse(QPointF(cx, cy), r_inner, r_inner)
            for i in range(8):
                angle = math.radians(i * 45)
                x1 = cx + r_inner * math.cos(angle)
                y1 = cy + r_inner * math.sin(angle)
                x2 = cx + r_outer * math.cos(angle)
                y2 = cy + r_outer * math.sin(angle)
                painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))
        elif self.icon_kind == "about":
            painter.drawEllipse(QPointF(cx, cy - 1.5), 3.8, 3.8)
            painter.drawLine(QPointF(cx - 2.4, cy + 2.5),
                             QPointF(cx + 2.4, cy + 2.5))
            painter.drawLine(QPointF(cx - 2.0, cy + 4.2),
                             QPointF(cx + 2.0, cy + 4.2))
        elif self.icon_kind == "user":
            painter.drawEllipse(QPointF(cx, cy - 2.5), 3.0, 3.0)
            path = QPainterPath()
            path.moveTo(cx - 5.2, cy + 5.2)
            path.cubicTo(cx - 5.2, cy + 1.6,
                         cx + 5.2, cy + 1.6,
                         cx + 5.2, cy + 5.2)
            painter.drawPath(path)

        painter.end()


# ----------------------------------------------------------------------
# ★★★ АНИМИРОВАННЫЙ ЧЕКБОКС ★★★
# ----------------------------------------------------------------------
class PulseCheckBox(QCheckBox):
    def __init__(self, text, base_color, hover_color, dim_bg, hover_bg,
                 flash_color, parent=None):
        super().__init__(text, parent)
        self.base_color = QColor(base_color)
        self.hover_color = QColor(hover_color)
        self.dim_bg = QColor(dim_bg)
        self.hover_bg = QColor(hover_bg)
        self.flash_color = QColor(flash_color)

        self._pulse_phase = 0.0
        self._pulse_active = False
        self._flash_phase = 0.0
        self._flash_active = False
        self._hover_progress = 0.0

        self._pulse_timer = QTimer(self)
        self._pulse_timer.setInterval(16)
        self._pulse_timer.timeout.connect(self._tick)

        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setStyleSheet(
            "QCheckBox { color: #FFFFFF; font-weight: bold; border: none;"
            "padding-left: 4px; }"
        )
        f = self.font()
        f.setPointSize(11)
        f.setBold(True)
        self.setFont(f)
        self.setMinimumHeight(28)

        self.stateChanged.connect(self._on_state_changed)

    def _on_state_changed(self, _state):
        self._pulse_phase = 0.0
        self._pulse_active = True
        self._flash_phase = 0.0
        self._flash_active = True
        if not self._pulse_timer.isActive():
            self._pulse_timer.start()
        self.update()

    def _tick(self):
        if self._pulse_active:
            self._pulse_phase += 0.05
            if self._pulse_phase >= 1.0:
                self._pulse_phase = 0.0
                self._pulse_active = False

        if self._flash_active:
            self._flash_phase += 0.08
            if self._flash_phase >= 1.0:
                self._flash_phase = 0.0
                self._flash_active = False

        target = 1.0 if self.underMouse() else 0.0
        self._hover_progress += (target - self._hover_progress) * 0.25

        if (not self._pulse_active and not self._flash_active
                and abs(target - self._hover_progress) < 0.01):
            self._hover_progress = target
            self._pulse_timer.stop()

        self.update()

    def enterEvent(self, event):
        if not self._pulse_timer.isActive():
            self._pulse_timer.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if not self._pulse_timer.isActive():
            self._pulse_timer.start()
        super().leaveEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

        rect = self.rect()
        box_size = 12
        margin = 6
        box_rect = QRectF(
            rect.left() + margin,
            rect.center().y() - box_size / 2.5,
            box_size, box_size,
        )
        radius = 5.0
        checked = self.isChecked()

        if checked:
            base = self.base_color
            border = self.hover_color
        else:
            base = self.base_color
            border = self.base_color

        if self._pulse_active:
            wave = 1.0 - self._pulse_phase
            expand = 6 * self._pulse_phase + 2
            wave_alpha = int(180 * wave)
            wave_color = QColor(self.flash_color)
            wave_color.setAlpha(wave_alpha)
            wave_rect = box_rect.adjusted(-expand, -expand, expand, expand)
            pen = QPen(wave_color, 2.0)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(wave_rect, radius + expand * 0.5,
                                    radius + expand * 0.5)

        if self._hover_progress > 0.01:
            glow = QColor(self.hover_color)
            glow.setAlpha(int(90 * self._hover_progress))
            painter.setBrush(glow)
            painter.setPen(Qt.PenStyle.NoPen)
            expand = 2
            painter.drawRoundedRect(
                box_rect.adjusted(-expand, -expand, expand, expand),
                radius + 1, radius + 1,
            )

        if checked:
            fill_col = self.base_color
        else:
            h = self._hover_progress
            dim = self.dim_bg
            hov = self.hover_bg
            fill_col = QColor(
                int(dim.red() * (1 - h) + hov.red() * h),
                int(dim.green() * (1 - h) + hov.green() * h),
                int(dim.blue() * (1 - h) + hov.blue() * h),
            )

        painter.setBrush(fill_col)
        pen = QPen(border, 2.0)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawRoundedRect(box_rect, radius, radius)

        if checked:
            check_pen = QPen(QColor("#0A1A0E"), 2.4)
            check_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            check_pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
            painter.setPen(check_pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)

            p1 = QPointF(box_rect.left() + box_size * 0.25,
                         box_rect.top() + box_size * 0.55)
            p2 = QPointF(box_rect.left() + box_size * 0.45,
                         box_rect.top() + box_size * 0.72)
            p3 = QPointF(box_rect.left() + box_size * 0.78,
                         box_rect.top() + box_size * 0.30)

            if self._flash_active:
                flashes = 3
                f = self._flash_phase
                pulse = abs(math.sin(f * math.pi * flashes))
                alpha = int(120 + 135 * pulse)
                check_pen.setColor(QColor(10, 26, 14, alpha))
                painter.setPen(check_pen)

            painter.drawLine(p1, p2)
            painter.drawLine(p2, p3)

        text_rect = QRectF(
            box_rect.right() + 10, rect.top(),
            rect.width() - box_rect.right() - 10, rect.height(),
        )
        painter.setPen(QColor("#FFFFFF"))
        painter.setFont(self.font())
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            self.text(),
        )
        
        painter.end()


# ----------------------------------------------------------------------
# ★★★ TKINTER-ДИАЛОГ ВВОДА НОМЕРА ШКОЛЫ ★★★
# ----------------------------------------------------------------------
def tkinter_ask_school_number(parent_title="Создание структуры папок"):
    import tkinter as tk

    result = {"value": None}

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

    root.geometry(f"{win_w}x{win_h}+0+0")
    root.update_idletasks()
    root.update()

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

    btn_row = tk.Frame(container, bg=BG)
    btn_row.pack(fill=tk.X)

    def make_rounded_button(parent, text, fill_color, text_color,
                            hover_color, command,
                            width=210, height=44, radius=16):
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
        fill_color="#10243A",
        text_color="#7CC4FF",
        hover_color="#183A5A",
        command=on_ok,
    )

    make_rounded_button(
        btn_row,
        text="Отмена",
        fill_color="#3A0A1A",
        text_color="#FF4081",
        hover_color="#5A1028",
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
    ("RAW + LEVELS + COLOR BALANCE", "CAMERA RAW", "raw_levels_colorbalance", "raw_levels_color_balance"),
    ("BRIGHT +0.15", "BRIGHT ", "+0.15 bright", "bright_plus"),
    ("BRIGHT -0.15", "BRIGHT ", "-0.15 bright", "bright_minus"),
    ("BLUR фон", "BACKGROUND_BLUR", "select_inverse_blursurface", "blur_bg"),
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
        "raw_levels_color_balance": ("#2A2A10", "#FFE082", "#3A3A18", "#FFF0B0", "#E8C858"),
        "raw_export": ("#10243A", "#7CC4FF", "#183A5A", "#A8D8FF", "#58A8E8"),
        "blur_bg": ("#2A103A", "#D89BFF", "#3A185A", "#E8B7FF", "#B87BE8"),
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
            border-radius: 13px;  /* Скругленные углы */
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
        QPushButton:disabled {{
            background-color: #1A1A1A;
            color: #555555;
            border: 2px solid #2A2A2A;
        }}
    """

# ----------------------------------------------------------------------
# RGB ANIMATED BORDER
# ----------------------------------------------------------------------
class RGBBorderWidget(QWidget):
    def __init__(self, child_widget, palette=None, speed=4, radius=0, parent=None):
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
        try:
            screen = self.screen() or QApplication.primaryScreen()
            if screen is not None:
                avail = screen.availableGeometry()
                combo_bottom_y = self.mapToGlobal(QPoint(0, self.height())).y()
                space_below = avail.bottom() - combo_bottom_y
                item_height = 40
                max_items_below = max(3, (space_below - 20) // item_height)
                self.setMaxVisibleItems(int(max_items_below))
        except Exception:
            pass

        super().showPopup()

        try:
            popup = self.view().window()
            if popup is not None:
                popup.setStyleSheet(
                    "QFrame { background-color: #1E2229; border: 1px solid #3A4150;"
                    "border-radius: 18px; }"
                )
        except Exception:
            pass


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


class FileCheckItemWidget(QWidget):
    """Кастомная строка списка: квадратик с галочкой + название файла."""
    def __init__(self, text, checked=False, parent=None):
        super().__init__(parent)
        self._checked = checked
        self._text = text
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def is_checked(self):
        return self._checked

    def set_checked(self, state):
        self._checked = state
        self.update()

    def toggle(self):
        self._checked = not self._checked
        self.update()
        return self._checked

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

        rect = self.rect()
        if self._checked:
            p.setBrush(QColor("#1E3A5F"))
        else:
            p.setBrush(Qt.BrushStyle.NoBrush)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(rect.adjusted(0, 0, -1, -1), 6, 6)

        box_size = 16
        box_x = 10
        box_y = (rect.height() - box_size) // 2
        box_rect = QRectF(box_x, box_y, box_size, box_size)

        if self._checked:
            p.setBrush(QColor("#00BFFF"))
            p.setPen(QPen(QColor("#66D9FF"), 1.2))
        else:
            p.setBrush(QColor("#0E0E0E"))
            p.setPen(QPen(QColor("#555555"), 1.5))
        p.drawRoundedRect(box_rect, 3, 3)

        if self._checked:
            pen = QPen(QColor("#0A1A2A"), 2.0)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)
            p1 = QPointF(box_rect.left() + box_size * 0.22,
                         box_rect.top() + box_size * 0.52)
            p2 = QPointF(box_rect.left() + box_size * 0.42,
                         box_rect.top() + box_size * 0.72)
            p3 = QPointF(box_rect.left() + box_size * 0.78,
                         box_rect.top() + box_size * 0.28)
            p.drawLine(p1, p2)
            p.drawLine(p2, p3)

        p.setPen(QColor("#FFFFFF"))
        f = self.font()
        f.setFamily("Consolas")
        f.setPointSize(11)
        p.setFont(f)
        text_rect = QRectF(box_rect.right() + 10, 0,
                           rect.width() - box_rect.right() - 15, rect.height())
        p.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            self._text,
        )
        p.end()


class PickPSDFilesForInsertQRDialog(QDialog):
    """Выбор PSD-файлов для пофайловой вставки QR-кодов (кнопка 2)."""
    def __init__(self, parent=None, school_folder="", preselected=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор PSD-файлов для вставки QR-кодов")
        self.setFixedSize(720, 560)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.school_folder = school_folder
        self.selected_files = list(preselected or [])
        self.file_rows = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(10)

        title = QLabel("Выберите PSD-файлы для вставки QR-кодов во все классы 2030")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setStyleSheet("color: #FF70D8; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setWordWrap(True)
        layout.addWidget(title)

        hint = QLabel(
            "Клик по строке — поставить/снять галочку.\n"
            "«Добавить» — вставить QR только в выбранные файлы."
        )
        hint.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(
            """
            QListWidget {
                background-color: #1A1A1A; color: #FFFFFF;
                border: 1px solid #333333; border-radius: 8px;
                padding: 4px;
                outline: none;
            }
            QListWidget::item {
                border-radius: 6px;
                margin: 1px 0px;
            }
            QListWidget::item:selected {
                background: transparent;
            }
            """
        )
        self.list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        layout.addWidget(self.list_widget, 1)
        self._populate_files()

        self.lbl_count = QLabel("Выбрано: 0")
        self.lbl_count.setStyleSheet(
            "color: #FF70D8; font-weight: bold; font-size: 11px; border: none;"
        )
        self.lbl_count.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_count)
        self._update_count()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_all = QPushButton("Выбрать все")
        btn_all.setStyleSheet(
            "QPushButton { background-color: #103A18; color: #88FFA8;"
            "border: 2px solid #185A28; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #185A28;"
            "border: 1px solid #88FFA8; color: #B0FFC8; }"
            "QPushButton:pressed { background-color: #58E878;"
            "border: 1px solid #88FFA8; color: #0A1A0E; }"
        )
        btn_all.setMinimumHeight(38)
        btn_all.clicked.connect(self._on_select_all)
        btn_row.addWidget(btn_all, 1)

        btn_add = QPushButton("Добавить")
        btn_add.setStyleSheet(
            "QPushButton { background-color: #3A3410; color: #FFE55C;"
            "border: 2px solid #5A5018; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A5018;"
            "border: 1px solid #FFE55C; color: #FFF088; }"
            "QPushButton:pressed { background-color: #FFD600;"
            "border: 1px solid #FFE55C; color: #1A1A00; }"
        )
        btn_add.setMinimumHeight(38)
        btn_add.clicked.connect(self._on_add)
        btn_row.addWidget(btn_add, 1)

        btn_back = QPushButton("Назад")
        btn_back.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF4081; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_back.setMinimumHeight(38)
        btn_back.clicked.connect(self.reject)
        btn_row.addWidget(btn_back, 1)
        layout.addLayout(btn_row)

    def _populate_files(self):
        if not self.school_folder or not os.path.isdir(self.school_folder):
            return
        for dp, dn, filenames in os.walk(self.school_folder):
            folder_name = os.path.basename(dp).lower()
            if "2030" not in folder_name:
                continue
            for f in filenames:
                if f.lower().endswith(".psd"):
                    full = os.path.join(dp, f)
                    rel = os.path.relpath(full, self.school_folder)
                    is_checked = full in self.selected_files
                    item = QListWidgetItem()
                    item.setData(Qt.ItemDataRole.UserRole, full)
                    item.setSizeHint(QSize(0, 34))
                    widget = FileCheckItemWidget(rel, checked=is_checked)
                    self.list_widget.addItem(item)
                    self.list_widget.setItemWidget(item, widget)
                    widget.mousePressEvent = (
                        lambda ev, w=widget, it=item:
                        self._on_row_clicked(w, it, ev)
                    )
                    self.file_rows.append((item, widget))

    def _on_row_clicked(self, widget, item, event):
        if event.button() == Qt.MouseButton.LeftButton:
            widget.toggle()
            self._update_count()

    def _on_select_all(self):
        all_checked = all(w.is_checked() for _, w in self.file_rows) if self.file_rows else False
        new_state = not all_checked
        for _, widget in self.file_rows:
            widget.set_checked(new_state)
        self._update_count()

    def _update_count(self):
        cnt = sum(1 for _, w in self.file_rows if w.is_checked())
        self.lbl_count.setText(f"Выбрано: {cnt}")
        self.lbl_count.setStyleSheet(
            "color: #FF70D8; font-weight: bold; font-size: 11px; border: none;"
        )

    def _on_add(self):
        checked_files = [
            item.data(Qt.ItemDataRole.UserRole)
            for item, widget in self.file_rows
            if widget.is_checked()
        ]
        if not checked_files:
            self.lbl_count.setText("❌ Выберите хотя бы один PSD файл!")
            self.lbl_count.setStyleSheet(
                "color: #FF5555; font-weight: bold; font-size: 11px; border: none;"
            )
            return
        self.selected_files = checked_files
        self.accept()


# ----------------------------------------------------------------------
# ДИАЛОГ BATCH ACTIONS (пункт 4)
# ----------------------------------------------------------------------
class BatchActionsDialog(QDialog):
    def __init__(self, parent=None, default_set="", default_actions=""):
        super().__init__(parent)
        self.setWindowTitle("БЫСТРАЯ ПАКЕТНАЯ ОБРАБОТКА")
        self.setFixedSize(720, 760)
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
        title.setStyleSheet("color: #c5fa75; border: none;")
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

        self.presets_grid = QGridLayout()
        self.presets_grid.setSpacing(8)
        self.preset_buttons = []

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
            self.presets_grid.addWidget(p_btn, row, col)
            self.preset_buttons.append((p_btn, color_key))

        layout.addLayout(self.presets_grid)
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
            "QPushButton { background-color: #0E3A18; color: #5CFF8A;"
            "border: 2px solid #1A5A28; border-radius: 0px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #1A5A28;"
            "border: 1px solid #5CFF8A; color: #9CFFB8; }"
            "QPushButton:pressed { background-color: #2ECC71;"
            "border: 1px solid #5CFF8A; color: #0A1A0E; }"
            "QPushButton:disabled { background-color: #1A1A1A;"
            "color: #555555; border: 2px solid #2A2A2A; }"
        )
        self.btn_start.clicked.connect(self._on_start)
        btn_row.addWidget(self.btn_start, 1)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 0px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF4081; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel, 1)
        layout.addLayout(btn_row)

        self.input_set.textChanged.connect(self._validate)
        self.input_actions.textChanged.connect(self._validate)

        self._update_buttons_state()
        self._validate()

    def _get_frame_flags(self):
        if self.main_window is None:
            return False, False, False
        try:
            h1520 = self.main_window.chk_1520.isChecked()
            h2030 = self.main_window.chk_2030.isChecked()
            return h1520, h2030, (h1520 or h2030)
        except Exception:
            return False, False, False

    def _update_buttons_state(self):
        has_1520, has_2030, has_any = self._get_frame_flags()

        for btn, color_key in self.preset_buttons:
            if color_key in ("replace_2030",):
                btn.setEnabled(has_2030)
            elif color_key in ("replace_1520",):
                btn.setEnabled(has_1520)
            else:
                btn.setEnabled(has_any)

    def _apply_preset(self, preset_set, preset_actions, color_key):
        has_1520, has_2030, has_any = self._get_frame_flags()

        allowed = False
        if color_key == "replace_2030":
            allowed = has_2030
        elif color_key == "replace_1520":
            allowed = has_1520
        else:
            allowed = has_any

        if not allowed:
            try:
                if self.main_window is not None:
                    if color_key == "replace_2030":
                        msg = "Пресет 'Заменить рамку 2030' доступен только при выборе папки 2030!"
                    elif color_key == "replace_1520":
                        msg = "Пресет 'Заменить рамку 1520' доступен только при выборе папки 1520!"
                    else:
                        msg = "Пресет не применён: не выбраны папки 1520 или 2030! Поставьте галочку под школой."
                    self.main_window.append_log(f"__ERROR__{msg}")
            except Exception:
                pass
            self.lbl_status.setText("❌ Выберите подходящие папки в главном окне!")
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
                "color: #66DD66; border: none; font-size: 14px; font-weight: bold;"
            )
        else:
            self.lbl_status.setText("❌ Заполните Action Set и Actions")
            self.lbl_status.setStyleSheet(
                "color: #FF5555; border: none; font-size: 14px; font-weight: bold;"
            )
        self.btn_start.setEnabled(set_ok and act_ok)

    def _on_start(self):
        self.action_set = self.input_set.toPlainText().strip()
        self.actions_text = self.input_actions.toPlainText().strip()
        self.fast_mode = self.chk_fast.isChecked()
        self.auto_alert = self.chk_alert.isChecked()
        self.accept()


# ----------------------------------------------------------------------
# ★★★ ПРЫГАЮЩИЙ ШАРИК (для ожидания) ★★★
# ----------------------------------------------------------------------
class BouncingBallWidget(QWidget):
    """Прыгающий шарик с тенью, как в CSS ldr-bouncing-ball."""
    def __init__(self, color="#3b82f6", size=24, parent=None):
        super().__init__(parent)
        self._color = QColor(color)
        self._size = size
        self.setFixedSize(size + 4, int(size * 1.6))
        self._phase = 0.0
        self._active = False
        self._timer = QTimer(self)
        self._timer.setInterval(30)
        self._timer.timeout.connect(self._tick)

    def set_active(self, active: bool):
        self._active = active
        if active:
            if not self._timer.isActive():
                self._timer.start()
        else:
            self._timer.stop()
            self._phase = 0.0
        self.update()

    def _tick(self):
        self._phase = (self._phase + 0.028) % 1.0
        self.update()

    def paintEvent(self, event):
        if not self._active:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        t = self._phase * 2.0
        if t > 1.0:
            t = 2.0 - t
        t_eased = t * t

        s = self._size
        drop_dist = s * 0.45

        ball_y = t_eased * drop_dist
        scale_y = 1.0 + 0.08 * t_eased
        scale_x = 1.0 - 0.08 * t_eased

        cx = self.width() / 2.0
        ball_cy = s / 2.0 + ball_y

        shadow_w = s * (0.5 + 0.5 * t_eased)
        shadow_h = s / 6.0
        shadow_rect = QRectF(
            cx - shadow_w / 2.0,
            self.height() - shadow_h - 1,
            shadow_w, shadow_h
        )
        shadow_color = QColor(self._color)
        shadow_color.setAlpha(int(60 + 100 * t_eased))
        p.setBrush(shadow_color)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(shadow_rect)

        p.save()
        p.translate(cx, ball_cy)
        p.scale(scale_x, scale_y)
        p.setBrush(self._color)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QPointF(0, 0), s / 2.0, s / 2.0)
        p.restore()
        p.end()


# ----------------------------------------------------------------------
# ★★★ GENERATING BADGE ★★★
# ----------------------------------------------------------------------
class GeneratingBadgeWidget(QWidget):
    """Анимированный бейдж 'Generating' (уменьшенная скорость, центрирован)."""
    def __init__(self, text="Generating", color="#3b82f6", parent=None):
        super().__init__(parent)
        self._text = text
        self._color = QColor(color)
        self._angle = 0.0
        self._twinkle_phase = 0.0
        self._active = False

        f = QFont("Segoe UI", 10, QFont.Weight.Bold)
        self.setFont(f)
        self.setMinimumHeight(38)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self._timer = QTimer(self)
        self._timer.setInterval(40)
        self._timer.timeout.connect(self._tick)

    def set_text(self, text):
        self._text = text
        self.update()

    def set_active(self, active: bool):
        self._active = active
        if active:
            if not self._timer.isActive():
                self._timer.start()
        else:
            self._timer.stop()
            self._angle = 0.0
            self._twinkle_phase = 0.0
        self.update()

    def _tick(self):
        self._angle = (self._angle + 2.5) % 360.0
        self._twinkle_phase = (self._twinkle_phase + 0.03) % 1.0
        self.update()

    def paintEvent(self, event):
        if not self._active:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

        fm = painter.fontMetrics()
        text = self._text or "Generating"
        text_w = fm.horizontalAdvance(text)
        icon_size = fm.height() + 2
        pad_h = 14
        pad_v = 6
        gap = 8
        pill_w = icon_size + gap + text_w + pad_h * 2
        pill_h = max(icon_size, fm.height()) + pad_v * 2

        rect = QRectF((self.width() - pill_w) / 2.0,
                      (self.height() - pill_h) / 2.0, pill_w, pill_h)
        radius = pill_h / 2.0

        bg = QColor(self._color)
        bg.setAlpha(25)
        painter.setBrush(bg)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(rect, radius, radius)

        border = QColor(self._color)
        border.setAlpha(70)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.setPen(QPen(border, 1))
        painter.drawRoundedRect(rect.adjusted(0.5, 0.5, -0.5, -0.5), radius, radius)

        gradient = QConicalGradient(QPointF(rect.center()), -self._angle)
        bright = QColor(self._color)
        bright.setAlpha(255)
        transparent = QColor(self._color)
        transparent.setAlpha(0)
        gradient.setColorAt(0.00, transparent)
        gradient.setColorAt(0.08, transparent)
        gradient.setColorAt(0.18, bright)
        gradient.setColorAt(0.30, transparent)
        gradient.setColorAt(0.55, transparent)
        gradient.setColorAt(0.68, bright)
        gradient.setColorAt(0.80, transparent)
        gradient.setColorAt(1.00, transparent)

        ring_pen = QPen(QBrush(gradient), 2.0)
        painter.setPen(ring_pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(rect.adjusted(1, 1, -1, -1), radius - 1, radius - 1)

        icon_cx = rect.left() + pad_h + icon_size / 2.0
        icon_cy = rect.center().y()
        t = self._twinkle_phase
        scale = 1.0 - 0.3 * abs(math.sin(t * math.pi * 2))
        rotation = t * 90.0

        painter.save()
        painter.translate(icon_cx, icon_cy)
        painter.rotate(rotation)
        painter.scale(scale, scale)

        star_color = QColor(self._color)
        star_color.setAlpha(230)
        painter.setBrush(star_color)
        painter.setPen(Qt.PenStyle.NoPen)

        r = icon_size / 2.0
        star_path = QPainterPath()
        star_path.moveTo(0, -r)
        star_path.cubicTo(r * 0.08, -r * 0.42, r * 0.42, -r * 0.08, r, 0)
        star_path.cubicTo(r * 0.42, r * 0.08, r * 0.08, r * 0.42, 0, r)
        star_path.cubicTo(-r * 0.08, r * 0.42, -r * 0.42, r * 0.08, -r, 0)
        star_path.cubicTo(-r * 0.42, -r * 0.08, -r * 0.08, -r * 0.42, 0, -r)
        painter.drawPath(star_path)
        painter.restore()

        text_color = QColor(self._color)
        painter.setPen(text_color)
        painter.setFont(self.font())
        text_rect = QRectF(
            icon_cx + icon_size / 2.0 + gap,
            rect.top(),
            text_w + 4,
            rect.height(),
        )
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            text,
        )
        painter.end()


# ----------------------------------------------------------------------
# ★★★ КОНТЕЙНЕР-ПЕРЕКЛЮЧАТЕЛЬ ★★★
# ----------------------------------------------------------------------
class StatusIndicatorWidget(QWidget):
    """Показывает BouncingBall при ожидании, Generating при работе."""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.ball = BouncingBallWidget(color="#3b82f6", size=24)
        self.badge = GeneratingBadgeWidget(text="Generating", color="#3b82f6")

        layout.addWidget(self.ball)
        layout.addWidget(self.badge)

        self.set_waiting()

    def set_waiting(self):
        self.ball.show()
        self.ball.set_active(True)
        self.badge.set_active(False)
        self.badge.hide()

    def set_generating(self):
        self.ball.set_active(False)
        self.ball.hide()
        self.badge.set_active(True)
        self.badge.show()

    def set_text(self, text):
        self.badge.set_text(text)


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
        gradient.setColorAt(0.0, QColor("#2563eb"))
        gradient.setColorAt(0.5, QColor("#3b82f6"))
        gradient.setColorAt(1.0, QColor("#60a5fa"))
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
# ГРАДИЕНТНЫЙ LABEL
# ----------------------------------------------------------------------
class GradientLabel(QLabel):
    def __init__(self, text="", parent=None,
                 colors=("#B8E62E", "#6EE8A0", "#4DE8B1", "#B8E62E"),
                 font=None, speed=0.006):
        super().__init__(text, parent)
        if colors and len(colors) > 0:
            self._colors = [QColor(c) for c in colors]
        else:
            self._colors = [
                QColor("#FF0000"),
                QColor("#00FF00"),
                QColor("#0000FF"),
                QColor("#FF00FF"),
                QColor("#00FFFF"),
                QColor("#FFFF00"),
                QColor("#FF0000")
            ]
        self._offset = 0.0
        self._speed = speed
        self._timer = QTimer(self)
        self._timer.setInterval(30)
        self._timer.timeout.connect(self._tick)
        self._active = False
        if font is not None:
            self.setFont(font)

    def start_gradient(self):
        if not self._active:
            self._active = True
            self._timer.start()

    def stop_gradient(self):
        self._active = False
        self._timer.stop()
        self._offset = 0.0
        self.update()

    def _tick(self):
        self._offset = (self._offset + self._speed) % 1.0
        self.update()

    def paintEvent(self, event):
        if not self._active or not self.text():
            super().paintEvent(event)
            return
            
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
        
        rect = self.rect()
        grad = QLinearGradient(rect.left(), 0, rect.left() + rect.width() * 2, 0)
        
        n = len(self._colors)
        for i, color in enumerate(self._colors):
            pos = i / (n - 1)
            grad.setColorAt(pos, color)
            
        grad.setStart(rect.left() - rect.width() * self._offset, 0)
        grad.setFinalStop(rect.left() + rect.width() * (2 - self._offset), 0)
        
        pen = QPen(QBrush(grad), 0)
        painter.setPen(pen)
        painter.setFont(self.font())
        painter.drawText(rect, int(self.alignment()), self.text())
        painter.end()


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
        title_lbl = QLabel("Photoshop WorkStation+ (v5.2)")
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
        <p><b>Новое в v5.2:</b> кнопки управления окном в стиле
        «Выход» / «Отменить выполнение» (фиолетовая / зелёная / красная),
        анимированные чекбоксы 1520 (бирюзовый) и 2030 (#F72F86)
        с пульсацией и мерцанием.</p>
        <p><b>Версия:</b> v5.2</p>
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
            "QPushButton { background-color: #ed053b; color: white;"
            "border: none; border-radius: 15px; padding: 8px;"
            "font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #ff668a; }"
        )
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


# ----------------------------------------------------------------------
# ДИАЛОГ ПРЕДПРОСМОТРА JPG
# ----------------------------------------------------------------------
class ImagePreviewDialog(QDialog):
    EXCLUDED_KEYWORDS = ("фото", "коды", "классы")
    HIDDEN_BTN_KEYWORDS = ("qr", "образц")

    def __init__(self, parent=None, folder=""):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр JPG")
        self.setFixedSize(1280, 800)
        self.setStyleSheet("QDialog { background-color: #161616; color: #FFFFFF; }")
        self.folder = folder

        self.all_items = []
        self.image_paths = []
        self.sort_mode = "name"
        self.sort_ascending = True
        self.subfolder_filter = "__ALL__"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel(f"Папка: {folder}" if folder else "Папка не выбрана")
        title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        title.setStyleSheet("color: #7CC4FF; border: none;")
        title.setWordWrap(True)
        layout.addWidget(title)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        lbl_sort = QLabel("Сортировка:")
        lbl_sort.setStyleSheet(
            "color: #AAAAAA; font-size: 11px; font-weight: bold; border: none;"
        )
        top_row.addWidget(lbl_sort)

        self.btn_sort_name = QPushButton("По имени А-Я")
        self.btn_sort_name.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_sort_name.setMinimumHeight(28)
        self.btn_sort_name.clicked.connect(lambda: self._set_sort("name"))
        top_row.addWidget(self.btn_sort_name)

        self.btn_sort_mtime = QPushButton("По дате ↓")
        self.btn_sort_mtime.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_sort_mtime.setMinimumHeight(28)
        self.btn_sort_mtime.clicked.connect(lambda: self._set_sort("mtime"))
        top_row.addWidget(self.btn_sort_mtime)

        self.lbl_total = QLabel("Всего найдено: 0 фото")
        self.lbl_total.setStyleSheet(
            "color: #CCFF00; font-size: 12px; font-weight: bold; border: none;"
        )
        top_row.addWidget(self.lbl_total)

        lbl_folder = QLabel("Папка:")
        lbl_folder.setStyleSheet(
            "color: #AAAAAA; font-size: 11px; font-weight: bold; border: none;"
        )
        top_row.addWidget(lbl_folder)

        self.combo_subfolder = QComboBox()
        self.combo_subfolder.setFixedHeight(28)
        self.combo_subfolder.setMinimumWidth(240)
        self.combo_subfolder.setStyleSheet(
            "QComboBox { background-color: #1A1A1A; color: #FFFFFF;"
            "border: 1px solid #333333; border-radius: 12px;"
            "padding: 3px 10px; font-size: 11px; font-weight: bold; }"
            "QComboBox:hover { border-color: #7CC4FF; }"
            "QComboBox::drop-down { border: none; width: 22px; }"
            "QComboBox::down-arrow { image: none;"
            "border-left: 4px solid transparent;"
            "border-right: 4px solid transparent;"
            "border-top: 5px solid #7CC4FF; margin-right: 6px; }"
            "QComboBox QAbstractItemView {"
            "  background-color: #1E2229; color: #FFFFFF;"
            "  border: 1px solid #3A4150; border-radius: 10px;"
            "  padding: 4px; outline: none;"
            "  selection-background-color: #1A2740;"
            "  selection-color: #7CC4FF;"
            "}"
        )
        self.combo_subfolder.currentIndexChanged.connect(self._on_subfolder_changed)
        top_row.addWidget(self.combo_subfolder)

        top_row.addStretch(1)
        layout.addLayout(top_row)

        body = QHBoxLayout()
        body.setSpacing(10)

        self.list_widget = QListWidget()
        self.list_widget.setFixedWidth(520)
        self.list_widget.setStyleSheet(
            """
            QListWidget {
                background-color: #1A1A1A; color: #FFFFFF;
                border: 1px solid #333333; border-radius: 8px;
                padding: 6px; font-family: Consolas; font-size: 11px;
            }
            QListWidget::item { padding: 5px 8px; border-radius: 6px; }
            QListWidget::item:selected { background-color: #10243A; color: #7CC4FF; }
            QListWidget::item:hover { background-color: #2A2A2A; }
            QScrollBar:vertical { background: #1A1A1A; width: 8px;
                margin: 4px 2px 4px 0px; border-radius: 4px; }
            QScrollBar::handle:vertical { background: #3A4150;
                min-height: 20px; border-radius: 4px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            """
        )
        self.list_widget.currentRowChanged.connect(self._on_select)
        body.addWidget(self.list_widget)

        self.preview = QLabel("Выберите изображение слева")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setStyleSheet(
            "QLabel { background-color: #0E0E0E; color: #777777;"
            "border: 1px solid #333333; border-radius: 8px;"
            "font-size: 13px; font-weight: bold; }"
        )
        self.preview.setMinimumSize(700, 580)
        self.preview.setScaledContents(False)
        body.addWidget(self.preview, 1)

        layout.addLayout(body, 1)

        info_row = QHBoxLayout()
        self.lbl_info = QLabel("")
        self.lbl_info.setStyleSheet("color: #AAAAAA; font-size: 11px; border: none;")
        info_row.addWidget(self.lbl_info, 1)

        self.btn_open_ps = QPushButton("Открыть проект в PS")
        self.btn_open_ps.setFixedWidth(200)
        self.btn_open_ps.setMinimumHeight(34)
        self.btn_open_ps.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_open_ps.setStyleSheet(
            "QPushButton { background-color: #10243A; color: #7CC4FF;"
            "border: 2px solid #183A5A; border-radius: 15px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #183A5A;"
            "border: 1px solid #7CC4FF; color: #A8D8FF; }"
            "QPushButton:pressed { background-color: #58A8E8;"
            "border: 1px solid #7CC4FF; color: #10243A; }"
            "QPushButton:disabled { background-color: #1A1A1A; color: #555555;"
            "border: 2px solid #2A2A2A; }"
        )
        self.btn_open_ps.clicked.connect(self._open_psd_in_photoshop)
        self.btn_open_ps.setVisible(False)
        info_row.addWidget(self.btn_open_ps)

        btn_close = QPushButton("Закрыть")
        btn_close.setFixedWidth(140)
        btn_close.setMinimumHeight(34)
        btn_close.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_close.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 15px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF4081; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_close.clicked.connect(self.accept)
        info_row.addWidget(btn_close)
        layout.addLayout(info_row)

        self._refresh_sort_buttons()
        self._scan_files()
        self._populate_subfolders()
        self._populate()

    def _scan_files(self):
        self.all_items = []
        if not self.folder or not os.path.isdir(self.folder):
            return

        for dp, dn, filenames in os.walk(self.folder):
            dn[:] = [d for d in dn if not self._is_excluded_folder(d)]
            if self._is_excluded_folder(os.path.basename(dp)):
                continue

            for f in filenames:
                if not f.lower().endswith(".jpg"):
                    continue
                full = os.path.join(dp, f)
                rel = os.path.relpath(full, self.folder)
                folder_name = os.path.basename(dp)
                try:
                    mtime = float(os.path.getmtime(full))
                except Exception:
                    mtime = 0.0
                self.all_items.append((rel, full, mtime, folder_name))

    @classmethod
    def _is_excluded_folder(cls, name):
        if not name:
            return False
        low = name.strip().lower()
        return any(kw in low for kw in cls.EXCLUDED_KEYWORDS)

    @classmethod
    def _is_qr_or_samples_folder(cls, path):
        if not path:
            return False
        parts = path.replace("\\", "/").split("/")
        for part in parts:
            low = part.strip().lower()
            for kw in cls.HIDDEN_BTN_KEYWORDS:
                if kw in low:
                    return True
        return False

    def _populate_subfolders(self):
        self.combo_subfolder.blockSignals(True)
        self.combo_subfolder.clear()
        self.combo_subfolder.addItem("Все папки с фото", "__ALL__")

        folders_with_photos = set()
        for rel, full, mtime, folder_name in self.all_items:
            d = os.path.dirname(rel)
            folders_with_photos.add(d if d else ".")

        final_folders = set()
        for folder in folders_with_photos:
            has_subfolder_with_photos = False
            prefix = folder + os.sep if folder != "." else ""
            for other in folders_with_photos:
                if other != folder and other.startswith(prefix):
                    has_subfolder_with_photos = True
                    break
            if not has_subfolder_with_photos:
                final_folders.add(folder)

        for sf in sorted(final_folders, key=lambda s: s.lower()):
            display = sf if sf != "." else "(корень)"
            self.combo_subfolder.addItem(display, sf)

        idx = self.combo_subfolder.findData(self.subfolder_filter)
        if idx >= 0:
            self.combo_subfolder.setCurrentIndex(idx)
        self.combo_subfolder.blockSignals(False)

    def _on_subfolder_changed(self, _idx):
        self.subfolder_filter = self.combo_subfolder.currentData() or "__ALL__"
        self._populate()

    def _set_sort(self, mode):
        if self.sort_mode == mode:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_mode = mode
            self.sort_ascending = True
        self._refresh_sort_buttons()
        self._populate()

    def _refresh_sort_buttons(self):
        if self.sort_mode == "name":
            name_text = "По имени А-Я" if self.sort_ascending else "По имени Я-А"
            mtime_text = "По дате"
        else:
            name_text = "По имени"
            mtime_text = "По дате ↑" if self.sort_ascending else "По дате ↓"
        self.btn_sort_name.setText(name_text)
        self.btn_sort_mtime.setText(mtime_text)

        active_name_style = (
            "QPushButton { background-color: #3A1030; color: #FF70D8;"
            "border: 1px solid #FF70D8; border-radius: 12px;"
            "padding: 4px 14px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A1850; color: #FFA0E8;"
            "border-color: #FFA0E8; }"
        )
        active_mtime_style = (
            "QPushButton { background-color: #3A2410; color: #FFB74D;"
            "border: 1px solid #FFB74D; border-radius: 12px;"
            "padding: 4px 14px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A3818; color: #FFCC80;"
            "border-color: #FFCC80; }"
        )
        inactive_style = (
            "QPushButton { background-color: #1A1A1A; color: #AAAAAA;"
            "border: 1px solid #333333; border-radius: 12px;"
            "padding: 4px 14px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #262626; color: #FFFFFF;"
            "border-color: #555555; }"
        )
        if self.sort_mode == "name":
            self.btn_sort_name.setStyleSheet(active_name_style)
            self.btn_sort_mtime.setStyleSheet(inactive_style)
        else:
            self.btn_sort_name.setStyleSheet(inactive_style)
            self.btn_sort_mtime.setStyleSheet(active_mtime_style)

    @staticmethod
    def _name_sort_key(filename):
        base = os.path.splitext(filename)[0].strip()
        if base.upper().startswith("PK"):
            return (1, 0, base.upper())
        i = 0
        while i < len(base) and base[i].isdigit():
            i += 1
        if i > 0:
            number = int(base[:i])
            suffix = base[i:].upper()
            return (0, number, suffix)
        return (0, 999999, base.upper())

    def _populate(self):
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        self.image_paths = []
        self.preview.setPixmap(QPixmap())
        self.preview.setText("Выберите изображение слева")
        self.lbl_info.setText("")
        self.btn_open_ps.setVisible(False)

        items = []
        for rel, full, mtime, folder_name in self.all_items:
            if self.subfolder_filter == "__ALL__":
                items.append((rel, full, mtime, folder_name))
            else:
                d = os.path.dirname(rel)
                if self.subfolder_filter == ".":
                    if d == "":
                        items.append((rel, full, mtime, folder_name))
                else:
                    if d == self.subfolder_filter or d.startswith(
                        self.subfolder_filter + os.sep
                    ):
                        items.append((rel, full, mtime, folder_name))

        import datetime
        reverse = not self.sort_ascending

        if self.sort_mode == "name":
            items.sort(
                key=lambda x: self._name_sort_key(os.path.basename(x[1])),
                reverse=reverse,
            )
        else:
            items.sort(
                key=lambda x: float(x[2]) if x[2] is not None else 0.0,
                reverse=reverse,
            )

        for rel, full, mtime, folder_name in items:
            self.image_paths.append(full)
            base_name = os.path.basename(full)

            if mtime:
                try:
                    date_str = datetime.datetime.fromtimestamp(
                        float(mtime)
                    ).strftime("%d.%m.%Y %H:%M:%S")
                except Exception:
                    date_str = "—"
            else:
                date_str = "—"

            name_part = base_name if len(base_name) <= 35 else base_name[:32] + "..."
            padding = " " * max(1, 40 - len(name_part))
            item = QListWidgetItem(f"{name_part}{padding}{date_str}")
            item.setData(Qt.ItemDataRole.UserRole, full)
            item.setToolTip(f"{full}\nИзменён: {date_str}")
            self.list_widget.addItem(item)

        self.lbl_total.setText(f"Всего найдено: {len(items)} фото")
        if not items:
            self.preview.setText("В выбранной папке нет .jpg файлов")

        self.list_widget.blockSignals(False)

    def _on_select(self, row):
        if row < 0 or row >= len(self.image_paths):
            return
        path = self.image_paths[row]
        pix = QPixmap(path)
        if pix.isNull():
            self.preview.setText("Не удалось открыть файл")
            self.lbl_info.setText(path)
            return
        avail = self.preview.size()
        scaled = pix.scaled(
            avail,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.preview.setPixmap(scaled)
        try:
            size_kb = os.path.getsize(path) / 1024
            import datetime
            mtime_str = datetime.datetime.fromtimestamp(
                os.path.getmtime(path)
            ).strftime("%d.%m.%Y %H:%M:%S")
            self.lbl_info.setText(
                f"{os.path.basename(path)}  •  {pix.width()}×{pix.height()}  "
                f"•  {size_kb:.0f} КБ  •  Изменён: {mtime_str}"
            )
        except Exception:
            self.lbl_info.setText(os.path.basename(path))

        if self._is_qr_or_samples_folder(path):
            self.btn_open_ps.setVisible(False)
        else:
            self.btn_open_ps.setVisible(True)

    def _find_psd_for_current(self):
        row = self.list_widget.currentRow()
        if row < 0 or row >= len(self.image_paths):
            return None

        jpg_path = self.image_paths[row]
        jpg_dir = os.path.dirname(jpg_path)
        jpg_base = os.path.splitext(os.path.basename(jpg_path))[0].strip()

        if not jpg_base:
            return None

        prefix = jpg_base + " "

        if not os.path.isdir(jpg_dir):
            return None

        try:
            entries = os.listdir(jpg_dir)
        except Exception:
            return None

        for f in entries:
            if not f.lower().endswith(".psd"):
                continue
            if f.startswith(prefix):
                return os.path.join(jpg_dir, f)

        return None

    def _open_psd_in_photoshop(self):
        psd_path = self._find_psd_for_current()
        if not psd_path:
            row = self.list_widget.currentRow()
            jpg_name = "?"
            if 0 <= row < len(self.image_paths):
                jpg_name = os.path.basename(self.image_paths[row])
            self.lbl_info.setText(
                f"❌ Не найден .psd, начинающийся с '{jpg_name} ' рядом с этим файлом"
            )
            self.lbl_info.setStyleSheet(
                "color: #FF5555; font-size: 11px; font-weight: bold; border: none;"
            )
            return

        try:
            self.lbl_info.setText(f"→ Открываю: {psd_path}")
            self.lbl_info.setStyleSheet(
                "color: #7CC4FF; font-size: 11px; font-weight: bold; border: none;"
            )
        except Exception:
            pass

        try:
            os.startfile(psd_path)
        except Exception:
            try:
                subprocess.Popen(["start", "", psd_path], shell=True)
            except Exception as e2:
                self.lbl_info.setText(f"❌ Ошибка открытия: {e2}")
                self.lbl_info.setStyleSheet(
                    "color: #FF5555; font-size: 11px; font-weight: bold; border: none;"
                )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        row = self.list_widget.currentRow()
        if row >= 0:
            self._on_select(row)


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
            "QPushButton { background-color: #103A18; color: #88FFA8;"
            "border: 2px solid #185A28; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #185A28;"
            "border: 1px solid #B0FFC8; color: #B0FFC8; }"
            "QPushButton:pressed { background-color: #58E878;"
            "border: 1px solid #88FFA8; color: #0A1A0E; }"
        )
        btn_default.setMinimumHeight(38)
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
            "QPushButton { background-color: #10243A; color: #7CC4FF;"
            "border: 2px solid #183A5A; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #183A5A;"
            "border: 1px solid #A8D8FF; color: #A8D8FF; }"
            "QPushButton:pressed { background-color: #58A8E8;"
            "border: 1px solid #7CC4FF; color: #10243A; }"
        )
        btn_alt.setMinimumHeight(38)
        btn_alt.clicked.connect(self._choose_alternate)
        ca_layout.addWidget(btn_alt)
        layout.addWidget(card_alt)
        layout.addSpacing(10)

        btn_cancel = QPushButton("Назад (оставить текущие пути)")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 0px;"
            "padding: 8px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF80AB; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_cancel.setMinimumHeight(38)
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
                "border: none; border-radius: 15px; padding: 8px;"
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
            "QPushButton { background-color: #3A3410; color: #FFE55C;"
            "border: 2px solid #5A5018; border-radius: 0px;"
            "padding: 9px 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #5A5018;"
            "border: 1px solid #FFE55C; color: #FFF088; }"
            "QPushButton:pressed { background-color: #FFD600;"
            "border: 1px solid #FFE55C; color: #1A1A00; }"
        )
        btn_defaults.setMinimumHeight(38)
        btn_defaults.clicked.connect(self._reset_to_defaults)
        btn_row.addWidget(btn_defaults, 1)

        btn_save = QPushButton("Сохранить")
        btn_save.setStyleSheet(
            "QPushButton { background-color: #103A18; color: #88FFA8;"
            "border: 2px solid #185A28; border-radius: 0px;"
            "padding: 9px 6px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #185A28;"
            "border: 1px solid #B0FFC8; color: #B0FFC8; }"
            "QPushButton:pressed { background-color: #58E878;"
            "border: 1px solid #88FFA8; color: #0A1A0E; }"
        )
        btn_save.setMinimumHeight(38)
        btn_save.clicked.connect(self.accept)
        btn_row.addWidget(btn_save, 1)

        btn_cancel = QPushButton("Назад")
        btn_cancel.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 0px;"
            "padding: 9px 6px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF80AB; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_cancel.setMinimumHeight(38)
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
# ПОТОК ОБРАБОТКИ (Photoshop)
# ----------------------------------------------------------------------
class WorkerThread(QThread):
    progress_signal = pyqtSignal(int, int)
    file_status_signal = pyqtSignal(str, str)
    log_signal = pyqtSignal(str)
    card_signal = pyqtSignal(str, str, str)
    finished_signal = pyqtSignal(int, str)
    typing_signal = pyqtSignal(bool)

    def __init__(self, mode, files, scripts_dir, script_filename, script_title,
                 school_name, is_visible=False, qr_frame_path=None,
                 qr_folder_path=None, action_set=None, actions_list=None,
                 fast_mode=True, auto_alert=True, frame_1520_path=None,
                 frame_2030_path=None, target_folder=None):
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
        self.target_folder = target_folder

    def _build_jsx_path(self):
        original = os.path.join(self.scripts_dir, self.script_filename)
        
        if self.mode == "insert_qr":
            if not os.path.exists(original):
                return original
            try:
                with open(original, "r", encoding="utf-8") as f:
                    content = f.read()
                def esc(p):
                    return p.replace("\\", "\\\\")

                school_dir = self.target_folder or self.school_name
                qr_path = os.path.join(school_dir, "коды")

                target_path = school_dir
                psd_path = ""
                if self.files:
                    psd_path = self.files[0]
                    target_path = os.path.dirname(psd_path)

                content = content.replace("__QR_FOLDER_PATH__", esc(qr_path))
                content = content.replace("__TARGET_FOLDER__", esc(target_path))
                content = content.replace("__PSD_FRAME_PATH__", esc(self.frame_2030_path or ""))
                content = content.replace("__PSD_PATH__", esc(psd_path))
                content = content.replace("__SCHOOL_DIR__", esc(school_dir))

                tmp_path = os.path.join(
                    tempfile.gettempdir(),
                    f"_insert_qr_{os.getpid()}_{id(self)}.jsx",
                )
                with open(tmp_path, "w", encoding="utf-8") as f:
                    f.write(content)
                return tmp_path
            except Exception as e:
                self.log_signal.emit(f"Ошибка генерации временного JSX: {e}")
                return original

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
                self.file_status_signal.emit("Обновление", "")
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

            if self.mode == "insert_qr":
                self._run_insert_qr()
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
                    self.file_status_signal.emit("Обновление", "Выполнение...")
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

            for idx, psd_path in enumerate(self.files, start=1):
                if self.is_cancelled:
                    break
                file_name = os.path.basename(psd_path)
                parent_dir = os.path.basename(os.path.dirname(psd_path))
                self.file_status_signal.emit(f"Обновление {idx}/{total}", file_name)
                self.typing_signal.emit(True)
                try:
                    doc = self.ps_app.Open(psd_path)
                    self.ps_app.DoJavaScriptFile(jsx_path)
                    doc.Close(1)
                    processed_count += 1
                    
                    self.card_signal.emit(file_name, "done", "Обновлён")
                    self.log_signal.emit(f"✅ Файл {file_name} Обновлён")
                except Exception as e:
                    self.card_signal.emit(file_name, "error", "Не удалось обработать")
                    self.log_signal.emit(f"Ошибка при обработке {file_name}: {e}")
                self.typing_signal.emit(False)
                self.progress_signal.emit(idx, total)

            if not self.is_cancelled:
                self.finished_signal.emit(processed_count, self.script_title)

        except Exception as fatal:
            self.log_signal.emit(f"__ERROR__КРИТИЧЕСКАЯ ОШИБКА ПОТОКА: {fatal}")
            try:
                self.finished_signal.emit(0, self.script_title)
            except Exception:
                pass

    def _run_insert_qr(self):
        total = len(self.files)
        if total == 0:
            self.log_signal.emit("__ERROR__Не выбраны PSD-файлы для вставки QR.")
            self.finished_signal.emit(0, self.script_title)
            return
        processed = 0
        errors = 0
        template_path = os.path.join(self.scripts_dir, self.script_filename)
        if not os.path.exists(template_path):
            self.log_signal.emit(f"__ERROR__Шаблон JSX не найден: {template_path}")
            self.finished_signal.emit(0, self.script_title)
            return
        school_dir = self.target_folder or self.school_name
        qr_path = os.path.join(school_dir, "коды")

        for idx, psd_path in enumerate(self.files, start=1):
            if self.is_cancelled:
                break
            file_name = os.path.basename(psd_path)
            self.file_status_signal.emit(f"Обновление {idx}/{total}", file_name)
            self.typing_signal.emit(True)
            psd_doc = None
            try:
                try:
                    _ = self.ps_app.Version
                except Exception:
                    self.log_signal.emit("__ERROR__Photoshop перестал отвечать. Прерывание.")
                    break

                # ★ ОТКРЫВАЕМ PSD ЧЕРЕЗ COM — без этого JSX не увидит документ
                psd_doc = self.ps_app.Open(psd_path)

                with open(template_path, "r", encoding="utf-8") as f:
                    content = f.read()

                def esc(p):
                    return str(p).replace("\\", "\\\\").replace('"', '\\"')

                content = content.replace("__QR_FOLDER_PATH__", esc(qr_path))
                content = content.replace("__PSD_PATH__", esc(psd_path))
                content = content.replace("__TARGET_FOLDER__", esc(os.path.dirname(psd_path)))
                content = content.replace("__SCHOOL_DIR__", esc(school_dir))

                tmp_jsx = os.path.join(
                    tempfile.gettempdir(),
                    f"_insert_qr_{os.getpid()}_{id(self)}_{idx}.jsx",
                )
                with open(tmp_jsx, "w", encoding="utf-8") as f:
                    f.write(content)

                try:
                    self.ps_app.DoJavaScriptFile(tmp_jsx)
                    # ★ JSX сам сохраняет targetDoc.save(), поэтому после вызова
                    # закрываем без сохранения (чтобы не было двойного save)
                    try:
                        psd_doc.Close(1)  # 1 = DoNotSaveChanges
                    except Exception:
                        pass
                    processed += 1
                    self.card_signal.emit(file_name, "done", "QR вставлен")
                    self.log_signal.emit(f"✅ QR-код успешно вставлен в {file_name}")
                except Exception as e:
                    errors += 1
                    self.log_signal.emit(f"__WARN__Ошибка при обработке {file_name}: {e}")
                    self.card_signal.emit(file_name, "error", "Ошибка вставки QR")
                    try:
                        if psd_doc is not None:
                            psd_doc.Close(1)
                    except Exception:
                        pass
                finally:
                    try:
                        os.remove(tmp_jsx)
                    except Exception:
                        pass
            except Exception as e:
                errors += 1
                self.log_signal.emit(f"__ERROR__Не удалось подготовить JSX для {file_name}: {e}")
                try:
                    if psd_doc is not None:
                        psd_doc.Close(1)
                except Exception:
                    pass
            self.typing_signal.emit(False)
            self.progress_signal.emit(idx, total)

        if not self.is_cancelled:
            self.log_signal.emit(f"__INFO__📊 Итого: обработано {processed} из {total}, ошибок: {errors}")
            self.finished_signal.emit(processed, self.script_title)

    def _run_smart_frames(self):
        total = len(self.files)
        processed_count = 0
        errors_count = 0
        skipped_count = 0

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
                self.card_signal.emit(file_name, "error", "Пропущен (нет папки 1520/2030)")
                skipped_count += 1
                self.progress_signal.emit(idx, total)
                continue
            self.file_status_signal.emit(f"Обновление {idx}/{total} ({kind})", file_name)
            self.typing_signal.emit(True)
            try:
                _ = self.ps_app.Version
            except Exception:
                self.log_signal.emit("__ERROR__Photoshop перестал отвечать. Прерывание.")
                break
            jsx_for_file = self._build_smart_jsx_for_file(template_path, psd_path, kind)
            if not jsx_for_file:
                errors_count += 1
                self.card_signal.emit(file_name, "error", "Ошибка подготовки JSX")
                self.progress_signal.emit(idx, total)
                self.typing_signal.emit(False)
                continue
            try:
                self.ps_app.DoJavaScriptFile(jsx_for_file)
                processed_count += 1
                self.card_signal.emit(file_name, "done", f"Обновлён ({kind})")
                self.log_signal.emit(f"✅ Файл {file_name} Обновлён (рамка {kind})")
            except Exception as e:
                errors_count += 1
                self.card_signal.emit(file_name, "error", "Не удалось обработать")
                self.log_signal.emit(f"Ошибка при обработке {file_name}: {e}")
            finally:
                try:
                    os.remove(jsx_for_file)
                except Exception:
                    pass
            self.typing_signal.emit(False)
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
            self.file_status_signal.emit(f"Обновление {idx}/{total}", file_name)
            self.typing_signal.emit(True)
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
                self.card_signal.emit(file_name, "done", "Обновлён")
                self.log_signal.emit(f"✅ Файл {file_name} Обновлён (Batch Actions)")
            except Exception as e:
                errors_count += 1
                self.card_signal.emit(file_name, "error", "Не удалось обработать")
                self.log_signal.emit(f"Ошибка при обработке {file_name}: {e}")
                try:
                    if doc is not None:
                        doc.Close(1)
                except Exception:
                    pass
            self.typing_signal.emit(False)
            self.progress_signal.emit(idx, total)

        if not self.is_cancelled:
            self.log_signal.emit(
                f"__INFO__📊 Итого: Обновлено {processed_count} из {total}, "
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
# ★★★ ВИДЖЕТ СООБЩЕНИЯ В ЧАТЕ ★★★
# ----------------------------------------------------------------------
class ChatMessageWidget(QWidget):
    """Одно сообщение в чате: аватарка слева, пурпурный пузырь с белым текстом,
    центрированная метка 'доставлено' под ним с анимацией shimmer."""
    def __init__(self, text, avatar_pixmap=None, show_delivered=True, parent=None):
        super().__init__(parent)
        self.text = text
        self.show_delivered = show_delivered

        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        # --- Круглая аватарка ---
        self.avatar_label = QLabel()
        self.avatar_label.setFixedSize(30, 30)
        self.avatar_label.setStyleSheet("border: none; background: transparent;")

        if avatar_pixmap and not avatar_pixmap.isNull():
            rounded = QPixmap(30, 30)
            rounded.fill(Qt.GlobalColor.transparent)
            p = QPainter(rounded)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            path = QPainterPath()
            path.addEllipse(0, 0, 30, 30)
            p.setClipPath(path)
            p.drawPixmap(
                0, 0,
                avatar_pixmap.scaled(
                    30, 30,
                    Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                    Qt.TransformationMode.SmoothTransformation
                )
            )
            p.end()
            self.avatar_label.setPixmap(rounded)
        else:
            fallback = QPixmap(36, 36)
            fallback.fill(Qt.GlobalColor.transparent)
            p = QPainter(fallback)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            p.setBrush(QColor("#7CFFCB"))
            p.setPen(Qt.PenStyle.NoPen)
            p.drawEllipse(0, 0, 36, 36)
            p.setPen(QColor("#0A1A0E"))
            p.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
            p.drawText(fallback.rect(), Qt.AlignmentFlag.AlignCenter, "?")
            p.end()
            self.avatar_label.setPixmap(fallback)

        layout.addWidget(self.avatar_label, 0, Qt.AlignmentFlag.AlignTop)

        # --- Контейнер для пузыря и Delivered ---
        msg_container = QVBoxLayout()
        msg_container.setSpacing(4)
        msg_container.setAlignment(Qt.AlignmentFlag.AlignLeft)

        # Используем QLabel с WordWrap — он идеально переносит текст
        self.bubble = QLabel()
        self.bubble.setTextFormat(Qt.TextFormat.RichText)
        self.bubble.setWordWrap(True)
        
        # Текст будет автоматически переноситься и занимать всю доступную ширину
        self.bubble.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        safe_text = html.escape(text).replace("\n", "<br>")
        self.bubble.setText(
            f"<div style='color: white; font-family: \"Segoe UI\"; "
            f"font-size: 13px; line-height: 1.3;'>{safe_text}</div>"
        )
        
        self.bubble.setStyleSheet("""
            QLabel {
                background-color: #3030fc;
                color: white;
                border-radius: 14px;
                padding: 8px 14px;
                border: none;
            }
        """)

        # Чтобы текст не растягивался на всю ширину, а переносился, 
        # нужно позволить виджету сжиматься, но не растягиваться больше нужного
        self.bubble.setMaximumWidth(650)  # Максимальная ширина пузыря сообщения
        msg_container.addWidget(self.bubble, 0, Qt.AlignmentFlag.AlignLeft)

        if show_delivered:
            self.delivered_label = QLabel("Доставлено")
            self.delivered_label.setStyleSheet(
                "color: #CCCCCC; font-size: 10px; border: none;"
                "background: transparent; font-family: 'Segoe UI';"
            )
            self.delivered_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            msg_container.addWidget(
                self.delivered_label, 0, Qt.AlignmentFlag.AlignHCenter
            )

        msg_container.setSizeConstraint(QVBoxLayout.SizeConstraint.SetMinimumSize)
        layout.addLayout(msg_container, 1)

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
        self.drag_pos = None

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint | Qt.WindowType.Window
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        self.setWindowTitle("Photoshop WorkStation+ v5.2")
        self.setFixedSize(1480, 760)

        # --- Загрузка иконки и аватарки ---
        self.chat_avatar_pixmap = None
        possible_names = [
            "avatarka3.jpg",
            "iconapp_Photoshop Workstation+.png",
            "iconapp.png"
        ]
        
        # Ищем иконку везде, где только можно
        found_icon_path = None
        for name in possible_names:
            # 1. В папке со скриптом (resource_path)
            p = resource_path(name)
            if os.path.exists(p):
                found_icon_path = p
                break
            # 2. В текущей рабочей директории
            p = os.path.join(os.getcwd(), name)
            if os.path.exists(p):
                found_icon_path = p
                break
            # 3. В папке с exe-файлом (если собрано в PyInstaller)
            if getattr(sys, "frozen", False):
                p = os.path.join(os.path.dirname(sys.executable), name)
                if os.path.exists(p):
                    found_icon_path = p
                    break
        
        if found_icon_path:
            self.setWindowIcon(QIcon(found_icon_path))
            # Важно: QPixmap создается только после QApplication!
            self.chat_avatar_pixmap = QPixmap(found_icon_path)
        else:
            print("Иконка не найдена! Аватарка будет заглушкой.")

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

        self.btn_win_min = None
        self.btn_win_max = None
        self.btn_win_close = None

        # Для индикатора "печатает..."
        self._typing_widget = None
        self._typing_hide_timer = QTimer(self)
        self._typing_hide_timer.setSingleShot(True)
        self._typing_hide_timer.timeout.connect(self._hide_typing_indicator)

        self.init_ui()
        self.center_window()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            widget = self.childAt(event.position().toPoint())
            if widget is not None and isinstance(widget, (QPushButton, QComboBox,
                                                          QCheckBox, QTextEdit,
                                                          QListWidget, QProgressBar)):
                return
            self.drag_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.drag_pos is not None and event.buttons() == Qt.MouseButton.LeftButton:
            delta = event.globalPosition().toPoint() - self.drag_pos
            self.move(self.pos() + delta)
            self.drag_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.drag_pos = None

    def _toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
            if self.btn_win_max is not None:
                self.btn_win_max.set_icon_kind("max")
        else:
            self.showMaximized()
            if self.btn_win_max is not None:
                self.btn_win_max.set_icon_kind("restore")

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
                folders.sort(
                    key=lambda name: os.path.getmtime(
                        os.path.join(self.default_schools_dir, name)
                    ),
                    reverse=True
                )
                return folders
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

    def open_profile_dialog(self):
        default_paths = dict(self.DEFAULT_PATHS)
        alternate_paths = dict(self.ALTERNATE_PATHS)
        dlg = ResetPathsDialog(
            self,
            default_paths=default_paths,
            alternate_paths=alternate_paths,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        if dlg.chosen_set == "default":
            target = default_paths
        elif dlg.chosen_set == "alternate":
            target = alternate_paths
        else:
            return
        self.default_schools_dir = target.get("schools_dir", self.default_schools_dir)
        self.default_frame_dir = target.get("frame_dir", self.default_frame_dir)
        self.scripts_dir = target.get("scripts_dir", self.scripts_dir)
        save_config({
            "schools_dir": self.default_schools_dir,
            "frame_dir": self.default_frame_dir,
            "scripts_dir": self.scripts_dir,
            "ambient_video": self.ambient_video_path,
        })
        self._refresh_frame_paths()
        self._load_ambient_video()
        self.refresh_school_list()
        self.append_log("__OK__Профиль путей успешно изменён.")

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
            for c in (resource_path("ambient_vid1.mp4"),
                      os.path.join(os.getcwd(), "ambient_vid1.mp4"),
                      os.path.join(os.path.dirname(sys.executable), "ambient_vid1.mp4")):
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
        central_widget.setStyleSheet(
            "QWidget { background-color: #121212; border-radius: 20px; }"
        )

        self.lbl_bg = QLabel(central_widget)
        self.lbl_bg.setGeometry(0, 0, 1480, 760)
        self.lbl_bg.setStyleSheet(
            "background-color: #121212; border-radius: 20px;"
        )
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
        welcome_layout.setContentsMargins(260, 0, 0, 0)
        welcome_layout.setSpacing(5)
        welcome_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome_layout.addStretch(1)

        lbl_welcome_part1 = QLabel("Добро пожаловать в")
        font_welcome = QFont("Segoe UI", 13, QFont.Weight.Bold)
        lbl_welcome_part1.setFont(font_welcome)
        lbl_welcome_part1.setStyleSheet(
            "color: #FFFFFF; border: none; background: transparent;"
        )

        self.lbl_welcome_part2 = GradientLabel(
            "Photoshop WorkStation+",
            colors=("#FF70D8", "#5ca5ff", "#fa00ab", "#FF70D8"),
            font=font_welcome,
            speed=0.012,
        )
        self.lbl_welcome_part2.setStyleSheet(
            "border: none; background: transparent;"
        )
        self.lbl_welcome_part2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_welcome_part2.start_gradient()

        welcome_layout.addWidget(lbl_welcome_part1)
        welcome_layout.addWidget(self.lbl_welcome_part2)
        welcome_layout.addStretch(1)
        header_layout.addWidget(welcome_container, 1)

        right_side = QWidget()
        right_side.setFixedWidth(SIDE_WIDTH + 260)
        right_side_layout = QHBoxLayout(right_side)
        right_side_layout.setContentsMargins(0, 0, 0, 0)
        right_side_layout.setSpacing(8)
        right_side_layout.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )

        btn_profile = CircleIconButton(
            icon_kind="user", style_key="user",
            tooltip="Сменить профиль путей",
        )
        btn_profile.clicked.connect(self.open_profile_dialog)
        right_side_layout.addStretch()
        right_side_layout.addWidget(btn_profile)

        btn_paths = CircleIconButton(
            icon_kind="settings", style_key="settings",
            tooltip="Настройки путей",
        )
        btn_paths.clicked.connect(self.open_path_settings)
        right_side_layout.addWidget(btn_paths)

        btn_about = CircleIconButton(
            icon_kind="about", style_key="about",
            tooltip="О программе",
        )
        btn_about.clicked.connect(self.open_about_dialog)
        right_side_layout.addWidget(btn_about)

        self.btn_win_min = CircleWindowButton(icon_kind="min", style_key="min")
        self.btn_win_min.setToolTip("Свернуть")
        self.btn_win_min.clicked.connect(self.showMinimized)
        right_side_layout.addWidget(self.btn_win_min)

        self.btn_win_max = CircleWindowButton(icon_kind="max", style_key="max")
        self.btn_win_max.setToolTip("Развернуть")
        self.btn_win_max.clicked.connect(self._toggle_maximize)
        right_side_layout.addWidget(self.btn_win_max)

        self.btn_win_close = CircleWindowButton(icon_kind="close", style_key="close")
        self.btn_win_close.setToolTip("Закрыть")
        self.btn_win_close.clicked.connect(self.close)
        right_side_layout.addWidget(self.btn_win_close)

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
            "border-radius: 0px; padding: 6px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: rgba(255, 255, 255, 30);"
            "border: 1px solid rgba(255, 255, 255, 60); }"
            "QPushButton:disabled { background-color: rgba(255, 255, 255, 5);"
            "color: #AAAAAA; border: 1px solid rgba(255, 255, 255, 15); }"
        )
        text_color = "white"
        sub_bg = "#1A1A1A"
        header_color = "#00BFFF"

        # ==============================================================
        # КОЛОНКА 1 — РАБОЧАЯ ПАПКА (ШКОЛЫ)
        # ==============================================================
        card_school = QFrame()
        card_school.setStyleSheet(card_style)
        layout_school = QVBoxLayout(card_school)
        layout_school.setContentsMargins(15, 15, 15, 15)
        layout_school.setSpacing(10)

        title_school = QLabel("1. Рабочая папка")
        title_school.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title_school.setStyleSheet(f"color: {header_color}; border: none;")
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
        layout_school.addSpacing(10)

        lbl_select = QLabel("📂 Выберите школу:")
        lbl_select.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        lbl_select.setStyleSheet("color: #8391fc; border: none;")
        lbl_select.setAlignment(Qt.AlignmentFlag.AlignCenter)
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

        self.lbl_sub = QLabel("Обрабатывать подпапки:")
        self.lbl_sub.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.lbl_sub.setStyleSheet("color: #ffc642; border: none;")
        self.lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_sub.addWidget(self.lbl_sub)

        layout_checks = QHBoxLayout()
        layout_checks.setSpacing(30)

        self.chk_1520 = PulseCheckBox(
            "Папка 1520",
            base_color="#B4FF2E",
            hover_color="#D4FF80",
            dim_bg="#26360E",
            hover_bg="#3C5A18",
            flash_color="#B4FF2E",
        )
        self.chk_1520.stateChanged.connect(self.update_file_list)

        self.chk_2030 = PulseCheckBox(
            "Папка 2030",
            base_color="#FF2EC9",
            hover_color="#FF80DC",
            dim_bg="#3A0E2E",
            hover_bg="#5A1850",
            flash_color="#FF2EC9",
        )
        self.chk_2030.stateChanged.connect(self.update_file_list)
        self.chk_2030.stateChanged.connect(self._update_insert_qr_button_state)

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
        layout_school.addSpacing(8)

        btn_preview = QPushButton("🔍 Предпросмотр JPG")
        btn_preview.setStyleSheet(
            "QPushButton { background-color: #0F282E; color: #00E5FF;"
            "border: 1px solid #194A54; border-radius: 15px;"
            "padding: 7px; font-weight: bold; font-size: 11px; }"
            "QPushButton:hover { background-color: #194A54;"
            "border-color: #00E5FF; color: #80F2FF; }"
        )
        btn_preview.clicked.connect(self.open_preview_dialog)
        layout_school.addWidget(btn_preview)

        self.lbl_path_info = QLabel("Школа не выбрана")
        self.lbl_path_info.setStyleSheet(
            "color: #FF4444; border: none; font-size: 12px; font-weight: bold;"
        )
        self.lbl_path_info.setWordWrap(True)
        self.lbl_path_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_school.addWidget(self.lbl_path_info)

        self.lbl_count_info = GradientLabel(
            "",
        colors=("#5CE8FF", "#3187FF", "#FF57D5", "#5CE8FF"),

            font=QFont("Segoe UI", 11, QFont.Weight.Bold),
            speed=0.015,
        )
        self.lbl_count_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_count_info.setFrameShape(QFrame.Shape.NoFrame)
        _pal = self.lbl_count_info.palette()
        _pal.setColor(self.lbl_count_info.foregroundRole(), QColor("#5CE8FF"))
        self.lbl_count_info.setPalette(_pal)
        self.lbl_count_info.setStyleSheet(
            "border: none; background: transparent;"
        )
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
        title_scripts.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title_scripts.setStyleSheet(f"color: {header_color}; border: none;")
        title_scripts.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_scripts.addWidget(title_scripts)
        layout_scripts.addStretch(1)

        self.lbl_hint_scripts = QLabel(
            "← Сначала выберите школу\nв первой колонке"
        )
        self.lbl_hint_scripts.setWordWrap(True)
        self.lbl_hint_scripts.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_hint_scripts.setStyleSheet(
            "QLabel { color: #B8E858; background-color: #232A14;"
            "border: 1px solid #3A4A1A; border-radius: 0px; padding: 8px;"
            "font-family: 'Segoe UI'; font-size: 13px; font-weight: bold; }"
        )
        layout_scripts.addWidget(self.lbl_hint_scripts)

        buttons_data = [
            ("1. СОЗДАТЬ ПАПКИ ШКОЛЫ",
             "create_folders", "", "1. Создать структуру папок", True),
            ("2. ВСТАВИТЬ QR-КОДЫ ВО ВСЕ КЛАССЫ 2030",
             "insert_qr", "26_QR_Insert_Into_Frames.jsx", "2. QR-коды во все классы 2030", True),
            ("3. БЫСТРЫЙ ЭКСПОРТ JPG (2_FAST_Export jpg)",
            "folder_fast", "2_FAST_Export jpg.jsx", "3. ЭКСПОРТ JPG", True),
            ("4. ДОБАВИТЬ QR-КОДЫ В РАМКУ 2030",
             "qr", "23_QR_add_FRAME2030.jsx", "4. Добавление QR-кодов в рамку 2030", True),
            ("5. БЫСТРАЯ ПАКЕТНАЯ ОБРАБОТКА (25_batch_run_action)",
             "batch_actions", "25_batch_run_action.jsx", "5. BATCH ACTIONS", True),
            ("6. ВСТАВКА РАМОК (ДОБАВИТЬ РАМКИ) (24_AddFramesSmart)",
             "smart_frames", "24_AddFramesSmart.jsx", "6. Умная вставка рамок", True),
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
            if mode == "create_folders":
                btn.setEnabled(True)
            else:
                btn.setEnabled(False)
            pal = script_palettes[idx % len(script_palettes)]
            border_wrap = RGBBorderWidget(btn, palette=pal["colors"],
                                          speed=pal["speed"], radius=0)
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

        layout_scripts.addSpacing(10)
        btn_exit = QPushButton("Выход")
        btn_exit.setMinimumHeight(36)
        btn_exit.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_exit.setStyleSheet(
            "QPushButton { background-color: #3A0A1A; color: #FF4081;"
            "border: 2px solid #5A1028; border-radius: 0px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #5A1028;"
            "border: 1px solid #FF4081; color: #FF80AB; }"
            "QPushButton:pressed { background-color: #E91E63;"
            "border: 1px solid #FF4081; color: #FFFFFF; }"
        )
        btn_exit.clicked.connect(self.close)
        layout_scripts.addWidget(btn_exit)

        self.btn_cancel = QPushButton("Отменить выполнение / закрыть PS")
        self.btn_cancel.setMinimumHeight(36)
        self.btn_cancel.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_cancel.setStyleSheet(
            "QPushButton { background-color: #10243A; color: #7CC4FF;"
            "border: 2px solid #183A5A; border-radius: 0px;"
            "padding: 8px; font-weight: bold; font-size: 12px; }"
            "QPushButton:hover { background-color: #183A5A;"
            "border: 1px solid #7CC4FF; color: #A8D8FF; }"
            "QPushButton:pressed { background-color: #58A8E8;"
            "border: 1px solid #7CC4FF; color: #10243A; }"
        )
        self.btn_cancel.clicked.connect(self.cancel_process)
        layout_scripts.addWidget(self.btn_cancel)
        layout_scripts.addStretch(1)

        frame_footer = QFrame()
        frame_footer.setStyleSheet("border: none; background: transparent;")
        layout_footer = QVBoxLayout(frame_footer)
        layout_footer.setContentsMargins(0, 5, 0, 0)
        layout_footer.setSpacing(2)
        layout_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)

        authors_row = QHBoxLayout()
        authors_row.setContentsMargins(0, 0, 0, 0)
        authors_row.setSpacing(0)
        authors_row.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_dev = QLabel("made by ")
        lbl_dev.setStyleSheet(
            f"color: {text_color}; font-weight: bold; font-size: 12px;"
        )
        lbl_author = QLabel(
            '<a href="https://github.com/DirtSmoke44" '
            'style="color:#9cfc8b; text-decoration:underline;">'
            '@DirtSmoke44</a>'
            '<span style="color:' + text_color + ';"> and </span>'
            '<a href="https://github.com/JohnnySuon" '
            'style="color:#d17dff; text-decoration:underline;">'
            '@JohnnySuon</a>'
        )
        lbl_author.setOpenExternalLinks(False)
        lbl_author.setStyleSheet(
            "font-weight: bold; font-size: 12px; border: none;"
        )
        lbl_author.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        lbl_author.linkActivated.connect(self._open_github)
        authors_row.addWidget(lbl_dev)
        authors_row.addWidget(lbl_author)
        layout_footer.addLayout(authors_row)

        lbl_copyright = QLabel("All Rights Reserved. 2026 ©")
        lbl_copyright.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_copyright.setStyleSheet(
            "color: #888888; font-size: 11px; font-weight: bold; border: none;"
        )
        layout_footer.addWidget(lbl_copyright)

        layout_scripts.addWidget(frame_footer)

        # ==============================================================
        # КОЛОНКА 3 — СТАТУС ВЫПОЛНЕНИЯ (ЧАТ)
        # ==============================================================
        card_status = QFrame()
        card_status.setStyleSheet(card_style)
        layout_status = QVBoxLayout(card_status)
        layout_status.setContentsMargins(18, 16, 18, 16)
        layout_status.setSpacing(10)

        self.lbl_status_header = QLabel("3. Статус выполнения")
        self.lbl_status_header.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.lbl_status_header.setStyleSheet(
            f"color: {header_color}; border: none; background: transparent;"
        )
        self.lbl_status_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_status.addWidget(self.lbl_status_header)

        # --- Контейнер для чата (вместо QTextEdit) ---
        self.chat_scroll = QScrollArea()
        self.chat_scroll.setWidgetResizable(True)
        self.chat_scroll.setStyleSheet("""
            QScrollArea {
                background-color: #171717;
                border: 1px solid #2E2E2E;
                border-radius: 10px;
            }
            QScrollBar:vertical {
                background: #171717;
                width: 8px;
                margin: 4px 2px 4px 0px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #3A4150;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)
        self.chat_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.chat_scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        self.chat_container = QWidget()
        self.chat_container.setStyleSheet("background: transparent;")
        # Жестко задаем, что контейнер должен подстраиваться под содержимое
        self.chat_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setContentsMargins(8, 8, 8, 8)
        self.chat_layout.setSpacing(10)
        self.chat_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        # ★ ДОБАВЛЯЕМ РАСПОРКУ В КОНЕЦ: она будет съедать всё пустое место снизу, 
        # прижимая сообщения к верху.
        self.chat_layout.addStretch(1)

        self.chat_scroll.setWidget(self.chat_container)
        layout_status.addWidget(self.chat_scroll, 1)

        self.progress_bar = YellowGlowProgressBar()
        layout_status.addWidget(self.progress_bar)

        panel_anim = QFrame()
        panel_anim.setStyleSheet(
            "QFrame { background-color: transparent; border: none; }"
        )
        layout_panel_anim = QHBoxLayout(panel_anim)
        layout_panel_anim.setContentsMargins(0, 8, 0, 8)
        layout_panel_anim.setSpacing(0)
        layout_panel_anim.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.status_indicator = StatusIndicatorWidget()
        layout_panel_anim.addWidget(self.status_indicator)

        layout_status.addWidget(panel_anim)

        self.progress_bar.valueChanged.connect(self._update_percent_label)

        card_school.setFixedWidth(475)
        card_scripts.setFixedWidth(475)
        card_status.setMinimumWidth(470)
        card_status.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        columns_layout.addWidget(card_school, 0)
        columns_layout.addWidget(card_scripts, 0)
        columns_layout.addWidget(card_status, 1)

        main_layout.addLayout(columns_layout)
        
        # --- Измененное начальное сообщение ---
        # --- Отложенное приветствие ---
        self._show_typing_indicator()
        QTimer.singleShot(2000, self._show_welcome_message)

    def _open_github(self, url):
        QDesktopServices.openUrl(QUrl(url))

    def _update_percent_label(self, value):
        pass

    def _update_insert_qr_button_state(self):
        if not hasattr(self, "script_buttons") or not self.script_buttons:
            return
        enabled = bool(self.selected_school_name) and self.chk_2030.isChecked()
        for btn in self.script_buttons:
            if btn.property("mode") == "insert_qr":
                btn.setEnabled(enabled)
                break

    def open_preview_dialog(self):
        folder = self.selected_folder or self.default_schools_dir
        if not folder or not os.path.isdir(folder):
            self.append_log("__ERROR__Папка школы не найдена для предпросмотра.")
            return
        dlg = ImagePreviewDialog(self, folder=folder)
        dlg.exec()

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
            self.append_log(f"__INFO__• {school_num} QR")
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
        elif self.selected_script_mode == "insert_qr":
            self.frame_sub.setVisible(True)
            self.qr_panel.setVisible(False)
            if not self.chk_2030.isChecked():
                self.clear_log()
                self.append_log(
                    "__ERROR__Для вставки QR нужно отметить чекбокс 'Папка 2030'."
                )
                self._reset_script_selection()
                return
            school_folder = self.selected_folder
            if not school_folder or not os.path.isdir(school_folder):
                self.clear_log()
                self.append_log("__ERROR__Папка школы не найдена.")
                self._reset_script_selection()
                return
            preselected = [
                f for f in self.selected_files
                if "2030" in os.path.basename(os.path.dirname(f)).lower()
            ]
            pick_dlg = PickPSDFilesForInsertQRDialog(
                self, school_folder=school_folder, preselected=preselected,
            )
            if pick_dlg.exec() != QDialog.DialogCode.Accepted:
                self._reset_script_selection()
                return
            self.selected_files = pick_dlg.selected_files
            self.lbl_count_info.setText(
                f"PSD файлов выбрано: {len(self.selected_files)}"
            )
            try:
                self.lbl_count_info.start_gradient()
            except Exception:
                pass
            self.append_log(
                f"__INFO__Выбрано для вставки QR: {len(self.selected_files)} PSD"
            )
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

    # ==============================================================
    # ЛОГИКА ЧАТА
    # ==============================================================
    def append_log(self, text):
        clean_text = text
        for prefix in ("__OK__", "__ERROR__", "__INFO__", "__WARN__"):
            if clean_text.startswith(prefix):
                clean_text = clean_text[len(prefix):]
                break

        msg_widget = ChatMessageWidget(
            text=clean_text,
            avatar_pixmap=self.chat_avatar_pixmap,
            show_delivered=True
        )
        
        # ★ ВАЖНО: вставляем виджет ПЕРЕД последней распоркой (addStretch)
        # Иначе он улетит вниз
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, msg_widget)
        
        # Заставляем контейнер пересчитать свой размер
        self.chat_container.adjustSize()
        # Даем Qt немного времени (50 мс) на отрисовку нового сообщения, 
        # чтобы правильно рассчитать высоту, и только потом скроллим
        QTimer.singleShot(50, self._scroll_chat_to_bottom)

    def _scroll_chat_to_bottom(self):
        if hasattr(self, 'chat_scroll'):
            # Прокручиваем в самый низ
            self.chat_scroll.verticalScrollBar().setValue(
                self.chat_scroll.verticalScrollBar().maximum()
            )

    def clear_log(self):
        # Удаляем все виджеты, КРОМЕ последней распорки
        while self.chat_layout.count() > 1:
            item = self.chat_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self._typing_widget = None


    def _show_welcome_message(self):
        # Убираем "печатает..."
        self._hide_typing_indicator()
        # Показываем сообщение
        self.append_log("__INFO__Жду твоей команды!")

    def _show_typing_indicator(self):
        if self._typing_widget is not None:
            return

        self._typing_widget = QWidget()
        self._typing_widget.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(self._typing_widget)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        avatar = QLabel()
        avatar.setFixedSize(30, 30)
        avatar.setStyleSheet("border: none; background: transparent;")
        if self.chat_avatar_pixmap and not self.chat_avatar_pixmap.isNull():
            rounded = QPixmap(30, 30)
            rounded.fill(Qt.GlobalColor.transparent)
            p = QPainter(rounded)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            path = QPainterPath()
            path.addEllipse(0, 0, 30, 30)
            p.setClipPath(path)
            p.drawPixmap(
                0, 0,
                self.chat_avatar_pixmap.scaled(
                    30, 30,
                    Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                    Qt.TransformationMode.SmoothTransformation
                )
            )
            p.end()
            avatar.setPixmap(rounded)
        layout.addWidget(avatar, 0, Qt.AlignmentFlag.AlignTop)

        typing_label = QLabel("печатает...")
        typing_label.setStyleSheet(
            "color: #E0E0E0; font-size: 12px; font-family: 'Segoe UI';"
            "background: transparent; border: none;"
            "padding-top: 6px;"
        )
        layout.addWidget(typing_label, 0, Qt.AlignmentFlag.AlignTop)
        layout.addStretch(1)

        self.chat_layout.addWidget(self._typing_widget)
        QTimer.singleShot(10, self._scroll_chat_to_bottom)

    def _hide_typing_indicator(self):
        if self._typing_widget is not None:
            self.chat_layout.removeWidget(self._typing_widget)
            self._typing_widget.deleteLater()
            self._typing_widget = None

    def on_typing_signal(self, show):
        if show:
            self._show_typing_indicator()
            self._typing_hide_timer.stop()
        else:
            self._typing_hide_timer.start(1100)

    def on_school_selected(self, text):
        if (text and text != "Выберите школу из списка..."
                and text != "Папки не найдены"):
            self.selected_school_name = text
            self.selected_folder = os.path.join(self.default_schools_dir, text)
            self.lbl_path_info.setText("")
            self.lbl_path_info.setToolTip(self.selected_folder)
            self.lbl_path_info.setStyleSheet(
                "color: #FFFFFF; border: none; font-size: 13px; font-weight: bold;"
            )
            self.update_file_list()
            self.combo_border.start_border()
            for btn in self.script_buttons:
                mode = btn.property("mode")
                if mode in ("qr",):
                    btn.setEnabled(True)        # кнопка 4 — всегда доступна
                elif mode not in ("create_folders", "insert_qr"):
                    btn.setEnabled(True)        # кнопки 3, 5, 6
            self._update_insert_qr_button_state()
        else:
            if hasattr(self, "combo_border"):
                self.combo_border.stop_border()
            for btn in self.script_buttons:
                mode = btn.property("mode")
                if mode in ("qr",):
                    btn.setEnabled(True)        # кнопка 4 — всегда доступна
                elif mode not in ("create_folders", "insert_qr"):
                    btn.setEnabled(False)       # кнопки 3, 5, 6 — выключаем
            self.selected_school_name = ""
            self.selected_folder = ""
            self.selected_files = []
            self.lbl_count_info.setText("")
            self.lbl_count_info.stop_gradient()
            self.lbl_count_info.setStyleSheet("border: none;")
            self.lbl_path_info.setText("Школа не выбрана")
            self.lbl_path_info.setStyleSheet(
                "color: #FF4444; border: none; font-size: 14px; font-weight: bold;"
            )
            self._update_insert_qr_button_state()

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

        if self.selected_files:
            self.lbl_count_info.start_gradient()
        else:
            self.lbl_count_info.stop_gradient()
            self.lbl_count_info.update()

    def check_ready_to_run(self):
        if self.selected_script_mode == "insert_qr":
            if not self.selected_school_name:
                self.clear_log()
                self.append_log("__ERROR__Для вставки QR выберите школу.")
                return False
            if not self.chk_2030.isChecked():
                self.clear_log()
                self.append_log(
                    "__ERROR__Для вставки QR отметьте 'Папка 2030'."
                )
                return False
            if not self.selected_files:
                self.clear_log()
                self.append_log(
                    "__ERROR__Не выбраны PSD-файлы для вставки QR."
                )
                return False
            return True

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
                    "галочку 'Папка 1520' или 'Папка 2030'!"
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
            self.append_log("__INFO__— В прошлый раз: "
                            f"{self.last_script_used} для {self.last_school_used}")
        header_text = f"📝 Запущен скрипт: {current_script_name}\n"
        if self.selected_school_name:
            header_text += f"🗂️ Школа: {self.selected_school_name}\n"
        else:
            header_text += "Школа: не выбрана (QR-режим)\n"
        if self.selected_script_mode == "batch_actions":
            header_text += f"⚙️ Action Set: {self.batch_action_set}\n"
            header_text += f"▶️ Actions: {self.batch_actions_text}"
        else:
            header_text = header_text.rstrip("\n")
        self.append_log(header_text)

    def launch_worker(self):
        if not self.check_ready_to_run():
            return
        self.prepare_log_header(self.selected_script_title)

        self.status_indicator.set_generating()
        self.status_indicator.set_text("Обновление")

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
            target_folder=self.selected_folder if self.selected_script_mode == "insert_qr" else None,
        )
        self.worker.progress_signal.connect(self.on_worker_progress)
        self.worker.file_status_signal.connect(self.on_file_status_update)
        self.worker.log_signal.connect(self.append_log)
        self.worker.typing_signal.connect(self.on_typing_signal)
        self.worker.finished_signal.connect(self.on_worker_finished)
        self.worker.start()

    def on_file_status_update(self, action_text, filename):
        if filename:
            self.status_indicator.set_text(f"{action_text}")
        else:
            self.status_indicator.set_text(action_text)

    def on_worker_progress(self, current, total):
        val = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(val)

    def on_worker_finished(self, processed_count, script_title):
        self.status_indicator.set_text("Завершено")
        self.progress_bar.stop_anim()
        self.progress_bar.setValue(100)
        # Строка с self.append_log("— Готово —") удалена
        self.last_script_used = script_title
        self.last_school_used = self.selected_school_name
        QTimer.singleShot(3000, self._restart_icon_after_finish)

    def _restart_icon_after_finish(self):
        if self.worker is None or not self.worker.isRunning():
            self.status_indicator.set_waiting()

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
        self.status_indicator.set_text("Отменено")
        self.progress_bar.stop_anim()
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

    icon_path = resource_path("avatarka2.jpg")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = ModernPhotoshopWorkstation()
    window.show()

    app.aboutToQuit.connect(window._kill_photoshop)
    sys.exit(app.exec())