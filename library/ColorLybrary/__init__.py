from PyQt6 import QtCore, QtWidgets
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QSlider, QLineEdit
from PyQt6.QtGui import QCursor, QPixmap, QColor, QPainter, QPen, QBrush
from PyQt6.QtCore import QTimer, Qt, pyqtSignal, QPoint

def hsv_to_rgb(h: int, s: int, v: int):
    s = s / 100
    v = v / 100

    hi = int(h / 60) % 6
    f = h / 60 - hi
    p = v * (1 - s)
    q = v * (1 - f * s)
    t = v * (1 - (1 - f) * s)

    if hi == 0:
        r, g, b = v, t, p
    elif hi == 1:
        r, g, b = q, v, p
    elif hi == 2:
        r, g, b = p, v, t
    elif hi == 3:
        r, g, b = p, q, v
    elif hi == 4:
        r, g, b = t, p, v
    elif hi == 5:
        r, g, b = v, p, q

    return int(r * 255), int(g * 255), int(b * 255)

def argb_to_hsva(argb: str):

    a = int(argb[0:2], 16)
    r = int(argb[2:4], 16)
    g = int(argb[4:6], 16)
    b = int(argb[6:8], 16)

    # Convert RGB to the range [0, 1]
    r, g, b = r / 255.0, g / 255.0, b / 255.0

    max_c = max(r, g, b)
    min_c = min(r, g, b)
    delta = max_c - min_c

    # Value
    v = max_c

    # Saturation
    s = 0 if max_c == 0 else delta / max_c

    # Hue
    if delta == 0:
        h = 0
    elif max_c == r:
        h = (g - b) / delta % 6
    elif max_c == g:
        h = (b - r) / delta + 2
    elif max_c == b:
        h = (r - g) / delta + 4
    h = h * 60  # Convert to degrees

    # Ensure hue is between 0 and 360
    if h < 0:
        h += 360

    # Convert saturation and value to the range [0, 100]
    s *= 100
    v *= 100

    return int(a), int(h), int(s), int(v)

class SelectedIcon(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(36, 5)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        width, height = self.width() - 1, self.height()
        color = QColor(255, 255, 255)
        painter.setPen(QPen(color, 1, Qt.PenStyle.SolidLine))

        z = 2
        for i in range(height):
            painter.drawLine(0, i, z, i)
            painter.drawLine(width - z, i, width, i)
            if i <= (height // 2 - 1):
                z += 1
            else:
                z -= 1

class CaroLabel(QLabel):

    def __init__(self, parent=None, width=12, height=12, pointSize=3):
        super().__init__(parent)
        self.setFixedSize(width, height)
        self.pointSize = pointSize

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        wlen, hlen = self.width() // self.pointSize + 1, self.height() // self.pointSize + 1

        for row in range(wlen):
            for col in range(hlen):
                color = QColor('white') if (row + col) % 2 == 1 else QColor('gray')
                painter.setPen(QPen(color, 1, Qt.PenStyle.SolidLine))
                painter.setBrush(QBrush(color, Qt.BrushStyle.SolidPattern))
                painter.drawRect(row * self.pointSize - 1, col * self.pointSize - 1, self.pointSize, self.pointSize)


class RePainterLabel(QLabel):
    positionChanged = pyqtSignal(int)

    def __init__(self, parent=None, start_value=0, _min=0, _max=100, width=100, height=100, drawing=None):
        super().__init__(parent)
        self.setFixedSize(width, height)
        self.min = _min
        self.max = _max
        self.drawing = drawing

        self.selected_value = SelectedIcon(parent=self)
        self.setValue((self.height() - 5) - int(start_value / _max * (self.height() - 5)))

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.readEvent)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        width, height = self.width() - 5, self.height() - 2
        for i in range(2, height):
            color = self.drawing(int(i / height * self.max), self.max)
            painter.setPen(QPen(color, 1, Qt.PenStyle.SolidLine))
            painter.drawLine(4, i, width, i)

    def setValue(self, h):
        self.selected_value.move(0, h)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.timer.start(30)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.timer.stop()

    def readEvent(self):
        pos = self.mapFromGlobal(QCursor.pos())
        y = max(min(pos.y(), (self.height() - 5)), 0)
        self.setValue(y)
        self.positionChanged.emit(self.max - int(y / (self.height() - 5) * self.max))


class ColorSquareLabel(QLabel):
    positionChanged = pyqtSignal(int, int)

    def __init__(self, parent=None, start_color="FFFF0000"):
        super().__init__(parent)
        self.alpha, self.hue, self.s, self.v = argb_to_hsva(start_color)
        self.bSize = 300
        self.cSize = 100
        self.setFixedSize(self.bSize, self.bSize)
        self.setMouseTracking(True)

        self.circle = QLabel(self)
        self.circle.setGeometry(QtCore.QRect(50, 50, 10, 10))
        self.circle.setStyleSheet("border-radius: 5px; border: 2px solid #fff;")

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.readEvent)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.NonCosmeticBrushPatterns)
        poisize = self.bSize // self.cSize
        height, width = self.bSize // poisize, self.bSize // poisize
        for y in range(height):
            for x in range(width):
                s = (x / width) * 100
                v = ((height - y) / height) * 100
                r, g, b = hsv_to_rgb(self.hue, int(s), int(v))
                color = QColor(r, g, b)
                painter.setBrush(QBrush(color, Qt.BrushStyle.SolidPattern))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawRect(x * poisize, y * poisize, poisize, poisize)

        # Draw the circle cursor
        if hasattr(self, 'cursor_pos'):
            self.draw_circle_cursor(self.cursor_pos.x(), self.cursor_pos.y())
        else:
            self.draw_circle_cursor(int(self.s * poisize), self.bSize - int(self.v * poisize))

    def draw_circle_cursor(self, x, y):
        x = x - 5
        y = y - 5
        cursor_color = 'black' if (x < 90 and y < 180) else 'white'
        self.circle.setStyleSheet(f"border-radius: 5px; border: 2px solid {cursor_color};")
        self.circle.move(x, y)

    def setHue(self, hue, x, y):
        if self.hue != hue:
            self.hue = hue

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.timer.start(30)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.timer.stop()

    def readEvent(self):
        pos = self.mapFromGlobal(QCursor.pos())
        self.cursor_pos = pos
        self.positionChanged.emit(pos.x(), pos.y())
        self.set_cursor_shape()

    def enterEvent(self, event):
        self.set_cursor_shape()

    def leaveEvent(self, event):
        self.unsetCursor()

    def set_cursor_shape(self):
        cursor_pixmap = QPixmap(20, 20)
        cursor_pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(cursor_pixmap)
        s, v = 0, 0
        if hasattr(self, 'cursor_pos'):
            s = int(self.cursor_pos.x() / 300 * 100)
            v = int((290 - self.cursor_pos.y()) / 300 * 100)
        cursor_color = QColor('black') if (s < 30 and v > 35) else QColor('white')
        painter.setPen(QPen(cursor_color, 1))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawEllipse(5, 5, 10, 10)
        painter.end()
        self.setCursor(QCursor(cursor_pixmap))


class PickerColorWindow(QMainWindow):
    colorChanged = pyqtSignal(str)

    def __init__(self, start_color: str):
        super().__init__()
        self.setWindowTitle("Color Picker")
        self.setWindowFlags(
            Qt.WindowType.Window | Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowStaysOnTopHint
            )
        self.resize(600, 400)

        self.a, self.h, self.s, self.v = argb_to_hsva(start_color[1:])
        text_color = "black" if (self.s < 40 and self.v > 40) or self.a < 150 else "white"

        self.label = ColorSquareLabel(parent=self, start_color=start_color[1:])
        self.label.setGeometry(QtCore.QRect(10, 65, 300, 300))
        self.label.positionChanged.connect(self.update_viewLabel)

        self.color_label = RePainterLabel(
            self, self.h, 0, 359, 35, 355, self.drawHueSlider
        )
        self.color_label.setGeometry(QtCore.QRect(325, 10, 35, 355))
        self.color_label.positionChanged.connect(self.update_hue)

        self.alpha_chanel_back = CaroLabel(self, 28, 350, 4)
        self.alpha_chanel_back.setGeometry(QtCore.QRect(380, 12, 28, 350))

        self.alpha_chanel = RePainterLabel(
            self, self.a, 0, 255, 36, 355, self.drawAlphaChanel
        )
        self.alpha_chanel.setGeometry(QtCore.QRect(376, 10, 36, 355))
        self.alpha_chanel.positionChanged.connect(self.update_alpha)

        self.view_chanel_back = CaroLabel(self, 120, 100, 4)
        self.view_chanel_back.setGeometry(QtCore.QRect(430, 10, 120, 100))

        self.currentColorDisplay = QLabel(parent=self)
        self.currentColorDisplay.setGeometry(QtCore.QRect(430, 10, 120, 50))
        self.currentColorDisplay.setStyleSheet(
            f"background-color: {start_color}; color: {text_color}; font-weight: bold;"
        )
        self.currentColorDisplay.setText("Current Color")
        self.currentColorDisplay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.oldColorDisplay = QLabel(parent=self)
        self.oldColorDisplay.setGeometry(QtCore.QRect(430, 60, 120, 50))
        self.oldColorDisplay.setStyleSheet(f"background-color: {start_color}; color: {text_color}; font-weight: bold;")
        self.oldColorDisplay.setText("Old Color")
        self.oldColorDisplay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.r_input = QLineEdit(parent=self)
        self.r_input.setGeometry(QtCore.QRect(430, 180, 50, 20))

    def drawHueSlider(self, value, max):
        r, g, b = hsv_to_rgb((360 - value) % 360, 100, 100)
        return QColor(r, g, b)

    def drawAlphaChanel(self, value, max):
        r, g, b = hsv_to_rgb(self.h, self.s, self.v)
        return QColor(r, g, b, int(max - value))

    def update_hue(self, h):
        x, y = self.get_current_position()
        self.h = h
        self.label.setHue(h, x, y)
        self.update_color_display(x, y)

    def update_alpha(self, a):
        x, y = self.get_current_position()
        self.a = a
        self.update_color_display(x, y)

    def update_viewLabel(self, x, y):
        move_x, move_y = 10 + min(max(x, 0), 300), 60 + min(max(y, 0), 300)
        self.label.cursor_pos = QPoint(move_x - 10, move_y - 60)
        self.update_color_display(x, y)

    def update_color_display(self, x, y):
        x = max(min(x, 300), 0)
        y = max(min(y, 300), 0)

        s = int(x / 300 * 100)
        v = int((300 - y) / 300 * 100)
        r, g, b = hsv_to_rgb(self.label.hue, s, v)
        self.s = s
        self.v = v
        text_color = "black" if (s < 40 and v > 40) or self.a < 150 else "white"
        color = f"#{self.a:02X}{r:02X}{g:02X}{b:02X}" if self.a != 255 else f"#{r:02X}{g:02X}{b:02X}"
        self.currentColorDisplay.setStyleSheet(f"background-color: {color}; color: {text_color}; font-weight: bold;")
        self.colorChanged.emit(color)
        self.label.update()
        self.alpha_chanel.update()

    def get_current_position(self):
        if hasattr(self, 'cursor_pos'):
            return self.label.cursor_pos.x(), self.label.cursor_pos.y()
        else:
            return int(self.s / 100 * 300), 100 - int(self.v / 100 * 300)
