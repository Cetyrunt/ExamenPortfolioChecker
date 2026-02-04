from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

class StatusCard(QFrame):
    """A colored status card representing an ID and its REA check status."""

    def __init__(self, id_label, status, files, excel_data, parent=None):
        super().__init__(parent)
        self.id_label = id_label
        self.status = status
        self.files = files
        self.excel_data = excel_data
        self.setup_ui()

    def setup_ui(self):
        self.setFixedSize(120, 100)
        self.setCursor(Qt.PointingHandCursor)
        layout = QVBoxLayout(self)

        colors = {
            'GOOD': '#d4edda',
            'WRONG': '#f8d7da',
            'EMPTY': '#fff3cd',
            'NODATA': '#e2e3e5',
            'INCOMPLETE': '#ffe0b3'
        }
        border_colors = {
            'GOOD': '#28a745',
            'WRONG': '#dc3545',
            'EMPTY': '#ffc107',
            'NODATA': '#6c757d',
            'INCOMPLETE': '#ff8800'
        }

        self.setStyleSheet(f"""
            QFrame {{
                background-color: {colors.get(self.status, '#ffffff')};
                border: 2px solid {border_colors.get(self.status, '#000000')};
                border-radius: 8px;
            }}
            QFrame:hover {{
                background-color: white;
                border: 3px solid {border_colors.get(self.status, '#000000')};
            }}
        """)

        id_text = QLabel(self.id_label)
        id_text.setFont(QFont("Arial", 11, QFont.Bold))
        id_text.setAlignment(Qt.AlignCenter)
        id_text.setStyleSheet("border: none; background: transparent;")

        status_map = {
            'GOOD': "✅",
            'WRONG': "❌",
            'EMPTY': "⚠️",
            'NODATA': "❓",
            'INCOMPLETE': "🟠"
        }
        icon_label = QLabel(status_map.get(self.status, "❓"))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("border: none; background: transparent; font-size: 18px;")

        layout.addWidget(id_text)
        layout.addWidget(icon_label)

    def mousePressEvent(self, event):
        """Open details panel on click."""
        self.window().display_details(self)
        super().mousePressEvent(event)
