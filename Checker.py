import sys
import os
import glob
import pandas as pd
from urllib.parse import quote
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QGridLayout, QLabel, QFrame,
    QScrollArea, QPushButton, QFileDialog
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont


class StatusCard(QFrame):
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
        self.window().display_details(self)
        super().mousePressEvent(event)


class FileInspector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Portfolio Validator (Positional Scan)")
        self.resize(1200, 850)

        self.evidence_folder = None
        self.excel_path = None
        self.loading = False

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_vbox = QVBoxLayout(central_widget)

        self.top_bar = QHBoxLayout()
        self.btn_select_project = QPushButton("📁 Select Project Folder")
        self.btn_select_project.setStyleSheet("padding: 10px; font-weight: bold;")
        self.btn_select_project.clicked.connect(self.select_project_folder)

        self.btn_refresh = QPushButton("🔄 Refresh")
        self.btn_refresh.setEnabled(False)
        self.btn_refresh.clicked.connect(self.safe_refresh)

        self.status_label = QLabel("Ready.")
        self.top_bar.addWidget(self.btn_select_project)
        self.top_bar.addWidget(self.btn_refresh)
        self.top_bar.addWidget(self.status_label)
        self.top_bar.addStretch()
        self.main_vbox.addLayout(self.top_bar)

        self.content_layout = QHBoxLayout()
        self.main_vbox.addLayout(self.content_layout)

        left_container = QWidget()
        self.grid_layout = QGridLayout(left_container)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(left_container)
        self.content_layout.addWidget(scroll, 2)

        self.detail_panel = QFrame()
        self.detail_panel.setMinimumWidth(420)
        self.detail_panel.setStyleSheet("background-color: #ffffff; border-left: 1px solid #ddd;")
        self.detail_layout = QVBoxLayout(self.detail_panel)

        self.detail_title = QLabel("Info")
        self.detail_title.setFont(QFont("Arial", 14, QFont.Bold))

        self.detail_info = QLabel("Details will show here.")
        self.detail_info.setWordWrap(True)
        self.detail_info.setTextFormat(Qt.RichText)
        self.detail_info.setOpenExternalLinks(True)

        self.detail_layout.addWidget(self.detail_title)
        self.detail_layout.addWidget(self.detail_info)
        self.detail_layout.addStretch()
        self.content_layout.addWidget(self.detail_panel, 1)

    def clean_val(self, val):
        if pd.isna(val):
            return "---"

        # Real pandas datetime
        if isinstance(val, pd.Timestamp):
            if val.year == 2000:
                return "---"
            return val.strftime("%d-%m-%Y")

        s = str(val).strip()

        # Excel default empty date (string form)
        if s.startswith("2000-01-01") or s.startswith("1899-12-30"):
            return "---"

        # ISO date string -> EU date
        if len(s) >= 10 and s[4] == "-" and s[7] == "-":
            try:
                y, m, d = s[:10].split("-")
                return f"{d}-{m}-{y}"
            except Exception:
                pass

        if "dropdown" in s.lower() or s == "0":
            return "---"

        return s



    def select_project_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Project Folder")
        if not folder:
            return

        excel_files = [
            f for f in glob.glob(os.path.join(folder, "*.xlsx"))
            if not os.path.basename(f).startswith("~$")
        ]
        self.excel_path = excel_files[0] if excel_files else None

        evidence_path = os.path.join(folder, "Bewijslasten")
        self.evidence_folder = evidence_path if os.path.isdir(evidence_path) else None

        if self.excel_path and self.evidence_folder:
            self.btn_refresh.setEnabled(True)
            self.safe_refresh()
        else:
            self.detail_info.setText("Missing Excel file or 'Bewijslasten' folder.")

    def safe_refresh(self):
        if self.loading:
            return
        self.loading = True
        self.btn_refresh.setEnabled(False)
        QTimer.singleShot(0, self.load_data)

    def load_data(self):
        try:
            if not self.excel_path or not os.path.isfile(self.excel_path):
                self.detail_info.setText("Excel file missing!")
                return

            if not self.evidence_folder or not os.path.isdir(self.evidence_folder):
                self.detail_info.setText("Evidence folder missing!")
                return

            for i in reversed(range(self.grid_layout.count())):
                widget = self.grid_layout.itemAt(i).widget()
                if widget:
                    widget.deleteLater()

            xl = pd.ExcelFile(self.excel_path, engine="openpyxl")
            sheet = xl.sheet_names[1] if len(xl.sheet_names) > 1 else xl.sheet_names[0]
            df = pd.read_excel(xl, sheet_name=sheet, skiprows=43, header=None)

            rows, cols = 0, 0
            for _, row in df.iterrows():
                try:
                    eid = str(row[0]).strip().upper()
                    if not eid.isalpha() or len(eid) != 1:
                        continue

                    excel_info = {
                        "Criteria": self.clean_val(
                            str(row[1]).split(".", 1)[1] if "." in str(row[1]) else "---"
                        ),
                        "Type": self.clean_val(row[2]),
                        "Context": self.clean_val(row[3]),
                        "Omschr": self.clean_val(row[4]),
                        "Niveau": self.clean_val(row[5]),
                        "Rel": self.clean_val(row[6]),
                        "Auth": self.clean_val(row[7]),
                        "Act": self.clean_val(row[8]),
                        "Datum": self.clean_val(row[9]),
                    }

                    folder_path = os.path.join(self.evidence_folder, eid)
                    found_files = (
                        [f for f in os.listdir(folder_path) if not f.startswith(".")]
                        if os.path.isdir(folder_path)
                        else []
                    )

                    has_excel_entry = excel_info["Type"] != "---" or excel_info["Context"] != "---"
                    has_files = bool(found_files)

                    if has_excel_entry and has_files:
                        if any(
                            excel_info[k] == "---" or excel_info[k].upper() == "NEE"
                            for k in ["Niveau", "Rel", "Auth", "Act", "Datum"]
                        ):
                            status = "INCOMPLETE"
                        else:
                            status = "GOOD"
                    elif has_excel_entry and not has_files:
                        status = "EMPTY"
                    elif not has_excel_entry and has_files:
                        status = "WRONG"
                    else:
                        status = "NODATA"

                    card = StatusCard(eid, status, found_files, excel_info)
                    self.grid_layout.addWidget(card, rows, cols)

                    cols += 1
                    if cols > 4:
                        cols = 0
                        rows += 1
                except Exception as e:
                    print(f"Row error ({eid if 'eid' in locals() else '?'}) → {e}")
                    continue

        finally:
            QTimer.singleShot(3000, self.enable_refresh)

    def enable_refresh(self):
        self.loading = False
        self.btn_refresh.setEnabled(True)

    def display_details(self, card):
        self.detail_title.setText(f"ID {card.id_label} Inspection")
        d = card.excel_data

        if card.files:
            files_html = "".join(
                f"<li><a href='file:///{quote(os.path.abspath(os.path.join(self.evidence_folder, card.id_label, f)))}'>{f}</a></li>"
                for f in card.files
            )
        else:
            files_html = "<li><i>No files found</i></li>"

        def color_val(val):
            v = val.upper()
            if v == "JA":
                return f"<span style='color:green;font-weight:bold'>{val}</span>"
            if v == "NEE":
                return f"<span style='color:red;font-weight:bold'>{val}</span>"
            return val
   
        todos = []

        def is_empty(v):
            return v == "---"

        def is_no(v):
            return v.upper() == "NEE"

        if not card.files:
            todos.append("❌ Geen bestanden aanwezig")

        if is_empty(d["Niveau"]):
            todos.append("⚠️ Niveau niet ingevuld")

        if is_empty(d["Rel"]):
            todos.append("⚠️ Relevant niet ingevuld")
        elif is_no(d["Rel"]):
            todos.append("❌ Bewijs is als niet relevant gemarkeerd")

        if is_empty(d["Auth"]):
            todos.append("⚠️ Authentiek niet ingevuld")
        elif is_no(d["Auth"]):
            todos.append("❌ Bewijs is als niet authentiek gemarkeerd")

        if is_empty(d["Act"]):
            todos.append("⚠️ Actueel niet ingevuld")
        elif is_no(d["Act"]):
            todos.append("❌ Bewijs is als niet actueel gemarkeerd")

        if is_empty(d["Datum"]):
            todos.append("⚠️ Datum niet ingevuld")

        todo_html = (
            "<ul>" + "".join(f"<li>{t}</li>" for t in todos) + "</ul>"
            if todos else "<i>Bewijs is compleet 🎉</i>"
        )

        self.detail_info.setText(
            f"<div style='font-size:13px; line-height:1.4;'>"
            f"<b>📂 Files:</b><ul>{files_html}</ul><hr>"
            f"<b>📊 REA Check:</b><br>"
            f"Niveau: {color_val(d['Niveau'])}<br>"
            f"Relevant: {color_val(d['Rel'])}<br>"
            f"Authentiek: {color_val(d['Auth'])}<br>"
            f"Actueel: {color_val(d['Act'])}<br><hr>"
            f"<b>📊 Excel Info:</b><br>"
            f"Criteria: {d['Criteria']}<br>"
            f"Type: {d['Type']}<br>"
            f"Context: {d['Context']}<br>"
            f"Omschrijving: {d['Omschr']}<br>"
            f"Datum: {d['Datum']}"
            f"</div><hr>"
            f"<b>📝 TODO:</b>{todo_html}"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileInspector()
    window.show()
    sys.exit(app.exec())
