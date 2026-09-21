from pathlib import Path

import openpyxl

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QMainWindow,
    QMenuBar,
    QMenu,
    QPushButton,
    QStatusBar,
    QTreeView,
    QVBoxLayout,
    QWidget,
)


class ExcelTreeWidget(QGroupBox):
    def __init__(self, title="Cargar Excel", parent=None):
        super().__init__(title, parent)
        self.model: QStandardItemModel | None = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        self.load_button = QPushButton("Cargar archivo Excel")
        self.load_button.setObjectName("btnLoadExcel")
        self.load_button.clicked.connect(self._load_excel)
        layout.addWidget(self.load_button)

        self.tree_view = QTreeView()
        self.tree_view.setAlternatingRowColors(True)
        self.tree_view.setAnimated(False)
        self.tree_view.setIndentation(10)
        self.tree_view.setHeaderHidden(False)
        self.tree_view.setItemsExpandable(False)
        self.tree_view.setExpandsOnDoubleClick(False)
        layout.addWidget(self.tree_view)

        self.model = QStandardItemModel(self)
        self.model.setHorizontalHeaderLabels([])
        self.tree_view.setModel(self.model)

    def _load_excel(self):
        app = QApplication.instance()
        if app is None:
            app = QApplication([])
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            str(Path.home()),
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)",
        )
        if not file_path:
            return
        self._populate_from_excel(file_path)

    def _populate_from_excel(self, file_path: str):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            return

        headers = [str(cell) if cell is not None else "" for cell in rows[0]]
        data_rows = rows[1:]

        self.model = QStandardItemModel(self)
        self.model.setHorizontalHeaderLabels(headers)

        for row_data in data_rows:
            row_items: list[QStandardItem] = []
            for col_idx, value in enumerate(row_data):
                if col_idx >= len(headers):
                    break
                text = str(value) if value is not None else ""
                item = QStandardItem(text)
                item.setEditable(False)
                row_items.append(item)
            if row_items:
                self.model.appendRow(row_items)

        self.tree_view.setModel(self.model)

        for col in range(len(headers)):
            self.tree_view.resizeColumnToContents(col)


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Aplicación")
        self.setMinimumSize(1000, 700)

        self._build_menu_bar()
        self._build_central_widget()
        self._build_status_bar()

    def _build_menu_bar(self):
        menu_bar: QMenuBar = self.menuBar()

        archivo_menu: QMenu = menu_bar.addMenu("Archivo")
        exit_action: QAction = QAction("Salir", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        archivo_menu.addAction(exit_action)

        ayuda_menu: QMenu = menu_bar.addMenu("Ayuda")
        about_action: QAction = QAction("Acerca de", self)
        about_action.triggered.connect(self._show_about)
        ayuda_menu.addAction(about_action)

    def _build_central_widget(self):
        central: QWidget = QWidget()
        self.setCentralWidget(central)

        layout: QHBoxLayout = QHBoxLayout(central)
        layout.setSpacing(10)

        self.tree_widget_1 = ExcelTreeWidget("Tabla 1")
        self.tree_widget_2 = ExcelTreeWidget("Tabla 2")

        layout.addWidget(self.tree_widget_1, stretch=1)
        layout.addWidget(self.tree_widget_2, stretch=1)

    def _build_status_bar(self):
        status_bar: QStatusBar = QStatusBar()
        status_bar.setObjectName("statusBar")
        self.setStatusBar(status_bar)
        status_bar.showMessage("Listo")

    def _show_about(self):
        from PySide6.QtWidgets import QMessageBox

        QMessageBox.about(
            self,
            "Acerca de",
            "Aplicación plantilla\nVersión 1.0.0",
        )

    def add_custom_widget(self, widget, stretch=1):
        central = self.centralWidget()
        if central is None:
            central = QWidget()
            self.setCentralWidget(central)
        layout = central.layout()
        if layout is None:
            layout = QHBoxLayout(central)
        layout.addWidget(widget, stretch=stretch)