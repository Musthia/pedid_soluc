from pathlib import Path

import openpyxl

from PySide6.QtCore import Qt

BASE_DIR = Path(__file__).resolve().parent.parent
DOCS_DIR = BASE_DIR / "docs"
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
    QToolBar,
    QTreeView,
    QVBoxLayout,
    QWidget,
    QGridLayout,
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
        start_dir = str(DOCS_DIR) if DOCS_DIR.exists() else str(Path.home())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            start_dir,
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
        self._build_toolbar()
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

    def _build_toolbar(self):
        toolbar = QToolBar("Herramientas")
        self.addToolBar(toolbar)

        compare_btn = QPushButton("Comparar CUIL/CUIT")
        compare_btn.setObjectName("btnCompare")
        compare_btn.clicked.connect(self.compare_tables)
        toolbar.addWidget(compare_btn)

        export_btn = QPushButton("Exportar Excel")
        export_btn.setObjectName("btnExport")
        export_btn.clicked.connect(self.export_results)
        toolbar.addWidget(export_btn)

    def _build_central_widget(self):
        central: QWidget = QWidget()
        self.setCentralWidget(central)

        layout = QGridLayout(central)
        layout.setSpacing(10)

        self.tree_widget_1 = ExcelTreeWidget("Pedidos")
        self.tree_widget_2 = ExcelTreeWidget("Listados")

        # Widget de resultados coincidencias
        self.results_group = QGroupBox("Resultados coincidencias")
        results_layout = QVBoxLayout(self.results_group)
        self.results_view = QTreeView()
        self.results_view.setAlternatingRowColors(True)
        self.results_view.setHeaderHidden(False)
        self.results_view.setItemsExpandable(False)
        self.results_model = QStandardItemModel(self)
        self.results_view.setModel(self.results_model)
        results_layout.addWidget(self.results_view)

        # Widget de pedidos sin coincidencia
        self.unmatched_group = QGroupBox("Pedidos sin coincidencia")
        unmatched_layout = QVBoxLayout(self.unmatched_group)
        self.btn_export_unmatched = QPushButton("Exportar sin coincidencias")
        self.btn_export_unmatched.clicked.connect(self.export_unmatched)
        unmatched_layout.addWidget(self.btn_export_unmatched)
        self.unmatched_view = QTreeView()
        self.unmatched_view.setAlternatingRowColors(True)
        self.unmatched_view.setHeaderHidden(False)
        self.unmatched_view.setItemsExpandable(False)
        self.unmatched_model = QStandardItemModel(self)
        self.unmatched_view.setModel(self.unmatched_model)
        unmatched_layout.addWidget(self.unmatched_view)

        layout.addWidget(self.tree_widget_1, 0, 0)
        layout.addWidget(self.tree_widget_2, 0, 1)
        layout.addWidget(self.results_group, 1, 0, 1, 2)
        layout.addWidget(self.unmatched_group, 2, 0, 1, 2)

        # Configurar proporciones
        layout.setRowStretch(0, 2)
        layout.setRowStretch(1, 3)
        layout.setRowStretch(2, 2)
        layout.setColumnStretch(0, 1)
        layout.setColumnStretch(1, 1)

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

    def compare_tables(self):
        model1 = self.tree_widget_1.model
        model2 = self.tree_widget_2.model
        if model1 is None or model2 is None:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Comparar", "Carga ambas tablas primero.")
            return

        # Buscar columna cuil_cuit
        headers1 = [model1.horizontalHeaderItem(i).text() for i in range(model1.columnCount())]
        headers2 = [model2.horizontalHeaderItem(i).text() for i in range(model2.columnCount())]

        def find_col(headers):
            for i, h in enumerate(headers):
                if h.strip().lower() == "cuil_cuit":
                    return i
            return None

        col1 = find_col(headers1)
        col2 = find_col(headers2)

        if col1 is None or col2 is None:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Comparar", "No se encontró columna 'cuil_cuit' en ambas tablas.")
            return

        # Construir índice de tabla 2
        index2 = {}
        for row in range(model2.rowCount()):
            val = model2.item(row, col2).text().strip()
            if val:
                if val not in index2:
                    index2[val] = row

        # Preparar modelos
        self.results_model.clear()
        self.results_model.setHorizontalHeaderLabels(headers1)

        self.unmatched_model.clear()
        self.unmatched_model.setHorizontalHeaderLabels(headers1)

        count_pairs = 0
        count_unmatched = 0
        for row1 in range(model1.rowCount()):
            val_item = model1.item(row1, col1)
            val = val_item.text().strip() if val_item else ""
            if not val:
                continue
            if val in index2:
                row2_idx = index2[val]
                # Añadir fila de tabla 1
                items1 = []
                for col in range(model1.columnCount()):
                    text = model1.item(row1, col).text() if model1.item(row1, col) else ""
                    items1.append(QStandardItem(text))
                self.results_model.appendRow(items1)

                # Añadir fila de tabla 2 directamente debajo
                items2 = []
                for col in range(model2.columnCount()):
                    text = model2.item(row2_idx, col).text() if model2.item(row2_idx, col) else ""
                    items2.append(QStandardItem(text))
                self.results_model.appendRow(items2)
                count_pairs += 1
            else:
                # Sin coincidencia
                items_u = []
                for col in range(model1.columnCount()):
                    text = model1.item(row1, col).text() if model1.item(row1, col) else ""
                    items_u.append(QStandardItem(text))
                self.unmatched_model.appendRow(items_u)
                count_unmatched += 1

        # Ajustar columnas
        for col in range(self.results_model.columnCount()):
            self.results_view.resizeColumnToContents(col)
        for col in range(self.unmatched_model.columnCount()):
            self.unmatched_view.resizeColumnToContents(col)

        self.statusBar().showMessage(f"Comparación completada. Pares encontrados: {count_pairs}, sin coincidencia: {count_unmatched}")

    def export_results(self):
        if self.results_model.rowCount() == 0:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(self, "Exportar", "No hay resultados para exportar.")
            return

        import openpyxl

        base_dir = DOCS_DIR if DOCS_DIR.exists() else Path.home()
        default_file = base_dir / "resultados_comparacion.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar resultados Excel",
            str(default_file),
            "Archivos Excel (*.xlsx)",
        )
        if not file_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resultados"

        # Headers
        headers = []
        for i in range(self.results_model.columnCount()):
            hdr = self.results_model.horizontalHeaderItem(i)
            headers.append(hdr.text() if hdr else "")
        ws.append(headers)

        # Data
        for row in range(self.results_model.rowCount()):
            row_data = []
            for col in range(self.results_model.columnCount()):
                item = self.results_model.item(row, col)
                row_data.append(item.text() if item else "")
            ws.append(row_data)

        wb.save(file_path)
        self.statusBar().showMessage(f"Exportado a {file_path}")

    def export_unmatched(self):
        if self.unmatched_model.rowCount() == 0:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(self, "Exportar", "No hay pedidos sin coincidencia para exportar.")
            return

        import openpyxl

        base_dir = DOCS_DIR if DOCS_DIR.exists() else Path.home()
        default_file = base_dir / "pedidos_sin_coincidencia.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar pedidos sin coincidencia",
            str(default_file),
            "Archivos Excel (*.xlsx)",
        )
        if not file_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sin coincidencia"

        headers = []
        for i in range(self.unmatched_model.columnCount()):
            hdr = self.unmatched_model.horizontalHeaderItem(i)
            headers.append(hdr.text() if hdr else "")
        ws.append(headers)

        for row in range(self.unmatched_model.rowCount()):
            row_data = []
            for col in range(self.unmatched_model.columnCount()):
                item = self.unmatched_model.item(row, col)
                row_data.append(item.text() if item else "")
            ws.append(row_data)

        wb.save(file_path)
        self.statusBar().showMessage(f"Exportado a {file_path}")

    def add_custom_widget(self, widget, stretch=1):
        central = self.centralWidget()
        if central is None:
            central = QWidget()
            self.setCentralWidget(central)
        layout = central.layout()
        if layout is None:
            layout = QHBoxLayout(central)
        layout.addWidget(widget, stretch=stretch)