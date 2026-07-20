import sys
import os
from datetime import datetime
from openpyxl import load_workbook

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QLineEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QGridLayout, QAbstractItemView, 
    QDialog, QSpinBox, QDoubleSpinBox, QMessageBox, QInputDialog, QFileDialog, QTabWidget
)
from PySide6.QtCore import Qt

from core.models import Product
from core.models import ClientDetails
from core.catalog_manager import CatalogManager
from core.quote_engine import QuoteEngine
from core.pdf_generator import PDFGenerator
from core.excel_generator import ExcelGenerator

# Ensure directories exist
os.makedirs("data", exist_ok=True)
os.makedirs("drafts", exist_ok=True)
os.makedirs("output", exist_ok=True)

LIGHT_THEME = {
    "bg": "#F8FAFC", "surface": "#FFFFFF", "primary": "#2563EB",
    "text": "#0F172A", "text_muted": "#64748B", "border": "#E2E8F0"
}
DARK_THEME = {
    "bg": "#0F172A", "surface": "#1E293B", "primary": "#3B82F6",
    "text": "#F8FAFC", "text_muted": "#94A3B8", "border": "#334155",
    "danger": "#EF4444"
}

def get_stylesheet(is_dark: bool) -> str:
    theme = DARK_THEME if is_dark else LIGHT_THEME
    return f"""
    QMainWindow, QWidget {{ background-color: {theme['bg']}; color: {theme['text']}; font-family: 'Inter', sans-serif; }}
    QFrame#surface {{ background-color: {theme['surface']}; border-radius: 8px; border: 1px solid {theme['border']}; }}
    QPushButton {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 4px; padding: 8px 16px; color: {theme['text']}; }}
    QPushButton:hover {{ background-color: {theme['border']}; }}
    QPushButton#primary_btn {{ background-color: {theme['primary']}; color: #FFFFFF; border: none; font-weight: bold; }}
    QPushButton#danger_btn {{ background-color: {theme.get('danger', '#EF4444')}; color: #FFFFFF; border: none; font-weight: bold; padding: 4px 8px; }}
    QPushButton#sidebar_btn {{ text-align: left; padding-left: 20px; border: none; background: transparent; font-size: 14px; border-radius: 0px; }}
    QPushButton#sidebar_btn:checked {{ background-color: {theme['border']}; border-left: 4px solid {theme['primary']}; font-weight: bold; }}
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 4px; padding: 6px; color: {theme['text']}; }}
    QTableWidget {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 8px; gridline-color: {theme['border']}; }}
    QHeaderView::section {{ background-color: {theme['surface']}; color: {theme['text_muted']}; font-weight: bold; padding: 4px; border: none; border-bottom: 1px solid {theme['border']}; }}
    QLabel#h1 {{ font-size: 24px; font-weight: bold; }}
    QLabel#h2 {{ font-size: 18px; font-weight: bold; }}
    QLabel#stat_val {{ font-size: 28px; font-weight: bold; color: {theme['primary']}; }}
    QLabel#muted {{ color: {theme['text_muted']}; }}
    """

# ==========================================
# VIEWS
# ==========================================
class DashboardView(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        layout.addWidget(QLabel("Dashboard", objectName="h1"))
        
        stats_layout = QHBoxLayout()
        self.lbl_total_prods = self._create_stat_card("TOTAL PRODUCTS", "0", "In Database")
        self.lbl_pending = self._create_stat_card("PENDING DRAFTS", "0", "Requires attention")
        self.lbl_revenue = self._create_stat_card("REVENUE", "₹0.00", "Finalized Quotes")
        
        stats_layout.addWidget(self.lbl_total_prods[0])
        stats_layout.addWidget(self.lbl_pending[0])
        stats_layout.addWidget(self.lbl_revenue[0])
        layout.addLayout(stats_layout)
        
        layout.addWidget(QLabel("Recent Activity", objectName="h2"))
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "CLIENT", "TOTAL", "DATE", "STATUS"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)

    def _create_stat_card(self, title, val, subtitle):
        card = QFrame(objectName="surface")
        l = QVBoxLayout(card)
        l.addWidget(QLabel(title, objectName="muted"))
        val_lbl = QLabel(val, objectName="stat_val")
        l.addWidget(val_lbl)
        l.addWidget(QLabel(subtitle, objectName="muted"))
        return card, val_lbl, None

    def refresh_dashboard(self, catalog_mgr):
        # Update Products count badge safely
        self.lbl_total_prods[1].setText(f"{len(catalog_mgr.load_catalog()):,}")
        
        history = []
        total_revenue = 0.0
        
        # Scan Drafts folder cleanly
        if os.path.exists("drafts"):
            drafts = [f for f in os.listdir("drafts") if f.endswith(".xlsx")]
            self.lbl_pending[1].setText(str(len(drafts)))
            for d in drafts:
                info = self.extract_excel_meta(os.path.join("drafts", d), "Draft")
                if info: 
                    history.append(info)
        else:
            self.lbl_pending[1].setText("0")
            
        # Scan Output (Finalized) folder cleanly
        if os.path.exists("output"):
            outputs = [f for f in os.listdir("output") if f.endswith(".xlsx")]
            for o in outputs:
                info = self.extract_excel_meta(os.path.join("output", o), "Completed")
                if info: 
                    history.append(info)
                    total_revenue += info['total']
                
        self.lbl_revenue[1].setText(f"₹{total_revenue:,.2f}")
        
        # Populate GUI Table with zero-edit flags locked
        history.sort(key=lambda x: x['timestamp'], reverse=True)
        self.table.setRowCount(len(history))
        for row, item in enumerate(history):
            self.table.setItem(row, 0, QTableWidgetItem(item['id']))
            self.table.setItem(row, 1, QTableWidgetItem(item['client']))
            self.table.setItem(row, 2, QTableWidgetItem(f"₹{item['total']:,.2f}"))
            self.table.setItem(row, 3, QTableWidgetItem(item['date']))
            self.table.setItem(row, 4, QTableWidgetItem(item['status']))

    def extract_excel_meta(self, path, status):
        """Resiliently scans generated files to read quote info back to the dashboard table."""
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            
            client = "Unknown"
            qid = os.path.splitext(os.path.basename(path))[0] # Fallback to filename if cell parse drops
            total = 0.0
            
            # Smart text search through the sheet matrix to parse meta details
            for row in ws.iter_rows(values_only=True):
                for val in row:
                    if not val: 
                        continue
                    val_str = str(val).strip()
                    if "Client:" in val_str:
                        client = val_str.replace("Client:", "").strip()
                    elif "Quote #:" in val_str:
                        qid = val_str.replace("Quote #:", "").strip()
                    elif "Total:" in val_str or "Total price" in val_str:
                        # Grab the numeric value if it sits nearby in the row array
                        for potential_val in row:
                            if isinstance(potential_val, (int, float)):
                                total = float(potential_val)
            
            # Fallback evaluation: if the file has items but Excel hasn't calculated formulas yet
            if total == 0.0:
                # Scan columns to build an estimation if cells are cached as None
                for r in range(1, ws.max_row + 1):
                    val_e = ws.cell(row=r, column=5).value  # Column E calculation check
                    if isinstance(val_e, (int, float)):
                        total += float(val_e)

            date_str = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%b %d, %Y')
            return {
                "id": qid, 
                "client": client, 
                "total": total, 
                "date": date_str, 
                "status": status, 
                "timestamp": os.path.getmtime(path)
            }
        except Exception as e:
            print(f"Skipping unreadable activity file {path}: {e}")
            return None


class ProductDatabaseView(QWidget):
    def __init__(self, catalog_mgr, main_window):
        super().__init__()
        self.catalog = catalog_mgr
        self.main_window = main_window
        self.all_products = []
        self.selection_mode = False
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # --- Header ---
        header = QHBoxLayout()
        self.title_label = QLabel("Product Database", objectName="h1")
        header.addWidget(self.title_label)
        header.addStretch()
        
        self.btn_edit = QPushButton("✏️ Edit Database")
        self.btn_edit.clicked.connect(self.toggle_edit_mode)
        
        self.btn_save = QPushButton("💾 Save Changes")
        self.btn_save.setObjectName("primary_btn")
        self.btn_save.hide()
        self.btn_save.clicked.connect(self.save_database)
        
        self.btn_add_row = QPushButton("+ Add Empty Row")
        self.btn_add_row.hide()
        self.btn_add_row.clicked.connect(self.add_empty_row)
        
        header.addWidget(self.btn_edit)
        header.addWidget(self.btn_add_row)
        header.addWidget(self.btn_save)
        layout.addLayout(header)
        
        # --- Filters ---
        filter_layout = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search by Name, Category, or Supplier...")
        self.search_box.textChanged.connect(self.apply_filters)
        
        self.category_cb = QComboBox()
        self.category_cb.currentTextChanged.connect(self.apply_filters)
        self.supplier_cb = QComboBox()
        self.supplier_cb.currentTextChanged.connect(self.apply_filters)
        
        filter_layout.addWidget(QLabel("🔍 Filter:"))
        filter_layout.addWidget(self.search_box)
        filter_layout.addWidget(QLabel("Category:"))
        filter_layout.addWidget(self.category_cb)
        filter_layout.addWidget(QLabel("Supplier:"))
        filter_layout.addWidget(self.supplier_cb)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)
        
        # --- Tabs ---
        self.tabs = QTabWidget()
        
        # TAB 1: PRODUCTS (Non-Service)
        self.tab_products = QWidget()
        prod_layout = QVBoxLayout(self.tab_products)
        prod_layout.setContentsMargins(0, 10, 0, 0)
        self.products_table = QTableWidget(0, 10)
        self.products_table.setHorizontalHeaderLabels([
            "Select", "NAME", "CATEGORY", "MAKE", "MODEL", "SPECIFICATION", 
            "PRICE (₹)", "SUPPLIER", "CONTACT INFO", "LAST UPDATED"
        ])
        self.setup_table(self.products_table)
        prod_layout.addWidget(self.products_table)
        self.tabs.addTab(self.tab_products, "🛒 Products")

        # TAB 2: SERVICES
        self.tab_services = QWidget()
        serv_layout = QVBoxLayout(self.tab_services)
        serv_layout.setContentsMargins(0, 10, 0, 0)
        self.services_table = QTableWidget(0, 10)
        self.services_table.setHorizontalHeaderLabels([
            "Select", "NAME", "CATEGORY", "SKILLSET", "L1 / BASE (₹)", 
            "L2 (₹)", "L3 (₹)", "SUPPLIER", "CONTACT INFO", "LAST UPDATED"
        ])
        self.setup_table(self.services_table)
        serv_layout.addWidget(self.services_table)
        self.tabs.addTab(self.tab_services, "🛠 Services")

        layout.addWidget(self.tabs)
        
        # --- Selection Footer ---
        self.footer = QHBoxLayout()
        self.footer.addStretch()
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.cancel_selection)
        self.btn_proceed = QPushButton("Proceed with Selection →")
        self.btn_proceed.setObjectName("primary_btn")
        self.btn_proceed.clicked.connect(self.process_selection)
        
        self.footer.addWidget(self.btn_cancel)
        self.footer.addWidget(self.btn_proceed)
        self.footer_widget = QWidget()
        self.footer_widget.setLayout(self.footer)
        self.footer_widget.hide()
        layout.addWidget(self.footer_widget)

    def setup_table(self, table):
        """Helper to apply standard styling to both tables."""
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().resizeSection(0, 50)  
        table.horizontalHeader().resizeSection(1, 200) 
        table.setColumnHidden(0, True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSortingEnabled(True)
        # Pass the specific table into the cell click handler
        table.cellClicked.connect(lambda row, col, t=table: self.on_cell_clicked(row, col, t))

    def load_data(self):
        self.all_products = self.catalog.load_catalog()
        categories = ["All Categories"] + sorted(list(set(p.category for p in self.all_products if p.category)))
        suppliers = ["All Suppliers"] + sorted(list(set(p.supplier for p in self.all_products if p.supplier)))
        self.category_cb.blockSignals(True)
        self.supplier_cb.blockSignals(True)
        self.category_cb.clear()
        self.supplier_cb.clear()
        self.category_cb.addItems(categories)
        self.supplier_cb.addItems(suppliers)
        self.category_cb.blockSignals(False)
        self.supplier_cb.blockSignals(False)
        self.apply_filters()

    def apply_filters(self):
        search = self.search_box.text().lower()
        cat = self.category_cb.currentText()
        sup = self.supplier_cb.currentText()
        
        prods = []
        servs = []
        
        for p in self.all_products:
            match_search = (search in p.name.lower() or search in p.category.lower() or search in p.supplier.lower())
            match_cat = (cat == "All Categories" or p.category == cat)
            match_sup = (sup == "All Suppliers" or p.supplier == sup)
            
            if match_search and match_cat and match_sup:
                if p.category.strip().lower() == 'service':
                    servs.append(p)
                else:
                    prods.append(p)
                    
        self.populate_tables(prods, servs)

    def populate_tables(self, products, services):
        self.products_table.setSortingEnabled(False)
        self.services_table.setSortingEnabled(False)
        
        # Populate Products Tab
        self.products_table.setRowCount(len(products))
        for row, p in enumerate(products):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Unchecked)
            self.products_table.setItem(row, 0, chk)
            self.products_table.setItem(row, 1, QTableWidgetItem(p.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(p.category))
            self.products_table.setItem(row, 3, QTableWidgetItem(p.make))
            self.products_table.setItem(row, 4, QTableWidgetItem(p.model))
            self.products_table.setItem(row, 5, QTableWidgetItem(p.specification))
            
            price_item = QTableWidgetItem()
            price_item.setData(Qt.DisplayRole, p.unit_price)
            self.products_table.setItem(row, 6, price_item)
            
            self.products_table.setItem(row, 7, QTableWidgetItem(p.supplier))
            self.products_table.setItem(row, 8, QTableWidgetItem(p.supplier_contact))
            self.products_table.setItem(row, 9, QTableWidgetItem(p.last_updated))

        # Populate Services Tab
        self.services_table.setRowCount(len(services))
        for row, s in enumerate(services):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Unchecked)
            self.services_table.setItem(row, 0, chk)
            self.services_table.setItem(row, 1, QTableWidgetItem(s.name))
            self.services_table.setItem(row, 2, QTableWidgetItem(s.category))
            self.services_table.setItem(row, 3, QTableWidgetItem(s.skillset))
            
            l1_item = QTableWidgetItem()
            l1_item.setData(Qt.DisplayRole, s.unit_price)
            self.services_table.setItem(row, 4, l1_item)
            
            l2_item = QTableWidgetItem()
            l2_item.setData(Qt.DisplayRole, s.price_l2)
            self.services_table.setItem(row, 5, l2_item)
            
            l3_item = QTableWidgetItem()
            l3_item.setData(Qt.DisplayRole, s.price_l3)
            self.services_table.setItem(row, 6, l3_item)
            
            self.services_table.setItem(row, 7, QTableWidgetItem(s.supplier))
            self.services_table.setItem(row, 8, QTableWidgetItem(s.supplier_contact))
            self.services_table.setItem(row, 9, QTableWidgetItem(s.last_updated))

        self.products_table.setSortingEnabled(True)
        self.services_table.setSortingEnabled(True)

    def on_cell_clicked(self, row, col, table):
        if self.selection_mode and col != 0:
            chk = table.item(row, 0)
            if chk:
                new_state = Qt.Checked if chk.checkState() == Qt.Unchecked else Qt.Unchecked
                chk.setCheckState(new_state)

    def enable_selection_mode(self):
        self.selection_mode = True
        self.title_label.setText("Select Products for Quote")
        self.btn_edit.hide()
        self.products_table.setColumnHidden(0, False)
        self.services_table.setColumnHidden(0, False)
        self.footer_widget.show()
        self.load_data()

    def cancel_selection(self):
        self.disable_selection_mode()
        self.main_window.switch_view(2) 

    def disable_selection_mode(self):
        self.selection_mode = False
        self.title_label.setText("Product Database")
        self.btn_edit.show()
        self.products_table.setColumnHidden(0, True)
        self.services_table.setColumnHidden(0, True)
        self.footer_widget.hide()
        
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                if table.item(row, 0):
                    table.item(row, 0).setCheckState(Qt.Unchecked)

    def process_selection(self):
        selected = []
        # Grab from Products Table
        for row in range(self.products_table.rowCount()):
            if self.products_table.item(row, 0) and self.products_table.item(row, 0).checkState() == Qt.Checked:
                selected.append(Product(
                    name=self.products_table.item(row, 1).text() if self.products_table.item(row, 1) else "",
                    category=self.products_table.item(row, 2).text() if self.products_table.item(row, 2) else "",
                    make=self.products_table.item(row, 3).text() if self.products_table.item(row, 3) else "",
                    model=self.products_table.item(row, 4).text() if self.products_table.item(row, 4) else "",
                    specification=self.products_table.item(row, 5).text() if self.products_table.item(row, 5) else "",
                    description="",
                    unit_price=float(self.products_table.item(row, 6).text() or 0) if self.products_table.item(row, 6) else 0.0,
                    supplier=self.products_table.item(row, 7).text() if self.products_table.item(row, 7) else "",
                    supplier_contact=self.products_table.item(row, 8).text() if self.products_table.item(row, 8) else "",
                    last_updated=self.products_table.item(row, 9).text() if self.products_table.item(row, 9) else datetime.now().strftime("%Y-%m-%d")
                ))
                
        # Grab from Services Table
        for row in range(self.services_table.rowCount()):
            if self.services_table.item(row, 0) and self.services_table.item(row, 0).checkState() == Qt.Checked:
                selected.append(Product(
                    name=self.services_table.item(row, 1).text() if self.services_table.item(row, 1) else "",
                    category=self.services_table.item(row, 2).text() if self.services_table.item(row, 2) else "",
                    skillset=self.services_table.item(row, 3).text() if self.services_table.item(row, 3) else "",
                    description="",
                    unit_price=float(self.services_table.item(row, 4).text() or 0) if self.services_table.item(row, 4) else 0.0,
                    price_l2=float(self.services_table.item(row, 5).text() or 0) if self.services_table.item(row, 5) else 0.0,
                    price_l3=float(self.services_table.item(row, 6).text() or 0) if self.services_table.item(row, 6) else 0.0,
                    supplier=self.services_table.item(row, 7).text() if self.services_table.item(row, 7) else "",
                    supplier_contact=self.services_table.item(row, 8).text() if self.services_table.item(row, 8) else "",
                    last_updated=self.services_table.item(row, 9).text() if self.services_table.item(row, 9) else datetime.now().strftime("%Y-%m-%d")
                ))
                
        self.disable_selection_mode()
        self.main_window.finish_product_selection(selected)

    def toggle_edit_mode(self):
        self.products_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.services_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.btn_edit.hide()
        self.btn_save.show()
        self.btn_add_row.show()

    def add_empty_row(self):
        """Adds a row to whichever tab is currently active."""
        if self.tabs.currentIndex() == 0:
            target_table = self.products_table
        else:
            target_table = self.services_table
            
        target_table.setSortingEnabled(False)
        row = target_table.rowCount()
        target_table.insertRow(row)
        for c in range(1, 10):
            target_table.setItem(row, c, QTableWidgetItem(""))
            
        target_table.setItem(row, 9, QTableWidgetItem(datetime.now().strftime("%Y-%m-%d")))
        target_table.setSortingEnabled(True)

    def save_database(self):
        final_list = []
        today_str = datetime.now().strftime("%Y-%m-%d")
        
        # Read Products
        for row in range(self.products_table.rowCount()):
            pname = self.products_table.item(row, 1)
            if pname and pname.text().strip():
                date_val = self.products_table.item(row, 9).text() if self.products_table.item(row, 9) and self.products_table.item(row, 9).text().strip() else today_str
                final_list.append(Product(
                    name=pname.text(),
                    category=self.products_table.item(row, 2).text() if self.products_table.item(row, 2) else "",
                    make=self.products_table.item(row, 3).text() if self.products_table.item(row, 3) else "",
                    model=self.products_table.item(row, 4).text() if self.products_table.item(row, 4) else "",
                    specification=self.products_table.item(row, 5).text() if self.products_table.item(row, 5) else "",
                    unit_price=float(self.products_table.item(row, 6).text() or 0) if self.products_table.item(row, 6) else 0.0,
                    supplier=self.products_table.item(row, 7).text() if self.products_table.item(row, 7) else "",
                    supplier_contact=self.products_table.item(row, 8).text() if self.products_table.item(row, 8) else "",
                    last_updated=date_val,
                    description=""
                ))
                
        # Read Services
        for row in range(self.services_table.rowCount()):
            sname = self.services_table.item(row, 1)
            if sname and sname.text().strip():
                date_val = self.services_table.item(row, 9).text() if self.services_table.item(row, 9) and self.services_table.item(row, 9).text().strip() else today_str
                
                cat_val = self.services_table.item(row, 2).text() if self.services_table.item(row, 2) else ""
                if not cat_val: cat_val = "Service" # Ensure it gets caught as a service
                
                final_list.append(Product(
                    name=sname.text(),
                    category=cat_val,
                    skillset=self.services_table.item(row, 3).text() if self.services_table.item(row, 3) else "",
                    unit_price=float(self.services_table.item(row, 4).text() or 0) if self.services_table.item(row, 4) else 0.0,
                    price_l2=float(self.services_table.item(row, 5).text() or 0) if self.services_table.item(row, 5) else 0.0,
                    price_l3=float(self.services_table.item(row, 6).text() or 0) if self.services_table.item(row, 6) else 0.0,
                    supplier=self.services_table.item(row, 7).text() if self.services_table.item(row, 7) else "",
                    supplier_contact=self.services_table.item(row, 8).text() if self.services_table.item(row, 8) else "",
                    last_updated=date_val,
                    description=""
                ))
                
        self.catalog.save_catalog(final_list)
        self.products_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.services_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.btn_save.hide()
        self.btn_add_row.hide()
        self.btn_edit.show()
        QMessageBox.information(self, "Success", "Database Updated!")
        self.load_data()


class NewQuotationView(QWidget):
    def __init__(self, quote_engine, main_window):
        super().__init__()
        self.engine = quote_engine
        self.main_window = main_window
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # --- Top Client Bar ---
        top_bar = QHBoxLayout()
        self.client_name = QLineEdit()
        self.client_name.setPlaceholderText("Client Name")
        self.client_name.textChanged.connect(self.update_client_info)
        self.quote_num = QLineEdit()
        self.quote_num.setReadOnly(True)
        top_bar.addWidget(QLabel("CLIENT:"))
        top_bar.addWidget(self.client_name)
        top_bar.addWidget(QLabel("ID:"))
        top_bar.addWidget(self.quote_num)
        
        btn_import = QPushButton("📂 Import Draft")
        btn_import.clicked.connect(self.import_draft)
        btn_delete_draft = QPushButton("🗑 Delete Draft")
        btn_delete_draft.clicked.connect(self.delete_draft_file)
        top_bar.addWidget(btn_import)
        top_bar.addWidget(btn_delete_draft)
        layout.addLayout(top_bar)
        
        # --- Action Bar ---
        act_bar = QHBoxLayout()
        act_bar.addStretch()
        
        # New Preview Toggle Button
        self.btn_preview = QPushButton("👁 Show Preview")
        self.btn_preview.clicked.connect(self.toggle_preview)
        
        btn_draft = QPushButton("💾 Save Draft")
        btn_draft.clicked.connect(self.save_draft)
        btn_excel = QPushButton("📊 Export Final Excel")
        btn_excel.setObjectName("primary_btn")
        btn_excel.clicked.connect(self.export_final)
        
        act_bar.addWidget(self.btn_preview)
        act_bar.addWidget(btn_draft)
        act_bar.addWidget(btn_excel)
        layout.addLayout(act_bar)
        
        self.split_layout = QHBoxLayout()
        
        # --- Left Pane (Data Entry) ---
        self.left_pane = QFrame(objectName="surface")
        left_l = QVBoxLayout(self.left_pane)
        header_l = QHBoxLayout()
        header_l.addWidget(QLabel("Quotation Items", objectName="h2"))
        header_l.addStretch()
        btn_clear_all = QPushButton("🗑 Clear All")
        btn_clear_all.setObjectName("danger_btn") 
        btn_clear_all.clicked.connect(self.clear_all_items)
        btn_add_line = QPushButton("+ Add Products/Services")
        btn_add_line.clicked.connect(self.add_line_item)
        header_l.addWidget(btn_clear_all)
        header_l.addWidget(btn_add_line)
        left_l.addLayout(header_l)
        
        # Products Table
        left_l.addWidget(QLabel("Products", objectName="h2"))
        self.products_table = QTableWidget(0, 9)
        self.products_table.setHorizontalHeaderLabels(["#", "DESCRIPTION", "SPECIFICATION", "MAKE", "MODEL", "QTY", "UNIT PRICE", "TOTAL", "ACT"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.products_table.horizontalHeader().resizeSection(0, 40)
        self.products_table.horizontalHeader().resizeSection(5, 60) # Qty
        self.products_table.horizontalHeader().resizeSection(8, 50) # Act
        left_l.addWidget(self.products_table)
        
        # Services Table
        left_l.addWidget(QLabel("Services", objectName="h2"))
        self.services_table = QTableWidget(0, 7)
        self.services_table.setHorizontalHeaderLabels(["#", "DESCRIPTION", "SKILLSET", "LVL", "NO. OF DAYS", "PRICE", "ACT"])
        self.services_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.services_table.horizontalHeader().resizeSection(0, 40)
        self.services_table.horizontalHeader().resizeSection(4, 80) # No of days
        self.services_table.horizontalHeader().resizeSection(6, 50) # Act
        left_l.addWidget(self.services_table)
        
        # Totals
        summary_l = QGridLayout()
        summary_l.addWidget(QLabel("Subtotal"), 0, 0)
        self.lbl_subtotal = QLabel("₹0.00", alignment=Qt.AlignRight)
        summary_l.addWidget(self.lbl_subtotal, 0, 1)
        summary_l.addWidget(QLabel("Profit And Overhead %:"), 1, 0)
        self.tax_spinner = QDoubleSpinBox()
        self.tax_spinner.setRange(0, 100)
        self.tax_spinner.setValue(15.0)
        self.tax_spinner.valueChanged.connect(self.update_tax)
        summary_l.addWidget(self.tax_spinner, 1, 1)
        summary_l.addWidget(QLabel("Total Value", objectName="h2"), 2, 0)
        self.lbl_total = QLabel("₹0.00", objectName="stat_val", alignment=Qt.AlignRight)
        summary_l.addWidget(self.lbl_total, 2, 1)
        left_l.addLayout(summary_l)
        
        self.split_layout.addWidget(self.left_pane, 6)
        
        # --- Right Pane (Excel Preview - Hidden by default) ---
        self.right_pane = QFrame(objectName="surface")
        right_l = QVBoxLayout(self.right_pane)
        right_l.addWidget(QLabel("Excel Template Preview", objectName="h2"))
        
        self.excel_preview = QTableWidget(12, 8) # Expanded to 8 columns to match new format
        self.excel_preview.setHorizontalHeaderLabels(["A", "B", "C", "D", "E", "F", "G", "H"])
        self.excel_preview.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.excel_preview.horizontalHeader().setStretchLastSection(True)
        self.excel_preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        right_l.addWidget(self.excel_preview)
        
        self.right_pane.hide() # Starts hidden!
        self.split_layout.addWidget(self.right_pane, 4)
        
        layout.addLayout(self.split_layout)

    def toggle_preview(self):
        """Toggles the visibility of the preview pane and generates the data if opening."""
        is_hidden = self.right_pane.isHidden()
        self.right_pane.setVisible(is_hidden)
        self.btn_preview.setText("👁 Hide Preview" if is_hidden else "👁 Show Preview")
        
        # FIX: Force the preview to render immediately when the user unhides it!
        if is_hidden:
            self.update_excel_preview()

    def update_excel_preview(self):
        """Dynamically reads format.xlsx, injects data in memory, and renders it to the UI."""
        import os
        import openpyxl
        import string

        if not self.engine.current_quote: 
            return
            
        template_path = "data/format.xlsx"
        if not os.path.exists(template_path):
            self.excel_preview.clear()
            self.excel_preview.setRowCount(1)
            self.excel_preview.setColumnCount(1)
            self.excel_preview.setItem(0, 0, QTableWidgetItem("Template format.xlsx not found!"))
            return

        quote = self.engine.current_quote
        
        # 1. Load the template into memory
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        # 2. Define replacements
        replacements = {
            "[CLIENT_NAME]": quote.client.name,
            "[QUOTE_ID]": quote.quote_id,
            "[DATE]": quote.date.strftime('%b %d, %Y'),
            "[SUBTOTAL]": f"₹{quote.subtotal:,.2f}",
            "[TAX_RATE]": f"{(quote.tax_rate*100):.0f}%",
            "[TAX_AMOUNT]": f"₹{quote.tax_amount:,.2f}",
            "[TOTAL]": f"₹{quote.total:,.2f}"
        }

        p_start_row, p_start_col = None, None
        s_start_row, s_start_col = None, None

        # 3. Apply static tags and find anchor points
        for row in ws.iter_rows():
            for cell in row:
                if type(cell).__name__ == 'MergedCell': 
                    continue
                if isinstance(cell.value, str):
                    for tag, val in replacements.items():
                        if tag in cell.value:
                            cell.value = str(cell.value).replace(tag, str(val))
                    if "[START_PRODUCTS]" in str(cell.value):
                        p_start_row, p_start_col = cell.row, cell.column
                        cell.value = ""
                    elif "[START_SERVICES]" in str(cell.value):
                        s_start_row, s_start_col = cell.row, cell.column
                        cell.value = ""

        # 4. Helper to inject items while jumping over merged cells
        def get_next_col(r, c):
            while type(ws.cell(row=r, column=c)).__name__ == 'MergedCell':
                c += 1
            return c

        def write_items(start_r, start_c, items, is_service):
            for i, item in enumerate(items):
                r = start_r + i
                if is_service:
                    vals = [i+1, item.product.name, item.product.skillset, item.quantity, f"₹{item.subtotal:,.2f}"]
                else:
                    vals = [i+1, item.product.name, item.product.specification, item.product.make, item.product.model, item.quantity, f"₹{item.active_price:,.2f}", f"₹{item.subtotal:,.2f}"]
                
                c = start_c
                for val in vals:
                    c = get_next_col(r, c)
                    ws.cell(row=r, column=c, value=val)
                    c += 1

        products = [i for i in quote.items if i.product.category.strip().lower() != 'service']
        services = [i for i in quote.items if i.product.category.strip().lower() == 'service']

        if p_start_row: write_items(p_start_row, p_start_col, products, False)
        if s_start_row: write_items(s_start_row, s_start_col, services, True)

        # 5. Render the in-memory Excel sheet to the UI Table
        self.excel_preview.clear()
        self.excel_preview.setRowCount(ws.max_row)
        self.excel_preview.setColumnCount(ws.max_column)

        # Generate A, B, C column headers
        cols = list(string.ascii_uppercase) + [f"A{c}" for c in string.ascii_uppercase]
        self.excel_preview.setHorizontalHeaderLabels(cols[:ws.max_column])

        # Copy data
        for row in ws.iter_rows():
            for cell in row:
                if type(cell).__name__ == 'MergedCell': 
                    continue
                if cell.value is not None:
                    # PySide6 uses 0-based indexing, openpyxl uses 1-based
                    self.excel_preview.setItem(cell.row - 1, cell.column - 1, QTableWidgetItem(str(cell.value)))

        # 6. Replicate Merged Cells in the UI!
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            self.excel_preview.setSpan(min_row - 1, min_col - 1, max_row - min_row + 1, max_col - min_col + 1)

    def start_session(self):
        if not self.engine.current_quote:
            client = ClientDetails(self.client_name.text(), "", "")
            qid = self.engine.start_new_quote(client, self.tax_spinner.value() / 100.0)
            self.quote_num.setText(qid)

    def add_line_item(self):
        self.start_session()
        self.main_window.switch_to_db_selection_mode()

    def update_client_info(self):
        if self.engine.current_quote:
            self.engine.current_quote.client.name = self.client_name.text()
            self.update_excel_preview()

    def update_tax(self):
        if self.engine.current_quote:
            self.engine.current_quote.tax_rate = self.tax_spinner.value() / 100.0
            self.refresh_ui()

    def update_qty(self, absolute_index, new_qty):
        if self.engine.current_quote:
            self.engine.current_quote.items[absolute_index].quantity = new_qty
            self.refresh_ui()

    def update_service_level(self, absolute_index, combo_box_text):
        if self.engine.current_quote:
            lvl_map = {"Level 1": 1, "Level 2": 2, "Level 3": 3}
            self.engine.current_quote.items[absolute_index].service_level = lvl_map.get(combo_box_text, 1)
            self.refresh_ui()

    def delete_item(self, absolute_index):
        if self.engine.current_quote:
            self.engine.current_quote.items.pop(absolute_index)
            self.refresh_ui()

    def clear_all_items(self):
        if self.engine.current_quote and self.engine.current_quote.items:
            reply = QMessageBox.question(self, "Clear All", "Remove all items?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.engine.current_quote.items.clear()
                self.refresh_ui()

    def delete_draft_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Delete Draft", "drafts", "Excel Files (*.xlsx)")
        if file:
            filename = os.path.basename(file)
            reply = QMessageBox.question(self, "Confirm", f"Permanently delete {filename}?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    os.remove(file)
                    QMessageBox.information(self, "Deleted", "Draft deleted.")
                    self.main_window.refresh_dashboard() 
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Could not delete: {e}")

    def refresh_ui(self):
        if not self.engine.current_quote: return
        quote = self.engine.current_quote
        self.client_name.setText(quote.client.name)
        self.quote_num.setText(quote.quote_id)
        
        self.products_table.setRowCount(0)
        self.services_table.setRowCount(0)
        
        p_row = 0
        s_row = 0
        
        for absolute_idx, item in enumerate(quote.items):
            is_service = (item.product.category.strip().lower() == 'service')
            target_table = self.services_table if is_service else self.products_table
            current_row = s_row if is_service else p_row
            
            target_table.insertRow(current_row)
            target_table.setItem(current_row, 0, QTableWidgetItem(str(absolute_idx+1)))
            target_table.setItem(current_row, 1, QTableWidgetItem(item.product.name))
            
            qty_spin = QSpinBox()
            qty_spin.setRange(1, 9999)
            qty_spin.setValue(item.quantity)
            qty_spin.valueChanged.connect(lambda val, idx=absolute_idx: self.update_qty(idx, val))
            
            btn_del = QPushButton("X")
            btn_del.setObjectName("danger_btn")
            btn_del.clicked.connect(lambda checked, idx=absolute_idx: self.delete_item(idx))
            
            if is_service:
                target_table.setItem(current_row, 2, QTableWidgetItem(item.product.skillset))
                
                lvl_combo = QComboBox()
                lvl_combo.addItems(["Level 1", "Level 2", "Level 3"])
                lvl_combo.setCurrentText(f"Level {item.service_level}")
                lvl_combo.currentTextChanged.connect(lambda txt, idx=absolute_idx: self.update_service_level(idx, txt))
                
                target_table.setCellWidget(current_row, 3, lvl_combo)
                target_table.setCellWidget(current_row, 4, qty_spin)
                target_table.setItem(current_row, 5, QTableWidgetItem(f"₹{item.subtotal:,.2f}"))
                target_table.setCellWidget(current_row, 6, btn_del)
                s_row += 1
            else:
                target_table.setItem(current_row, 2, QTableWidgetItem(item.product.specification))
                target_table.setItem(current_row, 3, QTableWidgetItem(item.product.make))
                target_table.setItem(current_row, 4, QTableWidgetItem(item.product.model))
                target_table.setCellWidget(current_row, 5, qty_spin)
                target_table.setItem(current_row, 6, QTableWidgetItem(f"₹{item.active_price:,.2f}"))
                target_table.setItem(current_row, 7, QTableWidgetItem(f"₹{item.subtotal:,.2f}"))
                target_table.setCellWidget(current_row, 8, btn_del)
                p_row += 1
            
        summary = self.engine.get_summary()
        self.lbl_subtotal.setText(f"₹{summary['subtotal']:,.2f}")
        self.lbl_total.setText(f"₹{summary['total']:,.2f}")
        
        # We only update the preview if it is visible to save computing power
        if not self.right_pane.isHidden():
            self.update_excel_preview()

    def handle_save_overwrite(self, folder, default_name):
        base_path = os.path.join(folder, f"{default_name}.xlsx")
        if not os.path.exists(base_path): return base_path
        msg = QMessageBox(self)
        msg.setWindowTitle("File Exists")
        msg.setText("A file with this name already exists.")
        btn_over = msg.addButton("Overwrite", QMessageBox.AcceptRole)
        btn_copy = msg.addButton("Save as Copy", QMessageBox.ActionRole)
        btn_can = msg.addButton("Cancel", QMessageBox.RejectRole)
        msg.exec()
        if msg.clickedButton() == btn_can: return None
        if msg.clickedButton() == btn_over: return base_path
        
        count = 2
        while os.path.exists(os.path.join(folder, f"{default_name}_{count}.xlsx")):
            count += 1
        return os.path.join(folder, f"{default_name}_{count}.xlsx")

    def save_draft(self):
        # Prevent saving if the quote doesn't exist or has no items
        if not self.engine.current_quote or not self.engine.current_quote.items:
            QMessageBox.warning(self, "Empty Quote", "Please add at least one product or service before saving a draft.")
            return
            
        default = self.client_name.text() or self.quote_num.text()
        name, ok = QInputDialog.getText(self, "Save Draft", "Draft Name:", QLineEdit.Normal, default)
        
        if ok and name:
            filepath = self.handle_save_overwrite("drafts", name)
            if filepath:
                try:
                    ExcelGenerator.generate_quote_excel(self.engine.current_quote, filepath)
                    QMessageBox.information(self, "Success", "Draft saved successfully.")
                    self.main_window.refresh_dashboard()
                except Exception as e:
                    QMessageBox.critical(self, "Template Error", str(e))

    def export_final(self):
        # Prevent exporting if the quote doesn't exist or has no items
        if not self.engine.current_quote or not self.engine.current_quote.items:
            QMessageBox.warning(self, "Empty Quote", "Please add at least one product or service before exporting.")
            return
            
        filepath = self.handle_save_overwrite("output", self.quote_num.text())
        if filepath:
            try:
                ExcelGenerator.generate_quote_excel(self.engine.current_quote, filepath)
                QMessageBox.information(self, "Success", "Final quote exported.")
                self.main_window.refresh_dashboard()
            except Exception as e:
                QMessageBox.critical(self, "Template Error", str(e))

    def import_draft(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Draft", "drafts", "Excel Files (*.xlsx)")
        if file:
            try:
                self.engine.current_quote = ExcelGenerator.load_quote_excel(file, self.main_window.catalog_mgr)
                self.refresh_ui()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load draft: {e}")

# ==========================================
# MAIN WINDOW CONTROLLER
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quotation Pro")
        self.resize(1280, 720)
        self.is_dark = False 
        
        self.catalog_mgr = CatalogManager("data/catalog.xlsx")
        self.quote_engine = QuoteEngine()
        
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Sidebar
        self.sidebar = QFrame(objectName="surface")
        self.sidebar.setFixedWidth(250)
        sidebar_layout = QVBoxLayout(self.sidebar)
        
        sidebar_layout.addWidget(QLabel("Quotation Pro", objectName="h2"))
        sidebar_layout.addSpacing(30)
        
        self.btn_dash = self._make_nav_btn("📊 Dashboard", 0)
        self.btn_db = self._make_nav_btn("🗄️ Product Database", 1)
        self.btn_quote = self._make_nav_btn("📝 New Quotation", 2)
        
        sidebar_layout.addWidget(self.btn_dash)
        sidebar_layout.addWidget(self.btn_db)
        sidebar_layout.addWidget(self.btn_quote)
        sidebar_layout.addStretch()
        
        self.btn_theme = QPushButton("🌙 Toggle Theme", objectName="sidebar_btn")
        self.btn_theme.clicked.connect(self.toggle_theme)
        sidebar_layout.addWidget(self.btn_theme)
        main_layout.addWidget(self.sidebar)
        
        # Stacked Views
        self.content_stack = QStackedWidget()
        self.view_dashboard = DashboardView()
        self.view_product_db = ProductDatabaseView(self.catalog_mgr, self)
        self.view_new_quote = NewQuotationView(self.quote_engine, self)
        
        self.content_stack.addWidget(self.view_dashboard)
        self.content_stack.addWidget(self.view_product_db)
        self.content_stack.addWidget(self.view_new_quote)
        main_layout.addWidget(self.content_stack)
        
        self.apply_theme()
        self.switch_view(0)

    def _make_nav_btn(self, text, index):
        btn = QPushButton(text, objectName="sidebar_btn")
        btn.setCheckable(True)
        btn.clicked.connect(lambda: self.switch_view(index))
        return btn

    def switch_view(self, index):
        self.content_stack.setCurrentIndex(index)
        self.btn_dash.setChecked(index == 0)
        self.btn_db.setChecked(index == 1)
        self.btn_quote.setChecked(index == 2)
        
        if index == 0: self.refresh_dashboard()
        elif index == 1: self.view_product_db.load_data()

    def refresh_dashboard(self):
        self.view_dashboard.refresh_dashboard(self.catalog_mgr)

    def switch_to_db_selection_mode(self):
        """Routing from Quote View -> DB view to pick products."""
        self.switch_view(1)
        self.view_product_db.enable_selection_mode()

    def finish_product_selection(self, selected_products):
        """Routing from DB view back to Quote View with data."""
        for p in selected_products:
            self.quote_engine.add_item(p, quantity=1)
            
        self.switch_view(2)
        
        # Block signals on both new tables to prevent UI loops while rendering
        self.view_new_quote.products_table.blockSignals(True)
        self.view_new_quote.services_table.blockSignals(True)
        
        self.view_new_quote.refresh_ui()
        
        self.view_new_quote.products_table.blockSignals(False)
        self.view_new_quote.services_table.blockSignals(False)

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.btn_theme.setText("☀️ Toggle Theme" if not self.is_dark else "🌙 Toggle Theme")
        self.apply_theme()

    def apply_theme(self):
        self.setStyleSheet(get_stylesheet(self.is_dark))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())