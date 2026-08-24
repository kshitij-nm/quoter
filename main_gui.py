import sys
import os
import json
import uuid
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QLineEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QGridLayout, QAbstractItemView, 
    QDialog, QSpinBox, QDoubleSpinBox, QMessageBox, QTabWidget, QFileDialog,
    QMenu, QInputDialog, QSizePolicy, QSplitter
)
from PySide6.QtCore import Qt, QSettings
from PySide6.QtGui import QColor

from core.models import Product, QuoteItem
from core.catalog_manager import CatalogManager
from core.quote_engine import QuoteEngine
from core.excel_generator import ExcelGenerator

os.makedirs("data", exist_ok=True)
os.makedirs("output", exist_ok=True)
os.makedirs("drafts", exist_ok=True)

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
    svg_check = "data:image/svg+xml;utf8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='4' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='20 6 9 17 4 12'%3E%3C/polyline%3E%3C/svg%3E"
    
    return f"""
    QMainWindow, QWidget, QDialog {{ background-color: {theme['bg']}; color: {theme['text']}; font-family: 'Inter', sans-serif; }}
    QFrame#surface {{ background-color: {theme['surface']}; border-radius: 8px; border: 1px solid {theme['border']}; }}
    QPushButton {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 4px; padding: 8px 16px; color: {theme['text']}; }}
    QPushButton:hover {{ background-color: {theme['border']}; }}
    QPushButton#primary_btn {{ background-color: {theme['primary']}; color: #FFFFFF; border: none; font-weight: bold; }}
    QPushButton#danger_btn {{ background-color: {theme.get('danger', '#EF4444')}; color: #FFFFFF; border: none; font-weight: bold; padding: 4px 8px; }}
    
    QPushButton#sidebar_btn {{ text-align: left; padding-left: 20px; border: none; background: transparent; font-size: 14px; border-radius: 0px; }}
    QPushButton#sidebar_btn:checked {{ background-color: {theme['border']}; border-left: 4px solid {theme['primary']}; font-weight: bold; }}
    
    QPushButton#sidebar_btn_collapsed {{ text-align: center; padding: 8px 0px; border: none; background: transparent; font-size: 18px; border-radius: 0px; }}
    QPushButton#sidebar_btn_collapsed:checked {{ background-color: {theme['border']}; border-left: 4px solid {theme['primary']}; font-weight: bold; }}
    
    QPushButton#collapse_btn {{
        background-color: transparent;
        border: 1px solid {theme['border']};
        border-left: none;
        border-top-right-radius: 4px;
        border-bottom-right-radius: 4px;
        border-top-left-radius: 0px;
        border-bottom-left-radius: 0px;
        color: {theme['text_muted']};
        padding: 0px;
        font-size: 10px;
    }}
    QPushButton#collapse_btn:hover {{ background-color: {theme['border']}; color: {theme['text']}; }}
    
    QSplitter::handle {{ background-color: {theme['border']}; width: 1px; }}
    
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 4px; padding: 6px; color: {theme['text']}; }}
    
    QTableWidget {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 8px; gridline-color: {theme['border']}; outline: none; }}
    QTableView::item:focus {{ outline: none; border: none; }}
    QTableView::item:selected {{ background-color: {theme['border']}; color: {theme['text']}; }}
    
    QHeaderView::section {{ background-color: {theme['surface']}; color: {theme['text_muted']}; font-weight: bold; padding: 4px; border: none; border-bottom: 1px solid {theme['border']}; }}
    QLabel#h1 {{ font-size: 24px; font-weight: bold; }}
    QLabel#h2 {{ font-size: 18px; font-weight: bold; }}
    QLabel#stat_val {{ font-size: 32px; font-weight: bold; color: {theme['primary']}; }}
    QLabel#muted {{ color: {theme['text_muted']}; font-weight: bold; }}
    QTabWidget::pane {{ border: 1px solid {theme['border']}; border-radius: 4px; }}
    QTabBar::tab {{ background: {theme['bg']}; border: 1px solid {theme['border']}; padding: 8px 16px; border-top-left-radius: 4px; border-top-right-radius: 4px; }}
    QTabBar::tab:selected {{ background: {theme['surface']}; font-weight: bold; border-bottom: 2px solid {theme['primary']}; }}
    
    QTableView::indicator {{ width: 18px; height: 18px; border: 2px solid #94A3B8; border-radius: 4px; background-color: {theme['surface']}; margin-left: 6px; }}
    QTableView::indicator:checked {{ background-color: {theme['primary']}; border: 2px solid {theme['primary']}; image: url("{svg_check}"); }}
    """

# ==========================================
# CART DIALOG WITH DRAFTS
# ==========================================
class CartDialog(QDialog):
    def __init__(self, engine, customer_name, parent_dashboard):
        super().__init__()
        self.engine = engine
        self.customer_name = customer_name
        self.parent_dashboard = parent_dashboard
        
        self.setWindowTitle("Shopping Cart")
        self.resize(1100, 700)
        
        layout = QVBoxLayout(self)
        
        header = QHBoxLayout()
        header.addWidget(QLabel("Cart for:", objectName="h1"))
        
        self.draft_dropdown = QComboBox()
        self.draft_dropdown.setMinimumWidth(200)
        self.refresh_drafts_dropdown()
        
        idx = self.draft_dropdown.findText(self.customer_name)
        if idx >= 0:
            self.draft_dropdown.setCurrentIndex(idx)
        else:
            self.draft_dropdown.addItem(self.customer_name)
            self.draft_dropdown.setCurrentText(self.customer_name)
            
        header.addWidget(self.draft_dropdown)
        
        btn_import = QPushButton("📂 Import Draft")
        btn_import.clicked.connect(self.import_draft)
        header.addWidget(btn_import)
        
        header.addStretch()
        
        btn_clear = QPushButton("🗑 Empty Cart")
        btn_clear.setObjectName("danger_btn")
        btn_clear.clicked.connect(self.empty_cart)
        header.addWidget(btn_clear)
        layout.addLayout(header)
        
        layout.addWidget(QLabel("Products", objectName="h2"))
        self.products_table = QTableWidget(0, 7)
        self.products_table.setHorizontalHeaderLabels(["#", "DESCRIPTION", "QTY", "PRICE (L1)", "DISC %", "TOTAL", "ACT"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.products_table.horizontalHeader().resizeSection(0, 40)
        self.products_table.horizontalHeader().resizeSection(2, 60)
        self.products_table.horizontalHeader().resizeSection(4, 60)
        self.products_table.horizontalHeader().resizeSection(6, 50)
        self.products_table.verticalHeader().setVisible(False)
        layout.addWidget(self.products_table)
        
        layout.addWidget(QLabel("Services", objectName="h2"))
        self.services_table = QTableWidget(0, 8)
        self.services_table.setHorizontalHeaderLabels(["#", "DESCRIPTION", "SERVICE LVL", "QTY", "PRICE", "DISC %", "TOTAL", "ACT"])
        self.services_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.services_table.horizontalHeader().resizeSection(0, 40)
        self.services_table.horizontalHeader().resizeSection(3, 60)
        self.services_table.horizontalHeader().resizeSection(5, 60)
        self.services_table.horizontalHeader().resizeSection(7, 50) 
        self.services_table.verticalHeader().setVisible(False)
        layout.addWidget(self.services_table)
        
        bottom_bar = QHBoxLayout()
        summary_l = QGridLayout()
        
        summary_l.addWidget(QLabel("Total Value:", objectName="h2"), 0, 0)
        self.lbl_total = QLabel("₹0.00", objectName="stat_val", alignment=Qt.AlignRight)
        summary_l.addWidget(self.lbl_total, 0, 1)
        
        bottom_bar.addLayout(summary_l)
        bottom_bar.addStretch()
        
        btn_export = QPushButton("📊 Export to Excel")
        btn_export.setObjectName("primary_btn")
        btn_export.setMinimumHeight(50)
        btn_export.setMinimumWidth(200)
        btn_export.clicked.connect(self.export_final)
        bottom_bar.addWidget(btn_export)
        layout.addLayout(bottom_bar)
        self.refresh_ui()

    def refresh_drafts_dropdown(self):
        self.draft_dropdown.clear()
        drafts = [f.replace(".json", "") for f in os.listdir("drafts") if f.endswith(".json")]
        self.draft_dropdown.addItems(drafts)

    def import_draft(self):
        selected = self.draft_dropdown.currentText()
        if not selected: return
        
        filepath = f"drafts/{selected}.json"
        if not os.path.exists(filepath): return
        
        self.parent_dashboard.customer_input.setText(selected)
        self.customer_name = selected
        
        self.refresh_ui()
        self.parent_dashboard.update_cart_btn()
        self.parent_dashboard.sync_checkboxes_with_cart()
        QMessageBox.information(self, "Imported", f"Draft '{selected}' loaded successfully.")

    def update_qty(self, absolute_index, new_qty):
        cart = self.engine.get_cart(self.customer_name)
        cart.items[absolute_index].quantity = new_qty
        self.refresh_ui()
        self.parent_dashboard.silent_save_draft(self.customer_name)
        
    def update_discount(self, absolute_index, new_disc):
        cart = self.engine.get_cart(self.customer_name)
        cart.items[absolute_index].discount_percent = new_disc
        self.refresh_ui()
        self.parent_dashboard.silent_save_draft(self.customer_name)

    def update_service_level(self, absolute_index, combo_box_text):
        cart = self.engine.get_cart(self.customer_name)
        lvl_map = {"Level 1": 1, "Level 2": 2, "Level 3": 3}
        cart.items[absolute_index].service_level = lvl_map.get(combo_box_text, 1)
        self.refresh_ui()
        self.parent_dashboard.silent_save_draft(self.customer_name)

    def delete_item(self, absolute_index):
        self.engine.remove_from_cart(self.customer_name, absolute_index)
        self.refresh_ui()
        self.parent_dashboard.update_cart_btn()
        self.parent_dashboard.sync_checkboxes_with_cart()
        self.parent_dashboard.silent_save_draft(self.customer_name)

    def empty_cart(self):
        reply = QMessageBox.question(self, "Clear Cart", "Remove all items?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.engine.clear_cart(self.customer_name)
            self.parent_dashboard.update_cart_btn()
            self.parent_dashboard.sync_checkboxes_with_cart()
            self.parent_dashboard.silent_save_draft(self.customer_name)
            self.close()

    def refresh_ui(self):
        cart = self.engine.get_cart(self.customer_name)
        self.products_table.setRowCount(0)
        self.services_table.setRowCount(0)
        p_row, s_row = 0, 0
        
        for absolute_idx, item in enumerate(cart.items):
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
            
            disc_spin = QDoubleSpinBox()
            disc_spin.setRange(0, 100)
            disc_spin.setValue(item.discount_percent)
            disc_spin.valueChanged.connect(lambda val, idx=absolute_idx: self.update_discount(idx, val))
            
            btn_del = QPushButton("X")
            btn_del.setObjectName("danger_btn")
            btn_del.clicked.connect(lambda checked, idx=absolute_idx: self.delete_item(idx))
            
            if is_service:
                lvl_combo = QComboBox()
                lvl_combo.addItems(["Level 1", "Level 2", "Level 3"])
                lvl_combo.setCurrentText(f"Level {item.service_level}")
                lvl_combo.currentTextChanged.connect(lambda txt, idx=absolute_idx: self.update_service_level(idx, txt))
                
                target_table.setCellWidget(current_row, 2, lvl_combo)
                target_table.setCellWidget(current_row, 3, qty_spin) 
                target_table.setItem(current_row, 4, QTableWidgetItem(f"₹{item.active_price:,.2f}"))
                target_table.setCellWidget(current_row, 5, disc_spin)
                target_table.setItem(current_row, 6, QTableWidgetItem(f"₹{item.subtotal:,.2f}"))
                target_table.setCellWidget(current_row, 7, btn_del)
                s_row += 1
            else:
                target_table.setCellWidget(current_row, 2, qty_spin) 
                target_table.setItem(current_row, 3, QTableWidgetItem(f"₹{item.active_price:,.2f}"))
                target_table.setCellWidget(current_row, 4, disc_spin)
                target_table.setItem(current_row, 5, QTableWidgetItem(f"₹{item.subtotal:,.2f}"))
                target_table.setCellWidget(current_row, 6, btn_del)
                p_row += 1
            
        self.lbl_total.setText(f"₹{cart.total:,.2f}")

    def export_final(self):
        cart = self.engine.get_cart(self.customer_name)
        if not cart.items:
            QMessageBox.warning(self, "Empty Cart", "Cart is empty!")
            return
            
        default_name = cart.quote_id
        base_path = os.path.join("output", f"{default_name}.xlsx")
        
        if os.path.exists(base_path):
            count = 2
            while os.path.exists(os.path.join("output", f"{default_name}_{count}.xlsx")):
                count += 1
            base_path = os.path.join("output", f"{default_name}_{count}.xlsx")

        try:
            ExcelGenerator.generate_quote_excel(cart, base_path)
            QMessageBox.information(self, "Success", f"Quote exported successfully to:\n{base_path}")
            
            self.engine.clear_cart(self.customer_name)
            self.parent_dashboard.update_cart_btn()
            self.parent_dashboard.sync_checkboxes_with_cart()
            self.parent_dashboard.silent_save_draft(self.customer_name)
            
            self.close()
            self.parent_dashboard.main_window.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))


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
        stats_layout.setSpacing(20)
        self.lbl_total_prods = self._create_stat_card("TOTAL PRODUCTS", "0", "In Database")
        self.lbl_revenue = self._create_stat_card("LIFETIME REVENUE", "₹0.00", "Finalized Quotes")
        
        stats_layout.addWidget(self.lbl_total_prods[0])
        stats_layout.addWidget(self.lbl_revenue[0])
        stats_layout.addStretch()
        layout.addLayout(stats_layout)
        
        layout.addSpacing(20)
        layout.addWidget(QLabel("Recent Activity", objectName="h2"))
        
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "CLIENT", "TOTAL", "DATE"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

    def _create_stat_card(self, title, val, subtitle):
        card = QFrame(objectName="surface")
        card.setMinimumSize(300, 120)
        card.setMaximumWidth(400)
        l = QVBoxLayout(card)
        l.setContentsMargins(25, 20, 25, 20)
        l.addWidget(QLabel(title, objectName="muted"))
        val_lbl = QLabel(val, objectName="stat_val")
        l.addWidget(val_lbl)
        l.addWidget(QLabel(subtitle, objectName="muted"))
        return card, val_lbl, None

    def refresh_dashboard(self, catalog_mgr):
        self.lbl_total_prods[1].setText(f"{len(catalog_mgr.load_catalog()):,}")
        history = []
        total_revenue = 0.0
        
        if os.path.exists("output"):
            import openpyxl
            outputs = [f for f in os.listdir("output") if f.endswith(".xlsx")]
            for o in outputs:
                path = os.path.join("output", o)
                try:
                    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                    ws = wb.active
                    client = "Unknown"
                    qid = os.path.splitext(os.path.basename(path))[0]
                    total = 0.0
                    for row in ws.iter_rows(values_only=True):
                        for val in row:
                            if not val: continue
                            val_str = str(val).strip()
                            if "Client:" in val_str: client = val_str.replace("Client:", "").strip()
                            elif "Quote #:" in val_str: qid = val_str.replace("Quote #:", "").strip()
                            elif "Total:" in val_str or "Total price" in val_str:
                                for p_val in row:
                                    if isinstance(p_val, (int, float)): total = float(p_val)
                    
                    date_str = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d-%m-%Y')
                    history.append({
                        "id": qid, "client": client, "total": total, 
                        "date": date_str, "timestamp": os.path.getmtime(path)
                    })
                    total_revenue += total
                    wb.close()
                except Exception:
                    pass
                
        self.lbl_revenue[1].setText(f"₹{total_revenue:,.2f}")
        history.sort(key=lambda x: x['timestamp'], reverse=True)
        self.table.setRowCount(len(history))
        for row, item in enumerate(history):
            self.table.setItem(row, 0, QTableWidgetItem(item['id']))
            self.table.setItem(row, 1, QTableWidgetItem(item['client']))
            self.table.setItem(row, 2, QTableWidgetItem(f"₹{item['total']:,.2f}"))
            self.table.setItem(row, 3, QTableWidgetItem(item['date']))


class ProductDatabaseView(QWidget):
    def __init__(self, catalog_mgr, quote_engine, main_window):
        super().__init__()
        self.catalog = catalog_mgr
        self.engine = quote_engine
        self.main_window = main_window
        self.all_products = []
        self._is_loading = False 
        
        self.active_highlight_names = set()
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        cart_bar = QFrame(objectName="surface")
        cart_l = QHBoxLayout(cart_bar)
        cart_l.addWidget(QLabel("Current Customer:"))
        self.customer_input = QLineEdit()
        self.customer_input.setPlaceholderText("Walk-in Customer")
        
        # Connect to live draft listener
        self.customer_input.textChanged.connect(self.on_customer_changed)
        
        cart_l.addWidget(self.customer_input)
        cart_l.addStretch()
        
        self.btn_view_cart = QPushButton("🛒 View Cart (0)")
        self.btn_view_cart.setObjectName("primary_btn")
        self.btn_view_cart.clicked.connect(self.open_cart)
        cart_l.addWidget(self.btn_view_cart)
        layout.addWidget(cart_bar)
        
        header = QHBoxLayout()
        header.addWidget(QLabel("Product Database", objectName="h1"))
        header.addStretch()
        btn_add_row = QPushButton("+ Add Empty Row")
        btn_add_row.clicked.connect(self.add_empty_row)
        header.addWidget(btn_add_row)
        layout.addLayout(header)
        
        filter_layout = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search all columns...")
        self.search_box.textChanged.connect(self.apply_filters)
        
        self.category_cb = QComboBox()
        self.category_cb.currentTextChanged.connect(self.apply_filters)
        self.supplier_cb = QComboBox()
        self.supplier_cb.currentTextChanged.connect(self.apply_filters)
        
        self.btn_clear_filters = QPushButton("❌ Clear Filters")
        self.btn_clear_filters.clicked.connect(self.clear_filters)
        
        filter_layout.addWidget(QLabel("🔍 Filter:"))
        filter_layout.addWidget(self.search_box)
        filter_layout.addWidget(QLabel("Category:"))
        filter_layout.addWidget(self.category_cb)
        filter_layout.addWidget(QLabel("Supplier:"))
        filter_layout.addWidget(self.supplier_cb)
        filter_layout.addWidget(self.btn_clear_filters)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)
        
        self.tabs = QTabWidget()
        
        self.tab_products = QWidget()
        prod_layout = QVBoxLayout(self.tab_products)
        prod_layout.setContentsMargins(0, 10, 0, 0)
        self.products_table = QTableWidget(0, 13)
        self.products_table.setHorizontalHeaderLabels([
            "Add", "NAME", "CATEGORY", "MAKE", "MODEL", "SPECIFICATION", 
            "PRICE (₹)", "DISC (%)", "TOTAL (₹)", "SUPPLIER", "CONTACT INFO", "REFERENCE", "LAST UPDATED"
        ])
        self.setup_table(self.products_table, self.on_product_edited)
        prod_layout.addWidget(self.products_table)
        self.tabs.addTab(self.tab_products, "🛒 Products")

        self.tab_services = QWidget()
        serv_layout = QVBoxLayout(self.tab_services)
        serv_layout.setContentsMargins(0, 10, 0, 0)
        self.services_table = QTableWidget(0, 13)
        self.services_table.setHorizontalHeaderLabels([
            "Add", "NAME", "CATEGORY", "SKILLSET", "L1 / BASE (₹)", 
            "L2 (₹)", "L3 (₹)", "DISC (%)", "TOTAL (₹)", "SUPPLIER", "CONTACT INFO", "REFERENCE", "LAST UPDATED"
        ])
        self.setup_table(self.services_table, self.on_service_edited)
        serv_layout.addWidget(self.services_table)
        self.tabs.addTab(self.tab_services, "🛠 Services")

        layout.addWidget(self.tabs)

    def on_customer_changed(self, text):
        cust = text.strip() or "Walk-in Customer"
        self.load_draft_silently(cust)
        self.update_cart_btn()
        self.sync_checkboxes_with_cart()

    def load_draft_silently(self, cust):
        filepath = f"drafts/{cust}.json"
        cart = self.engine.get_cart(cust)
        if not os.path.exists(filepath):
            return
            
        try:
            with open(filepath, "r") as f:
                data = json.load(f)
                
            cart.items = [] 
            for item_data in data.get("items", []):
                p_data = item_data["product"]
                p = Product(
                    name=p_data.get("name", ""),
                    category=p_data.get("category", ""),
                    make=p_data.get("make", ""),
                    model=p_data.get("model", ""),
                    specification=p_data.get("specification", ""),
                    unit_price=p_data.get("unit_price", 0.0),
                    price_l2=p_data.get("price_l2", 0.0),
                    price_l3=p_data.get("price_l3", 0.0),
                    discount=p_data.get("discount", 0.0),
                    supplier=p_data.get("supplier", ""),
                    supplier_contact=p_data.get("supplier_contact", ""),
                    reference=p_data.get("reference", ""),
                    skillset=p_data.get("skillset", ""),
                    description=p_data.get("description", ""),
                    last_updated=p_data.get("last_updated", ""),
                    uid=p_data.get("uid", uuid.uuid4().hex)
                )
                qi = QuoteItem(
                    product=p,
                    quantity=item_data.get("quantity", 1),
                    discount_percent=item_data.get("discount_percent", p.discount),
                    service_level=item_data.get("service_level", 1)
                )
                cart.items.append(qi)
        except Exception:
            pass

    def setup_table(self, table, change_handler):
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        table.horizontalHeader().resizeSection(0, 35)  
        table.horizontalHeader().resizeSection(1, 200) 
        table.verticalHeader().setVisible(False)
        table.setSortingEnabled(True)
        
        # Connect new double click clearing mechanism
        table.cellDoubleClicked.connect(lambda r, c, t=table: self.on_cell_double_clicked(r, c, t))
        
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos, t=table: self.show_context_menu(pos, t))
        table.cellChanged.connect(change_handler)

    def count_name_occurrences(self, name):
        count = 0
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                item = table.item(row, 1)
                if item and item.text().strip().lower() == name:
                    count += 1
        return count

    def apply_name_highlights(self):
        highlight_color = QColor(239, 68, 68, 50)
        self._is_loading = True
        
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                name_item = table.item(row, 1)
                name = name_item.text().strip().lower() if name_item else ""
                
                is_highlighted = bool(name and name in self.active_highlight_names)
                
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item:
                        if is_highlighted:
                            item.setBackground(highlight_color)
                        else:
                            item.setData(Qt.BackgroundRole, None)
        self._is_loading = False

    def on_cell_double_clicked(self, row, col, table):
        name_item = table.item(row, 1)
        if name_item:
            name = name_item.text().strip().lower()
            if name in self.active_highlight_names:
                self.active_highlight_names.discard(name)
                self.apply_name_highlights()

    def clear_filters(self):
        self.search_box.blockSignals(True)
        self.category_cb.blockSignals(True)
        self.supplier_cb.blockSignals(True)
        self.search_box.clear()
        self.category_cb.setCurrentIndex(0)
        self.supplier_cb.setCurrentIndex(0)
        self.search_box.blockSignals(False)
        self.category_cb.blockSignals(False)
        self.supplier_cb.blockSignals(False)
        self.apply_filters()

    def show_context_menu(self, pos, table):
        row = table.rowAt(pos.y())
        if row >= 0:
            menu = QMenu(self)
            del_action = menu.addAction("🗑 Delete Row")
            action = menu.exec(table.viewport().mapToGlobal(pos))
            if action == del_action:
                reply = QMessageBox.question(self, "Verify Deletion", "Are you sure you want to permanently delete this row?", QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    self._is_loading = True
                    table.removeRow(row)
                    self._is_loading = False
                    self.save_database()

    def update_cart_btn(self):
        cust = self.customer_input.text().strip() or "Walk-in Customer"
        cart = self.engine.get_cart(cust)
        count = len(cart.items) 
        self.btn_view_cart.setText(f"🛒 View Cart ({count})")

    def sync_checkboxes_with_cart(self):
        self._is_loading = True
        cust = self.customer_input.text().strip() or "Walk-in Customer"
        cart = self.engine.get_cart(cust)
        
        # We now match by UID securely to ensure exact identical-twin rows are remembered!
        cart_uids = set(item.product.uid for item in cart.items)
        
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                chk = table.item(row, 0)
                if chk:
                    row_uid = chk.data(Qt.UserRole)
                    if row_uid in cart_uids:
                        chk.setCheckState(Qt.Checked)
                    else:
                        chk.setCheckState(Qt.Unchecked)
        self._is_loading = False

    def silent_save_draft(self, customer_name=None):
        cust = customer_name or (self.customer_input.text().strip() or "Walk-in Customer")
        cart = self.engine.get_cart(cust)
        
        if not cart.items:
            filepath = f"drafts/{cust}.json"
            if os.path.exists(filepath):
                try: os.remove(filepath)
                except: pass
            return
            
        draft_data = {"items": []}
        for i in cart.items:
            draft_data["items"].append({
                "product": {
                    "name": i.product.name,
                    "category": i.product.category,
                    "description": i.product.description,
                    "unit_price": i.product.unit_price,
                    "price_l2": i.product.price_l2,
                    "price_l3": i.product.price_l3,
                    "discount": i.product.discount,
                    "supplier": i.product.supplier,
                    "supplier_contact": i.product.supplier_contact,
                    "make": i.product.make,
                    "model": i.product.model,
                    "specification": i.product.specification,
                    "skillset": i.product.skillset,
                    "reference": i.product.reference,
                    "last_updated": i.product.last_updated,
                    "uid": i.product.uid # Ensure UID is saved!
                },
                "quantity": i.quantity,
                "discount_percent": i.discount_percent,
                "service_level": i.service_level
            })
            
        with open(f"drafts/{cust}.json", "w") as f:
            json.dump(draft_data, f)

    def open_cart(self):
        cust = self.customer_input.text().strip() or "Walk-in Customer"
        dialog = CartDialog(self.engine, cust, self)
        dialog.exec()

    def process_auto_add(self, row, table, is_service):
        chk = table.item(row, 0)
        if not chk: return

        name_item = table.item(row, 1)
        if not name_item or not name_item.text().strip():
            if chk.checkState() == Qt.Checked:
                self._is_loading = True
                chk.setCheckState(Qt.Unchecked)
                self._is_loading = False
                QMessageBox.warning(self, "Invalid Product", "You cannot add an empty row to the cart.")
            return

        cust = self.customer_input.text().strip() or "Walk-in Customer"
        product_name = name_item.text().strip()
        uid_val = chk.data(Qt.UserRole) 

        if chk.checkState() == Qt.Checked:
            qty, ok = QInputDialog.getInt(self, "Add to Cart", f"Quantity for {product_name}:", 1, 1, 9999)
            if ok:
                if is_service:
                    p = Product(
                        name=product_name,
                        category=table.item(row, 2).text() if table.item(row, 2) else "",
                        skillset=table.item(row, 3).text() if table.item(row, 3) else "",
                        unit_price=float(table.item(row, 4).text() or 0) if table.item(row, 4) else 0.0,
                        price_l2=float(table.item(row, 5).text() or 0) if table.item(row, 5) else 0.0,
                        price_l3=float(table.item(row, 6).text() or 0) if table.item(row, 6) else 0.0,
                        discount=float(table.item(row, 7).text() or 0) if table.item(row, 7) else 0.0,
                        supplier=table.item(row, 9).text() if table.item(row, 9) else "",
                        supplier_contact=table.item(row, 10).text() if table.item(row, 10) else "",
                        reference=table.item(row, 11).text() if table.item(row, 11) else "",
                        description="",
                        uid=uid_val
                    )
                else:
                    p = Product(
                        name=product_name,
                        category=table.item(row, 2).text() if table.item(row, 2) else "",
                        make=table.item(row, 3).text() if table.item(row, 3) else "",
                        model=table.item(row, 4).text() if table.item(row, 4) else "",
                        specification=table.item(row, 5).text() if table.item(row, 5) else "",
                        unit_price=float(table.item(row, 6).text() or 0) if table.item(row, 6) else 0.0,
                        discount=float(table.item(row, 7).text() or 0) if table.item(row, 7) else 0.0,
                        supplier=table.item(row, 9).text() if table.item(row, 9) else "",
                        supplier_contact=table.item(row, 10).text() if table.item(row, 10) else "",
                        reference=table.item(row, 11).text() if table.item(row, 11) else "",
                        description="",
                        uid=uid_val
                    )
                self.engine.add_to_cart(cust, p, quantity=qty)
                self.update_cart_btn()
                self.silent_save_draft(cust) 
            else:
                self._is_loading = True
                chk.setCheckState(Qt.Unchecked)
                self._is_loading = False
        else:
            cart = self.engine.get_cart(cust)
            items_to_remove = [i for i, item in enumerate(cart.items) if item.product.uid == uid_val]
            for i in reversed(items_to_remove):
                self.engine.remove_from_cart(cust, i)
            
            self.update_cart_btn()
            self.silent_save_draft(cust) 

    def on_product_edited(self, row, col):
        if self._is_loading: return
        if col == 0:
            self.process_auto_add(row, self.products_table, is_service=False)
            return
            
        self._is_loading = True
        
        # New Highlighting Trigger: Only highlight if editing NAME results in a duplicate
        if col == 1:
            name_item = self.products_table.item(row, 1)
            if name_item:
                name = name_item.text().strip().lower()
                if name:
                    count = self.count_name_occurrences(name)
                    if count > 1:
                        self.active_highlight_names.add(name)
                    else:
                        self.active_highlight_names.discard(name)
            self.apply_name_highlights()
            
        if col in [6, 7]:
            try:
                price = float(self.products_table.item(row, 6).text() or 0)
                disc = float(self.products_table.item(row, 7).text() or 0)
                total = price * (1 - (disc / 100))
                t_item = QTableWidgetItem(f"{total:,.2f}")
                t_item.setFlags(t_item.flags() & ~Qt.ItemIsEditable) 
                self.products_table.setItem(row, 8, t_item)
            except ValueError:
                pass
                
        if col == 6:
            self.products_table.setItem(row, 12, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
        
        self.save_database()
        self._is_loading = False

    def on_service_edited(self, row, col):
        if self._is_loading: return
        if col == 0:
            self.process_auto_add(row, self.services_table, is_service=True)
            return
            
        self._is_loading = True
        
        # New Highlighting Trigger
        if col == 1:
            name_item = self.services_table.item(row, 1)
            if name_item:
                name = name_item.text().strip().lower()
                if name:
                    count = self.count_name_occurrences(name)
                    if count > 1:
                        self.active_highlight_names.add(name)
                    else:
                        self.active_highlight_names.discard(name)
            self.apply_name_highlights()
            
        if col in [4, 7]:
            try:
                price = float(self.services_table.item(row, 4).text() or 0)
                disc = float(self.services_table.item(row, 7).text() or 0)
                total = price * (1 - (disc / 100))
                t_item = QTableWidgetItem(f"{total:,.2f}")
                t_item.setFlags(t_item.flags() & ~Qt.ItemIsEditable)
                self.services_table.setItem(row, 8, t_item)
            except ValueError:
                pass
                
        if col in [4, 5, 6]:
            self.services_table.setItem(row, 12, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
        
        self.save_database()
        self._is_loading = False

    def add_empty_row(self):
        self._is_loading = True
        target_table = self.products_table if self.tabs.currentIndex() == 0 else self.services_table
        target_table.setSortingEnabled(False)
        
        row = target_table.rowCount()
        target_table.insertRow(row)
        
        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        chk.setCheckState(Qt.Unchecked)
        
        # Assign a brand new UID to this new row securely!
        new_uid = uuid.uuid4().hex
        chk.setData(Qt.UserRole, new_uid)
        target_table.setItem(row, 0, chk)
        
        for c in range(1, 12):
            target_table.setItem(row, c, QTableWidgetItem(""))
            
        target_table.setItem(row, 12, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
        target_table.setSortingEnabled(True)
        self._is_loading = False

    def load_data(self):
        self._is_loading = True
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
        
        prods = [p for p in self.all_products if p.category.strip().lower() != 'service']
        servs = [p for p in self.all_products if p.category.strip().lower() == 'service']
        self.populate_tables(prods, servs)
        self.apply_filters()
        self._is_loading = False

    def apply_filters(self):
        search = self.search_box.text().lower()
        cat = self.category_cb.currentText()
        sup = self.supplier_cb.currentText()
        
        for row in range(self.products_table.rowCount()):
            name = self.products_table.item(row, 1).text() if self.products_table.item(row, 1) else ""
            category = self.products_table.item(row, 2).text() if self.products_table.item(row, 2) else ""
            make = self.products_table.item(row, 3).text() if self.products_table.item(row, 3) else ""
            model = self.products_table.item(row, 4).text() if self.products_table.item(row, 4) else ""
            spec = self.products_table.item(row, 5).text() if self.products_table.item(row, 5) else ""
            supplier = self.products_table.item(row, 9).text() if self.products_table.item(row, 9) else ""
            reference = self.products_table.item(row, 11).text() if self.products_table.item(row, 11) else ""
            
            massive_string = f"{name} {category} {make} {model} {spec} {supplier} {reference}".lower()
            match_search = (search in massive_string)
            match_cat = (cat == "All Categories" or category == cat)
            match_sup = (sup == "All Suppliers" or supplier == sup)
            
            self.products_table.setRowHidden(row, not (match_search and match_cat and match_sup))
            
        for row in range(self.services_table.rowCount()):
            name = self.services_table.item(row, 1).text() if self.services_table.item(row, 1) else ""
            category = self.services_table.item(row, 2).text() if self.services_table.item(row, 2) else ""
            skillset = self.services_table.item(row, 3).text() if self.services_table.item(row, 3) else ""
            supplier = self.services_table.item(row, 9).text() if self.services_table.item(row, 9) else ""
            reference = self.services_table.item(row, 11).text() if self.services_table.item(row, 11) else ""
            
            massive_string = f"{name} {category} {skillset} {supplier} {reference}".lower()
            match_search = (search in massive_string)
            match_cat = (cat == "All Categories" or category == cat)
            match_sup = (sup == "All Suppliers" or supplier == sup)
            
            self.services_table.setRowHidden(row, not (match_search and match_cat and match_sup))

    def populate_tables(self, products, services):
        self.products_table.setSortingEnabled(False)
        self.services_table.setSortingEnabled(False)
        
        self.products_table.setRowCount(len(products))
        for row, p in enumerate(products):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Unchecked)
            
            # Mount the UID so it can be retrieved!
            chk.setData(Qt.UserRole, p.uid)
            
            self.products_table.setItem(row, 0, chk)
            self.products_table.setItem(row, 1, QTableWidgetItem(p.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(p.category))
            self.products_table.setItem(row, 3, QTableWidgetItem(p.make))
            self.products_table.setItem(row, 4, QTableWidgetItem(p.model))
            self.products_table.setItem(row, 5, QTableWidgetItem(p.specification))
            
            p_item = QTableWidgetItem(); p_item.setData(Qt.DisplayRole, p.unit_price)
            d_item = QTableWidgetItem(); d_item.setData(Qt.DisplayRole, p.discount)
            
            total_price = p.unit_price * (1 - (p.discount / 100))
            t_item = QTableWidgetItem(f"{total_price:,.2f}")
            t_item.setFlags(t_item.flags() & ~Qt.ItemIsEditable) 
            
            self.products_table.setItem(row, 6, p_item)
            self.products_table.setItem(row, 7, d_item)
            self.products_table.setItem(row, 8, t_item)
            self.products_table.setItem(row, 9, QTableWidgetItem(p.supplier))
            self.products_table.setItem(row, 10, QTableWidgetItem(p.supplier_contact))
            self.products_table.setItem(row, 11, QTableWidgetItem(p.reference))
            self.products_table.setItem(row, 12, QTableWidgetItem(p.last_updated))

        self.services_table.setRowCount(len(services))
        for row, s in enumerate(services):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Unchecked)
            
            # Mount the UID so it can be retrieved!
            chk.setData(Qt.UserRole, s.uid)
            
            self.services_table.setItem(row, 0, chk)
            self.services_table.setItem(row, 1, QTableWidgetItem(s.name))
            self.services_table.setItem(row, 2, QTableWidgetItem(s.category))
            self.services_table.setItem(row, 3, QTableWidgetItem(s.skillset))
            
            l1 = QTableWidgetItem(); l1.setData(Qt.DisplayRole, s.unit_price)
            l2 = QTableWidgetItem(); l2.setData(Qt.DisplayRole, s.price_l2)
            l3 = QTableWidgetItem(); l3.setData(Qt.DisplayRole, s.price_l3)
            disc = QTableWidgetItem(); disc.setData(Qt.DisplayRole, s.discount)
            
            total_price = s.unit_price * (1 - (s.discount / 100))
            t_item = QTableWidgetItem(f"{total_price:,.2f}")
            t_item.setFlags(t_item.flags() & ~Qt.ItemIsEditable) 
            
            self.services_table.setItem(row, 4, l1)
            self.services_table.setItem(row, 5, l2)
            self.services_table.setItem(row, 6, l3)
            self.services_table.setItem(row, 7, disc)
            self.services_table.setItem(row, 8, t_item)
            self.services_table.setItem(row, 9, QTableWidgetItem(s.supplier))
            self.services_table.setItem(row, 10, QTableWidgetItem(s.supplier_contact))
            self.services_table.setItem(row, 11, QTableWidgetItem(s.reference))
            self.services_table.setItem(row, 12, QTableWidgetItem(s.last_updated))

        self.products_table.setSortingEnabled(True)
        self.services_table.setSortingEnabled(True)
        self.sync_checkboxes_with_cart()

    def save_database(self):
        final_list = []
        today_str = datetime.now().strftime("%d-%m-%Y")
        
        for row in range(self.products_table.rowCount()):
            pname = self.products_table.item(row, 1)
            if pname and pname.text().strip():
                uid_val = self.products_table.item(row, 0).data(Qt.UserRole)
                d_val = self.products_table.item(row, 12).text() if self.products_table.item(row, 12) and self.products_table.item(row, 12).text().strip() else today_str
                final_list.append(Product(
                    name=pname.text(),
                    category=self.products_table.item(row, 2).text() if self.products_table.item(row, 2) else "",
                    make=self.products_table.item(row, 3).text() if self.products_table.item(row, 3) else "",
                    model=self.products_table.item(row, 4).text() if self.products_table.item(row, 4) else "",
                    specification=self.products_table.item(row, 5).text() if self.products_table.item(row, 5) else "",
                    unit_price=float(self.products_table.item(row, 6).text() or 0) if self.products_table.item(row, 6) else 0.0,
                    discount=float(self.products_table.item(row, 7).text() or 0) if self.products_table.item(row, 7) else 0.0,
                    supplier=self.products_table.item(row, 9).text() if self.products_table.item(row, 9) else "",
                    supplier_contact=self.products_table.item(row, 10).text() if self.products_table.item(row, 10) else "",
                    reference=self.products_table.item(row, 11).text() if self.products_table.item(row, 11) else "",
                    last_updated=d_val,
                    uid=uid_val,
                    description=""
                ))
                
        for row in range(self.services_table.rowCount()):
            sname = self.services_table.item(row, 1)
            if sname and sname.text().strip():
                uid_val = self.services_table.item(row, 0).data(Qt.UserRole)
                d_val = self.services_table.item(row, 12).text() if self.services_table.item(row, 12) and self.services_table.item(row, 12).text().strip() else today_str
                c_val = self.services_table.item(row, 2).text() if self.services_table.item(row, 2) else "Service"
                if not c_val: c_val = "Service"
                
                final_list.append(Product(
                    name=sname.text(),
                    category=c_val,
                    skillset=self.services_table.item(row, 3).text() if self.services_table.item(row, 3) else "",
                    unit_price=float(self.services_table.item(row, 4).text() or 0) if self.services_table.item(row, 4) else 0.0,
                    price_l2=float(self.services_table.item(row, 5).text() or 0) if self.services_table.item(row, 5) else 0.0,
                    price_l3=float(self.services_table.item(row, 6).text() or 0) if self.services_table.item(row, 6) else 0.0,
                    discount=float(self.services_table.item(row, 7).text() or 0) if self.services_table.item(row, 7) else 0.0,
                    supplier=self.services_table.item(row, 9).text() if self.services_table.item(row, 9) else "",
                    supplier_contact=self.services_table.item(row, 10).text() if self.services_table.item(row, 10) else "",
                    reference=self.services_table.item(row, 11).text() if self.services_table.item(row, 11) else "",
                    last_updated=d_val,
                    uid=uid_val,
                    description=""
                ))
                
        self.catalog.save_catalog(final_list)
        self.all_products = final_list


# ==========================================
# MAIN WINDOW CONTROLLER
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.settings = QSettings("QuotationPro", "AppConfig")
        self.setWindowTitle("Quotation Pro")
        self.resize(1280, 720)
        
        self.is_dark = self.settings.value("is_dark", False, type=bool)
        self.is_collapsed = self.settings.value("sidebar_collapsed", False, type=bool)
        
        self.catalog_mgr = CatalogManager("data/catalog.xlsx")
        self.quote_engine = QuoteEngine()
        
        central = QWidget()
        self.setCentralWidget(central)
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        
        self.main_splitter = QSplitter(Qt.Horizontal)
        central_layout.addWidget(self.main_splitter)
        
        # --- FIXED SIDEBAR WRAPPER ---
        # This completely guarantees the 0-gap connection directly to the screen edge.
        self.sidebar_wrapper = QFrame(objectName="surface")
        wrapper_layout = QHBoxLayout(self.sidebar_wrapper)
        wrapper_layout.setContentsMargins(0, 0, 0, 0)
        wrapper_layout.setSpacing(0)
        
        self.sidebar_content = QWidget()
        sidebar_layout = QVBoxLayout(self.sidebar_content)
        
        self.lbl_title = QLabel()
        self.lbl_title.setObjectName("h2")
        sidebar_layout.addWidget(self.lbl_title)
        sidebar_layout.addSpacing(30)
        
        self.btn_dash = self._make_nav_btn("📊", "Dashboard", 0)
        self.btn_db = self._make_nav_btn("🗄️", "Product Database", 1)
        
        sidebar_layout.addWidget(self.btn_dash)
        sidebar_layout.addWidget(self.btn_db)
        sidebar_layout.addStretch()
        
        self.btn_theme = QPushButton()
        self.btn_theme.clicked.connect(self.toggle_theme)
        sidebar_layout.addWidget(self.btn_theme)
        
        wrapper_layout.addWidget(self.sidebar_content)
        
        # Collapse Arrow Container
        toggle_container = QWidget()
        toggle_container.setFixedWidth(16)
        toggle_layout = QVBoxLayout(toggle_container)
        toggle_layout.setContentsMargins(0, 0, 0, 0)
        toggle_layout.addStretch()
        
        self.btn_toggle_sidebar = QPushButton()
        self.btn_toggle_sidebar.setObjectName("collapse_btn")
        self.btn_toggle_sidebar.setFixedSize(16, 50)
        self.btn_toggle_sidebar.clicked.connect(self.toggle_sidebar)
        
        toggle_layout.addWidget(self.btn_toggle_sidebar)
        toggle_layout.addStretch()
        
        wrapper_layout.addWidget(toggle_container)
        self.main_splitter.addWidget(self.sidebar_wrapper)
        
        # --- RIGHT CONTENT STACK ---
        self.content_stack = QStackedWidget()
        self.view_dashboard = DashboardView()
        self.view_product_db = ProductDatabaseView(self.catalog_mgr, self.quote_engine, self)
        
        self.content_stack.addWidget(self.view_dashboard)
        self.content_stack.addWidget(self.view_product_db)
        
        self.main_splitter.addWidget(self.content_stack)
        
        self.expanded_width = 250
        self._update_sidebar_ui()
        
        geom = self.settings.value("geometry")
        if geom: self.restoreGeometry(geom)
        
        state = self.settings.value("windowState")
        if state: self.restoreState(state)
        
        splitter_state = self.settings.value("splitterState")
        if splitter_state: self.main_splitter.restoreState(splitter_state)
        
        self.switch_view(1)
        
        dash_st = self.settings.value("dash_table")
        if dash_st: self.view_dashboard.table.horizontalHeader().restoreState(dash_st)
            
        prod_st = self.settings.value("prod_table")
        if prod_st: self.view_product_db.products_table.horizontalHeader().restoreState(prod_st)
            
        serv_st = self.settings.value("serv_table")
        if serv_st: self.view_product_db.services_table.horizontalHeader().restoreState(serv_st)

    def _make_nav_btn(self, icon, text, index):
        btn = QPushButton()
        btn.icon_text = icon
        btn.full_text = text
        btn.setCheckable(True)
        btn.clicked.connect(lambda: self.switch_view(index))
        return btn

    def _update_sidebar_ui(self):
        if self.is_collapsed:
            self.sidebar_content.setFixedWidth(60)
            self.sidebar_wrapper.setFixedWidth(76) 
            
            self.lbl_title.setText("QP")
            self.lbl_title.setAlignment(Qt.AlignCenter)
            
            self.btn_dash.setText(self.btn_dash.icon_text)
            self.btn_db.setText(self.btn_db.icon_text)
            
            theme_icon = "☀️" if self.is_dark else "🌙"
            self.btn_theme.setText(theme_icon)
            
            obj_name = "sidebar_btn_collapsed"
            self.btn_toggle_sidebar.setText("▶")
        else:
            self.sidebar_content.setMinimumWidth(200)
            self.sidebar_content.setMaximumWidth(16777215)  
            
            self.sidebar_wrapper.setMinimumWidth(216)
            self.sidebar_wrapper.setMaximumWidth(16777215)
            
            self.lbl_title.setText("Quotation Pro")
            self.lbl_title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            
            self.btn_dash.setText(f"{self.btn_dash.icon_text}  {self.btn_dash.full_text}")
            self.btn_db.setText(f"{self.btn_db.icon_text}  {self.btn_db.full_text}")
            
            theme_icon = "☀️" if self.is_dark else "🌙"
            theme_text = "Light Mode" if self.is_dark else "Dark Mode"
            self.btn_theme.setText(f"{theme_icon}  {theme_text}")
            
            obj_name = "sidebar_btn"
            self.btn_toggle_sidebar.setText("◀")
            
        self.btn_dash.setObjectName(obj_name)
        self.btn_db.setObjectName(obj_name)
        self.btn_theme.setObjectName(obj_name)
        
        self.apply_theme()

    def toggle_sidebar(self):
        if not self.is_collapsed:
            self.expanded_width = self.sidebar_wrapper.width()
            self.is_collapsed = True
        else:
            self.is_collapsed = False
            
        self._update_sidebar_ui()
        
        if not self.is_collapsed:
            target_w = getattr(self, 'expanded_width', 266)
            self.main_splitter.setSizes([target_w, self.width() - target_w])

    def switch_view(self, index):
        self.content_stack.setCurrentIndex(index)
        self.btn_dash.setChecked(index == 0)
        self.btn_db.setChecked(index == 1)
        
        if index == 0: 
            self.refresh_dashboard()
        elif index == 1: 
            if not self.view_product_db.all_products:
                self.view_product_db.load_data()

    def refresh_dashboard(self):
        self.view_dashboard.refresh_dashboard(self.catalog_mgr)

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self._update_sidebar_ui() 

    def apply_theme(self):
        self.setStyleSheet(get_stylesheet(self.is_dark))
        
    def closeEvent(self, event):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        self.settings.setValue("splitterState", self.main_splitter.saveState())
        self.settings.setValue("is_dark", self.is_dark)
        self.settings.setValue("sidebar_collapsed", self.is_collapsed)
        
        self.settings.setValue("dash_table", self.view_dashboard.table.horizontalHeader().saveState())
        self.settings.setValue("prod_table", self.view_product_db.products_table.horizontalHeader().saveState())
        self.settings.setValue("serv_table", self.view_product_db.services_table.horizontalHeader().saveState())
        
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())