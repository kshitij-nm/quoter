import sys
import os
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QLineEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QGridLayout, QAbstractItemView, 
    QDialog, QSpinBox, QDoubleSpinBox, QMessageBox, QTabWidget, QFileDialog,
    QMenu, QInputDialog, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor  # Added to support duplicate highlighting

from core.models import Product
from core.catalog_manager import CatalogManager
from core.quote_engine import QuoteEngine
from core.excel_generator import ExcelGenerator

os.makedirs("data", exist_ok=True)
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
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{ background-color: {theme['surface']}; border: 1px solid {theme['border']}; border-radius: 4px; padding: 6px; color: {theme['text']}; }}
    
    QTableWidget {{ 
        background-color: {theme['surface']}; 
        border: 1px solid {theme['border']}; 
        border-radius: 8px; 
        gridline-color: {theme['border']}; 
        outline: none; 
    }}
    QTableView::item:focus {{ outline: none; border: none; }}
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
# CART DIALOG
# ==========================================
class CartDialog(QDialog):
    def __init__(self, engine, customer_name, parent_dashboard):
        super().__init__()
        self.engine = engine
        self.customer_name = customer_name
        self.parent_dashboard = parent_dashboard
        
        self.setWindowTitle(f"Shopping Cart - {self.customer_name}")
        self.resize(1100, 700)
        
        layout = QVBoxLayout(self)
        header = QHBoxLayout()
        header.addWidget(QLabel(f"Cart for: {self.customer_name}", objectName="h1"))
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

    def update_qty(self, absolute_index, new_qty):
        cart = self.engine.get_cart(self.customer_name)
        cart.items[absolute_index].quantity = new_qty
        self.refresh_ui()
        
    def update_discount(self, absolute_index, new_disc):
        cart = self.engine.get_cart(self.customer_name)
        cart.items[absolute_index].discount_percent = new_disc
        self.refresh_ui()

    def update_service_level(self, absolute_index, combo_box_text):
        cart = self.engine.get_cart(self.customer_name)
        lvl_map = {"Level 1": 1, "Level 2": 2, "Level 3": 3}
        cart.items[absolute_index].service_level = lvl_map.get(combo_box_text, 1)
        self.refresh_ui()

    def delete_item(self, absolute_index):
        self.engine.remove_from_cart(self.customer_name, absolute_index)
        self.refresh_ui()
        self.parent_dashboard.update_cart_btn()

    def empty_cart(self):
        reply = QMessageBox.question(self, "Clear Cart", "Remove all items?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.engine.clear_cart(self.customer_name)
            self.parent_dashboard.update_cart_btn()
            self.parent_dashboard.uncheck_all_boxes()
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
            self.parent_dashboard.uncheck_all_boxes()
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
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        cart_bar = QFrame(objectName="surface")
        cart_l = QHBoxLayout(cart_bar)
        cart_l.addWidget(QLabel("Current Customer:"))
        self.customer_input = QLineEdit()
        self.customer_input.setPlaceholderText("Walk-in Customer")
        self.customer_input.textChanged.connect(self.update_cart_btn)
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
        self.products_table = QTableWidget(0, 12)
        self.products_table.setHorizontalHeaderLabels([
            "Add", "NAME", "CATEGORY", "MAKE", "MODEL", "SPECIFICATION", 
            "PRICE (₹)", "DISC (%)", "SUPPLIER", "CONTACT INFO", "REFERENCE", "LAST UPDATED"
        ])
        self.setup_table(self.products_table, self.on_product_edited)
        prod_layout.addWidget(self.products_table)
        self.tabs.addTab(self.tab_products, "🛒 Products")

        self.tab_services = QWidget()
        serv_layout = QVBoxLayout(self.tab_services)
        serv_layout.setContentsMargins(0, 10, 0, 0)
        self.services_table = QTableWidget(0, 12)
        self.services_table.setHorizontalHeaderLabels([
            "Add", "NAME", "CATEGORY", "SKILLSET", "L1 / BASE (₹)", 
            "L2 (₹)", "L3 (₹)", "DISC (%)", "SUPPLIER", "CONTACT INFO", "REFERENCE", "LAST UPDATED"
        ])
        self.setup_table(self.services_table, self.on_service_edited)
        serv_layout.addWidget(self.services_table)
        self.tabs.addTab(self.tab_services, "🛠 Services")

        layout.addWidget(self.tabs)

    def setup_table(self, table, change_handler):
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        table.horizontalHeader().resizeSection(0, 35)  
        table.horizontalHeader().resizeSection(1, 200) 
        table.verticalHeader().setVisible(False)
        table.setSortingEnabled(True)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos, t=table: self.show_context_menu(pos, t))
        table.cellChanged.connect(change_handler)

    def highlight_duplicates(self):
        """Finds matches where the ENTIRE row's data is identical (ignoring the date)."""
        self._is_loading = True
        row_counts = {}
        
        # Pass 1: Build a signature for the entire row (Columns 1 through 10)
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                row_data = []
                for col in range(1, 11): # Checks Name, Category, Specs, Prices, Supplier, Ref, etc.
                    item = table.item(row, col)
                    row_data.append(item.text().strip().lower() if item else "")
                
                # Only count rows that actually have a name
                if row_data[0]: 
                    signature = "|".join(row_data)
                    row_counts[signature] = row_counts.get(signature, 0) + 1
                    
        # Soft Red Highlight color for duplicates
        highlight_color = QColor(239, 68, 68, 50)  
        
        # Pass 2: Apply coloring to rows whose full signature appears more than once
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                row_data = []
                for col in range(1, 11):
                    item = table.item(row, col)
                    row_data.append(item.text().strip().lower() if item else "")
                
                is_duplicate = False
                if row_data[0]:
                    signature = "|".join(row_data)
                    if row_counts.get(signature, 0) > 1:
                        is_duplicate = True
                        
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item:
                        if is_duplicate:
                            item.setBackground(highlight_color)
                        else:
                            item.setData(Qt.BackgroundRole, None)  # Clears the highlight
        self._is_loading = False

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

    def uncheck_all_boxes(self):
        self._is_loading = True
        for table in [self.products_table, self.services_table]:
            for row in range(table.rowCount()):
                chk = table.item(row, 0)
                if chk: chk.setCheckState(Qt.Unchecked)
        self._is_loading = False

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
                        supplier=table.item(row, 8).text() if table.item(row, 8) else "",
                        supplier_contact=table.item(row, 9).text() if table.item(row, 9) else "",
                        reference=table.item(row, 10).text() if table.item(row, 10) else "",
                        description="",
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
                        supplier=table.item(row, 8).text() if table.item(row, 8) else "",
                        supplier_contact=table.item(row, 9).text() if table.item(row, 9) else "",
                        reference=table.item(row, 10).text() if table.item(row, 10) else "",
                        description="",
                    )
                self.engine.add_to_cart(cust, p, quantity=qty)
                self.update_cart_btn()
            else:
                self._is_loading = True
                chk.setCheckState(Qt.Unchecked)
                self._is_loading = False
        else:
            cart = self.engine.get_cart(cust)
            items_to_remove = [i for i, item in enumerate(cart.items) if item.product.name == product_name]
            for i in reversed(items_to_remove):
                self.engine.remove_from_cart(cust, i)
            if items_to_remove:
                self.update_cart_btn()

    def on_product_edited(self, row, col):
        if self._is_loading: return
        if col == 0:
            self.process_auto_add(row, self.products_table, is_service=False)
            return
            
        self._is_loading = True
        if col == 6:
            self.products_table.setItem(row, 11, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
        
        self.save_database()
        self._is_loading = False

    def on_service_edited(self, row, col):
        if self._is_loading: return
        if col == 0:
            self.process_auto_add(row, self.services_table, is_service=True)
            return
            
        self._is_loading = True
        if col in [4, 5, 6]:
            self.services_table.setItem(row, 11, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
        
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
        target_table.setItem(row, 0, chk)
        
        for c in range(1, 11):
            target_table.setItem(row, c, QTableWidgetItem(""))
            
        target_table.setItem(row, 11, QTableWidgetItem(datetime.now().strftime("%d-%m-%Y")))
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
            supplier = self.products_table.item(row, 8).text() if self.products_table.item(row, 8) else ""
            reference = self.products_table.item(row, 10).text() if self.products_table.item(row, 10) else ""
            
            massive_string = f"{name} {category} {make} {model} {spec} {supplier} {reference}".lower()
            match_search = (search in massive_string)
            match_cat = (cat == "All Categories" or category == cat)
            match_sup = (sup == "All Suppliers" or supplier == sup)
            
            self.products_table.setRowHidden(row, not (match_search and match_cat and match_sup))
            
        for row in range(self.services_table.rowCount()):
            name = self.services_table.item(row, 1).text() if self.services_table.item(row, 1) else ""
            category = self.services_table.item(row, 2).text() if self.services_table.item(row, 2) else ""
            skillset = self.services_table.item(row, 3).text() if self.services_table.item(row, 3) else ""
            supplier = self.services_table.item(row, 8).text() if self.services_table.item(row, 8) else ""
            reference = self.services_table.item(row, 10).text() if self.services_table.item(row, 10) else ""
            
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
            self.products_table.setItem(row, 0, chk)
            self.products_table.setItem(row, 1, QTableWidgetItem(p.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(p.category))
            self.products_table.setItem(row, 3, QTableWidgetItem(p.make))
            self.products_table.setItem(row, 4, QTableWidgetItem(p.model))
            self.products_table.setItem(row, 5, QTableWidgetItem(p.specification))
            
            p_item = QTableWidgetItem(); p_item.setData(Qt.DisplayRole, p.unit_price)
            d_item = QTableWidgetItem(); d_item.setData(Qt.DisplayRole, p.discount)
            
            self.products_table.setItem(row, 6, p_item)
            self.products_table.setItem(row, 7, d_item)
            self.products_table.setItem(row, 8, QTableWidgetItem(p.supplier))
            self.products_table.setItem(row, 9, QTableWidgetItem(p.supplier_contact))
            self.products_table.setItem(row, 10, QTableWidgetItem(p.reference))
            self.products_table.setItem(row, 11, QTableWidgetItem(p.last_updated))

        self.services_table.setRowCount(len(services))
        for row, s in enumerate(services):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Unchecked)
            self.services_table.setItem(row, 0, chk)
            self.services_table.setItem(row, 1, QTableWidgetItem(s.name))
            self.services_table.setItem(row, 2, QTableWidgetItem(s.category))
            self.services_table.setItem(row, 3, QTableWidgetItem(s.skillset))
            
            l1 = QTableWidgetItem(); l1.setData(Qt.DisplayRole, s.unit_price)
            l2 = QTableWidgetItem(); l2.setData(Qt.DisplayRole, s.price_l2)
            l3 = QTableWidgetItem(); l3.setData(Qt.DisplayRole, s.price_l3)
            disc = QTableWidgetItem(); disc.setData(Qt.DisplayRole, s.discount)
            
            self.services_table.setItem(row, 4, l1)
            self.services_table.setItem(row, 5, l2)
            self.services_table.setItem(row, 6, l3)
            self.services_table.setItem(row, 7, disc)
            self.services_table.setItem(row, 8, QTableWidgetItem(s.supplier))
            self.services_table.setItem(row, 9, QTableWidgetItem(s.supplier_contact))
            self.services_table.setItem(row, 10, QTableWidgetItem(s.reference))
            self.services_table.setItem(row, 11, QTableWidgetItem(s.last_updated))

        self.products_table.setSortingEnabled(True)
        self.services_table.setSortingEnabled(True)
        self.highlight_duplicates()

    def save_database(self):
        final_list = []
        today_str = datetime.now().strftime("%d-%m-%Y")
        
        for row in range(self.products_table.rowCount()):
            pname = self.products_table.item(row, 1)
            if pname and pname.text().strip():
                d_val = self.products_table.item(row, 11).text() if self.products_table.item(row, 11) and self.products_table.item(row, 11).text().strip() else today_str
                final_list.append(Product(
                    name=pname.text(),
                    category=self.products_table.item(row, 2).text() if self.products_table.item(row, 2) else "",
                    make=self.products_table.item(row, 3).text() if self.products_table.item(row, 3) else "",
                    model=self.products_table.item(row, 4).text() if self.products_table.item(row, 4) else "",
                    specification=self.products_table.item(row, 5).text() if self.products_table.item(row, 5) else "",
                    unit_price=float(self.products_table.item(row, 6).text() or 0) if self.products_table.item(row, 6) else 0.0,
                    discount=float(self.products_table.item(row, 7).text() or 0) if self.products_table.item(row, 7) else 0.0,
                    supplier=self.products_table.item(row, 8).text() if self.products_table.item(row, 8) else "",
                    supplier_contact=self.products_table.item(row, 9).text() if self.products_table.item(row, 9) else "",
                    reference=self.products_table.item(row, 10).text() if self.products_table.item(row, 10) else "",
                    last_updated=d_val,
                    description=""
                ))
                
        for row in range(self.services_table.rowCount()):
            sname = self.services_table.item(row, 1)
            if sname and sname.text().strip():
                d_val = self.services_table.item(row, 11).text() if self.services_table.item(row, 11) and self.services_table.item(row, 11).text().strip() else today_str
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
                    supplier=self.services_table.item(row, 8).text() if self.services_table.item(row, 8) else "",
                    supplier_contact=self.services_table.item(row, 9).text() if self.services_table.item(row, 9) else "",
                    reference=self.services_table.item(row, 10).text() if self.services_table.item(row, 10) else "",
                    last_updated=d_val,
                    description=""
                ))
                
        self.catalog.save_catalog(final_list)
        self.all_products = final_list
        self.highlight_duplicates()


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
        
        self.sidebar = QFrame(objectName="surface")
        self.sidebar.setFixedWidth(250)
        sidebar_layout = QVBoxLayout(self.sidebar)
        
        sidebar_layout.addWidget(QLabel("Quotation Pro", objectName="h2"))
        sidebar_layout.addSpacing(30)
        
        self.btn_dash = self._make_nav_btn("📊 Dashboard", 0)
        self.btn_db = self._make_nav_btn("🗄️ Product Database", 1)
        
        sidebar_layout.addWidget(self.btn_dash)
        sidebar_layout.addWidget(self.btn_db)
        sidebar_layout.addStretch()
        
        self.btn_theme = QPushButton("🌙 Toggle Theme", objectName="sidebar_btn")
        self.btn_theme.clicked.connect(self.toggle_theme)
        sidebar_layout.addWidget(self.btn_theme)
        main_layout.addWidget(self.sidebar)
        
        self.content_stack = QStackedWidget()
        self.view_dashboard = DashboardView()
        self.view_product_db = ProductDatabaseView(self.catalog_mgr, self.quote_engine, self)
        
        self.content_stack.addWidget(self.view_dashboard)
        self.content_stack.addWidget(self.view_product_db)
        main_layout.addWidget(self.content_stack)
        
        self.apply_theme()
        self.btn_theme.setText("☀️ Toggle Theme" if not self.is_dark else "🌙 Toggle Theme")
        self.switch_view(1)

    def _make_nav_btn(self, text, index):
        btn = QPushButton(text, objectName="sidebar_btn")
        btn.setCheckable(True)
        btn.clicked.connect(lambda: self.switch_view(index))
        return btn

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
        self.btn_theme.setText("☀️ Toggle Theme" if not self.is_dark else "🌙 Toggle Theme")
        self.apply_theme()

    def apply_theme(self):
        self.setStyleSheet(get_stylesheet(self.is_dark))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())