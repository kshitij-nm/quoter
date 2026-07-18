import os
import logging
from openpyxl import load_workbook
from .models import Quotation, ClientDetails, QuoteItem

logger = logging.getLogger(__name__)

class ExcelGenerator:
    @staticmethod
    def generate_quote_excel(quote: Quotation, output_path: str, template_path: str = "data/format.xlsx"):
        logger.info(f"Generating Excel via template for quote: {quote.quote_id}")
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found at {template_path}. Please create a format.xlsx file.")

        wb = load_workbook(template_path)
        ws = wb.active

        replacements = {
            "[CLIENT_NAME]": quote.client.name,
            "[QUOTE_ID]": quote.quote_id,
            "[DATE]": quote.date.strftime('%b %d, %Y'),
            "[SUBTOTAL]": round(quote.subtotal, 2),
            "[TAX_RATE]": f"{(quote.tax_rate*100):.0f}%",
            "[TAX_AMOUNT]": round(quote.tax_amount, 2),
            "[TOTAL]": round(quote.total, 2)
        }

        p_start_row, p_start_col = None, None
        s_start_row, s_start_col = None, None
        legacy_start_row, legacy_start_col = None, None

        # 1. Replace static tags
        for row in ws.iter_rows():
            for cell in row:
                # Safely skip background merged cells
                if type(cell).__name__ == 'MergedCell':
                    continue

                if isinstance(cell.value, str):
                    for tag, value in replacements.items():
                        if tag in cell.value:
                            if cell.value.strip() == tag and isinstance(value, (int, float)):
                                cell.value = value
                            else:
                                cell.value = str(cell.value).replace(tag, str(value))
                    
                    if "[START_PRODUCTS]" in str(cell.value):
                        p_start_row, p_start_col = cell.row, cell.column
                        cell.value = ""
                    elif "[START_SERVICES]" in str(cell.value):
                        s_start_row, s_start_col = cell.row, cell.column
                        cell.value = ""
                    elif "[START_ITEMS]" in str(cell.value):
                        legacy_start_row, legacy_start_col = cell.row, cell.column
                        cell.value = ""

        # 2. Split Items
        products = [i for i in quote.items if i.product.category.strip().lower() != 'service']
        services = [i for i in quote.items if i.product.category.strip().lower() == 'service']

        # 3. Write Items
        if legacy_start_row: 
            ExcelGenerator._write_products(ws, legacy_start_row, legacy_start_col, quote.items)
        else:
            if p_start_row:
                ExcelGenerator._write_products(ws, p_start_row, p_start_col, products)
            if s_start_row:
                ExcelGenerator._write_services(ws, s_start_row, s_start_col, services)

        wb.save(output_path)

    @staticmethod
    def _get_next_unmerged_col(worksheet, row, start_col):
        """Helper to find the next available column that isn't a locked MergedCell."""
        col = start_col
        while type(worksheet.cell(row=row, column=col)).__name__ == 'MergedCell':
            col += 1
        return col

    @staticmethod
    def _write_products(worksheet, start_row, start_col, items_list):
        """Matches: SL. No. | Description | Specification | Make | Model | Qty | Unit Price | Total price"""
        for i, item in enumerate(items_list):
            r = start_row + i
            values_to_write = [
                i + 1,
                item.product.name,
                item.product.specification,
                item.product.make,
                item.product.model,
                item.quantity,
                item.active_price,
                item.subtotal
            ]
            
            c = start_col
            for val in values_to_write:
                c = ExcelGenerator._get_next_unmerged_col(worksheet, r, c)
                worksheet.cell(row=r, column=c, value=val)
                c += 1 

    @staticmethod
    def _write_services(worksheet, start_row, start_col, items_list):
        """Matches: SL.No. | Description | SkillSet | No. of days | Price"""
        for i, item in enumerate(items_list):
            r = start_row + i
            values_to_write = [
                i + 1,
                item.product.name,
                item.product.skillset,
                item.quantity,  # No. of days
                item.subtotal   # Final Price
            ]
            
            c = start_col
            for val in values_to_write:
                c = ExcelGenerator._get_next_unmerged_col(worksheet, r, c)
                worksheet.cell(row=r, column=c, value=val)
                c += 1

    @staticmethod
    def load_quote_excel(filepath: str, catalog_mgr) -> Quotation:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        
        client_name = ""
        quote_id = "Imported-Draft"
        
        for row in ws.iter_rows(min_row=1, max_row=20):
            for cell in row:
                if type(cell).__name__ == 'MergedCell':
                    continue
                if isinstance(cell.value, str) and "QT-" in cell.value:
                    quote_id = cell.value
        
        quote = Quotation(quote_id=quote_id, client=ClientDetails(name=client_name, company="", email=""))
        catalog_products = {p.name: p for p in catalog_mgr.load_catalog()}
        
        for row in ws.iter_rows():
            for cell in row:
                if type(cell).__name__ == 'MergedCell':
                    continue

                if isinstance(cell.value, str) and cell.value in catalog_products:
                    product_name = cell.value
                    
                    # Find Qty and Price safely, skipping merged cells
                    qty_col = ExcelGenerator._get_next_unmerged_col(ws, cell.row, cell.column + 1)
                    qty_cell = ws.cell(row=cell.row, column=qty_col).value
                    
                    price_col = ExcelGenerator._get_next_unmerged_col(ws, cell.row, qty_col + 1)
                    price_cell = ws.cell(row=cell.row, column=price_col).value
                    
                    try: qty = int(qty_cell) if qty_cell else 1
                    except ValueError: qty = 1
                    
                    try: loaded_price = float(price_cell) if price_cell else 0.0
                    except ValueError: loaded_price = 0.0
                    
                    p = catalog_products[product_name]
                    
                    level = 1
                    if p.category.lower() == 'service':
                        if loaded_price == p.price_l3 and p.price_l3 > 0: level = 3
                        elif loaded_price == p.price_l2 and p.price_l2 > 0: level = 2
                        
                    quote.items.append(QuoteItem(product=p, quantity=qty, service_level=level))

        return quote