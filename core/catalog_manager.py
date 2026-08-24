import logging
import pandas as pd
from datetime import datetime
import uuid
from typing import List
from .models import Product

logger = logging.getLogger(__name__)

class CatalogManager:
    def __init__(self, filepath: str):
        self.filepath = filepath

    def load_catalog(self) -> List[Product]:
        try:
            products = []
            with pd.ExcelFile(self.filepath, engine='openpyxl') as xls:
                sheet_names = xls.sheet_names
                if 'Products' in sheet_names and 'Services' in sheet_names:
                    df_p = pd.read_excel(xls, sheet_name='Products', dtype=str).fillna('')
                    df_s = pd.read_excel(xls, sheet_name='Services', dtype=str).fillna('')
                    dfs = [(df_p, False), (df_s, True)]
                else:
                    df = pd.read_excel(xls, sheet_name=0, dtype=str).fillna('')
                    dfs = [(df, False)]

            for df, is_service_sheet in dfs:
                for _, row in df.iterrows():
                    raw_date = row.get('Last Updated', '')
                    if not str(raw_date).strip():
                        date_str = datetime.now().strftime("%d-%m-%Y")
                    else:
                        date_str = str(raw_date).split()[0]
                        
                    contact = str(row.get('Contact Info', '')).strip()
                    if contact.endswith('.0'):
                        contact = contact[:-2]
                        
                    def parse_float(val):
                        try: return float(val) if str(val).strip() else 0.0
                        except ValueError: return 0.0

                    cat = str(row.get('Category', ''))
                    if is_service_sheet and not cat:
                        cat = 'Service'
                        
                    products.append(Product(
                        name=str(row.get('Name', '')),
                        category=cat,
                        description=str(row.get('Description', '')),
                        unit_price=parse_float(row.get('Price L1', row.get('Unit Price', 0.0))),
                        price_l2=parse_float(row.get('Price L2', 0.0)),
                        price_l3=parse_float(row.get('Price L3', 0.0)),
                        supplier=str(row.get('Supplier', '')),
                        supplier_contact=contact,
                        make=str(row.get('Make', '')),
                        model=str(row.get('Model', '')),
                        specification=str(row.get('Specification', '')),
                        skillset=str(row.get('SkillSet', '')),
                        discount=parse_float(row.get('Discount (%)', 0.0)),  
                        reference=str(row.get('Reference', '')),             
                        last_updated=date_str,
                        uid=str(row.get('_UID', uuid.uuid4().hex))
                    ))
            return products
        except FileNotFoundError:
            return []

    def save_catalog(self, products: List[Product]):
        prods_data = []
        servs_data = []
        
        for p in products:
            total_price = round(p.unit_price * (1 - (p.discount / 100)), 2)
            
            if p.category.strip().lower() == 'service':
                servs_data.append({
                    "Name": p.name,
                    "Category": p.category,
                    "SkillSet": p.skillset,
                    "Price L1": p.unit_price,
                    "Price L2": p.price_l2,
                    "Price L3": p.price_l3,
                    "Discount (%)": p.discount,
                    "Total (₹)": total_price,
                    "Supplier": p.supplier,
                    "Contact Info": p.supplier_contact,
                    "Reference": p.reference,
                    "Last Updated": p.last_updated,
                    "_UID": p.uid
                })
            else:
                prods_data.append({
                    "Name": p.name,
                    "Category": p.category,
                    "Make": p.make,
                    "Model": p.model,
                    "Specification": p.specification,
                    "Price L1": p.unit_price,
                    "Discount (%)": p.discount,
                    "Total (₹)": total_price,
                    "Supplier": p.supplier,
                    "Contact Info": p.supplier_contact,
                    "Reference": p.reference,
                    "Last Updated": p.last_updated,
                    "_UID": p.uid
                })
                
        with pd.ExcelWriter(self.filepath, engine='openpyxl') as writer:
            pd.DataFrame(prods_data).to_excel(writer, sheet_name='Products', index=False)
            pd.DataFrame(servs_data).to_excel(writer, sheet_name='Services', index=False)