import logging
import pandas as pd
from datetime import datetime
from typing import List
from .models import Product

logger = logging.getLogger(__name__)

class CatalogManager:
    def __init__(self, filepath: str):
        self.filepath = filepath

    def load_catalog(self) -> List[Product]:
        try:
            df = pd.read_excel(self.filepath, engine='openpyxl')
            products = []
            for _, row in df.iterrows():
                
                # Safely parse the date, fallback to today if missing
                raw_date = row.get('Last Updated', '')
                if pd.isna(raw_date) or not str(raw_date).strip():
                    date_str = datetime.now().strftime("%Y-%m-%d")
                else:
                    date_str = str(raw_date).split()[0] # clean up time if present
                    
                products.append(Product(
                    name=str(row.get('Name', '')),
                    category=str(row.get('Category', '')),
                    description=str(row.get('Description', '')),
                    unit_price=float(row.get('Price L1', row.get('Unit Price', 0.0))),
                    price_l2=float(row.get('Price L2', 0.0) if pd.notna(row.get('Price L2')) else 0.0),
                    price_l3=float(row.get('Price L3', 0.0) if pd.notna(row.get('Price L3')) else 0.0),
                    supplier=str(row.get('Supplier', '')),
                    supplier_contact=str(row.get('Contact Info', '')),
                    make=str(row.get('Make', '')),
                    model=str(row.get('Model', '')),
                    specification=str(row.get('Specification', '')),
                    skillset=str(row.get('SkillSet', '')),
                    last_updated=date_str
                ))
            return products
        except FileNotFoundError:
            return []

    def save_catalog(self, products: List[Product]):
        data = [{
            "Name": p.name,
            "Category": p.category,
            "Description": p.description,
            "Price L1": p.unit_price,
            "Price L2": p.price_l2,
            "Price L3": p.price_l3,
            "Supplier": p.supplier,
            "Contact Info": p.supplier_contact,
            "Make": p.make,
            "Model": p.model,
            "Specification": p.specification,
            "SkillSet": p.skillset,
            "Last Updated": p.last_updated
        } for p in products]
        
        df = pd.DataFrame(data)
        df.to_excel(self.filepath, index=False, engine='openpyxl')