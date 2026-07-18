from dataclasses import dataclass, field
from typing import List, Optional
from datetime import datetime

@dataclass
class Product:
    name: str
    category: str
    description: str
    unit_price: float  # Base Price / Level 1
    price_l2: float = 0.0
    price_l3: float = 0.0
    supplier: str = ""
    supplier_contact: str = ""
    # New Fields for the updated layouts
    make: str = ""
    model: str = ""
    specification: str = ""
    skillset: str = ""

@dataclass
class QuoteItem:
    product: Product
    quantity: int
    discount_percent: float = 0.0
    service_level: int = 1  # Tracks which level is selected (1, 2, or 3)
    
    @property
    def active_price(self) -> float:
        """Returns the correct price based on the selected service level."""
        if self.product.category.strip().lower() == 'service':
            if self.service_level == 2:
                return self.product.price_l2
            elif self.service_level == 3:
                return self.product.price_l3
        return self.product.unit_price

    @property
    def subtotal(self) -> float:
        base_price = self.active_price * self.quantity
        discount_amount = base_price * (self.discount_percent / 100)
        return round(base_price - discount_amount, 2)

@dataclass
class ClientDetails:
    name: str
    company: str
    email: str

@dataclass
class Quotation:
    quote_id: str
    client: ClientDetails
    date: datetime = field(default_factory=datetime.now)
    items: List[QuoteItem] = field(default_factory=list)
    tax_rate: float = 0.0
    
    @property
    def subtotal(self) -> float:
        return sum(item.subtotal for item in self.items)
        
    @property
    def tax_amount(self) -> float:
        return round(self.subtotal * self.tax_rate, 2)
        
    @property
    def total(self) -> float:
        return round(self.subtotal + self.tax_amount, 2)