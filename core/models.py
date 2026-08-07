from dataclasses import dataclass, field
from typing import List
from datetime import datetime

@dataclass
class Product:
    name: str
    category: str
    description: str
    unit_price: float 
    price_l2: float = 0.0
    price_l3: float = 0.0
    supplier: str = ""
    supplier_contact: str = ""
    make: str = ""
    model: str = ""
    specification: str = ""
    skillset: str = ""
    discount: float = 0.0  # NEW: Base discount %
    reference: str = ""    # NEW: Blank reference field
    last_updated: str = field(default_factory=lambda: datetime.now().strftime("%d-%m-%y"))

@dataclass
class QuoteItem:
    product: Product
    quantity: int
    discount_percent: float = -1.0 
    service_level: int = 1  
    
    def __post_init__(self):
        # Automatically pull the product's default discount if not explicitly changed in cart
        if self.discount_percent == -1.0:
            self.discount_percent = self.product.discount

    @property
    def active_price(self) -> float:
        if self.product.category.strip().lower() == 'service':
            if self.service_level == 2:
                return self.product.price_l2
            elif self.service_level == 3:
                return self.product.price_l3
        return self.product.unit_price

    @property
    def subtotal(self) -> float:
        # Subtracts the discount percentage from the base total
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