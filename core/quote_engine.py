import logging
from datetime import datetime
from typing import Dict
from .models import Product, QuoteItem, ClientDetails, Quotation

logger = logging.getLogger(__name__)

class QuoteEngine:
    def __init__(self):
        # Dictionary to hold multiple carts keyed by Customer Name
        self.carts: Dict[str, Quotation] = {}

    def get_cart(self, customer_name: str, tax_rate: float = 0.10) -> Quotation:
        """Retrieves or creates a cart for the specific customer."""
        name_key = customer_name.strip() if customer_name.strip() else "Walk-in Customer"
        
        if name_key not in self.carts:
            quote_id = f"QT-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            client = ClientDetails(name=name_key, company="", email="")
            self.carts[name_key] = Quotation(quote_id=quote_id, client=client, tax_rate=tax_rate)
            
        return self.carts[name_key]

    def add_to_cart(self, customer_name: str, product: Product, quantity: int = 1):
        cart = self.get_cart(customer_name)
        # Check if product is already in cart to just increase qty
        for item in cart.items:
            if item.product.name == product.name:
                item.quantity += quantity
                return
        # Otherwise add new line item
        cart.items.append(QuoteItem(product=product, quantity=quantity))

    def remove_from_cart(self, customer_name: str, index: int):
        cart = self.get_cart(customer_name)
        if 0 <= index < len(cart.items):
            cart.items.pop(index)

    def clear_cart(self, customer_name: str):
        name_key = customer_name.strip() if customer_name.strip() else "Walk-in Customer"
        if name_key in self.carts:
            del self.carts[name_key]