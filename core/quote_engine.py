import logging
from datetime import datetime
from typing import Optional
from .models import Product, QuoteItem, ClientDetails, Quotation

logger = logging.getLogger(__name__)

class QuoteEngine:
    def __init__(self):
        self.current_quote: Optional[Quotation] = None

    def start_new_quote(self, client: ClientDetails, tax_rate: float = 0.0) -> str:
        # Format: QT-YYYYMMDD-HHMMSS
        quote_id = f"QT-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
        self.current_quote = Quotation(quote_id=quote_id, client=client, tax_rate=tax_rate)
        return quote_id

    def add_item(self, product: Product, quantity: int = 1, discount_percent: float = 0.0):
        if self.current_quote:
            self.current_quote.items.append(
                QuoteItem(product=product, quantity=quantity, discount_percent=discount_percent)
            )

    def remove_item(self, index: int):
        if self.current_quote and 0 <= index < len(self.current_quote.items):
            self.current_quote.items.pop(index)

    def get_summary(self) -> dict:
        if not self.current_quote: return {}
        return {
            "subtotal": self.current_quote.subtotal,
            "tax_amount": self.current_quote.tax_amount,
            "total": self.current_quote.total
        }