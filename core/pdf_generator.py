import logging
from fpdf import FPDF
from .models import Quotation

logger = logging.getLogger(__name__)

class QuotePDF(FPDF):
    def header(self):
        self.set_font("helvetica", "B", 20)
        self.cell(0, 10, "QUOTATION", align="R", new_x="LMARGIN", new_y="NEXT")
        self.ln(10)

class PDFGenerator:
    @staticmethod
    def generate_quote_pdf(quote: Quotation, output_path: str):
        """Compiles the Quotation object into a formatted PDF."""
        logger.info(f"Generating PDF for quote: {quote.quote_id}")
        
        try:
            pdf = QuotePDF()
            pdf.add_page()
            
            # Client & Quote Info
            pdf.set_font("helvetica", size=10)
            pdf.cell(0, 5, f"Quote ID: {quote.quote_id}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 5, f"Date: {quote.date.strftime('%Y-%m-%d')}", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(5)
            
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 6, "Billed To:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("helvetica", size=10)
            pdf.cell(0, 5, f"{quote.client.name} | {quote.client.company}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 5, quote.client.address, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(10)
            
            # Table Header
            pdf.set_font("helvetica", "B", 10)
            col_widths = (30, 70, 20, 20, 30, 20) # Total 190mm
            headers = ["Item ID", "Description", "Qty", "Price", "Discount", "Total"]
            
            for width, header in zip(col_widths, headers):
                pdf.cell(width, 8, header, border=1)
            pdf.ln(8)
            
            # Table Body
            pdf.set_font("helvetica", size=10)
            for item in quote.items:
                pdf.cell(col_widths[0], 8, item.product.product_id, border=1)
                pdf.cell(col_widths[1], 8, item.product.name[:40], border=1)
                pdf.cell(col_widths[2], 8, str(item.quantity), border=1)
                pdf.cell(col_widths[3], 8, f"${item.product.unit_price:.2f}", border=1)
                pdf.cell(col_widths[4], 8, f"{item.discount_percent}%", border=1)
                pdf.cell(col_widths[5], 8, f"${item.subtotal:.2f}", border=1, new_x="LMARGIN", new_y="NEXT")
                
            pdf.ln(10)
            
            # Totals
            pdf.set_font("helvetica", "B", 10)
            pdf.cell(140) # Push to right
            pdf.cell(30, 8, "Subtotal:", border=1)
            pdf.cell(20, 8, f"${quote.subtotal:.2f}", border=1, new_x="LMARGIN", new_y="NEXT")
            
            pdf.cell(140)
            pdf.cell(30, 8, f"Tax ({(quote.tax_rate*100):.0f}%):", border=1)
            pdf.cell(20, 8, f"${quote.tax_amount:.2f}", border=1, new_x="LMARGIN", new_y="NEXT")
            
            pdf.cell(140)
            pdf.cell(30, 8, "Total:", border=1)
            pdf.cell(20, 8, f"${quote.total:.2f}", border=1)
            
            # Output
            pdf.output(output_path)
            logger.info(f"PDF successfully saved to {output_path}")
            
        except Exception as e:
            logger.error(f"Failed to generate PDF: {str(e)}", exc_info=True)
            raise