# calculations.py
from decimal import Decimal, ROUND_HALF_UP

def money(v):
    """Round value to 2 decimal places safely."""
    return float(Decimal(v).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def calculate_totals(items):
    """
    items: list of dicts with keys:
      description, uom, qty, rate, tax (percentage)
    Returns dict with processed items and summary totals.
    """
    processed_items = []
    subtotal = 0.0
    total_tax = 0.0
    slno = 1

    for it in items:
        qty = float(it.get("qty", 0) or 0)
        rate = float(it.get("rate", 0) or 0)
        tax_percent = float(it.get("tax", 0) or 0)

        # Product total BEFORE tax
        product_total = qty * rate

        # Tax amount for the line
        tax_amount = product_total * (tax_percent / 100.0)

        # Final line total including tax
        total_amount = product_total + tax_amount

        processed_items.append({
            "slno": slno,
            "description": it.get("description", ""),
            "uom": it.get("uom", ""),
            "qty": qty,
            "rate": money(rate),
            "product_total": money(product_total),
            "tax": money(tax_percent),     # store percentage for template display
            "tax_amount": money(tax_amount),
            "total": money(total_amount)
        })

        slno += 1
        subtotal += product_total
        total_tax += tax_amount

    subtotal = money(subtotal)
    total_tax = money(total_tax)
    grand_total = money(subtotal + total_tax)

    # rounding difference (fractional part) — you may change how you want rounding shown
    rounding = money(grand_total - float(int(grand_total)))

    return {
        "items": processed_items,
        "subtotal": subtotal,
        "total_tax": total_tax,
        "grand_total": grand_total,
        "rounding": rounding
    }
