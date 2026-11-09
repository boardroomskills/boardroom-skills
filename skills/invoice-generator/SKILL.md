---
name: invoice-generator
description: Generate a professional invoice as a PDF (and optional Word .docx) from JSON or CSV line items, with logo placeholder, seller and buyer blocks, quantities, line and invoice-level discounts, VAT/sales-tax grouped by rate, shipping, payments received, balance due, due date, and USD/EUR/GBP formatting. Includes three themes and an automatic arithmetic QA that fails on any total mismatch. Use when a freelancer, agency or small business asks to create, draft, make or send an invoice, bill a client, or turn a spreadsheet of work into invoices. Not for tax advice, e-invoicing formats (Peppol, ZUGFeRD) or credit-note law.
---

# Invoice Generator

Produces `<invoice-number>.pdf` (A4, reportlab, one page for typical invoices, header repeats on long tables), optional `.docx` (editable, python-docx), and `<name>.totals.json` (all totals, for bookkeeping or checks). Money is computed with `Decimal` and ROUND_HALF_UP, never floats.

## Workflow

1. **Collect inputs. Never invent them.** Ask for anything missing:
   - Seller: legal name, address, tax/VAT ID (if they charge tax), email, bank or payment details.
   - Buyer: name, address, tax ID (needed for B2B VAT in most countries), PO number if any.
   - Invoice number (ask for their sequence; if they have none propose `INV-YYYY-0001` and say it is a proposal), issue date, payment terms or due date.
   - Line items: description, quantity, unit, unit price, tax rate per line, any line discount.
   - Currency, any invoice-level discount, shipping, amounts already paid.
   If tax rate or tax ID is unknown, ask. Do not guess a VAT rate. If they cannot say, set `tax_rate` 0 and tell them clearly that no tax is charged. Payment details are never made up: if absent, leave the `payment` object empty and say the invoice has no way to pay.
2. **Write the JSON** following `examples/invoice.json`. Line items may instead come from a CSV (`--items items.csv`, columns `description,details,quantity,unit,unit_price,discount_percent,tax_rate`; appended after `items` in the JSON). If the user hands you a spreadsheet, convert to this CSV first.
3. **Build:** `python scripts/build_invoice.py invoice.json out_dir [--items items.csv] [--name file-stem] [--theme classic|minimal|bold|theme.json] [--docx] [--no-pdf]`. Exit code 2 means a QA error; never hand over that file.
4. **QA** (below), fix, rebuild. Hand over the files and list every assumption (tax rates applied, due date derived from terms, anything left blank).

## JSON fields

`currency` (USD, EUR, GBP; other codes print as a prefix), `number_style` (`en` = 1,234.50, `eu` = 1.234,50 with symbol after), `theme`, `default_tax_rate`, `tax_label` (VAT, GST, Sales tax), `tax_note` (e.g. "Reverse charge: VAT to be accounted for by the recipient"; use with 0% lines and only if the user states it applies), `invoice{number, issue_date (ISO), due_days or due_date, po_number, notes}`, `seller{name, address[], tax_id, email, phone, logo_text}`, `buyer{name, contact, address[], tax_id}`, `items[]`, `discount{label, percent | amount}` (applied pro rata over lines so VAT per rate stays right), `shipping{label, amount, tax_rate}`, `payment{method, bank, account_name, iban, bic, account_number, sort_code, account_ref, link, reference, terms}`, `amount_paid`.

## Logo

No images are bundled. The header draws a rounded monogram from `seller.logo_text` (default: initials). Tell the user to replace it with their logo in the .docx, or add an image to the PDF script if they supply a file. Do not fabricate a logo.

## Conventions by currency

- USD: `$1,234.50`; set `tax_label` to "Sales tax" if relevant. EUR: `eu` style is usual in DE/FR/ES/IT (`1.234,50 €`); `en` style (`€1,234.50`) in IE/NL-English contexts. GBP: `£1,234.50`, show VAT number.
- Dates print as `02 Nov 2026` to avoid DD/MM ambiguity.

## Design system

One typeface (Liberation Sans in the PDF, Calibri in docx), a single primary colour and accent, hairline rules, generous whitespace. Themes in `templates/themes.json`: `classic` (navy and teal, filled header), `minimal` (black rules, no fills, prints well in mono), `bold` (red fills). Add a theme by copying a block and passing `--theme path.json` (a file with `primary, accent, text, muted, line, band, header_text, style` where `style` is `bar` or `rule`). Keep contrast high: `header_text` must be readable on `primary`.

## QA checklist (mandatory)

1. Read the script output: every `QA ERROR` must be zero. It independently re-derives: qty x price - discount = line amount; sum of lines = subtotal; subtotal - discount + shipping = net; tax bases sum to net; tax per rate recomputed; net + tax = total; total - paid = balance; due >= issue date.
2. Read `QA WARNING`s (missing tax ID while charging tax, placeholders, no payment details, negative balance) and resolve or report each.
3. Render and look: `pdftoppm -png -r 60 -singlefile out.pdf /tmp/pv`, open the PNG. Check: nothing clipped, long descriptions wrap, totals right-aligned and readable, balance due prominent, page count sensible. For docx: `soffice --headless --convert-to pdf` and check the same.
4. Mentally cross-check one line by hand and the grand total against the user's own figures. If the user supplied a total, compare it with `totals.json`; any difference is reported, not forced.
5. Remind the user: invoice numbers must be sequential and unique; legal invoice requirements vary by country; have an accountant confirm tax treatment.

## Limits

Prices are net (tax-exclusive). Tax-inclusive pricing, multiple currencies on one invoice, recurring schedules and e-invoicing XML are out of scope. Tell the user instead of faking them.
