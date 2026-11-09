# Brief: invoice for Alder & Finch Ltd

Seller: Northwind Studio GmbH (fictional), Berlin, VAT DE000000000, bank transfer, IBAN placeholder.
Buyer: Alder & Finch Ltd, Dublin, accounts payable.
Invoice INV-2026-0042 issued 2 Nov 2026, net 30, PO-7781, currency EUR.

Lines:
1. Brand strategy workshop, 2 days at 1,450.00, 19% VAT
2. Visual identity system, 1 project at 6,800.00, 10% line discount, 19% VAT
3. Website copywriting, 12 pages at 185.50, 19% VAT
4. Printed brand books, 25 pcs at 18.40, 7% VAT

Invoice-level: 5% loyalty discount on goods and services, courier 42.00 (19% VAT). Nothing paid yet.
Output: examples/invoice.json -> output.pdf and output.docx (classic theme).

Extra examples: `invoice-gbp-minimal.json` (minimal theme; GBP, UK VAT 20%, bank transfer with sort code and account number) and `invoice-usd-bold-eu.json` (bold theme, EU number style, fixed-amount discount).
Commands: `python ../scripts/build_invoice.py invoice.json . --name output --docx`
