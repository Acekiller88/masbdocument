# MIMOS Purchase Order — Final Visual QA

**Target DOCX:** `PO Template/MIMOS_Academy_Blank_Purchase_Order_Template.docx`  
**Visual source of truth:** `PO Template/contoh po.PDF` (6 pages, A4 portrait, 595 × 842 pt)  
**Existing V1 DOCX** was inspected and **not** copied. It uses MIMOS Academy branding, Malay placeholders, and rewritten acknowledgement text. The PDF always wins.

**Logo decision:** `MIMOS_Academy_LOGO.png` is the Academy wordmark on black and **does not match** the PDF (MIMOS Solutions). The template embeds the logo extracted from the PDF (`MIMOS_Solutions_LOGO_from_PO.png`, PDF xref 9, 1000×372, placed at 120 × 44.65 pt).

**Font substitutes (original Type1 families are not embeddable as named fonts):**

| PDF (CONFIRMED) | Word substitute |
|---|---|
| Helvetica / Helvetica-Bold / Helvetica-Oblique | Arial / Arial Bold / Arial Italic |
| Courier / Courier-Bold | Courier New / Courier New Bold |

Purple fill measured from the 150 dpi reference render: banner `#801070` (RGB 128,16,112); TOTAL bar `#8F1B7E` (RGB 143,27,126).

QA PNGs were produced by walking the DOCX table geometry (exact row heights and column widths written into the file). They are not Microsoft Word’s layout engine; Word with Arial + Courier New will be closer than these PNGs (which fall back to DejaVu).

**Visual QA iterations performed: 4**

---

## PAGE 1 — Main Purchase Order (`1/02`)

### Differences found (iteration 1)
- Nested vendor table doubled inner borders; extra empty paragraphs opened a gap under the supplier block.
- Description notes wrapped inside the Description column (PDF prints them across Quantity/Unit/Per because SAP does not clip the text run).
- `borders=False` did not actually clear table borders (`edges=()` was treated as “all edges”).

### Corrections made
- Supplier block rebuilt as a 3-column grid with a merged address row (Vendor No / Tel. No / Fax No share the measured widths 70.85 / 99.20 / 113.40 pt).
- Nested item/rest tables use interior grid lines only; parent cell draws the outer box once.
- Empty body paragraphs between tables removed so measured row heights control vertical rhythm.
- Table-border helper fixed so borderless spacers, header chrome, acknowledgement and signature stay borderless.
- TOTAL bar uses measured purple shading and white Arial Bold; “Continue Next Page” remains in the Total Price cell.
- Page number is the visible reference value `1/02` (not reinterpreted).

### Remaining differences
- Item notes wrap inside Description instead of overflowing the column (usable Word behaviour; PDF overflow is a SAP print artefact).
- Unit and Per still show a short empty band beside TOTAL (PDF continues those columns through the TOTAL row without a break).
- 4 pt italic “(formerly known as …)” sits slightly tighter under the logo than the 4 pt Helvetica-Oblique overlay.
- Courier New metrics vs Courier: character width is equivalent (10 CPI) but glyph shapes differ (this QA render uses DejaVu Sans Mono, which slashes zeros).

---

## PAGE 2 — Continuation + totals (`2/02`)

### Differences found
- Same header/footer chrome as page 1 (correct). First build put note 3 only in Description; totals sat in the Unit Price / Total Price columns.

### Corrections made
- Page number `2/02`.
- Note 3 wording copied exactly from the PDF (wrapped in Description).
- Sub Total `5,000.00` / Service Tax `400.00` / Sales Tax `0.00` / TOTAL `5,400.00` in the measured columns.
- Same acknowledgement + signature block as page 1.

### Remaining differences
- Same note-wrapping and TOTAL-row unit/per band as page 1.
- Subtotal labels sit in the Unit Price column (as in the PDF) but Word vertical-align bottom is an approximation of the SAP y=496.8 placement.

---

## PAGE 3 — Terms & Conditions (clauses 1–13 start)

### Differences found (iteration 1)
- Clause titles rendered as `INVOICE1.` because PDF number and title spans sit 0.6 pt apart on y and were concatenated in the wrong x-order.

### Corrections made
- Spans on the same line are sorted by x before writing, so titles are `1. INVOICE`, `2. TAX`, … matching the PDF.
- Body lines (including doubled spaces used by SAP as poor-man’s justification) copied from PDF spans.
- Title `TERMS & CONDITIONS OF PURCHASE ORDER` is Arial Bold 12 pt, centred.
- Clause 7 QUANTITY body remains bold, as in the PDF.

### Remaining differences
- Line-wrap points follow the PDF’s stored line breaks, so they match; Courier New vs Courier may reflow by a character if the user edits a line.
- Title-to-body gap is ~16 pt vs measured 29.6 pt from title top to first body top (12 pt title + 16 pt after ≈ 28 pt — close).

---

## PAGE 4 — Terms continuation (warranty remainder through 20(f))

### Differences found
- None structural after the x-order fix. Wording matches the PDF, including `LEINS` (not corrected to LIENS) and `P.O.is`.

### Corrections made
- Hard page break at the PDF split (warranty first line stays on page 3; remainder starts page 4).

### Remaining differences
- Same font-substitute and micro-spacing notes as page 3.

---

## PAGE 5 — Terms continuation (20(f) remainder through 31 heading)

### Differences found
- None structural. Clause numbers `21.`–`31.` match the PDF (`21.LIQUIDATED…` has a space before LAD in the PDF source).

### Remaining differences
- Same font-substitute notes.

---

## PAGE 6 — Terms close (31 body, 32, 33)

### Differences found (iteration 1)
- Headings appeared as `BRIBERY 32.` until x-order sort.

### Corrections made
- `32.BRIBERY` and `33.BID-RIGGING/ANTI-COMPETITION` match the PDF (no space after the number on these lines).
- Large trailing white space on page 6 is preserved (no extra clauses invented).

### Remaining differences
- Same font-substitute notes.

---

## Variable fields represented (28)

Sample values from the reference PO are kept so the template still *looks* like the PDF. Cells are ordinary editable Word text (no “ENTER NAME HERE” placeholders).

1. Purchase Order No. — `7000000306`  
2. Purchase Order Release No. — `N/A`  
3. Order Date — `25.05.2026`  
4. Contract No. — `N/A`  
5. Revision No. — `N/A`  
6. Page No. — `1/02` / `2/02`  
7. Supplier name / address / SST  
8. Attn To (supplier) — `Mohamad Fauzi`  
9. Vendor No. — `20009041`  
10. Tel. No. — `60389955000`  
11. Fax No. — (blank)  
12. Deliver To  
13. Expected Delivery — `24.06.2026`  
14. Completion Work — `24.06.2026`  
15. Attn to (Requestor) — `Saszwani`  
16. Tel. or Ext. no. — `-`  
17. Delivery Terms  
18. Payment Terms  
19. Purchase Requisition (PR) No. — `6000000481`  
20. Buyers Ref / Tel. or Ext No.  
21. Refer to RFQ / ESS / SebutHarga / Tender No. — `34512`  
22. Item No. / Description / Quantity / Unit / Per  
23. Unit Price / Total Price  
24. Currency — `MYR`  
25. Sub Total  
26. Service Tax  
27. Sales Tax / TOTAL  
28. Checked by Buyer / Approved By  

Buyer-side identity on the form (MIMOS SOLUTIONS SDN. BHD., SST, address, tel, URL) is treated as letterhead, not a 29th variable.

---

## Overall fidelity

**GOOD**

Justified: six A4 portrait pages; all 33 terms clauses with PDF wording (including original spelling); six-column item table with measured (not equal) widths; purple banners and TOTAL bar; `1/02` and `2/02` page numbers; Solutions logo; acknowledgement and signature lines. Not EXCELLENT because Word cannot reproduce SAP’s unclipped description overflow, Type1 Courier/Helvetica, or the raster banner bitmaps pixel-for-pixel.

---

## Files

| File | Role |
|---|---|
| `PO Template/MIMOS_Academy_Blank_Purchase_Order_Template.docx` | Final template |
| `PO Template/06_DOCX_RENDER/page-1.png` … `page-6.png` | DOCX geometry render |
| `PO Template/01_REFERENCE_RENDER/page-1.png` … `page-6.png` | PDF reference render |
| `PO Template/MIMOS_Solutions_LOGO_from_PO.png` | Logo extracted from the PDF |
| `PO Template/build_po_template.py` | Rebuild script |
