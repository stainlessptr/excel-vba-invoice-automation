# Excel VBA Invoice Automation

A portfolio case study of an Excel VBA workflow for preparing Japanese-language billing/request documents from a customer list and a worksheet template.

> **Status:** Source-code showcase only. The original workbook is intentionally excluded because it may contain confidential business and personal information. This repository is **not yet a runnable standalone demo**. The code has not been fully tested against a sanitized workbook.

## What it does

- Adds a new entry to a `list` worksheet with input validations and action buttons.
- Calculates line-item totals from up to six amount fields and applies tax-related formulas.
- Copies a `template` worksheet to generate a customer-specific document.
- Fills in dates, customer information, line items, totals and withholding-related fields.
- Adds a return link to the source entry and opens print preview.
- Includes helpers for handling merged cells and normalizing Japanese-formatted monetary input.

## Source

`src/InvoiceAutomation.bas` contains the provided VBA module, with only a standard module-name declaration added for import. The original logic is preserved for review; it is **not claimed to be bug-free**.

## Requirements and setup

- Microsoft Excel desktop with VBA macro support.
- A workbook containing sheets named `list` and `template`, with the exact cell layout expected by the code.
- Import the `.bas` module through the VBA editor (`Alt+F11` → File → Import File), or review it as plain text.
- Enable macros only for workbooks you trust.

The original workbook is not distributed. A runnable demonstration will require a new, fictional-data template and testing.

## Known issues / improvement backlog

1. Validate generated sheet names (Excel's 31-character limit and forbidden characters).
2. Avoid deleting an existing generated sheet before confirming that replacement succeeds.
3. Make numeric input errors explicit rather than silently treating invalid amounts as zero.
4. Reconcile and test the list total, document total, consumption tax and withholding formulas against confirmed business rules.
5. Add `Option Explicit`, declare variables explicitly and remove duplicate button assignments.
6. Add error handling and regression tests for blank inputs, duplicates, negative amounts and merged cells.

## Portfolio use

This project demonstrates business-process automation, Excel VBA event-driven forms, template-based document generation, input validation and financial-document workflows. It does **not** represent a certified accounting or tax calculation product.

## Privacy and ownership

No original `.xlsm` workbook or real customer records are included. Before publishing any screenshots or workbook, replace all personal, client and financial information and check hidden sheets, metadata, comments, formulas, macros and embedded objects. Publish only material you are authorized to share.

## Next steps

See [`docs/DEMO_PLAN.md`](docs/DEMO_PLAN.md) for a plan to create a safe runnable demo.
