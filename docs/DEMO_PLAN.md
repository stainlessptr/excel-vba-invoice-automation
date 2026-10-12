# Runnable demo checklist

This package intentionally does not contain the original business workbook.

1. Create a **new** `.xlsm` workbook with fictional company names and amounts; do not copy confidential records into the demo.
2. Recreate the `list` and `template` sheet structure, including expected merged cells and formulas.
3. Import `src/InvoiceAutomation.bas` and test `AddNewRow`, `GenerateRequestSheet` and `PrintRequestSheet` in desktop Excel.
4. Verify the tax and withholding rules with appropriate domain expertise before making accuracy claims.
5. Test duplicate customer/date combinations, long or invalid sheet names, blank inputs and error recovery.
6. Add sanitized screenshots of list view, generated document and print preview to `screenshots/`.
7. Only after all checks, add the fictional-data `.xlsm` under `demo/`.
8. Check that you have permission to publish any work originally developed for a third party.

**Do not** rename the original private workbook and upload it as a demo: renaming does not anonymize it.
