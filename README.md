# Excel Trial Balance Formatter

This project is an Excel VBA macro that formats a raw trial balance into a clean, readable layout.

## What it does

- Formats header row (bold, centered, shaded)
- Applies number formatting to Debit and Credit columns
- Adds borders around the trial balance
- Auto-fits column widths

## How to use

1. Download `TrialBalanceFormatter.xlsm`.
2. Open it in Excel and enable macros if prompted.
3. Paste or enter your trial balance starting in cell A1 with these columns:
   - Account Number
   - Account Name
   - Debit
   - Credit
4. Go to the **Developer** tab → click **Macros**.
5. Run `FormatTrialBalance`.

The macro will clean up the layout automatically.

## Why this is useful for accounting

Accountants often receive messy trial balances exported from ERP systems.
This macro standardizes the format in one click, saving time and reducing manual formatting.
