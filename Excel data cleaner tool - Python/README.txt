================================================================================
  EXCEL DATA CLEANER TOOL
================================================================================

A lightweight, production-ready Python utility for cleaning messy Excel
datasets. Automates the most common data-cleaning chores so you can go
from raw input.xlsx to a tidy cleaned_output.xlsx in one command.


--------------------------------------------------------------------------------
  FEATURES
--------------------------------------------------------------------------------

  * Removes completely empty rows (NaN, blank, or whitespace-only).
  * Removes duplicate rows (keeps the first occurrence).
  * Standardizes column names:
        - strips surrounding whitespace
        - converts to lowercase
        - replaces spaces / special characters with underscores
        - collapses repeated underscores
        (e.g. "  First Name  "  ->  "first_name"
              "E-mail Address"  ->  "e_mail_address"
              "Total ($)"       ->  "total")
  * Saves the cleaned data to a new Excel file (original is never modified).
  * Clean logging at every step so you can see what was removed.
  * Friendly error handling (missing file, unsupported format, file locked
    by Excel, etc.).
  * Configurable input / output paths via command-line arguments.


--------------------------------------------------------------------------------
  REQUIREMENTS
--------------------------------------------------------------------------------

  * Python 3.8 or newer
  * pandas
  * openpyxl   (pandas uses this under the hood to read/write .xlsx)

  Install dependencies:

      pip install pandas openpyxl


--------------------------------------------------------------------------------
  FOLDER STRUCTURE
--------------------------------------------------------------------------------

  Excel_Cleaner_Tool/
  |-- cleaner.py             <- the tool
  |-- input.xlsx             <- your messy input file (sample provided)
  |-- cleaned_output.xlsx    <- cleaned result (created after running)
  |-- README.txt             <- this file


--------------------------------------------------------------------------------
  USAGE -- STEP BY STEP
--------------------------------------------------------------------------------

  1. Place your messy Excel file in the project folder and name it
     "input.xlsx" (or use a custom name with the --input flag).

  2. Open a terminal in the project folder.

  3. (One-time) Install dependencies:

         pip install pandas openpyxl

  4. Run the cleaner with default paths:

         python cleaner.py

     Or with custom paths:

         python cleaner.py --input my_data.xlsx --output cleaned.xlsx
         python cleaner.py -i my_data.xlsx -o cleaned.xlsx

  5. Open "cleaned_output.xlsx" to see the cleaned data.


--------------------------------------------------------------------------------
  EXAMPLE
--------------------------------------------------------------------------------

  Input (input.xlsx)  -- 8 rows, messy column names, empty + duplicate rows:

      '  First Name  ', 'LAST NAME', 'E-mail Address', 'Total ($)', 'Signup Date'
      Alice, Smith,    a@x.com, 100, 2024-01-05
      Bob,   Jones,    b@x.com, 250, 2024-02-10
      Alice, Smith,    a@x.com, 100, 2024-01-05     <- duplicate
      (empty row)
      Charlie, Brown,  c@x.com,  75, 2024-03-15
      (whitespace-only row)
      Diana, Patel,    d@x.com, 500, 2024-04-20
      Bob,   Jones,    b@x.com, 250, 2024-02-10    <- duplicate

  Command:

      python cleaner.py

  Log output:

      Loading file: input.xlsx
      Loaded 8 rows x 5 columns
      Standardized 5 column names
      Removed 2 empty rows
      Removed 2 duplicate rows
      Saved cleaned file: cleaned_output.xlsx (4 rows)
      Cleaning complete.

  Output (cleaned_output.xlsx) -- 4 rows, clean column names:

      first_name, last_name, e_mail_address, total, signup_date
      Alice,      Smith,     a@x.com,        100,   2024-01-05
      Bob,        Jones,     b@x.com,        250,   2024-02-10
      Charlie,    Brown,     c@x.com,         75,   2024-03-15
      Diana,      Patel,     d@x.com,        500,   2024-04-20


--------------------------------------------------------------------------------
  TROUBLESHOOTING
--------------------------------------------------------------------------------

  * "Input file not found"
      -> Make sure input.xlsx is in the same folder as cleaner.py,
         or pass a full path with --input.

  * "Cannot write to cleaned_output.xlsx. Is the file open in Excel?"
      -> Close the output file in Excel, then re-run.

  * "ModuleNotFoundError: No module named 'pandas'" (or 'openpyxl')
      -> Run:  pip install pandas openpyxl


--------------------------------------------------------------------------------
  EXIT CODES
--------------------------------------------------------------------------------

  0  -> Success
  1  -> Expected error (missing file, bad format, file locked, etc.)
  2  -> Unexpected internal error


================================================================================
