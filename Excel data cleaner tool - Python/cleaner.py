"""
Pipeline:
    1. Load Excel file
    2. Standardize column names
    3. Remove empty rows (fully blank / whitespace-only)
    4. Remove duplicate rows
    5. Save cleaned output to a new Excel file

Usage:
    python cleaner.py
    python cleaner.py -i input.xlsx -o cleaned_output.xlsx

Requires: Python 3.8+, pandas, openpyxl
"""

from __future__ import annotations
import argparse
import logging
import re
import sys
from pathlib import Path
import pandas as pd

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("excel_cleaner")

# ---------------------------------------------------------------------------
# Core cleaning functions
# ---------------------------------------------------------------------------
def load_excel(file_path: str | Path) -> pd.DataFrame:
    """
    Load an Excel file into a pandas DataFrame.

    Raises:
        FileNotFoundError: if the file does not exist.
        ValueError:        if the file extension is not supported.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    if path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        raise ValueError(
            f"Unsupported file format '{path.suffix}'. "
            "Expected .xlsx, .xls, or .xlsm."
        )

    logger.info("Loading file: %s", path)
    df = pd.read_excel(path)
    logger.info("Loaded %d rows x %d columns", len(df), len(df.columns))
    return df

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names:
        - Convert to string
        - Strip surrounding whitespace
        - Lowercase
        - Replace any non-alphanumeric sequence with a single underscore
        - Trim leading/trailing underscores

    Example:
        "  First Name  "   -> "first_name"
        "E-mail Address"   -> "e_mail_address"
        "Total ($)"        -> "total"
    """
    def _clean(name: object) -> str:
        text = str(name).strip().lower()
        text = re.sub(r"[^\w]+", "_", text)   # any run of non-word chars -> "_"
        text = re.sub(r"_+", "_", text)       # collapse repeats
        return text.strip("_")

    df = df.copy()
    df.columns = [_clean(c) for c in df.columns]
    logger.info("Standardized %d column names", len(df.columns))
    return df

def remove_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove rows that are entirely empty.

    A row is considered empty if every cell is either:
        - NaN / None, or
        - An empty / whitespace-only string.
    """
    before = len(df)

    # Treat empty / whitespace-only strings as missing, then drop rows that
    # are all missing. This handles pure-NaN rows, pure-whitespace rows,
    # and mixed NaN+whitespace rows uniformly.
    cleaned = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.loc[cleaned.notna().any(axis=1)].reset_index(drop=True)

    logger.info("Removed %d empty rows", before - len(df))
    return df

def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Remove duplicate rows (keeping the first occurrence)."""
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    logger.info("Removed %d duplicate rows", before - len(df))
    return df

def save_excel(df: pd.DataFrame, output_path: str | Path) -> None:
    """
    Save DataFrame to an Excel file.

    Raises:
        PermissionError: if the file is locked (e.g. open in Excel).
    """
    path = Path(output_path)
    try:
        df.to_excel(path, index=False)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot write to '{path}'. Is the file open in Excel? "
            "Close it and try again."
        ) from exc

    logger.info("Saved cleaned file: %s (%d rows)", path, len(df))

# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------
def clean_excel(input_path: str | Path, output_path: str | Path) -> pd.DataFrame:
    """
    Run the full cleaning pipeline on an Excel file and save the result.

    Returns:
        The cleaned DataFrame.
    """
    df = load_excel(input_path)

    if df.empty:
        logger.warning("Input file has no data — writing empty output.")
        save_excel(df, output_path)
        return df

    df = standardize_columns(df)
    df = remove_empty_rows(df)
    df = remove_duplicates(df)
    save_excel(df, output_path)

    logger.info("Cleaning complete.")
    return df

# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Clean an Excel file: remove empty rows, duplicates, "
                    "and standardize column names.",
    )
    parser.add_argument(
        "-i", "--input",
        default="input.xlsx",
        help="Path to the input Excel file (default: input.xlsx)",
    )
    parser.add_argument(
        "-o", "--output",
        default="cleaned_output.xlsx",
        help="Path for the cleaned Excel file (default: cleaned_output.xlsx)",
    )
    return parser.parse_args(argv)

def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        clean_excel(args.input, args.output)
    except (FileNotFoundError, ValueError, PermissionError) as exc:
        logger.error("%s", exc)
        return 1
    except Exception as exc:  # pragma: no cover
        logger.exception("Unexpected error: %s", exc)
        return 2
    return 0

if __name__ == "__main__":
    sys.exit(main())
