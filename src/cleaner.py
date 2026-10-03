import re
import time
import pandas as pd
import numpy as np
from pathlib import Path

class DataCleaner:
    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.raw_df = None
        self.cleaned_df = None
        self.audit_log = {
            "file_name": self.file_path.name,
            "file_size_kb": 0,
            "raw_rows": 0,
            "raw_columns": 0,
            "duplicates_removed": 0,
            "missing_values_before": 0,
            "missing_values_after": 0,
            "cleaned_columns": [],
            "processing_time_sec": 0.0,
            "notes": []
        }

    def load_data(self):
        """Loads data from CSV or Excel with fallback encodings."""
        if not self.file_path.exists():
            raise FileNotFoundError(f"File not found: {self.file_path}")

        self.audit_log["file_size_kb"] = round(self.file_path.stat().st_size / 1024, 2)
        suffix = self.file_path.suffix.lower()

        if suffix in ['.xlsx', '.xls']:
            self.raw_df = pd.read_excel(self.file_path)
        elif suffix == '.csv':
            try:
                self.raw_df = pd.read_csv(self.file_path, encoding='utf-8')
            except UnicodeDecodeError:
                self.raw_df = pd.read_csv(self.file_path, encoding='windows-1252')
        else:
            raise ValueError(f"Unsupported file format: {suffix}. Please provide .csv or .xlsx")

        self.audit_log["raw_rows"] = len(self.raw_df)
        self.audit_log["raw_columns"] = len(self.raw_df.columns)
        self.audit_log["missing_values_before"] = int(self.raw_df.isna().sum().sum())
        return self.raw_df

    def _clean_currency_numeric(self, series):
        """Cleans dirty currency strings, negative parenthesis, commas, and rogue text."""
        def parse_val(v):
            if pd.isna(v):
                return np.nan
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip()
            if s.lower() in ['n/a', 'na', 'null', 'none', '-', '--', 'pending', '?']:
                return np.nan
            # Check for parenthesized negatives: (123.45)
            is_negative = False
            if s.startswith('(') and s.endswith(')'):
                is_negative = True
                s = s[1:-1]
            # Strip currency symbols, commas, USD/EUR labels, and extra spaces
            s = re.sub(r'[^\d.-]', '', s)
            try:
                num = float(s)
                return -num if is_negative else num
            except ValueError:
                return np.nan
        return series.apply(parse_val)

    def clean(self):
        """Runs the full cleaning pipeline and tracks audit metrics."""
        start_time = time.time()
        if self.raw_df is None:
            self.load_data()

        df = self.raw_df.copy()

        # 1. Clean Column Headers (strip whitespace, normalize spaces)
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Clean Text Columns (strip rogue whitespace, standardize title-case)
        text_cols = df.select_dtypes(include=['object']).columns
        for col in text_cols:
            sample_non_null = df[col].dropna().astype(str).str.strip()
            if len(sample_non_null) == 0:
                continue

            currency_match = sample_non_null.str.contains(r'[$€£]|USD|EUR', regex=True).mean()
            numeric_match = sample_non_null.str.contains(r'^-?\$?\d+(?:[.,]\d+)?', regex=True).mean()

            if currency_match > 0.3 or numeric_match > 0.7:
                df[col] = self._clean_currency_numeric(df[col])
                median_val = df[col].median()
                df[col] = df[col].fillna(median_val if not np.isnan(median_val) else 0.0)
                self.audit_log["cleaned_columns"].append(f"{col} (Parsed Currency/Numeric, imputed missing with median)")
            else:
                # Attempt date parsing if contains 'date' in column name
                if 'date' in col.lower():
                    parsed_dates = pd.to_datetime(df[col], format='mixed', errors='coerce')
                    if parsed_dates.notna().mean() > 0.5:
                        df[col] = parsed_dates.ffill().bfill()
                        df[col] = df[col].dt.strftime('%Y-%m-%d')
                        self.audit_log["cleaned_columns"].append(f"{col} (Standardized to YYYY-MM-DD)")
                        continue

                # Standard text cleaning
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan': 'Unknown', 'None': 'Unknown', '': 'Unknown'})
                if col.lower() in ['segment', 'category', 'sub-category', 'ship mode', 'region', 'city', 'state']:
                    df[col] = df[col].str.title()
                self.audit_log["cleaned_columns"].append(f"{col} (Trimmed whitespace & standardized casing)")

        # 3. Deduplication (Post-normalization & Key-based)
        initial_count = len(df)
        # Check if composite business key exists
        if 'Order ID' in df.columns and 'Product ID' in df.columns:
            df = df.drop_duplicates(subset=['Order ID', 'Product ID'])
        else:
            # Check all columns except index/Row ID
            subset_cols = [c for c in df.columns if c.lower() not in ['row id', 'index', 'id']]
            df = df.drop_duplicates(subset=subset_cols if subset_cols else None)

        dups_removed = initial_count - len(df)
        self.audit_log["duplicates_removed"] = dups_removed
        if dups_removed > 0:
            self.audit_log["notes"].append(f"Deduplicated {dups_removed} records based on transaction keys.")

        # 4. Fill remaining missing values
        for col in df.columns:
            if df[col].isna().sum() > 0:
                if np.issubdtype(df[col].dtype, np.number):
                    df[col] = df[col].fillna(0)
                else:
                    df[col] = df[col].fillna("Unknown")

        # 5. Type coercion for numeric types if column names suggest metrics
        for col in df.columns:
            if col.lower() in ['sales', 'profit', 'discount', 'quantity']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        self.audit_log["missing_values_after"] = int(df.isna().sum().sum())
        self.audit_log["final_rows"] = len(df)
        self.audit_log["processing_time_sec"] = round(time.time() - start_time, 3)
        self.cleaned_df = df
        return self.cleaned_df, self.audit_log

if __name__ == "__main__":
    cleaner = DataCleaner("data/messy_sales_data.xlsx")
    cleaned_df, audit = cleaner.clean()
    print("Cleaning complete successfully!")
    print(f"Duplicates removed: {audit['duplicates_removed']}")
    print(f"Time taken: {audit['processing_time_sec']}s")
    print(f"Columns processed: {len(audit['cleaned_columns'])}")
    print(f"Raw rows: {audit['raw_rows']} -> Cleaned rows: {audit['final_rows']}")
