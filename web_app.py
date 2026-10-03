import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import sys
import os

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))
from cleaner import DataCleaner
from reporter import ExcelReportGenerator

st.set_page_config(
    page_title="DataCleaner Pro | Enterprise Automation",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ DataCleaner Pro — Enterprise Data Cleaning & Reporting")
st.markdown("Automated ingestion, deduplication, and executive multi-tab Excel reporting pipeline.")

uploaded_file = st.file_uploader("Upload raw/messy Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # Save to temp file
    suffix = Path(uploaded_file.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        tmp_path = tmp_file.name

    st.info(f"File loaded: **{uploaded_file.name}** ({round(len(uploaded_file.getvalue()) / 1024, 2)} KB)")

    if st.button("🚀 Run Automation Pipeline", type="primary"):
        with st.spinner("Processing dataset, running deduplication, and styling executive report..."):
            try:
                cleaner = DataCleaner(tmp_path)
                cleaned_df, audit = cleaner.clean()

                # Generate Excel
                out_path = Path(tempfile.gettempdir()) / f"{Path(uploaded_file.name).stem}_Cleaned_Report.xlsx"
                reporter = ExcelReportGenerator(cleaned_df, audit, output_path=out_path)
                reporter.save()

                st.success("✅ Cleaning & Reporting Pipeline Completed Successfully!")

                # Display Metrics
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Raw Rows", f"{audit['raw_rows']:,}")
                col2.metric("Duplicates Removed", f"{audit['duplicates_removed']:,}")
                col3.metric("Missing Values Treated", f"{audit['missing_values_before']} -> {audit['missing_values_after']}")
                col4.metric("Execution Time", f"{audit['processing_time_sec']}s")

                st.subheader("Preview of Cleaned Data")
                st.dataframe(cleaned_df.head(25), use_container_width=True)

                # Download Button for the Generated Excel Workbook
                with open(out_path, "rb") as f:
                    excel_bytes = f.read()

                st.download_button(
                    label="📥 Download Executive Styled Excel Report (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{Path(uploaded_file.name).stem}_Executive_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error processing file: {e}")
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
else:
    st.write("👆 Upload a messy sales or operational spreadsheet above to test the automated pipeline.")
