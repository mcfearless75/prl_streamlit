
import streamlit as st
import os
import tempfile
import pandas as pd
from timesheet_processor_streamlined import (
    process_folder,
    autofit_and_style,
    apply_excel_formula_to_pay_column,
    add_weekly_summary_sheet
)

st.set_page_config(page_title="PRL Timesheet Processor", layout="centered")

st.title("📄 PRL Timesheet Processor")
st.markdown("Upload `.docx` timesheets and receive a calculated Excel file with overtime breakdown and a weekly summary.")

uploaded_files = st.file_uploader("Upload Timesheet Files", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        for uploaded_file in uploaded_files:
            file_path = os.path.join(tmpdir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

        df = pd.DataFrame(process_folder(tmpdir))
        output_path = os.path.join(tmpdir, "All_Timesheets_Combined.xlsx")
        df.to_excel(output_path, sheet_name="Timesheets", index=False)

        autofit_and_style(output_path)
        apply_excel_formula_to_pay_column(output_path)
        add_weekly_summary_sheet(output_path)

        with open(output_path, "rb") as f:
            st.success("✅ Done! Download your Excel file below:")
            st.download_button("📥 Download Excel", f, file_name="All_Timesheets_Combined.xlsx")
