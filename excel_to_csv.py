import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel to CSV Converter", layout="wide")

st.title("📂 Excel to CSV Converter")
st.write("Upload one or more Excel files and convert them into CSV format.")

# File uploader (multiple files allowed)
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")
        
        try:
            # Read Excel file
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names

            st.write(f"Sheets found: {', '.join(sheet_names)}")

            for sheet in sheet_names:
                df = excel_data.parse(sheet)

                # Convert dataframe to CSV
                csv_buffer = io.StringIO()
                df.to_csv(csv_buffer, index=False)

                csv_data = csv_buffer.getvalue()

                # Download button
                st.download_button(
                    label=f"⬇️ Download CSV ({sheet})",
                    data=csv_data,
                    file_name=f"{uploaded_file.name.replace('.xlsx','').replace('.xls','')}_{sheet}.csv",
                    mime="text/csv"
                )

                # Optional preview
                with st.expander(f"Preview: {sheet}"):
                    st.dataframe(df)

        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")