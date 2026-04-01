import streamlit as st
import pandas as pd
import io
import zipfile
import openpyxl

st.set_page_config(page_title="Excel to CSV Converter", layout="wide")

st.title("📂 Excel to CSV Converter")
st.write("Upload one or more Excel files and convert them into CSV format (ZIP for multiple files).")

# File uploader (multiple files allowed)
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    # If multiple files → prepare ZIP
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED)

    multiple_files = len(uploaded_files) > 1

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")
        
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names

            st.write(f"Sheets found: {', '.join(sheet_names)}")

            for sheet in sheet_names:
                df = excel_data.parse(sheet)

                # Convert dataframe to CSV
                csv_buffer = io.StringIO()
                df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()

                filename = f"{uploaded_file.name.rsplit('.',1)[0]}_{sheet}.csv"

                if multiple_files:
                    # Add to ZIP
                    zip_file.writestr(filename, csv_data)
                else:
                    # Single file → direct download
                    st.download_button(
                        label=f"⬇️ Download CSV ({sheet})",
                        data=csv_data,
                        file_name=filename,
                        mime="text/csv"
                    )

                # Preview
                with st.expander(f"Preview: {sheet}"):
                    st.dataframe(df)

        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    if multiple_files:
        zip_file.close()
        zip_buffer.seek(0)

        st.download_button(
            label="⬇️ Download All as ZIP",
            data=zip_buffer,
            file_name="converted_csv_files.zip",
            mime="application/zip"
        )
