import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Excel to CSV Converter", layout="wide")

# Add centered logo
col1, col2, col3 = st.columns([1,2,1])
with col2:
    st.image("logo.png", width=150)

st.title("📂 Excel to CSV Converter")
st.write("Upload one or more Excel files and convert them into CSV format (same filename, only format change).")

# File uploader (multiple files allowed)
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED)

    multiple_files = len(uploaded_files) > 1

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")
        
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names

            st.write(f"Sheets found: {', '.join(sheet_names)}")

            # Combine all sheets into ONE dataframe (no extra columns)
            combined_df = []

            for sheet in sheet_names:
                df = excel_data.parse(sheet)
                combined_df.append(df)

                # Preview
                with st.expander(f"Preview: {sheet}"):
                    st.dataframe(df)

            final_df = pd.concat(combined_df, ignore_index=True)

            # Convert to CSV
            csv_buffer = io.StringIO()
            final_df.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue()

            # SAME filename, only extension changed
            filename = uploaded_file.name.rsplit('.', 1)[0] + ".csv"

            if multiple_files:
                zip_file.writestr(filename, csv_data)
            else:
                st.download_button(
                    label="⬇️ Download CSV",
                    data=csv_data,
                    file_name=filename,
                    mime="text/csv"
                )

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
