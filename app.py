import io
import re
import pdfplumber
import pandas as pd
import streamlit as st


def extract_records_from_pdf(file) -> pd.DataFrame:
    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

    records = []
    for line in text.split("\n"):
        match = re.search(r"┃\s*(.*?)\s*│\s*(.*?)\s*┃", line)
        if match:
            records.append({
                "住所": match.group(1).strip(),
                "所有者名": match.group(2).strip()
            })

    return pd.DataFrame(records)


def main():
    st.set_page_config(page_title="PDF Data Extractor", layout="wide")
    st.title("PDF Data Extractor")

    uploaded_file = st.file_uploader(
        "Upload a PDF file", type=["pdf"], help="Only PDF files are supported"
    )

    if uploaded_file:
        df = extract_records_from_pdf(uploaded_file)

        if not df.empty:
            st.subheader("Extracted Data")
            st.dataframe(df)

            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            towrite.seek(0)

            st.download_button(
                label="Download as Excel",
                data=towrite,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No matching rows found in the uploaded PDF.")
    else:
        st.info("Please upload a PDF to begin extraction.")


if __name__ == "__main__":
    main()