import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO

from parser import (
    extract_text,
    extract_name,
    extract_email,
    extract_mobile,
    extract_highest_qualification_details,
    extract_latest_company,
    extract_total_experience
)

# ==================================================
# PAGE CONFIG
# ==================================================

st.set_page_config(
    page_title="Resume Extractor",
    layout="wide"
)

st.title("📄 Bulk Resume Extractor")
st.write("Upload PDF and DOCX resumes to extract candidate details.")

# ==================================================
# FILE UPLOADER
# ==================================================

uploaded_files = st.file_uploader(
    "Upload Resumes",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

# ==================================================
# PROCESS RESUMES
# ==================================================

if uploaded_files:

    st.info(f"{len(uploaded_files)} file(s) selected.")

    if st.button("🚀 Process Resumes"):

        records = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):

            status_text.text(
                f"Processing {idx + 1}/{len(uploaded_files)} : {file.name}"
            )

            temp_path = None

            try:

                file_ext = os.path.splitext(file.name)[1].lower()

                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=file_ext
                ) as tmp:

                    tmp.write(file.read())
                    temp_path = tmp.name

                text = extract_text(temp_path)

                qual = extract_highest_qualification_details(text)

                records.append({
                    "Resume": file.name,
                    "Name": extract_name(text),
                    "Email": extract_email(text),
                    "Mobile": extract_mobile(text),
                    "Qualification": qual["qualification"],
                    "Passing Year": qual["passing_year"],
                    "Latest Company": extract_latest_company(text),
                    "Total Experience": extract_total_experience(text),
                    "Status": "Success"
                })

            except Exception as e:

                records.append({
                    "Resume": file.name,
                    "Name": "Error",
                    "Email": "Error",
                    "Mobile": "Error",
                    "Qualification": "Error",
                    "Passing Year": "Error",
                    "Latest Company": "Error",
                    "Total Experience": "Error",
                    "Status": str(e)
                })

            finally:

                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except:
                        pass

            progress_bar.progress((idx + 1) / len(uploaded_files))

        # ==========================================
        # RESULTS
        # ==========================================

        df = pd.DataFrame(records)

        success_count = len(df[df["Status"] == "Success"])
        failed_count = len(df[df["Status"] != "Success"])

        st.success(
            f"Completed | Success: {success_count} | Failed: {failed_count}"
        )

        st.dataframe(
            df,
            use_container_width=True,
            height=500
        )

        # ==========================================
        # CSV DOWNLOAD
        # ==========================================

        csv_data = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="📥 Download CSV",
            data=csv_data,
            file_name="resume_output.csv",
            mime="text/csv"
        )

        # ==========================================
        # EXCEL DOWNLOAD
        # ==========================================

        excel_buffer = BytesIO()

        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Resume Data"
            )

        excel_buffer.seek(0)

        st.download_button(
            label="📊 Download Excel",
            data=excel_buffer,
            file_name="resume_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ==========================================
        # FAILED FILES
        # ==========================================

        failed_df = df[df["Status"] != "Success"]

        if not failed_df.empty:

            st.subheader("❌ Failed Files")

            st.dataframe(
                failed_df[["Resume", "Status"]],
                use_container_width=True
            )