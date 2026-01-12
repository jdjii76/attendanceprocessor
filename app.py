import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# --- HELPER FUNCTIONS ---
def clean_name(value) -> str | None:
    if pd.isna(value): return None
    s = str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s or None

def format_workbook_in_memory(output_bytes):
    output_bytes.seek(0)
    wb = load_workbook(output_bytes)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"
        for col in ws.columns:
            col_letter = col[0].column_letter
            max_len = 0
            for cell in col[:200]:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)

    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue()


# --- MAIN APP ---
st.set_page_config(page_title="Attendance Processor", layout="centered")
st.title("📊 Attendance Processor")
st.info("Upload your MS Forms or Google Forms exports (.xlsx) to generate a summary.")

uploaded_files = st.file_uploader("Upload Attendance Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    if st.button("GENERATE COMPLETE REPORT"):
        try:
            all_dfs = []
            file_reports = []

            for f in uploaded_files:
                temp_df = pd.read_excel(f)
                # Universal mapping
                mapping = {"Timestamp": "Start time", "Email Address": "Email", "Full Name": "Name"}
                temp_df = temp_df.rename(columns=mapping)
                all_dfs.append(temp_df)

            df = pd.concat(all_dfs, ignore_index=True)
            df["SubmissionDateTime_ET"] = pd.to_datetime(df["Start time"])
            df["ClassDate"] = df["SubmissionDateTime_ET"].dt.normalize()
            df["Email"] = df["Email"].astype(str).str.strip().str.lower()
            
            if "Name" in df.columns:
                df["DisplayName"] = df["Name"].apply(clean_name)
            else:
                df["DisplayName"] = "Unknown"

            # Logic processing
            df_sorted = df.sort_values(["ClassDate", "Email", "SubmissionDateTime_ET"])
            daily = df_sorted.drop_duplicates(subset=["ClassDate", "Email"], keep="first")
            
            name_mode = daily.groupby("Email")["DisplayName"].agg(
                lambda s: s.value_counts().idxmax() if not s.dropna().empty else "Unknown"
            )

            student_summary = daily.groupby("Email").agg(
                DaysPresent=("ClassDate", "nunique"), 
                LastSeen=("ClassDate", "max")
            ).reset_index()
            
            total_days = daily["ClassDate"].nunique()
            student_summary["DisplayName"] = student_summary["Email"].map(name_mode)
            student_summary["TotalClassDays"] = total_days
            student_summary["AttendancePercent"] = (student_summary["DaysPresent"] / total_days * 100).round(1)

            most_recent_by_student = daily.sort_values("ClassDate").drop_duplicates(subset="Email", keep="last")
            most_recent_by_student["DisplayName"] = most_recent_by_student["Email"].map(name_mode)

            status_report = most_recent_by_student.merge(
                student_summary[["Email", "DaysPresent", "TotalClassDays", "AttendancePercent"]],
                on="Email", how="left"
            ).sort_values("ClassDate", ascending=False)

            per_day = daily.groupby("ClassDate").agg(PresentCount=("Email", "nunique")).reset_index()

            # Create Excel object
            excel_data = BytesIO()
            with pd.ExcelWriter(excel_data, engine="openpyxl") as writer:
                student_summary.to_excel(writer, sheet_name="Student_Summary", index=False)
                daily.to_excel(writer, sheet_name="Daily_Attendance", index=False)
                per_day.to_excel(writer, sheet_name="Per_Day_Counts", index=False)
                most_recent_by_student.to_excel(writer, sheet_name="Most_Recent_By_Student", index=False)
                status_report.to_excel(writer, sheet_name="Student_Status_Report", index=False)

            formatted_excel = format_workbook_in_memory(excel_data)
            st.success("✅ Report Generated!")
            st.download_button(
                label="📥 Download Excel Report",
                data=formatted_excel,
                file_name="Attendance_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")