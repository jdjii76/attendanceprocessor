import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# --- HELPER FUNCTIONS ---
def clean_name(value) -> str:
    if pd.isna(value) or str(value).strip() == "": return ""
    s = str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s

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
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)
    
    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

# --- MAIN APP ---
st.set_page_config(page_title="Attendance Processor", layout="centered")
st.title("📊 Attendance Processor")

uploaded_files = st.file_uploader("Upload Attendance Files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    if st.button("GENERATE COMPLETE REPORT"):
        try:
            all_dfs = []
            for f in uploaded_files:
                raw_df = pd.read_excel(f)
                
                # Standardize Time and Email
                time_map = {"Timestamp": "Start time"}
                email_map = {"Email Address": "Email"}
                raw_df = raw_df.rename(columns={**time_map, **email_map})
                
                # Consolidate Names: Check Name, Name1, Name2, Full Name
                name_cols = ["Name", "Name1", "Name2", "Full Name"]
                # Create a single DisplayName by taking the first non-empty value found in name columns
                found_name_cols = [c for c in name_cols if c in raw_df.columns]
                
                if found_name_cols:
                    display_name = raw_df[found_name_cols].bfill(axis=1).iloc[:, 0].apply(clean_name)
                else:
                    display_name = "Unknown"

                # Create a clean dataframe with ONLY these 3 columns to prevent Name1/Name2 leak
                temp_df = pd.DataFrame({
                    "CheckInTime": pd.to_datetime(raw_df["Start time"]),
                    "Email": raw_df["Email"].astype(str).str.strip().str.lower(),
                    "DisplayName": display_name
                })
                all_dfs.append(temp_df)

            df = pd.concat(all_dfs, ignore_index=True)
            df["ClassDate"] = df["CheckInTime"].dt.normalize()
            
            # --- CALCULATIONS ---
            # 1. Deduplicate: First entry per student per day
            daily = df.sort_values(["ClassDate", "Email", "CheckInTime"]).drop_duplicates(subset=["ClassDate", "Email"], keep="first")
            
            # Stable Name Mapping (Mode Name)
            name_map = daily.groupby("Email")["DisplayName"].agg(lambda s: s.value_counts().idxmax() if not s.empty else "Unknown")

            # --- SHEET PREPARATION (Matching your example exactly) ---

            # Sheet 1: Student_Summary
            total_days = daily["ClassDate"].nunique()
            student_summary = daily.groupby("Email").agg(
                DaysPresent=("ClassDate", "nunique"), 
                LastSeen=("ClassDate", "max")
            ).reset_index()
            student_summary["DisplayName"] = student_summary["Email"].map(name_map)
            student_summary["TotalClassDays"] = total_days
            student_summary["AttendancePercent"] = (student_summary["DaysPresent"] / total_days * 100).round(1)
            # Match Order: DisplayName, Email, DaysPresent, TotalClassDays, AttendancePercent, LastSeen
            student_summary = student_summary[["DisplayName", "Email", "DaysPresent", "TotalClassDays", "AttendancePercent", "LastSeen"]]

            # Sheet 2: Daily_Attendance
            daily_attendance = daily[["ClassDate", "Email", "DisplayName", "CheckInTime"]].copy()
            daily_attendance.rename(columns={"CheckInTime": "SubmissionDateTime_ET"}, inplace=True)

            # Sheet 3: Per_Day_Counts
            per_day = daily.groupby("ClassDate").agg(PresentCount=("Email", "nunique")).reset_index()

            # Sheet 4: Most_Recent_By_Student
            most_recent = daily.sort_values("ClassDate").drop_duplicates(subset="Email", keep="last")
            most_recent = most_recent[["ClassDate", "Email", "DisplayName", "CheckInTime"]].copy()
            most_recent.rename(columns={"CheckInTime": "SubmissionDateTime_ET"}, inplace=True)

            # Sheet 5: Student_Status_Report
            status_report = most_recent.merge(
                student_summary[["Email", "DaysPresent", "TotalClassDays", "AttendancePercent"]],
                on="Email", how="left"
            ).sort_values("ClassDate", ascending=False)
            # Match Order: ClassDate, Email, DisplayName, SubmissionDateTime_ET, DaysPresent, TotalClassDays, AttendancePercent
            status_report = status_report[["ClassDate", "Email", "DisplayName", "SubmissionDateTime_ET", "DaysPresent", "TotalClassDays", "AttendancePercent"]]

            # --- EXCEL EXPORT ---
            excel_data = BytesIO()
            with pd.ExcelWriter(excel_data, engine="openpyxl") as writer:
                student_summary.to_excel(writer, sheet_name="Student_Summary", index=False)
                daily_attendance.to_excel(writer, sheet_name="Daily_Attendance", index=False)
                per_day.to_excel(writer, sheet_name="Per_Day_Counts", index=False)
                most_recent.to_excel(writer, sheet_name="Most_Recent_By_Student", index=False)
                status_report.to_excel(writer, sheet_name="Student_Status_Report", index=False)

            formatted_excel = format_workbook_in_memory(excel_data)
            st.success("✅ Clean Report Generated matching your template!")
            st.download_button(label="📥 Download Excel Report", data=formatted_excel, file_name="Attendance_Summary.xlsx")
            
        except Exception as e:
            st.error(f"Error: {e}")
