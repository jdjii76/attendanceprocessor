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
        # Style Header
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"
        # Auto-size columns
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
st.markdown("Works with **Microsoft Forms** and **Google Forms**.")

uploaded_files = st.file_uploader("Upload Attendance Files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    if st.button("GENERATE COMPLETE REPORT"):
        try:
            all_dfs = []
            for f in uploaded_files:
                raw_df = pd.read_excel(f)
                
                # 1. FIND TIME COLUMN (MS: Start time, Google: Timestamp)
                time_alias = ["Start time", "Timestamp", "Submission Date"]
                time_col = next((c for c in time_alias if c in raw_df.columns), None)
                
                # 2. FIND EMAIL COLUMN (MS: Email, Google: Email Address)
                email_alias = ["Email", "Email Address"]
                email_col = next((c for c in email_alias if c in raw_df.columns), None)
                
                if not time_col or not email_col:
                    st.warning(f"Skipping {f.name}: Missing Time or Email column.")
                    continue

                # 3. CONSOLIDATE NAMES (Check all possible name columns)
                # This merges Name, Name1, Name2, etc. into one column
                name_aliases = ["Name", "Name1", "Name2", "Full Name", "Student Name"]
                found_name_cols = [c for c in name_aliases if c in raw_df.columns]
                
                # Coalesce: take the first non-null value across all name candidates
                if found_name_cols:
                    display_name = raw_df[found_name_cols].bfill(axis=1).iloc[:, 0].apply(clean_name)
                else:
                    display_name = "Unknown"

                # 4. CREATE CLEAN TEMP DATAFRAME (Purges all other columns)
                temp_df = pd.DataFrame({
                    "RawTime": pd.to_datetime(raw_df[time_col]),
                    "Email": raw_df[email_col].astype(str).str.strip().str.lower(),
                    "DisplayName": display_name
                })
                all_dfs.append(temp_df)

            if not all_dfs:
                st.error("No valid data found in uploaded files.")
                st.stop()

            df = pd.concat(all_dfs, ignore_index=True)
            df["ClassDate"] = df["RawTime"].dt.normalize()
            
            # --- PROCESSING ---
            # 1. Daily Attendance
            daily = df.sort_values(["ClassDate", "Email", "RawTime"]).drop_duplicates(subset=["ClassDate", "Email"], keep="first")
            
            # Pick the most frequent name for each email (Standardization)
            name_map = daily.groupby("Email")["DisplayName"].agg(lambda s: s.value_counts().idxmax() if not s.empty else "Unknown")

            # 2. Student Summary
            total_days = daily["ClassDate"].nunique()
            student_summary = daily.groupby("Email").agg(
                DaysPresent=("ClassDate", "nunique"), 
                LastSeen=("ClassDate", "max")
            ).reset_index()
            student_summary["DisplayName"] = student_summary["Email"].map(name_map)
            student_summary["TotalClassDays"] = total_days
            student_summary["AttendancePercent"] = (student_summary["DaysPresent"] / total_days * 100).round(1)
            student_summary = student_summary[["DisplayName", "Email", "DaysPresent", "TotalClassDays", "AttendancePercent", "LastSeen"]]

            # 3. Status Report
            status_report = daily.sort_values("ClassDate").drop_duplicates(subset="Email", keep="last")
            status_report = status_report.merge(student_summary[["Email", "DaysPresent", "AttendancePercent"]], on="Email", how="left")
            status_report = status_report[["DisplayName", "Email", "ClassDate", "AttendancePercent", "DaysPresent"]]
            status_report.rename(columns={"ClassDate": "LastSeenDate"}, inplace=True)
            status_report = status_report.sort_values("LastSeenDate", ascending=False)

            # 4. Daily Log
            daily_log = daily[["ClassDate", "DisplayName", "Email", "RawTime"]].rename(columns={"RawTime": "CheckInTime"})

            # 5. Counts
            per_day = daily.groupby("ClassDate").agg(PresentCount=("Email", "nunique")).reset_index()

            # --- EXPORT ---
            excel_data = BytesIO()
            with pd.ExcelWriter(excel_data, engine="openpyxl") as writer:
                student_summary.to_excel(writer, sheet_name="Student_Summary", index=False)
                daily_log.to_excel(writer, sheet_name="Daily_Attendance", index=False)
                per_day.to_excel(writer, sheet_name="Per_Day_Counts", index=False)
                status_report.to_excel(writer, sheet_name="Student_Status_Report", index=False)

            formatted_excel = format_workbook_in_memory(excel_data)
            st.success("✅ Perfectly Clean Report Generated!")
            st.download_button(
                label="📥 Download Excel Report",
                data=formatted_excel,
                file_name="Attendance_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
