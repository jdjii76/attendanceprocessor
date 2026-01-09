# attendanceprocessor
# 📊 Attendance Processor (Multi-Form)

A web-based tool built with **Streamlit** to combine multiple Excel attendance exports from **Microsoft Forms** or **Google Forms** into a single, formatted summary.

## 🚀 How to Use

1. **Prepare your files:** - All files must be in `.xlsx` format.
   - **Google Forms users:** Download your data as a `.csv`, open it in Excel, and "Save As" an `.xlsx` workbook before uploading.
2. **Access the App:** Open the Streamlit link (provided after deployment).
3. **Upload:** Drag and drop all attendance files for the semester/session into the uploader.
4. **Generate:** Click **"GENERATE COMPLETE REPORT"**.
5. **Download:** Click the download button to receive your formatted `Attendance_Summary.xlsx`.

## 📋 Features

This tool generates an Excel workbook with five specialized sheets:
* **Student_Summary:** Overall statistics (Total Days Present, Attendance %, and Last Seen date).
* **Daily_Attendance:** A cleaned log showing the first valid check-in per student, per day.
* **Per_Day_Counts:** Unique student headcounts for every class date.
* **Most_Recent_By_Student:** A dedicated list of the last time every individual student was seen.
* **Student_Status_Report:** A "Who is at risk?" view, sorted by students who haven't attended in the longest time.

## 🛠 Required Column Headers

The script automatically detects headers from both major platforms:
- **Time:** Looks for `Start time` (MS) or `Timestamp` (Google).
- **Email:** Looks for `Email` (MS) or `Email Address` (Google).
- **Name:** Looks for `Name` (MS) or `Full Name` (Google).

## 🧰 Tech Stack

- **Language:** Python
- **Interface:** Streamlit
- **Data Engine:** Pandas
- **Excel Formatting:** Openpyxl
