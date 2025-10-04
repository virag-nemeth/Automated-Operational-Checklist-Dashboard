# 🧾 Automated Operational Checklist & Performance Dashboard

An advanced Excel-based automation system designed to **streamline operational compliance tracking** and **automate routine check resets**, while generating **real-time performance dashboards** — all without manual input.

> 🧠 **Note:** This project was developed as part of a **client engagement**.  
> All proprietary details, company identifiers, and client-specific data have been **safely removed**.  
> The shared version focuses solely on the **technical structure, logic, and automation framework**.

---

## ⚙️ Key Features

### 🔄 Automated Reset Engine
- Automatically resets checklist items on set schedules:
  - **Daily checks** (3.5h, 8h, 24h intervals)
  - **Weekly checks** (every Monday, 6:00 AM)
  - **Monthly checks** (first working day of each month)
  - **6+ week checks** (6, 12, and 24-week intervals)
- Includes a **manual “Run All Resets”** button as a safety override.
- Tracks all reset timestamps in a dedicated **Reset Tracker** sheet.

---

### ⏰ Background Timer Check
- Runs an **automatic validation every 15 minutes** while the file is open.
- Detects and executes missed resets automatically.
- Logs the **last refresh timestamp** in the Reset Tracker for audit visibility.

---

### 📊 Dynamic Completion Dashboard
- **Real-time dashboard** powered by Power Query and Pivot Charts.
- Displays completion percentages for:
  - Daily, Weekly, Monthly, and Long-Term (6, 12, 24 weeks)
- Interactive **slicers** to filter by date, week, quarter, month, or year.
- Includes a **“Missed Checks” panel** showing incomplete items.
- Trends tracked across time to monitor operational performance.

---

### ✅ Smart Checklist System
- Each checkbox click automatically logs:
  - Timestamp  
  - User Initials  
  - Check type (Daily, Weekly, Monthly, or 6 Weeks+)  
- Timestamps and logs are stored on a **central Log Sheet** for traceability.  
- VBA ensures accurate **completion rate calculations**.  
- Adding new checks or rows is supported without breaking the automation.

---

### 🎨 Professional BI-Style Dashboard Design
- Clean, modern layout inspired by Power BI visuals.  
- Uses **traffic-light indicators** (red/yellow/green) for instant insights.  
- Includes a **custom “Refresh Data” button** for one-click updates.  
- Dark-blue dashboard theme for clear data contrast and readability.

---

## 🧠 Tech Stack
- **Microsoft Excel (.xlsm)**  
- **VBA Macros** – time-based automation and event triggers  
- **Power Query** – data transformation and aggregation  
- **Pivot Tables & Charts** – dynamic visualization  
- **Conditional Formatting** – visual KPI indicators  

---

## 📂 Main Components

| Module | Purpose |
|--------|----------|
| `modDailyReset` | Automates daily checklist resets (3.5h, 8h, 24h intervals). |
| `modWeeklyReset` | Resets weekly checks every Monday 6:00 AM. |
| `modMonthlyReset` | Resets monthly checks on the first working day of the month. |
| `mod6WeeksReset` | Resets long-term checks (6, 12, 24-week intervals). |
| `modScheduledReset` | Runs timed reset checks every 15 minutes while open. |
| `modManualRunner` | Manual “Run All Resets” button to trigger all macros. |
| `modDashboardRefresh` | Refreshes Power Queries and updates dashboard visuals. |

---

## 🗂️ Sheet Overview

| Sheet | Function |
|--------|-----------|
| **Daily checks** | Operational checklist resetting multiple times per day. |
| **Weekly checks** | Weekly tasks resetting automatically each Monday. |
| **Monthly checks** | Tasks that reset on the first working day of the month. |
| **6 weeks+ checks** | Long-term checks grouped by 6, 12, and 24-week cycles. |
| **ResetTracker** | Stores all reset timestamps and auto-refresh logs. |
| **Log sheet** | Central audit log of all completed tasks with timestamps. |
| **Completion Rate** | Interactive dashboard showing KPIs and performance trends. |

---

## 🚀 How It Works

1. **Automatic startup:**  
   When the workbook opens, all reset checks are executed immediately.

2. **Ongoing monitoring:**  
   A timer runs every 15 minutes to verify and trigger missed resets.

3. **Data refresh:**  
   Power Queries pull data from the log sheet to update the dashboard.

4. **Dashboard insights:**  
   KPIs update automatically, showing completion rates, trends, and missed checks.

---

## 💡 Ideal For

Teams managing **operational, compliance, or safety checklists** that need:
- Automated reset logic  
- Audit-ready logging  
- Real-time performance tracking  
- Professional, BI-style Excel dashboards  

---

## 📅 Example Use Cases
- Manufacturing daily/weekly QA inspections  
- Facilities and maintenance compliance logs  
- Financial reconciliation task tracking  
- Periodic safety or procedural reviews  

---

## 🧩 Future Enhancements
- Email or Teams notifications for missed checks  
- Power BI integration  
- Role-based user tracking for multi-department reporting  

---

## 📘 License
This project is for **educational and portfolio purposes only**.  
Free to adapt and customize for non-commercial applications.

---

### ✨ Author
Developed by **Virág Németh**  
Originally built for a private client project and shared with identifying data removed.  
Focused on **automating repetitive reporting and compliance tasks in Excel**.

