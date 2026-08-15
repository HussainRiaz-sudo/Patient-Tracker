# 🩺 Dr Naila Patient Tracker & Financial Ledger

A modern, responsive, offline-first web application designed for medical practitioners to manage patient intake across multiple clinical locations, automate location-specific revenue split calculations, archive records, generate itemized monthly PDF settlement reports, and analyze clinical trends.

---

## 🚀 Key Features

### 🏥 1. Multi-Hospital Switcher & Isolated Storage
- **Instant Hospital Toggle:** Seamlessly switch between multiple clinical practices with a single tap:
  - **🏥 Naeem Surgical:** Configured for **Doctor Share 30% / Hospital Share 70%**.
  - **🏥 Cavalry Hospital:** Configured for **Doctor Share 70% / Hospital Share 30%**.
- **Completely Isolated Databases:** Each hospital operates on its own isolated data keys (`doctor_naeem_*` vs `doctor_cavalry_*`), ensuring patient records, daily counts, and financial stats never mix.
- **Zero Server Overhead:** Built on browser LocalStorage (no Redis cache required), providing **100% free, offline, instant multi-tenant storage** capable of holding 50,000+ records.

### 📋 2. Daily Log (Sheet 1)
- **Real-Time Patient Counting:** Automatically generates and previews sequential serial numbers (`#1`, `#2`, `#3`...) as patient names are entered.
- **Smart Patient Autocomplete:** Dynamically suggests names of recurring patients to speed up intake on touchscreen devices.
- **Location-Aware Financial Totals:** Real-time summary cards displaying *Today's Patients*, *Today's Gross Billing (PKR)*, and *Doctor Share* calculated according to the active hospital's split rate.
- **Automatic Midnight Reset:** Automatically clears the daily intake log at the end of the day while preserving all entries in the permanent ledger.

### 🗄️ 3. All-Time Ledger & Filtered Summary (Sheet 2)
- **Permanent Archival Database:** Retains all historical patient entries across all days for the selected hospital.
- **Dynamic Revenue Split Calculation:** Computes revenue distribution instantly in **Pakistani Rupees (PKR)** based on active hospital split rules:
  - **Naeem Surgical:** 30% Doctor / 70% Hospital
  - **Cavalry Hospital:** 70% Doctor / 30% Hospital
- **Quick Month Selector:** Pick any specific month directly from the filter toolbar to auto-populate date boundaries and view monthly totals.
- **Dynamic Filtered Summary Banner:** Dedicated real-time summary block above the ledger table that calculates:
  - **Filtered Matching Patients**
  - **Filtered Gross Charges (PKR)**
  - **Filtered Doctor Share (PKR)** *(30% or 70%)*
  - **Filtered Hospital Share (PKR)** *(70% or 30%)*
- **Search & Multi-Filter Engine:** Filter entries instantly by patient name, month, or custom date ranges (*From Date / To Date*).
- **Data Export & Controls:**
  - **📊 CSV Backup Export:** One-click CSV download labeled with the active hospital's name and split rates.
  - **🗑️ Delete All Records:** Modal-protected reset option per hospital.

### 🖨️ 4. Itemized Monthly PDF Settlement Report Export
- **Official Hospital Proof of Earnings:** One-click **"Print Monthly PDF"** button on Sheet 2 that generates an official, itemized financial settlement statement formatted for hospital administration.
- **Complete Patient Intake Breakdown:**
  - **Patient Serial Numbers (`#1`, `#2`, `#3`...)**
  - **Patient Names**
  - **Consultation Date & Creation Timestamp** *(e.g. Jul 15, 2026 | 02:30 PM)*
  - **Gross Billing Charges (PKR)**
  - **Doctor Share (PKR)** *(30% or 70%)*
  - **Hospital Share (PKR)** *(70% or 30%)*
- **Summary Stat Cards & Verification Footer:**
  - Includes top summary cards for Total Volume, Gross Revenue, Doctor Payout, and Hospital Share.
  - Features official signature verification blocks for **Dr. Naila (Practitioner)** and **Hospital Management**.
- **Multi-Page Print Engine:** Styled with `@media print` CSS rules enabling seamless multi-page pagination with `page-break-inside: avoid` on table rows and signatures.

### 👤 5. Patient History & Clinical Record Modal
- **Retroactive Record Scanner:** Tapping any patient name on **Sheet 1 (Daily Log)** or **Sheet 2 (All-Time Ledger)** aggregates all past and present visits across both hospital practices (*Naeem Surgical* & *Cavalry Hospital*).
- **Comprehensive Patient Profile:** Displays total visit count, cumulative gross charges, total doctor earnings, and first/last visit timestamps.
- **Chronological Timeline Table:** Lists all historical consultations with date/time, hospital facility, gross charges, and net payout.

### 📊 6. Analytics & Trends (Sheet 3)
- **Interactive Visual Dashboard:** Powered by **Chart.js** with location-specific charting modes:
  - 📈 **Daily Patient Flow (Line Chart):** Tracks day-by-day patient volume for the selected month and hospital.
  - 📊 **Revenue & Split Breakdown (Stacked Bar Chart):** Visualizes daily earnings split dynamically formatted for the active hospital.
  - 📅 **Past 6-Month Volume Comparison:** Bar graph comparing total patient volume across the last 6 months.
- **Dynamic Month Selector:** View analytics for the current month or switch to any historical month archived in the database.

### 🌓 7. Dark Mode & Light Mode Theme Engine
- **1-Tap Header Toggle:** Sun/Moon toggle button in top bar header for seamless switching.
- **High-Contrast Text Guarantee:** Explicit CSS theme overrides ensuring inputs, select dropdowns, labels, and table cells remain 100% legible without dark-on-dark text overlap.
- **Persistent Theme Preference:** Remembers user preference in `LocalStorage` (`doctor_theme`) to prevent white-screen flashes on load.
- **Chart.js Theme Adaptation:** Dynamically updates graph gridlines and labels for dark/light themes.

### 🩺 8. Cavalry Hospital Procedure Priority Schedule & 100% Doctor Payout
- **Exclusive Hospital Visibility:** Tab 4 (**Procedures**) is scoped exclusively to **Cavalry Hospital** (`activeHospital === 'cavalry'`) and hides automatically when Naeem Surgical is active.
- **Smart Patient Autocomplete:** Allows choosing existing patients from intake records or entering new patients.
- **Earliest-First Priority Engine:** Automatically sorts procedures chronologically by due date so the most urgent procedures remain at the very top of the list.
- **Interactive Completion Fee Prompt:** Tapping "Mark Completed" opens a prompt for the doctor to enter the procedure fee (PKR).
- **100% Doctor Share (Zero Hospital Ratio Split):** Procedure earnings bypass the standard 70/30 hospital split and assign **100% of revenue to the Doctor Share** (`Doctor = 100%`, `Hospital = 0%`).
- **Separate All-Time Ledger Record:** Completed procedures automatically post to Sheet 2 tagged with a distinct purple `Procedure` badge.
- **Visual Priority Badges:**
  - 🔴 **Overdue / Due Today** (Red High Priority)
  - 🟠 **Due Soon (1-3 Days)** (Orange Medium Priority)
  - 🟢 **Scheduled** (Blue Standard Priority)
  - ⚪ **Mark Completed Toggle** (Green Done Status)

---

## 🛠️ Technology Stack

| Component | Technology |
|---|---|
| **Core Structure** | HTML5 (Semantic Markup) |
| **Styling & Theme** | Vanilla CSS3 (CSS Variables, Flexbox, CSS Grid, `@media print`) |
| **Logic & Engine** | Vanilla JavaScript (ES6+, LocalStorage API, Fetch API) |
| **Data Visualization** | Chart.js (CDN) |
| **Iconography & Favicon** | Lucide Icons (CDN), Custom SVG Favicon |
| **Deployment** | Vercel Static Hosting (`vercel.json`) |

---

## 🔒 Security & Privacy Audit

This repository contains **NO confidential credentials, secret keys, API tokens, or personal patient data**.

- **No Secrets Hardcoded:** The application uses client-side LocalStorage and user-configured environment endpoints.
- **Public Repository Safe:** You can safely fork, publicize, or deploy this repository without any risk of exposing sensitive keys.
- **Google Sheets Endpoint Privacy:** Google Apps Script Web App URLs are configured by the user via the in-app settings UI and stored securely in their own browser's LocalStorage.

---

## 💻 Getting Started

### Local Running (No Installation Required)
1. **Clone the repository:**
   ```bash
   git clone https://github.com/HussainRiaz-sudo/Patient-Tracker.git
   ```
2. **Launch the application:**
   - Double-click `index.html` to open it directly in any modern web browser (Chrome, Safari, Edge, Firefox).
   - Alternatively, serve it locally using VS Code Live Server or python HTTP server:
     ```bash
     python -m http.server 8000
     ```

---

## ⚙️ Google Sheets Background Sync Setup (Optional)

To sync records automatically to a Google Sheet:
1. Open the application and click the **Settings (gear icon)** at the top right.
2. Click **Copy Code** to copy the provided Google Apps Script snippet.
3. Open a new [Google Sheet](https://sheets.new), navigate to **Extensions > Apps Script**, paste the code, and click **Save**.
4. Click **Deploy > New Deployment**, choose **Web App**, set *Execute as: "Me"*, and *Who has access: "Anyone"*.
5. Copy the generated **Web App URL**, paste it into the app's settings panel, and click **Test & Save**.

---

## 🌐 Deploying to Vercel

This repository includes a pre-configured `vercel.json` for one-click static hosting:

1. Import the repository into your [Vercel Dashboard](https://vercel.com).
2. Keep all default build settings (Framework: *Other / Static*).
3. Click **Deploy**. Vercel will host the web application on a global CDN with clean URLs.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).
