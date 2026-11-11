HOSPITAL PATIENT ANALYTICS — PLOTLY DASHBOARD

Files
- Hospital_Patient_Analytics_Plotly_Dashboard.html: standalone browser dashboard with the June summary and grouped patient-level chart aggregates embedded. It needs an internet connection for Plotly.js from cdn.plot.ly.
- Code.gs and Index.html: Google Apps Script web-app source. The app reads Dashboard and Chart Data summary tabs and aggregates the Patient Data tab on the server.

Google Sheets setup
1. Open the Google Sheet already connected to the deployed web app. Copy the Patient Data tab from the updated XLSX into this same spreadsheet. If you prefer, import clean_patients_cleaned.csv there as a new sheet named exactly Patient Data. Keep the first row of column headers unchanged.
2. In that Google Sheet, choose Extensions > Apps Script. Replace Code.gs with the supplied Code.gs. Add an HTML file named Index and paste the supplied Index.html.
3. Save. Select setupDashboard in the function menu and click Run. This stores the active spreadsheet as the dashboard source. Google will ask you to authorize the script to read it.
4. Choose Deploy > Manage deployments > Edit for the existing web app, select New version, and deploy. Restrict access to yourself, approved Google accounts, or your organization. Avoid anonymous/public access to patient-level source data.
5. Open the existing web-app URL and confirm the new Age distribution, Animated scatter, Association map, and Blood group & severity tabs show data.

Dashboard sections
- Daily records, Top diseases, Emergency & insurance, Cost by rating, Cost per patient-day, and Service time each have their own tab.
- Related existing charts are paired in Patient mix & outcomes, Severity & payment cost, Patient ratings, and Payment & coverage.
- New patient-level views compare age bands by gender, animate patient-type/severity counts by day, map categorical association strengths, and compare blood group/severity counts.

Data and interpretation notes
- The report covers 1–30 June 2026. The daily date selector updates only daily records and the selected-date total; the other existing summary metrics remain June totals.
- New patient-level charts filter admission dates to June 2026. Patient Data must use the cleaned CSV headers, including admission_date, patient_type, severity_level, gender, age, blood_group, and disease.
- The animated scatter uses one bubble per patient-type/severity pair per day; bubble size represents record count. The blood-group/severity scatter also uses aggregated bubbles to avoid misleading jitter for categories.
- The association map uses Cramér’s V for categorical pairs and the correlation ratio (η) for age-to-category pairs. These are association measures, not evidence of causation. Disease is a recorded label, not a verified diagnosis.
- Patient-level rows are aggregated in Apps Script; individual patient and doctor identifiers are not returned to the browser. Keep the Google Sheet and deployment restricted to approved viewers.
- Currency is INR. Cost per patient-day and waiting times retain the workbook’s definitions.
