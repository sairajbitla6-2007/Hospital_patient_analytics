HOSPITAL PATIENT ANALYTICS — PLOTLY FILES

Files
- Hospital_Patient_Analytics_Plotly_Dashboard.html: opens in a browser and includes the workbook's summary data. It needs an internet connection to load the pinned Plotly.js library from cdn.plot.ly.
- Code.gs and Index.html: Google Apps Script web-app source. The app reads the Dashboard and Chart Data tabs from a Google Sheet.

Google Sheets setup
1. Upload hospital_patient_analytics_dash_1.xlsx to Google Drive and open it with Google Sheets. Confirm that the Dashboard and Chart Data tabs are present.
2. In that Google Sheet, choose Extensions > Apps Script.
3. Replace Code.gs with the supplied Code.gs. Add an HTML file named Index and paste the supplied Index.html.
4. Save. Select setupDashboard in the function menu and click Run. Google will ask you to authorize the script to read this spreadsheet. Run setupDashboard again if needed.
5. Choose Deploy > New deployment > Web app. For initial testing, restrict access to yourself. For intended sharing, choose only the approved Google accounts or organization. Execute as you only if authorized viewers should view the returned aggregate summaries using your sheet access. Avoid anonymous/public access.
6. Deploy and open the web-app URL. When updating code later, use Deploy > Manage deployments > Edit, create a new version, and redeploy.

What is interactive now
- Plotly hover labels, zoom, pan, selection, legend toggles, responsive sizing, and image export.
- A from/to date control filters daily records and the selected-date record KPI.

Data limitation
The provided workbook contains summaries, not patient-level records. Thus the date filter cannot recalculate disease, severity, insurance, patient-type, or cost summaries across the selected range. Those metrics remain June 2026 totals. Add an approved cleaned source-data tab with the required fields before implementing cross-filters. Do not use real identifiable patient records unless your organization approves the transfer and app access model.
