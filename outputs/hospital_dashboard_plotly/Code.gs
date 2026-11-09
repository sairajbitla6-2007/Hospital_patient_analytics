const DASHBOARD_ID_KEY = 'PATIENT_ANALYTICS_SPREADSHEET_ID';

// Run once from the bound Google Sheet to remember its ID and authorize read access.
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the target Google Sheet and run setupDashboard again.');
  if (!ss.getSheetByName('Chart Data') || !ss.getSheetByName('Dashboard')) {
    throw new Error('This spreadsheet must contain Dashboard and Chart Data sheets.');
  }
  PropertiesService.getScriptProperties().setProperty(DASHBOARD_ID_KEY, ss.getId());
  return 'Dashboard source saved: ' + ss.getName();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hospital Patient Analytics');
}

function getDashboardData() {
  const id = PropertiesService.getScriptProperties().getProperty(DASHBOARD_ID_KEY);
  if (!id) throw new Error('Run setupDashboard once from the Google Sheet first.');
  const ss = SpreadsheetApp.openById(id);
  const c = ss.getSheetByName('Chart Data');
  const d = ss.getSheetByName('Dashboard');
  if (!c || !d) throw new Error('Expected Dashboard and Chart Data sheets.');
  const pair = (a1) => c.getRange(a1).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({label: String(r[0]), value: Number(r[1])}));
  const daily = c.getRange('A2:B31').getValues().filter(r => r[0] !== '' && r[0] !== null).map(r => {
    const label = String(r[0]);
    const match = label.match(/^(\d{1,2})\s+Jun$/i);
    const day = match ? Number(match[1]) : 1;
    return {label, date: '2026-06-' + String(day).padStart(2, '0'), records: Number(r[1])};
  });
  const ratingByType = c.getRange('AJ2:AL6').getValues().filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({rating:Number(r[0]),opd:Number(r[1]),ipd:Number(r[2])}));
  const ratingBox = c.getRange('AN2:AT3').getValues().filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({type:String(r[0]),n:Number(r[1]),min:Number(r[2]),q1:Number(r[3]),median:Number(r[4]),q3:Number(r[5]),max:Number(r[6])}));
  const paymentLabels = c.getRange('AI2:AI6').getValues().flat();
  const paymentValues = c.getRange('AH2:AH6').getValues().flat();
  return {
    period:{start:'2026-06-01',end:'2026-06-30',label:'1–30 June 2026'},
    kpis:{totalRecords:Number(d.getRange('A8').getValue()),opd:Number(d.getRange('C8').getValue()),ipd:Number(d.getRange('E8').getValue()),totalCost:Number(d.getRange('G8').getValue()),avgRating:Number(d.getRange('I8').getValue()),insuredAvg:Number(d.getRange('K8').getValue()),uninsuredAvg:Number(d.getRange('M8').getValue()),ipdMaxDaily:Number(d.getRange('E10').getValue()),opdMaxDaily:Number(d.getRange('L10').getValue())},
    daily, patientType:pair('C2:D3'), diseases:pair('E2:F11'), costPerPatientDay:pair('J2:K3'), serviceTime:pair('L2:M3'), outcomes:pair('N2:O6'), severity:pair('P2:Q5'), ratingDistribution:pair('S2:T6'), paymentRecords:pair('V2:W6'), insuranceRecords:pair('Y2:Z3'), costByRating:pair('AB2:AC6'), emergencyByInsurance:pair('AE2:AF3'),
    paymentCost:paymentLabels.map((label,i)=>({label:String(label),value:Number(paymentValues[i])})),
    ratingByType, ratingBox
  };
}
