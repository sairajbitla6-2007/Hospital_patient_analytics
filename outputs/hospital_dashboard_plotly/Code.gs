const DASHBOARD_ID_KEY = 'PATIENT_ANALYTICS_SPREADSHEET_ID';
const PATIENT_SHEET_NAME = 'Patient Data';
const REPORT_START = '2026-06-01';
const REPORT_END = '2026-06-30';

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
  const patientAnalysis = getPatientAnalysis_(ss);
  return {
    period:{start:REPORT_START,end:REPORT_END,label:'1–30 June 2026'},
    kpis:{totalRecords:Number(d.getRange('A8').getValue()),opd:Number(d.getRange('C8').getValue()),ipd:Number(d.getRange('E8').getValue()),totalCost:Number(d.getRange('G8').getValue()),avgRating:Number(d.getRange('I8').getValue()),insuredAvg:Number(d.getRange('K8').getValue()),uninsuredAvg:Number(d.getRange('M8').getValue()),ipdMaxDaily:Number(d.getRange('E10').getValue()),opdMaxDaily:Number(d.getRange('L10').getValue())},
    daily, patientType:pair('C2:D3'), diseases:pair('E2:F11'), costPerPatientDay:pair('J2:K3'), serviceTime:pair('L2:M3'), outcomes:pair('N2:O6'), severity:pair('P2:Q5'), ratingDistribution:pair('S2:T6'), paymentRecords:pair('V2:W6'), insuranceRecords:pair('Y2:Z3'), costByRating:pair('AB2:AC6'), emergencyByInsurance:pair('AE2:AF3'),
    paymentCost:paymentLabels.map((label,i)=>({label:String(label),value:Number(paymentValues[i])})),
    ratingByType, ratingBox,
    ...patientAnalysis
  };
}

// Aggregate patient-level records on the server. Individual rows and identifiers are
// never sent to the browser; only grouped chart counts and association scores return.
function getPatientAnalysis_(ss) {
  const sheet = ss.getSheetByName(PATIENT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return {patientDataAvailable:false, ageDistribution:[], animationFrames:[], bloodSeverity:[], associations:[]};
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const required = ['admission_date','patient_type','severity_level','gender','age','blood_group','disease'];
  const columns = {};
  required.forEach(name => {
    columns[name] = headers.indexOf(name);
    if (columns[name] < 0) throw new Error('Patient Data is missing the required column: ' + name);
  });

  const ageBins = {};
  const frameCounts = {};
  const bloodCounts = {};
  const records = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const date = asISODate_(row[columns.admission_date]);
    if (date < REPORT_START || date > REPORT_END) continue;
    const patientType = String(row[columns.patient_type] || '').trim();
    const severity = String(row[columns.severity_level] || '').trim();
    const gender = String(row[columns.gender] || '').trim();
    const bloodGroup = String(row[columns.blood_group] || '').trim();
    const disease = String(row[columns.disease] || '').trim();
    const age = Number(row[columns.age]);
    if (!patientType || !severity || !gender || !bloodGroup || !disease || !isFinite(age)) continue;
    const item = {date,patientType,severity,gender,bloodGroup,disease,age};
    records.push(item);

    const lower = Math.floor(age / 10) * 10;
    const ageLabel = lower + '–' + (lower + 9);
    const ageKey = gender + '|' + ageLabel;
    ageBins[ageKey] = (ageBins[ageKey] || 0) + 1;
    const frameKey = [date,patientType,severity].join('|');
    frameCounts[frameKey] = (frameCounts[frameKey] || 0) + 1;
    const bloodKey = [bloodGroup,severity].join('|');
    bloodCounts[bloodKey] = (bloodCounts[bloodKey] || 0) + 1;
  }

  const ageLabels = [];
  for (let lower = 0; lower <= 120; lower += 10) ageLabels.push(lower + '–' + (lower + 9));
  const genders = Array.from(new Set(records.map(x => x.gender))).sort();
  const ageDistribution = [];
  genders.forEach(gender => ageLabels.forEach(label => {
    ageDistribution.push({gender,ageBand:label,count:ageBins[gender + '|' + label] || 0});
  }));

  const frameMap = {};
  Object.keys(frameCounts).forEach(key => {
    const [date,patientType,severity] = key.split('|');
    if (!frameMap[date]) frameMap[date] = [];
    frameMap[date].push({patientType,severity,count:frameCounts[key]});
  });
  const animationFrames = Object.keys(frameMap).sort().map(date => ({date,rows:frameMap[date]}));
  const bloodSeverity = Object.keys(bloodCounts).map(key => {
    const [bloodGroup,severity] = key.split('|');
    return {bloodGroup,severity,count:bloodCounts[key]};
  });
  const associations = associationMatrix_(records);
  return {patientDataAvailable:true,ageDistribution,animationFrames,bloodSeverity,associations};
}

function asISODate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) return match[1];
  const dayFirst = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (dayFirst) return dayFirst[3] + '-' + String(dayFirst[2]).padStart(2,'0') + '-' + String(dayFirst[1]).padStart(2,'0');
  return '';
}

function associationMatrix_(records) {
  const vars = ['patient_type','gender','age','blood_group','disease'];
  const labels = {patient_type:'Patient type',gender:'Gender',age:'Age',blood_group:'Blood group',disease:'Disease'};
  const result = [];
  for (let i = 0; i < vars.length; i++) {
    for (let j = 0; j < vars.length; j++) {
      if (i === j) {
        result.push({x:labels[vars[i]],y:labels[vars[j]],value:1,method:vars[i] === 'age' ? 'Self association' : 'Cramér’s V'});
      } else if (vars[i] === 'age' || vars[j] === 'age') {
        const groupField = vars[i] === 'age' ? vars[j] : vars[i];
        result.push({x:labels[vars[i]],y:labels[vars[j]],value:eta_(records,groupField),method:'Correlation ratio (η)'});
      } else {
        result.push({x:labels[vars[i]],y:labels[vars[j]],value:cramersV_(records,vars[i],vars[j]),method:'Cramér’s V'});
      }
    }
  }
  return result;
}

function cramersV_(records, fieldA, fieldB) {
  const counts = {}, rows = {}, cols = {};
  records.forEach(r => {
    const a = String(r[fieldA]), b = String(r[fieldB]);
    if (!counts[a]) counts[a] = {};
    counts[a][b] = (counts[a][b] || 0) + 1;
    rows[a] = (rows[a] || 0) + 1;
    cols[b] = (cols[b] || 0) + 1;
  });
  const n = records.length, rowKeys = Object.keys(rows), colKeys = Object.keys(cols);
  if (n < 1 || Math.min(rowKeys.length - 1,colKeys.length - 1) < 1) return 0;
  let chi2 = 0;
  rowKeys.forEach(a => colKeys.forEach(b => {
    const expected = rows[a] * cols[b] / n;
    const observed = counts[a][b] || 0;
    if (expected > 0) chi2 += Math.pow(observed - expected,2) / expected;
  }));
  return Math.sqrt(chi2 / (n * Math.min(rowKeys.length - 1,colKeys.length - 1)));
}

function eta_(records, groupField) {
  const groups = {}, ages = records.map(r => Number(r.age));
  const mean = ages.reduce((a,b) => a + b,0) / (ages.length || 1);
  const total = ages.reduce((sum,x) => sum + Math.pow(x - mean,2),0);
  if (!total) return 0;
  records.forEach(r => {
    const key = String(r[groupField]);
    if (!groups[key]) groups[key] = [];
    groups[key].push(Number(r.age));
  });
  let within = 0;
  Object.keys(groups).forEach(key => {
    const xs = groups[key], avg = xs.reduce((a,b) => a + b,0) / xs.length;
    xs.forEach(x => within += Math.pow(x - avg,2));
  });
  return Math.sqrt(Math.max(0,(total - within) / total));
}
