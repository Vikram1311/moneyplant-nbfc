const SHEETS = {
  MEMBERS: ['ID', 'Name', 'Mobile', 'Salt', 'PasswordHash', 'Role', 'Status', 'NomineeName', 'NomineeRelation', 'CreatedAt'],
  CONTRIBUTIONS: ['EntryID', 'Date', 'MemberID', 'Amount', 'Penalty', 'Remark'],
  LOANS: ['LoanID', 'Date', 'MemberID', 'Principal', 'MonthlyRate', 'Months', 'EMI', 'TotalPayable', 'PaidEMIs', 'PaidAmount', 'Balance', 'Status', 'GuarantorID'],
  EMI: ['EntryID', 'Date', 'LoanID', 'MemberID', 'Amount', 'Penalty', 'Remark'],
  LOAN_REQUESTS: ['RequestID', 'Date', 'MemberID', 'Amount', 'Months', 'Purpose', 'Status'],
  WITHDRAWALS: ['RequestID', 'Date', 'MemberID', 'Amount', 'Status'],
  EXPENSES: ['ExpenseID', 'Date', 'Title', 'Amount', 'Remark']
};
const SESSION_SECONDS = 21600;

function doGet() {
  setupAllSheets();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MoneyPlant NBFC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once from the Apps Script editor. Set ADMIN_NAME, ADMIN_MOBILE and ADMIN_PASSWORD in Script Properties first. */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, SHEETS[name].length).setBackground('#4a148c').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  const members = ss.getSheetByName('MEMBERS');
  if (members.getLastRow() <= 1) {
    const props = PropertiesService.getScriptProperties();
    const name = props.getProperty('ADMIN_NAME');
    const mobile = props.getProperty('ADMIN_MOBILE');
    const password = props.getProperty('ADMIN_PASSWORD');
    if (name && mobile && password) {
      const salt = Utilities.getUuid();
      members.appendRow(['ADMIN01', name, mobile, salt, hashPassword_(salt, password), 'Admin', 'Active', 'Not Set', '-', new Date()]);
    }
  }
  return 'शीट तैयार हैं।';
}

function login(mobile, password) {
  setupAllSheets();
  mobile = String(mobile || '').trim();
  password = String(password || '');
  if (!/^\d{10}$/.test(mobile) || password.length < 4) return { success: false, message: 'मोबाइल नंबर या पासवर्ड सही नहीं है।' };
  const rows = readRows_('MEMBERS');
  const row = rows.find(r => String(r.Mobile).trim() === mobile && r.Status === 'Active');
  if (!row || hashPassword_(row.Salt, password) !== row.PasswordHash) return { success: false, message: 'मोबाइल नंबर या पासवर्ड सही नहीं है।' };
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({ id: row.ID, role: row.Role }), SESSION_SECONDS);
  return { success: true, token: token, user: { id: row.ID, name: row.Name, role: row.Role } };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function getDashboard(token) {
  const user = requireUser_(token);
  const members = readRows_('MEMBERS').filter(m => m.Status === 'Active');
  const allContributions = readRows_('CONTRIBUTIONS');
  const allLoans = readRows_('LOANS');
  const allEmis = readRows_('EMI');
  const allExpenses = readRows_('EXPENSES');
  const allRequests = readRows_('LOAN_REQUESTS');
  const allWithdrawals = readRows_('WITHDRAWALS');
  const visibleMembers = user.role === 'Admin' ? members : members.filter(m => m.ID === user.id);
  const visibleIds = visibleMembers.map(m => m.ID);
  const contributions = allContributions.filter(r => visibleIds.indexOf(r.MemberID) >= 0);
  const loans = allLoans.filter(r => visibleIds.indexOf(r.MemberID) >= 0 || (user.role === 'Admin' && r.Status === 'Running'));
  const emis = allEmis.filter(r => visibleIds.indexOf(r.MemberID) >= 0);
  const withdrawals = user.role === 'Admin' ? allWithdrawals : allWithdrawals.filter(r => r.MemberID === user.id);
  const requests = user.role === 'Admin' ? allRequests : allRequests.filter(r => r.MemberID === user.id);
  const expenses = user.role === 'Admin' ? allExpenses : [];
  const memberCards = visibleMembers.map(m => {
    const depositTotal = sum_(allContributions.filter(x => x.MemberID === m.ID), 'Amount') - sum_(allWithdrawals.filter(x => x.MemberID === m.ID && x.Status === 'Approved'), 'Amount');
    const memberLoans = allLoans.filter(x => x.MemberID === m.ID);
    const loanBalance = sum_(memberLoans.filter(x => x.Status === 'Running'), 'Balance');
    const passbook = [];
    allContributions.filter(x => x.MemberID === m.ID).forEach(x => passbook.push({ date: dateText_(x.Date), type: 'जमा', credit: num_(x.Amount), debit: 0, note: x.Remark || '' }));
    allEmis.filter(x => x.MemberID === m.ID).forEach(x => passbook.push({ date: dateText_(x.Date), type: 'किस्त जमा', credit: num_(x.Amount), debit: 0, note: x.LoanID }));
    allWithdrawals.filter(x => x.MemberID === m.ID && x.Status === 'Approved').forEach(x => passbook.push({ date: dateText_(x.Date), type: 'निकासी', credit: 0, debit: num_(x.Amount), note: 'स्वीकृत निकासी' }));
    return { id: m.ID, name: m.Name, mobile: m.Mobile, role: m.Role, nomineeName: m.NomineeName || 'Not Set', nomineeRelation: m.NomineeRelation || '-', deposit: Math.max(0, depositTotal), loanBalance: loanBalance, sharePct: 0, loans: memberLoans, passbook: passbook };
  });
  const groupFund = Math.max(0, sum_(allContributions, 'Amount') - sum_(allWithdrawals.filter(x => x.Status === 'Approved'), 'Amount'));
  const runningLoans = allLoans.filter(x => x.Status === 'Running');
  const outstanding = sum_(runningLoans, 'Balance');
  const interestReceived = allEmis.reduce((n, e) => {
    const loan = allLoans.find(l => l.LoanID === e.LoanID);
    return n + (loan ? Math.min(num_(e.Amount), num_(loan.EMI)) * num_(loan.MonthlyRate) / 100 : 0);
  }, 0);
  const totalMemberDeposits = memberCards.reduce((n, m) => n + m.deposit, 0) || 1;
  memberCards.forEach(m => m.sharePct = Number((m.deposit * 100 / totalMemberDeposits).toFixed(2)));
  const recent = [];
  allContributions.slice(-30).reverse().forEach(x => recent.push({ id: x.EntryID, date: dateText_(x.Date), detail: 'जमा: ' + memberName_(members, x.MemberID) + ' (' + money_(x.Amount) + ')', type: 'Contribution' }));
  allLoans.slice(-30).reverse().forEach(x => recent.push({ id: x.LoanID, date: dateText_(x.Date), detail: 'लोन: ' + memberName_(members, x.MemberID) + ' (' + money_(x.Principal) + ')', type: 'Loan' }));
  return {
    user: { id: user.id, role: user.role },
    summary: { totalFund: groupFund, totalLoanGiven: outstanding, availableCash: Math.max(0, groupFund - outstanding - sum_(allExpenses, 'Amount')), totalInterest: interestReceived },
    members: memberCards,
    activeLoans: runningLoans.map(l => Object.assign({}, l, { name: memberName_(members, l.MemberID) })),
    loanRequests: requests,
    withdrawals: withdrawals,
    expenses: expenses,
    recentEntries: user.role === 'Admin' ? recent.slice(0, 60) : [],
    pendingDues: []
  };
}

function addMember(token, payload) {
  requireAdmin_(token);
  payload = payload || {};
  const name = clean_(payload.name, 80);
  const mobile = String(payload.mobile || '').trim();
  const password = String(payload.password || '');
  if (!name || !/^\d{10}$/.test(mobile) || password.length < 8) throw new Error('नाम, 10 अंकों का मोबाइल और कम से कम 8 अक्षर का पासवर्ड भरें।');
  const rows = readRows_('MEMBERS');
  if (rows.some(m => String(m.Mobile) === mobile)) throw new Error('यह मोबाइल नंबर पहले से जुड़ा है।');
  const id = 'C' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  const salt = Utilities.getUuid();
  append_('MEMBERS', [id, name, mobile, salt, hashPassword_(salt, password), 'Member', 'Active', 'Not Set', '-', new Date()]);
  return { success: true, message: 'नया सदस्य जोड़ दिया गया।', id: id };
}

function changePassword(token, oldPassword, newPassword) {
  const user = requireUser_(token);
  newPassword = String(newPassword || '');
  if (newPassword.length < 8) throw new Error('नया पासवर्ड कम से कम 8 अक्षर का रखें।');
  const sheet = sheet_('MEMBERS');
  const rowNum = findRow_('MEMBERS', 'ID', user.id);
  if (!rowNum) throw new Error('सदस्य नहीं मिला।');
  const row = sheet.getRange(rowNum, 1, 1, SHEETS.MEMBERS.length).getValues()[0];
  if (hashPassword_(row[3], String(oldPassword || '')) !== row[4]) throw new Error('पुराना पासवर्ड सही नहीं है।');
  const salt = Utilities.getUuid();
  sheet.getRange(rowNum, 4, 1, 2).setValues([[salt, hashPassword_(salt, newPassword)]]);
  return { success: true, message: 'पासवर्ड बदल गया।' };
}

function updateNominee(token, name, relation) {
  const user = requireUser_(token);
  const rowNum = findRow_('MEMBERS', 'ID', user.id);
  if (!rowNum) throw new Error('सदस्य नहीं मिला।');
  sheet_('MEMBERS').getRange(rowNum, 8, 1, 2).setValues([[clean_(name, 80) || 'Not Set', clean_(relation, 50) || '-']]);
  return { success: true, message: 'नॉमिनी जानकारी सेव हो गई।' };
}

function addContribution(token, payload) {
  requireAdmin_(token);
  const memberId = String(payload.memberId || '');
  const amount = positive_(payload.amount, 'जमा राशि');
  const penalty = Math.max(0, num_(payload.penalty));
  if (!readRows_('MEMBERS').some(m => m.ID === memberId && m.Status === 'Active')) throw new Error('सदस्य नहीं मिला।');
  append_('CONTRIBUTIONS', [newId_('DEP'), new Date(), memberId, amount, penalty, clean_(payload.remark, 200)]);
  return { success: true, message: 'जमा दर्ज हो गई।' };
}

function issueLoan(token, payload) {
  requireAdmin_(token);
  const memberId = String(payload.memberId || '');
  const principal = positive_(payload.amount, 'लोन राशि');
  const rate = num_(payload.rate);
  const months = Math.floor(num_(payload.months));
  if (rate < 0 || rate > 100 || months < 1 || months > 120) throw new Error('ब्याज दर या महीनों की संख्या ठीक नहीं है।');
  const members = readRows_('MEMBERS');
  if (!members.some(m => m.ID === memberId && m.Status === 'Active' && m.Role === 'Member')) throw new Error('सदस्य चुनें।');
  const dash = getDashboard(token);
  if (principal > dash.summary.availableCash) throw new Error('उपलब्ध नकद से अधिक लोन नहीं दिया जा सकता।');
  const r = rate / 100;
  const emi = r === 0 ? principal / months : principal * r * Math.pow(1 + r, months) / (Math.pow(1 + r, months) - 1);
  const total = emi * months;
  const loanId = newId_('LN');
  append_('LOANS', [loanId, new Date(), memberId, principal, rate, months, round2_(emi), round2_(total), 0, 0, round2_(total), 'Running', String(payload.guarantorId || '')]);
  return { success: true, message: 'लोन दर्ज हो गया।' };
}

function collectEmi(token, payload) {
  requireAdmin_(token);
  const loanId = String(payload.loanId || '');
  const amount = positive_(payload.amount, 'किस्त राशि');
  const penalty = Math.max(0, num_(payload.penalty));
  const rowNum = findRow_('LOANS', 'LoanID', loanId);
  if (!rowNum) throw new Error('लोन नहीं मिला।');
  const sh = sheet_('LOANS');
  const v = sh.getRange(rowNum, 1, 1, SHEETS.LOANS.length).getValues()[0];
  if (v[11] !== 'Running') throw new Error('यह लोन चालू नहीं है।');
  if (amount > num_(v[10]) + 0.01) throw new Error('किस्त बाकी रकम से अधिक है।');
  const newPaid = num_(v[9]) + amount;
  const balance = Math.max(0, num_(v[7]) - newPaid);
  const paidEmis = num_(v[8]) + 1;
  sh.getRange(rowNum, 9, 1, 4).setValues([[paidEmis, round2_(newPaid), round2_(balance), balance <= 0.01 ? 'Closed' : 'Running']]);
  append_('EMI', [newId_('EMI'), new Date(), loanId, v[2], amount, penalty, clean_(payload.remark, 200)]);
  return { success: true, message: 'किस्त दर्ज हो गई।' };
}

function requestLoan(token, payload) {
  const user = requireUser_(token);
  if (user.role !== 'Member') throw new Error('यह सुविधा सदस्य के लिए है।');
  const amount = positive_(payload.amount, 'लोन राशि');
  const months = Math.floor(num_(payload.months));
  if (months < 1 || months > 120) throw new Error('महीनों की संख्या 1 से 120 के बीच रखें।');
  append_('LOAN_REQUESTS', [newId_('REQ'), new Date(), user.id, amount, months, clean_(payload.purpose, 200), 'Pending']);
  return { success: true, message: 'लोन आवेदन भेज दिया गया।' };
}

function requestWithdrawal(token, amount) {
  const user = requireUser_(token);
  if (user.role !== 'Member') throw new Error('यह सुविधा सदस्य के लिए है।');
  amount = positive_(amount, 'निकासी राशि');
  append_('WITHDRAWALS', [newId_('WDR'), new Date(), user.id, amount, 'Pending']);
  return { success: true, message: 'निकासी आवेदन भेज दिया गया।' };
}

function reviewRequest(token, requestId, approve) {
  requireAdmin_(token);
  const rowNum = findRow_('LOAN_REQUESTS', 'RequestID', requestId);
  if (!rowNum) throw new Error('आवेदन नहीं मिला।');
  const sh = sheet_('LOAN_REQUESTS');
  const v = sh.getRange(rowNum, 1, 1, SHEETS.LOAN_REQUESTS.length).getValues()[0];
  if (v[6] !== 'Pending') throw new Error('इस आवेदन पर पहले कार्रवाई हो चुकी है।');
  if (!approve) { sh.getRange(rowNum, 7).setValue('Rejected'); return { success: true, message: 'आवेदन अस्वीकार किया।' }; }
  const dash = getDashboard(token);
  if (num_(v[3]) > dash.summary.availableCash) throw new Error('उपलब्ध नकद पर्याप्त नहीं है।');
  sh.getRange(rowNum, 7).setValue('Approved');
  const member = readRows_('MEMBERS').find(m => m.ID === v[2]);
  issueLoan(token, { memberId: v[2], amount: v[3], rate: 1, months: v[4], guarantorId: '' });
  return { success: true, message: 'आवेदन मंजूर हुआ।' };
}

function reviewWithdrawal(token, requestId, approve) {
  requireAdmin_(token);
  const rowNum = findRow_('WITHDRAWALS', 'RequestID', requestId);
  if (!rowNum) throw new Error('आवेदन नहीं मिला।');
  const sh = sheet_('WITHDRAWALS');
  const v = sh.getRange(rowNum, 1, 1, SHEETS.WITHDRAWALS.length).getValues()[0];
  if (v[4] !== 'Pending') throw new Error('इस आवेदन पर पहले कार्रवाई हो चुकी है।');
  if (approve) {
    const available = sum_(readRows_('CONTRIBUTIONS').filter(x => x.MemberID === v[2]), 'Amount') - sum_(readRows_('WITHDRAWALS').filter(x => x.MemberID === v[2] && x.Status === 'Approved'), 'Amount');
    if (num_(v[3]) > available) throw new Error('सदस्य के जमा बैलेंस से अधिक निकासी नहीं हो सकती।');
  }
  sh.getRange(rowNum, 5).setValue(approve ? 'Approved' : 'Rejected');
  return { success: true, message: approve ? 'निकासी मंजूर हुई।' : 'निकासी अस्वीकार हुई।' };
}

function addExpense(token, payload) {
  requireAdmin_(token);
  const title = clean_(payload.title, 100);
  const amount = positive_(payload.amount, 'खर्च राशि');
  if (!title) throw new Error('खर्च का नाम लिखें।');
  append_('EXPENSES', [newId_('EXP'), new Date(), title, amount, clean_(payload.remark, 200)]);
  return { success: true, message: 'खर्च दर्ज हो गया।' };
}

function requireUser_(token) {
  const raw = token && CacheService.getScriptCache().get('session_' + token);
  if (!raw) throw new Error('लॉगिन का समय पूरा हुआ। कृपया दोबारा लॉगिन करें।');
  return JSON.parse(raw);
}
function requireAdmin_(token) {
  const user = requireUser_(token);
  if (user.role !== 'Admin') throw new Error('यह काम केवल एडमिन कर सकता है।');
  return user;
}
function hashPassword_(salt, password) {
  let value = String(salt) + ':' + String(password);
  for (let i = 0; i < 1000; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
    value = bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  }
  return value;
}
function readRows_(name) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  const headers = SHEETS[name];
  return sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues().filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('शीट नहीं मिली। पहले setupAllSheets चलाएँ।');
  return sh;
}
function append_(name, row) { sheet_(name).appendRow(row); }
function findRow_(name, header, value) {
  const headers = SHEETS[name];
  const col = headers.indexOf(header) + 1;
  const sh = sheet_(name);
  if (col < 1 || sh.getLastRow() < 2) return 0;
  const found = sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}
function sum_(rows, key) { return rows.reduce((total, row) => total + num_(row[key]), 0); }
function num_(value) { const n = Number(value); return isFinite(n) ? n : 0; }
function positive_(value, label) { const n = num_(value); if (!(n > 0) || n > 100000000) throw new Error(label + ' सही भरें।'); return round2_(n); }
function round2_(n) { return Math.round((num_(n) + Number.EPSILON) * 100) / 100; }
function clean_(value, max) { return String(value || '').trim().replace(/[\x00-\x1F]/g, '').slice(0, max); }
function newId_(prefix) { return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase(); }
function dateText_(value) { return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd-MM-yyyy') : String(value || ''); }
function money_(value) { return '₹' + num_(value).toLocaleString('en-IN'); }
function memberName_(members, id) { const m = members.find(x => x.ID === id); return m ? m.Name : id; }
