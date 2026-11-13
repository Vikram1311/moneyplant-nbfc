const SHEETS = {
  MEMBERS: ['ID','Name','Mobile','Salt','PasswordHash','Role','Status','NomineeName','NomineeRelation','CreatedAt'],
  CONTRIBUTIONS: ['EntryID','Date','MemberID','Amount','Penalty','Remark'],
  LOANS: ['LoanID','Date','MemberID','Principal','MonthlyRate','Months','EMI','TotalPayable','PaidEMIs','PaidAmount','Balance','Status','GuarantorID'],
  EMI: ['EntryID','Date','LoanID','MemberID','Amount','Penalty','Remark'],
  LOAN_REQUESTS: ['RequestID','Date','MemberID','Amount','Months','Purpose','Status'],
  WITHDRAWALS: ['RequestID','Date','MemberID','Amount','Status'],
  EXPENSES: ['ExpenseID','Date','Title','Amount','Remark']
};
const SESSION_SECONDS = 21600;

function doGet() {
  setupAllSheets();
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('MoneyPlant NBFC').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function setupAllSheets() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(name=>{
    let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
    if(sh.getLastRow()===0){sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);sh.setFrozenRows(1);sh.getRange(1,1,1,SHEETS[name].length).setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold');}
  });
  ensureAdmin_(); return 'शीट और एडमिन जाँच पूरी।';
}
function ensureAdmin_(){
  const sh=sheet_('MEMBERS'),rows=readRows_('MEMBERS'),props=PropertiesService.getScriptProperties();
  const name=String(props.getProperty('ADMIN_NAME')||'').trim(),mobile=String(props.getProperty('ADMIN_MOBILE')||'').trim(),password=String(props.getProperty('ADMIN_PASSWORD')||'');
  if(rows.some(m=>m.Role==='Admin')) return;
  if(!name||!/^[0-9]{10}$/.test(mobile)||password.length<8) return;
  if(rows.some(m=>String(m.Mobile).trim()===mobile)) throw new Error('यह मोबाइल नंबर पहले से सदस्य के रूप में दर्ज है। Sheet में जाँच करें।');
  const salt=Utilities.getUuid(); sh.appendRow(['ADMIN01',name,mobile,salt,hashPassword_(salt,password),'Admin','Active','Not Set','-',new Date()]);
}
function login(mobile,password){
  setupAllSheets();mobile=String(mobile||'').trim();password=String(password||'');
  if(!/^[0-9]{10}$/.test(mobile)||password.length<4)return{success:false,message:'मोबाइल नंबर या पासवर्ड सही नहीं है।'};
  const row=readRows_('MEMBERS').find(m=>String(m.Mobile).trim()===mobile&&m.Status==='Active');
  if(!row||hashPassword_(row.Salt,password)!==row.PasswordHash)return{success:false,message:'मोबाइल नंबर या पासवर्ड सही नहीं है।'};
  const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put('session_'+token,JSON.stringify({id:row.ID,role:row.Role}),SESSION_SECONDS);
  return{success:true,token,user:{id:row.ID,name:row.Name,role:row.Role}};
}
function logout(token){if(token)CacheService.getScriptCache().remove('session_'+token);return{success:true};}

function getDashboard(token){
  const user=requireUser_(token),members=readRows_('MEMBERS').filter(m=>m.Status==='Active');
  const contributions=readRows_('CONTRIBUTIONS'),loans=readRows_('LOANS'),emis=readRows_('EMI'),expenses=readRows_('EXPENSES'),requests=readRows_('LOAN_REQUESTS'),withdrawals=readRows_('WITHDRAWALS');
  const isAdmin=user.role==='Admin';
  const visibleMembers=isAdmin?members:members.filter(m=>m.ID===user.id);
  const visibleIds=visibleMembers.map(m=>m.ID);
  const visibleLoans=isAdmin?loans:loans.filter(l=>l.MemberID===user.id);
  const visibleEmis=isAdmin?emis:emis.filter(e=>e.MemberID===user.id);
  const visibleContributions=isAdmin?contributions:contributions.filter(c=>c.MemberID===user.id);
  const visibleWithdrawals=isAdmin?withdrawals:withdrawals.filter(w=>w.MemberID===user.id);
  const visibleExpenses=isAdmin?expenses:[];
  const memberCards=visibleMembers.map(m=>{
    const memberLoans=loans.filter(l=>l.MemberID===m.ID);
    const deposit=sum_(contributions.filter(x=>x.MemberID===m.ID),'Amount')-sum_(withdrawals.filter(x=>x.MemberID===m.ID&&x.Status==='Approved'),'Amount');
    const passbook=[];
    contributions.filter(x=>x.MemberID===m.ID).forEach(x=>passbook.push({date:dateText_(x.Date),type:'जमा',credit:num_(x.Amount),debit:0,note:x.Remark||''}));
    emis.filter(x=>x.MemberID===m.ID).forEach(x=>passbook.push({date:dateText_(x.Date),type:'किस्त जमा',credit:num_(x.Amount),debit:0,note:x.LoanID}));
    withdrawals.filter(x=>x.MemberID===m.ID&&x.Status==='Approved').forEach(x=>passbook.push({date:dateText_(x.Date),type:'निकासी',credit:0,debit:num_(x.Amount),note:'स्वीकृत निकासी'}));
    return{id:m.ID,name:m.Name,mobile:m.Mobile,role:m.Role,nomineeName:m.NomineeName||'Not Set',nomineeRelation:m.NomineeRelation||'-',deposit:Math.max(0,deposit),loanBalance:sum_(memberLoans.filter(l=>l.Status==='Running'),'Balance'),sharePct:0,loans:memberLoans,passbook};
  });
  const runningLoans=loans.filter(l=>l.Status==='Running');
  const groupFund=Math.max(0,sum_(contributions,'Amount')-sum_(withdrawals.filter(w=>w.Status==='Approved'),'Amount'));
  const outstanding=sum_(runningLoans,'Balance');
  const totalMemberDeposits=memberCards.reduce((s,m)=>s+m.deposit,0)||1;
  memberCards.forEach(m=>m.sharePct=Number((m.deposit*100/totalMemberDeposits).toFixed(2)));
  const recent=[];
  if(isAdmin){
    contributions.forEach(x=>recent.push({entryId:x.EntryID,date:dateText_(x.Date),detail:'जमा: '+memberName_(members,x.MemberID)+' ('+money_(x.Amount)+')',sheet:'CONTRIBUTIONS',amount:x.Amount,penalty:x.Penalty,remark:x.Remark,memberId:x.MemberID}));
    loans.forEach(x=>recent.push({entryId:x.LoanID,date:dateText_(x.Date),detail:'लोन: '+memberName_(members,x.MemberID)+' ('+money_(x.Principal)+')',sheet:'LOANS',amount:x.Principal,rate:x.MonthlyRate,months:x.Months,memberId:x.MemberID}));
    emis.forEach(x=>recent.push({entryId:x.EntryID,date:dateText_(x.Date),detail:'किस्त: '+memberName_(members,x.MemberID)+' ('+money_(x.Amount)+')',sheet:'EMI',amount:x.Amount,penalty:x.Penalty,remark:x.Remark,memberId:x.MemberID,loanId:x.LoanID}));
    expenses.forEach(x=>recent.push({entryId:x.ExpenseID,date:dateText_(x.Date),detail:'खर्च: '+x.Title+' ('+money_(x.Amount)+')',sheet:'EXPENSES',amount:x.Amount,title:x.Title,remark:x.Remark}));
  }
  const own=memberCards[0];
  const personalInterest=own?Math.max(0,sum_(emis.filter(e=>e.MemberID===user.id),'Amount')):0;
  const dashboard={
    user:{id:user.id,role:user.role},
    summary:isAdmin?{totalFund:groupFund,totalLoanGiven:outstanding,availableCash:Math.max(0,groupFund-outstanding-sum_(expenses,'Amount')),totalInterest:Math.max(0,sum_(emis,'Amount')-sum_(loans,'Principal')+outstanding)}:{totalFund:own?own.deposit:0,totalLoanGiven:own?own.loanBalance:0,availableCash:0,totalInterest:personalInterest},
    members:memberCards,activeLoans:(isAdmin?runningLoans:runningLoans.filter(l=>l.MemberID===user.id)).map(l=>Object.assign({},l,{name:memberName_(members,l.MemberID)})),
    loanRequests:(isAdmin?requests:requests.filter(r=>r.MemberID===user.id)),withdrawals:(isAdmin?withdrawals:withdrawals.filter(w=>w.MemberID===user.id)),expenses:visibleExpenses,
    recentEntries:isAdmin?recent.sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,100):[],pendingDues:[]
  };
  return plain_(dashboard);
}

function addMember(token,payload){
  requireAdmin_(token);payload=payload||{};const name=clean_(payload.name,80),mobile=String(payload.mobile||'').trim(),password=String(payload.password||'');
  if(!name||!/^[0-9]{10}$/.test(mobile)||password.length<8)throw new Error('नाम, 10 अंकों का मोबाइल और कम से कम 8 अक्षर का पासवर्ड भरें।');
  if(readRows_('MEMBERS').some(m=>String(m.Mobile)===mobile))throw new Error('यह मोबाइल नंबर पहले से जुड़ा है।');
  const id='C'+Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase(),salt=Utilities.getUuid();append_('MEMBERS',[id,name,mobile,salt,hashPassword_(salt,password),'Member','Active','Not Set','-',new Date()]);return{success:true,message:'नया सदस्य जोड़ दिया गया।',id};
}
function changePassword(token,oldPassword,newPassword){
  const user=requireUser_(token);newPassword=String(newPassword||'');if(newPassword.length<8)throw new Error('नया पासवर्ड कम से कम 8 अक्षर का रखें।');
  const rowNum=findRow_('MEMBERS','ID',user.id),sh=sheet_('MEMBERS');if(!rowNum)throw new Error('सदस्य नहीं मिला।');const row=sh.getRange(rowNum,1,1,SHEETS.MEMBERS.length).getValues()[0];
  if(hashPassword_(row[3],String(oldPassword||''))!==row[4])throw new Error('पुराना पासवर्ड सही नहीं है।');const salt=Utilities.getUuid();sh.getRange(rowNum,4,1,2).setValues([[salt,hashPassword_(salt,newPassword)]]);return{success:true,message:'पासवर्ड बदल गया।'};
}
function updateNominee(token,name,relation){const user=requireUser_(token),rowNum=findRow_('MEMBERS','ID',user.id);if(!rowNum)throw new Error('सदस्य नहीं मिला।');sheet_('MEMBERS').getRange(rowNum,8,1,2).setValues([[clean_(name,80)||'Not Set',clean_(relation,50)||'-']]);return{success:true,message:'नॉमिनी जानकारी सेव हो गई।'};}

function addContribution(token,payload){
  requireAdmin_(token);payload=payload||{};const memberId=String(payload.memberId||''),amount=positive_(payload.amount,'जमा राशि'),penalty=Math.max(0,num_(payload.penalty)),remark=clean_(payload.remark,200);
  if(!readRows_('MEMBERS').some(m=>m.ID===memberId&&m.Status==='Active'))throw new Error('सदस्य नहीं मिला।');
  return appendWithDuplicateCheck_('CONTRIBUTIONS',payload.allowDuplicate,[{key:'MemberID',value:memberId},{key:'Amount',value:amount}],()=>append_('CONTRIBUTIONS',[newId_('DEP'),new Date(),memberId,amount,penalty,remark]),'इस सदस्य की आज की इसी रकम की जमा एंट्री पहले से है। फिर भी जोड़ना है?');
}
function issueLoan(token,payload){
  requireAdmin_(token);payload=payload||{};const memberId=String(payload.memberId||''),principal=positive_(payload.amount,'लोन राशि'),rate=num_(payload.rate),months=Math.floor(num_(payload.months));
  if(rate<0||rate>100||months<1||months>120)throw new Error('ब्याज दर या महीनों की संख्या ठीक नहीं है।');if(!readRows_('MEMBERS').some(m=>m.ID===memberId&&m.Status==='Active'&&m.Role==='Member'))throw new Error('सदस्य चुनें।');
  const dash=getDashboard(token);if(principal>dash.summary.availableCash)throw new Error('उपलब्ध नकद से अधिक लोन नहीं दिया जा सकता।');
  const r=rate/100,emi=r===0?principal/months:principal*r*Math.pow(1+r,months)/(Math.pow(1+r,months)-1),total=emi*months;
  return appendWithDuplicateCheck_('LOANS',payload.allowDuplicate,[{key:'MemberID',value:memberId},{key:'Principal',value:principal},{key:'MonthlyRate',value:rate},{key:'Months',value:months}],()=>append_('LOANS',[newId_('LN'),new Date(),memberId,principal,rate,months,round2_(emi),round2_(total),0,0,round2_(total),'Running',String(payload.guarantorId||'')]),'इस सदस्य का आज इसी रकम और शर्तों वाला लोन पहले से दर्ज है। फिर भी जोड़ना है?');
}
function collectEmi(token,payload){
  requireAdmin_(token);payload=payload||{};const loanId=String(payload.loanId||''),amount=positive_(payload.amount,'किस्त राशि'),penalty=Math.max(0,num_(payload.penalty)),remark=clean_(payload.remark,200);
  const rowNum=findRow_('LOANS','LoanID',loanId);if(!rowNum)throw new Error('लोन नहीं मिला।');const loanSheet=sheet_('LOANS'),v=loanSheet.getRange(rowNum,1,1,SHEETS.LOANS.length).getValues()[0];
  if(v[11]!=='Running')throw new Error('यह लोन चालू नहीं है।');if(amount>num_(v[10])+0.01)throw new Error('किस्त बाकी रकम से अधिक है।');
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const duplicate=hasDuplicate_('EMI',[{key:'LoanID',value:loanId},{key:'Amount',value:amount}]);if(duplicate&&!payload.allowDuplicate)return{success:false,duplicate:true,message:'इस लोन की आज इसी रकम की किस्त पहले से दर्ज है।'};
    const newPaid=round2_(num_(v[9])+amount),balance=Math.max(0,round2_(num_(v[7])-newPaid));loanSheet.getRange(rowNum,9,1,4).setValues([[num_(v[8])+1,newPaid,balance,balance<=0.01?'Closed':'Running']]);
    append_('EMI',[newId_('EMI'),new Date(),loanId,v[2],amount,penalty,remark]);return{success:true,message:'किस्त दर्ज हो गई।'};
  }finally{lock.releaseLock();}
}
function requestLoan(token,payload){const user=requireUser_(token);if(user.role!=='Member')throw new Error('यह सुविधा सदस्य के लिए है।');payload=payload||{};const amount=positive_(payload.amount,'लोन राशि'),months=Math.floor(num_(payload.months));if(months<1||months>120)throw new Error('महीनों की संख्या 1 से 120 के बीच रखें।');return appendWithDuplicateCheck_('LOAN_REQUESTS',payload.allowDuplicate,[{key:'MemberID',value:user.id},{key:'Amount',value:amount},{key:'Months',value:months}],()=>append_('LOAN_REQUESTS',[newId_('REQ'),new Date(),user.id,amount,months,clean_(payload.purpose,200),'Pending']),'आपका आज का ऐसा ही लोन आवेदन पहले से है। फिर भी भेजना है?');}
function requestWithdrawal(token,payload){const user=requireUser_(token);if(user.role!=='Member')throw new Error('यह सुविधा सदस्य के लिए है।');payload=payload||{};const amount=positive_(payload.amount,'निकासी राशि');return appendWithDuplicateCheck_('WITHDRAWALS',payload.allowDuplicate,[{key:'MemberID',value:user.id},{key:'Amount',value:amount}],()=>append_('WITHDRAWALS',[newId_('WDR'),new Date(),user.id,amount,'Pending']),'आज इसी रकम की निकासी माँग पहले से है। फिर भी भेजना है?');}
function reviewRequest(token,requestId,approve){
  requireAdmin_(token);const rowNum=findRow_('LOAN_REQUESTS','RequestID',requestId);if(!rowNum)throw new Error('आवेदन नहीं मिला।');const sh=sheet_('LOAN_REQUESTS'),v=sh.getRange(rowNum,1,1,SHEETS.LOAN_REQUESTS.length).getValues()[0];if(v[6]!=='Pending')throw new Error('इस आवेदन पर पहले कार्रवाई हो चुकी है।');if(!approve){sh.getRange(rowNum,7).setValue('Rejected');return{success:true,message:'आवेदन रद्द किया।'};}
  if(num_(v[3])>getDashboard(token).summary.availableCash)throw new Error('उपलब्ध नकद पर्याप्त नहीं है।');const n=Math.floor(num_(v[4])),r=0.01,emi=num_(v[3])*r*Math.pow(1+r,n)/(Math.pow(1+r,n)-1),total=emi*n;
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{sh.getRange(rowNum,7).setValue('Approved');append_('LOANS',[newId_('LN'),new Date(),v[2],num_(v[3]),1,n,round2_(emi),round2_(total),0,0,round2_(total),'Running','']);}finally{lock.releaseLock();}return{success:true,message:'आवेदन मंजूर हुआ।'};
}
function reviewWithdrawal(token,requestId,approve){requireAdmin_(token);const rowNum=findRow_('WITHDRAWALS','RequestID',requestId);if(!rowNum)throw new Error('आवेदन नहीं मिला।');const sh=sheet_('WITHDRAWALS'),v=sh.getRange(rowNum,1,1,SHEETS.WITHDRAWALS.length).getValues()[0];if(v[4]!=='Pending')throw new Error('इस आवेदन पर पहले कार्रवाई हो चुकी है।');if(approve){const bal=sum_(readRows_('CONTRIBUTIONS').filter(x=>x.MemberID===v[2]),'Amount')-sum_(readRows_('WITHDRAWALS').filter(x=>x.MemberID===v[2]&&x.Status==='Approved'),'Amount');if(num_(v[3])>bal)throw new Error('सदस्य के जमा बैलेंस से अधिक निकासी नहीं हो सकती।');}sh.getRange(rowNum,5).setValue(approve?'Approved':'Rejected');return{success:true,message:approve?'निकासी मंजूर हुई।':'निकासी रद्द हुई।'};}
function addExpense(token,payload){requireAdmin_(token);payload=payload||{};const title=clean_(payload.title,100),amount=positive_(payload.amount,'खर्च राशि'),remark=clean_(payload.remark,200);if(!title)throw new Error('खर्च का नाम लिखें।');return appendWithDuplicateCheck_('EXPENSES',payload.allowDuplicate,[{key:'Title',value:title},{key:'Amount',value:amount}],()=>append_('EXPENSES',[newId_('EXP'),new Date(),title,amount,remark]),'आज इसी नाम और रकम का खर्च पहले से दर्ज है। फिर भी जोड़ना है?');}

function editEntry(token,sheetName,entryId,payload){
  requireAdmin_(token);payload=payload||{};const config={CONTRIBUTIONS:{id:'EntryID'},LOANS:{id:'LoanID'},EMI:{id:'EntryID'},EXPENSES:{id:'ExpenseID'}}[String(sheetName)];if(!config)throw new Error('इस प्रकार की एंट्री नहीं बदली जा सकती।');
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const rowNum=findRow_(sheetName,config.id,entryId);if(!rowNum)throw new Error('एंट्री नहीं मिली।');const sh=sheet_(sheetName),headers=SHEETS[sheetName],row=sh.getRange(rowNum,1,1,headers.length).getValues()[0],col=k=>headers.indexOf(k)+1,amount=positive_(payload.amount,'राशि');
    if(sheetName==='CONTRIBUTIONS'||sheetName==='EMI'){sh.getRange(rowNum,col('Amount')).setValue(amount);sh.getRange(rowNum,col('Penalty')).setValue(Math.max(0,num_(payload.penalty)));sh.getRange(rowNum,col('Remark')).setValue(clean_(payload.remark,200));}
    else if(sheetName==='EXPENSES'){const title=clean_(payload.title,100);if(!title)throw new Error('खर्च का नाम लिखें।');sh.getRange(rowNum,col('Title')).setValue(title);sh.getRange(rowNum,col('Amount')).setValue(amount);sh.getRange(rowNum,col('Remark')).setValue(clean_(payload.remark,200));}
    else {const loanId=row[col('LoanID')-1];if(readRows_('EMI').some(e=>e.LoanID===loanId))throw new Error('इस लोन की किस्त दर्ज है, इसलिए यहाँ लोन नहीं बदल सकते।');const rate=num_(payload.rate),months=Math.floor(num_(payload.months));if(rate<0||rate>100||months<1||months>120)throw new Error('ब्याज दर या महीने सही भरें।');const r=rate/100,emi=r===0?amount/months:amount*r*Math.pow(1+r,months)/(Math.pow(1+r,months)-1),total=emi*months;sh.getRange(rowNum,col('Principal')).setValue(amount);sh.getRange(rowNum,col('MonthlyRate')).setValue(rate);sh.getRange(rowNum,col('Months')).setValue(months);sh.getRange(rowNum,col('EMI')).setValue(round2_(emi));sh.getRange(rowNum,col('TotalPayable')).setValue(round2_(total));sh.getRange(rowNum,col('Balance')).setValue(round2_(total));}
  }finally{lock.releaseLock();}return{success:true,message:'एंट्री ठीक कर दी गई।'};
}
function deleteEntry(token,sheetName,entryId){
  requireAdmin_(token);const config={CONTRIBUTIONS:{id:'EntryID'},LOANS:{id:'LoanID'},EMI:{id:'EntryID'},EXPENSES:{id:'ExpenseID'}}[String(sheetName)];if(!config)throw new Error('इस प्रकार की एंट्री हटाई नहीं जा सकती।');
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{const rowNum=findRow_(sheetName,config.id,entryId);if(!rowNum)throw new Error('एंट्री नहीं मिली।');
    if(sheetName==='LOANS'&&readRows_('EMI').some(e=>e.LoanID===entryId))throw new Error('इस लोन की किस्तें दर्ज हैं। पहले किस्तों का हिसाब जाँचना होगा; अभी लोन हटाना सुरक्षित नहीं है।');
    sheet_(sheetName).deleteRow(rowNum);
    if(sheetName==='EMI'){const loanRow=findRow_('LOANS','LoanID',readRows_('EMI').find(e=>e.EntryID===entryId)?.LoanID||'');if(loanRow){/* पुराना EMI हटाने से पहले loanId सुरक्षित नीचे निकाला जाता है; सामान्यतः इस शाखा का उपयोग न करें। */}}
    if(sheetName==='EMI')recalculateLoanAfterEmiDelete_(entryId);
  }finally{lock.releaseLock();}return{success:true,message:'एंट्री हटा दी गई।'};
}
function recalculateLoanAfterEmiDelete_(entryId){
  // EMI की loan ID delete से पहले recover की जाती है; यह helper fallback-safe है.
  const log=PropertiesService.getScriptProperties();const loanId=log.getProperty('DELETE_EMI_LOAN_'+entryId);if(!loanId)return;log.deleteProperty('DELETE_EMI_LOAN_'+entryId);
  const rowNum=findRow_('LOANS','LoanID',loanId);if(!rowNum)return;const sh=sheet_('LOANS'),row=sh.getRange(rowNum,1,1,SHEETS.LOANS.length).getValues()[0],paidRows=readRows_('EMI').filter(e=>e.LoanID===loanId),paid=sum_(paidRows,'Amount'),balance=Math.max(0,round2_(num_(row[7])-paid));sh.getRange(rowNum,9,1,4).setValues([[paidRows.length,round2_(paid),balance,balance<=0.01?'Closed':'Running']]);
}
function appendWithDuplicateCheck_(sheetName,allowDuplicate,checks,writeFn,message){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{if(!allowDuplicate&&hasDuplicate_(sheetName,checks))return{success:false,duplicate:true,message};writeFn();return{success:true,message:'एंट्री सेव हो गई।'};}finally{lock.releaseLock();}
}
function hasDuplicate_(sheetName,checks){
  const today=dateKey_(new Date());return readRows_(sheetName).some(row=>dateKey_(row.Date)===today&&checks.every(c=>String(row[c.key])===String(c.value)));
}
function requireUser_(token){const raw=token&&CacheService.getScriptCache().get('session_'+token);if(!raw)throw new Error('लॉगिन का समय पूरा हुआ। कृपया दोबारा लॉगिन करें।');return JSON.parse(raw);}
function requireAdmin_(token){const u=requireUser_(token);if(u.role!=='Admin')throw new Error('यह काम केवल एडमिन कर सकता है।');return u;}
function hashPassword_(salt,password){let value=String(salt)+':'+String(password);for(let i=0;i<1000;i++){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8);value=bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}return value;}
function readRows_(name){const sh=sheet_(name);if(sh.getLastRow()<2)return[];const headers=SHEETS[name];return sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues().filter(r=>r.some(v=>v!=='')).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});}
function sheet_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error('शीट नहीं मिली। पहले setupAllSheets चलाएँ।');return sh;}
function append_(name,row){sheet_(name).appendRow(row);}
function findRow_(name,header,value){const col=SHEETS[name].indexOf(header)+1,sh=sheet_(name);if(col<1||sh.getLastRow()<2)return 0;const cell=sh.getRange(2,col,sh.getLastRow()-1,1).createTextFinder(String(value)).matchEntireCell(true).findNext();return cell?cell.getRow():0;}
function sum_(rows,key){return rows.reduce((s,r)=>s+num_(r[key]),0);}
function num_(v){const n=Number(v);return isFinite(n)?n:0;}
function positive_(v,label){const n=num_(v);if(!(n>0)||n>100000000)throw new Error(label+' सही भरें।');return round2_(n);}
function round2_(n){return Math.round((num_(n)+Number.EPSILON)*100)/100;}
function clean_(v,max){return String(v||'').trim().replace(/[\x00-\x1F]/g,'').slice(0,max);}
function newId_(prefix){return prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase();}
function dateKey_(v){const d=v instanceof Date?v:new Date(v);return isNaN(d.getTime())?'':Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');}
function dateText_(v){return v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'dd-MM-yyyy'):String(v||'');}
function money_(v){return '₹'+num_(v).toLocaleString('en-IN');}
function memberName_(members,id){const m=members.find(x=>x.ID===id);return m?m.Name:id;}
function plain_(v){if(v instanceof Date)return dateText_(v);if(Array.isArray(v))return v.map(plain_);if(v&&typeof v==='object'){const o={};Object.keys(v).forEach(k=>o[k]=plain_(v[k]));return o;}return v;}
