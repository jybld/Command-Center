// ============================================================
//  OUTSIDE SEARCH COMMAND CENTER — Code.gs  v4.2
// ============================================================

const ENGINE_ID          = '1nGssGvwIdgIg7PAS_QINPG9YFhYn1LM8tgHENB-hM-o';
const DASHBOARD_ID       = '1gHkeuz5SsPixghH8z3q6orMXs_3zSMjHf4Pr7lbk0Yw';
const TRACKER_ID         = '1DFzm0LS1MywIRdOJwRaJJz-_CfhrcXt_TpEnEnVnmqY';
const INSIDE_SEARCH_ID   = '1_mQ5-cuZfHZiChfF3lZPXaOCzBmdESktIS5o2UPxRzE';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Outside Search Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══ RECOMMENDATION ENGINE ═══════════════════════════════════

function getRecommendation() {
  var ss   = SpreadsheetApp.openById(ENGINE_ID);
  var sh   = ss.getSheetByName('Client_Input');
  var data = sh.getRange('A5:F24').getValues();

  var backup1Id    = data[7][4] ? String(data[7][4]).trim() : '';
  var backup2Id    = data[8][4] ? String(data[8][4]).trim() : '';
  var backup1Login = backup1Id ? getLoginForAccount_(backup1Id) : '';
  var backup2Login = backup2Id ? getLoginForAccount_(backup2Id) : '';
  var backup1Owner = backup1Id ? getOwnerForAccount_(backup1Id) : '';
  var backup2Owner = backup2Id ? getOwnerForAccount_(backup2Id) : '';

  return {
    clientName    : data[0][1],
    appNeeded     : data[1][1],
    city          : data[2][1],
    agePref       : data[3][1],
    clientGender  : data[4][1],
    bestAccountId : data[0][4],
    bestLogin     : data[1][4],
    bestHolder    : data[2][4],
    bestScore     : data[3][4],
    bestOwnerVA   : getOwnerForAccount_(String(data[0][4] || '').trim()),
    whyWon        : data[5][4],
    backup1       : backup1Id,
    backup1Login  : backup1Login,
    backup1Owner  : backup1Owner,
    backup2       : backup2Id,
    backup2Login  : backup2Login,
    backup2Owner  : backup2Owner,
    selectedAccount : data[15][1],
    searcherType    : data[16][1],
    currentHolder   : data[17][1],
    saveStatus      : data[18][1]
  };
}

function setClientInput(clientName, app, city, agePref, gender) {
  var ss = SpreadsheetApp.openById(ENGINE_ID);
  var sh = ss.getSheetByName('Client_Input');
  sh.getRange('B5').setValue(clientName);
  sh.getRange('B6').setValue(app);
  sh.getRange('B7').setValue(city.toUpperCase());
  sh.getRange('B8').setValue(agePref);
  sh.getRange('B9').setValue(gender);
  SpreadsheetApp.flush();
  Utilities.sleep(2500);
  return getRecommendation();
}
function saveAssignment(params) {
  params = params || {};

  var selectedAccount = params.selectedAccount || '';
  var searcherType    = params.searcherType || '';
  var holderName      = params.holderName || '';
  var accountId       = params.accountId || '';
  var accountLogin    = params.accountLogin || '';
  var clientName      = params.clientName || '';
  var city            = params.city || '';
  var ownerVA         = params.ownerVA || '';
  var app             = params.app || '';
  var agePref         = params.agePref || '';

  if (!accountId)    throw new Error('Missing account ID.');
  if (!accountLogin) throw new Error('Missing login email.');
  if (!holderName)   throw new Error('Missing Assign to VA / current holder.');
  if (!clientName)   throw new Error('Missing client name.');

  confirmAssignment(selectedAccount, searcherType, ownerVA, holderName);

  var logResult = addUsageLog({
    accountId: accountId,
    accountLogin: accountLogin,
    clientName: clientName,
    city: city,
    ownerVA: ownerVA,
    holderName: holderName,
    app: app,
    agePref: agePref,
    searcherType: searcherType
  });

  if (logResult !== 'logged') {
    throw new Error('Usage log did not save: ' + logResult);
  }

  var masterResult = updateAccountMaster(accountId, clientName, city, holderName);

  if (masterResult !== 'updated') {
    throw new Error('Accounts_Master did not update: ' + masterResult);
  }

  return {
    success: true,
    message: 'Assignment saved',
    accountId: accountId,
    accountLogin: accountLogin
  };
}
function confirmAssignment(selectedAccount, searcherType, ownerVA, holderName) {
  var ss = SpreadsheetApp.openById(ENGINE_ID);
  var sh = ss.getSheetByName('Client_Input');
  sh.getRange('E19').setValue(selectedAccount);
  sh.getRange('E20').setValue(searcherType);
  sh.getRange('E22').setValue(holderName);
  SpreadsheetApp.flush();
  return 'saved';
}

// ═══ HELPERS — Accounts_Master lookups ════════════════════════

function getMasterRows_() {
  var ss  = SpreadsheetApp.openById(ENGINE_ID);
  var sh  = ss.getSheetByName('Accounts_Master');
  var lr  = sh.getLastRow();
  if (lr < 4) return [];
  return sh.getRange(3, 1, lr - 2, 17).getValues();
}

function getLoginForAccount_(accountId) {
  if (!accountId) return '';
  var rows = getMasterRows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === accountId.trim())
      return String(rows[i][2]).trim();
  }
  return '';
}

function getOwnerForAccount_(accountId) {
  if (!accountId) return '';
  var rows = getMasterRows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === accountId.trim())
      return String(rows[i][4]).trim();
  }
  return '';
}

// ═══ ACTIVE ACCOUNTS — ENGINE → Accounts_Master ══════════════
// FIX: both getAccountsMaster and getActiveAccountsUsers defined

function getAccountsMaster() {
  return getMasterRows_();
}

function getActiveAccountsUsers() {
  var ss      = SpreadsheetApp.openById(ENGINE_ID);
  var sh      = ss.getSheetByName('Accounts_Master');
  var lastRow = sh.getLastRow();
  if (lastRow < 4) return [];
  return sh.getRange(3, 1, lastRow - 2, 17).getValues();
}

function updateAccountMaster(accountId, clientName, city, holderName) {
  var ss   = SpreadsheetApp.openById(ENGINE_ID);
  var sh   = ss.getSheetByName('Accounts_Master');
  var data = sh.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(accountId).trim()) {
      sh.getRange(i + 1, 10).setValue(clientName);
      sh.getRange(i + 1, 11).setValue(city);
      sh.getRange(i + 1, 16).setValue(holderName);
      sh.getRange(i + 1, 17).setValue(new Date());
      SpreadsheetApp.flush();
      return 'updated';
    }
  }
  return 'account not found';
}

// ═══ ACCOUNT LOOKUP — TRACKER → Data Source ══════════════════

function getTrackerRows_() {
  var ss  = SpreadsheetApp.openById(TRACKER_ID);
  var sh  = ss.getSheetByName('Data Source');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 10).getValues();
}

function lookupAccount(query, searchType) {
  searchType = searchType || 'email';
  query = (query || '').trim().toLowerCase();

  var rows    = getTrackerRows_();
  var results = [];

  // Return all rows if searchType is 'all' or query is empty
  var returnAll = (searchType === 'all' || query === '');

  for (var i = 0; i < rows.length; i++) {
    var r            = rows[i];
    var email        = String(r[0] || '').trim().toLowerCase();
    var vaOwner      = String(r[2] || '').trim().toLowerCase();
    var platform     = String(r[3] || '').trim();
    var currentUser  = String(r[4] || '').trim().toLowerCase();
    var clientAssign = String(r[5] || '').trim().toLowerCase();
    var status       = String(r[6] || '').trim();

    if (!email) continue;

    var match = returnAll;
    if (!returnAll) {
      if (searchType === 'email'  && email.indexOf(query) !== -1)                                              match = true;
      if (searchType === 'va'     && (vaOwner.indexOf(query) !== -1 || currentUser.indexOf(query) !== -1))     match = true;
      if (searchType === 'client' && clientAssign.indexOf(query) !== -1)                                       match = true;
    }

    if (match) {
      results.push({
        rowIndex      : i + 2,
        email         : String(r[0] || '').trim(),
        gender        : String(r[1] || '').trim(),
        ownerVA       : String(r[2] || '').trim(),
        app           : platform,
        currentUser   : String(r[4] || '').trim(),
        clientAssigned: String(r[5] || '').trim(),
        status        : status,
        lastUpdated   : String(r[7] || '').trim(),
        priorityTag   : String(r[8] || '').trim(),
        notes         : String(r[9] || '').trim()
      });
    }
  }
  return results;
}

function setAccountStatus(params) {
  var email      = (params.email      || '').trim().toLowerCase();
  var app        = (params.app        || '').trim();
  var status     = (params.status     || '').trim();
  var vaName     = (params.vaName     || '').trim();
  var clientName = (params.clientName || '').trim();

  if (!email) return { success: false, message: 'No email provided.' };
  if (!app)   return { success: false, message: 'No app specified.' };

  var ss   = SpreadsheetApp.openById(TRACKER_ID);
  var sh   = ss.getSheetByName('Data Source');
  if (!sh) return { success: false, message: 'Sheet "Data Source" not found.' };

  var lr   = sh.getLastRow();
  var data = sh.getRange(2, 1, lr - 1, 10).getValues();
  var ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yy | h:mm a');
  var updated = 0;

  for (var i = 0; i < data.length; i++) {
    var rowEmail    = String(data[i][0] || '').trim().toLowerCase();
    var rowPlatform = String(data[i][3] || '').trim();

    if (rowEmail !== email) continue;
    if (rowPlatform.toLowerCase() !== app.toLowerCase()) continue;

    var sheetRow = i + 2;

    if (status === 'In Use') {
      sh.getRange(sheetRow, 5).setValue(vaName);
      sh.getRange(sheetRow, 6).setValue(clientName);
      sh.getRange(sheetRow, 7).setValue('In Use');
      sh.getRange(sheetRow, 8).setValue(ts);
    } else {
      sh.getRange(sheetRow, 5).setValue('');
      sh.getRange(sheetRow, 6).setValue('');
      sh.getRange(sheetRow, 7).setValue('Available');
      sh.getRange(sheetRow, 8).setValue(ts);
    }
    updated++;
    break;
  }

  SpreadsheetApp.flush();

  if (updated === 0) return { success: false, message: 'No match for ' + params.email + ' / ' + app };
  return { success: true, message: 'Status set to ' + status + ' (' + app + ')' };
}

// ═══ USAGE LOG ════════════════════════════════════════════════

function addUsageLog(params) {
  var ss    = SpreadsheetApp.openById(ENGINE_ID);
  var sh    = ss.getSheetByName('Account_Usage_Log')
             || ss.getSheetByName('Account Usage Log')
             || ss.getSheetByName('Account Usage Log (Engine Source)');
  if (!sh) return 'sheet not found';
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var city  = (params.city || '').trim();

  sh.appendRow([
    today,
    params.clientName   || '',  // B: Client
    params.app          || '',  // C: App
    params.accountId    || '',  // D: Account_ID
    params.accountLogin || '',  // E: Account_Login
    params.ownerVA      || '',  // F: Owner_VA
    city.replace(/\w\S*/g, function(w){return w.charAt(0).toUpperCase()+w.slice(1).toLowerCase();}), // G: Target_City — Title Case
    city.toUpperCase(),         // H: City_Normalized — uppercase
    params.agePref      || '',  // I: Age_Preference
    getAgeBucket_(params.agePref), // J: Target_Age_Bucket
    params.searcherType || '',  // K: Searcher_Type
    'Client Assignments',       // L: Notes
    params.holderName   || ''   // M: Current_Holder
  ]);
  return 'logged';
}

function getAgeBucket_(agePref) {
  agePref = String(agePref || '').trim();
  if (!agePref) return '';
  var first = Number(agePref.split('-')[0].trim());
  if (isNaN(first)) return '';
  return first < 35 ? 'Young' : 'Older';
}

// ═══ CLIENT LIST — DASHBOARD → Client Assignments ═════════════
// FIX: explicitly uses getSheetByName('Client Assignments')

function getClientAssignments() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Client Assignments');
  if (!sh) {
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      var n = all[i].getName().toLowerCase();
      if (n.indexOf('client') !== -1 && n.indexOf('assign') !== -1) { sh = all[i]; break; }
    }
  }
  if (!sh) return [['Sheet "Client Assignments" not found']];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(1, 1, lr, sh.getLastColumn()).getValues();
}

function updateClientAssignment(row, col, value) {
  var ss = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh = ss.getSheetByName('Client Assignments') || ss.getSheets()[0];
  sh.getRange(row, col).setValue(value);
  return 'updated';
}

function addClientAssignment(rowData) {
  var ss = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh = ss.getSheetByName('Client Assignments') || ss.getSheets()[0];
  sh.appendRow(rowData);
  return 'added';
}

// ═══ EOD REPORT — Return all Usage Log rows as raw strings ═══
// Returns all data rows with date as a formatted string M/D/YYYY
// Frontend does the date filtering to avoid all timezone issues

function getUsageLogAllRows() {
  var ss = SpreadsheetApp.openById(ENGINE_ID);
  var sh = ss.getSheetByName('Account_Usage_Log')
          || ss.getSheetByName('Account Usage Log')
          || ss.getSheetByName('Account Usage Log (Engine Source)');
  if (!sh) return { error: 'Sheet not found', rows: [] };
  var lr = sh.getLastRow();
  if (lr < 2) return { rows: [] };

  // Read all data — row 1 = headers, data from row 2
  var data = sh.getRange(2, 1, lr - 1, 13).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var cellVal = data[i][0];
    if (!cellVal) continue;
    var dateStr = '';
    if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
      // Format as M/D/YYYY using local timezone
      dateStr = Utilities.formatDate(cellVal, Session.getScriptTimeZone(), 'M/d/yyyy');
    } else {
      dateStr = String(cellVal).trim();
    }
    // Replace index 0 with formatted date string so frontend can compare easily
    var row = data[i].slice(); // copy
    row[0] = dateStr;
    rows.push(row);
  }
  return { rows: rows };
}

// Keep old function name as alias for router compatibility
function getUsageLogByDate(params) {
  var dateStr = (params.date || '').trim();
  if (!dateStr) return [];
  var result = getUsageLogAllRows();
  if (result.error || !result.rows) return [];
  return result.rows.filter(function(r) { return r[0] === dateStr; });
}

// ═══ INSIDE SEARCH DASHBOARD ══════════════════════════════════
// Reads VA summary table from 'Dashboard' tab
// Cols: A=VA Name, B=Total, C=Due Today, D=Overdue,
//       E=Due in 3 Days, F=Search in Process, G=Not Started,
//       H=Reviewing, I=Done Searching

function getInsideSearchDashboard() {
  var ss  = SpreadsheetApp.openById(INSIDE_SEARCH_ID);
  var sh  = ss.getSheetByName('Dashboard');
  if (!sh) return { error: 'Sheet "Dashboard" not found in Inside Search spreadsheet.' };
  var lr  = sh.getLastRow();
  if (lr < 2) return { rows: [] };
  // Row 1 = headers, data from row 2
  var data = sh.getRange(2, 1, lr - 1, 9).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    rows.push({
      vaName      : name,
      total       : Number(data[i][1]) || 0,
      dueToday    : Number(data[i][2]) || 0,
      overdue     : Number(data[i][3]) || 0,
      due3Days    : Number(data[i][4]) || 0,
      inProcess   : Number(data[i][5]) || 0,
      notStarted  : Number(data[i][6]) || 0,
      reviewing   : Number(data[i][7]) || 0,
      done        : Number(data[i][8]) || 0
    });
  }
  return { rows: rows };
}

// ═══ OVERDUE CLIENTS (Outside Search Dashboard) ════════════════
// Reads Client Assignments — col A=Client, col B=Due Date,
// col R (index 17) = Overall Status
// Returns clients where Due Date < today and not blank

function getOverdueClients() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Client Assignments');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  var data    = sh.getRange(2, 1, lr - 1, 26).getValues(); // A-Z covers col R
  var today   = new Date(); today.setHours(0, 0, 0, 0);
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var client  = String(data[i][0] || '').trim();
    var dueVal  = data[i][1]; // col B
    var status  = String(data[i][17] || '').trim(); // col R
    if (!client || !dueVal) continue;
    var dueDate = (dueVal instanceof Date) ? dueVal : new Date(dueVal);
    dueDate.setHours(0, 0, 0, 0);
    if (isNaN(dueDate.getTime())) continue;
    if (dueDate < today) {
      var diffDays = Math.floor((today - dueDate) / (1000 * 60 * 60 * 24));
      results.push({
        client   : client,
        dueDate  : (dueDate.getMonth()+1) + '/' + dueDate.getDate() + '/' + dueDate.getFullYear(),
        status   : status,
        daysOver : diffDays
      });
    }
  }
  // Sort most overdue first
  results.sort(function(a, b){ return b.daysOver - a.daysOver; });
  return results;
}

// ═══ OUTSIDE SEARCH CLIENTS ═══════════════════════════════════
// Reads Client Assignments — col A=Client, col B=Due Date,
// col V=Age Preference (index 21), col W=City (index 22),
// col R=Overall Status (index 17), col C=# of Women (index 2)
// Returns array of [client, dueDate, agePref, city, status, women]

function getOutsideSearchClients() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Client Assignments');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  // Read cols A-W (1-23); col E=5(idx4), col K=11(idx10), col R=18(idx17), col V=22(idx21), col W=23(idx22)
  var data    = sh.getRange(2, 1, lr - 1, 23).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var client = String(data[i][0] || '').trim();
    if (!client) continue;
    var dueVal     = data[i][1];          // col B
    var women      = data[i][2];          // col C
    var primaryVA  = String(data[i][4]  || '').trim();  // col E
    var secondaryVA= String(data[i][10] || '').trim();  // col K
    var status     = String(data[i][17] || '').trim();  // col R
    var agePref    = String(data[i][21] || '').trim();  // col V
    var city       = String(data[i][22] || '').trim();  // col W
    var dueStr = '';
    if (dueVal instanceof Date && !isNaN(dueVal.getTime())) {
      dueStr = (dueVal.getMonth()+1) + '/' + dueVal.getDate() + '/' + dueVal.getFullYear();
    } else if (dueVal) {
      dueStr = String(dueVal).trim();
    }
    // rowIndex = i+2 (data starts row 2, header is row 1)
    results.push([client, dueStr, agePref, city, status, women, i+2, primaryVA, secondaryVA]);
  }
  return results;
}
// result index map: 0=client 1=dueDate 2=agePref 3=city 4=status 5=women 6=rowIndex 7=primaryVA 8=secondaryVA

// ═══ OUTSIDE SEARCH — WRITE FUNCTIONS ═════════════════════════

// Helper: write status to col R (18) and also F(6), H(8), L(12), N(14)
// Skips F/H/L/N if status is Maxed Out or MAXED OUT ALL
function writeClientStatus_(sh, rowIndex, status) {
  var stLow = status.toLowerCase().trim();
  var isMaxed = stLow.includes('maxed out') || stLow.includes('maxed out all');
  var isSIP   = stLow === 'search in progress';
  sh.getRange(rowIndex, 18).setValue(status);  // col R — always
  // Skip F/H/L/N for Maxed Out (preserve individual statuses)
  // and for Search In Progress (VA sets manually when they start)
  if (!isMaxed && !isSIP) {
    sh.getRange(rowIndex, 6).setValue(status);   // col F — Primary Hinge Status
    sh.getRange(rowIndex, 8).setValue(status);   // col H — Primary Bumble Status
    sh.getRange(rowIndex, 12).setValue(status);  // col L — Secondary Hinge Status
    sh.getRange(rowIndex, 14).setValue(status);  // col N — Secondary Bumble Status
  }
}

// Update date (col B) and/or status (col R + F/H/L/N) for a row
function updateClientDateStatus(params) {
  var rowIndex = params.rowIndex;
  if (!rowIndex) return { success: false, message: 'No row index' };
  var ss = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh = ss.getSheetByName('Client Assignments');
  if (!sh) return { success: false, message: 'Sheet not found' };
  if (params.dueDate) {
    var d = new Date(params.dueDate);
    sh.getRange(rowIndex, 2).setValue(isNaN(d.getTime()) ? params.dueDate : d);
  }
  if (params.status) {
    writeClientStatus_(sh, rowIndex, params.status);
    if (params.status.toLowerCase() !== 'search in progress') {
      sh.getRange(rowIndex, 24).clearContent(); // col X — clear timer
    }
    // Log maxed out account if provided
    var stLow = params.status.toLowerCase();
    if ((stLow.includes('maxed out')) && params.maxedAccount) {
      logMaxedOut({
        clientName   : params.clientName || '',
        maxedAccount : params.maxedAccount,
        status       : params.status,
        primaryVA    : String(sh.getRange(rowIndex, 5).getValue() || ''),
        app          : params.maxedApp   || '',
        notes        : params.maxedNotes || ''
      });
    }
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// Assign VA(s), profiles (×3), date, and auto-set status to Search In Progress
function assignClient(params) {
  var rowIndex = params.rowIndex;
  if (!rowIndex) return { success: false, message: 'No row index' };
  var ss = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh = ss.getSheetByName('Client Assignments');
  if (!sh) return { success: false, message: 'Sheet not found' };
  if (params.primaryVA   !== undefined) sh.getRange(rowIndex, 5).setValue(params.primaryVA);   // col E
  if (params.secondaryVA !== undefined) sh.getRange(rowIndex, 11).setValue(params.secondaryVA); // col K
  if (params.profiles    !== undefined) sh.getRange(rowIndex, 3).setValue(Number(params.profiles) * 3);
  if (params.dueDate) {
    var d = new Date(params.dueDate);
    sh.getRange(rowIndex, 2).setValue(isNaN(d.getTime()) ? params.dueDate : d);
  }
  // Auto-set Search In Progress on R, F, H, L, N
  writeClientStatus_(sh, rowIndex, 'Search In Progress');
  // Record search start timestamp in col X (24) for timer tracking
  sh.getRange(rowIndex, 24).setValue(new Date());
  SpreadsheetApp.flush();
  return { success: true };
}

// ═══ INSIDE SEARCH — OVERDUE FROM VA SHEETS ═══════════════════
// Reads each VA's individual tab in the Inside Search spreadsheet
// Each tab: col A=Client, col B=Due Date, col C=Status (adjust if different)
// Skips the 'Dashboard' tab and any tab with <2 rows

function getInsideSearchOverdue() {
  var ss      = SpreadsheetApp.openById(INSIDE_SEARCH_ID);
  var sheets  = ss.getSheets();
  var today   = new Date(); today.setHours(0,0,0,0);
  var results = [];
  var EXCLUDE_STATUS = ['done', 'batch sent'];

  for (var i = 0; i < sheets.length; i++) {
    var sh   = sheets[i];
    var name = sh.getName();
    if (name === 'Dashboard' || name.toLowerCase().indexOf('ref') !== -1) continue;
    var lr = sh.getLastRow();
    if (lr < 2) continue;

    var numCols = Math.min(sh.getLastColumn(), 10);
    var data = sh.getRange(2, 1, lr - 1, numCols).getValues();

    for (var j = 0; j < data.length; j++) {
      var client = String(data[j][0] || '').trim();
      if (!client) continue;

      // Find due date: first Date object or date-like string in cols B-E (indices 1-4)
      var dueDate = null;
      var dueCellVal = null;
      for (var c = 1; c <= Math.min(4, numCols-1); c++) {
        var cv = data[j][c];
        if (cv instanceof Date && !isNaN(cv.getTime())) { dueCellVal=cv; break; }
        if (cv && String(cv).match(/\d{1,2}\/\d{1,2}\/\d{2,4}/)) { dueCellVal=cv; break; }
      }
      if (!dueCellVal) continue;
      dueDate = new Date(dueCellVal);
      if (isNaN(dueDate.getTime())) continue;
      dueDate.setHours(0,0,0,0);
      if (dueDate > today) continue;

      // Find status: look in cols B-J for non-empty string that is not a date or boolean
      var status = '';
      for (var c2 = 1; c2 < numCols; c2++) {
        var sv = data[j][c2];
        if (typeof sv === 'string' && sv.trim() && !(String(sv).match(/\d{1,2}\/\d{1,2}\/\d{2,4}/))) {
          status = sv.trim(); break;
        }
      }

      // Exclude Done and Batch Sent
      var stLow = status.toLowerCase();
      if (EXCLUDE_STATUS.some(function(ex){ return stLow.includes(ex); })) continue;

      var diffDays = Math.floor((today - dueDate) / (1000*60*60*24));
      results.push({
        vaName   : name,
        client   : client,
        dueDate  : (dueDate.getMonth()+1)+'/'+dueDate.getDate()+'/'+dueDate.getFullYear(),
        status   : status,
        daysOver : diffDays,
        dueToday : diffDays === 0
      });
    }
  }
  results.sort(function(a,b){ return b.daysOver - a.daysOver; });
  return results;
}

// ═══ EOD — INSIDE SEARCH SUMMARY (for EOD report) ═════════════

function getEODInsideSummary() {
  var data = getInsideSearchDashboard();
  if (!data || data.error || !data.rows) return { total:0, overdue:0, dueToday:0, inProcess:0, done:0, notStarted:0 };
  var totals = { total:0, overdue:0, dueToday:0, inProcess:0, done:0, notStarted:0 };
  data.rows.forEach(function(r){
    totals.total     += r.total     || 0;
    totals.overdue   += r.overdue   || 0;
    totals.dueToday  += r.dueToday  || 0;
    totals.inProcess += r.inProcess || 0;
    totals.done      += r.done      || 0;
    totals.notStarted+= r.notStarted|| 0;
  });
  return totals;
}

// ═══ SEARCH TIMER — record/check ══════════════════════════════
// Uses col X (index 24, col 24) in Client Assignments to store search start timestamp
// When assignClient is called, writes current timestamp to col X

function getSearchTimerFlags() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Client Assignments');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  // Read A, B, E, R, X (cols 1,2,5,18,24)
  var data = sh.getRange(2, 1, lr-1, 24).getValues();
  var today= new Date();
  var TWO_DAYS = 2 * 24 * 60 * 60 * 1000;
  var flagged = [];
  for (var i = 0; i < data.length; i++) {
    var client  = String(data[i][0] || '').trim();
    var status  = String(data[i][17] || '').trim(); // col R
    var timerVal= data[i][23]; // col X
    if (!client) continue;
    if (status.toLowerCase() !== 'search in progress') continue;
    if (!timerVal) continue;
    var started = (timerVal instanceof Date) ? timerVal : new Date(timerVal);
    if (isNaN(started.getTime())) continue;
    var elapsed = today - started;
    var elapsedHours = Math.floor(elapsed / (1000*60*60));
    var elapsedDays  = Math.floor(elapsed / (1000*60*60*24));
    var primaryVA    = String(data[i][4] || '').trim(); // col E
    flagged.push({
      client      : client,
      primaryVA   : primaryVA,
      startedDate : (started.getMonth()+1)+'/'+started.getDate(),
      elapsedHours: elapsedHours,
      elapsedDays : elapsedDays,
      flagged     : elapsed >= TWO_DAYS
    });
  }
  return flagged;
}

// ═══ DEBUG — Usage Log ════════════════════════════════════════
function getUsageLogDebug() {
  var ss = SpreadsheetApp.openById(ENGINE_ID);
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(s){ return s.getName(); });
  var sh = ss.getSheetByName('Account_Usage_Log')
          || ss.getSheetByName('Account Usage Log')
          || ss.getSheetByName('Account Usage Log (Engine Source)');
  var result = { allSheets: sheetNames, foundSheet: sh ? sh.getName() : 'NOT FOUND' };
  if (!sh) return result;
  result.lastRow = sh.getLastRow();
  // Read rows 1-5 col A
  var peek = sh.getRange(1, 1, Math.min(5, sh.getLastRow()), 1).getValues();
  result.colA_rows1to5 = peek.map(function(r){
    var v = r[0];
    if (v instanceof Date) return 'DATE:' + (v.getMonth()+1)+'/'+v.getDate()+'/'+v.getFullYear();
    return 'STRING:' + String(v);
  });
  // Read last 3 rows col A
  var lr = sh.getLastRow();
  if (lr >= 3) {
    var last3 = sh.getRange(lr-2, 1, 3, 1).getValues();
    result.colA_lastRows = last3.map(function(r){
      var v = r[0];
      if (v instanceof Date) return 'DATE:' + (v.getMonth()+1)+'/'+v.getDate()+'/'+v.getFullYear();
      return 'STRING:' + String(v);
    });
  }
  return result;
}

// ═══ OUTSIDE SEARCH STATUS SUMMARY (for EOD) ══════════════════
// Overdue = past due date AND status is in the active list
// Active overdue statuses: Search In Progress, Rework Needed, Pending Next Batch,
//   Need More Profile, Maxed Out, MAXED OUT ALL, Not Started, PAUSE SEARCH

function getOutsideSearchStatus() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Client Assignments');
  if (!sh) return { inProgress:0, activeOverdue:0, dueToday:0, notStarted:0, pending:0, total:0 };
  var lr  = sh.getLastRow();
  if (lr < 2) return { inProgress:0, activeOverdue:0, dueToday:0, notStarted:0, pending:0, total:0 };

  var data  = sh.getRange(2, 1, lr-1, 18).getValues();
  var today = new Date(); today.setHours(0,0,0,0);

  // Statuses that count as "active overdue" when past due
  var ACTIVE_OVERDUE_STATUSES = [
    'search in progress','rework needed','pending next batch',
    'need more profile','not started','pause search'
  ];
  var EXCLUDE = ['done','fc needed','completed','maxed out'];

  var inProgress=0, activeOverdue=0, dueToday=0, notStarted=0, pending=0, total=0;

  for (var i=0; i<data.length; i++) {
    var client = String(data[i][0]||'').trim();
    if (!client) continue;
    total++;
    var status = String(data[i][17]||'').trim().toLowerCase(); // col R
    var dueVal = data[i][1]; // col B

    if (status==='search in progress') inProgress++;
    if (status==='not started') notStarted++;
    if (status.includes('pending')) pending++;

    if (dueVal) {
      var dueDate = (dueVal instanceof Date) ? dueVal : new Date(dueVal);
      if (!isNaN(dueDate.getTime())) {
        var d = new Date(dueDate); d.setHours(0,0,0,0);
        if (d.getTime() === today.getTime()) dueToday++;
        if (d < today) {
          // Only count as overdue if status is in active list
          var isExcluded = EXCLUDE.some(function(ex){ return status.includes(ex); });
          var isActive   = ACTIVE_OVERDUE_STATUSES.some(function(st){ return status.includes(st); });
          if (!isExcluded && isActive) activeOverdue++;
        }
      }
    }
  }
  return { inProgress:inProgress, activeOverdue:activeOverdue, dueToday:dueToday, notStarted:notStarted, pending:pending, total:total };
}

// ═══ MAXED OUT LOG ════════════════════════════════════════════
// Cols: A=Date, B=Client, C=Maxed Account Email, D=Primary VA, E=Status, F=App, G=Notes

function getOrCreateMaxedOutLogSheet_() {
  var ss = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh = ss.getSheetByName('Maxed_Out_Log');
  if (!sh) {
    sh = ss.insertSheet('Maxed_Out_Log');
    sh.getRange(1, 1, 1, 7).setValues([['Date','Client','Maxed Account Email','Primary VA','Status','App','Notes']]);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function logMaxedOut(params) {
  var sh  = getOrCreateMaxedOutLogSheet_();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  // Use explicit app if provided, otherwise auto-detect from account email
  var app = (params.app || '').trim();
  if (!app) {
    var acct = String(params.maxedAccount || '').toLowerCase();
    app = acct.indexOf('atlasvas') !== -1 ? 'HNG' : '';
  }
  sh.appendRow([
    ts,
    params.clientName    || '',
    params.maxedAccount  || '',
    params.primaryVA     || '',
    params.status        || '',
    app,
    params.notes         || ''
  ]);
  SpreadsheetApp.flush();
  return 'logged';
}

function getMaxedOutHistory(params) {
  var clientName = (params.clientName || '').trim().toLowerCase();
  if (!clientName) return [];
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Maxed_Out_Log');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  var numCols = Math.min(sh.getLastColumn(), 7);
  var data    = sh.getRange(2, 1, lr - 1, numCols).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1] || '').trim().toLowerCase() === clientName) {
      results.push(data[i].map(function(v){
        return v instanceof Date
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'M/d/yyyy')
          : String(v || '');
      }));
    }
  }
  return results;
}

// Returns all maxed out entries for today (for EOD report)
function getMaxedOutToday() {
  var ss  = SpreadsheetApp.openById(DASHBOARD_ID);
  var sh  = ss.getSheetByName('Maxed_Out_Log');
  if (!sh) return [];
  var lr  = sh.getLastRow();
  if (lr < 2) return [];
  var today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var numCols = Math.min(sh.getLastColumn(), 7);
  var data    = sh.getRange(2, 1, lr - 1, numCols).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    var dateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'M/d/yyyy')
      : String(rowDate || '').trim();
    if (dateStr === today) {
      results.push({
        date    : dateStr,
        client  : String(data[i][1] || '').trim(),
        account : String(data[i][2] || '').trim(),
        va      : String(data[i][3] || '').trim(),
        status  : String(data[i][4] || '').trim(),
        app     : String(data[i][5] || '').trim(),
        notes   : String(data[i][6] || '').trim()
      });
    }
  }
  return results;
}

// ═══ SERVER CALL ROUTER ═══════════════════════════════════════

function serverCall(action, params) {
  params = params || {};
  switch (action) {
    case 'getOutsideSearchClients':    return getOutsideSearchClients();
    case 'getOutsideSearchStatus':     return getOutsideSearchStatus();
    case 'getUsageLogDebug':           return getUsageLogDebug();
    case 'getMaxedOutHistory':         return getMaxedOutHistory(params);
    case 'getMaxedOutToday':           return getMaxedOutToday();
    case 'updateClientDateStatus':     return updateClientDateStatus(params);
    case 'assignClient':               return assignClient(params);
    case 'getUsageLogByDate':          return getUsageLogByDate(params);
    case 'getUsageLogAllRows':         return getUsageLogAllRows();
    case 'getInsideSearchDashboard':   return getInsideSearchDashboard();
    case 'getInsideSearchOverdue':     return getInsideSearchOverdue();
    case 'getEODInsideSummary':        return getEODInsideSummary();
    case 'getSearchTimerFlags':        return getSearchTimerFlags();
    case 'getOverdueClients':          return getOverdueClients();
    case 'getRecommendation':      return getRecommendation();
    case 'setClientInput':         return setClientInput(params.clientName, params.app, params.city, params.agePref, params.gender);
    case 'saveAssignment':         return saveAssignment(params);
    case 'confirmAssignment':      return confirmAssignment(params.selectedAccount, params.searcherType, params.ownerVA, params.holderName);
    case 'getAccountsMaster':      return getAccountsMaster();
    case 'addUsageLog':            return addUsageLog(params);
    case 'lookupAccount':          return lookupAccount(params.query, params.searchType);
    case 'setAccountStatus':       return setAccountStatus(params);
    case 'getClientAssignments':   return getClientAssignments();
    case 'getActiveAccountsUsers': return getActiveAccountsUsers();
    case 'updateClientAssignment': return updateClientAssignment(params.row, params.col, params.value);
    case 'addClientAssignment':    return addClientAssignment(params.rowData);
    case 'updateAccountMaster':    return updateAccountMaster(params.accountId, params.clientName, params.city, params.holderName);
    default: return 'Unknown action: ' + action;
  }
}

// ═══ ON EDIT TRIGGER ══════════════════════════════════════════

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (range.getColumn() === 13 && range.getRow() > 1) {
    var dateCell = sheet.getRange(range.getRow(), 14);
    if (range.getValue() !== '' && dateCell.getValue() === '') dateCell.setValue(new Date());
    if (range.getValue() === '') dateCell.clearContent();
  }
}
