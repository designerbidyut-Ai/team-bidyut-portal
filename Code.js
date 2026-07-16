/**
 * Bidyut Content Tracker - Google Apps Script Backend
 * File: Code.js
 * Location: D:\SCT\Google sheet\Code.js
 */

/**
 * Maps spreadsheet column headers to standard frontend camelCase/PascalCase keys
 */
function mapHeaderToKey(header, sheetName) {
  let key = header.replace(/\s+/g, '');
  
  if (sheetName === 'dailyLogs') {
    if (key === "6sContent") return "Content6s";
    if (key === "10sContent") return "Content10s";
    if (key === "15sContent") return "Content15s";
    if (key === "20sContent") return "Content20s";
    if (key === "30sContent") return "Content30s";
    if (key === "1mContent") return "Content1m";
  }
  
  if (sheetName === 'accounts') {
    if (key === "Username/Email") return "UsernameEmail";
  }
  
  if (sheetName === 'prompts') {
    if (key === "Name") return "Title";
    if (key === "Prompttext") return "PromptText";
    if (key === "Targettools") return "Category";
  }
  
  return key;
}

// Global configuration
const ADMIN_PASSWORD = "1234"; // Default admin password (can be customized)

/**
 * Serves the HTML Web UI
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Team Bidyut Content Tracker & Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to include files inside template
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize or get sheets database
 */
function getDbSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = {
    members: ss.getSheetByName("Members") || ss.insertSheet("Members"),
    dailyLogs: ss.getSheetByName("DailyLogs") || ss.insertSheet("DailyLogs"),
    assignments: ss.getSheetByName("Assignments") || ss.insertSheet("Assignments"),
    fundTransactions: ss.getSheetByName("FundTransactions") || ss.insertSheet("FundTransactions"),
    accounts: ss.getSheetByName("Accounts") || ss.insertSheet("Accounts"),
    prompts: ss.getSheetByName("Prompts") || ss.insertSheet("Prompts"),
    brands: ss.getSheetByName("Brands") || ss.insertSheet("Brands")
  };
  
  // Set headers if sheets are empty/new
  setupHeaders(sheets);
  
  return sheets;
}

/**
 * Setup standard column headers for the relational tables
 */
function setupHeaders(sheets) {
  const headers = {
    members: ["Name", "Email", "Role", "Phone", "Status", "PIN", "Recommended Brands"],
    dailyLogs: ["Date", "Member Name", "6s Content", "10s Content", "15s Content", "20s Content", "30s Content", "1m Content", "Social Post", "Note"],
    assignments: ["Task ID", "Assign Date", "MPRO ID", "MPRO Link", "Group", "Quantity", "Completed Qty", "Status", "Deadline", "Assigned To", "Note", "Submit Date", "Approve Date"],
    fundTransactions: ["Date", "Type", "Amount", "Member Name", "Note", "Status"],
    accounts: ["Service Name", "Login URL", "Username/Email", "Password", "Expiry Date", "Note"],
    prompts: ["Title", "Category", "Prompt Text", "Description"],
    brands: ["Brand Name"]
  };
  
  for (let key in headers) {
    const sheet = sheets[key];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers[key]);
      
      // If it is the brands sheet, populate with default brands
      if (key === "brands") {
        const defaultBrands = ["PAL", "NAL", "PFL", "PCL", "PBL", "Taste Treet", "BBL", "RFL", "UAE", "pingo", "Italiano", "RPLM", "RPLE", "vision", "Dairy", "ROADBEAT", "AMCL", "Vision THA", "RYDO", "Mother Touch", "Vision Emporium", "RPL", "Sunny ToothBrush", "BBML", "DPL", "MUL", "falcon", "Getwell", "REL", "RAC", "Active Plus", "Rainbow Paints", "Pran Pulse", "PDL", "Bizli", "Good Luck", "Regal furniture", "Daily Dowyat", "TEL", "RMIL", "Golpo", "Daily Shoping", "RFL Best Buy", "BML", "Proton", "Malaysia"];
        defaultBrands.forEach(function(b) {
          sheet.appendRow([b]);
        });
      }
      
      // Format header row
      const range = sheet.getRange(1, 1, 1, headers[key].length);
      range.setFontWeight("bold");
      range.setBackground("#2c3e50");
      range.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  }
}

/**
 * Fetch list of team members (stripping PIN for security)
 */
function getMembers() {
  const db = getDbSheets();
  const data = db.members.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const members = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const member = {};
    headers.forEach((h, idx) => {
      if (h === "PIN") return; // Strip PIN for client security
      member[h.replace(/\s+/g, '')] = row[idx];
    });
    members.push(member);
  }
  return members;
}

/**
 * Add a new team member with default PIN 1234
 */
function addMember(name, email, role, phone) {
  if (!name) throw new Error("Name is required");
  const db = getDbSheets();
  
  // Check if member already exists
  const members = getMembers();
  const exists = members.some(m => m.Name.toLowerCase().trim() === name.toLowerCase().trim());
  if (exists) throw new Error("Member already exists");
  
  db.members.appendRow([name, email || "", role || "Designer", phone || "", "Active", "1234"]);
  return { success: true, message: `Member ${name} added successfully with default PIN 1234!` };
}

/**
 * Remove or Deactivate a team member
 */
function deleteMember(name) {
  const db = getDbSheets();
  const data = db.members.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      db.members.deleteRow(i + 1);
      return { success: true, message: `Member ${name} deleted successfully!` };
    }
  }
  throw new Error("Member not found");
}

/**
 * Log daily work done by a member (validating PIN, allowing admin override)
 */
function logDailyWork(logData) {
  const isAdmin = String(logData.pin).trim() === ADMIN_PASSWORD;
  if (!isAdmin && !verifyMemberPIN(logData.memberName, logData.pin)) {
    throw new Error("Authentication failed: Invalid PIN. Default PIN is 1234.");
  }

  const db = getDbSheets();
  const dateStr = logData.date || new Date().toISOString().split('T')[0];
  
  db.dailyLogs.appendRow([
    dateStr,
    logData.memberName,
    Number(logData.content6s || 0),
    Number(logData.content10s || 0),
    Number(logData.content15s || 0),
    Number(logData.content20s || 0),
    Number(logData.content30s || 0),
    Number(logData.content1m || 0),
    Number(logData.socialPost || 0),
    logData.note || ""
  ]);
  
  return { success: true, message: "Work logged successfully!" };
}

/**
 * Verify a member's PIN (with Admin override)
 */
function verifyMemberPIN(memberName, pin) {
  if (String(pin).trim() === ADMIN_PASSWORD) return true; // Admin override
  
  const db = getDbSheets();
  const data = db.members.getDataRange().getValues();
  if (data.length <= 1) return false;
  
  const headers = data[0];
  const pinIdx = headers.indexOf("PIN");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memberName) {
      const storedPin = pinIdx === -1 ? "" : String(data[i][pinIdx] || "").trim();
      const expectedPin = storedPin === "" ? "1234" : storedPin;
      return expectedPin === String(pin).trim();
    }
  }
  return false;
}

/**
 * Change member PIN
 */
function changeMemberPIN(memberName, oldPin, newPin) {
  if (!newPin || newPin.length < 4) {
    throw new Error("PIN must be at least 4 digits/characters long");
  }
  
  const db = getDbSheets();
  const data = db.members.getDataRange().getValues();
  const headers = data[0];
  
  let pinIdx = headers.indexOf("PIN");
  if (pinIdx === -1) {
    // If PIN header is missing in Sheet, add it
    db.members.getRange(1, headers.length + 1).setValue("PIN");
    pinIdx = headers.length;
    // Reload format
    db.members.getRange(1, 1, 1, headers.length + 1).setFontWeight("bold");
  }
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memberName) {
      const storedPin = String(data[i][pinIdx] || "").trim();
      const expectedOldPin = storedPin === "" ? "1234" : storedPin;
      
      if (expectedOldPin !== String(oldPin).trim()) {
        throw new Error("Incorrect current PIN");
      }
      
      db.members.getRange(i + 1, pinIdx + 1).setValue(String(newPin).trim());
      return { success: true, message: "PIN changed successfully!" };
    }
  }
  throw new Error("Member not found");
}

/**
 * Fetch daily logs for a specific member
 */
function getDailyLogs(memberName) {
  const db = getDbSheets();
  const data = db.dailyLogs.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const logs = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (memberName && row[1] !== memberName) continue; // Filter by name
    
    const log = {};
    headers.forEach((h, idx) => {
      let val = row[idx];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      log[mapHeaderToKey(h, 'dailyLogs')] = val;
    });
    log.RowIndex = i + 1; // 1-indexed row number
    logs.push(log);
  }
  return logs.reverse(); // Newest first
}

/**
 * Update daily log entry (with Admin override)
 */
function updateDailyLog(rowIndex, logData, memberName, pin) {
  const isAdmin = String(pin).trim() === ADMIN_PASSWORD;
  if (!isAdmin && !verifyMemberPIN(memberName, pin)) {
    throw new Error("Authentication failed: Invalid PIN");
  }
  
  const db = getDbSheets();
  const rowVals = db.dailyLogs.getRange(rowIndex, 1, 1, 10).getValues()[0];
  
  if (!isAdmin && rowVals[1] !== memberName) {
    throw new Error("Permission denied: You can only edit your own logs");
  }
  
  db.dailyLogs.getRange(rowIndex, 1, 1, 10).setValues([[
    logData.date || rowVals[0],
    isAdmin ? (logData.memberName || rowVals[1]) : memberName, // Let admin change member name if needed
    Number(logData.content6s || 0),
    Number(logData.content10s || 0),
    Number(logData.content15s || 0),
    Number(logData.content20s || 0),
    Number(logData.content30s || 0),
    Number(logData.content1m || 0),
    Number(logData.socialPost || 0),
    logData.note || ""
  ]]);
  
  return { success: true, message: "Daily log updated successfully!" };
}

/**
 * Delete daily log entry (with Admin override)
 */
function deleteDailyLog(rowIndex, memberName, pin) {
  const isAdmin = String(pin).trim() === ADMIN_PASSWORD;
  if (!isAdmin && !verifyMemberPIN(memberName, pin)) {
    throw new Error("Authentication failed: Invalid PIN");
  }
  
  const db = getDbSheets();
  const rowVals = db.dailyLogs.getRange(rowIndex, 1, 1, 2).getValues()[0];
  
  if (!isAdmin && rowVals[1] !== memberName) {
    throw new Error("Permission denied: You can only delete your own logs");
  }
  
  db.dailyLogs.deleteRow(rowIndex);
  return { success: true, message: "Daily log deleted successfully!" };
}

/**
 * Fetch all task assignments
 */
function getAssignments() {
  const db = getDbSheets();
  const data = db.assignments.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const assignments = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const formatDateVal = function(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return val ? String(val).trim() : "";
    };
    
    // Hardcoded mapping by column index to prevent header rename bugs
    const task = {
      TaskID: row[0] ? String(row[0]).trim() : "",
      AssignDate: formatDateVal(row[1]),
      MPROID: row[2] ? String(row[2]).trim() : "",
      MPROLink: row[3] ? String(row[3]).trim() : "",
      Group: row[4] ? String(row[4]).trim() : "", // Brand Name
      Quantity: row[5] !== undefined ? Number(row[5]) || 0 : 0,
      CompletedQty: row[6] !== undefined ? Number(row[6]) || 0 : 0,
      Status: row[7] ? String(row[7]).trim() : "Not Started",
      Deadline: row[8] ? String(row[8]).trim() : "",
      AssignedTo: row[9] ? String(row[9]).trim() : "",
      Note: row[10] ? String(row[10]).trim() : "",
      SubmitDate: formatDateVal(row[11]),
      ApproveDate: formatDateVal(row[12]),
      RowIndex: i + 1
    };
    assignments.push(task);
  }
  return assignments;
}

/**
 * Assign a new task to a member
 */
function assignTask(taskData) {
  const db = getDbSheets();
  const taskId = "T-" + Math.floor(100000 + Math.random() * 900000);
  const assignDate = taskData.assignDate || new Date().toISOString().split('T')[0];
  
  db.assignments.appendRow([
    taskId,
    assignDate,
    taskData.mproId || "",
    taskData.mproLink || "",
    taskData.group || "", // stores selected Brand
    Number(taskData.quantity || 1),
    0, // Completed Qty (Col 7) defaults to 0
    taskData.status || "Not Started",
    taskData.deadline || "",
    taskData.assignedTo,
    taskData.note || ""
  ]);
  
  return { success: true, taskId: taskId, message: "Task assigned successfully!" };
}

/**
 * Delete an assigned task from the database
 */
function deleteAssignment(taskId) {
  const db = getDbSheets();
  const data = db.assignments.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      db.assignments.deleteRow(i + 1);
      return { success: true, message: "Task deleted successfully!" };
    }
  }
  throw new Error("Task ID not found");
}

/**
 * Update the status, note, MPRO Link, or Completed Qty of an assigned task
 */
function updateAssignmentStatus(taskId, status, note, mproLink, completedQty) {
  const db = getDbSheets();
  const data = db.assignments.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      const rowIndex = i + 1;
      
      if (status) {
        db.assignments.getRange(rowIndex, 8).setValue(status);
        
        const stLower = status.toLowerCase();
        if (stLower.includes("submit")) {
          db.assignments.getRange(rowIndex, 12).setValue(todayStr); // Submit Date is Col 12
        } else if (stLower.includes("approved") || stLower.includes("ok") || stLower.includes("pass") || stLower.includes("completed")) {
          db.assignments.getRange(rowIndex, 13).setValue(todayStr); // Approve Date is Col 13
          // Force completed qty to total quantity
          const totalQty = Number(data[i][5]) || 1;
          db.assignments.getRange(rowIndex, 7).setValue(totalQty);
        }
      }
      
      if (note !== undefined) db.assignments.getRange(rowIndex, 11).setValue(note);
      if (mproLink) db.assignments.getRange(rowIndex, 4).setValue(mproLink);
      
      if (completedQty !== undefined && (!status || (!status.toLowerCase().includes("approved") && !status.toLowerCase().includes("ok") && !status.toLowerCase().includes("pass") && !status.toLowerCase().includes("completed")))) {
        db.assignments.getRange(rowIndex, 7).setValue(Number(completedQty));
      }
      return { success: true, message: "Task updated successfully!" };
    }
  }
  throw new Error("Task ID not found");
}

/**
 * Load all brands from database
 */
function getBrands() {
  const db = getDbSheets();
  const data = db.brands.getDataRange().getValues();
  const brands = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) brands.push(data[i][0]);
  }
  return brands;
}

/**
 * Add a new brand
 */
function addBrand(brandName) {
  const db = getDbSheets();
  const data = db.brands.getDataRange().getValues();
  const cleanBrand = String(brandName).trim();
  if (!cleanBrand) throw new Error("Brand name cannot be empty");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toLowerCase() === cleanBrand.toLowerCase()) {
      return { success: false, message: "Brand already exists" };
    }
  }
  db.brands.appendRow([cleanBrand]);
  return { success: true, message: "Brand '" + cleanBrand + "' added successfully!" };
}

/**
 * Delete a brand from database
 */
function deleteBrand(brandName) {
  const db = getDbSheets();
  const data = db.brands.getDataRange().getValues();
  const cleanBrand = String(brandName).trim();
  if (!cleanBrand) throw new Error("Brand name cannot be empty");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toLowerCase() === cleanBrand.toLowerCase()) {
      db.brands.deleteRow(i + 1);
      return { success: true, message: "Brand '" + cleanBrand + "' deleted successfully!" };
    }
  }
  return { success: false, message: "Brand not found" };
}

/**
 * Fetch Fund ledger details
 */
function getFundLedger() {
  const db = getDbSheets();
  const data = db.fundTransactions.getDataRange().getValues();
  const ledger = [];
  
  let totalDeposits = 0;
  let totalExpenses = 0;
  
  if (data.length > 1) {
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tx = {};
      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        tx[h.replace(/\s+/g, '')] = val;
      });
      
      const amt = Number(tx.Amount || 0);
      if (tx.Type === "Deposit") {
        totalDeposits += amt;
      } else if (tx.Type === "Expense") {
        totalExpenses += amt;
      }
      ledger.push(tx);
    }
  }
  
  return {
    ledger: ledger.reverse(), // Show newest transactions first
    summary: {
      totalDeposits: totalDeposits,
      totalExpenses: totalExpenses,
      balance: totalDeposits - totalExpenses
    }
  };
}

/**
 * Record a new fund transaction (Deposit or Expense)
 */
function addFundTransaction(txData) {
  const db = getDbSheets();
  const dateStr = txData.date || new Date().toISOString().split('T')[0];
  
  db.fundTransactions.appendRow([
    dateStr,
    txData.type, // Deposit or Expense
    Number(txData.amount || 0),
    txData.type === "Deposit" ? txData.memberName : "",
    txData.note || "",
    txData.status || "Paid"
  ]);
  
  return { success: true, message: "Fund transaction recorded successfully!" };
}

/**
 * Fetch accounts credentials and prompts library
 */
function getAccountsAndPrompts() {
  const db = getDbSheets();
  
  // Read Accounts
  const accData = db.accounts.getDataRange().getValues();
  const accounts = [];
  if (accData.length > 1) {
    const headers = accData[0];
    for (let i = 1; i < accData.length; i++) {
      const row = accData[i];
      const acc = {};
      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        acc[mapHeaderToKey(h, 'accounts')] = val;
      });
      accounts.push(acc);
    }
  }
  
  // Read Prompts
  const prData = db.prompts.getDataRange().getValues();
  const prompts = [];
  if (prData.length > 1) {
    const headers = prData[0];
    for (let i = 1; i < prData.length; i++) {
      const row = prData[i];
      const pr = {};
      headers.forEach((h, idx) => {
        pr[mapHeaderToKey(h, 'prompts')] = row[idx];
      });
      prompts.push(pr);
    }
  }
  
  return {
    accounts: accounts,
    prompts: prompts
  };
}

/**
 * Add shared account credentials
 */
function addAccount(accData) {
  const db = getDbSheets();
  db.accounts.appendRow([
    accData.serviceName,
    accData.loginUrl || "",
    accData.username || "",
    accData.password || "",
    accData.expiryDate || "",
    accData.note || ""
  ]);
  return { success: true, message: "Account credentials added successfully!" };
}

/**
 * Add AI prompt
 */
function addPrompt(promptData) {
  const db = getDbSheets();
  db.prompts.appendRow([
    promptData.title,
    promptData.category || "General",
    promptData.promptText,
    promptData.description || ""
  ]);
  return { success: true, message: "Prompt added successfully!" };
}

/**
 * Verify Admin Password
 */
function verifyAdminPassword(password) {
  return password === ADMIN_PASSWORD;
}

/**
 * Fetch full dashboard data
 */
function getDashboardData() {
  const db = getDbSheets();
  
  // Get members
  const members = getMembers();
  
  // Read all Daily Logs
  const logsData = db.dailyLogs.getDataRange().getValues();
  const logs = [];
  
  let totals = {
    content6s: 0,
    content10s: 0,
    content15s: 0,
    content20s: 0,
    content30s: 0,
    content1m: 0,
    socialPost: 0,
    grandTotal: 0
  };
  
  const memberPerformance = {}; // { MemberName: { content6s, ..., total } }
  
  if (logsData.length > 1) {
    const headers = logsData[0];
    for (let i = 1; i < logsData.length; i++) {
      const row = logsData[i];
      const log = {};
      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        log[mapHeaderToKey(h, 'dailyLogs')] = val;
      });
      
      const c6 = Number(log.Content6s || 0);
      const c10 = Number(log.Content10s || 0);
      const c15 = Number(log.Content15s || 0);
      const c20 = Number(log.Content20s || 0);
      const c30 = Number(log.Content30s || 0);
      const c1m = Number(log.Content1m || 0);
      const sp = Number(log.SocialPost || 0);
      const rowTotal = c6 + c10 + c15 + c20 + c30 + c1m + sp;
      
      totals.content6s += c6;
      totals.content10s += c10;
      totals.content15s += c15;
      totals.content20s += c20;
      totals.content30s += c30;
      totals.content1m += c1m;
      totals.socialPost += sp;
      totals.grandTotal += rowTotal;
      
      // Member specific aggregate
      const mName = log.MemberName;
      if (mName) {
        if (!memberPerformance[mName]) {
          memberPerformance[mName] = { c6s: 0, c10s: 0, c15s: 0, c20s: 0, c30s: 0, c1m: 0, sp: 0, total: 0 };
        }
        memberPerformance[mName].c6s += c6;
        memberPerformance[mName].c10s += c10;
        memberPerformance[mName].c15s += c15;
        memberPerformance[mName].c20s += c20;
        memberPerformance[mName].c30s += c30;
        memberPerformance[mName].c1m += c1m;
        memberPerformance[mName].sp += sp;
        memberPerformance[mName].total += rowTotal;
      }
      
      logs.push(log);
    }
  }
  
  // Get active assignments counts
  const tasks = getAssignments();
  const taskStats = {
    total: tasks.length,
    pending: tasks.filter(t => t.Status === "Not Started").length,
    inProgress: tasks.filter(t => t.Status === "Work in Progress" || t.Status === "Running" || t.Status === "Motion Running").length,
    submitted: tasks.filter(t => t.Status === "Work Submitted").length,
    completed: tasks.filter(t => t.Status === "HOM Approved" || t.Status === "Static Approved").length
  };
  
  // Get fund status
  const fund = getFundLedger();
  
  return {
    members: members,
    totals: totals,
    memberPerformance: memberPerformance,
    taskStats: taskStats,
    fundBalance: fund.summary.balance,
    recentLogs: logs.reverse().slice(0, 10), // Show last 10 logs
    dailyLogs: logs // Return all daily logs to count submitted work
  };
}

/**
 * Migration Script - Runs once to convert old multi-sheet structure
 * into clean standardized relational tables.
 */
function runMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = getDbSheets(); // Initializes sheets and headers
  
  // 1. MIGRATE MEMBERS
  db.members.clearContents();
  db.members.appendRow(["Name", "Email", "Role", "Phone", "Status", "PIN", "Recommended Brands"]);
  
  const memberList = [
    { name: "Bidyut Kumar", email: "designer.bidyut@gmail.com", recBrands: "All" },
    { name: "Emon Kumer", email: "getemonkumar@gmail.com", recBrands: "All" },
    { name: "Md. Maruf Hossain", email: "marufgd2525@gmail.com", recBrands: "Favarit Brand, Dairy, winner, export, good life, latina, junior drinks, But all brand working interested" },
    { name: "Mst. Airen Akter Orni", email: "aa.orni2001@gmail.com", recBrands: "want to work with all brand but want to keep work with Support mat,Hdpe pipe, Icy." },
    { name: "Most. Shumaya Khatun", email: "shumayasimu549@gmail.com", recBrands: "All" },
    { name: "Md. Mahafujur Rahaman", email: "mdmahafujurrahaman11@gmail.com", recBrands: "PRAN,RFL,Export" },
    { name: "Md. Sabbir Ahmed", email: "mssabbir121@gmail.com", recBrands: "as your wish dada. I am ready to do everything," },
    { name: "Md. A. Malek", email: "maleksct@gmail.com", recBrands: "Favarit Brand, Italiano, Dairy, Vision, Click, But all brand working interested" },
    { name: "Pranto Sarkar", email: "prantosarkar32@gmail.com", recBrands: "All" },
    { name: "MD Naimur Rahman", email: "naimursct@gmail.com", recBrands: "All" },
    { name: "Arman Hossain", email: "armansct@gmail.com", recBrands: "I would like do anything but preferable brand: Italiano Melamine, Opal, Dairy UHT Milk, Winner, Mother Touch, Best Buy, Golpo, Mr. Noodles, Sara's Cloud Kitchen, Sara Pizza" },
    { name: "Nurjahan Akter Sumy", email: "nurjahansct@gmail.com", recBrands: "All" },
    { name: "Md Nuruzzaman Jehady", email: "jehadysct@gmail.com", recBrands: "All" }
  ];
  
  memberList.forEach(function(m) {
    db.members.appendRow([m.name, m.email, "Designer", "", "Active", "1234", m.recBrands]);
  });
  
  // 2. MIGRATE ACCOUNTS AND PROMPTS
  const accPromptSheet = ss.getSheetByName("AccountPrompt");
  if (accPromptSheet) {
    const accData = accPromptSheet.getDataRange().getValues();
    db.accounts.clearContents();
    db.accounts.appendRow(["Service Name", "Login URL", "Username/Email", "Password", "Expiry Date", "Note"]);
    
    db.prompts.clearContents();
    db.prompts.appendRow(["Title", "Category", "Prompt Text", "Description"]);
    
    // Columns A-F are Accounts: Name, Email, Password, Buy Date, Taka, Expiry Date
    // Columns I-L are Website logins: Website, Link, User Name, Password
    for (let r = 2; r < accData.length; r++) {
      const row = accData[r];
      // Type 1 Accounts (Col A-F)
      if (row[0] && row[0].trim()) {
        db.accounts.appendRow([
          row[0].trim(),
          "",
          row[1] ? String(row[1]).trim() : "",
          row[2] ? String(row[2]).trim() : "",
          row[5] ? String(row[5]).trim() : "",
          row[4] ? "Taka: " + row[4] : ""
        ]);
      }
      
      // Type 2 Websites (Col I-L)
      if (row[8] && row[8].trim()) {
        db.accounts.appendRow([
          row[8].trim(),
          row[9] ? String(row[9]).trim() : "",
          row[10] ? String(row[10]).trim() : "",
          row[11] ? String(row[11]).trim() : "",
          "",
          ""
        ]);
      }
    }
    
    // Add default prompts if none exist
    db.prompts.appendRow(["Suno Music Generator Prompt", "Suno", "[Genre: Synthpop] Melodic and catchy chorus with heavy 80s drums, energetic vocals.", "Used for 15s content bgm"]);
    db.prompts.appendRow(["Leonardo Premium Artistic Prompt", "Leonardo", "masterpiece, ultra detailed, cinematic lighting, 8k resolution, photorealistic", "Used for high quality static posters"]);
  }
  
  // 3. MIGRATE TEAM ASSIGNMENTS
  const teamSheet = ss.getSheetByName("Team");
  if (teamSheet) {
    const teamData = teamSheet.getDataRange().getValues();
    db.assignments.clearContents();
    db.assignments.appendRow(["Task ID", "Assign Date", "MPRO ID", "MPRO Link", "Group", "Quantity", "Completed Qty", "Status", "Deadline", "Assigned To", "Note"]);
    
    // Columns are grouped horizontally:
    // Col 3-10: Bidyut (Assign Date, MPRO ID, MPRO Link, GROUP, Quantity, Status, Note, Deadline)
    // Col 11-17: Emon
    // Col 18-24: Maruf
    // Col 25-32: Orni
    // and so on...
    const memberColumns = [
      { name: "Bidyut Kumar", startIdx: 3, width: 8 },
      { name: "Emon Kumer", startIdx: 11, width: 7 },
      { name: "Maruf Hossain", startIdx: 18, width: 7 },
      { name: "Airen Akter Orni", startIdx: 25, width: 8 },
      { name: "Shumaya Khatun", startIdx: 33, width: 7 },
      { name: "Mahafujur Rahaman", startIdx: 40, width: 8 },
      { name: "Sabbir Ahmed", startIdx: 48, width: 7 },
      { name: "A.Malek", startIdx: 55, width: 8 },
      { name: "Pranto Sarkar", startIdx: 63, width: 7 },
      { name: "MD Naimur Rahman", startIdx: 70, width: 7 },
      { name: "Arman Hossain", startIdx: 77, width: 7 }
    ];
    
    let assignCount = 0;
    
    memberColumns.forEach(member => {
      for (let r = 2; r < teamData.length; r++) {
        const row = teamData[r];
        const mproId = row[member.startIdx + 1];
        if (mproId && String(mproId).trim()) {
          assignCount++;
          const taskId = "T-" + (100000 + assignCount);
          
          let assignDate = row[member.startIdx];
          if (assignDate instanceof Date) {
            assignDate = Utilities.formatDate(assignDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          
          let deadline = "";
          let note = "";
          
          // Width based offsets
          if (member.width === 8) {
            note = row[member.startIdx + 6] ? String(row[member.startIdx + 6]).trim() : "";
            deadline = row[member.startIdx + 7] ? String(row[member.startIdx + 7]).trim() : "";
          } else {
            deadline = row[member.startIdx + 6] ? String(row[member.startIdx + 6]).trim() : "";
          }
          
          const statusVal = row[member.startIdx + 5] ? String(row[member.startIdx + 5]).trim() : "Not Started";
          const qtyVal = row[member.startIdx + 4] ? Number(row[member.startIdx + 4]) || 1 : 1;
          
          // Estimate completed qty based on status
          const isCompleted = (statusVal.toLowerCase().includes("approved") || statusVal.toLowerCase().includes("completed") || statusVal.toLowerCase().includes("pass") || statusVal.toLowerCase().includes("ok"));
          const completedQtyVal = isCompleted ? qtyVal : 0;
          
          db.assignments.appendRow([
            taskId,
            assignDate || "",
            String(mproId).trim(),
            row[member.startIdx + 2] ? String(row[member.startIdx + 2]).trim() : "",
            row[member.startIdx + 3] ? String(row[member.startIdx + 3]).trim() : "",
            qtyVal,
            completedQtyVal, // Completed Qty
            statusVal,
            deadline,
            member.name,
            note
          ]);
        }
      }
    });
  }
  
  // 4. MIGRATE FUND TRANSACTIONS
  const fundSheet = ss.getSheetByName("Fund");
  if (fundSheet) {
    const fundData = fundSheet.getDataRange().getValues();
    db.fundTransactions.clearContents();
    db.fundTransactions.appendRow(["Date", "Type", "Amount", "Member Name", "Note", "Status"]);
    
    for (let r = 1; r < fundData.length; r++) {
      const row = fundData[r];
      let dateVal = row[0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      // Deposits (Col B is Deposit amount, Col C is note, Col D is Receivable Name, Col E is Receivable Amount)
      if (row[1] && Number(row[1]) > 0) {
        db.fundTransactions.appendRow([
          dateVal || "",
          "Deposit",
          Number(row[1]),
          "", // General deposit
          row[2] ? String(row[2]).trim() : "",
          "Paid"
        ]);
      }
      
      // Receivable (Col D, E)
      if (row[3] && row[4] && Number(row[4]) > 0) {
        db.fundTransactions.appendRow([
          dateVal || "",
          "Deposit",
          Number(row[4]),
          String(row[3]).trim(),
          "Receivable payment",
          "Paid"
        ]);
      }
      
      // Expenses (Col F is amount, Col G is Color/Note, Col H is Name/Details, Col I is status)
      if (row[5] && Number(row[5]) > 0) {
        db.fundTransactions.appendRow([
          dateVal || "",
          "Expense",
          Number(row[5]),
          "",
          (row[6] ? String(row[6]).trim() : "") + (row[7] ? " (By " + row[7] + ")" : ""),
          row[8] ? String(row[8]).trim() : "Paid"
        ]);
      }
    }
  }
  
  // 5. MIGRATE MEMBER LOGS (DailyWorkLogs)
  db.dailyLogs.clearContents();
  db.dailyLogs.appendRow(["Date", "Member Name", "6s Content", "10s Content", "15s Content", "20s Content", "30s Content", "1m Content", "Social Post", "Note"]);
  
  const memberSheets = [
    "Bidyut Kumar", "Emon Kumer", "A.Malek", "Airen Akter Orni", "Maruf Hossain",
    "Shumaya Khatun", "Mahafujur Rahaman", "Sabbir Ahmed", "Pranto Sarkar",
    "MD Naimur Rahman", "Arman Hossain ", "Nurjahan Akter Sumy", "Md Nuruzzaman Jehady"
  ];
  
  memberSheets.forEach(mSheetName => {
    const cleanSheetName = mSheetName.trim();
    const sheet = ss.getSheetByName(cleanSheetName) || ss.getSheetByName(mSheetName);
    
    const getCleanMemberName = function(sName) {
      if (sName === "A.Malek") return "Md. A. Malek";
      if (sName === "Maruf Hossain") return "Md. Maruf Hossain";
      if (sName === "Shumaya Khatun") return "Most. Shumaya Khatun";
      return sName;
    };
    
    if (sheet) {
      const sData = sheet.getDataRange().getValues();
      if (sData.length > 2) {
        // Row 2 is columns list: e.g. Date, 6 Sec, 10 Sec, Total, etc.
        const cols = sData[1].map(c => c ? String(c).toLowerCase() : "");
        
        // Find positions of headers
        const pos = {
          date: cols.indexOf("date"),
          c6s: cols.findIndex(c => c.includes("6 sec") || c.includes("6sec")),
          c10s: cols.findIndex(c => c.includes("10 sec") || c.includes("10+ sec")),
          c15s: cols.findIndex(c => c.includes("15 sec") || c.includes("15+ sec")),
          c20s: cols.findIndex(c => c.includes("20 sec") || c.includes("20+ sec")),
          c30s: cols.findIndex(c => c.includes("30 sec") || c.includes("30+ sec")),
          c1m: cols.findIndex(c => c.includes("1min") || c.includes("1 min") || c.includes("1m")),
          sp: cols.findIndex(c => c.includes("static") || c.includes("social post") || c.includes("socialpost"))
        };
        
        for (let r = 2; r < sData.length; r++) {
          const row = sData[r];
          let dVal = row[pos.date];
          if (dVal instanceof Date) {
            dVal = Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          
          if (dVal) {
            // Check if there is any logged content on this day
            const q6 = pos.c6s !== -1 ? Number(row[pos.c6s]) || 0 : 0;
            const q10 = pos.c10s !== -1 ? Number(row[pos.c10s]) || 0 : 0;
            const q15 = pos.c15s !== -1 ? Number(row[pos.c15s]) || 0 : 0;
            const q20 = pos.c20s !== -1 ? Number(row[pos.c20s]) || 0 : 0;
            const q30 = pos.c30s !== -1 ? Number(row[pos.c30s]) || 0 : 0;
            const q1m = pos.c1m !== -1 ? Number(row[pos.c1m]) || 0 : 0;
            const qsp = pos.sp !== -1 ? Number(row[pos.sp]) || 0 : 0;
            
            if (q6 + q10 + q15 + q20 + q30 + q1m + qsp > 0) {
              db.dailyLogs.appendRow([
                dVal,
                getCleanMemberName(cleanSheetName),
                q6, q10, q15, q20, q30, q1m, qsp,
                "" // note
              ]);
            }
          }
        }
      }
    }
  });
  
  // Format tables nicely
  setupHeaders(db);
  
  return { success: true, message: "Migration completed successfully! All offline data structured." };
}

/**
 * Import accounts from JSON array
 */
function importAccounts(accounts) {
  if (!accounts || !Array.isArray(accounts)) {
    throw new Error("Invalid input: Must be an array of accounts");
  }
  
  const db = getDbSheets();
  let accountsAdded = 0;
  
  db.accounts.clearContents();
  db.accounts.appendRow(["Service Name", "Login URL", "Username/Email", "Password", "Expiry Date", "Note"]);
  
  accounts.forEach(acc => {
    db.accounts.appendRow([
      acc.serviceName || "",
      acc.loginUrl || "",
      acc.username || "",
      acc.password || "",
      acc.expiryDate || "",
      acc.note || ""
    ]);
    accountsAdded++;
  });
  
  // Format headers
  setupHeaders(db);
  
  return { success: true, message: `Successfully imported ${accountsAdded} accounts!` };
}

/**
 * Import prompts from JSON array
 */
function importPrompts(prompts) {
  if (!prompts || !Array.isArray(prompts)) {
    throw new Error("Invalid input: Must be an array of prompts");
  }
  
  const db = getDbSheets();
  let promptsAdded = 0;
  
  db.prompts.clearContents();
  db.prompts.appendRow(["Title", "Category", "Prompt Text", "Description"]);
  
  prompts.forEach(pr => {
    db.prompts.appendRow([
      pr.title || "",
      pr.category || "General",
      pr.promptText || "",
      pr.description || ""
    ]);
    promptsAdded++;
  });
  
  // Format headers
  setupHeaders(db);
  
  return { success: true, message: `Successfully imported ${promptsAdded} prompts!` };
}

/**
 * Custom Menu inside Google Sheets
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Team Portal Tools")
      .addItem("Create / Reset Dashboard V2", "createResetDashboardV2")
      .addItem("Compile Member Logs to DailyLogs", "compileAndMergeMemberLogs")
      .addItem("Migrate Sheets Data to Supabase", "migrateAllDataToSupabase")
      .addToUi();
  } catch(e) {
    Logger.log("Not running in spreadsheet context: " + e.message);
  }
}

/**
 * Compiles all individual member daily log sheets into the centralized 'DailyLogs' sheet.
 */
function compileAndMergeMemberLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    "Compile Member Logs",
    "Are you sure you want to compile and merge all member worksheets into the 'DailyLogs' worksheet? This will clear existing rows in 'DailyLogs'.",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const dailyLogsSheet = ss.getSheetByName("DailyLogs") || ss.insertSheet("DailyLogs");
  
  // Clear all data except headers (row 1)
  const lastRow = dailyLogsSheet.getLastRow();
  if (lastRow > 1) {
    dailyLogsSheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  }
  
  // Set headers just in case
  dailyLogsSheet.getRange(1, 1, 1, 10).setValues([[
    "Date", "Member Name", "6s Content", "10s Content", "15s Content", "20s Content", "30s Content", "1m Content", "Social Post", "Note"
  ]]);
  
  const sheets = ss.getSheets();
  const compiledRows = [];
  
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    
    // Skip system sheets
    if (["Dashboard", "Members", "DailyLogs", "Assignments", "FundTransactions", "Accounts", "Prompts", "Brands", "Incentive", "Information"].indexOf(sheetName) !== -1) {
      continue;
    }
    
    // Check if it matches a member sheet structure
    const hasStructure = sheet.getLastRow() >= 3 &&
                         sheet.getLastColumn() >= 8 &&
                         String(sheet.getRange(3, 1).getValue()).toLowerCase().indexOf("date") !== -1 &&
                         String(sheet.getRange(3, 2).getValue()).toLowerCase().indexOf("10+") !== -1;
                         
    if (!hasStructure) continue;
    
    // Member Name from Cell A1
    let memberName = String(sheet.getRange(1, 1).getValue()).trim();
    if (!memberName) {
      memberName = sheetName; // fallback to sheet name
    }
    
    const rangeData = sheet.getRange(4, 1, sheet.getLastRow() - 3, 8).getValues();
    
    for (let r = 0; r < rangeData.length; r++) {
      const row = rangeData[r];
      const dateVal = row[0];
      const total = Number(row[7]) || 0;
      
      // Only compile if there is a valid date and work was actually done
      if (dateVal && total > 0) {
        let formattedDate = dateVal;
        if (dateVal instanceof Date) {
          formattedDate = Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        }
        
        compiledRows.push([
          formattedDate,
          memberName,
          0, // 6s Content
          Number(row[1]) || 0, // 10s Content
          Number(row[2]) || 0, // 15s Content
          Number(row[3]) || 0, // 20s Content
          Number(row[4]) || 0, // 30s Content
          Number(row[5]) || 0, // 1m Content
          Number(row[6]) || 0, // Social Post (Static)
          "" // Note
        ]);
      }
    }
  }
  
  if (compiledRows.length > 0) {
    dailyLogsSheet.getRange(2, 1, compiledRows.length, 10).setValues(compiledRows);
    ui.alert("Success", "Successfully compiled " + compiledRows.length + " entries from member sheets into the 'DailyLogs' sheet.\n\nYou can now click 'Migrate Sheets Data to Supabase' to sync this data with the database!", ui.ButtonSet.OK);
  } else {
    ui.alert("Information", "No active work entries found in member worksheets to compile.", ui.ButtonSet.OK);
  }
}

/**
 * Recreate or reset the Dashboard worksheet tab with correct styles and formula mapping
 */
function createResetDashboardV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard");
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet("Dashboard");
  }
  
  // Set active
  ss.setActiveSheet(sheet);
  
  // Apply gridlines (show gridlines by setting hidden to false)
  sheet.setHiddenGridlines(false);
  
  // 1. Header Title Banner (Rows 1 & 2)
  sheet.getRange("A1:J2").merge();
  const titleCell = sheet.getRange("A1");
  titleCell.setValue("TEAM BIDYUT CONTENT PORTAL & DASHBOARD");
  titleCell.setFontFamily("Outfit");
  titleCell.setFontSize(18);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  titleCell.setBackground("#1b2a4a");
  titleCell.setFontColor("#ffffff");
  
  // 2. Summary Metric Cards (Rows 4 & 5)
  // Total Members Card (A4:B5)
  sheet.getRange("A4").setValue("Total Members");
  sheet.getRange("A5").setFormula("=COUNTA(Members!A2:A)");
  
  // Total Content Items Card (C4:D5)
  sheet.getRange("C4").setValue("Total Content Items");
  sheet.getRange("C5").setFormula("=SUM(DailyLogs!C2:I)");
  
  // Total Tasks Assigned Card (E4:F5)
  sheet.getRange("E4").setValue("Total Tasks Assigned");
  sheet.getRange("E5").setFormula("=COUNTA(Assignments!A2:A)");
  
  // Completed Tasks Card (G4:H5)
  sheet.getRange("G4").setValue("Completed Tasks");
  sheet.getRange("G5").setFormula("=COUNTIF(Assignments!H2:H, \"*Approved*\") + COUNTIF(Assignments!H2:H, \"*Completed*\") + COUNTIF(Assignments!H2:H, \"*Pass*\") + COUNTIF(Assignments!H2:H, \"*Ok*\")");
  
  // Current Fund Balance Card (I4:J5)
  sheet.getRange("I4").setValue("Current Fund Balance");
  sheet.getRange("I5").setFormula("=SUMIF(FundTransactions!B2:B, \"Deposit\", FundTransactions!C2:C) - SUMIF(FundTransactions!B2:B, \"Expense\", FundTransactions!C2:C)");
  
  // Style Summary Cards
  const cardsHeaderRange = sheet.getRange("A4:J4");
  cardsHeaderRange.setFontWeight("bold");
  cardsHeaderRange.setFontSize(9);
  cardsHeaderRange.setHorizontalAlignment("center");
  cardsHeaderRange.setFontColor("#7f8c8d");
  
  const cardsValueRange = sheet.getRange("A5:J5");
  cardsValueRange.setFontWeight("bold");
  cardsValueRange.setFontSize(16);
  cardsValueRange.setHorizontalAlignment("center");
  cardsValueRange.setFontColor("#2c3e50");
  
  // Format Currency for Fund Balance
  sheet.getRange("I5").setNumberFormat("$#,##0.00");
  
  // 3. MEMBER PERFORMANCE SUMMARY title (Row 7)
  sheet.getRange("A7:J7").merge();
  const perfTitle = sheet.getRange("A7");
  perfTitle.setValue("MEMBER PERFORMANCE SUMMARY (ROLLING LOGS)");
  perfTitle.setFontWeight("bold");
  perfTitle.setBackground("#2c3e50");
  perfTitle.setFontColor("#ffffff");
  perfTitle.setHorizontalAlignment("left");
  
  // 4. Performance Table Headers (Row 8)
  const tableHeaders = ["Name", "6s", "10s", "15s", "20s", "30s", "1m", "Social Post", "Total Items", "Points"];
  tableHeaders.forEach((h, idx) => {
    sheet.getRange(8, idx + 1).setValue(h);
  });
  
  const tableHeaderRange = sheet.getRange("A8:J8");
  tableHeaderRange.setFontWeight("bold");
  tableHeaderRange.setBackground("#ecf0f1");
  tableHeaderRange.setHorizontalAlignment("center");
  
  // 5. Populate Members & Formulas (Rows 9 to 21)
  const members = [
    "Bidyut Kumar", "Emon Kumer", "Md. Maruf Hossain", "Mst. Airen Akter Orni", 
    "Most. Shumaya Khatun", "Md. Mahafujur Rahaman", "Md. Sabbir Ahmed", 
    "Md. A. Malek", "Pranto Sarkar", "MD Naimur Rahman", 
    "Arman Hossain", "Nurjahan Akter Sumy", "Md Nuruzzaman Jehady"
  ];
  
  members.forEach((m, idx) => {
    const rowNum = 9 + idx;
    sheet.getRange(rowNum, 1).setValue(m);
    
    // Formula columns B to H (6s to Social Post) referencing DailyLogs sheet
    sheet.getRange(rowNum, 2).setFormula(`=SUMIFS(DailyLogs!C:C, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 3).setFormula(`=SUMIFS(DailyLogs!D:D, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 4).setFormula(`=SUMIFS(DailyLogs!E:E, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 5).setFormula(`=SUMIFS(DailyLogs!F:F, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 6).setFormula(`=SUMIFS(DailyLogs!G:G, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 7).setFormula(`=SUMIFS(DailyLogs!H:H, DailyLogs!B:B, A${rowNum})`);
    sheet.getRange(rowNum, 8).setFormula(`=SUMIFS(DailyLogs!I:I, DailyLogs!B:B, A${rowNum})`);
    
    // Total Items: SUM(B:H)
    sheet.getRange(rowNum, 9).setFormula(`=SUM(B${rowNum}:H${rowNum})`);
    
    // Points formula: (6s*1)+(10s*1)+(15s*1.7)+(20s*1.7)+(30s*2)+(1m*3)+(SocialPost*0.5)
    sheet.getRange(rowNum, 10).setFormula(`=(B${rowNum}*1)+(C${rowNum}*1)+(D${rowNum}*1.7)+(E${rowNum}*1.7)+(F${rowNum}*2)+(G${rowNum}*3)+(H${rowNum}*0.5)`);
  });
  
  // Format numbers in table
  sheet.getRange(9, 2, members.length, 8).setNumberFormat("0");
  sheet.getRange(9, 10, members.length, 1).setNumberFormat("0.0");
  
  // 6. OPERATIONAL BREAKDOWNS title (Row 23)
  sheet.getRange("A23:J23").merge();
  const opTitle = sheet.getRange("A23");
  opTitle.setValue("OPERATIONAL BREAKDOWNS (TASKS & FUNDS)");
  opTitle.setFontWeight("bold");
  opTitle.setBackground("#2c3e50");
  opTitle.setFontColor("#ffffff");
  
  // 7. Task Status Breakdown Table (A25:B30)
  sheet.getRange("A25").setValue("Task Status");
  sheet.getRange("B25").setValue("Count");
  sheet.getRange("A25:B25").setFontWeight("bold").setBackground("#ecf0f1").setHorizontalAlignment("center");
  
  sheet.getRange("A26").setValue("Not Started");
  sheet.getRange("B26").setFormula('=COUNTIF(Assignments!H2:H, "Not Started")');
  
  sheet.getRange("A27").setValue("In Progress");
  sheet.getRange("B27").setFormula('=COUNTIF(Assignments!H2:H, "Work in Progress") + COUNTIF(Assignments!H2:H, "Running")');
  
  sheet.getRange("A28").setValue("Submitted");
  sheet.getRange("B28").setFormula('=COUNTIF(Assignments!H2:H, "Work Submitted") + COUNTIF(Assignments!H2:H, "Submitted")');
  
  sheet.getRange("A29").setValue("Completed");
  sheet.getRange("B29").setFormula('=COUNTIF(Assignments!H2:H, "*Approved*") + COUNTIF(Assignments!H2:H, "*Completed*") + COUNTIF(Assignments!H2:H, "*Pass*") + COUNTIF(Assignments!H2:H, "*Ok*")');
  
  sheet.getRange("A30").setValue("Total Tasks");
  sheet.getRange("B30").setFormula('=SUM(B26:B29)');
  sheet.getRange("A30:B30").setFontWeight("bold");
  
  // 8. Fund Category Table (D25:E28)
  sheet.getRange("D25").setValue("Fund Category");
  sheet.getRange("E25").setValue("Total Amount");
  sheet.getRange("D25:E25").setFontWeight("bold").setBackground("#ecf0f1").setHorizontalAlignment("center");
  
  sheet.getRange("D26").setValue("Total Deposit");
  sheet.getRange("E26").setFormula('=SUMIF(FundTransactions!B2:B, "Deposit", FundTransactions!C2:C)');
  
  sheet.getRange("D27").setValue("Total Expense");
  sheet.getRange("E27").setFormula('=SUMIF(FundTransactions!B2:B, "Expense", FundTransactions!C2:C)');
  
  sheet.getRange("D28").setValue("Net Balance");
  sheet.getRange("E28").setFormula('=E26-E27');
  sheet.getRange("D28:E28").setFontWeight("bold");
  
  // Format currency for Fund Table
  sheet.getRange("E26:E28").setNumberFormat("$#,##0.00");
  
  // Style and autofit columns
  sheet.autoResizeColumns(1, 10);
  
  try {
    SpreadsheetApp.getUi().alert("Dashboard V2 has been successfully created and formatted!");
  } catch(e) {}
}

/**
 * Migrate all Google Sheets data to Supabase Database
 */
function migrateAllDataToSupabase() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Migrate to Supabase",
    "Are you sure you want to migrate all Google Sheets data to Supabase? Existing data in Supabase tables will be updated/inserted.",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const supabaseUrl = "https://zyeuvkqusnpuvjqyynbf.supabase.co";
  const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5ZXV2a3F1c25wdXZqcXl5bmJmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQxMDM5MjIsImV4cCI6MjA5OTY3OTkyMn0.Yz8jB4_2YMddwuejpSpldhp_2LbJjijqG2A6K8wr3EI";
  
  const headers = {
    "apikey": supabaseKey,
    "Authorization": "Bearer " + supabaseKey,
    "Content-Type": "application/json",
    "Prefer": "resolution=merge-duplicates"
  };
  
  let statusLog = [];
  
  // Helper function to post to Supabase with automatic conflict resolution parameter for upserting
  function postToSupabase(table, payload) {
    let conflictParam = "";
    if (table === "members") conflictParam = "?on_conflict=name";
    else if (table === "assignments") conflictParam = "?on_conflict=mpro_id";
    else if (table === "settings") conflictParam = "?on_conflict=key";
    else if (table === "daily_logs") conflictParam = "?on_conflict=id";
    else if (table === "fund_transactions") conflictParam = "?on_conflict=id";
    else if (table === "accounts") conflictParam = "?on_conflict=id";
    else if (table === "prompts") conflictParam = "?on_conflict=id";
    else if (table === "brands") conflictParam = "?on_conflict=name";
    
    const url = supabaseUrl + "/rest/v1/" + table + conflictParam;
    const options = {
      method: "post",
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return { success: true };
    } else {
      return { success: false, error: res.getContentText() };
    }
  }
  
  const db = getDbSheets();
  
  // 1. Migrate Members
  try {
    const membersData = db.members.getDataRange().getValues();
    if (membersData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < membersData.length; i++) {
        const r = membersData[i];
        if (!r[0]) continue;
        dbRows.push({
          name: String(r[0]).trim(),
          email: String(r[1] || "").trim(),
          role: String(r[2] || "").trim(),
          phone: String(r[3] || "").trim(),
          status: String(r[4] || "").trim(),
          pin: String(r[5] || "").trim(),
          recommended_brands: String(r[6] || "").trim()
        });
      }
      const res = postToSupabase("members", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " members.");
      } else {
        statusLog.push("❌ Failed members migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Members error: " + e.message);
  }
  
  // 2. Migrate Daily Logs
  try {
    const logsData = db.dailyLogs.getDataRange().getValues();
    if (logsData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < logsData.length; i++) {
        const r = logsData[i];
        if (!r[0]) continue;
        let dateVal = r[0];
        if (dateVal instanceof Date) {
          dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        dbRows.push({
          date: dateVal,
          member_name: String(r[1] || "").trim(),
          content_6s: Number(r[2]) || 0,
          content_10s: Number(r[3]) || 0,
          content_15s: Number(r[4]) || 0,
          content_20s: Number(r[5]) || 0,
          content_30s: Number(r[6]) || 0,
          content_1m: Number(r[7]) || 0,
          social_post: Number(r[8]) || 0,
          note: String(r[9] || "").trim()
        });
      }
      const res = postToSupabase("daily_logs", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " daily logs.");
      } else {
        statusLog.push("❌ Failed daily logs migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Daily logs error: " + e.message);
  }
  
  // 3. Migrate Assignments
  try {
    const assignData = db.assignments.getDataRange().getValues();
    if (assignData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < assignData.length; i++) {
        const r = assignData[i];
        if (!r[0] && !r[2]) continue; // Skip empty rows
        
        let mproId = r[2] ? String(r[2]).trim() : String(r[0]).trim();
        if (!mproId) continue;
        
        let assignD = r[1];
        if (assignD instanceof Date) {
          assignD = Utilities.formatDate(assignD, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          assignD = assignD ? String(assignD).trim() : new Date().toISOString().split('T')[0];
        }
        
        let submitD = r[11];
        if (submitD instanceof Date) {
          submitD = Utilities.formatDate(submitD, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          submitD = submitD ? String(submitD).trim() : null;
        }
        
        let approveD = r[12];
        if (approveD instanceof Date) {
          approveD = Utilities.formatDate(approveD, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          approveD = approveD ? String(approveD).trim() : null;
        }
        
        dbRows.push({
          task_id: String(r[0]).trim(),
          mpro_id: mproId,
          assign_date: assignD,
          brands: String(r[4] || "").trim(),
          member_name: String(r[9] || "").trim(),
          quantity: Number(r[5]) || 0,
          completed_qty: Number(r[6]) || 0,
          status: String(r[7] || "Not Started").trim(),
          submit_date: submitD,
          approve_date: approveD
        });
      }
      const res = postToSupabase("assignments", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " assignments.");
      } else {
        statusLog.push("❌ Failed assignments migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Assignments error: " + e.message);
  }
  
  // 4. Migrate Fund Transactions
  try {
    const fundData = db.fundTransactions.getDataRange().getValues();
    if (fundData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < fundData.length; i++) {
        const r = fundData[i];
        if (!r[0]) continue;
        let dateVal = r[0];
        if (dateVal instanceof Date) {
          dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        dbRows.push({
          date: dateVal,
          type: String(r[1] || "").trim(),
          amount: Number(r[2]) || 0,
          category: String(r[3] || "").trim(),
          description: String(r[4] || "").trim()
        });
      }
      const res = postToSupabase("fund_transactions", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " fund transactions.");
      } else {
        statusLog.push("❌ Failed fund transactions migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Fund error: " + e.message);
  }
  
  // 5. Migrate Accounts
  try {
    const accountsData = db.accounts.getDataRange().getValues();
    if (accountsData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < accountsData.length; i++) {
        const r = accountsData[i];
        if (!r[0]) continue;
        dbRows.push({
          service_name: String(r[0] || "").trim(),
          login_url: String(r[1] || "").trim(),
          username: String(r[2] || "").trim(),
          password: String(r[3] || "").trim(),
          expiry_date: String(r[4] || "").trim(),
          note: String(r[5] || "").trim()
        });
      }
      const res = postToSupabase("accounts", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " premium accounts.");
      } else {
        statusLog.push("❌ Failed accounts migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Accounts error: " + e.message);
  }
  
  // 6. Migrate Prompts
  try {
    const promptsData = db.prompts.getDataRange().getValues();
    if (promptsData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < promptsData.length; i++) {
        const r = promptsData[i];
        if (!r[0]) continue;
        dbRows.push({
          title: String(r[0] || "").trim(),
          category: String(r[1] || "").trim(),
          prompt_text: String(r[2] || "").trim(),
          description: String(r[3] || "").trim()
        });
      }
      const res = postToSupabase("prompts", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " prompts.");
      } else {
        statusLog.push("❌ Failed prompts migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Prompts error: " + e.message);
  }
  
  // 7. Migrate Brands
  try {
    const brandsData = db.brands.getDataRange().getValues();
    if (brandsData.length > 1) {
      const dbRows = [];
      for (let i = 1; i < brandsData.length; i++) {
        const r = brandsData[i];
        if (!r[0]) continue;
        dbRows.push({
          name: String(r[0]).trim()
        });
      }
      const res = postToSupabase("brands", dbRows);
      if (res.success) {
        statusLog.push("✅ Migrated " + dbRows.length + " brands.");
      } else {
        statusLog.push("❌ Failed brands migration: " + res.error);
      }
    }
  } catch (e) {
    statusLog.push("❌ Brands error: " + e.message);
  }
  
  ui.alert("Migration Complete", statusLog.join("\n"), ui.ButtonSet.OK);
}


