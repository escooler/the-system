/**
 * Marketing Team Points System - Version 12.2
 * Fixed: Buffer no longer shows as line item in manifest
 * Fixed: Total capacity formula now correctly references team net capacity
 * Removed: Confusing "Available for Allocation" field
 * Improved: Source tracking (PMM vs Workstream vs Team)
 */

// ==================== CONSTANTS ====================
const VERSION = "12.2";
const RELEASE_DATE = "2024-12-20";

const CONFIG = {
  DEFAULT_WORKSTREAMS: ['SoMe', 'PUA', 'ASO', 'Portal'],
  DEFAULT_ALLOCATIONS: [0.50, 0.20, 0.05, 0.25],
  DEFAULT_CAPACITY: 100,
  DEFAULT_BUFFER_PERCENT: 0.10, // 10% default buffer
  TSHIRT_SIZES: {
    'XS': 1,
    'S': 3,
    'M': 5,
    'L': 13,
    'XL': 21
  },
  COLORS: {
    HEADER_BLUE: '#4285F4',
    HEADER_GREEN: '#34A853',
    HEADER_ORANGE: '#FF9800',
    HEADER_PURPLE: '#9C27B0',
    LIGHT_BLUE: '#E8F0FE',
    LIGHT_GREEN: '#E8F5E9',
    LIGHT_ORANGE: '#FFE0B2',
    LIGHT_YELLOW: '#FFF3E0',
    LIGHT_PURPLE: '#F3E5F5',
    PMM_BLUE: '#CDDFF9',
    GRAY: '#F0F0F0'
  },
  MONTHS: [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ]
};

// ==================== MAIN SETUP ====================
function setupPointsSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear all sheets except first
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // Setup Allocation tab
  const allocationSheet = sheets[0];
  allocationSheet.clear();
  allocationSheet.setName('Allocation');
  setupAllocationTab(allocationSheet);
  
  // Create workstream tabs
  CONFIG.DEFAULT_WORKSTREAMS.forEach(name => {
    const wsSheet = ss.insertSheet(name);
    setupWorkstreamTab(wsSheet, name);
  });
  
  // Create default Creative team
  const teamSheet = ss.insertSheet('Creative Team');
  setupTeamTab(teamSheet, 'Creative');
  
  // Update team dropdowns after creating the default team
  updateTeamDropdowns();
  
  // Update the total capacity formula after creating teams
  updateTotalCapacity();
  
  ss.setActiveSheet(allocationSheet);
  
  SpreadsheetApp.getUi().alert(
    'Points System Setup Complete! 🎉',
    `System v${VERSION} ready with auto-calculated capacity from teams.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==================== ALLOCATION TAB ====================
function setupAllocationTab(sheet) {
  // Set column widths
  const widths = [250, 120, 120, 30, 250, 100, 80, 80, 80, 80];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Title with month reference
  setCell(sheet, 'A1:C1', '=CONCATENATE("ALLOCATION TAB - ",C3," Planning")', {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_BLUE, fontColor: '#FFFFFF'
  });
  
  setCell(sheet, 'A2:C2', 'Monthly Planning & Resource Allocation', {
    merge: true, fontSize: 11, background: CONFIG.COLORS.LIGHT_BLUE
  });
  
  // Add Month Selector
  setCell(sheet, 'A3', 'Planning Month:', {
    fontWeight: true
  });
  
  // Create month dropdown in C3
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();
  const monthValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.MONTHS, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange('C3').setDataValidation(monthValidation)
    .setValue(CONFIG.MONTHS[currentMonth])
    .setBackground(CONFIG.COLORS.LIGHT_YELLOW)
    .setFontWeight(true);
  
  // Add Year input
  setCell(sheet, 'D3', 'Year:');
  setCell(sheet, 'E3', currentYear, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
  });
  
  // Monthly Setup
  setCell(sheet, 'A5', 'MONTHLY SETUP', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  setCell(sheet, 'A6', 'Total Creative Capacity (Points):');
  
  // Auto-sum from all teams' net capacity (D7)
  const teamFormula = generateTeamCapacityFormula();
  setCell(sheet, 'B6', teamFormula, {
    background: '#E8F5E9', border: true, format: '0', fontWeight: true
  });
  
  // Workstream Allocation
  setCell(sheet, 'A8', 'WORKSTREAM ALLOCATION', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  
  sheet.getRange('A9:C9').setValues([['Workstream', 'Allocation %', 'Points']])
    .setFontWeight(true).setBackground('#E3F2FD');
  
  // Workstreams
  CONFIG.DEFAULT_WORKSTREAMS.forEach((ws, i) => {
    const row = 10 + i;
    setCell(sheet, `A${row}`, ws);
    setCell(sheet, `B${row}`, CONFIG.DEFAULT_ALLOCATIONS[i], {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    setCell(sheet, `C${row}`, `=ROUND($B$6*B${row},0)`, {
      format: '0', background: '#F5F5F5'
    });
  });
  
  // Total row
  const totalRow = 14;
  ['TOTAL', '=SUM(B10:B13)', '=SUM(C10:C13)'].forEach((val, i) => {
    setCell(sheet, `${String.fromCharCode(65 + i)}${totalRow}`, val, {
      fontWeight: true, background: '#E0E0E0',
      format: i > 0 ? (i === 1 ? '0%' : '0') : null
    });
  });
  
  sheet.getRange(9, 1, 6, 3).setBorder(true, true, true, true, true, true);
  
  // Strategic Priorities
  setCell(sheet, 'E5', 'STRATEGIC PRIORITIES', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  
  sheet.getRange('E6:J6').setValues([['Priority Name', 'Weight %', 'SoMe', 'PUA', 'ASO', 'Portal']])
    .setFontWeight(true).setBackground('#E3F2FD');
  
  // Sample priorities
  const priorities = [
    ['Q4 Campaign Launch', 0.40],
    ['Brand Awareness Push', 0.30],
    ['Product Feature Release', 0.20],
    ['Holiday Season Prep', 0.10]
  ];
  
  priorities.forEach((p, i) => {
    const row = 7 + i;
    setCell(sheet, `E${row}`, p[0]);
    setCell(sheet, `F${row}`, p[1], {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    // Add checkboxes for workstream columns only
    for (let col = 7; col <= 10; col++) {
      sheet.getRange(row, col).insertCheckboxes();
    }
  });
  
  // Empty rows for more priorities
  for (let row = 11; row <= 21; row++) {
    setCell(sheet, `E${row}`, '');
    setCell(sheet, `F${row}`, '', {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    // Add checkboxes only for workstream columns
    for (let col = 7; col <= 10; col++) {
      sheet.getRange(row, col).insertCheckboxes();
    }
  }
  
  // Add TOTAL row for Strategic Priorities
  setCell(sheet, 'E22', 'TOTAL', {
    fontWeight: true, background: '#E0E0E0'
  });
  setCell(sheet, 'F22', '=SUM(F7:F21)', {
    format: '0%', fontWeight: true, background: '#E0E0E0'
  });
  // Conditional formatting for total
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F$22>1')
    .setFontColor('#FF0000')
    .setRanges([sheet.getRange('F22')])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  sheet.getRange(6, 5, 17, 6).setBorder(true, true, true, true, true, true);
}

// ==================== GENERATE TEAM CAPACITY FORMULA ====================
function generateTeamCapacityFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    return CONFIG.DEFAULT_CAPACITY; // Default value if no teams
  }
  
  // Build formula to sum all teams' net capacity (D7)
  const teamRefs = teams.map(team => `'${team} Team'!D7`);
  return `=SUM(${teamRefs.join(',')})`;
}

// ==================== UPDATE TOTAL CAPACITY ====================
function updateTotalCapacity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  if (!allocSheet) return;
  
  const formula = generateTeamCapacityFormula();
  allocSheet.getRange('B6').setFormula(formula);
}

// ==================== WORKSTREAM TAB ====================
function setupWorkstreamTab(sheet, workstreamName) {
  sheet.clear();
  
  // Ensure 6 columns
  adjustColumns(sheet, 6);
  
  // Set widths
  const widths = [400, 120, 100, 80, 120, 120];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Header with month reference
  setCell(sheet, 'A1:F1', 
    `=CONCATENATE("${workstreamName.toUpperCase()} WORKSTREAM - ",Allocation!C3," ",Allocation!E3)`, {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_GREEN, fontColor: '#FFFFFF'
  });
  
  // Budget Summary
  setCell(sheet, 'A2', 'Total Points Allocated:');
  setCell(sheet, 'B2', `=IFERROR(VLOOKUP("${workstreamName}",Allocation!A:C,3,FALSE),0)`, {
    fontSize: 14, fontWeight: true, background: CONFIG.COLORS.LIGHT_GREEN, format: '0'
  });
  
  setCell(sheet, 'A3', 'Points Spent on Assets:');
  setCell(sheet, 'B3', '=SUMIF(D46:D95,">0")', {
    fontSize: 14, fontWeight: true, background: CONFIG.COLORS.LIGHT_ORANGE, format: '0'
  });
  
  setCell(sheet, 'C2', 'Remaining:');
  setCell(sheet, 'D2', '=B2-B3', {
    fontSize: 14, fontWeight: true, background: '#E1F5FE', format: '0'
  });
  
  // Priorities Table
  sheet.getRange('A5:D5').setValues([['Priority Name', 'Source', 'Allocation %', 'Points']])
    .setFontWeight(true).setBackground('#E3F2FD');
  
  // Workstream Priorities Section
  setCell(sheet, 'A6:D6', '--- Workstream Team Priorities (Direct %) ---', {
    merge: true, fontStyle: true, background: '#FFF9C4'
  });
  
  // Workstream priority rows
  for (let i = 0; i < 10; i++) {
    const row = 7 + i;
    setCell(sheet, `A${row}`, '', { background: CONFIG.COLORS.LIGHT_YELLOW });
    setCell(sheet, `B${row}`, 'Workstream');
    setCell(sheet, `C${row}`, '', { format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW });
    setCell(sheet, `D${row}`, `=IF(C${row}="","",ROUND(C${row}*$B$2,0))`, {
      format: '0', background: CONFIG.COLORS.LIGHT_GREEN
    });
  }
  
  // Remaining capacity
  setCell(sheet, 'A18', 'Remaining for PMM:', { fontWeight: true });
  setCell(sheet, 'B18', '=MAX(0,100%-SUMIF(C7:C16,">0"))', {
    format: '0%', fontWeight: true, background: '#E3F2FD'
  });
  setCell(sheet, 'C18', '←');
  setCell(sheet, 'D18', '=ROUND(B18*B2,0)', {
    format: '0', fontWeight: true, background: '#E3F2FD'
  });
  
  // PMM Priorities Section
  setCell(sheet, 'A20:D20', '--- PMM Strategic Priorities (Auto-scaled) ---', {
    merge: true, fontStyle: true, background: CONFIG.COLORS.PMM_BLUE
  });
  
  // Setup PMM formulas
  setupPMMFormulas(sheet, workstreamName);
  
  // Summary
  setCell(sheet, 'A37', 'WORKSTREAM %:', { fontWeight: true });
  setCell(sheet, 'B37', '=SUMIF(C7:C16,">0")', { format: '0%', fontWeight: true });
  setCell(sheet, 'C37', 'PMM %:', { fontWeight: true });
  setCell(sheet, 'D37', '=SUMIF(C21:C35,">0")', { format: '0%', fontWeight: true });
  setCell(sheet, 'A38', 'TOTAL POINTS:', { fontWeight: true });
  setCell(sheet, 'B38', '=SUM(D7:D16,D21:D35)', { format: '0', fontWeight: true });
  
  // Borders
  sheet.getRange(5, 1, 12, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(20, 1, 16, 4).setBorder(true, true, true, true, true, true);
  
  // Asset Planning Section
  setupAssetSection(sheet, workstreamName);
}

// ==================== PMM FORMULAS ====================
function setupPMMFormulas(sheet, workstreamName) {
  const checkboxCol = findCheckboxColumn(workstreamName);
  if (checkboxCol < 0) return;
  
  const colLetter = columnToLetter(checkboxCol);
  
  for (let i = 0; i < 15; i++) {
    const row = 21 + i;
    const allocRow = 7 + i;
    
    setCell(sheet, `A${row}`, 
      `=IF(Allocation!${colLetter}${allocRow}=TRUE,Allocation!E${allocRow},"")`, 
      { background: CONFIG.COLORS.PMM_BLUE });
    
    setCell(sheet, `B${row}`, 
      `=IF(A${row}<>"","PMM","")`, 
      { background: CONFIG.COLORS.PMM_BLUE });
    
    const percentFormula = `=IF(A${row}="","",` +
      `IF(SUMPRODUCT(Allocation!${colLetter}7:${colLetter}21*Allocation!F7:F21)=0,0,` +
      `(Allocation!F${allocRow}*Allocation!${colLetter}${allocRow})/` +
      `SUMPRODUCT(Allocation!${colLetter}7:${colLetter}21*Allocation!F7:F21)*B18))`;
    
    setCell(sheet, `C${row}`, percentFormula, 
      { format: '0%', background: CONFIG.COLORS.PMM_BLUE });
    
    setCell(sheet, `D${row}`, 
      `=IF(C${row}="","",ROUND(C${row}*$B$2,0))`, 
      { format: '0', background: CONFIG.COLORS.LIGHT_GREEN });
  }
}

// ==================== TEAM TAB ====================
function setupTeamTab(sheet, teamName) {
  sheet.clear();
  
  // Set columns
  const widths = [120, 400, 100, 80, 120, 120];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Header with month reference
  setCell(sheet, 'A1:F1', 
    `=CONCATENATE("${teamName.toUpperCase()} TEAM - ",Allocation!C3," ",Allocation!E3," Manifest")`, {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_PURPLE, fontColor: '#FFFFFF'
  });
  
  // Capacity Section
  setCell(sheet, 'A3:F3', 'TEAM CAPACITY & PLANNING', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  // Capacity inputs
  const capacityData = [
    ['A4', 'Team Members:', 'B4', 5],
    ['C4', 'Working Days/Month:', 'D4', 20],
    ['A5', 'Total Days Off (All Members):', 'B5', 0]
  ];
  
  capacityData.forEach(([labelCell, label, valueCell, value]) => {
    setCell(sheet, labelCell, label);
    setCell(sheet, valueCell, value, {
      background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
    });
  });
  
  // Buffer percentage
  setCell(sheet, 'A6', 'Buffer % (sick/team projects):');
  setCell(sheet, 'B6', CONFIG.DEFAULT_BUFFER_PERCENT, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0%'
  });
  
  // Creative Planning with next month reference
  setCell(sheet, 'A7', `=CONCATENATE("Creative Planning Days (for ",IF(Allocation!C3="December","January",INDEX({"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"},MATCH(Allocation!C3,{"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November"},0)+1)),"):")`);
  setCell(sheet, 'B7', 0, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
  });
  
  setCell(sheet, 'C5', 'Gross Capacity:');
  setCell(sheet, 'D5', '=(B4*D4)-B5', {
    fontWeight: true, background: '#E8F5E9', format: '0'
  });
  
  setCell(sheet, 'C6', 'Buffer Points:');
  setCell(sheet, 'D6', '=ROUND(D5*B6,0)', {
    fontWeight: true, background: '#FFE0B2', format: '0'
  });
  
  setCell(sheet, 'C7', 'Net Capacity:');
  setCell(sheet, 'D7', '=D5-D6-B7', {
    fontWeight: true, background: CONFIG.COLORS.LIGHT_GREEN, format: '0'
  });
  
  // Assignment Summary
  setCell(sheet, 'A9:F9', 'ASSIGNMENT SUMMARY', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  setCell(sheet, 'A10', 'Workstream Assigned:');
  setCell(sheet, 'B10', '=SUMIF(F14:F200,"Workstream",D14:D200)+SUMIF(F14:F200,"PMM",D14:D200)', {
    fontWeight: true, fontSize: 14, background: CONFIG.COLORS.LIGHT_ORANGE, format: '0'
  });
  
  setCell(sheet, 'C10', 'Team Initiated:');
  setCell(sheet, 'D10', '=SUMIF(F14:F200,"Team",D14:D200)', {
    fontWeight: true, fontSize: 14, background: '#E1F5FE', format: '0'
  });
  
  setCell(sheet, 'E10', 'Buffer/Team Projects:');
  setCell(sheet, 'F10', '=D6', {
    fontWeight: true, fontSize: 14, background: '#FFE0B2', format: '0'
  });
  
  setCell(sheet, 'A11', `=CONCATENATE("Creative Planning (",IF(Allocation!C3="December","Jan",TEXT(DATE(2000,MATCH(Allocation!C3,{"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"},0)+1,1),"mmm")),"):")`);
  setCell(sheet, 'B11', '=B7', {
    fontWeight: true, fontSize: 14, background: '#FFECB3', format: '0'
  });
  
  setCell(sheet, 'C11', 'Total Allocated:');
  setCell(sheet, 'D11', '=B10+D10+F10+B11', {
    fontWeight: true, fontSize: 14, background: '#FFD54F', format: '0'
  });
  
  setCell(sheet, 'E11', 'Utilization:');
  setCell(sheet, 'F11', '=IF(D5=0,"",D11/D5)', {
    fontWeight: true, fontSize: 14, format: '0%'
  });
  
  setCell(sheet, 'A12:F12', 
    '=IF(D11>D5,"⚠️ OVER by "&(D11-D5)&" pts",IF(D11=D5,"✅ FULL","✅ "&(D5-D11)&" pts available"))', {
    merge: true, fontWeight: true, fontSize: 12
  });
  
  // Conditional formatting for over capacity
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D$11>$D$5')
    .setFontColor('#FF0000')
    .setRanges([sheet.getRange('D11')])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Table headers
  sheet.getRange('A13:F13').setValues([['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source']])
    .setFontWeight(true).setBackground('#E1BEE7');
  
  setCell(sheet, 'A15', 'Click "Refresh Team Assignments" to load workstream assignments...', {
    fontStyle: true, fontColor: '#666666'
  });
  
  // Team-initiated section
  setCell(sheet, 'A61:F61', '--- TEAM-INITIATED WORK ---', {
    merge: true, fontWeight: true, fontStyle: true, background: '#E1F5FE'
  });
  
  // Team rows
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CONFIG.TSHIRT_SIZES), true)
    .setAllowInvalid(false)
    .build();
  
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const currentDay = today.getDate();
  
  for (let i = 0; i < 30; i++) {
    const row = 62 + i;
    
    setCell(sheet, `A${row}`, teamName, { background: CONFIG.COLORS.GRAY });
    setCell(sheet, `B${row}`, '', { background: CONFIG.COLORS.LIGHT_YELLOW });
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
    const defaultDate = new Date(currentYear, currentMonth, currentDay);
    setCell(sheet, `E${row}`, defaultDate, { background: '#FFF9C4' });
    sheet.getRange(row, 5).setNumberFormat('yyyy-MM-dd');
    setCell(sheet, `F${row}`, 'Team', { background: '#E1F5FE' });
  }
  
  // Borders
  sheet.getRange(3, 1, 5, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(9, 1, 4, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(61, 1, 31, 6).setBorder(true, true, true, true, true, true);
}

// ==================== TEAM ASSIGNMENTS ====================
function refreshTeamAssignments(sortBy = 'workstream') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No teams found. Please add teams first.');
    return;
  }
  
  // Initialize team assignments
  const teamAssignments = {};
  teams.forEach(team => {
    teamAssignments[team] = [];
    const teamSheet = ss.getSheetByName(team + ' Team');
    if (teamSheet) {
      // Clear workstream assignments area only (rows 14-60)
      teamSheet.getRange(14, 1, 47, 6).clear();
    }
  });
  
  // Add Creative Planning as a regular item for each team (but NOT buffer)
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (teamSheet) {
      const creativePlanningDays = teamSheet.getRange('B7').getValue();
      if (creativePlanningDays > 0) {
        const allocSheet = ss.getSheetByName('Allocation');
        const monthName = allocSheet.getRange('C3').getValue();
        const year = allocSheet.getRange('E3').getValue();
        const monthIndex = CONFIG.MONTHS.indexOf(monthName);
        
        let nextMonthIndex = (monthIndex + 1) % 12;
        let nextYear = year;
        if (nextMonthIndex === 0) {
          nextYear = year + 1;
        }
        const planningDate = new Date(nextYear, nextMonthIndex, 1);
        const nextMonthName = CONFIG.MONTHS[nextMonthIndex];
        
        teamAssignments[teamName].push({
          origin: teamName,
          description: `Creative Planning & Ideation (for ${nextMonthName})`,
          size: '-',
          points: creativePlanningDays,
          goLiveDate: Utilities.formatDate(planningDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          source: 'Team'
        });
      }
    }
  });
  
  // Collect from workstreams (with improved source tracking)
  const workstreams = getWorkstreamNames();
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    const data = wsSheet.getRange(46, 1, 50, 6).getValues();
    data.forEach((row, i) => {
      const [description, goLiveDate, tShirtSize, points, origin, teamAssignment] = row;
      
      if (description && teamAssignment && teams.includes(teamAssignment)) {
        let formattedDate = goLiveDate;
        if (goLiveDate instanceof Date) {
          formattedDate = Utilities.formatDate(goLiveDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        // Determine if this is PMM or Workstream origin
        const sourceType = (origin === 'PMM') ? 'PMM' : 'Workstream';
        
        teamAssignments[teamAssignment].push({
          origin: wsName,
          description,
          size: tShirtSize,
          points,
          goLiveDate: formattedDate,
          source: sourceType  // Now properly tracks PMM vs Workstream
        });
      }
    });
  });
  
  // Collect team-initiated work
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const teamData = teamSheet.getRange(62, 1, 30, 6).getValues();
    teamData.forEach(row => {
      const [origin, description, tShirtSize, points, goLiveDate, source] = row;
      
      if (description && points > 0) {
        let formattedDate = goLiveDate;
        if (goLiveDate instanceof Date) {
          formattedDate = Utilities.formatDate(goLiveDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        teamAssignments[teamName].push({
          origin: teamName,
          description,
          size: tShirtSize,
          points,
          goLiveDate: formattedDate,
          source: 'Team'
        });
      }
    });
  });
  
  // Write to team sheets
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const assignments = teamAssignments[teamName];
    if (assignments.length === 0) {
      setCell(teamSheet, 'A14', 'No assignments', {
        fontStyle: true, fontColor: '#666666'
      });
      return;
    }
    
    // Sort and write assignments
    if (sortBy === 'date') {
      assignments.sort((a, b) => {
        const dateA = a.goLiveDate ? String(a.goLiveDate) : '';
        const dateB = b.goLiveDate ? String(b.goLiveDate) : '';
        
        if (dateA && dateB && dateA !== '-' && dateB !== '-') {
          const dateCompare = dateA.localeCompare(dateB);
          if (dateCompare !== 0) return dateCompare;
        } else if (dateA && dateA !== '-') {
          return -1;
        } else if (dateB && dateB !== '-') {
          return 1;
        }
        return a.origin.localeCompare(b.origin);
      });
      
      let currentRow = 14;
      assignments.forEach(a => {
        if (currentRow >= 61) return;
        
        teamSheet.getRange(currentRow, 1, 1, 6).setValues([[
          a.origin,
          a.description,
          a.size,
          a.points,
          a.goLiveDate || '',
          a.source
        ]]);
        
        if (a.goLiveDate && a.goLiveDate !== '-') {
          teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
        }
        teamSheet.getRange(currentRow, 4).setNumberFormat('0');
        
        // Color code based on source
        if (a.source === 'PMM') {
          teamSheet.getRange(currentRow, 6).setBackground(CONFIG.COLORS.PMM_BLUE);
        } else if (a.source === 'Workstream') {
          teamSheet.getRange(currentRow, 6).setBackground(CONFIG.COLORS.LIGHT_YELLOW);
        } else if (a.source === 'Team') {
          teamSheet.getRange(currentRow, 6).setBackground('#E1F5FE');
        }
        
        currentRow++;
      });
      
      if (currentRow > 14) {
        teamSheet.getRange(14, 1, currentRow - 14, 6)
          .setBorder(true, true, true, true, true, false);
      }
    } else {
      // Default: Group by workstream
      const grouped = {};
      assignments.forEach(a => {
        if (!grouped[a.origin]) grouped[a.origin] = [];
        grouped[a.origin].push(a);
      });
      
      let currentRow = 14;
      
      // Sort keys to put team's own work first
      const sortedKeys = Object.keys(grouped).sort((a, b) => {
        if (a === teamName) return -1;
        if (b === teamName) return 1;
        return a.localeCompare(b);
      });
      
      sortedKeys.forEach(wsName => {
        if (currentRow >= 61) return;
        
        // Add section header only for external workstreams
        if (wsName !== teamName) {
          setCell(teamSheet, `A${currentRow}:F${currentRow}`, `--- ${wsName} ---`, {
            merge: true, fontWeight: true, background: '#F5F5F5', fontStyle: true
          });
          currentRow++;
        }
        
        // Write assignments
        grouped[wsName].forEach(a => {
          if (currentRow >= 61) return;
          
          teamSheet.getRange(currentRow, 1, 1, 6).setValues([[
            a.origin,
            a.description,
            a.size,
            a.points,
            a.goLiveDate || '',
            a.source
          ]]);
          
          if (a.goLiveDate && a.goLiveDate !== '-') {
            teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
          }
          teamSheet.getRange(currentRow, 4).setNumberFormat('0');
          
          // Color code based on source
          if (a.source === 'PMM') {
            teamSheet.getRange(currentRow, 6).setBackground(CONFIG.COLORS.PMM_BLUE);
          } else if (a.source === 'Workstream') {
            teamSheet.getRange(currentRow, 6).setBackground(CONFIG.COLORS.LIGHT_YELLOW);
          } else if (a.source === 'Team') {
            teamSheet.getRange(currentRow, 6).setBackground('#E1F5FE');
          }
          
          currentRow++;
        });
      });
      
      if (currentRow > 14) {
        teamSheet.getRange(14, 1, currentRow - 14, 6)
          .setBorder(true, true, true, true, true, false);
      }
    }
  });
  
  SpreadsheetApp.getUi().alert('Team assignments refreshed successfully!');
}

// ==================== ASSET PLANNING SECTION ====================
function setupAssetSection(sheet, workstreamName) {
  // Header
  setCell(sheet, 'A41:F41', 'ASSET PLANNING', {
    merge: true, fontSize: 14, fontWeight: true,
    background: CONFIG.COLORS.HEADER_ORANGE, fontColor: '#FFFFFF'
  });
  
  // Budget info
  setCell(sheet, 'A42', 'Budget:');
  setCell(sheet, 'B42', '=B2', {
    fontWeight: true, format: '0', background: CONFIG.COLORS.LIGHT_GREEN
  });
  setCell(sheet, 'C42', 'Spent:');
  setCell(sheet, 'D42', '=SUMIF(D46:D95,">0")', {
    fontWeight: true, format: '0', background: CONFIG.COLORS.LIGHT_ORANGE
  });
  setCell(sheet, 'E42:F42', 
    '=IF(B42-D42<0,"⚠️ OVER by "&ABS(B42-D42),"💰 "&(B42-D42)&" remaining")', {
    merge: true, fontWeight: true
  });
  
  // T-shirt legend
  setCell(sheet, 'A43:F43', 
    'T-Shirt Sizes: XS=1 point | S=3 points | M=5 points | L=13 points | XL=21 points', {
    merge: true, fontSize: 10, fontStyle: true, background: CONFIG.COLORS.LIGHT_YELLOW
  });
  
  // Table headers
  const headers = ['Asset Description', 'Go Live Date', 'T-Shirt Size', 'Cost', 'Origin', 'Team'];
  headers.forEach((header, i) => {
    setCell(sheet, `${String.fromCharCode(65 + i)}45`, header, {
      fontWeight: true, background: CONFIG.COLORS.LIGHT_ORANGE
    });
  });
  
  // Asset rows
  const teamNames = getTeamNames();
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CONFIG.TSHIRT_SIZES), true)
    .setAllowInvalid(false)
    .build();
  
  // Origin validation for Workstream or PMM
  const originValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Workstream', 'PMM'], true)
    .setAllowInvalid(false)
    .build();
  
  const teamValidation = teamNames.length > 0 ? 
    SpreadsheetApp.newDataValidation()
      .requireValueInList(teamNames, true)
      .setAllowInvalid(false)
      .build() : null;
  
  // Get current date for defaults
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const currentDay = today.getDate();
  
  // Create all rows first
  for (let i = 0; i < 50; i++) {
    const row = 46 + i;
    
    setCell(sheet, `A${row}`, '', { background: '#FFFFFF' });
    
    // Set default date to current date
    const defaultDate = new Date(currentYear, currentMonth, currentDay);
    setCell(sheet, `B${row}`, defaultDate, { background: '#FFF9C4' });
    sheet.getRange(row, 2).setNumberFormat('yyyy-MM-dd');
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    // T-shirt size formula
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
    // Origin dropdown - default to Workstream with yellow background
    sheet.getRange(row, 5).setDataValidation(originValidation);
    setCell(sheet, `E${row}`, 'Workstream', {
      background: CONFIG.COLORS.LIGHT_YELLOW
    });
    
    if (teamValidation) {
      sheet.getRange(row, 6).setDataValidation(teamValidation)
        .setBackground(CONFIG.COLORS.LIGHT_PURPLE);
    } else {
      setCell(sheet, `F${row}`, '', { background: CONFIG.COLORS.LIGHT_PURPLE });
    }
  }
  
  // Add conditional formatting for Origin column (E46:E95) after all rows are created
  const originRange = sheet.getRange('E46:E95');
  
  // Create rules for coloring based on value
  const wsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Workstream')
    .setBackground(CONFIG.COLORS.LIGHT_YELLOW)
    .setRanges([originRange])
    .build();
  
  const pmmRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('PMM')
    .setBackground(CONFIG.COLORS.PMM_BLUE)
    .setRanges([originRange])
    .build();
  
  // Apply the conditional formatting rules
  const existingRules = sheet.getConditionalFormatRules();
  existingRules.push(wsRule, pmmRule);
  sheet.setConditionalFormatRules(existingRules);
  
  sheet.getRange(45, 1, 51, 6).setBorder(true, true, true, true, true, true);
}

// ==================== WORKSTREAM MANAGEMENT ====================
function addWorkstream() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Add New Workstream', 
    'Enter the name for the new workstream:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const name = response.getResponseText().trim();
  if (!name) {
    ui.alert('Error', 'Workstream name cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) {
    ui.alert('Error', `"${name}" already exists.`, ui.ButtonSet.OK);
    return;
  }
  
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Find TOTAL row
  let totalRow = 10;
  while (allocSheet.getRange(totalRow, 1).getValue() !== 'TOTAL' && totalRow < 20) {
    totalRow++;
  }
  
  // Insert in allocation table
  allocSheet.insertRowBefore(totalRow);
  setCell(allocSheet, `A${totalRow}`, name, {
    background: '#FFFFFF'  // Ensure white background for workstream name
  });
  setCell(allocSheet, `B${totalRow}`, 0, {
    format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
  });
  setCell(allocSheet, `C${totalRow}`, `=ROUND($B$6*B${totalRow},0)`, {
    format: '0', background: '#F5F5F5'
  });
  
  // Apply border to the new row
  allocSheet.getRange(totalRow, 1, 1, 3).setBorder(true, true, true, true, true, false);
  
  // Update TOTAL formulas and ensure TOTAL row is properly formatted
  const newTotalRow = totalRow + 1;
  setCell(allocSheet, `A${newTotalRow}`, 'TOTAL', {
    fontWeight: true, background: '#E0E0E0'
  });
  setCell(allocSheet, `B${newTotalRow}`, `=SUM(B10:B${totalRow})`, {
    fontWeight: true, background: '#E0E0E0', format: '0%'
  });
  setCell(allocSheet, `C${newTotalRow}`, `=SUM(C10:C${totalRow})`, {
    fontWeight: true, background: '#E0E0E0', format: '0'
  });
  
  // Re-apply border to include the new row
  allocSheet.getRange(9, 1, newTotalRow - 8, 3).setBorder(true, true, true, true, true, true);
  
  // Find next available column for checkbox header
  let nextCol = 7;
  while (allocSheet.getRange(6, nextCol).getValue() && nextCol < 20) {
    nextCol++;
  }
  
  // Add header for new workstream's checkbox column
  setCell(allocSheet, `${columnToLetter(nextCol)}6`, name, {
    fontWeight: true, background: '#E3F2FD'
  });
  allocSheet.setColumnWidth(nextCol, 80);
  
  // Add checkboxes for all priority rows
  for (let row = 7; row <= 21; row++) {
    allocSheet.getRange(row, nextCol).insertCheckboxes();
  }
  
  // Create workstream sheet
  const wsSheet = ss.insertSheet(name);
  setupWorkstreamTab(wsSheet, name);
  
  // Update team dropdowns
  updateTeamDropdowns();
  
  ui.alert('Success', `"${name}" added. Allocate points in the Allocation tab.`, ui.ButtonSet.OK);
}

function removeWorkstream() {
  const ui = SpreadsheetApp.getUi();
  const workstreams = getWorkstreamNames();
  
  if (workstreams.length === 0) {
    ui.alert('Error', 'No workstreams to remove.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt('Remove Workstream',
    'Enter the name to remove:\n\n' + workstreams.join(', '), ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const name = response.getResponseText().trim();
  if (!workstreams.includes(name)) {
    ui.alert('Error', `"${name}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Remove from allocation table
  let row = 10;
  while (allocSheet.getRange(row, 1).getValue() !== 'TOTAL') {
    if (allocSheet.getRange(row, 1).getValue() === name) {
      allocSheet.deleteRow(row);
      break;
    }
    row++;
  }
  
  // Update TOTAL formulas with formatting
  setCell(allocSheet, `B${row}`, `=SUM(B10:B${row-1})`, {
    fontWeight: true, background: '#E0E0E0', format: '0%'
  });
  setCell(allocSheet, `C${row}`, `=SUM(C10:C${row-1})`, {
    fontWeight: true, background: '#E0E0E0', format: '0'
  });
  
  // Remove column
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(6, col).getValue() === name) {
      allocSheet.deleteColumn(col);
      break;
    }
  }
  
  // Delete sheet
  const wsSheet = ss.getSheetByName(name);
  if (wsSheet) ss.deleteSheet(wsSheet);
  
  ui.alert('Success', `"${name}" removed.`, ui.ButtonSet.OK);
}

// ==================== TEAM MANAGEMENT ====================
function addTeam() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Add New Team', 
    'Enter the team name:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const name = response.getResponseText().trim();
  if (!name) {
    ui.alert('Error', 'Team name cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name + ' Team';
  
  if (ss.getSheetByName(sheetName)) {
    ui.alert('Error', `Team "${name}" already exists.`, ui.ButtonSet.OK);
    return;
  }
  
  const teamSheet = ss.insertSheet(sheetName);
  setupTeamTab(teamSheet, name);
  updateTeamDropdowns();
  updateTotalCapacity(); // Update the formula when adding a team
  
  ui.alert('Success', `Team "${name}" added.`, ui.ButtonSet.OK);
}

function removeTeam() {
  const ui = SpreadsheetApp.getUi();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    ui.alert('Error', 'No teams to remove.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt('Remove Team',
    'Enter the team name:\n\n' + teams.join(', '), ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const name = response.getResponseText().trim();
  if (!teams.includes(name)) {
    ui.alert('Error', `Team "${name}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamSheet = ss.getSheetByName(name + ' Team');
  if (teamSheet) ss.deleteSheet(teamSheet);
  
  updateTeamDropdowns();
  updateTotalCapacity(); // Update the formula when removing a team
  ui.alert('Success', `Team "${name}" removed.`, ui.ButtonSet.OK);
}

// Sort manifest by date
function sortManifestByDate() {
  refreshTeamAssignments('date');
}

// Sort manifest by workstream
function sortManifestByWorkstream() {
  refreshTeamAssignments('workstream');
}

// ==================== HELPER FUNCTIONS ====================
function setCell(sheet, range, value, options = {}) {
  const cell = sheet.getRange(range);
  
  if (options.merge) cell.merge();
  
  if (value !== undefined) {
    if (typeof value === 'string' && value.startsWith('=')) {
      cell.setFormula(value);
    } else {
      cell.setValue(value);
    }
  }
  
  if (options.fontSize) cell.setFontSize(options.fontSize);
  if (options.fontWeight) cell.setFontWeight('bold');
  if (options.fontStyle) cell.setFontStyle('italic');
  if (options.fontColor) cell.setFontColor(options.fontColor);
  if (options.background) cell.setBackground(options.background);
  if (options.format) cell.setNumberFormat(options.format);
  if (options.border) cell.setBorder(true, true, true, true, false, false);
  
  return cell;
}

function adjustColumns(sheet, targetCols) {
  const maxCols = sheet.getMaxColumns();
  if (maxCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, maxCols - targetCols);
  } else if (maxCols < targetCols) {
    sheet.insertColumnsAfter(maxCols, targetCols - maxCols);
  }
}

function findCheckboxColumn(workstreamName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(6, col).getValue() === workstreamName) {
      return col;
    }
  }
  return -1;
}

function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

function getTeamNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name.endsWith(' Team'))
    .map(name => name.replace(' Team', ''));
}

function getWorkstreamNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  const workstreams = [];
  
  let row = 10;
  let wsName = allocSheet.getRange(row, 1).getValue();
  while (wsName && wsName !== 'TOTAL' && row < 20) {
    workstreams.push(wsName);
    row++;
    wsName = allocSheet.getRange(row, 1).getValue();
  }
  
  return workstreams;
}

function updateTeamDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  const workstreams = getWorkstreamNames();
  
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    for (let row = 46; row <= 95; row++) {
      if (teams.length === 0) {
        wsSheet.getRange(row, 6).clearDataValidations();
      } else {
        const validation = SpreadsheetApp.newDataValidation()
          .requireValueInList(teams, true)
          .setAllowInvalid(false)
          .build();
        wsSheet.getRange(row, 6).setDataValidation(validation);
      }
    }
  });
}

// ==================== ABOUT FUNCTION ====================
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const message = `Marketing Team Points System

Version: ${VERSION}
Release Date: ${RELEASE_DATE}

© 2024 Marketing Team`;
  
  ui.alert('About Points System', message, ui.ButtonSet.OK);
}

// ==================== MENU SETUP ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Points System')
    .addItem('🚀 Initial Setup', 'setupPointsSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('Workstreams')
      .addItem('➕ Add Workstream', 'addWorkstream')
      .addItem('➖ Remove Workstream', 'removeWorkstream'))
    .addSubMenu(ui.createMenu('Teams')
      .addItem('➕ Add Team', 'addTeam')
      .addItem('➖ Remove Team', 'removeTeam')
      .addSeparator()
      .addItem('🔄 Refresh Team Assignments', 'refreshTeamAssignments')
      .addItem('📅 Sort Manifest by Date', 'sortManifestByDate')
      .addItem('📊 Sort Manifest by Workstream', 'sortManifestByWorkstream'))
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}