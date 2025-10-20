/**
 * Marketing Team Points System - Version 13.3
 * 
 * FIXES IN v13.3:
 * - Fixed creative planning double-counting issue
 * - Creative planning now only reduces capacity, not added as work
 * - Corrected Total Allocated formula to exclude creative planning
 * - Fixed manifest generation to not include creative planning as line item
 */

const VERSION = "13.3";
const RELEASE_DATE = "2024-12-21";

const CONFIG = {
  DEFAULT_WORKSTREAMS: ['SoMe', 'PUA', 'ASO', 'Portal'],
  DEFAULT_ALLOCATIONS: [0.50, 0.20, 0.05, 0.25],
  DEFAULT_CAPACITY: 100,
  DEFAULT_BUFFER_PERCENT: 0.10,
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
  
  // Update team dropdowns and capacity
  updateTeamDropdowns();
  updateTotalCapacity();
  
  ss.setActiveSheet(allocationSheet);
  
  SpreadsheetApp.getUi().alert(
    'Points System Setup Complete! 🎉',
    `System v${VERSION} ready with fixed creative planning calculations.`,
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
  setCell(sheet, 'A3', 'Planning Month:', { fontWeight: true });
  
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
  
  // Auto-sum from all teams' net capacity
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

// ==================== WORKSTREAM TAB ====================
function setupWorkstreamTab(sheet, workstreamName) {
  sheet.clear();
  
  // Set column count and widths
  adjustColumns(sheet, 8);
  const widths = [400, 120, 100, 80, 120, 120, 200, 150];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Header with month reference
  setCell(sheet, 'A1:H1', 
    `=CONCATENATE("${workstreamName.toUpperCase()} WORKSTREAM - ",Allocation!C3," ",Allocation!E3)`, {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_GREEN, fontColor: '#FFFFFF'
  });
  
  // Budget Summary (Left side)
  setCell(sheet, 'A2', 'Total Points Allocated:');
  setCell(sheet, 'B2', `=IFERROR(VLOOKUP("${workstreamName}",Allocation!A:C,3,FALSE),0)`, {
    fontSize: 14, fontWeight: true, background: CONFIG.COLORS.LIGHT_GREEN, format: '0'
  });
  
  setCell(sheet, 'A3', 'Points Spent on Assets:');
  setCell(sheet, 'B3', '=SUMIF(D46:D95,">0",D46:D95)', {
    fontSize: 14, fontWeight: true, background: CONFIG.COLORS.LIGHT_ORANGE, format: '0'
  });
  
  setCell(sheet, 'C2', 'Remaining:');
  setCell(sheet, 'D2', '=B2-B3', {
    fontSize: 14, fontWeight: true, background: '#E1F5FE', format: '0'
  });
  
  // Workstream Owner Info (Right side)
  setCell(sheet, 'F2', 'Workstream Owner:');
  setCell(sheet, 'G2', '', { background: CONFIG.COLORS.LIGHT_YELLOW });
  
  setCell(sheet, 'F3', 'Email:');
  setCell(sheet, 'G3', '', { background: '#FFF9C4' });
  
  setCell(sheet, 'F4', 'Jira Username:');
  setCell(sheet, 'G4', '', { background: '#E1F5FE' });
  
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
  sheet.getRange(2, 6, 3, 2).setBorder(true, true, true, true, true, true);
  
  // Asset Planning Section
  setupAssetSection(sheet, workstreamName);
}

// ==================== TEAM TAB WITH FIXED FORMULAS ====================
function setupTeamTab(sheet, teamName) {
  sheet.clear();
  
  // Set column widths for compact layout
  const widths = [120, 300, 80, 80, 120, 120, 150, 200, 100, 80];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Header with month reference
  setCell(sheet, 'A1:J1', 
    `=CONCATENATE("${teamName.toUpperCase()} TEAM - ",Allocation!C3," ",Allocation!E3," Manifest")`, {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_PURPLE, fontColor: '#FFFFFF'
  });
  
  // Team Capacity & Planning Section (Compact - Left side)
  setCell(sheet, 'A3:F3', 'TEAM CAPACITY & PLANNING', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  // Row 4 - Team Members and Working Days
  setCell(sheet, 'A4', 'Team Members:');
  setCell(sheet, 'B4', 5, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
  });
  setCell(sheet, 'C4', 'Working Days/Month:');
  setCell(sheet, 'D4', 20, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
  });
  
  // Row 5 - Days Off and Gross Capacity
  setCell(sheet, 'A5', 'Total Days Off (All Members):');
  setCell(sheet, 'B5', '=SUM(J5:J14)', {
    background: '#E8F5E9', format: '0', fontWeight: true
  });
  setCell(sheet, 'C5', 'Gross Capacity:');
  setCell(sheet, 'D5', '=(B4*D4)-B5', {
    fontWeight: true, background: '#E8F5E9', format: '0'
  });
  
  // Row 6 - Buffer and Buffer Points
  setCell(sheet, 'A6', 'Buffer % (sick/team projects):');
  setCell(sheet, 'B6', CONFIG.DEFAULT_BUFFER_PERCENT, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0%'
  });
  setCell(sheet, 'C6', 'Buffer Points:');
  setCell(sheet, 'D6', '=ROUND(D5*B6,0)', {
    fontWeight: true, background: '#FFE0B2', format: '0'
  });
  
  // Row 7 - Creative Planning and Net Capacity
  setCell(sheet, 'A7', `=CONCATENATE("Creative Planning Days (for ",IF(Allocation!C3="December","January",INDEX({"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"},MATCH(Allocation!C3,{"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November"},0)+1)),"):")`);
  setCell(sheet, 'B7', 0, {
    background: CONFIG.COLORS.LIGHT_YELLOW, format: '0'
  });
  setCell(sheet, 'C7', 'Net Capacity:');
  setCell(sheet, 'D7', '=D5-D6-B7', {
    fontWeight: true, background: CONFIG.COLORS.LIGHT_GREEN, format: '0'
  });
  
  // Team Members Section (Right side) - Limited to 10 members
  setCell(sheet, 'G3:J3', 'TEAM MEMBERS', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  sheet.getRange('G4:J4').setValues([['Name', 'Email', 'Jira Username', 'Holiday (Days)']])
    .setFontWeight(true).setBackground('#E1BEE7');
  
  // Auto-populate with Team Member 1, 2, etc. (limited to 10)
  for (let i = 0; i < 10; i++) {
    const row = 5 + i;
    setCell(sheet, `G${row}`, `Team Member ${i + 1}`, { background: CONFIG.COLORS.LIGHT_YELLOW });
    setCell(sheet, `H${row}`, '', { background: '#FFF9C4' });
    setCell(sheet, `I${row}`, '', { background: '#E1F5FE' });
    setCell(sheet, `J${row}`, 0, { background: '#FFE0B2', format: '0' });
  }
  
  // Assignment Summary (FIXED FORMULAS)
  setCell(sheet, 'A9:F9', 'ASSIGNMENT SUMMARY', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  setCell(sheet, 'A10', 'Workstream Assigned:');
  // FIXED: Use correct SUMIF syntax to sum all points from manifest where Source is "Workstream" or "PMM"
  setCell(sheet, 'B10', '=SUMIF(F16:F62,"Workstream",D16:D62)+SUMIF(F16:F62,"PMM",D16:D62)', {
    fontWeight: true, fontSize: 14, background: CONFIG.COLORS.LIGHT_ORANGE, format: '0'
  });
  
  setCell(sheet, 'C10', 'Team Initiated:');
  // FIXED: Sum points from team-initiated section AND manifest rows where Source is "Team"
  setCell(sheet, 'D10', '=SUMIF(D65:D94,">0",D65:D94)+SUMIF(F16:F62,"Team",D16:D62)', {
    fontWeight: true, fontSize: 14, background: '#E1F5FE', format: '0'
  });
  
  setCell(sheet, 'E10', 'Buffer/Team Projects:');
  setCell(sheet, 'F10', '=D6', {
    fontWeight: true, fontSize: 14, background: '#FFE0B2', format: '0'
  });
  
  // Creative planning shown but not counted in total
  setCell(sheet, 'A11', `=CONCATENATE("Creative Planning (",IF(Allocation!C3="December","Jan",TEXT(DATE(2000,MATCH(Allocation!C3,{"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"},0)+1,1),"mmm")),"):")`);
  setCell(sheet, 'B11', '=B7', {
    fontWeight: true, fontSize: 14, background: '#FFECB3', format: '0'
  });
  
  setCell(sheet, 'C11', 'Total Allocated:');
  // FIXED: Don't add creative planning (B11) to the total - it's already deducted from capacity
  setCell(sheet, 'D11', '=B10+D10+F10', {
    fontWeight: true, fontSize: 14, background: '#FFD54F', format: '0'
  });
  
  setCell(sheet, 'E11', 'Utilization:');
  setCell(sheet, 'F11', '=IF(D7=0,"",D11/D7)', {
    fontWeight: true, fontSize: 14, format: '0%'
  });
  
  // Status message comparing to NET capacity (D7)
  setCell(sheet, 'A12:F12', 
    '=IF(D11>D7,"⚠️ OVER by "&(D11-D7)&" pts",IF(D11=D7,"✅ FULL","✅ "&(D7-D11)&" pts available"))', {
    merge: true, fontWeight: true, fontSize: 12
  });
  
  // Conditional formatting for over capacity
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D$11>$D$7')
    .setFontColor('#FF0000')
    .setRanges([sheet.getRange('D11')])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Table headers at ROW 15
  sheet.getRange('A15:F15').setValues([['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source']])
    .setFontWeight(true).setBackground('#E1BEE7');
  
  // Initial placeholder text
  setCell(sheet, 'A16', 'Click "Refresh Team Assignments" to load workstream assignments...', {
    fontStyle: true, fontColor: '#666666'
  });
  
  // Team-initiated section
  setCell(sheet, 'A64:F64', '--- TEAM-INITIATED WORK ---', {
    merge: true, fontWeight: true, fontStyle: true, background: '#E1F5FE'
  });
  
  // Team rows with T-shirt validation
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CONFIG.TSHIRT_SIZES), true)
    .setAllowInvalid(false)
    .build();
  
  const today = new Date();
  
  for (let i = 0; i < 30; i++) {
    const row = 65 + i;
    
    setCell(sheet, `A${row}`, teamName, { background: CONFIG.COLORS.GRAY });
    setCell(sheet, `B${row}`, '', { background: CONFIG.COLORS.LIGHT_YELLOW });
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
    setCell(sheet, `E${row}`, today, { background: '#FFF9C4' });
    sheet.getRange(row, 5).setNumberFormat('yyyy-MM-dd');
    setCell(sheet, `F${row}`, 'Team', { background: '#E1F5FE' });
  }
  
  // Borders
  sheet.getRange(3, 1, 5, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(3, 7, 12, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(9, 1, 4, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(64, 1, 31, 6).setBorder(true, true, true, true, true, true);
}

// ==================== REFRESH TEAM ASSIGNMENTS (FIXED) ====================
function refreshTeamAssignments(sortBy = 'workstream') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No teams found. Please add teams first.');
    return;
  }
  
  // Check for over-allocations first
  const warnings = validateWorkstreamAllocations();
  if (warnings.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Warning: Team Over-allocation Detected',
      warnings.join('\n') + '\n\nContinue anyway?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }
  
  // Initialize team assignments
  const teamAssignments = {};
  teams.forEach(team => {
    teamAssignments[team] = [];
    const teamSheet = ss.getSheetByName(team + ' Team');
    if (teamSheet) {
      // Clear ONLY the data area (row 16 onwards) - keep headers intact
      teamSheet.getRange(16, 1, 47, 6).clear();
      teamSheet.getRange(16, 1, 47, 6).clearFormat();
    }
  });
  
  // REMOVED: Creative planning is NO LONGER added as a work item
  // It's already accounted for by reducing net capacity
  
  // Collect from workstreams
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
        
        const sourceType = (origin === 'PMM') ? 'PMM' : 'Workstream';
        
        teamAssignments[teamAssignment].push({
          origin: wsName,
          description,
          size: tShirtSize,
          points,
          goLiveDate: formattedDate,
          source: sourceType
        });
      }
    });
  });
  
  // Collect team-initiated work
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const teamData = teamSheet.getRange(65, 1, 30, 6).getValues();
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
  
  // Write to team sheets - START AT ROW 16
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const assignments = teamAssignments[teamName];
    if (assignments.length === 0) {
      // Write "No assignments" at row 16
      setCell(teamSheet, 'A16', 'No assignments', {
        fontStyle: true, fontColor: '#666666'
      });
      return;
    }
    
    // Prepare data arrays for batch operations
    const rowsData = [];
    const formats = [];
    const backgrounds = [];
    
    if (sortBy === 'date') {
      // Sort by date
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
      
      // Prepare data
      assignments.forEach((a, index) => {
        if (index >= 47) return; // Respect row limit
        
        rowsData.push([
          a.origin,
          a.description,
          a.size,
          a.points,
          a.goLiveDate || '',
          a.source
        ]);
        
        formats.push({
          row: 16 + index,
          source: a.source,
          hasDate: a.goLiveDate && a.goLiveDate !== '-'
        });
      });
      
    } else {
      // Default: Group by workstream
      const grouped = {};
      assignments.forEach(a => {
        if (!grouped[a.origin]) grouped[a.origin] = [];
        grouped[a.origin].push(a);
      });
      
      // Sort keys to put team's own work first
      const sortedKeys = Object.keys(grouped).sort((a, b) => {
        if (a === teamName) return -1;
        if (b === teamName) return 1;
        return a.localeCompare(b);
      });
      
      let currentIndex = 0;
      
      sortedKeys.forEach(wsName => {
        if (currentIndex >= 47) return;
        
        // Add section header only for external workstreams
        if (wsName !== teamName) {
          rowsData.push([`--- ${wsName} ---`, '', '', '', '', '']);
          formats.push({
            row: 16 + currentIndex,
            isHeader: true,
            source: null
          });
          currentIndex++;
        }
        
        // Add assignments
        grouped[wsName].forEach(a => {
          if (currentIndex >= 47) return;
          
          rowsData.push([
            a.origin,
            a.description,
            a.size,
            a.points,
            a.goLiveDate || '',
            a.source
          ]);
          
          formats.push({
            row: 16 + currentIndex,
            source: a.source,
            hasDate: a.goLiveDate && a.goLiveDate !== '-',
            isHeader: false
          });
          currentIndex++;
        });
      });
    }
    
    // Write all data at once - START AT ROW 16
    if (rowsData.length > 0) {
      const maxRows = Math.min(rowsData.length, 47);
      const range = teamSheet.getRange(16, 1, maxRows, 6);
      range.setValues(rowsData.slice(0, maxRows));
      
      // Apply formatting in batch
      formats.forEach(formatInfo => {
        const row = formatInfo.row;
        
        if (formatInfo.isHeader) {
          // Format header rows
          teamSheet.getRange(row, 1, 1, 6).merge()
            .setFontWeight('bold')
            .setBackground('#F5F5F5')
            .setFontStyle('italic');
        } else {
          // Format date column if needed
          if (formatInfo.hasDate) {
            teamSheet.getRange(row, 5).setNumberFormat('yyyy-MM-dd');
          }
          
          // Format points column
          teamSheet.getRange(row, 4).setNumberFormat('0');
          
          // Color code based on source
          let backgroundColor = '';
          if (formatInfo.source === 'PMM') {
            backgroundColor = CONFIG.COLORS.PMM_BLUE;
          } else if (formatInfo.source === 'Workstream') {
            backgroundColor = CONFIG.COLORS.LIGHT_YELLOW;
          } else if (formatInfo.source === 'Team') {
            backgroundColor = '#E1F5FE';
          }
          
          if (backgroundColor) {
            teamSheet.getRange(row, 6).setBackground(backgroundColor);
          }
        }
      });
      
      // Apply borders to entire range
      range.setBorder(true, true, true, true, true, true);
    }
  });
  
  SpreadsheetApp.getUi().alert('Team assignments refreshed successfully!');
}

// ==================== VALIDATION FUNCTION ====================
function validateWorkstreamAllocations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  const workstreams = getWorkstreamNames();
  
  // Create a map of team capacities
  const teamCapacities = {};
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (teamSheet) {
      const netCapacity = teamSheet.getRange('D7').getValue() || 0;
      teamCapacities[teamName] = {
        net: netCapacity,
        allocated: 0
      };
    }
  });
  
  // Check each workstream's asset allocations
  const warnings = [];
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    // Check asset rows (46-95)
    const assetData = wsSheet.getRange(46, 1, 50, 6).getValues();
    assetData.forEach((row, i) => {
      const [description, goLiveDate, size, points, origin, teamAssignment] = row;
      if (description && teamAssignment && points > 0) {
        if (teamCapacities[teamAssignment]) {
          teamCapacities[teamAssignment].allocated += points;
        }
      }
    });
  });
  
  // Also check team-initiated work (but NOT creative planning)
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (teamSheet) {
      const teamData = teamSheet.getRange(65, 1, 30, 4).getValues();
      teamData.forEach(row => {
        if (row[1] && row[3] > 0) { // If description exists and points > 0
          teamCapacities[teamName].allocated += row[3];
        }
      });
      
      // REMOVED: No longer add creative planning to allocated amount
      // Creative planning reduces capacity, it's not work to be done
    }
  });
  
  // Check for over-allocations
  Object.keys(teamCapacities).forEach(teamName => {
    const team = teamCapacities[teamName];
    if (team.allocated > team.net) {
      warnings.push(`${teamName}: Allocated ${team.allocated} pts but only has ${team.net} pts capacity (OVER by ${team.allocated - team.net} pts)`);
    }
  });
  
  return warnings;
}

// ==================== ASSET SECTION ====================
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
  setCell(sheet, 'D42', '=SUMIF(D46:D95,">0",D46:D95)', {
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
  
  const originValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Workstream', 'PMM'], true)
    .setAllowInvalid(false)
    .build();
  
  const teamValidation = teamNames.length > 0 ? 
    SpreadsheetApp.newDataValidation()
      .requireValueInList(teamNames, true)
      .setAllowInvalid(false)
      .build() : null;
  
  const today = new Date();
  
  for (let i = 0; i < 50; i++) {
    const row = 46 + i;
    
    setCell(sheet, `A${row}`, '', { background: '#FFFFFF' });
    
    setCell(sheet, `B${row}`, today, { background: '#FFF9C4' });
    sheet.getRange(row, 2).setNumberFormat('yyyy-MM-dd');
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
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
  
  // Conditional formatting for Origin column
  const originRange = sheet.getRange('E46:E95');
  
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
  
  const existingRules = sheet.getConditionalFormatRules();
  existingRules.push(wsRule, pmmRule);
  sheet.setConditionalFormatRules(existingRules);
  
  sheet.getRange(45, 1, 51, 6).setBorder(true, true, true, true, true, true);
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

function generateTeamCapacityFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    return CONFIG.DEFAULT_CAPACITY;
  }
  
  const teamRefs = teams.map(team => `'${team} Team'!D7`);
  return `=SUM(${teamRefs.join(',')})`;
}

function updateTotalCapacity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  if (!allocSheet) return;
  
  const formula = generateTeamCapacityFormula();
  allocSheet.getRange('B6').setFormula(formula);
}

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

// ==================== UTILITY FUNCTIONS ====================
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
  
  let totalRow = 10;
  while (allocSheet.getRange(totalRow, 1).getValue() !== 'TOTAL' && totalRow < 20) {
    totalRow++;
  }
  
  allocSheet.insertRowBefore(totalRow);
  setCell(allocSheet, `A${totalRow}`, name, { background: '#FFFFFF' });
  setCell(allocSheet, `B${totalRow}`, 0, {
    format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
  });
  setCell(allocSheet, `C${totalRow}`, `=ROUND($B$6*B${totalRow},0)`, {
    format: '0', background: '#F5F5F5'
  });
  
  allocSheet.getRange(totalRow, 1, 1, 3).setBorder(true, true, true, true, true, false);
  
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
  
  allocSheet.getRange(9, 1, newTotalRow - 8, 3).setBorder(true, true, true, true, true, true);
  
  let nextCol = 7;
  while (allocSheet.getRange(6, nextCol).getValue() && nextCol < 20) {
    nextCol++;
  }
  
  setCell(allocSheet, `${columnToLetter(nextCol)}6`, name, {
    fontWeight: true, background: '#E3F2FD'
  });
  allocSheet.setColumnWidth(nextCol, 80);
  
  for (let row = 7; row <= 21; row++) {
    allocSheet.getRange(row, nextCol).insertCheckboxes();
  }
  
  const wsSheet = ss.insertSheet(name);
  setupWorkstreamTab(wsSheet, name);
  
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
  
  let row = 10;
  while (allocSheet.getRange(row, 1).getValue() !== 'TOTAL') {
    if (allocSheet.getRange(row, 1).getValue() === name) {
      allocSheet.deleteRow(row);
      break;
    }
    row++;
  }
  
  setCell(allocSheet, `B${row}`, `=SUM(B10:B${row-1})`, {
    fontWeight: true, background: '#E0E0E0', format: '0%'
  });
  setCell(allocSheet, `C${row}`, `=SUM(C10:C${row-1})`, {
    fontWeight: true, background: '#E0E0E0', format: '0'
  });
  
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(6, col).getValue() === name) {
      allocSheet.deleteColumn(col);
      break;
    }
  }
  
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
  updateTotalCapacity();
  
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
  updateTotalCapacity();
  ui.alert('Success', `Team "${name}" removed.`, ui.ButtonSet.OK);
}

// Sort manifest functions
function sortManifestByDate() {
  refreshTeamAssignments('date');
}

function sortManifestByWorkstream() {
  refreshTeamAssignments('workstream');
}

// About function
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const message = `Marketing Team Points System

Version: ${VERSION}
Release Date: ${RELEASE_DATE}

FIXES in v13.3:
• Fixed creative planning double-counting issue
• Creative planning now only reduces capacity
• Corrected Total Allocated formula
• Fixed manifest generation

Features:
• Team member management with holidays
• Auto-calculated capacity from team holidays
• PMM strategic priority distribution
• Workstream team priorities with owner info
• Asset planning with T-shirt sizing
• Complete manifest generation

© 2024 Marketing Team`;
  
  ui.alert('About Points System', message, ui.ButtonSet.OK);
}

// ==================== MENU SETUP ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Points System')
    .addItem('🚀 Initial Setup', 'setupPointsSystem')
    .addItem('🔄 Refresh Team Assignments', 'refreshTeamAssignments')
    .addSeparator()
    .addSubMenu(ui.createMenu('Workstreams')
      .addItem('➕ Add Workstream', 'addWorkstream')
      .addItem('➖ Remove Workstream', 'removeWorkstream'))
    .addSubMenu(ui.createMenu('Teams')
      .addItem('➕ Add Team', 'addTeam')
      .addItem('➖ Remove Team', 'removeTeam')
      .addSeparator()
      .addItem('📅 Sort Manifest by Date', 'sortManifestByDate')
      .addItem('📊 Sort Manifest by Workstream', 'sortManifestByWorkstream'))
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}