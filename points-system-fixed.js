/**
 * Marketing Team Points System - Version 8.1
 * Bug fixes: Date formatting, Origin selection, Team-initiated work in manifest
 * New features: Sort manifest by date or workstream
 */

// ==================== CONSTANTS ====================
const CONFIG = {
  DEFAULT_WORKSTREAMS: ['SoMe', 'PUA', 'ASO', 'Portal'],
  DEFAULT_ALLOCATIONS: [0.50, 0.20, 0.05, 0.25],
  DEFAULT_CAPACITY: 100,
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
    GRAY: '#F0F0F0'
  }
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
  
  ss.setActiveSheet(allocationSheet);
  
  SpreadsheetApp.getUi().alert(
    'Points System Setup Complete! 🎉',
    'System ready with asset planning and team assignments.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==================== ALLOCATION TAB ====================
function setupAllocationTab(sheet) {
  // Set column widths
  const widths = [250, 120, 120, 30, 250, 100, 80, 80, 80, 80];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Title
  setCell(sheet, 'A1:C1', 'ALLOCATION TAB - PMM Control Panel', {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_BLUE, fontColor: '#FFFFFF'
  });
  
  setCell(sheet, 'A2:C2', 'Monthly Planning & Resource Allocation', {
    merge: true, fontSize: 11, background: CONFIG.COLORS.LIGHT_BLUE
  });
  
  // Monthly Setup
  setCell(sheet, 'A4', 'MONTHLY SETUP', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  setCell(sheet, 'A5', 'Total Creative Capacity (Points):');
  setCell(sheet, 'B5', CONFIG.DEFAULT_CAPACITY, {
    background: CONFIG.COLORS.LIGHT_YELLOW, border: true, format: '0'
  });
  
  // Workstream Allocation
  setCell(sheet, 'A7', 'WORKSTREAM ALLOCATION', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  
  sheet.getRange('A8:C8').setValues([['Workstream', 'Allocation %', 'Points']])
    .setFontWeight(true).setBackground('#E3F2FD');
  
  // Workstreams
  CONFIG.DEFAULT_WORKSTREAMS.forEach((ws, i) => {
    const row = 9 + i;
    setCell(sheet, `A${row}`, ws);
    setCell(sheet, `B${row}`, CONFIG.DEFAULT_ALLOCATIONS[i], {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    setCell(sheet, `C${row}`, `=ROUND($B$5*B${row},0)`, {
      format: '0', background: '#F5F5F5'
    });
  });
  
  // Total row
  const totalRow = 13;
  ['TOTAL', '=SUM(B9:B12)', '=SUM(C9:C12)'].forEach((val, i) => {
    setCell(sheet, `${String.fromCharCode(65 + i)}${totalRow}`, val, {
      fontWeight: true, background: '#E0E0E0',
      format: i > 0 ? (i === 1 ? '0%' : '0') : null
    });
  });
  
  sheet.getRange(8, 1, 6, 3).setBorder(true, true, true, true, true, true);
  
  // Strategic Priorities
  setCell(sheet, 'E4', 'STRATEGIC PRIORITIES', {
    fontWeight: true, background: CONFIG.COLORS.GRAY
  });
  
  sheet.getRange('E5:J5').setValues([['Priority Name', 'Weight %', 'SoMe', 'PUA', 'ASO', 'Portal']])
    .setFontWeight(true).setBackground('#E3F2FD');
  
  // Sample priorities
  const priorities = [
    ['Q4 Campaign Launch', 0.40],
    ['Brand Awareness Push', 0.30],
    ['Product Feature Release', 0.20],
    ['Holiday Season Prep', 0.10]
  ];
  
  priorities.forEach((p, i) => {
    const row = 6 + i;
    setCell(sheet, `E${row}`, p[0]);
    setCell(sheet, `F${row}`, p[1], {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    // Add checkboxes
    for (let col = 7; col <= 10; col++) {
      sheet.getRange(row, col).insertCheckboxes();
    }
  });
  
  // Empty rows for more priorities
  for (let row = 10; row <= 20; row++) {
    setCell(sheet, `F${row}`, '', {
      format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
    });
    for (let col = 7; col <= 10; col++) {
      sheet.getRange(row, col).insertCheckboxes();
    }
  }
  
  sheet.getRange(5, 5, 16, 6).setBorder(true, true, true, true, true, true);
}

// ==================== WORKSTREAM TAB ====================
function setupWorkstreamTab(sheet, workstreamName) {
  sheet.clear();
  
  // Ensure 6 columns
  adjustColumns(sheet, 6);
  
  // Set widths
  const widths = [400, 120, 100, 80, 120, 120];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Header
  setCell(sheet, 'A1:F1', `${workstreamName.toUpperCase()} WORKSTREAM`, {
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
  setCell(sheet, 'C18', '→');
  setCell(sheet, 'D18', '=ROUND(B18*B2,0)', {
    format: '0', fontWeight: true, background: '#E3F2FD'
  });
  
  // PMM Priorities Section
  setCell(sheet, 'A20:D20', '--- PMM Strategic Priorities (Auto-scaled) ---', {
    merge: true, fontStyle: true, background: '#F5F5F5'
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
  
  // Get current year for default dates
  const currentYear = new Date().getFullYear();
  const defaultDate = new Date(currentYear, 0, 1); // January 1st of current year
  
  for (let i = 0; i < 50; i++) {
    const row = 46 + i;
    
    setCell(sheet, `A${row}`, '', { background: '#FFFFFF' });
    
    // Set default date to current year
    setCell(sheet, `B${row}`, '', { background: '#FFF9C4' });
    sheet.getRange(row, 2).setNumberFormat('yyyy-MM-dd');
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    // T-shirt size formula
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
    // Origin dropdown (Workstream or PMM)
    sheet.getRange(row, 5).setDataValidation(originValidation)
      .setBackground(CONFIG.COLORS.LIGHT_GREEN);
    setCell(sheet, `E${row}`, 'Workstream'); // Default to Workstream
    
    if (teamValidation) {
      sheet.getRange(row, 6).setDataValidation(teamValidation)
        .setBackground(CONFIG.COLORS.LIGHT_PURPLE);
    } else {
      setCell(sheet, `F${row}`, '', { background: CONFIG.COLORS.LIGHT_PURPLE });
    }
  }
  
  sheet.getRange(45, 1, 51, 6).setBorder(true, true, true, true, true, true);
}

// ==================== PMM FORMULAS ====================
function setupPMMFormulas(sheet, workstreamName) {
  const checkboxCol = findCheckboxColumn(workstreamName);
  if (checkboxCol < 0) return;
  
  const colLetter = columnToLetter(checkboxCol);
  
  for (let i = 0; i < 15; i++) {
    const row = 21 + i;
    const allocRow = 6 + i;
    
    // Priority name
    setCell(sheet, `A${row}`, 
      `=IF(Allocation!${colLetter}${allocRow}=TRUE,Allocation!E${allocRow},"")`, 
      { background: CONFIG.COLORS.GRAY });
    
    // Source
    setCell(sheet, `B${row}`, 
      `=IF(A${row}<>"","PMM","")`, 
      { background: CONFIG.COLORS.GRAY });
    
    // Percentage - simplified formula
    const percentFormula = `=IF(A${row}="","",` +
      `IF(SUMPRODUCT(Allocation!${colLetter}6:${colLetter}20*Allocation!F6:F20)=0,0,` +
      `(Allocation!F${allocRow}*Allocation!${colLetter}${allocRow})/` +
      `SUMPRODUCT(Allocation!${colLetter}6:${colLetter}20*Allocation!F6:F20)*B18))`;
    
    setCell(sheet, `C${row}`, percentFormula, 
      { format: '0%', background: CONFIG.COLORS.GRAY });
    
    // Points
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
  
  // Header
  setCell(sheet, 'A1:F1', `${teamName.toUpperCase()} TEAM`, {
    merge: true, fontSize: 16, fontWeight: true,
    background: CONFIG.COLORS.HEADER_PURPLE, fontColor: '#FFFFFF'
  });
  
  // Capacity Section
  setCell(sheet, 'A3:F3', 'TEAM CAPACITY', {
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
  
  setCell(sheet, 'C5', 'Team Capacity:');
  setCell(sheet, 'D5', '=(B4*D4)-B5', {
    fontWeight: true, background: CONFIG.COLORS.LIGHT_GREEN, format: '0'
  });
  
  // Assignment Summary
  setCell(sheet, 'A7:F7', 'ASSIGNMENT SUMMARY', {
    merge: true, fontWeight: true, background: CONFIG.COLORS.LIGHT_PURPLE
  });
  
  setCell(sheet, 'A8', 'Workstream Assigned:');
  setCell(sheet, 'B8', '=SUMIF(F12:F200,"<>Team",D12:D200)', {
    fontWeight: true, fontSize: 14, background: CONFIG.COLORS.LIGHT_ORANGE, format: '0'
  });
  
  setCell(sheet, 'C8', 'Team Initiated:');
  setCell(sheet, 'D8', '=SUMIF(F12:F200,"Team",D12:D200)', {
    fontWeight: true, fontSize: 14, background: '#E1F5FE', format: '0'
  });
  
  setCell(sheet, 'E8', 'Total:');
  setCell(sheet, 'F8', '=B8+D8', {
    fontWeight: true, fontSize: 14, background: '#FFD54F', format: '0'
  });
  
  // Utilization
  setCell(sheet, 'A9', 'Utilization:');
  setCell(sheet, 'B9', '=IF(D5=0,"",F8/D5)', {
    fontWeight: true, fontSize: 14, format: '0%'
  });
  
  setCell(sheet, 'C9:F9', 
    '=IF(F8>D5,"⚠️ OVER CAPACITY by "&(F8-D5)&" points",IF(F8=D5,"✅ FULL","✅ "&(D5-F8)&" points available"))', {
    merge: true, fontWeight: true
  });
  
  // Conditional formatting for over capacity
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F$8>$D$5')
    .setFontColor('#FF0000')
    .setRanges([sheet.getRange('F8')])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Table headers
  sheet.getRange('A11:F11').setValues([['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source']])
    .setFontWeight(true).setBackground('#E1BEE7');
  
  setCell(sheet, 'A12', 'Click "Refresh Team Assignments" to load...', {
    fontStyle: true, fontColor: '#666666'
  });
  
  // Team-initiated section
  setCell(sheet, 'A50:F50', '--- TEAM-INITIATED WORK ---', {
    merge: true, fontWeight: true, fontStyle: true, background: '#E1F5FE'
  });
  
  // Team rows
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CONFIG.TSHIRT_SIZES), true)
    .setAllowInvalid(false)
    .build();
  
  const currentYear = new Date().getFullYear();
  
  for (let i = 0; i < 30; i++) {
    const row = 51 + i;
    
    setCell(sheet, `A${row}`, teamName, { background: CONFIG.COLORS.GRAY });
    setCell(sheet, `B${row}`, '', { background: CONFIG.COLORS.LIGHT_YELLOW });
    
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    const sizeFormula = `=IF(C${row}="","",SWITCH(C${row},` +
      Object.entries(CONFIG.TSHIRT_SIZES)
        .map(([size, points]) => `"${size}",${points}`)
        .join(',') + ',0))';
    setCell(sheet, `D${row}`, sizeFormula, { format: '0', background: CONFIG.COLORS.GRAY });
    
    setCell(sheet, `E${row}`, '', { background: '#FFF9C4' });
    sheet.getRange(row, 5).setNumberFormat('yyyy-MM-dd');
    setCell(sheet, `F${row}`, 'Team', { background: '#E1F5FE' });
  }
  
  // Borders
  sheet.getRange(3, 1, 3, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(7, 1, 3, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(50, 1, 31, 6).setBorder(true, true, true, true, true, true);
}

// ==================== TEAM ASSIGNMENTS ====================
function refreshTeamAssignments(sortBy = 'date') {
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
      // Clear workstream assignments area only (rows 12-49)
      teamSheet.getRange(12, 1, 38, 6).clear();
    }
  });
  
  // Collect from workstreams
  const workstreams = getWorkstreamNames();
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    const data = wsSheet.getRange(46, 1, 50, 6).getValues();
    data.forEach((row, i) => {
      const [description, goLiveDate, tShirtSize, points, origin, teamAssignment] = row;
      
      if (description && teamAssignment && teams.includes(teamAssignment)) {
        // Format date properly
        let formattedDate = goLiveDate;
        if (goLiveDate instanceof Date) {
          formattedDate = Utilities.formatDate(goLiveDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        teamAssignments[teamAssignment].push({
          origin: wsName,
          description,
          size: tShirtSize,
          points,
          goLiveDate: formattedDate,
          source: origin
        });
      }
    });
  });
  
  // Collect team-initiated work
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const teamData = teamSheet.getRange(51, 1, 30, 6).getValues();
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
      setCell(teamSheet, 'A12', 'No assignments', {
        fontStyle: true, fontColor: '#666666'
      });
      return;
    }
    
    // Sort based on preference
    if (sortBy === 'date') {
      // Sort by go live date (earliest first), then by workstream
      assignments.sort((a, b) => {
        if (a.goLiveDate && b.goLiveDate) {
          const dateCompare = a.goLiveDate.localeCompare(b.goLiveDate);
          if (dateCompare !== 0) return dateCompare;
        } else if (a.goLiveDate) {
          return -1; // Items with dates come first
        } else if (b.goLiveDate) {
          return 1;
        }
        return a.origin.localeCompare(b.origin);
      });
    } else {
      // Default: Group by workstream
      const grouped = {};
      assignments.forEach(a => {
        if (!grouped[a.origin]) grouped[a.origin] = [];
        grouped[a.origin].push(a);
      });
      
      let currentRow = 12;
      Object.keys(grouped).sort().forEach(wsName => {
        if (currentRow >= 50) return;
        
        // Workstream header
        setCell(teamSheet, `A${currentRow}:F${currentRow}`, `--- ${wsName} ---`, {
          merge: true, fontWeight: true, background: '#F5F5F5', fontStyle: true
        });
        currentRow++;
        
        // Assignments
        grouped[wsName].forEach(a => {
          if (currentRow >= 50) return;
          
          teamSheet.getRange(currentRow, 1, 1, 6).setValues([[
            a.origin,
            a.description,
            a.size,
            a.points,
            a.goLiveDate || '',
            a.source
          ]]);
          
          if (a.goLiveDate) {
            teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
          }
          teamSheet.getRange(currentRow, 4).setNumberFormat('0');
          currentRow++;
        });
      });
      
      if (currentRow > 12) {
        teamSheet.getRange(12, 1, currentRow - 12, 6)
          .setBorder(true, true, true, true, true, false);
      }
      return; // Exit here for workstream grouping
    }
    
    // Write all assignments sorted by date
    let currentRow = 12;
    assignments.forEach(a => {
      if (currentRow >= 50) return;
      
      teamSheet.getRange(currentRow, 1, 1, 6).setValues([[
        a.origin,
        a.description,
        a.size,
        a.points,
        a.goLiveDate || '',
        a.source
      ]]);
      
      if (a.goLiveDate) {
        teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
      }
      teamSheet.getRange(currentRow, 4).setNumberFormat('0');
      currentRow++;
    });
    
    if (currentRow > 12) {
      teamSheet.getRange(12, 1, currentRow - 12, 6)
        .setBorder(true, true, true, true, true, false);
    }
  });
  
  SpreadsheetApp.getUi().alert('Team assignments refreshed successfully!');
}

// Sort manifest by date
function sortManifestByDate() {
  refreshTeamAssignments('date');
}

// Sort manifest by workstream
function sortManifestByWorkstream() {
  refreshTeamAssignments('workstream');
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
  let totalRow = 9;
  while (allocSheet.getRange(totalRow, 1).getValue() !== 'TOTAL' && totalRow < 20) {
    totalRow++;
  }
  
  // Insert in allocation table
  allocSheet.insertRowBefore(totalRow);
  setCell(allocSheet, `A${totalRow}`, name);
  setCell(allocSheet, `B${totalRow}`, 0, {
    format: '0%', background: CONFIG.COLORS.LIGHT_YELLOW
  });
  setCell(allocSheet, `C${totalRow}`, `=ROUND($B$5*B${totalRow},0)`, {
    format: '0', background: '#F5F5F5'
  });
  
  // Update TOTAL formulas
  const newTotalRow = totalRow + 1;
  setCell(allocSheet, `B${newTotalRow}`, `=SUM(B9:B${totalRow})`);
  setCell(allocSheet, `C${newTotalRow}`, `=SUM(C9:C${totalRow})`);
  
  // Add checkbox column
  let nextCol = 7;
  while (allocSheet.getRange(5, nextCol).getValue() && nextCol < 20) {
    nextCol++;
  }
  
  setCell(allocSheet, `${columnToLetter(nextCol)}5`, name, {
    fontWeight: true, background: '#E3F2FD'
  });
  allocSheet.setColumnWidth(nextCol, 80);
  
  for (let row = 6; row <= 20; row++) {
    allocSheet.getRange(row, nextCol).insertCheckboxes();
  }
  
  // Create workstream sheet
  const wsSheet = ss.insertSheet(name);
  setupWorkstreamTab(wsSheet, name);
  
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
  let row = 9;
  while (allocSheet.getRange(row, 1).getValue() !== 'TOTAL') {
    if (allocSheet.getRange(row, 1).getValue() === name) {
      allocSheet.deleteRow(row);
      break;
    }
    row++;
  }
  
  // Update TOTAL formulas
  setCell(allocSheet, `B${row}`, `=SUM(B9:B${row-1})`);
  setCell(allocSheet, `C${row}`, `=SUM(C9:C${row-1})`);
  
  // Remove column
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(5, col).getValue() === name) {
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
  ui.alert('Success', `Team "${name}" removed.`, ui.ButtonSet.OK);
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
    if (allocSheet.getRange(5, col).getValue() === workstreamName) {
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
  
  let row = 9;
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
    .addToUi();
}