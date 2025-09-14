/**
 * Marketing Team Points System - Version 6.1 with Team Assignments
 * Complete script with all functionality
 */

// ==================== MAIN SETUP FUNCTION ====================

function setupPointsSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete all existing sheets except the first one
  const sheets = ss.getSheets();
  sheets.forEach((sheet, index) => {
    if (index > 0) {
      ss.deleteSheet(sheet);
    }
  });
  
  // Clear and rename first sheet to Allocation
  const allocationSheet = sheets[0];
  allocationSheet.clear();
  allocationSheet.setName('Allocation');
  
  // Set up the Allocation tab
  setupAllocationTab(allocationSheet);
  
  // Create default workstream tabs with asset planning
  const defaultWorkstreams = ['SoMe', 'PUA', 'ASO', 'Portal'];
  defaultWorkstreams.forEach(name => {
    const wsSheet = ss.insertSheet(name);
    setupWorkstreamTabWithAssets(wsSheet, name);
  });
  
  // Create default team - just Creative
  const teamSheet = ss.insertSheet('Creative Team');
  setupTeamTab(teamSheet, 'Creative');
  
  // Set the Allocation tab as active
  ss.setActiveSheet(allocationSheet);
  
  SpreadsheetApp.getUi().alert(
    'Points System Setup Complete! 🎉',
    'System ready with asset planning and team assignments.\nDefault team created: Creative',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==================== ALLOCATION TAB SETUP ====================

function setupAllocationTab(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 250); // A
  sheet.setColumnWidth(2, 120); // B
  sheet.setColumnWidth(3, 120); // C
  sheet.setColumnWidth(4, 30);  // D (spacer)
  sheet.setColumnWidth(5, 250); // E
  sheet.setColumnWidth(6, 100); // F
  sheet.setColumnWidth(7, 80);  // G - SoMe
  sheet.setColumnWidth(8, 80);  // H - PUA
  sheet.setColumnWidth(9, 80);  // I - ASO
  sheet.setColumnWidth(10, 80); // J - Portal
  
  // Title and Header
  sheet.getRange('A1:C1').merge()
    .setValue('ALLOCATION TAB - PMM Control Panel')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');
  
  sheet.getRange('A2:C2').merge()
    .setValue('Monthly Planning & Resource Allocation')
    .setFontSize(11)
    .setBackground('#E8F0FE');
  
  // Monthly Setup Section
  sheet.getRange('A4').setValue('MONTHLY SETUP').setFontWeight('bold').setBackground('#F0F0F0');
  sheet.getRange('A5').setValue('Total Creative Capacity (Points):');
  
  // Manual value for capacity (simplified)
  sheet.getRange('B5').setValue(100)
    .setBackground('#FFF3E0')
    .setBorder(true, true, true, true, false, false)
    .setNumberFormat('0');
  
  // Workstream Allocation Table
  sheet.getRange('A7').setValue('WORKSTREAM ALLOCATION').setFontWeight('bold').setBackground('#F0F0F0');
  
  // Headers
  sheet.getRange('A8:C8').setValues([['Workstream', 'Allocation %', 'Points']])
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // Default workstreams
  const workstreams = [
    ['SoMe', 0.50],
    ['PUA', 0.20],
    ['ASO', 0.05],
    ['Portal', 0.25]
  ];
  
  workstreams.forEach((ws, i) => {
    const row = 9 + i;
    sheet.getRange(row, 1).setValue(ws[0]);
    sheet.getRange(row, 2).setValue(ws[1])
      .setNumberFormat('0%')
      .setBackground('#FFF3E0');
    sheet.getRange(row, 3).setFormula(`=ROUND($B$5*B${row},0)`)
      .setNumberFormat('0')
      .setBackground('#F5F5F5');
  });
  
  // Total row
  const totalRow = 13;
  sheet.getRange(totalRow, 1).setValue('TOTAL')
    .setFontWeight('bold')
    .setBackground('#E0E0E0');
  sheet.getRange(totalRow, 2).setFormula('=SUM(B9:B12)')
    .setNumberFormat('0%')
    .setFontWeight('bold')
    .setBackground('#E0E0E0');
  sheet.getRange(totalRow, 3).setFormula('=SUM(C9:C12)')
    .setNumberFormat('0')
    .setFontWeight('bold')
    .setBackground('#E0E0E0');
  
  // Add borders to allocation table
  sheet.getRange(8, 1, 6, 3).setBorder(true, true, true, true, true, true);
  
  // Strategic Priorities Section
  sheet.getRange('E4').setValue('STRATEGIC PRIORITIES').setFontWeight('bold').setBackground('#F0F0F0');
  
  // Priority headers
  sheet.getRange('E5:J5').setValues([['Priority Name', 'Weight %', 'SoMe', 'PUA', 'ASO', 'Portal']])
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // Add sample priorities with checkboxes
  const priorities = [
    ['Q4 Campaign Launch', 0.40],
    ['Brand Awareness Push', 0.30],
    ['Product Feature Release', 0.20],
    ['Holiday Season Prep', 0.10]
  ];
  
  priorities.forEach((priority, i) => {
    const row = 6 + i;
    sheet.getRange(row, 5).setValue(priority[0]);
    sheet.getRange(row, 6).setValue(priority[1])
      .setNumberFormat('0%')
      .setBackground('#FFF3E0');
    // Insert checkboxes for each workstream
    sheet.getRange(row, 7).insertCheckboxes();
    sheet.getRange(row, 8).insertCheckboxes();
    sheet.getRange(row, 9).insertCheckboxes();
    sheet.getRange(row, 10).insertCheckboxes();
  });
  
  // Add empty rows for more priorities
  for (let row = 10; row <= 20; row++) {
    sheet.getRange(row, 6).setNumberFormat('0%').setBackground('#FFF3E0');
    sheet.getRange(row, 7).insertCheckboxes();
    sheet.getRange(row, 8).insertCheckboxes();
    sheet.getRange(row, 9).insertCheckboxes();
    sheet.getRange(row, 10).insertCheckboxes();
  }
  
  // Add borders to priorities table
  sheet.getRange(5, 5, 16, 6).setBorder(true, true, true, true, true, true);
}

// ==================== WORKSTREAM TAB WITH ASSETS ====================

function setupWorkstreamTabWithAssets(sheet, workstreamName) {
  sheet.clear();
  
  // Set to 6 columns for team assignment
  const maxCols = sheet.getMaxColumns();
  if (maxCols > 6) {
    sheet.deleteColumns(7, maxCols - 6);
  } else if (maxCols < 6) {
    sheet.insertColumnsAfter(maxCols, 6 - maxCols);
  }
  
  // Set column widths
  sheet.setColumnWidth(1, 400); // A - Description
  sheet.setColumnWidth(2, 120); // B - Go Live Date
  sheet.setColumnWidth(3, 100); // C - T-Shirt Size
  sheet.setColumnWidth(4, 80);  // D - Cost
  sheet.setColumnWidth(5, 120); // E - Origin
  sheet.setColumnWidth(6, 120); // F - Team Assignment
  
  // Header Section
  sheet.getRange('A1:F1').merge()
    .setValue(`${workstreamName.toUpperCase()} WORKSTREAM`)
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');
  
  // Budget Summary Section
  sheet.getRange('A2').setValue('Total Points Allocated:');
  sheet.getRange('B2').setFormula(`=IFERROR(VLOOKUP("${workstreamName}",Allocation!A:C,3,FALSE),0)`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#E8F5E9')
    .setNumberFormat('0');
  
  sheet.getRange('A3').setValue('Points Spent on Assets:');
  sheet.getRange('B3').setFormula('=SUMIF(D46:D95,">0")')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#FFE0B2')
    .setNumberFormat('0');
  
  sheet.getRange('C2').setValue('Remaining:');
  sheet.getRange('D2').setFormula('=B2-B3')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#E1F5FE')
    .setNumberFormat('0');
  
  // Priorities Table Header
  sheet.getRange('A5:D5').setValues([['Priority Name', 'Source', 'Allocation %', 'Points']])
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // Workstream Priorities Section
  sheet.getRange('A6:D6').merge()
    .setValue('--- Workstream Team Priorities (Direct %) ---')
    .setFontStyle('italic')
    .setBackground('#FFF9C4');
  
  // Add rows for workstream priorities
  for (let i = 0; i < 10; i++) {
    const row = 7 + i;
    sheet.getRange(row, 1).setBackground('#FFF3E0');
    sheet.getRange(row, 2).setValue('Workstream');
    sheet.getRange(row, 3).setNumberFormat('0%').setBackground('#FFF3E0');
    sheet.getRange(row, 4).setFormula(
      `=IF(C${row}="","",ROUND(C${row}*$B$2,0))`
    ).setNumberFormat('0').setBackground('#E8F5E9');
  }
  
  // Remaining capacity indicator
  sheet.getRange('A18').setValue('Remaining for PMM:').setFontWeight('bold');
  sheet.getRange('B18').setFormula('=MAX(0,100%-SUMIF(C7:C16,">0"))')
    .setNumberFormat('0%')
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  sheet.getRange('C18').setValue('→');
  sheet.getRange('D18').setFormula('=ROUND(B18*B2,0)')
    .setNumberFormat('0')
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // PMM Priorities Section
  sheet.getRange('A20:D20').merge()
    .setValue('--- PMM Strategic Priorities (Auto-scaled) ---')
    .setFontStyle('italic')
    .setBackground('#F5F5F5');
  
  // Setup PMM priority formulas
  setupPMMPriorityFormulas(sheet, workstreamName);
  
  // Summary Section
  sheet.getRange('A37').setValue('WORKSTREAM %:').setFontWeight('bold');
  sheet.getRange('B37').setFormula('=SUMIF(C7:C16,">0")')
    .setNumberFormat('0%')
    .setFontWeight('bold');
  
  sheet.getRange('C37').setValue('PMM %:').setFontWeight('bold');
  sheet.getRange('D37').setFormula('=SUMIF(C21:C35,">0")')
    .setNumberFormat('0%')
    .setFontWeight('bold');
  
  sheet.getRange('A38').setValue('TOTAL POINTS:').setFontWeight('bold');
  sheet.getRange('B38').setFormula('=SUM(D7:D16,D21:D35)')
    .setNumberFormat('0')
    .setFontWeight('bold');
  
  // Add borders to priorities section
  sheet.getRange(5, 1, 12, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(18, 1, 1, 4).setBorder(true, true, true, true, false, false);
  sheet.getRange(20, 1, 16, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(37, 1, 2, 4).setBorder(true, true, true, true, false, false);
  
  // ==================== ASSET PLANNING SECTION ====================
  
  // Asset Planning Header
  sheet.getRange('A41:F41').merge()
    .setValue('ASSET PLANNING')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#FF9800')
    .setFontColor('#FFFFFF');
  
  // Budget info row
  sheet.getRange('A42').setValue('Budget:');
  sheet.getRange('B42').setFormula('=B2')
    .setFontWeight('bold')
    .setNumberFormat('0')
    .setBackground('#E8F5E9');
  sheet.getRange('C42').setValue('Spent:');
  sheet.getRange('D42').setFormula('=SUMIF(D46:D95,">0")')
    .setFontWeight('bold')
    .setNumberFormat('0')
    .setBackground('#FFE0B2');
  sheet.getRange('E42:F42').merge()
    .setFormula(
    '=IF(B42-D42<0,"⚠️ OVER by "&ABS(B42-D42),"💰 "&(B42-D42)&" remaining")'
  ).setFontWeight('bold');
  
  // T-shirt size legend
  sheet.getRange('A43:F43').merge()
    .setValue('T-Shirt Sizes: XS=1 point | S=3 points | M=5 points | L=13 points | XL=21 points')
    .setFontSize(10)
    .setFontStyle('italic')
    .setBackground('#FFF3E0');
  
  // Asset table headers
  sheet.getRange('A45').setValue('Asset Description')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  sheet.getRange('B45').setValue('Go Live Date')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  sheet.getRange('C45').setValue('T-Shirt Size')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  sheet.getRange('D45').setValue('Cost')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  sheet.getRange('E45').setValue('Origin')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  sheet.getRange('F45').setValue('Team')
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  
  // Get list of teams for dropdown
  const teamNames = getTeamNames();
  
  // Add 50 rows for assets with Origin set to Workstream
  for (let i = 0; i < 50; i++) {
    const row = 46 + i;
    
    // Asset Description
    sheet.getRange(row, 1).setBackground('#FFFFFF');
    
    // Go Live Date - ISO format for Jira
    sheet.getRange(row, 2).setBackground('#FFF9C4')
      .setNumberFormat('yyyy-mm-dd');
    
    // T-Shirt Size dropdown
    const sizeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['XS', 'S', 'M', 'L', 'XL'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    // Cost (Points) with formula
    sheet.getRange(row, 4).setFormula(
      `=IF(C${row}="","",SWITCH(C${row},"XS",1,"S",3,"M",5,"L",13,"XL",21,0))`
    ).setNumberFormat('0')
     .setBackground('#F0F0F0');
    
    // Origin - defaults to Workstream
    sheet.getRange(row, 5).setValue('Workstream')
      .setBackground('#E8F5E9');
    
    // Team Assignment dropdown
    if (teamNames.length > 0) {
      const teamValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(teamNames, true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(row, 6).setDataValidation(teamValidation)
        .setBackground('#F3E5F5');
    } else {
      sheet.getRange(row, 6).setBackground('#F3E5F5');
    }
  }
  
  // Add borders to asset table
  sheet.getRange(45, 1, 51, 6).setBorder(true, true, true, true, true, true);
}

// ==================== SETUP PMM PRIORITY FORMULAS ====================

function setupPMMPriorityFormulas(sheet, workstreamName) {
  const checkboxCol = findCheckboxColumn(workstreamName);
  
  if (checkboxCol > 0) {
    for (let i = 0; i < 15; i++) {
      const currentRow = 21 + i;
      const allocRow = 6 + i;
      
      // Priority name formula
      sheet.getRange(currentRow, 1).setFormula(
        `=IF(Allocation!${columnToLetter(checkboxCol)}${allocRow}=TRUE,Allocation!E${allocRow},"")`
      ).setBackground('#F0F0F0');
      
      // Source formula
      sheet.getRange(currentRow, 2).setFormula(
        `=IF(A${currentRow}<>"","PMM","")`
      ).setBackground('#F0F0F0');
      
      // Percentage formula
      const percentFormula = `=IF(A${currentRow}="","",` +
        `IF(SUMPRODUCT(Allocation!${columnToLetter(checkboxCol)}6:${columnToLetter(checkboxCol)}20*Allocation!F6:F20)=0,0,` +
        `(Allocation!F${allocRow}*Allocation!${columnToLetter(checkboxCol)}${allocRow})/` +
        `SUMPRODUCT(Allocation!${columnToLetter(checkboxCol)}6:${columnToLetter(checkboxCol)}20*Allocation!F6:F20)*B18))`;
      
      sheet.getRange(currentRow, 3).setFormula(percentFormula)
        .setNumberFormat('0%')
        .setBackground('#F0F0F0');
      
      // Points formula
      sheet.getRange(currentRow, 4).setFormula(
        `=IF(C${currentRow}="","",ROUND(C${currentRow}*$B$2,0))`
      ).setNumberFormat('0').setBackground('#E8F5E9');
    }
  }
}

// ==================== TEAM TAB SETUP ====================

function setupTeamTab(sheet, teamName) {
  sheet.clear();
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // A - Origin
  sheet.setColumnWidth(2, 400); // B - Description
  sheet.setColumnWidth(3, 100); // C - T-Shirt Size
  sheet.setColumnWidth(4, 80);  // D - Points
  sheet.setColumnWidth(5, 120); // E - Go Live Date
  sheet.setColumnWidth(6, 120); // F - Source Type
  
  // Header Section
  sheet.getRange('A1:F1').merge()
    .setValue(`${teamName.toUpperCase()} TEAM`)
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#9C27B0')
    .setFontColor('#FFFFFF');
  
  // Team Capacity Section
  sheet.getRange('A3:F3').merge()
    .setValue('TEAM CAPACITY')
    .setFontWeight('bold')
    .setBackground('#F3E5F5');
  
  // Capacity inputs
  sheet.getRange('A4').setValue('Team Members:');
  sheet.getRange('B4').setValue(5)
    .setBackground('#FFF3E0')
    .setNumberFormat('0');
  
  sheet.getRange('C4').setValue('Working Days/Month:');
  sheet.getRange('D4').setValue(20)
    .setBackground('#FFF3E0')
    .setNumberFormat('0');
  
  sheet.getRange('A5').setValue('Total Days Off (All Members):');
  sheet.getRange('B5').setValue(0)
    .setBackground('#FFF3E0')
    .setNumberFormat('0')
    .setNote('Enter total number of days off across all team members');
  
  // Capacity calculation
  sheet.getRange('C5').setValue('Team Capacity:');
  sheet.getRange('D5').setFormula('=(B4*D4)-B5')
    .setFontWeight('bold')
    .setBackground('#E8F5E9')
    .setNumberFormat('0');
  
  // Assignment Summary
  sheet.getRange('A7:F7').merge()
    .setValue('ASSIGNMENT SUMMARY')
    .setFontWeight('bold')
    .setBackground('#F3E5F5');
  
  sheet.getRange('A8').setValue('Workstream Assigned:');
  sheet.getRange('B8').setFormula('=SUMIF(F12:F200,"<>Team",D12:D200)')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#FFE0B2')
    .setNumberFormat('0');
  
  sheet.getRange('C8').setValue('Team Initiated:');
  sheet.getRange('D8').setFormula('=SUMIF(F12:F200,"Team",D12:D200)')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#E1F5FE')
    .setNumberFormat('0');
  
  sheet.getRange('E8').setValue('Total:');
  sheet.getRange('F8').setFormula('=B8+D8')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#FFD54F')
    .setNumberFormat('0');
  
  // Utilization calculation
  sheet.getRange('A9').setValue('Utilization:');
  sheet.getRange('B9').setFormula('=IF(D5=0,"",F8/D5)')
    .setFontWeight('bold')
    .setFontSize(14)
    .setNumberFormat('0%');
  
  // Status indicator
  sheet.getRange('C9:F9').merge()
    .setFormula(
    '=IF(F8>D5,"⚠️ OVER CAPACITY by "&(F8-D5)&" points",IF(F8=D5,"✅ FULL","✅ "&(D5-F8)&" points available"))'
  ).setFontWeight('bold');
  
  // Apply red color if over capacity
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F$8>$D$5')
    .setFontColor('#FF0000')
    .setRanges([sheet.getRange('F8')])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Assignments Table Header
  sheet.getRange('A11:F11').setValues([['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source']])
    .setFontWeight('bold')
    .setBackground('#E1BEE7');
  
  // Section for workstream assignments (will be populated by refresh)
  sheet.getRange('A12').setValue('Click "Refresh Team Assignments" to load workstream assignments...')
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  // Team-initiated work section
  sheet.getRange('A50:F50').merge()
    .setValue('--- TEAM-INITIATED WORK ---')
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setBackground('#E1F5FE');
  
  // Add 30 rows for team-initiated work
  for (let i = 0; i < 30; i++) {
    const row = 51 + i;
    
    // Origin - set to team name
    sheet.getRange(row, 1).setValue(teamName)
      .setBackground('#F0F0F0');
    
    // Description - editable
    sheet.getRange(row, 2).setBackground('#FFF3E0');
    
    // T-Shirt Size dropdown
    const sizeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['XS', 'S', 'M', 'L', 'XL'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, 3).setDataValidation(sizeValidation)
      .setBackground('#E1F5FE');
    
    // Cost (Points) with formula
    sheet.getRange(row, 4).setFormula(
      `=IF(C${row}="","",SWITCH(C${row},"XS",1,"S",3,"M",5,"L",13,"XL",21,0))`
    ).setNumberFormat('0')
     .setBackground('#F0F0F0');
    
    // Go Live Date
    sheet.getRange(row, 5).setBackground('#FFF9C4')
      .setNumberFormat('yyyy-mm-dd');
    
    // Source - set to Team
    sheet.getRange(row, 6).setValue('Team')
      .setBackground('#E1F5FE');
  }
  
  // Add borders
  sheet.getRange(3, 1, 3, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(7, 1, 3, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(11, 1, 1, 6).setBorder(true, true, true, true, true, true);
  sheet.getRange(50, 1, 31, 6).setBorder(true, true, true, true, true, true);
}

// ==================== TEAM ASSIGNMENT REFRESH ====================

function refreshTeamAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No teams found. Please add teams first.');
    return;
  }
  
  // Clear existing workstream assignments for all teams (preserve team-initiated work)
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (teamSheet) {
      // Clear only rows 12-49 (workstream assignments area)
      teamSheet.getRange(12, 1, 38, 6).clear();
    }
  });
  
  // Get all workstreams
  const workstreams = getWorkstreamNames();
  
  // Collect assignments from each workstream
  const teamAssignments = {};
  teams.forEach(team => {
    teamAssignments[team] = [];
  });
  
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    // Check asset planning rows (46-95)
    for (let row = 46; row <= 95; row++) {
      const description = wsSheet.getRange(row, 1).getValue();
      const goLiveDate = wsSheet.getRange(row, 2).getValue();
      const tShirtSize = wsSheet.getRange(row, 3).getValue();
      const points = wsSheet.getRange(row, 4).getValue();
      const origin = wsSheet.getRange(row, 5).getValue(); // Get origin (Workstream/PMM)
      const teamAssignment = wsSheet.getRange(row, 6).getValue();
      
      if (description && teamAssignment && teams.includes(teamAssignment)) {
        teamAssignments[teamAssignment].push({
          origin: wsName,
          description: description,
          size: tShirtSize,
          points: points,
          goLiveDate: goLiveDate,
          source: origin // Pass through the origin
        });
      }
    }
  });
  
  // Write assignments to team sheets
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const assignments = teamAssignments[teamName];
    if (assignments.length === 0) {
      teamSheet.getRange('A12').setValue('No workstream assignments')
        .setFontStyle('italic')
        .setFontColor('#666666');
      return;
    }
    
    // Group by workstream
    const groupedAssignments = {};
    assignments.forEach(a => {
      if (!groupedAssignments[a.origin]) {
        groupedAssignments[a.origin] = [];
      }
      groupedAssignments[a.origin].push(a);
    });
    
    let currentRow = 12;
    
    // Write assignments grouped by workstream
    Object.keys(groupedAssignments).sort().forEach(wsName => {
      // Add workstream header
      teamSheet.getRange(currentRow, 1, 1, 6).merge()
        .setValue(`--- ${wsName} ---`)
        .setFontWeight('bold')
        .setBackground('#F5F5F5')
        .setFontStyle('italic');
      currentRow++;
      
      // Add assignments
      groupedAssignments[wsName].forEach(assignment => {
        if (currentRow >= 50) return; // Don't overwrite team-initiated section
        
        teamSheet.getRange(currentRow, 1).setValue(wsName);
        teamSheet.getRange(currentRow, 2).setValue(assignment.description);
        teamSheet.getRange(currentRow, 3).setValue(assignment.size);
        teamSheet.getRange(currentRow, 4).setValue(assignment.points).setNumberFormat('0');
        
        // Format date for Jira (ISO format)
        if (assignment.goLiveDate) {
          const date = new Date(assignment.goLiveDate);
          teamSheet.getRange(currentRow, 5).setValue(date).setNumberFormat('yyyy-mm-dd');
        }
        
        // Set source (Workstream/PMM)
        teamSheet.getRange(currentRow, 6).setValue(assignment.source);
        
        currentRow++;
      });
    });
    
    // Add borders to the data
    if (currentRow > 12) {
      teamSheet.getRange(12, 1, currentRow - 12, 6).setBorder(true, true, true, true, true, false);
    }
  });
  
  SpreadsheetApp.getUi().alert('Team assignments refreshed successfully!');
}

// ==================== TEAM MANAGEMENT FUNCTIONS ====================

function addTeam() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add New Team',
    'Enter the name for the new team:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const teamName = response.getResponseText().trim();
  if (!teamName) {
    ui.alert('Error', 'Team name cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = teamName + ' Team';
  
  if (ss.getSheetByName(sheetName)) {
    ui.alert('Error', `Team "${teamName}" already exists.`, ui.ButtonSet.OK);
    return;
  }
  
  // Create the team sheet
  const teamSheet = ss.insertSheet(sheetName);
  setupTeamTab(teamSheet, teamName);
  
  // Update team dropdowns in all workstreams
  updateTeamDropdowns();
  
  ui.alert('Success', `Team "${teamName}" added successfully.`, ui.ButtonSet.OK);
}

function removeTeam() {
  const ui = SpreadsheetApp.getUi();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    ui.alert('Error', 'No teams to remove.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt(
    'Remove Team',
    'Enter the name of the team to remove:\n\nAvailable: ' + teams.join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const teamName = response.getResponseText().trim();
  
  if (!teams.includes(teamName)) {
    ui.alert('Error', `Team "${teamName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = teamName + ' Team';
  const teamSheet = ss.getSheetByName(sheetName);
  
  if (teamSheet) {
    ss.deleteSheet(teamSheet);
  }
  
  // Update team dropdowns in all workstreams
  updateTeamDropdowns();
  
  ui.alert('Success', `Team "${teamName}" removed successfully.`, ui.ButtonSet.OK);
}

// ==================== WORKSTREAM MANAGEMENT ====================

function addWorkstream() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add New Workstream',
    'Enter the name for the new workstream:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const workstreamName = response.getResponseText().trim();
  if (!workstreamName) {
    ui.alert('Error', 'Workstream name cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getSheetByName(workstreamName)) {
    ui.alert('Error', `"${workstreamName}" already exists.`, ui.ButtonSet.OK);
    return;
  }
  
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Find TOTAL row
  let totalRow = 9;
  while (allocSheet.getRange(totalRow, 1).getValue() !== 'TOTAL' && totalRow < 20) {
    totalRow++;
  }
  
  // Insert new row before TOTAL
  allocSheet.insertRowBefore(totalRow);
  
  // Add workstream to allocation table
  allocSheet.getRange(totalRow, 1).setValue(workstreamName);
  allocSheet.getRange(totalRow, 2).setValue(0)
    .setNumberFormat('0%')
    .setBackground('#FFF3E0');
  allocSheet.getRange(totalRow, 3).setFormula(`=ROUND($B$5*B${totalRow},0)`)
    .setNumberFormat('0')
    .setBackground('#F5F5F5');
  
  // Update TOTAL row formulas
  const newTotalRow = totalRow + 1;
  allocSheet.getRange(newTotalRow, 2).setFormula(`=SUM(B9:B${totalRow})`);
  allocSheet.getRange(newTotalRow, 3).setFormula(`=SUM(C9:C${totalRow})`);
  
  // Find next available column for checkboxes
  let nextCol = 7;
  while (allocSheet.getRange(5, nextCol).getValue() && nextCol < 20) {
    nextCol++;
  }
  
  // Add header for new workstream
  allocSheet.getRange(5, nextCol).setValue(workstreamName)
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  allocSheet.setColumnWidth(nextCol, 80);
  
  // Add checkboxes for all priority rows
  for (let row = 6; row <= 20; row++) {
    allocSheet.getRange(row, nextCol).insertCheckboxes();
  }
  
  // Create the workstream tab with asset planning
  const wsSheet = ss.insertSheet(workstreamName);
  setupWorkstreamTabWithAssets(wsSheet, workstreamName);
  
  ui.alert('Success', `"${workstreamName}" added with asset planning.\nAllocate points in the Allocation tab.`, ui.ButtonSet.OK);
}

function removeWorkstream() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Get list of current workstreams
  const workstreams = [];
  let row = 9;
  let wsName = allocSheet.getRange(row, 1).getValue();
  while (wsName && wsName !== 'TOTAL' && row < 20) {
    workstreams.push(wsName);
    row++;
    wsName = allocSheet.getRange(row, 1).getValue();
  }
  
  if (workstreams.length === 0) {
    ui.alert('Error', 'No workstreams to remove.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt(
    'Remove Workstream',
    'Enter the name of the workstream to remove:\n\nAvailable: ' + workstreams.join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const workstreamName = response.getResponseText().trim();
  
  if (!workstreams.includes(workstreamName)) {
    ui.alert('Error', `"${workstreamName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Find and remove from allocation table
  row = 9;
  while (allocSheet.getRange(row, 1).getValue() !== 'TOTAL') {
    if (allocSheet.getRange(row, 1).getValue() === workstreamName) {
      allocSheet.deleteRow(row);
      break;
    }
    row++;
  }
  
  // Update TOTAL row formulas
  allocSheet.getRange(row, 2).setFormula(`=SUM(B9:B${row-1})`);
  allocSheet.getRange(row, 3).setFormula(`=SUM(C9:C${row-1})`);
  
  // Find and remove column from priorities
  const headerRow = 5;
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(headerRow, col).getValue() === workstreamName) {
      allocSheet.deleteColumn(col);
      break;
    }
  }
  
  // Delete the workstream sheet
  const wsSheet = ss.getSheetByName(workstreamName);
  if (wsSheet) {
    ss.deleteSheet(wsSheet);
  }
  
  ui.alert('Success', `"${workstreamName}" removed.`, ui.ButtonSet.OK);
}

// ==================== HELPER FUNCTIONS ====================

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
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getTeamNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const teams = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.endsWith(' Team')) {
      teams.push(name.replace(' Team', ''));
    }
  });
  
  return teams;
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
  
  if (teams.length === 0) {
    // If no teams, clear validations
    workstreams.forEach(wsName => {
      const wsSheet = ss.getSheetByName(wsName);
      if (wsSheet) {
        for (let row = 46; row <= 95; row++) {
          wsSheet.getRange(row, 6).clearDataValidations();
        }
      }
    });
    return;
  }
  
  // Update dropdowns in all workstreams
  const teamValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(teams, true)
    .setAllowInvalid(false)
    .build();
  
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (wsSheet) {
      for (let row = 46; row <= 95; row++) {
        wsSheet.getRange(row, 6).setDataValidation(teamValidation);
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
      .addItem('🔄 Refresh Team Assignments', 'refreshTeamAssignments'))
    .addToUi();
}