/**
 * Marketing Team Points System - Version 4.2 Dynamic Workstreams Fixed
 * Workstream priorities use direct percentages, PMM priorities fill remaining space
 * Properly handles dynamic addition/removal of workstreams
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
  
  // Create default workstream tabs
  const defaultWorkstreams = ['SoMe', 'PUA', 'ASO', 'Portal'];
  defaultWorkstreams.forEach(name => {
    const wsSheet = ss.insertSheet(name);
    setupWorkstreamTab(wsSheet, name);
  });
  
  // Set the Allocation tab as active
  ss.setActiveSheet(allocationSheet);
  
  SpreadsheetApp.getUi().alert(
    'Points System Setup Complete! 🎉',
    'System ready. Workstream priorities use exact percentages.\n' +
    'PMM priorities automatically fill remaining capacity.',
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
  sheet.getRange('B5').setValue(200)
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
    sheet.getRange(row, 7).insertCheckboxes(); // SoMe
    sheet.getRange(row, 8).insertCheckboxes(); // PUA
    sheet.getRange(row, 9).insertCheckboxes(); // ASO
    sheet.getRange(row, 10).insertCheckboxes(); // Portal
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
  
  // Instructions
  sheet.getRange('A23:J23').merge()
    .setValue('📝 Check boxes to assign priorities to workstreams. Team priorities use exact %, PMM fills remainder.')
    .setFontSize(10)
    .setFontColor('#666666')
    .setBackground('#FFFEF7');
}

// ==================== HELPER FUNCTION TO FIND CHECKBOX COLUMN ====================

function findCheckboxColumn(workstreamName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Search through the header row (row 5) starting from column G (7)
  for (let col = 7; col <= 20; col++) {
    if (allocSheet.getRange(5, col).getValue() === workstreamName) {
      return col;
    }
  }
  return -1; // Not found
}

// ==================== WORKSTREAM TAB SETUP ====================

function setupWorkstreamTab(sheet, workstreamName) {
  sheet.clear();
  
  // Set column widths
  sheet.setColumnWidth(1, 350); // A - Priority Name
  sheet.setColumnWidth(2, 120); // B - Source
  sheet.setColumnWidth(3, 120); // C - Allocation %
  sheet.setColumnWidth(4, 120); // D - Points
  
  // Header Section
  sheet.getRange('A1:D1').merge()
    .setValue(`${workstreamName.toUpperCase()} WORKSTREAM`)
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');
  
  // Display allocated points
  sheet.getRange('A2').setValue('Total Points Allocated:');
  sheet.getRange('B2').setFormula(`=IFERROR(VLOOKUP("${workstreamName}",Allocation!A:C,3,FALSE),0)`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#E8F5E9')
    .setNumberFormat('0');
  sheet.getRange('C2').setValue('Points');
  
  // Table Header
  sheet.getRange('A4:D4').setValues([['Priority Name', 'Source', 'Allocation %', 'Points']])
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // Workstream Priorities Section (FIRST)
  sheet.getRange('A5:D5').merge()
    .setValue('--- Workstream Team Priorities (Direct %) ---')
    .setFontStyle('italic')
    .setBackground('#FFF9C4');
  
  // Add rows for workstream priorities
  for (let i = 0; i < 10; i++) {
    const row = 6 + i;
    sheet.getRange(row, 1).setBackground('#FFF3E0');
    sheet.getRange(row, 2).setValue('Workstream');
    sheet.getRange(row, 3).setNumberFormat('0%').setBackground('#FFF3E0');
    // Direct calculation for workstream priorities
    sheet.getRange(row, 4).setFormula(
      `=IF(C${row}="","",ROUND(C${row}*$B$2,0))`
    ).setNumberFormat('0').setBackground('#E8F5E9');
  }
  
  // Remaining capacity indicator
  sheet.getRange('A17').setValue('Remaining for PMM:').setFontWeight('bold');
  sheet.getRange('B17').setFormula('=MAX(0,100%-SUMIF(C6:C15,">0"))')
    .setNumberFormat('0%')
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  sheet.getRange('C17').setValue('→');
  sheet.getRange('D17').setFormula('=ROUND(B17*B2,0)')
    .setNumberFormat('0')
    .setFontWeight('bold')
    .setBackground('#E3F2FD');
  
  // PMM Priorities Section
  sheet.getRange('A19:D19').merge()
    .setValue('--- PMM Strategic Priorities (Auto-scaled) ---')
    .setFontStyle('italic')
    .setBackground('#F5F5F5');
  
  // Find the checkbox column for this workstream
  const checkboxCol = findCheckboxColumn(workstreamName);
  
  if (checkboxCol > 0) {
    // Create formulas for PMM priorities using column number
    for (let i = 0; i < 15; i++) {
      const currentRow = 20 + i;
      const allocRow = 6 + i;
      
      // Priority name - pulls from Allocation sheet if checkbox is checked
      sheet.getRange(currentRow, 1).setFormula(
        `=IF(INDIRECT("Allocation!R${allocRow}C${checkboxCol}",FALSE)=TRUE,Allocation!E${allocRow},"")`
      ).setBackground('#F0F0F0');
      
      // Source
      sheet.getRange(currentRow, 2).setFormula(
        `=IF(A${currentRow}<>"","PMM","")`
      ).setBackground('#F0F0F0');
      
      // PMM Weight % - scaled to fit remaining capacity
      sheet.getRange(currentRow, 3).setFormula(
        `=IF(A${currentRow}="","",` +
        `IF(SUMPRODUCT(INDIRECT("Allocation!R6C${checkboxCol}:R20C${checkboxCol}",FALSE)*Allocation!F6:F20)=0,0,` +
        `(Allocation!F${allocRow}*INDIRECT("Allocation!R${allocRow}C${checkboxCol}",FALSE))/` +
        `SUMPRODUCT(INDIRECT("Allocation!R6C${checkboxCol}:R20C${checkboxCol}",FALSE)*Allocation!F6:F20)*B17))`
      ).setNumberFormat('0%').setBackground('#F0F0F0');
      
      // Points
      sheet.getRange(currentRow, 4).setFormula(
        `=IF(C${currentRow}="","",ROUND(C${currentRow}*$B$2,0))`
      ).setNumberFormat('0').setBackground('#E8F5E9');
    }
  }
  
  // Summary Section
  sheet.getRange('A36').setValue('WORKSTREAM %:').setFontWeight('bold');
  sheet.getRange('B36').setFormula('=SUMIF(C6:C15,">0")')
    .setNumberFormat('0%')
    .setFontWeight('bold');
  
  sheet.getRange('C36').setValue('PMM %:').setFontWeight('bold');
  sheet.getRange('D36').setFormula('=SUMIF(C20:C34,">0")')
    .setNumberFormat('0%')
    .setFontWeight('bold');
  
  sheet.getRange('A37').setValue('TOTAL POINTS:').setFontWeight('bold');
  sheet.getRange('B37').setFormula('=SUM(D6:D15,D20:D34)')
    .setNumberFormat('0')
    .setFontWeight('bold');
  
  sheet.getRange('C37').setValue('STATUS:').setFontWeight('bold');
  sheet.getRange('D37').setFormula(
    '=IF(B37=B2,"✅ Allocated",' +
    'IF(B36>1,"⚠️ Over 100%",' +
    '"📊 " & TEXT(B37/B2,"0%")))'
  ).setFontColor('#006600');
  
  // Add borders
  sheet.getRange(4, 1, 12, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(17, 1, 1, 4).setBorder(true, true, true, true, false, false);
  sheet.getRange(19, 1, 16, 4).setBorder(true, true, true, true, true, true);
  sheet.getRange(36, 1, 2, 4).setBorder(true, true, true, true, false, false);
  
  // Instructions
  sheet.getRange('A39:D39').merge()
    .setValue('📝 Enter team priorities with exact %. PMM priorities auto-scale to fill remaining capacity.')
    .setFontSize(10)
    .setFontColor('#666666')
    .setBackground('#FFFEF7');
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
  let nextCol = 7; // Start from column G
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
  
  // Create the workstream tab
  const wsSheet = ss.insertSheet(workstreamName);
  setupWorkstreamTab(wsSheet, workstreamName);
  
  ui.alert('Success', `"${workstreamName}" added.\nAllocate points in the Allocation tab.`, ui.ButtonSet.OK);
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
  
  // Simple prompt for workstream name
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

// ==================== REFRESH FUNCTION ====================

function refreshWorkstreamFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  
  // Get all workstreams from allocation table
  const workstreams = [];
  let row = 9;
  let wsName = allocSheet.getRange(row, 1).getValue();
  while (wsName && wsName !== 'TOTAL' && row < 20) {
    workstreams.push(wsName);
    row++;
    wsName = allocSheet.getRange(row, 1).getValue();
  }
  
  // Refresh each workstream's formulas
  workstreams.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (wsSheet) {
      // Find the checkbox column for this workstream
      const checkboxCol = findCheckboxColumn(wsName);
      
      if (checkboxCol > 0) {
        // Update PMM priority formulas
        for (let i = 0; i < 15; i++) {
          const currentRow = 20 + i;
          const allocRow = 6 + i;
          
          // Update priority name formula
          wsSheet.getRange(currentRow, 1).setFormula(
            `=IF(INDIRECT("Allocation!R${allocRow}C${checkboxCol}",FALSE)=TRUE,Allocation!E${allocRow},"")`
          );
          
          // Update weight formula
          wsSheet.getRange(currentRow, 3).setFormula(
            `=IF(A${currentRow}="","",` +
            `IF(SUMPRODUCT(INDIRECT("Allocation!R6C${checkboxCol}:R20C${checkboxCol}",FALSE)*Allocation!F6:F20)=0,0,` +
            `(Allocation!F${allocRow}*INDIRECT("Allocation!R${allocRow}C${checkboxCol}",FALSE))/` +
            `SUMPRODUCT(INDIRECT("Allocation!R6C${checkboxCol}:R20C${checkboxCol}",FALSE)*Allocation!F6:F20)*B17))`
          );
        }
      }
    }
  });
  
  SpreadsheetApp.getUi().alert('Formulas refreshed successfully!');
}

// ==================== MENU SETUP ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Points System')
    .addItem('🚀 Initial Setup', 'setupPointsSystem')
    .addSeparator()
    .addItem('➕ Add Workstream', 'addWorkstream')
    .addItem('➖ Remove Workstream', 'removeWorkstream')
    .addItem('🔄 Refresh Formulas', 'refreshWorkstreamFormulas')
    .addToUi();
}