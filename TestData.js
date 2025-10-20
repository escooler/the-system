/**
 * Updated Test Data Script for Points System v13.2
 * Based on actual Fashion Store campaign data from screenshots
 */

// ==================== TEST DATA CONFIGURATION ====================
const TEST_DATA = {
  // Three teams as shown in the screenshots
  TEAMS: [
    { name: 'Ice Cream', members: 4, workingDays: 20, daysOff: 0, bufferPercent: 0.05, creativePlanning: 12 },
    { name: 'Cake', members: 4, workingDays: 20, daysOff: 0, bufferPercent: 0.05, creativePlanning: 12 },
    { name: 'Cash Kings', members: 3, workingDays: 20, daysOff: 0, bufferPercent: 0.05, creativePlanning: 10 }
  ],
  
  // Workstream allocations matching screenshots
  WORKSTREAM_ALLOCATIONS: {
    'SoMe': 0.70,      // 123 points
    'PUA': 0.16,       // 28 points
    'ASO': 0.08,       // 14 points
    'Portal': 0.03,    // 5 points
    'Superfans': 0.03  // 5 points
  },
  
  // Strategic priorities from PMM (from Allocation tab)
  STRATEGIC_PRIORITIES: [
    { name: 'Fashion Store Free', weight: 0.50, workstreams: ['SoMe', 'PUA', 'ASO', 'Portal', 'Superfans'] },
    { name: 'Fashion Furniture Pack', weight: 0.18, workstreams: ['SoMe', 'PUA', 'ASO', 'Portal', 'Superfans'] },
    { name: 'Fashion Style Pack', weight: 0.17, workstreams: ['SoMe', 'PUA', 'ASO', 'Portal', 'Superfans'] },
    { name: 'Seasonal Gifts Events', weight: 0.07, workstreams: ['SoMe', 'Portal', 'Superfans'] },
    { name: 'Weekly Gifts', weight: 0.03, workstreams: ['Portal', 'Superfans'] },
    { name: 'Seasonal Meta Campaign', weight: 0.05, workstreams: ['SoMe'] }
  ],
  
  // Workstream-specific priorities
  WORKSTREAM_PRIORITIES: {
    'SoMe': [
      { name: 'Evergreen Social Content', weight: 0.35 }
    ],
    'PUA': [
      { name: 'PUA Wild Card Experiments', weight: 0.35 }
    ],
    'ASO': [
      { name: 'Product Page Optimations', weight: 0.20 }
    ],
    'Portal': [],
    'Superfans': []
  },
  
  // Asset items for each workstream (based on screenshots)
  WORKSTREAM_ASSETS: {
    'SoMe': [
      // Fashion Store Free assets
      { desc: 'Free fashion store Teaser 2', size: 'M', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Free fashion store Teaser 2', size: 'L', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Free fashion store Showcase', size: 'L', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Free Fashion Show Followup', size: 'M', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Free fashion store Trailer', size: 'L', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Free fashion store Outnow', size: 'M', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      
      // Evergreen content
      { desc: 'Evergreen Social Post 1', size: 'L', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Evergreen Social Post 2', size: 'M', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Evergreen Social Post 3', size: 'M', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Evergreen Social Post 4', size: 'M', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Evergreen Social Post 5', size: 'L', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Evergreen Social Post 6', size: 'S', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      
      // Seasonal content
      { desc: 'Seasonal Themed Post', size: 'M', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Seasonal Themed Post', size: 'M', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      
      // Fashion Show Style Pack
      { desc: 'Fashion Show Style Pack Teaser', size: 'S', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      
      // Additional posts
      { desc: 'RoadMap Post Feb', size: 'S', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 }
    ],
    
    'PUA': [
      // Free Fashion Store campaigns
      { desc: 'Free Fashion Store - Fashion Story', size: 'L', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'Free Fashion Store - Shopping day', size: 'M', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'Free Fashion Store - Outfit Focus.', size: 'M', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'Free Fashion Store Stills', size: 'S', origin: 'Workstream', team: 'Cash Kings', goLiveDays: 25 },
      
      // Wild Card experiments
      { desc: 'Wild Card Creative - 1', size: 'S', origin: 'Workstream', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'Wild Card Creative - 1', size: 'M', origin: 'Workstream', team: 'Cash Kings', goLiveDays: 25 }
    ],
    
    'ASO': [
      // In-app events
      { desc: 'In app Event - Free Fashion Store', size: 'S', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'In app Event - Fashion Furniture Pack', size: 'S', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'In app Event - Fashion Style Pack', size: 'S', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      
      // Product page optimizations
      { desc: 'Product Page Optimation Trailer', size: 'S', origin: 'Workstream', team: 'Cash Kings', goLiveDays: 25 },
      
      // Custom icons
      { desc: 'Custom Icon Free Fashion Store', size: 'XS', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 },
      { desc: 'Custom Icon - Fashion Style Pack', size: 'XS', origin: 'PMM', team: 'Cash Kings', goLiveDays: 25 }
    ],
    
    'Portal': [
      // Portal posts
      { desc: 'Portal Posts - Fashion Store Free X2', size: 'XS', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Portal Posts - Fashion Store Furniture Pack X2', size: 'XS', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Fashion Store Style Pack X2', size: 'XS', origin: 'PMM', team: 'Cake', goLiveDays: 25 },
      { desc: 'Seasonal Gifts Posts X2', size: 'XS', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Weekly Gifts X2', size: 'XS', origin: 'Workstream', team: 'Ice Cream', goLiveDays: 25 }
    ],
    
    'Superfans': [
      // Superfans content
      { desc: 'Superfans Fashion Post', size: 'S', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 },
      { desc: 'Superfans Free Content Post', size: 'XS', origin: 'PMM', team: 'Ice Cream', goLiveDays: 25 }
    ]
  },
  
  // Team-initiated work (to fill any remaining capacity)
  TEAM_INITIATIVES: {
    'Ice Cream': [],  // Already at capacity
    'Cake': [],       // Already at capacity  
    'Cash Kings': []  // Already at capacity
  }
};

// ==================== MAIN FUNCTIONS ====================

/**
 * Populates the sheet with test data
 */
function populateWithTestData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Populate Test Data',
    'This will:\n' +
    '1. Set up 3 teams (Ice Cream, Cake, Cash Kings)\n' +
    '2. Configure workstream allocations\n' +
    '3. Add Fashion Store campaign priorities\n' +
    '4. Add workstream priorities and assets\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if the system is set up
    if (!ss.getSheetByName('Allocation')) {
      ui.alert('Error', 'Please run Initial Setup first (Points System → Initial Setup)', ui.ButtonSet.OK);
      return;
    }
    
    // Step 1: Set up teams
    setupTestTeams(ss);
    
    // Step 2: Add Superfans workstream if it doesn't exist
    addSuperfansWorkstream(ss);
    
    // Step 3: Configure allocations
    configureAllocations(ss);
    
    // Step 4: Add strategic priorities
    addStrategicPriorities(ss);
    
    // Step 5: Add workstream priorities
    addWorkstreamPriorities(ss);
    
    // Step 6: Add assets
    addWorkstreamAssets(ss);
    
    // Step 7: Add team initiatives (if any)
    addTeamInitiatives(ss);
    
    // Success message
    ui.alert(
      'Test Data Populated! 🎉',
      'Successfully added:\n' +
      `• ${TEST_DATA.TEAMS.length} teams\n` +
      `• ${Object.keys(TEST_DATA.WORKSTREAM_ALLOCATIONS).length} workstreams\n` +
      `• ${TEST_DATA.STRATEGIC_PRIORITIES.length} strategic priorities\n' +
      `• Fashion Store campaign assets\n\n` +
      'Use "Points System → Teams → Refresh Team Assignments" to generate manifests.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', 'Failed to populate test data: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Clears all data while maintaining structure
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Data',
    'This will clear all data but keep the sheet structure intact.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Clear allocation data
    const allocSheet = ss.getSheetByName('Allocation');
    if (allocSheet) {
      // Clear strategic priorities
      allocSheet.getRange('E7:J21').clearContent();
      // Reset workstream allocations to 0
      allocSheet.getRange('B10:B14').setValue(0);
    }
    
    // Clear workstream data
    ['SoMe', 'PUA', 'ASO', 'Portal', 'Superfans'].forEach(wsName => {
      const wsSheet = ss.getSheetByName(wsName);
      if (wsSheet) {
        // Clear workstream priorities
        wsSheet.getRange('A7:C16').clearContent();
        // Clear assets - ONLY clear content in columns A, B, C, E, F (skip D which has formulas)
        for (let row = 46; row <= 95; row++) {
          wsSheet.getRange(`A${row}`).clearContent();
          wsSheet.getRange(`B${row}`).clearContent();
          wsSheet.getRange(`C${row}`).clearContent();
          wsSheet.getRange(`E${row}`).clearContent();
          wsSheet.getRange(`F${row}`).clearContent();
        }
      }
    });
    
    // Clear team data
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.endsWith(' Team')) {
        // Clear team configurations
        sheet.getRange('B4').clearContent();  // Team members
        sheet.getRange('B6').clearContent();  // Buffer %
        sheet.getRange('B7').clearContent();  // Creative planning
        sheet.getRange('D4').setValue(20);    // Reset to default working days
        
        // Clear team assignments (manifest area)
        sheet.getRange('A16:F63').clearContent();
        
        // Clear team initiatives
        for (let row = 65; row <= 94; row++) {
          sheet.getRange(`B${row}`).clearContent();  // Description
          sheet.getRange(`C${row}`).clearContent();  // T-shirt size
        }
        
        // Clear team member details
        for (let row = 5; row <= 14; row++) {
          sheet.getRange(`H${row}`).clearContent();  // Email
          sheet.getRange(`I${row}`).clearContent();  // Jira Username
          sheet.getRange(`J${row}`).setValue(0);     // Holiday days
        }
      }
    });
    
    ui.alert('Data Cleared', 'All data has been cleared. Sheet structure remains intact.', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', 'Failed to clear data: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ==================== HELPER FUNCTIONS ====================

function setupTestTeams(ss) {
  // First, remove any existing teams except the default
  const existingTeams = ss.getSheets()
    .filter(sheet => sheet.getName().endsWith(' Team'))
    .map(sheet => sheet.getName().replace(' Team', ''));
  
  // Delete existing teams
  existingTeams.forEach(teamName => {
    const sheet = ss.getSheetByName(teamName + ' Team');
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  });
  
  // Create new test teams
  TEST_DATA.TEAMS.forEach(teamData => {
    const teamSheet = ss.insertSheet(teamData.name + ' Team');
    
    // Call the main script's setup function
    if (typeof setupTeamTab !== 'undefined') {
      setupTeamTab(teamSheet, teamData.name);
    }
    
    // Configure team settings
    teamSheet.getRange('B4').setValue(teamData.members);
    teamSheet.getRange('D4').setValue(teamData.workingDays);
    teamSheet.getRange('B6').setValue(teamData.bufferPercent);
    teamSheet.getRange('B7').setValue(teamData.creativePlanning);
  });
  
  // Update dropdowns and capacity
  if (typeof updateTeamDropdowns !== 'undefined') {
    updateTeamDropdowns();
  }
  if (typeof updateTotalCapacity !== 'undefined') {
    updateTotalCapacity();
  }
}

function addSuperfansWorkstream(ss) {
  // Check if Superfans workstream exists
  if (!ss.getSheetByName('Superfans')) {
    // Add Superfans to the allocation tab
    const allocSheet = ss.getSheetByName('Allocation');
    
    // Find the TOTAL row
    let totalRow = 10;
    while (allocSheet.getRange(totalRow, 1).getValue() !== 'TOTAL' && totalRow < 20) {
      totalRow++;
    }
    
    // Insert new row for Superfans
    allocSheet.insertRowBefore(totalRow);
    allocSheet.getRange(`A${totalRow}`).setValue('Superfans');
    allocSheet.getRange(`B${totalRow}`).setValue(0);
    allocSheet.getRange(`B${totalRow}`).setNumberFormat('0%').setBackground('#FFF9C4');
    allocSheet.getRange(`C${totalRow}`).setFormula(`=ROUND($B$6*B${totalRow},0)`);
    allocSheet.getRange(`C${totalRow}`).setNumberFormat('0').setBackground('#F5F5F5');
    
    // Update TOTAL row formulas
    const newTotalRow = totalRow + 1;
    allocSheet.getRange(`B${newTotalRow}`).setFormula(`=SUM(B10:B${totalRow})`);
    allocSheet.getRange(`C${newTotalRow}`).setFormula(`=SUM(C10:C${totalRow})`);
    
    // Add checkbox column for Superfans
    let nextCol = 7;
    while (allocSheet.getRange(6, nextCol).getValue() && nextCol < 20) {
      nextCol++;
    }
    allocSheet.getRange(6, nextCol).setValue('Superfans');
    allocSheet.getRange(6, nextCol).setFontWeight('bold').setBackground('#E3F2FD');
    
    for (let row = 7; row <= 21; row++) {
      allocSheet.getRange(row, nextCol).insertCheckboxes();
    }
    
    // Create Superfans workstream sheet
    const superfansSheet = ss.insertSheet('Superfans');
    if (typeof setupWorkstreamTab !== 'undefined') {
      setupWorkstreamTab(superfansSheet, 'Superfans');
    }
  }
}

function configureAllocations(ss) {
  const allocSheet = ss.getSheetByName('Allocation');
  if (!allocSheet) return;
  
  // Set current month to February 2025 (as in screenshots)
  allocSheet.getRange('C3').setValue('February');
  allocSheet.getRange('E3').setValue(2025);
  
  // Set workstream allocations
  const workstreams = Object.keys(TEST_DATA.WORKSTREAM_ALLOCATIONS);
  workstreams.forEach((wsName, index) => {
    // Find the row for this workstream
    let row = 10;
    while (row < 20) {
      if (allocSheet.getRange(row, 1).getValue() === wsName) {
        allocSheet.getRange(`B${row}`).setValue(TEST_DATA.WORKSTREAM_ALLOCATIONS[wsName]);
        break;
      }
      row++;
    }
  });
}

function addStrategicPriorities(ss) {
  const allocSheet = ss.getSheetByName('Allocation');
  if (!allocSheet) return;
  
  // Clear existing priorities
  allocSheet.getRange('E7:J21').clearContent();
  
  TEST_DATA.STRATEGIC_PRIORITIES.forEach((priority, index) => {
    const row = 7 + index;
    
    // Set name and weight
    allocSheet.getRange(`E${row}`).setValue(priority.name);
    allocSheet.getRange(`F${row}`).setValue(priority.weight);
    
    // Clear all checkboxes first
    for (let col = 7; col <= 11; col++) {  // Extended to include Superfans column
      allocSheet.getRange(row, col).setValue(false);
    }
    
    // Check appropriate workstream boxes
    const workstreamList = ['SoMe', 'PUA', 'ASO', 'Portal', 'Superfans'];
    priority.workstreams.forEach(wsName => {
      const colIndex = workstreamList.indexOf(wsName);
      if (colIndex >= 0) {
        allocSheet.getRange(row, 7 + colIndex).setValue(true);
      }
    });
  });
}

function addWorkstreamPriorities(ss) {
  Object.keys(TEST_DATA.WORKSTREAM_PRIORITIES).forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    // Clear existing priorities
    wsSheet.getRange('A7:C16').clearContent();
    
    // Add new priorities
    TEST_DATA.WORKSTREAM_PRIORITIES[wsName].forEach((priority, index) => {
      const row = 7 + index;
      wsSheet.getRange(`A${row}`).setValue(priority.name);
      wsSheet.getRange(`C${row}`).setValue(priority.weight);
    });
  });
}

function addWorkstreamAssets(ss) {
  const currentDate = new Date();
  
  Object.keys(TEST_DATA.WORKSTREAM_ASSETS).forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    // Clear existing assets
    for (let row = 46; row <= 95; row++) {
      wsSheet.getRange(`A${row}`).clearContent();
      wsSheet.getRange(`B${row}`).clearContent();
      wsSheet.getRange(`C${row}`).clearContent();
      wsSheet.getRange(`E${row}`).clearContent();
      wsSheet.getRange(`F${row}`).clearContent();
    }
    
    // Add new assets
    TEST_DATA.WORKSTREAM_ASSETS[wsName].forEach((asset, index) => {
      const row = 46 + index;
      
      // Calculate go-live date (October 15, 2025 for most items as shown)
      const goLiveDate = new Date(2025, 9, 15);  // October 15, 2025
      
      wsSheet.getRange(`A${row}`).setValue(asset.desc);
      wsSheet.getRange(`B${row}`).setValue(goLiveDate);
      wsSheet.getRange(`C${row}`).setValue(asset.size);
      wsSheet.getRange(`E${row}`).setValue(asset.origin);
      wsSheet.getRange(`F${row}`).setValue(asset.team);
    });
  });
}

function addTeamInitiatives(ss) {
  const currentDate = new Date();
  
  Object.keys(TEST_DATA.TEAM_INITIATIVES).forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Clear existing initiatives
    for (let row = 65; row <= 94; row++) {
      teamSheet.getRange(`B${row}`).clearContent();
      teamSheet.getRange(`C${row}`).clearContent();
    }
    
    // Add new initiatives (if any)
    TEST_DATA.TEAM_INITIATIVES[teamName].forEach((initiative, index) => {
      const row = 65 + index;
      
      const goLiveDate = new Date(currentDate);
      goLiveDate.setDate(goLiveDate.getDate() + initiative.goLiveDays);
      
      teamSheet.getRange(`B${row}`).setValue(initiative.desc);
      teamSheet.getRange(`C${row}`).setValue(initiative.size);
      teamSheet.getRange(`E${row}`).setValue(goLiveDate);
    });
    
    SpreadsheetApp.flush();
  });
}

// ==================== MENU INSTALLATION ====================

/**
 * Installs the Test Data menu using an installable trigger
 * Run this once from the script editor
 */
function installTestDataMenu() {
  // Remove any existing triggers first
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createTestDataMenu') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('createTestDataMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  // Also create the menu now for immediate use
  createTestDataMenu();
  
  SpreadsheetApp.getUi().alert(
    'Test Data Menu Installed!',
    'The Test Data menu has been installed.\n\n' +
    'It will appear automatically each time you open the sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Creates the Test Data menu
 * This is called by the trigger on sheet open
 */
function createTestDataMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🧪 Test Data')
    .addItem('📊 Populate with Fashion Campaign Data', 'populateWithTestData')
    .addItem('🧹 Clear All Data', 'clearAllData')
    .addToUi();
}