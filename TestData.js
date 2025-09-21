/**
 * Fixed Test Data Script for Points System v12.2
 * 
 * FIXED: Team initiatives now properly trigger point calculations
 * - Ensures t-shirt size values are set correctly
 * - Forces formula recalculation after data population
 * - Matches the fix used for workstream assets
 */

// ==================== TEST DATA CONFIGURATION ====================
const TEST_DATA = {
  // Three teams with realistic configurations
  TEAMS: [
    { name: 'Creative', members: 6, workingDays: 20, daysOff: 3, bufferPercent: 0.10, creativePlanning: 3 },
    { name: 'Performance', members: 3, workingDays: 20, daysOff: 2, bufferPercent: 0.05, creativePlanning: 4 },
    { name: 'Content', members: 4, workingDays: 20, daysOff: 4, bufferPercent: 0.12, creativePlanning: 2 }
  ],
  
  // Workstream allocations (must total 100%)
  WORKSTREAM_ALLOCATIONS: {
    'SoMe': 0.50,
    'PUA': 0.25,
    'ASO': 0.15,
    'Portal': 0.10
  },
  
  // Strategic priorities from PMM (your actual priorities)
  STRATEGIC_PRIORITIES: [
    { name: 'Free House', weight: 0.40, workstreams: ['SoMe', 'PUA', 'Portal'] },
    { name: 'Sweet Pea Cottage', weight: 0.30, workstreams: ['SoMe', 'PUA'] },
    { name: 'Brand Focused Project', weight: 0.20, workstreams: ['SoMe', 'Portal', 'ASO'] },
    { name: 'Black Friday Sales', weight: 0.10, workstreams: ['ASO', 'Portal'] }
  ],
  
  // Workstream-specific priorities (each workstream can have their own)
  WORKSTREAM_PRIORITIES: {
    'SoMe': [
      { name: 'Instagram Content Calendar', weight: 0.20 },
      { name: 'TikTok Growth Strategy', weight: 0.15 },
      { name: 'Community Engagement Plan', weight: 0.10 },
      { name: 'Influencer Partnerships', weight: 0.08 }
    ],
    'PUA': [
      { name: 'Google Ads Optimization', weight: 0.25 },
      { name: 'Facebook Campaign Refresh', weight: 0.15 },
      { name: 'Retargeting Strategy', weight: 0.10 }
    ],
    'ASO': [
      { name: 'Keyword Research Update', weight: 0.30 },
      { name: 'A/B Testing Screenshots', weight: 0.20 },
      { name: 'Reviews & Ratings Campaign', weight: 0.15 }
    ],
    'Portal': [
      { name: 'User Dashboard Redesign', weight: 0.15 },
      { name: 'Mobile Experience Update', weight: 0.12 },
      { name: 'Analytics Integration', weight: 0.10 }
    ]
  },
  
  // Asset items for each workstream (comprehensive to use full budgets)
  WORKSTREAM_ASSETS: {
    'SoMe': [
      // Free House campaign assets
      { desc: 'Free House Teaser', size: 'M', origin: 'PMM', team: 'Creative', goLiveDays: 30 },
      { desc: 'Free House Trailer', size: 'L', origin: 'PMM', team: 'Creative', goLiveDays: 25 },
      { desc: 'Free House Stories', size: 'S', origin: 'PMM', team: 'Creative', goLiveDays: 20 },
      
      // Sweet Pea Cottage assets
      { desc: 'Sweet Pea Cottage Reveal Video', size: 'L', origin: 'PMM', team: 'Creative', goLiveDays: 35 },
      { desc: 'Sweet Pea Before/After Reels', size: 'M', origin: 'PMM', team: 'Creative', goLiveDays: 32 },
      { desc: 'Sweet Post', size: 'S', origin: 'Workstream', team: 'Creative', goLiveDays: 28 },
      
      // Brand Focused Project
      { desc: 'OK Street Location Tour', size: 'L', origin: 'PMM', team: 'Creative', goLiveDays: 45 },
      { desc: 'Brand Guidelines Social Templates', size: 'M', origin: 'Workstream', team: 'Creative', goLiveDays: 15 },
      
      // Black Friday Sales
      { desc: 'Black Friday Countdown Content', size: 'S', origin: 'PMM', team: 'Creative', goLiveDays: 50 },
      { desc: 'Flash Sale Graphics', size: 'S', origin: 'Workstream', team: 'Creative', goLiveDays: 48 },
      
      // Workstream initiatives
      { desc: 'Roadmap Post', size: 'S', origin: 'Workstream', team: 'Content', goLiveDays: 5 },
      { desc: 'Super Fans: Decoration', size: 'S', origin: 'Workstream', team: 'Creative', goLiveDays: 22 },
      { desc: 'Evergreen 1', size: 'XS', origin: 'Workstream', team: 'Content', goLiveDays: 8 },
      { desc: 'Everygreen 2', size: 'M', origin: 'Workstream', team: 'Creative', goLiveDays: 18 },
      { desc: 'Evergreen 3', size: 'M', origin: 'Workstream', team: 'Content', goLiveDays: 12 }
    ],
    
    'PUA': [
      // Free House campaigns
      { desc: 'Free House Trailer Cut', size: 'M', origin: 'PMM', team: 'Content', goLiveDays: 25 },
      { desc: 'Friendship story Edit', size: 'M', origin: 'PMM', team: 'Performance', goLiveDays: 20 },
      
      // Sweet Pea Cottage campaigns
      { desc: 'Sweet Pea Stills', size: 'XS', origin: 'PMM', team: 'Content', goLiveDays: 30 },
      { desc: 'Sweet Pea Life', size: 'S', origin: 'PMM', team: 'Performance', goLiveDays: 28 },
      
      // Black Friday Sales
      { desc: 'WILD CARD', size: 'M', origin: 'PMM', team: 'Content', goLiveDays: 48 },
      { desc: 'WILD CARD 2', size: 'M', origin: 'Workstream', team: 'Content', goLiveDays: 45 },
      
      // Workstream optimization
      { desc: 'Q1 Ad Copy Variants (50x)', size: 'M', origin: 'Workstream', team: 'Content', goLiveDays: 10 },
      { desc: 'Display Network Banners', size: 'M', origin: 'Workstream', team: 'Content', goLiveDays: 15 },
      { desc: 'Performance Dashboard Setup', size: 'M', origin: 'Workstream', team: 'Performance', goLiveDays: 8 },
      { desc: 'A/B Testing Framework', size: 'S', origin: 'Workstream', team: 'Performance', goLiveDays: 12 },
      { desc: 'Conversion Tracking Update', size: 'S', origin: 'Workstream', team: 'Performance', goLiveDays: 5 }
    ],
    
    'ASO': [
      // Brand Focused Project
      { desc: 'Brand App Store Graphics', size: 'S', origin: 'PMM', team: 'Performance', goLiveDays: 35 },
      { desc: 'Brand Messaging in App Desc', size: 'S', origin: 'PMM', team: 'Performance', goLiveDays: 30 },
      
      // Workstream optimization
      { desc: 'iOS Screenshots Optimization', size: 'S', origin: 'Workstream', team: 'Performance', goLiveDays: 20 },
      { desc: 'Android Store Listing Update', size: 'S', origin: 'Workstream', team: 'Performance', goLiveDays: 22 },
      { desc: 'App Preview Videos (2x)', size: 'M', origin: 'Workstream', team: 'Performance', goLiveDays: 25 },
      { desc: 'Localization Updates', size: 'S', origin: 'Workstream', team: 'Performance', goLiveDays: 15 }
    ],
    
    'Portal': [
      // Free House integration
      { desc: 'Free House Portal Section', size: 'XS', origin: 'PMM', team: 'Creative', goLiveDays: 40 },
      { desc: 'Free House Virtual Tour', size: 'XS', origin: 'PMM', team: 'Creative', goLiveDays: 35 },
      
      // Brand Focused Project
      { desc: 'Know your crumpet 1 Artical', size: 'XS', origin: 'PMM', team: 'Creative', goLiveDays: 30 },
      { desc: 'Know your crumpet 2 Artical', size: 'XS', origin: 'PMM', team: 'Content', goLiveDays: 28 },
      
      // Workstream improvements
      { desc: 'Portal Sweet Pea Gifts 1', size: 'XS', origin: 'Workstream', team: 'Creative', goLiveDays: 18 },
      { desc: 'Portal Sweet Pea Gifts 2', size: 'XS', origin: 'Workstream', team: 'Creative', goLiveDays: 25 },
      { desc: 'Portal Sweet shop link', size: 'XS', origin: 'Workstream', team: 'Content', goLiveDays: 12 },
      { desc: 'Portal Sweet Pea trailer cut', size: 'XS', origin: 'Workstream', team: 'Content', goLiveDays: 15 },
      { desc: 'Road map', size: 'XS', origin: 'Workstream', team: 'Content', goLiveDays: 8 },
      { desc: 'Portal Sweet Pea Gifts 3', size: 'XS', origin: 'Workstream', team: 'Content', goLiveDays: 10 }
    ]
  },
  
  // Team-initiated work (to fill remaining capacity)
  TEAM_INITIATIVES: {
    'Creative': [
      { desc: 'Design System Documentation', size: 'S', goLiveDays: 20 },
      { desc: 'Brand Asset Library Update', size: 'M', goLiveDays: 25 },
      { desc: 'Creative Process Optimization', size: 'XS', goLiveDays: 15 }
    ],
    'Performance': [
      { desc: 'Q1 Analytics Audit', size: 'S', goLiveDays: 25 },
      
    ],
    'Content': [
      { desc: 'Editorial Calendar System', size: 'S', goLiveDays: 22 },
      { desc: 'Content Style Guide Update', size: 'M', goLiveDays: 14 },
      
    ]
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
    '1. Set up 3 teams (Creative, Performance, Content)\n' +
    '2. Configure workstream allocations\n' +
    '3. Add strategic priorities\n' +
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
    
    // Step 2: Configure allocations
    configureAllocations(ss);
    
    // Step 3: Add strategic priorities
    addStrategicPriorities(ss);
    
    // Step 4: Add workstream priorities
    addWorkstreamPriorities(ss);
    
    // Step 5: Add assets
    addWorkstreamAssets(ss);
    
    // Step 6: Add team initiatives  
    addTeamInitiatives(ss);
    
    // Success message
    ui.alert(
      'Test Data Populated! 🎉',
      'Successfully added:\n' +
      `• ${TEST_DATA.TEAMS.length} teams\n` +
      `• ${TEST_DATA.STRATEGIC_PRIORITIES.length} strategic priorities\n` +
      `• Workstream priorities and assets\n\n` +
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
      allocSheet.getRange('B10:B13').setValue(0);
    }
    
    // Clear workstream data
    ['SoMe', 'PUA', 'ASO', 'Portal'].forEach(wsName => {
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
    
    // Clear team data - be careful not to clear formulas
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.endsWith(' Team')) {
        // Clear team configurations (but not formulas)
        sheet.getRange('B4').clearContent();  // Team members
        sheet.getRange('B5').clearContent();  // Days off
        sheet.getRange('B6').clearContent();  // Buffer %
        sheet.getRange('B7').clearContent();  // Creative planning
        sheet.getRange('D4').setValue(20);    // Reset to default working days
        
        // Clear team assignments (manifest area)
        sheet.getRange('A14:F60').clearContent();
        
        // Clear team initiatives - only B and C columns (A has origin, D has formula, E has date, F has source)
        for (let row = 62; row <= 91; row++) {
          sheet.getRange(`B${row}`).clearContent();  // Description
          sheet.getRange(`C${row}`).clearContent();  // T-shirt size
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
    teamSheet.getRange('B5').setValue(teamData.daysOff);
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

function configureAllocations(ss) {
  const allocSheet = ss.getSheetByName('Allocation');
  if (!allocSheet) return;
  
  // Set current month and year
  const currentDate = new Date();
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];
  
  allocSheet.getRange('C3').setValue(months[currentDate.getMonth()]);
  allocSheet.getRange('E3').setValue(currentDate.getFullYear());
  
  // Set workstream allocations
  const workstreams = Object.keys(TEST_DATA.WORKSTREAM_ALLOCATIONS);
  workstreams.forEach((wsName, index) => {
    const row = 10 + index;
    allocSheet.getRange(`B${row}`).setValue(TEST_DATA.WORKSTREAM_ALLOCATIONS[wsName]);
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
    for (let col = 7; col <= 10; col++) {
      allocSheet.getRange(row, col).setValue(false);
    }
    
    // Check appropriate workstream boxes
    const workstreamList = ['SoMe', 'PUA', 'ASO', 'Portal'];
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
    
    // Clear existing assets - BUT DO NOT clear column D (formulas) or validations
    for (let row = 46; row <= 95; row++) {
      wsSheet.getRange(`A${row}`).clearContent();  // Description
      wsSheet.getRange(`B${row}`).clearContent();  // Date
      wsSheet.getRange(`C${row}`).clearContent();  // Size (but keep validation)
      wsSheet.getRange(`E${row}`).clearContent();  // Origin (but keep validation)  
      wsSheet.getRange(`F${row}`).clearContent();  // Team (but keep validation)
    }
    
    // Add new assets
    TEST_DATA.WORKSTREAM_ASSETS[wsName].forEach((asset, index) => {
      const row = 46 + index;
      
      // Calculate go-live date
      const goLiveDate = new Date(currentDate);
      goLiveDate.setDate(goLiveDate.getDate() + asset.goLiveDays);
      
      wsSheet.getRange(`A${row}`).setValue(asset.desc);
      wsSheet.getRange(`B${row}`).setValue(goLiveDate);
      wsSheet.getRange(`C${row}`).setValue(asset.size);  // This should trigger the formula in D
      wsSheet.getRange(`E${row}`).setValue(asset.origin);
      wsSheet.getRange(`F${row}`).setValue(asset.team);
    });
  });
}

/**
 * FIXED: Team initiatives now properly trigger point calculations
 */
function addTeamInitiatives(ss) {
  const currentDate = new Date();
  
  Object.keys(TEST_DATA.TEAM_INITIATIVES).forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Clear existing initiatives - only clear content, not formulas or validations
    for (let row = 62; row <= 91; row++) {
      teamSheet.getRange(`B${row}`).clearContent();  // Description
      teamSheet.getRange(`C${row}`).clearContent();  // T-shirt size
      // Don't clear A, D, E, F as they contain formulas/defaults
    }
    
    // Add new initiatives
    TEST_DATA.TEAM_INITIATIVES[teamName].forEach((initiative, index) => {
      const row = 62 + index;
      
      // Calculate go-live date
      const goLiveDate = new Date(currentDate);
      goLiveDate.setDate(goLiveDate.getDate() + initiative.goLiveDays);
      
      // Set the values in the correct order to ensure formula calculation
      teamSheet.getRange(`B${row}`).setValue(initiative.desc);
      teamSheet.getRange(`C${row}`).setValue(initiative.size);
      teamSheet.getRange(`E${row}`).setValue(goLiveDate);
      
      // Force recalculation by setting the formula again
      // This ensures the t-shirt size gets converted to points
      const currentFormula = teamSheet.getRange(`D${row}`).getFormula();
      if (currentFormula) {
        teamSheet.getRange(`D${row}`).setFormula(currentFormula);
      }
    });
    
    // Force a recalculation of the entire sheet
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
    .addItem('📊 Populate with Test Data', 'populateWithTestData')
    .addItem('🧹 Clear All Data', 'clearAllData')
    .addToUi();
}