/**
 * Planning Tools - Standalone Add-on for Points System v12
 * Completely decoupled planning enhancement
 * Version: 3.0
 * 
 * INSTALLATION:
 * 1. Add this script to your Points System spreadsheet
 * 2. Run "installPlanningTools" function once
 * 3. Refresh the spreadsheet - both menus will appear
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  PLANNING_METHODS: ['Sprint', 'Waterfall'],
  SORT_OPTIONS: ['Sprint', 'Go-Live Date', 'Points'],
  COLORS: {
    HEADER: '#4285F4',
    CONFIG_BG: '#F5F5F5',
    SPRINT_1: '#E8F5E9',
    SPRINT_2: '#FFF9C4', 
    SPRINT_3: '#FFE0B2',
    SPRINT_4: '#F3E5F5',
    SPRINT_SEPARATOR: '#E0E0E0'
  }
};

// ==================== INSTALLATION ====================
/**
 * ONE-TIME SETUP FUNCTION
 * Run this once after adding the script to enable Planning Tools
 */
function installPlanningTools() {
  // Set up the trigger for planning menu
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove any existing planning triggers to avoid duplicates
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createPlanningMenu') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Add new trigger
  ScriptApp.newTrigger('createPlanningMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  // Create the menu immediately
  createPlanningMenu();
  
  SpreadsheetApp.getUi().alert(
    'Planning Tools Installed!',
    'Planning Tools menu has been added to your spreadsheet.\n\n' +
    'If you don\'t see both menus, please refresh the page.\n\n' +
    'The Planning Tools menu will now appear automatically when you open this spreadsheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Creates the Planning Tools menu
 * This is called by the trigger, not onOpen
 */
function createPlanningMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Planning Tools')
    .addItem('🎯 Apply Sprint Planning', 'applySprintPlanning')
    .addItem('🌊 Apply Waterfall Sorting', 'applyWaterfallSorting')
    .addSeparator()
    .addItem('⚙️ Planning Settings', 'openPlanningSettings')
    .addItem('🧹 Clear All Planning', 'clearAllPlanning')
    .addSeparator()
    .addItem('🔧 Reinstall Planning Tools', 'installPlanningTools')
    .addToUi();
}

// ==================== PLANNING CONFIG SETUP ====================
function getOrCreatePlanningConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(PLANNING_CONFIG.CONFIG_SHEET_NAME);
  
  if (!configSheet) {
    // Create config sheet
    configSheet = ss.insertSheet(PLANNING_CONFIG.CONFIG_SHEET_NAME);
    setupConfigSheet(configSheet);
  }
  
  return configSheet;
}

function setupConfigSheet(sheet) {
  sheet.clear();
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  
  // Title
  sheet.getRange('A1:B1').merge()
    .setValue('PLANNING CONFIGURATION')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.HEADER)
    .setFontColor('#FFFFFF');
  
  // Planning Method
  sheet.getRange('A3').setValue('Planning Method:');
  const methodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.PLANNING_METHODS, true)
    .build();
  sheet.getRange('B3').setDataValidation(methodValidation)
    .setValue('Sprint')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Sprint Duration
  sheet.getRange('A4').setValue('Sprint Duration:');
  const durationValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SPRINT_DURATIONS, true)
    .build();
  sheet.getRange('B4').setDataValidation(durationValidation)
    .setValue(PLANNING_CONFIG.DEFAULT_SPRINT_DURATION)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Start Date
  sheet.getRange('A5').setValue('Start Date:');
  const nextMonday = getNextMonday();
  sheet.getRange('B5').setValue(nextMonday)
    .setNumberFormat('yyyy-MM-dd')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Sort By
  sheet.getRange('A7').setValue('Sort Items By:');
  const sortValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SORT_OPTIONS, true)
    .build();
  sheet.getRange('B7').setDataValidation(sortValidation)
    .setValue('Sprint')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Add border
  sheet.getRange(3, 1, 5, 2).setBorder(true, true, true, true, true, true);
}

// ==================== APPLY SPRINT PLANNING ====================
function applySprintPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Verify Points System exists
  if (!verifyPointsSystem()) {
    SpreadsheetApp.getUi().alert(
      'Points System Not Found',
      'This tool requires Points System v12 to be set up first.\n\n' +
      'Please ensure you have an Allocation sheet and team sheets.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const configSheet = getOrCreatePlanningConfig();
  
  // Get configuration
  const sprintDuration = configSheet.getRange('B4').getValue();
  const startDate = new Date(configSheet.getRange('B5').getValue());
  
  // Process each team
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found. Please set up teams first.');
    return;
  }
  
  let successCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Add planning columns if they don't exist
    addPlanningColumns(teamSheet, true);
    
    // Get team capacity safely
    let netCapacity = 0;
    try {
      netCapacity = teamSheet.getRange('D7').getValue() || 0;
    } catch(e) {
      console.log(`Could not read capacity for ${teamName}`);
      return;
    }
    
    // Collect manifest items (rows 14-60)
    const manifestItems = [];
    for (let row = 14; row <= 60; row++) {
      try {
        const description = teamSheet.getRange(row, 2).getValue();
        const points = teamSheet.getRange(row, 4).getValue();
        
        if (description && points > 0 && 
            !description.toString().startsWith('---') &&
            description !== 'No assignments') {
          
          manifestItems.push({
            row: row,
            origin: teamSheet.getRange(row, 1).getValue(),
            description: description,
            size: teamSheet.getRange(row, 3).getValue() || '-',
            points: parseFloat(points) || 0,
            goLiveDate: teamSheet.getRange(row, 5).getValue(),
            source: teamSheet.getRange(row, 6).getValue() || 'Workstream'
          });
        }
      } catch(e) {
        // Skip rows that can't be read
        continue;
      }
    }
    
    // Collect team-initiated items (rows 62-91)
    for (let row = 62; row <= 91; row++) {
      try {
        const description = teamSheet.getRange(row, 2).getValue();
        const points = teamSheet.getRange(row, 4).getValue();
        
        if (description && points > 0) {
          manifestItems.push({
            row: row,
            origin: teamSheet.getRange(row, 1).getValue() || teamName,
            description: description,
            size: teamSheet.getRange(row, 3).getValue() || '-',
            points: parseFloat(points) || 0,
            goLiveDate: teamSheet.getRange(row, 5).getValue(),
            source: 'Team'
          });
        }
      } catch(e) {
        // Skip rows that can't be read
        continue;
      }
    }
    
    if (manifestItems.length > 0) {
      // Assign to sprints and reorganize
      assignSprintsAndReorganize(teamSheet, manifestItems, netCapacity, sprintDuration, startDate);
      successCount++;
    }
  });
  
  if (successCount === 0) {
    SpreadsheetApp.getUi().alert(
      'No Data Found',
      'No manifest items found to plan. Please run "Refresh Team Assignments" from the Points System menu first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'Sprint Planning Applied',
      `Successfully applied sprint planning to ${successCount} team(s).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== ASSIGN SPRINTS AND REORGANIZE ====================
function assignSprintsAndReorganize(teamSheet, items, capacity, sprintDuration, startDate) {
  // Calculate sprint capacity
  let sprintsInMonth, sprintCapacity;
  
  switch(sprintDuration) {
    case '1 week':
      sprintsInMonth = 4;
      break;
    case '1 month':
      sprintsInMonth = 1;
      break;
    default: // 2 weeks
      sprintsInMonth = 2;
  }
  
  sprintCapacity = capacity > 0 ? capacity / sprintsInMonth : 20; // Default if no capacity
  
  // Sort items by go-live date first
  items.sort((a, b) => {
    const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
    const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Assign to sprints
  let currentSprint = 1;
  let currentSprintLoad = 0;
  
  items.forEach(item => {
    if (currentSprintLoad + item.points > sprintCapacity * 1.2 && currentSprintLoad > 0) {
      currentSprint++;
      currentSprintLoad = 0;
    }
    
    item.sprint = currentSprint;
    item.sprintStart = calculateSprintDate(startDate, currentSprint - 1, sprintDuration, true);
    item.sprintEnd = calculateSprintDate(startDate, currentSprint - 1, sprintDuration, false);
    
    currentSprintLoad += item.points;
  });
  
  // Group items by sprint
  const sprintGroups = {};
  items.forEach(item => {
    if (!sprintGroups[item.sprint]) {
      sprintGroups[item.sprint] = [];
    }
    sprintGroups[item.sprint].push(item);
  });
  
  // Clear manifest areas first (but preserve formulas and structure)
  teamSheet.getRange(14, 1, 47, 9).clearContent();
  teamSheet.getRange(62, 1, 30, 9).clearContent();
  
  // Write back organized by sprint (workstream items area)
  let currentRow = 14;
  const sortedSprints = Object.keys(sprintGroups).sort((a, b) => parseInt(a) - parseInt(b));
  
  sortedSprints.forEach(sprint => {
    if (currentRow >= 61) return; // Don't overflow into team area
    
    // Add sprint header
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- SPRINT ${sprint} ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    // Add items in this sprint
    sprintGroups[sprint].forEach(item => {
      if (currentRow >= 61) return;
      
      // Write item data
      teamSheet.getRange(currentRow, 1).setValue(item.origin);
      teamSheet.getRange(currentRow, 2).setValue(item.description);
      teamSheet.getRange(currentRow, 3).setValue(item.size);
      teamSheet.getRange(currentRow, 4).setValue(item.points).setNumberFormat('0');
      if (item.goLiveDate) {
        teamSheet.getRange(currentRow, 5).setValue(item.goLiveDate);
      }
      teamSheet.getRange(currentRow, 6).setValue(item.source);
      teamSheet.getRange(currentRow, 7).setValue(`Sprint ${sprint}`);
      teamSheet.getRange(currentRow, 8).setValue(item.sprintStart).setNumberFormat('yyyy-MM-dd');
      teamSheet.getRange(currentRow, 9).setValue(item.sprintEnd).setNumberFormat('yyyy-MM-dd');
      
      // Apply sprint coloring
      const sprintColor = getSprintColor(parseInt(sprint));
      teamSheet.getRange(currentRow, 7, 1, 3).setBackground(sprintColor);
      
      currentRow++;
    });
    
    // Add spacing between sprints
    if (currentRow < 60) {
      currentRow++;
    }
  });
}

// ==================== APPLY WATERFALL SORTING ====================
function applyWaterfallSorting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Verify Points System exists
  if (!verifyPointsSystem()) {
    SpreadsheetApp.getUi().alert(
      'Points System Not Found',
      'This tool requires Points System v12 to be set up first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const configSheet = getOrCreatePlanningConfig();
  
  // Get configuration
  const startDate = new Date(configSheet.getRange('B5').getValue());
  
  // Process each team
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found.');
    return;
  }
  
  let successCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Add planning columns (no Sprint column for waterfall)
    addPlanningColumns(teamSheet, false);
    
    // Collect all manifest items
    const manifestItems = [];
    
    // Workstream items
    for (let row = 14; row <= 60; row++) {
      try {
        const description = teamSheet.getRange(row, 2).getValue();
        const points = teamSheet.getRange(row, 4).getValue();
        
        if (description && points > 0 && 
            !description.toString().startsWith('---') &&
            description !== 'No assignments') {
          
          manifestItems.push({
            row: row,
            data: [
              teamSheet.getRange(row, 1).getValue(),
              description,
              teamSheet.getRange(row, 3).getValue(),
              points,
              teamSheet.getRange(row, 5).getValue(),
              teamSheet.getRange(row, 6).getValue()
            ],
            goLiveDate: teamSheet.getRange(row, 5).getValue(),
            points: parseFloat(points) || 0
          });
        }
      } catch(e) {
        continue;
      }
    }
    
    // Team items
    for (let row = 62; row <= 91; row++) {
      try {
        const description = teamSheet.getRange(row, 2).getValue();
        const points = teamSheet.getRange(row, 4).getValue();
        
        if (description && points > 0) {
          manifestItems.push({
            row: row,
            data: [
              teamSheet.getRange(row, 1).getValue(),
              description,
              teamSheet.getRange(row, 3).getValue(),
              points,
              teamSheet.getRange(row, 5).getValue(),
              teamSheet.getRange(row, 6).getValue()
            ],
            goLiveDate: teamSheet.getRange(row, 5).getValue(),
            points: parseFloat(points) || 0,
            isTeamItem: true
          });
        }
      } catch(e) {
        continue;
      }
    }
    
    if (manifestItems.length > 0) {
      // Sort by go-live date and complexity
      manifestItems.sort((a, b) => {
        const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
        const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
        return dateA - dateB;
      });
      
      // Clear areas
      teamSheet.getRange(14, 1, 47, 8).clearContent();
      teamSheet.getRange(62, 1, 30, 8).clearContent();
      
      // Write back sorted items
      let wsRow = 14;
      let teamRow = 62;
      
      manifestItems.forEach(item => {
        const targetRow = item.isTeamItem ? teamRow++ : wsRow++;
        
        // Skip if we're out of space
        if ((!item.isTeamItem && targetRow >= 61) || (item.isTeamItem && targetRow >= 92)) {
          return;
        }
        
        // Write data
        item.data.forEach((value, col) => {
          if (value) {
            teamSheet.getRange(targetRow, col + 1).setValue(value);
          }
        });
        
        // Calculate simple start/end dates based on position
        const weekOffset = Math.floor((targetRow - (item.isTeamItem ? 62 : 14)) / 3);
        const itemStart = new Date(startDate);
        itemStart.setDate(itemStart.getDate() + (weekOffset * 7));
        
        const itemEnd = new Date(itemStart);
        itemEnd.setDate(itemEnd.getDate() + 6);
        
        teamSheet.getRange(targetRow, 7).setValue(itemStart).setNumberFormat('yyyy-MM-dd');
        teamSheet.getRange(targetRow, 8).setValue(itemEnd).setNumberFormat('yyyy-MM-dd');
        
        // Format
        teamSheet.getRange(targetRow, 4).setNumberFormat('0');
        teamSheet.getRange(targetRow, 7, 1, 2).setBackground('#E3F2FD');
      });
      
      successCount++;
    }
  });
  
  if (successCount === 0) {
    SpreadsheetApp.getUi().alert('No manifest items found to sort.');
  } else {
    SpreadsheetApp.getUi().alert(
      'Waterfall Sorting Applied',
      `Successfully sorted ${successCount} team(s) by timeline.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== CLEAR ALL PLANNING ====================
function clearAllPlanning() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Planning',
    'This will remove all planning columns and sprint groupings from team sheets.\n\n' +
    'You can regenerate the original manifests using the Points System menu:\n' +
    'Points System → Teams → Refresh Team Assignments\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    ui.alert('No team sheets found.');
    return;
  }
  
  let clearedCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    try {
      // Check if planning columns exist
      const headerG = teamSheet.getRange('G13').getValue();
      
      if (headerG) {
        // Clear planning columns (G, H, I)
        teamSheet.getRange(13, 7, 79, 3).clear();
        
        // Remove any sprint separators (rows starting with ---)
        for (let row = 14; row <= 60; row++) {
          const value = teamSheet.getRange(row, 1).getValue();
          if (value && value.toString().startsWith('---')) {
            teamSheet.getRange(row, 1, 1, 9).clear();
          }
        }
        
        clearedCount++;
      }
    } catch(e) {
      console.log(`Could not clear ${teamName}: ${e}`);
    }
  });
  
  ui.alert(
    'Planning Cleared',
    `Removed planning from ${clearedCount} team sheet(s).\n\n` +
    'Use Points System → Teams → Refresh Team Assignments to restore original manifests.',
    ui.ButtonSet.OK
  );
}

// ==================== HELPER FUNCTIONS ====================
function verifyPointsSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  return allocSheet !== null;
}

function addPlanningColumns(teamSheet, includeSprint) {
  try {
    // Check if columns already exist
    const existingHeader = teamSheet.getRange('G13').getValue();
    if (existingHeader) return;
    
    // Add headers
    if (includeSprint) {
      teamSheet.getRange('G13').setValue('Sprint').setFontWeight('bold').setBackground('#E1BEE7');
      teamSheet.getRange('H13').setValue('Start').setFontWeight('bold').setBackground('#E1BEE7');
      teamSheet.getRange('I13').setValue('End').setFontWeight('bold').setBackground('#E1BEE7');
    } else {
      teamSheet.getRange('G13').setValue('Start').setFontWeight('bold').setBackground('#E1BEE7');
      teamSheet.getRange('H13').setValue('End').setFontWeight('bold').setBackground('#E1BEE7');
    }
    
    // Apply border to extended table
    teamSheet.getRange(13, 1, 48, includeSprint ? 9 : 8).setBorder(true, true, true, true, true, true);
  } catch(e) {
    console.log('Could not add planning columns: ' + e);
  }
}

function getTeamNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = [];
  
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.endsWith(' Team')) {
      teams.push(name.replace(' Team', ''));
    }
  });
  
  return teams;
}

function getNextMonday() {
  const today = new Date();
  const day = today.getDay();
  const diff = day === 0 ? 1 : (8 - day) % 7 || 7;
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + diff);
  return nextMonday;
}

function calculateSprintDate(startDate, sprintIndex, duration, isStart) {
  const date = new Date(startDate);
  
  switch(duration) {
    case '1 week':
      date.setDate(date.getDate() + (sprintIndex * 7));
      if (!isStart) date.setDate(date.getDate() + 4);
      break;
    case '2 weeks':
      date.setDate(date.getDate() + (sprintIndex * 14));
      if (!isStart) date.setDate(date.getDate() + 13);
      break;
    case '1 month':
      date.setMonth(date.getMonth() + sprintIndex);
      if (!isStart) {
        date.setMonth(date.getMonth() + 1);
        date.setDate(0); // Last day of month
      }
      break;
  }
  
  return date;
}

function getSprintColor(sprintNumber) {
  const colors = [
    PLANNING_CONFIG.COLORS.SPRINT_1,
    PLANNING_CONFIG.COLORS.SPRINT_2,
    PLANNING_CONFIG.COLORS.SPRINT_3,
    PLANNING_CONFIG.COLORS.SPRINT_4
  ];
  return colors[(sprintNumber - 1) % colors.length];
}

function openPlanningSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = getOrCreatePlanningConfig();
  ss.setActiveSheet(configSheet);
}