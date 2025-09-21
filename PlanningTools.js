/**
 * Enhanced Planning Tools - Fixed Sprint Distribution & Team Work Preservation
 * Version: 4.0 - Smart Sprint Distribution & Bug Fixes
 * 
 * FIXES:
 * 1. Even distribution of points across sprints based on capacity
 * 2. Smart sorting by go-live dates within sprint constraints
 * 3. Configurable number of sprints
 * 4. Team names in sprint headers
 * 5. Preserves team-initiated work section when clearing
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  PLANNING_METHODS: ['Sprint', 'Waterfall'],
  SORT_OPTIONS: ['Sprint', 'Go-Live Date', 'Points'],
  DEFAULT_FIRST_SPRINT: 1,
  COLORS: {
    HEADER: '#4285F4',
    CONFIG_BG: '#F5F5F5',
    SPRINT_1: '#E8F5E9',
    SPRINT_2: '#FFF9C4', 
    SPRINT_3: '#FFE0B2',
    SPRINT_4: '#F3E5F5',
    SPRINT_5: '#E1F5FE',
    SPRINT_6: '#FCE4EC',
    SPRINT_SEPARATOR: '#E0E0E0',
    ASSIGNMENT_BG: '#E3F2FD',
    TIMELINE_BG: '#F8F9FA'
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
    .addItem('🌊 Apply Waterfall Planning', 'applyWaterfallSorting')
    .addSeparator()
    .addItem('⚙️ Planning Settings', 'openPlanningSettings')
    .addItem('🧹 Clear All Planning', 'clearAllPlanning')
    .addSeparator()
    .addItem('🔧 Reinstall Planning Tools', 'installPlanningTools')
    .addToUi();
}

// ==================== ENHANCED PLANNING CONFIG SETUP ====================
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
  
  // First Sprint Number (NEW)
  sheet.getRange('A5').setValue('First Sprint Number:');
  sheet.getRange('B5').setValue(PLANNING_CONFIG.DEFAULT_FIRST_SPRINT)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Start Date
  sheet.getRange('A6').setValue('Start Date:');
  const nextMonday = getNextMonday();
  sheet.getRange('B6').setValue(nextMonday)
    .setNumberFormat('yyyy-MM-dd')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Sort By
  sheet.getRange('A8').setValue('Sort Items By:');
  const sortValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SORT_OPTIONS, true)
    .build();
  sheet.getRange('B8').setDataValidation(sortValidation)
    .setValue('Sprint')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Add border
  sheet.getRange(3, 1, 6, 2).setBorder(true, true, true, true, true, true);
}

// ==================== FIXED CLEARING FUNCTION - PRESERVES TEAM WORK ====================
function clearPlanningAreas(teamSheet) {
  try {
    // ONLY clear the workstream assignment area (14-60), NOT the team work area (62-91)
    teamSheet.getRange(14, 1, 47, 9).clear();
    
    // Clear any merged cells in workstream area only
    for (let row = 14; row <= 60; row++) {
      try {
        teamSheet.getRange(row, 1, 1, 9).breakApart();
      } catch(e) {
        // Cell wasn't merged, continue
      }
    }
    
    // Reset background colors to white in workstream area only
    teamSheet.getRange(14, 1, 47, 9).setBackground('#FFFFFF');
    
    // Clear any borders in workstream area only
    teamSheet.getRange(14, 1, 47, 9).setBorder(false, false, false, false, false, false);
    
    // Reset font styles in workstream area only
    teamSheet.getRange(14, 1, 47, 9).setFontWeight('normal').setFontStyle('normal');
    
    // Clear planning headers (G, H, I in row 13) but preserve original headers (A-F)
    teamSheet.getRange(13, 7, 1, 3).clear();
    
  } catch(e) {
    console.log('Error clearing planning areas: ' + e);
  }
}

// ==================== ENHANCED SPRINT PLANNING ====================
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
  const firstSprintNumber = parseInt(configSheet.getRange('B5').getValue()) || PLANNING_CONFIG.DEFAULT_FIRST_SPRINT;
  const startDate = new Date(configSheet.getRange('B6').getValue());
  
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
    
    // Get team capacity safely
    let netCapacity = 0;
    try {
      netCapacity = teamSheet.getRange('D7').getValue() || 0;
    } catch(e) {
      console.log(`Could not read capacity for ${teamName}`);
      return;
    }
    
    // Collect manifest items
    const manifestItems = collectManifestItems(teamSheet, teamName);
    
    if (manifestItems.length > 0) {
      // Clear areas completely first (but preserve team work)
      clearPlanningAreas(teamSheet);
      
      // Add sprint planning headers
      addSprintHeaders(teamSheet);
      
      // Smart sprint assignment and reorganization
      smartSprintAssignment(teamSheet, manifestItems, netCapacity, sprintDuration, firstSprintNumber, startDate, teamName);
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
      `Successfully applied smart sprint planning to ${successCount} team(s).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== COLLECT MANIFEST ITEMS ====================
function collectManifestItems(teamSheet, teamName) {
  const manifestItems = [];
  
  // ONLY collect workstream items (rows 14-60) for sprint planning
  // Team-initiated work (rows 62-91) stays in its own section
  for (let row = 14; row <= 60; row++) {
    try {
      const description = teamSheet.getRange(row, 2).getValue();
      const points = teamSheet.getRange(row, 4).getValue();
      
      if (description && points > 0 && 
          !description.toString().startsWith('---') &&
          description !== 'No assignments' &&
          description.trim() !== '') {
        
        manifestItems.push({
          origin: teamSheet.getRange(row, 1).getValue() || '',
          description: description,
          size: teamSheet.getRange(row, 3).getValue() || '-',
          points: parseFloat(points) || 0,
          goLiveDate: teamSheet.getRange(row, 5).getValue(),
          source: teamSheet.getRange(row, 6).getValue() || 'Workstream'
        });
      }
    } catch(e) {
      continue;
    }
  }
  
  // NOTE: Team items (rows 62-91) are NOT included in sprint planning
  // They remain in their dedicated team section and are counted separately
  
  return manifestItems;
}

// ==================== ADD SPRINT HEADERS ====================
function addSprintHeaders(teamSheet) {
  try {
    teamSheet.getRange('G13').setValue('Sprint')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
    teamSheet.getRange('H13').setValue('Start Date')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
    teamSheet.getRange('I13').setValue('End Date')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
      
    // Apply border to headers
    teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
  } catch(e) {
    console.log('Error adding sprint headers: ' + e);
  }
}

// ==================== SMART SPRINT ASSIGNMENT ====================
function smartSprintAssignment(teamSheet, items, capacity, sprintDuration, firstSprintNumber, startDate, teamName) {
  // Calculate how many sprints we need based on capacity and total points
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const targetSprintCapacity = Math.max(10, capacity / 2); // Target roughly half capacity per sprint
  const estimatedSprintsNeeded = Math.max(2, Math.ceil(totalPoints / targetSprintCapacity));
  
  // Sort items by go-live date first
  items.sort((a, b) => {
    const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
    const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Initialize sprint buckets starting from firstSprintNumber
  const sprints = [];
  for (let i = 0; i < estimatedSprintsNeeded; i++) {
    const sprintNumber = firstSprintNumber + i;
    sprints.push({
      number: sprintNumber,
      items: [],
      totalPoints: 0,
      startDate: calculateSprintDate(startDate, i, sprintDuration, true),
      endDate: calculateSprintDate(startDate, i, sprintDuration, false)
    });
  }
  
  // Smart assignment algorithm
  items.forEach(item => {
    let bestSprintIndex = 0;
    let bestScore = -1;
    
    // Find the best sprint for this item
    for (let i = 0; i < sprints.length; i++) {
      const sprint = sprints[i];
      
      // Skip if sprint is severely over capacity (more than 150% of target)
      if (sprint.totalPoints >= targetSprintCapacity * 1.5) continue;
      
      // Calculate score based on:
      // 1. How close go-live date is to sprint end
      // 2. How much capacity is left in sprint
      // 3. Preference for earlier sprints if dates are flexible
      
      let score = 0;
      
      // Date preference (higher score for better date fit)
      if (item.goLiveDate && item.goLiveDate instanceof Date) {
        const daysDiff = Math.abs((item.goLiveDate - sprint.endDate) / (1000 * 60 * 60 * 24));
        score += Math.max(0, 100 - daysDiff); // Up to 100 points for perfect date fit
      } else {
        score += 50; // Neutral score for items without dates
      }
      
      // Capacity preference (higher score for sprints with more capacity)
      const capacityUtilization = sprint.totalPoints / targetSprintCapacity;
      score += Math.max(0, (1 - capacityUtilization) * 50); // Up to 50 points for available capacity
      
      // Earlier sprint preference (small bias toward earlier sprints)
      score += (sprints.length - i) * 5; // Up to 5 * sprintCount points for being earlier
      
      if (score > bestScore) {
        bestScore = score;
        bestSprintIndex = i;
      }
    }
    
    // Assign item to best sprint
    const targetSprint = sprints[bestSprintIndex];
    item.sprint = targetSprint.number;
    item.sprintStart = targetSprint.startDate;
    item.sprintEnd = targetSprint.endDate;
    
    targetSprint.items.push(item);
    targetSprint.totalPoints += item.points;
  });
  
  // Write organized data back to sheet
  writeSprintData(teamSheet, sprints, teamName);
}

// ==================== WRITE SPRINT DATA ====================
function writeSprintData(teamSheet, sprints, teamName) {
  let currentRow = 14;
  
  sprints.forEach(sprint => {
    if (sprint.items.length === 0) return; // Skip empty sprints
    if (currentRow >= 61) return; // Don't overflow into team area
    
    // Add sprint header with team name
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- ${teamName.toUpperCase()} SPRINT ${sprint.number} (${sprint.totalPoints} pts) ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    // Sort items within sprint by go-live date
    sprint.items.sort((a, b) => {
      const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
      const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
      return dateA - dateB;
    });
    
    // Add items in this sprint
    sprint.items.forEach(item => {
      if (currentRow >= 61) return;
      
      // Write item data
      teamSheet.getRange(currentRow, 1).setValue(item.origin);
      teamSheet.getRange(currentRow, 2).setValue(item.description);
      teamSheet.getRange(currentRow, 3).setValue(item.size);
      teamSheet.getRange(currentRow, 4).setValue(item.points).setNumberFormat('0');
      if (item.goLiveDate) {
        teamSheet.getRange(currentRow, 5).setValue(item.goLiveDate);
        if (item.goLiveDate instanceof Date) {
          teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
        }
      }
      teamSheet.getRange(currentRow, 6).setValue(item.source);
      teamSheet.getRange(currentRow, 7).setValue(`Sprint ${sprint.number}`);
      teamSheet.getRange(currentRow, 8).setValue(item.sprintStart).setNumberFormat('yyyy-MM-dd');
      teamSheet.getRange(currentRow, 9).setValue(item.sprintEnd).setNumberFormat('yyyy-MM-dd');
      
      // Apply sprint coloring
      const sprintColor = getSprintColor(sprint.number);
      teamSheet.getRange(currentRow, 7, 1, 3).setBackground(sprintColor);
      
      currentRow++;
    });
    
    // Add spacing between sprints
    if (currentRow < 60) {
      currentRow++;
    }
  });
  
  // Apply border to the used area
  if (currentRow > 14) {
    try {
      teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
    } catch(e) {
      console.log('Error applying borders: ' + e);
    }
  }
}

// ==================== WATERFALL PLANNING (UNCHANGED) ====================
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
  const startDate = new Date(configSheet.getRange('B6').getValue());
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found.');
    return;
  }
  
  let successCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Get team configuration
    let teamMembers = 1;
    let netCapacity = 20;
    
    try {
      teamMembers = teamSheet.getRange('B4').getValue() || 1;
      netCapacity = teamSheet.getRange('D7').getValue() || 20;
    } catch(e) {
      console.log(`Could not read team config for ${teamName}`);
    }
    
    // Collect all manifest items
    const manifestItems = collectManifestItems(teamSheet, teamName);
    
    if (manifestItems.length > 0) {
      // Apply sequential waterfall planning
      applySequentialWaterfall(teamSheet, manifestItems, teamMembers, startDate);
      successCount++;
    }
  });
  
  if (successCount === 0) {
    SpreadsheetApp.getUi().alert('No manifest items found to distribute.');
  } else {
    SpreadsheetApp.getUi().alert(
      'Waterfall Planning Applied',
      `Successfully applied sequential waterfall planning to ${successCount} team(s).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== APPLY SEQUENTIAL WATERFALL ====================
function applySequentialWaterfall(teamSheet, items, teamMembers, startDate) {
  // Sort items by go-live date
  items.sort((a, b) => {
    const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
    const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Calculate team member schedules
  const teamMemberSchedules = [];
  for (let i = 1; i <= teamMembers; i++) {
    teamMemberSchedules.push({
      id: i,
      name: `Team Member ${i}`,
      nextAvailableDate: new Date(startDate),
      totalPoints: 0
    });
  }
  
  // Assign items to team members sequentially
  const scheduledItems = [];
  
  items.forEach(item => {
    // Find team member who will be available earliest
    const availableMember = teamMemberSchedules.reduce((earliest, current) => 
      current.nextAvailableDate < earliest.nextAvailableDate ? current : earliest
    );
    
    // Calculate task duration (1 point = 1 day, minimum 1 day)
    const taskDays = Math.max(1, Math.round(item.points));
    
    // Set start date (when this team member is available)
    const taskStart = new Date(availableMember.nextAvailableDate);
    
    // Calculate end date (add working days)
    const taskEnd = new Date(taskStart);
    taskEnd.setDate(taskEnd.getDate() + taskDays - 1);
    
    // Update team member's next available date
    availableMember.nextAvailableDate = new Date(taskEnd);
    availableMember.nextAvailableDate.setDate(availableMember.nextAvailableDate.getDate() + 1);
    availableMember.totalPoints += item.points;
    
    // Add to scheduled items
    scheduledItems.push({
      ...item,
      assignedTo: availableMember.name,
      startDate: taskStart,
      endDate: taskEnd
    });
  });
  
  // Sort scheduled items by start date for display
  scheduledItems.sort((a, b) => a.startDate - b.startDate);
  
  // Clear and write to sheet
  clearPlanningAreas(teamSheet);
  addWaterfallHeaders(teamSheet);
  writeScheduledItems(teamSheet, scheduledItems);
}

// ==================== ADD WATERFALL HEADERS ====================
function addWaterfallHeaders(teamSheet) {
  try {
    teamSheet.getRange('G13').setValue('Assigned To')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
    teamSheet.getRange('H13').setValue('Start Date')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
    teamSheet.getRange('I13').setValue('End Date')
      .setFontWeight('bold')
      .setBackground('#E1BEE7');
      
    // Apply border to headers
    teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
  } catch(e) {
    console.log('Error adding headers: ' + e);
  }
}

// ==================== WRITE SCHEDULED ITEMS ====================
function writeScheduledItems(teamSheet, scheduledItems) {
  let currentRow = 14;
  
  scheduledItems.forEach(item => {
    if (currentRow >= 61) return; // Don't overflow
    
    try {
      // Write item data in a clean row
      teamSheet.getRange(currentRow, 1).setValue(item.origin);
      teamSheet.getRange(currentRow, 2).setValue(item.description);
      teamSheet.getRange(currentRow, 3).setValue(item.size);
      teamSheet.getRange(currentRow, 4).setValue(item.points).setNumberFormat('0');
      
      if (item.goLiveDate) {
        teamSheet.getRange(currentRow, 5).setValue(item.goLiveDate);
        if (item.goLiveDate instanceof Date) {
          teamSheet.getRange(currentRow, 5).setNumberFormat('yyyy-MM-dd');
        }
      }
      
      teamSheet.getRange(currentRow, 6).setValue(item.source);
      teamSheet.getRange(currentRow, 7).setValue(item.assignedTo);
      teamSheet.getRange(currentRow, 8).setValue(item.startDate).setNumberFormat('yyyy-MM-dd');
      teamSheet.getRange(currentRow, 9).setValue(item.endDate).setNumberFormat('yyyy-MM-dd');
      
      // Apply simple, clean formatting
      teamSheet.getRange(currentRow, 7).setBackground(PLANNING_CONFIG.COLORS.ASSIGNMENT_BG); // Assigned To column
      teamSheet.getRange(currentRow, 8, 1, 2).setBackground(PLANNING_CONFIG.COLORS.TIMELINE_BG); // Date columns
      
      currentRow++;
    } catch(e) {
      console.log(`Error writing item at row ${currentRow}: ${e}`);
      currentRow++;
    }
  });
  
  // Apply border to the used area only
  if (currentRow > 14) {
    try {
      teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
    } catch(e) {
      console.log('Error applying borders: ' + e);
    }
  }
}

// ==================== IMPROVED CLEAR ALL PLANNING ====================
function clearAllPlanning() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Planning',
    'This will remove planning data from workstream assignments but preserve team-initiated work.\n\n' +
    'You can regenerate the original manifests using:\n' +
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
      // Clear planning areas (preserving team-initiated work)
      clearPlanningAreas(teamSheet);
      clearedCount++;
    } catch(e) {
      console.log(`Could not clear ${teamName}: ${e}`);
    }
  });
  
  ui.alert(
    'Planning Cleared',
    `Cleared planning from ${clearedCount} team sheet(s).\n\n` +
    'Team-initiated work has been preserved.\n\n' +
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
      if (!isStart) date.setDate(date.getDate() + 6);
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
    PLANNING_CONFIG.COLORS.SPRINT_4,
    PLANNING_CONFIG.COLORS.SPRINT_5,
    PLANNING_CONFIG.COLORS.SPRINT_6
  ];
  return colors[(sprintNumber - 1) % colors.length];
}

function openPlanningSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = getOrCreatePlanningConfig();
  ss.setActiveSheet(configSheet);
}