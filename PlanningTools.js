/**
 * Enhanced Planning Tools - Weekend-Aware Sprint Planning
 * Version: 5.0 - Working Days Only
 * 
 * NEW FEATURES:
 * 1. Start and end dates skip weekends (Monday-Friday only)
 * 2. Sprint capacity based on working days (5 days = 5 points per person per week)
 * 3. Proper working day calculations for sprint durations
 * 4. Smart date handling that respects business schedules
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  PLANNING_METHODS: ['Sprint', 'Waterfall'],
  SORT_OPTIONS: ['Sprint', 'Go-Live Date', 'Points'],
  DEFAULT_FIRST_SPRINT: 1,
  WORKING_DAYS_PER_WEEK: 5, // Monday-Friday
  POINTS_PER_PERSON_PER_WEEK: 5, // 1 point = 1 working day
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

// ==================== WEEKEND-AWARE DATE FUNCTIONS ====================

/**
 * Adds working days to a date, skipping weekends
 */
function addWorkingDays(startDate, workingDays) {
  const result = new Date(startDate);
  let daysAdded = 0;
  
  while (daysAdded < workingDays) {
    result.setDate(result.getDate() + 1);
    // Skip weekends (0 = Sunday, 6 = Saturday)
    if (result.getDay() !== 0 && result.getDay() !== 6) {
      daysAdded++;
    }
  }
  
  return result;
}

/**
 * Gets the next Monday from a given date
 */
function getNextMonday(date = new Date()) {
  const result = new Date(date);
  const day = result.getDay();
  const daysUntilMonday = day === 0 ? 1 : (8 - day) % 7 || 7;
  result.setDate(result.getDate() + daysUntilMonday);
  return result;
}

/**
 * Gets the previous Friday from a given date
 */
function getPreviousFriday(date) {
  const result = new Date(date);
  const day = result.getDay();
  const daysSinceFriday = day === 0 ? 2 : (day + 2) % 7;
  result.setDate(result.getDate() - daysSinceFriday);
  return result;
}

/**
 * Calculates working days between two dates
 */
function getWorkingDaysBetween(startDate, endDate) {
  let workingDays = 0;
  const current = new Date(startDate);
  
  while (current <= endDate) {
    if (current.getDay() !== 0 && current.getDay() !== 6) {
      workingDays++;
    }
    current.setDate(current.getDate() + 1);
  }
  
  return workingDays;
}

/**
 * Calculates sprint dates based on working days only
 */
function calculateSprintDates(startDate, sprintIndex, duration, isStart) {
  // Ensure we start on a Monday
  const sprintStartMonday = getNextMonday(startDate);
  
  let workingDaysPerSprint;
  switch(duration) {
    case '1 week':
      workingDaysPerSprint = 5; // 1 week = 5 working days
      break;
    case '2 weeks':
      workingDaysPerSprint = 10; // 2 weeks = 10 working days
      break;
    case '1 month':
      workingDaysPerSprint = 20; // 1 month = ~20 working days
      break;
    default:
      workingDaysPerSprint = 10;
  }
  
  if (isStart) {
    // Calculate start date for this sprint
    const totalWorkingDaysToSprint = sprintIndex * workingDaysPerSprint;
    return addWorkingDays(new Date(sprintStartMonday.getTime() - 24 * 60 * 60 * 1000), totalWorkingDaysToSprint);
  } else {
    // Calculate end date for this sprint
    const totalWorkingDaysToSprint = (sprintIndex + 1) * workingDaysPerSprint - 1;
    return addWorkingDays(new Date(sprintStartMonday.getTime() - 24 * 60 * 60 * 1000), totalWorkingDaysToSprint);
  }
}

/**
 * Calculates realistic sprint capacity based on team size and working days
 */
function calculateSprintCapacity(teamMembers, sprintDuration) {
  const membersCount = parseInt(teamMembers) || 1;
  
  switch(sprintDuration) {
    case '1 week':
      return membersCount * PLANNING_CONFIG.POINTS_PER_PERSON_PER_WEEK; // 5 points per person per week
    case '2 weeks':
      return membersCount * PLANNING_CONFIG.POINTS_PER_PERSON_PER_WEEK * 2; // 10 points per person per 2 weeks
    case '1 month':
      return membersCount * PLANNING_CONFIG.POINTS_PER_PERSON_PER_WEEK * 4; // 20 points per person per month
    default:
      return membersCount * 10; // Default to 2-week capacity
  }
}

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
    'NEW: Sprint planning now respects weekends and working days!\n' +
    '• Sprint dates are Monday-Friday only\n' +
    '• Capacity calculated as 5 points per person per week\n' +
    '• All dates skip weekends automatically\n\n' +
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
  sheet.setColumnWidth(3, 250);
  
  // Title
  sheet.getRange('A1:C1').merge()
    .setValue('PLANNING CONFIGURATION - Working Days Only')
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
  
  // Working days explanation
  sheet.getRange('C4').setValue('(Working days: 1 week=5 days, 2 weeks=10 days, 1 month=20 days)')
    .setFontStyle('italic')
    .setFontSize(9);
  
  // First Sprint Number
  sheet.getRange('A5').setValue('First Sprint Number:');
  sheet.getRange('B5').setValue(PLANNING_CONFIG.DEFAULT_FIRST_SPRINT)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Start Date (will be adjusted to next Monday)
  sheet.getRange('A6').setValue('Start Date:');
  const nextMonday = getNextMonday();
  sheet.getRange('B6').setValue(nextMonday)
    .setNumberFormat('yyyy-MM-dd')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  sheet.getRange('C6').setValue('(Automatically adjusted to next Monday)')
    .setFontStyle('italic')
    .setFontSize(9);
  
  // Working days info
  sheet.getRange('A8').setValue('Working Days Info:');
  sheet.getRange('B8').setValue('Monday - Friday only')
    .setFontWeight('bold')
    .setBackground('#E8F5E9');
  
  // Capacity calculation
  sheet.getRange('A9').setValue('Sprint Capacity:');
  sheet.getRange('B9').setValue('5 points per person per week')
    .setFontWeight('bold')
    .setBackground('#E8F5E9');
  
  // Sort By
  sheet.getRange('A11').setValue('Sort Items By:');
  const sortValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SORT_OPTIONS, true)
    .build();
  sheet.getRange('B11').setDataValidation(sortValidation)
    .setValue('Sprint')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  // Add border
  sheet.getRange(3, 1, 9, 2).setBorder(true, true, true, true, true, true);
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

// ==================== ENHANCED SPRINT PLANNING WITH WORKING DAYS ====================
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
  let startDate = new Date(configSheet.getRange('B6').getValue());
  
  // Ensure start date is a Monday
  startDate = getNextMonday(startDate);
  configSheet.getRange('B6').setValue(startDate); // Update config with Monday
  
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
    
    // Get team capacity and member count
    let netCapacity = 0;
    let teamMembers = 1;
    
    try {
      netCapacity = teamSheet.getRange('D7').getValue() || 0;
      teamMembers = teamSheet.getRange('B4').getValue() || 1;
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
      
      // Smart sprint assignment with working days
      smartSprintAssignmentWithWorkingDays(teamSheet, manifestItems, teamMembers, sprintDuration, firstSprintNumber, startDate, teamName);
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
      `Successfully applied weekend-aware sprint planning to ${successCount} team(s).\n\n` +
      'All dates respect working days (Monday-Friday only).',
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

// ==================== SIMPLIFIED AND IMPROVED SPRINT ASSIGNMENT ====================
function smartSprintAssignmentWithWorkingDays(teamSheet, items, teamMembers, sprintDuration, firstSprintNumber, startDate, teamName) {
  // Calculate realistic sprint capacity based on working days
  const sprintCapacity = calculateSprintCapacity(teamMembers, sprintDuration);
  
  // Calculate how many sprints we need
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const estimatedSprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
  
  // Sort items by go-live date first - this is the primary organizing principle
  items.sort((a, b) => {
    const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
    const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
    
    // Adjust weekend dates to previous Friday for comparison
    const adjustedDateA = dateA.getDay() === 0 || dateA.getDay() === 6 ? getPreviousFriday(dateA) : dateA;
    const adjustedDateB = dateB.getDay() === 0 || dateB.getDay() === 6 ? getPreviousFriday(dateB) : dateB;
    
    return adjustedDateA - adjustedDateB;
  });
  
  // Initialize sprint buckets
  const sprints = [];
  for (let i = 0; i < estimatedSprintsNeeded; i++) {
    const sprintNumber = firstSprintNumber + i;
    sprints.push({
      number: sprintNumber,
      items: [],
      totalPoints: 0,
      startDate: calculateSprintDates(startDate, i, sprintDuration, true),
      endDate: calculateSprintDates(startDate, i, sprintDuration, false),
      capacity: sprintCapacity
    });
  }
  
  // SIMPLE ALGORITHM: Fill sprints sequentially with capacity limits
  let currentSprintIndex = 0;
  
  items.forEach(item => {
    // Try to fit in current sprint
    let targetSprintIndex = currentSprintIndex;
    
    // If current sprint would be over capacity, try next sprint
    if (sprints[currentSprintIndex].totalPoints + item.points > sprintCapacity) {
      // Move to next sprint if available
      if (currentSprintIndex + 1 < sprints.length) {
        currentSprintIndex++;
        targetSprintIndex = currentSprintIndex;
      }
      // If we're at the last sprint, we have to fit it somewhere
    }
    
    // Special case: if item has a late go-live date and early sprints aren't full, 
    // try to put it in a later sprint
    const itemGoLiveDate = item.goLiveDate ? new Date(item.goLiveDate) : null;
    if (itemGoLiveDate && itemGoLiveDate.getFullYear() < 2099) {
      // Find the latest sprint that ends before this item's go-live date
      for (let i = sprints.length - 1; i >= 0; i--) {
        const sprint = sprints[i];
        if (sprint.endDate <= itemGoLiveDate && 
            sprint.totalPoints + item.points <= sprintCapacity * 1.1) {
          targetSprintIndex = i;
          break;
        }
      }
    }
    
    // Assign to target sprint
    const targetSprint = sprints[targetSprintIndex];
    item.sprint = targetSprint.number;
    item.sprintStart = targetSprint.startDate;
    item.sprintEnd = targetSprint.endDate;
    
    targetSprint.items.push(item);
    targetSprint.totalPoints += item.points;
  });
  
  // POST-PROCESSING: Simple rebalancing if there are severe imbalances
  simpleRebalance(sprints, sprintCapacity);
  
  // Write organized data back to sheet
  writeSprintDataWithWorkingDays(teamSheet, sprints, teamName, sprintCapacity);
}

// ==================== SIMPLE REBALANCING ====================
function simpleRebalance(sprints, targetCapacity) {
  let maxIterations = 3;
  
  while (maxIterations > 0) {
    let moved = false;
    maxIterations--;
    
    // Find overloaded and underloaded sprints
    for (let i = 0; i < sprints.length - 1; i++) {
      const currentSprint = sprints[i];
      const nextSprint = sprints[i + 1];
      
      // If current sprint is over capacity and next sprint has room
      if (currentSprint.totalPoints > targetCapacity && 
          nextSprint.totalPoints < targetCapacity * 0.8) {
        
        // Find smallest item that can be moved
        const movableItems = currentSprint.items
          .filter(item => {
            // Don't move items that would violate date constraints
            if (item.goLiveDate && item.goLiveDate instanceof Date && item.goLiveDate.getFullYear() < 2099) {
              return nextSprint.endDate <= item.goLiveDate;
            }
            return true;
          })
          .sort((a, b) => a.points - b.points); // Smallest first
        
        for (const item of movableItems) {
          // Would this move improve the balance?
          const newCurrentPoints = currentSprint.totalPoints - item.points;
          const newNextPoints = nextSprint.totalPoints + item.points;
          
          if (newCurrentPoints >= targetCapacity * 0.5 && 
              newNextPoints <= targetCapacity * 1.1) {
            
            // Move the item
            currentSprint.items = currentSprint.items.filter(i => i !== item);
            currentSprint.totalPoints -= item.points;
            
            nextSprint.items.push(item);
            nextSprint.totalPoints += item.points;
            
            // Update item assignment
            item.sprint = nextSprint.number;
            item.sprintStart = nextSprint.startDate;
            item.sprintEnd = nextSprint.endDate;
            
            moved = true;
            break;
          }
        }
      }
      
      if (moved) break; // Only one move per iteration
    }
    
    if (!moved) break; // No more beneficial moves found
  }
}

// ==================== WRITE SPRINT DATA WITH WORKING DAYS ====================
function writeSprintDataWithWorkingDays(teamSheet, sprints, teamName, sprintCapacity) {
  let currentRow = 14;
  
  sprints.forEach(sprint => {
    if (sprint.items.length === 0) return; // Skip empty sprints
    if (currentRow >= 61) return; // Don't overflow into team area
    
    // Calculate utilization percentage
    const utilization = Math.round((sprint.totalPoints / sprint.capacity) * 100);
    const utilizationIcon = utilization > 100 ? '⚠️' : utilization > 90 ? '🔥' : '✅';
    
    // Add sprint header with capacity info
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- ${teamName.toUpperCase()} SPRINT ${sprint.number} (${sprint.totalPoints}/${sprint.capacity} pts - ${utilization}% ${utilizationIcon}) ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    // Sort items within sprint by go-live date (adjusted for working days)
    sprint.items.sort((a, b) => {
      const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
      const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
      
      // Adjust weekend dates to previous Friday for sorting
      const adjustedDateA = dateA.getDay() === 0 || dateA.getDay() === 6 ? getPreviousFriday(dateA) : dateA;
      const adjustedDateB = dateB.getDay() === 0 || dateB.getDay() === 6 ? getPreviousFriday(dateB) : dateB;
      
      return adjustedDateA - adjustedDateB;
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

// ==================== ENHANCED WATERFALL PLANNING WITH WORKING DAYS ====================
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
  let startDate = new Date(configSheet.getRange('B6').getValue());
  
  // Ensure start date is a Monday
  startDate = getNextMonday(startDate);
  configSheet.getRange('B6').setValue(startDate);
  
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
    
    try {
      teamMembers = teamSheet.getRange('B4').getValue() || 1;
    } catch(e) {
      console.log(`Could not read team config for ${teamName}`);
    }
    
    // Collect all manifest items
    const manifestItems = collectManifestItems(teamSheet, teamName);
    
    if (manifestItems.length > 0) {
      // Apply sequential waterfall planning with working days
      applySequentialWaterfallWithWorkingDays(teamSheet, manifestItems, teamMembers, startDate);
      successCount++;
    }
  });
  
  if (successCount === 0) {
    SpreadsheetApp.getUi().alert('No manifest items found to distribute.');
  } else {
    SpreadsheetApp.getUi().alert(
      'Waterfall Planning Applied',
      `Successfully applied working-day waterfall planning to ${successCount} team(s).\n\n` +
      'All dates respect weekends (Monday-Friday scheduling only).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== APPLY SEQUENTIAL WATERFALL WITH WORKING DAYS ====================
function applySequentialWaterfallWithWorkingDays(teamSheet, items, teamMembers, startDate) {
  // Sort items by go-live date (adjusted for working days)
  items.sort((a, b) => {
    const dateA = a.goLiveDate ? new Date(a.goLiveDate) : new Date('2099-12-31');
    const dateB = b.goLiveDate ? new Date(b.goLiveDate) : new Date('2099-12-31');
    
    // Adjust weekend dates to previous Friday
    const adjustedDateA = dateA.getDay() === 0 || dateA.getDay() === 6 ? getPreviousFriday(dateA) : dateA;
    const adjustedDateB = dateB.getDay() === 0 || dateB.getDay() === 6 ? getPreviousFriday(dateB) : dateB;
    
    return adjustedDateA - adjustedDateB;
  });
  
  // Calculate team member schedules (working days only)
  const teamMemberSchedules = [];
  for (let i = 1; i <= teamMembers; i++) {
    teamMemberSchedules.push({
      id: i,
      name: `Team Member ${i}`,
      nextAvailableDate: getNextMonday(startDate), // Start on Monday
      totalPoints: 0
    });
  }
  
  // Assign items to team members sequentially (working days only)
  const scheduledItems = [];
  
  items.forEach(item => {
    // Find team member who will be available earliest
    const availableMember = teamMemberSchedules.reduce((earliest, current) => 
      current.nextAvailableDate < earliest.nextAvailableDate ? current : earliest
    );
    
    // Calculate task duration in working days (1 point = 1 working day, minimum 1 day)
    const taskWorkingDays = Math.max(1, Math.round(item.points));
    
    // Set start date (when this team member is available - ensure it's a working day)
    let taskStart = new Date(availableMember.nextAvailableDate);
    if (taskStart.getDay() === 0 || taskStart.getDay() === 6) {
      taskStart = getNextMonday(taskStart);
    }
    
    // Calculate end date using working days
    const taskEnd = addWorkingDays(taskStart, taskWorkingDays - 1);
    
    // Update team member's next available date (next working day after task ends)
    availableMember.nextAvailableDate = addWorkingDays(taskEnd, 1);
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

// ==================== UTILITY FUNCTIONS FOR TESTING ====================

/**
 * Test function to validate working day calculations
 * Run this to verify the weekend logic is working correctly
 */
function testWorkingDayLogic() {
  const startDate = new Date('2025-01-15'); // Wednesday
  
  console.log('Testing Working Day Logic:');
  console.log('Start date:', startDate.toDateString());
  console.log('Next Monday:', getNextMonday(startDate).toDateString());
  
  // Test adding working days
  const after5Days = addWorkingDays(startDate, 5);
  console.log('After 5 working days:', after5Days.toDateString());
  
  // Test sprint calculations
  const sprint1Start = calculateSprintDates(startDate, 0, '2 weeks', true);
  const sprint1End = calculateSprintDates(startDate, 0, '2 weeks', false);
  console.log('Sprint 1 Start:', sprint1Start.toDateString());
  console.log('Sprint 1 End:', sprint1End.toDateString());
  
  // Test sprint capacity
  const capacity = calculateSprintCapacity(3, '2 weeks');
  console.log('Sprint capacity for 3 people, 2 weeks:', capacity, 'points');
  
  SpreadsheetApp.getUi().alert(
    'Working Day Logic Test',
    'Check the console log for test results.\n\n' +
    'Key validations:\n' +
    '• Sprint dates are Monday-Friday only\n' +
    '• Sprint capacity = team members × 5 points per week\n' +
    '• Working days skip weekends correctly',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}