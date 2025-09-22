/**
 * Planning Tools - Fixed Sprint Distribution
 * Version: 6.0
 * 
 * FIXES:
 * 1. Sprint dates now properly respect Planning Config start date
 * 2. Better distribution algorithm that actually spreads work across sprints
 * 3. Proper capacity-based allocation
 * 4. Respects working days (Monday-Friday only)
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  PLANNING_METHODS: ['Sprint', 'Waterfall'],
  DEFAULT_FIRST_SPRINT: 1,
  WORKING_DAYS_PER_WEEK: 5,
  COLORS: {
    HEADER: '#4285F4',
    CONFIG_BG: '#F5F5F5',
    SPRINT_1: '#E8F5E9',
    SPRINT_2: '#FFF9C4', 
    SPRINT_3: '#FFE0B2',
    SPRINT_4: '#F3E5F5',
    SPRINT_5: '#E1F5FE',
    SPRINT_6: '#FCE4EC',
    SPRINT_SEPARATOR: '#E0E0E0'
  }
};

// ==================== DATE FUNCTIONS ====================
function getNextMonday(date = new Date()) {
  const result = new Date(date);
  const day = result.getDay();
  const daysUntilMonday = day === 0 ? 1 : (8 - day) % 7 || 7;
  if (daysUntilMonday > 0) {
    result.setDate(result.getDate() + daysUntilMonday);
  }
  return result;
}

function addWorkingDays(startDate, workingDays) {
  const result = new Date(startDate);
  let daysAdded = 0;
  
  while (daysAdded < workingDays) {
    result.setDate(result.getDate() + 1);
    if (result.getDay() !== 0 && result.getDay() !== 6) {
      daysAdded++;
    }
  }
  
  return result;
}

// ==================== INSTALLATION ====================
function installPlanningTools() {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createPlanningMenu') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('createPlanningMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  createPlanningMenu();
  
  SpreadsheetApp.getUi().alert(
    'Planning Tools Installed!',
    'Sprint planning will now properly distribute work and respect configured start dates.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createPlanningMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Planning Tools')
    .addItem('🎯 Apply Sprint Planning', 'applySprintPlanning')
    .addItem('🌊 Apply Waterfall Planning', 'applyWaterfallPlanning')
    .addSeparator()
    .addItem('⚙️ Planning Settings', 'openPlanningSettings')
    .addItem('🧹 Clear All Planning', 'clearAllPlanning')
    .addSeparator()
    .addItem('🔧 Reinstall Planning Tools', 'installPlanningTools')
    .addToUi();
}

// ==================== CONFIG MANAGEMENT ====================
function getOrCreatePlanningConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(PLANNING_CONFIG.CONFIG_SHEET_NAME);
  
  if (!configSheet) {
    configSheet = ss.insertSheet(PLANNING_CONFIG.CONFIG_SHEET_NAME);
    setupConfigSheet(configSheet);
  }
  
  return configSheet;
}

function setupConfigSheet(sheet) {
  sheet.clear();
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  
  sheet.getRange('A1:C1').merge()
    .setValue('PLANNING CONFIGURATION')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.HEADER)
    .setFontColor('#FFFFFF');
  
  sheet.getRange('A3').setValue('Planning Method:');
  const methodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sprint', 'Waterfall'], true)
    .build();
  sheet.getRange('B3').setDataValidation(methodValidation)
    .setValue('Sprint')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  sheet.getRange('A4').setValue('Sprint Duration:');
  const durationValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SPRINT_DURATIONS, true)
    .build();
  sheet.getRange('B4').setDataValidation(durationValidation)
    .setValue(PLANNING_CONFIG.DEFAULT_SPRINT_DURATION)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  sheet.getRange('A5').setValue('First Sprint Number:');
  sheet.getRange('B5').setValue(PLANNING_CONFIG.DEFAULT_FIRST_SPRINT)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  
  sheet.getRange('A6').setValue('Start Date:');
  const nextMonday = getNextMonday();
  sheet.getRange('B6').setValue(nextMonday)
    .setNumberFormat('yyyy-MM-dd')
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  sheet.getRange('C6').setValue('(Adjusted to Monday)')
    .setFontStyle('italic')
    .setFontSize(9);
  
  sheet.getRange(3, 1, 4, 2).setBorder(true, true, true, true, true, true);
}

// ==================== SPRINT PLANNING ====================
function applySprintPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!verifyPointsSystem()) {
    SpreadsheetApp.getUi().alert('Points System not found. Please set up Points System first.');
    return;
  }
  
  const configSheet = getOrCreatePlanningConfig();
  const sprintDuration = configSheet.getRange('B4').getValue();
  const firstSprintNumber = parseInt(configSheet.getRange('B5').getValue()) || 1;
  let startDate = new Date(configSheet.getRange('B6').getValue());
  
  // Ensure start date is Monday
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
    
    const netCapacity = teamSheet.getRange('D7').getValue() || 0;
    const manifestItems = collectManifestItems(teamSheet);
    
    if (manifestItems.length > 0) {
      clearPlanningAreas(teamSheet);
      addSprintHeaders(teamSheet);
      distributeToSprints(teamSheet, manifestItems, netCapacity, sprintDuration, firstSprintNumber, startDate);
      successCount++;
    }
  });
  
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('Sprint Planning Applied', 
      `Successfully applied sprint planning to ${successCount} team(s).`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== IMPROVED SPRINT DISTRIBUTION ====================
function distributeToSprints(teamSheet, items, teamNetCapacity, sprintDuration, firstSprintNumber, startDate) {
  // Calculate working days per sprint
  const workingDaysPerSprint = getWorkingDaysForDuration(sprintDuration);
  
  // Calculate capacity per sprint (proportional to sprint duration)
  const sprintCapacity = Math.round(teamNetCapacity * (workingDaysPerSprint / 20)); // 20 working days per month
  
  // Calculate total points and number of sprints needed
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const sprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
  
  // Initialize sprints with CORRECT dates from config
  const sprints = [];
  for (let i = 0; i < sprintsNeeded; i++) {
    const sprintNum = firstSprintNumber + i;
    sprints.push({
      number: sprintNum,
      items: [],
      totalPoints: 0,
      capacity: sprintCapacity,
      startDate: calculateSprintStart(startDate, i, workingDaysPerSprint),
      endDate: calculateSprintEnd(startDate, i, workingDaysPerSprint)
    });
  }
  
  // BALANCED DISTRIBUTION ALGORITHM
  // Goal: Even distribution across all sprints while respecting date constraints
  
  // Step 1: Separate items by date constraints
  const dateConstrainedItems = [];
  const flexibleItems = [];
  
  items.forEach(item => {
    if (item.goLiveDate && item.goLiveDate instanceof Date && item.goLiveDate.getFullYear() < 2099) {
      // Find which sprints this item can fit in based on date
      const validSprints = sprints.filter(sprint => sprint.endDate <= item.goLiveDate);
      if (validSprints.length === 1) {
        // Only one sprint option - must go there
        dateConstrainedItems.push({...item, mustBeSprint: validSprints[0].number});
      } else if (validSprints.length > 1) {
        // Multiple sprint options - flexible but date-aware
        flexibleItems.push({...item, validSprintNumbers: validSprints.map(s => s.number)});
      } else {
        // No valid sprint based on date - treat as flexible
        flexibleItems.push({...item, validSprintNumbers: sprints.map(s => s.number)});
      }
    } else {
      // No date constraint - fully flexible
      flexibleItems.push({...item, validSprintNumbers: sprints.map(s => s.number)});
    }
  });
  
  // Step 2: Place date-constrained items first
  dateConstrainedItems.forEach(item => {
    const sprint = sprints.find(s => s.number === item.mustBeSprint);
    if (sprint) {
      sprint.items.push(item);
      sprint.totalPoints += item.points;
    }
  });
  
  // Step 3: Sort flexible items by size (largest first for better bin packing)
  flexibleItems.sort((a, b) => b.points - a.points);
  
  // Step 4: Place flexible items using "best fit decreasing" strategy
  flexibleItems.forEach(item => {
    // Find valid sprint with lowest current load
    let bestSprint = null;
    let minLoad = Infinity;
    
    sprints.forEach(sprint => {
      // Check if this sprint is valid for this item
      const isValid = !item.validSprintNumbers || item.validSprintNumbers.includes(sprint.number);
      
      if (isValid) {
        // Prefer sprints that are under capacity
        const loadAfterAdding = sprint.totalPoints + item.points;
        const loadRatio = loadAfterAdding / sprintCapacity;
        
        // Scoring function: prefer sprints that will be closest to target capacity after adding
        const score = Math.abs(1.0 - loadRatio); // Distance from 100% capacity
        
        if (sprint.totalPoints < minLoad || 
            (loadRatio <= 1.1 && score < Math.abs(1.0 - (minLoad / sprintCapacity)))) {
          bestSprint = sprint;
          minLoad = sprint.totalPoints;
        }
      }
    });
    
    if (bestSprint) {
      bestSprint.items.push(item);
      bestSprint.totalPoints += item.points;
    }
  });
  
  // Step 5: Final rebalancing pass
  rebalanceSprints(sprints, sprintCapacity);
  
  // Write to sheet
  writeSprintsToSheet(teamSheet, sprints);
}

// ==================== ENHANCED REBALANCING ====================
function rebalanceSprints(sprints, targetCapacity) {
  // Aggressive rebalancing for even distribution
  const maxIterations = 20; // More iterations
  const targetUtilization = 0.9; // Aim for 90% utilization
  
  for (let iter = 0; iter < maxIterations; iter++) {
    let moved = false;
    
    // Calculate average load
    const avgLoad = sprints.reduce((sum, s) => sum + s.totalPoints, 0) / sprints.length;
    
    // Find most imbalanced pair of sprints
    let maxDiff = 0;
    let overloadedSprint = null;
    let underloadedSprint = null;
    
    for (let i = 0; i < sprints.length; i++) {
      for (let j = 0; j < sprints.length; j++) {
        if (i === j) continue;
        
        const diff = sprints[i].totalPoints - sprints[j].totalPoints;
        if (diff > maxDiff && diff > targetCapacity * 0.2) { // Only rebalance if difference > 20%
          maxDiff = diff;
          overloadedSprint = sprints[i];
          underloadedSprint = sprints[j];
        }
      }
    }
    
    if (!overloadedSprint || !underloadedSprint) break;
    
    // Try to move items from overloaded to underloaded
    overloadedSprint.items.sort((a, b) => a.points - b.points); // Start with smallest
    
    for (let i = 0; i < overloadedSprint.items.length; i++) {
      const item = overloadedSprint.items[i];
      
      // Check if moving this item would improve balance
      const newOverloadedTotal = overloadedSprint.totalPoints - item.points;
      const newUnderloadedTotal = underloadedSprint.totalPoints + item.points;
      
      // Only move if it improves overall balance
      const currentDiff = Math.abs(overloadedSprint.totalPoints - underloadedSprint.totalPoints);
      const newDiff = Math.abs(newOverloadedTotal - newUnderloadedTotal);
      
      if (newDiff < currentDiff && newUnderloadedTotal <= targetCapacity * 1.1) {
        // Check date constraint if moving to later sprint
        let canMove = true;
        if (underloadedSprint.number > overloadedSprint.number && 
            item.goLiveDate && item.goLiveDate instanceof Date) {
          canMove = underloadedSprint.endDate <= item.goLiveDate;
        }
        
        if (canMove) {
          // Move the item
          overloadedSprint.items.splice(i, 1);
          overloadedSprint.totalPoints = newOverloadedTotal;
          underloadedSprint.items.push(item);
          underloadedSprint.totalPoints = newUnderloadedTotal;
          moved = true;
          break;
        }
      }
    }
    
    if (!moved) {
      // Try medium-sized items if small items didn't work
      const midSizeItems = overloadedSprint.items.filter(i => 
        i.points >= 3 && i.points <= 5
      );
      
      for (const item of midSizeItems) {
        const newOverloadedTotal = overloadedSprint.totalPoints - item.points;
        const newUnderloadedTotal = underloadedSprint.totalPoints + item.points;
        
        if (Math.abs(newOverloadedTotal - avgLoad) < Math.abs(overloadedSprint.totalPoints - avgLoad) &&
            Math.abs(newUnderloadedTotal - avgLoad) < Math.abs(underloadedSprint.totalPoints - avgLoad)) {
          
          // Check date constraint
          let canMove = true;
          if (underloadedSprint.number > overloadedSprint.number && 
              item.goLiveDate && item.goLiveDate instanceof Date) {
            canMove = underloadedSprint.endDate <= item.goLiveDate;
          }
          
          if (canMove) {
            const idx = overloadedSprint.items.indexOf(item);
            overloadedSprint.items.splice(idx, 1);
            overloadedSprint.totalPoints = newOverloadedTotal;
            underloadedSprint.items.push(item);
            underloadedSprint.totalPoints = newUnderloadedTotal;
            moved = true;
            break;
          }
        }
      }
    }
    
    if (!moved) break;
  }
}

// ==================== DATE CALCULATIONS ====================
function getWorkingDaysForDuration(duration) {
  switch(duration) {
    case '1 week': return 5;
    case '2 weeks': return 10;
    case '1 month': return 20;
    default: return 10;
  }
}

function calculateSprintStart(configStartDate, sprintIndex, workingDaysPerSprint) {
  const totalWorkingDays = sprintIndex * workingDaysPerSprint;
  
  if (sprintIndex === 0) {
    return new Date(configStartDate); // First sprint starts on config date
  }
  
  // Add working days from start date
  return addWorkingDays(configStartDate, totalWorkingDays);
}

function calculateSprintEnd(configStartDate, sprintIndex, workingDaysPerSprint) {
  const totalWorkingDays = ((sprintIndex + 1) * workingDaysPerSprint) - 1;
  return addWorkingDays(configStartDate, totalWorkingDays);
}

// ==================== WRITING TO SHEET ====================
function writeSprintsToSheet(teamSheet, sprints) {
  let currentRow = 14;
  const teamName = teamSheet.getName().replace(' Team', '');
  
  sprints.forEach(sprint => {
    if (sprint.items.length === 0 || currentRow >= 61) return;
    
    // Calculate utilization
    const utilization = Math.round((sprint.totalPoints / sprint.capacity) * 100);
    const icon = utilization > 100 ? '🔥' : '✅';
    
    // Sprint header
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- ${teamName.toUpperCase()} SPRINT ${sprint.number} (${sprint.totalPoints}/${sprint.capacity} pts - ${utilization}% ${icon}) ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    // Sort items within sprint by go-live date
    sprint.items.sort((a, b) => {
      const dateA = a.goLiveDate || new Date('2099-12-31');
      const dateB = b.goLiveDate || new Date('2099-12-31');
      return dateA - dateB;
    });
    
    // Write items
    sprint.items.forEach(item => {
      if (currentRow >= 61) return;
      
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
      teamSheet.getRange(currentRow, 8).setValue(sprint.startDate).setNumberFormat('yyyy-MM-dd');
      teamSheet.getRange(currentRow, 9).setValue(sprint.endDate).setNumberFormat('yyyy-MM-dd');
      
      // Color code
      const sprintColor = getSprintColor(sprint.number);
      teamSheet.getRange(currentRow, 7, 1, 3).setBackground(sprintColor);
      
      currentRow++;
    });
    
    // Add spacing
    if (currentRow < 60) currentRow++;
  });
  
  // Apply border
  if (currentRow > 14) {
    teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
  }
}

// ==================== WATERFALL PLANNING ====================
function applyWaterfallPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!verifyPointsSystem()) {
    SpreadsheetApp.getUi().alert('Points System not found.');
    return;
  }
  
  const configSheet = getOrCreatePlanningConfig();
  let startDate = new Date(configSheet.getRange('B6').getValue());
  startDate = getNextMonday(startDate);
  
  const teams = getTeamNames();
  let successCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const manifestItems = collectManifestItems(teamSheet);
    
    if (manifestItems.length > 0) {
      clearPlanningAreas(teamSheet);
      addWaterfallHeaders(teamSheet);
      applySequentialWaterfall(teamSheet, manifestItems, startDate);
      successCount++;
    }
  });
  
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('Waterfall Planning Applied', 
      `Applied to ${successCount} team(s).`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function applySequentialWaterfall(teamSheet, items, startDate) {
  // Sort by go-live date
  items.sort((a, b) => {
    const dateA = a.goLiveDate || new Date('2099-12-31');
    const dateB = b.goLiveDate || new Date('2099-12-31');
    return dateA - dateB;
  });
  
  let currentDate = new Date(startDate);
  
  items.forEach(item => {
    // Skip weekends
    while (currentDate.getDay() === 0 || currentDate.getDay() === 6) {
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    item.startDate = new Date(currentDate);
    const workingDays = Math.max(1, Math.round(item.points));
    item.endDate = addWorkingDays(currentDate, workingDays - 1);
    
    currentDate = addWorkingDays(item.endDate, 1);
  });
  
  writeWaterfallToSheet(teamSheet, items);
}

// ==================== HELPER FUNCTIONS ====================
function collectManifestItems(teamSheet) {
  const items = [];
  
  // Collect from rows 14-60 (workstream assignments)
  for (let row = 14; row <= 60; row++) {
    const description = teamSheet.getRange(row, 2).getValue();
    const points = teamSheet.getRange(row, 4).getValue();
    
    if (description && points > 0 && !description.toString().startsWith('---')) {
      items.push({
        origin: teamSheet.getRange(row, 1).getValue() || '',
        description: description,
        size: teamSheet.getRange(row, 3).getValue() || '-',
        points: parseFloat(points) || 0,
        goLiveDate: teamSheet.getRange(row, 5).getValue(),
        source: teamSheet.getRange(row, 6).getValue() || ''
      });
    }
  }
  
  return items;
}

function clearPlanningAreas(teamSheet) {
  teamSheet.getRange(14, 1, 47, 9).clear();
  teamSheet.getRange(14, 1, 47, 9).setBackground('#FFFFFF');
  teamSheet.getRange(13, 7, 1, 3).clear();
}

function addSprintHeaders(teamSheet) {
  teamSheet.getRange('G13').setValue('Sprint').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('H13').setValue('Start Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('I13').setValue('End Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
}

function addWaterfallHeaders(teamSheet) {
  teamSheet.getRange('G13').setValue('Phase').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('H13').setValue('Start Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('I13').setValue('End Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
}

function writeWaterfallToSheet(teamSheet, items) {
  let currentRow = 14;
  
  items.forEach((item, index) => {
    if (currentRow >= 61) return;
    
    teamSheet.getRange(currentRow, 1).setValue(item.origin);
    teamSheet.getRange(currentRow, 2).setValue(item.description);
    teamSheet.getRange(currentRow, 3).setValue(item.size);
    teamSheet.getRange(currentRow, 4).setValue(item.points).setNumberFormat('0');
    
    if (item.goLiveDate) {
      teamSheet.getRange(currentRow, 5).setValue(item.goLiveDate).setNumberFormat('yyyy-MM-dd');
    }
    
    teamSheet.getRange(currentRow, 6).setValue(item.source);
    teamSheet.getRange(currentRow, 7).setValue(`Phase ${index + 1}`);
    teamSheet.getRange(currentRow, 8).setValue(item.startDate).setNumberFormat('yyyy-MM-dd');
    teamSheet.getRange(currentRow, 9).setValue(item.endDate).setNumberFormat('yyyy-MM-dd');
    
    teamSheet.getRange(currentRow, 7, 1, 3).setBackground('#E3F2FD');
    
    currentRow++;
  });
  
  if (currentRow > 14) {
    teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
  }
}

function clearAllPlanning() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Clear All Planning', 
    'This will remove sprint/waterfall assignments. Continue?', 
    ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) return;
  
  const teams = getTeamNames();
  let clearedCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teamName + ' Team');
    if (teamSheet) {
      clearPlanningAreas(teamSheet);
      clearedCount++;
    }
  });
  
  ui.alert('Planning Cleared', `Cleared ${clearedCount} team sheet(s).`, ui.ButtonSet.OK);
}

function verifyPointsSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Allocation') !== null;
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