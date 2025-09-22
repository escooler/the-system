/**
 * Planning Tools - Enhanced with Tweakable Plans
 * Version: 7.0
 * 
 * NEW FEATURES:
 * 1. Waterfall now groups by team members instead of phases
 * 2. Sprint/Person assignments are now dropdown-editable
 * 3. Refresh function reorganizes based on dropdown changes
 * 4. Removed phase concept from waterfall
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  PLANNING_METHODS: ['Sprint', 'Waterfall'],
  DEFAULT_FIRST_SPRINT: 1,
  DEFAULT_TEAM_SIZE: 5, // Default number of team members
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
    SPRINT_SEPARATOR: '#E0E0E0',
    PERSON_1: '#E8F5E9',
    PERSON_2: '#FFF9C4',
    PERSON_3: '#FFE0B2',
    PERSON_4: '#F3E5F5',
    PERSON_5: '#E1F5FE'
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
    'Version 7.0 - Now with editable assignments and refresh capability!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createPlanningMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Planning Tools')
    .addItem('🎯 Apply Sprint Planning', 'applySprintPlanning')
    .addItem('👥 Apply Waterfall Planning (by Person)', 'applyWaterfallPlanning')
    .addSeparator()
    .addItem('🔄 Refresh Planning Display', 'refreshPlanningDisplay')
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
  
  // Add team size configuration for waterfall
  sheet.getRange('A7').setValue('Team Size:');
  sheet.getRange('B7').setValue(PLANNING_CONFIG.DEFAULT_TEAM_SIZE)
    .setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  sheet.getRange('C7').setValue('(For waterfall planning)')
    .setFontStyle('italic')
    .setFontSize(9);
  
  sheet.getRange(3, 1, 5, 2).setBorder(true, true, true, true, true, true);
}

// ==================== SPRINT PLANNING WITH DROPDOWNS ====================
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
      
      // Calculate sprints needed
      const workingDaysPerSprint = getWorkingDaysForDuration(sprintDuration);
      const sprintCapacity = Math.round(netCapacity * (workingDaysPerSprint / 20));
      const totalPoints = manifestItems.reduce((sum, item) => sum + item.points, 0);
      const sprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
      
      // Distribute to sprints
      distributeToSprints(teamSheet, manifestItems, netCapacity, sprintDuration, firstSprintNumber, startDate);
      
      // Add sprint dropdowns
      addSprintDropdowns(teamSheet, manifestItems.length, sprintsNeeded, firstSprintNumber);
      
      successCount++;
    }
  });
  
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('Sprint Planning Applied', 
      `Successfully applied sprint planning to ${successCount} team(s).\n\nYou can now adjust assignments using the dropdowns and click "Refresh Planning Display" to reorganize.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== ADD SPRINT DROPDOWNS ====================
function addSprintDropdowns(teamSheet, itemCount, sprintsNeeded, firstSprintNumber) {
  // Create list of sprint options
  const sprintOptions = [];
  for (let i = 0; i < sprintsNeeded; i++) {
    sprintOptions.push(`Sprint ${firstSprintNumber + i}`);
  }
  
  // Add validation to column G (Sprint assignment)
  let currentRow = 14;
  let actualItemCount = 0;
  
  // Count actual items (skip headers)
  for (let row = 14; row <= Math.min(60, 14 + itemCount * 1.5); row++) {
    const description = teamSheet.getRange(row, 2).getValue();
    if (description && !description.toString().startsWith('---')) {
      actualItemCount++;
    }
  }
  
  // Apply dropdowns to actual items
  currentRow = 14;
  let itemsProcessed = 0;
  
  while (currentRow <= 60 && itemsProcessed < actualItemCount) {
    const description = teamSheet.getRange(currentRow, 2).getValue();  // Fixed: was 'row', should be 'currentRow'
    
    if (description && !description.toString().startsWith('---')) {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInList(sprintOptions, true)
        .setAllowInvalid(false)
        .build();
      
      teamSheet.getRange(currentRow, 7).setDataValidation(validation);
      itemsProcessed++;
    }
    currentRow++;
  }
}

// ==================== WATERFALL PLANNING WITH PERSON ASSIGNMENT ====================
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
    
    // Read team size from the team sheet itself (B4 = Team Members)
    const teamSize = parseInt(teamSheet.getRange('B4').getValue()) || PLANNING_CONFIG.DEFAULT_TEAM_SIZE;
    const netCapacity = teamSheet.getRange('D7').getValue() || 0;
    const manifestItems = collectManifestItems(teamSheet);
    
    if (manifestItems.length > 0) {
      clearPlanningAreas(teamSheet);
      addWaterfallHeaders(teamSheet);
      
      // Distribute to people
      distributeToTeamMembers(teamSheet, manifestItems, teamSize, netCapacity, startDate);
      
      // Add person dropdowns
      addPersonDropdowns(teamSheet, manifestItems.length, teamSize);
      
      successCount++;
    }
  });
  
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('Waterfall Planning Applied', 
      `Applied to ${successCount} team(s) using each team's member count.\n\nYou can adjust assignments and click "Refresh Planning Display" to reorganize.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== DISTRIBUTE TO TEAM MEMBERS ====================
function distributeToTeamMembers(teamSheet, items, teamSize, teamNetCapacity, startDate) {
  // Calculate capacity per person
  const capacityPerPerson = Math.round(teamNetCapacity / teamSize);
  
  // Initialize team members
  const teamMembers = [];
  for (let i = 0; i < teamSize; i++) {
    teamMembers.push({
      number: i + 1,
      items: [],
      totalPoints: 0,
      capacity: capacityPerPerson,
      currentDate: new Date(startDate)
    });
  }
  
  // Sort items by go-live date priority
  items.sort((a, b) => {
    const dateA = a.goLiveDate || new Date('2099-12-31');
    const dateB = b.goLiveDate || new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Distribute items using round-robin with load balancing
  items.forEach(item => {
    // Find person with least load
    let selectedPerson = teamMembers.reduce((minPerson, person) => 
      person.totalPoints < minPerson.totalPoints ? person : minPerson
    );
    
    // Calculate dates for this item
    const workingDays = Math.max(1, Math.round(item.points));
    item.startDate = new Date(selectedPerson.currentDate);
    
    // Skip weekends for start date
    while (item.startDate.getDay() === 0 || item.startDate.getDay() === 6) {
      item.startDate.setDate(item.startDate.getDate() + 1);
    }
    
    item.endDate = addWorkingDays(item.startDate, workingDays - 1);
    item.assignedPerson = selectedPerson.number;
    
    // Add to person's list
    selectedPerson.items.push(item);
    selectedPerson.totalPoints += item.points;
    selectedPerson.currentDate = addWorkingDays(item.endDate, 1);
  });
  
  // Write to sheet grouped by person
  writePersonGroupsToSheet(teamSheet, teamMembers);
}

// ==================== WRITE PERSON GROUPS TO SHEET ====================
function writePersonGroupsToSheet(teamSheet, teamMembers) {
  let currentRow = 14;
  const teamName = teamSheet.getName().replace(' Team', '');
  
  teamMembers.forEach(person => {
    if (person.items.length === 0 || currentRow >= 61) return;
    
    // Calculate utilization
    const utilization = Math.round((person.totalPoints / person.capacity) * 100);
    const icon = utilization > 100 ? '🔥' : '✅';
    
    // Person header
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- PERSON ${person.number} (${person.totalPoints}/${person.capacity} pts - ${utilization}% ${icon}) ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    // Sort items by start date
    person.items.sort((a, b) => a.startDate - b.startDate);
    
    // Write items
    person.items.forEach(item => {
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
      teamSheet.getRange(currentRow, 7).setValue(`Person ${person.number}`);
      teamSheet.getRange(currentRow, 8).setValue(item.startDate).setNumberFormat('yyyy-MM-dd');
      teamSheet.getRange(currentRow, 9).setValue(item.endDate).setNumberFormat('yyyy-MM-dd');
      
      // Color code
      const personColor = getPersonColor(person.number);
      teamSheet.getRange(currentRow, 7, 1, 3).setBackground(personColor);
      
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

// ==================== ADD PERSON DROPDOWNS ====================
function addPersonDropdowns(teamSheet, itemCount, teamSize) {
  // Create list of person options
  const personOptions = [];
  for (let i = 1; i <= teamSize; i++) {
    personOptions.push(`Person ${i}`);
  }
  
  // Apply dropdowns to column G
  let currentRow = 14;
  let itemsProcessed = 0;
  
  while (currentRow <= 60 && itemsProcessed < itemCount * 1.5) {
    const description = teamSheet.getRange(currentRow, 2).getValue();
    
    if (description && !description.toString().startsWith('---')) {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInList(personOptions, true)
        .setAllowInvalid(false)
        .build();
      
      teamSheet.getRange(currentRow, 7).setDataValidation(validation);
      itemsProcessed++;
    }
    currentRow++;
  }
}

// ==================== REFRESH PLANNING DISPLAY ====================
function refreshPlanningDisplay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found.');
    return;
  }
  
  let refreshCount = 0;
  let notFoundCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // More robust detection - look for any dropdown or assignment value in column G
    let hasPlan = false;
    let firstValidAssignment = null;
    
    // Check multiple rows for planning evidence
    for (let checkRow = 14; checkRow <= 25; checkRow++) {
      const dropdownCheck = teamSheet.getRange(checkRow, 7).getDataValidation();
      const valueCheck = teamSheet.getRange(checkRow, 7).getValue();
      
      if (dropdownCheck || (valueCheck && (valueCheck.toString().includes('Sprint') || valueCheck.toString().includes('Person')))) {
        hasPlan = true;
        if (!firstValidAssignment && valueCheck) {
          firstValidAssignment = valueCheck;
        }
        break;
      }
    }
    
    if (!hasPlan) {
      notFoundCount++;
      return;
    }
    
    // Determine if it's Sprint or Person based
    const isSprint = firstValidAssignment && firstValidAssignment.toString().includes('Sprint');
    
    // Collect all items with their current assignments
    const items = [];
    for (let row = 14; row <= 60; row++) {
      const description = teamSheet.getRange(row, 2).getValue();
      const assignment = teamSheet.getRange(row, 7).getValue();
      
      // Skip empty rows and headers - be more flexible with header detection
      if (description && 
          !description.toString().includes('---') && 
          !description.toString().includes('SPRINT') && 
          !description.toString().includes('PERSON') &&
          assignment && 
          (assignment.toString().includes('Sprint') || assignment.toString().includes('Person'))) {
        
        // Store the current dropdown validation
        const validation = teamSheet.getRange(row, 7).getDataValidation();
        
        items.push({
          row: row,
          origin: teamSheet.getRange(row, 1).getValue() || '',
          description: description,
          size: teamSheet.getRange(row, 3).getValue() || '-',
          points: parseFloat(teamSheet.getRange(row, 4).getValue()) || 0,
          goLiveDate: teamSheet.getRange(row, 5).getValue(),
          source: teamSheet.getRange(row, 6).getValue() || '',
          assignment: assignment,
          startDate: teamSheet.getRange(row, 8).getValue(),
          endDate: teamSheet.getRange(row, 9).getValue(),
          validation: validation  // Store the dropdown validation
        });
      }
    }
    
    if (items.length === 0) {
      notFoundCount++;
      return;
    }
    
    // Group items by assignment
    const groups = {};
    items.forEach(item => {
      if (!groups[item.assignment]) {
        groups[item.assignment] = [];
      }
      groups[item.assignment].push(item);
    });
    
    // Clear the entire manifest area (14-60) but preserve column headers
    teamSheet.getRange(14, 1, 47, 9).clear();
    teamSheet.getRange(14, 1, 47, 9).clearDataValidations();
    teamSheet.getRange(14, 1, 47, 9).setBackground('#FFFFFF');
    
    // Rewrite organized by groups
    let currentRow = 14;
    
    // Sort group keys
    const sortedKeys = Object.keys(groups).sort((a, b) => {
      // Extract numbers for proper sorting
      const matchA = a.match(/\d+/);
      const matchB = b.match(/\d+/);
      const numA = matchA ? parseInt(matchA[0]) : 0;
      const numB = matchB ? parseInt(matchB[0]) : 0;
      return numA - numB;
    });
    
    // Get team capacity info for calculations
    const netCapacity = teamSheet.getRange('D7').getValue() || 100;
    const teamMembers = teamSheet.getRange('B4').getValue() || 1;
    
    // Calculate capacity per group
    let groupCapacity;
    if (isSprint) {
      // For sprints, divide by number of sprints
      groupCapacity = Math.round(netCapacity / sortedKeys.length);
    } else {
      // For persons, divide by number of team members
      groupCapacity = Math.round(netCapacity / teamMembers);
    }
    
    sortedKeys.forEach(groupKey => {
      if (currentRow >= 61) return;
      
      const groupItems = groups[groupKey];
      const totalPoints = groupItems.reduce((sum, item) => sum + item.points, 0);
      const utilization = Math.round((totalPoints / groupCapacity) * 100);
      const icon = utilization > 100 ? '🔥' : '✅';
      
      // Group header
      const headerText = isSprint ? 
        `--- ${teamName.toUpperCase()} ${groupKey.toUpperCase()} (${totalPoints}/${groupCapacity} pts - ${utilization}% ${icon}) ---` :
        `--- ${groupKey.toUpperCase()} (${totalPoints}/${groupCapacity} pts - ${utilization}% ${icon}) ---`;
      
      teamSheet.getRange(currentRow, 1, 1, 9).merge()
        .setValue(headerText)
        .setFontWeight('bold')
        .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
        .setFontStyle('italic');
      currentRow++;
      
      // Sort items within group by go-live date
      groupItems.sort((a, b) => {
        const dateA = a.goLiveDate || new Date('2099-12-31');
        const dateB = b.goLiveDate || new Date('2099-12-31');
        
        if (dateA instanceof Date && dateB instanceof Date) {
          return dateA.getTime() - dateB.getTime();
        }
        return String(dateA).localeCompare(String(dateB));
      });
      
      // Write items
      groupItems.forEach(item => {
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
        
        // Restore the dropdown with the current assignment
        if (item.validation) {
          teamSheet.getRange(currentRow, 7).setDataValidation(item.validation);
        }
        teamSheet.getRange(currentRow, 7).setValue(item.assignment);
        
        // Keep existing dates if they exist
        if (item.startDate) {
          teamSheet.getRange(currentRow, 8).setValue(item.startDate);
          if (item.startDate instanceof Date) {
            teamSheet.getRange(currentRow, 8).setNumberFormat('yyyy-MM-dd');
          }
        }
        if (item.endDate) {
          teamSheet.getRange(currentRow, 9).setValue(item.endDate);
          if (item.endDate instanceof Date) {
            teamSheet.getRange(currentRow, 9).setNumberFormat('yyyy-MM-dd');
          }
        }
        
        // Color code based on assignment type
        if (isSprint) {
          const sprintMatch = item.assignment.match(/\d+/);
          const sprintNum = sprintMatch ? parseInt(sprintMatch[0]) : 1;
          teamSheet.getRange(currentRow, 7, 1, 3).setBackground(getSprintColor(sprintNum));
        } else {
          const personMatch = item.assignment.match(/\d+/);
          const personNum = personMatch ? parseInt(personMatch[0]) : 1;
          teamSheet.getRange(currentRow, 7, 1, 3).setBackground(getPersonColor(personNum));
        }
        
        currentRow++;
      });
      
      // Add spacing between groups
      if (currentRow < 60 && currentRow < 14 + items.length + sortedKeys.length * 2) {
        currentRow++;
      }
    });
    
    // Apply border to the populated area
    if (currentRow > 14) {
      teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
    }
    
    refreshCount++;
  });
  
  // Provide appropriate feedback
  if (refreshCount > 0) {
    SpreadsheetApp.getUi().alert('Planning Refreshed', 
      `Successfully reorganized ${refreshCount} team sheet(s) based on dropdown assignments.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else if (notFoundCount > 0) {
    SpreadsheetApp.getUi().alert('No Planning Found', 
      'No valid planning assignments found to refresh. Make sure you have:\n\n' +
      '1. Applied Sprint or Waterfall planning first\n' +
      '2. Items have Sprint/Person assignments in column G\n\n' +
      'Check team sheets and try again.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Team Sheets', 
      'No team sheets found in the spreadsheet.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== EXISTING SPRINT DISTRIBUTION (ENHANCED) ====================
function distributeToSprints(teamSheet, items, teamNetCapacity, sprintDuration, firstSprintNumber, startDate) {
  const workingDaysPerSprint = getWorkingDaysForDuration(sprintDuration);
  const sprintCapacity = Math.round(teamNetCapacity * (workingDaysPerSprint / 20));
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const sprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
  
  // Initialize sprints
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
  
  // Sort items by priority/date
  items.sort((a, b) => {
    const dateA = a.goLiveDate || new Date('2099-12-31');
    const dateB = b.goLiveDate || new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Distribute items
  items.forEach(item => {
    // Find best sprint
    let bestSprint = sprints[0];
    let minLoad = Infinity;
    
    sprints.forEach(sprint => {
      if (sprint.totalPoints < minLoad) {
        bestSprint = sprint;
        minLoad = sprint.totalPoints;
      }
    });
    
    bestSprint.items.push(item);
    bestSprint.totalPoints += item.points;
  });
  
  // Balance sprints
  rebalanceSprints(sprints, sprintCapacity);
  
  // Write to sheet
  writeSprintsToSheet(teamSheet, sprints);
}

// ==================== HELPER FUNCTIONS ====================
function collectManifestItems(teamSheet) {
  const items = [];
  
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
  teamSheet.getRange(14, 1, 47, 9).clearDataValidations();
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
  teamSheet.getRange('G13').setValue('Person').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('H13').setValue('Start Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange('I13').setValue('End Date').setFontWeight('bold').setBackground('#E1BEE7');
  teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
}

function clearAllPlanning() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Clear All Planning', 
    'This will remove all sprint/person assignments. Continue?', 
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
    return new Date(configStartDate);
  }
  
  return addWorkingDays(configStartDate, totalWorkingDays);
}

function calculateSprintEnd(configStartDate, sprintIndex, workingDaysPerSprint) {
  const totalWorkingDays = ((sprintIndex + 1) * workingDaysPerSprint) - 1;
  return addWorkingDays(configStartDate, totalWorkingDays);
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

function getPersonColor(personNumber) {
  const colors = [
    PLANNING_CONFIG.COLORS.PERSON_1,
    PLANNING_CONFIG.COLORS.PERSON_2,
    PLANNING_CONFIG.COLORS.PERSON_3,
    PLANNING_CONFIG.COLORS.PERSON_4,
    PLANNING_CONFIG.COLORS.PERSON_5
  ];
  return colors[(personNumber - 1) % colors.length];
}

function openPlanningSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = getOrCreatePlanningConfig();
  ss.setActiveSheet(configSheet);
}

// ==================== WRITE SPRINTS TO SHEET ====================
function writeSprintsToSheet(teamSheet, sprints) {
  let currentRow = 14;
  const teamName = teamSheet.getName().replace(' Team', '');
  
  sprints.forEach(sprint => {
    if (sprint.items.length === 0 || currentRow >= 61) return;
    
    const utilization = Math.round((sprint.totalPoints / sprint.capacity) * 100);
    const icon = utilization > 100 ? '🔥' : '✅';
    
    teamSheet.getRange(currentRow, 1, 1, 9).merge()
      .setValue(`--- ${teamName.toUpperCase()} SPRINT ${sprint.number} (${sprint.totalPoints}/${sprint.capacity} pts - ${utilization}% ${icon}) ---`)
      .setFontWeight('bold')
      .setBackground(PLANNING_CONFIG.COLORS.SPRINT_SEPARATOR)
      .setFontStyle('italic');
    currentRow++;
    
    sprint.items.sort((a, b) => {
      const dateA = a.goLiveDate || new Date('2099-12-31');
      const dateB = b.goLiveDate || new Date('2099-12-31');
      return dateA - dateB;
    });
    
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
      
      const sprintColor = getSprintColor(sprint.number);
      teamSheet.getRange(currentRow, 7, 1, 3).setBackground(sprintColor);
      
      currentRow++;
    });
    
    if (currentRow < 60) currentRow++;
  });
  
  if (currentRow > 14) {
    teamSheet.getRange(14, 1, currentRow - 14, 9).setBorder(true, true, true, true, true, false);
  }
}

// ==================== REBALANCE SPRINTS ====================
function rebalanceSprints(sprints, targetCapacity) {
  const maxIterations = 20;
  
  for (let iter = 0; iter < maxIterations; iter++) {
    let moved = false;
    
    const avgLoad = sprints.reduce((sum, s) => sum + s.totalPoints, 0) / sprints.length;
    
    let maxDiff = 0;
    let overloadedSprint = null;
    let underloadedSprint = null;
    
    for (let i = 0; i < sprints.length; i++) {
      for (let j = 0; j < sprints.length; j++) {
        if (i === j) continue;
        
        const diff = sprints[i].totalPoints - sprints[j].totalPoints;
        if (diff > maxDiff && diff > targetCapacity * 0.2) {
          maxDiff = diff;
          overloadedSprint = sprints[i];
          underloadedSprint = sprints[j];
        }
      }
    }
    
    if (!overloadedSprint || !underloadedSprint) break;
    
    overloadedSprint.items.sort((a, b) => a.points - b.points);
    
    for (let i = 0; i < overloadedSprint.items.length; i++) {
      const item = overloadedSprint.items[i];
      
      const newOverloadedTotal = overloadedSprint.totalPoints - item.points;
      const newUnderloadedTotal = underloadedSprint.totalPoints + item.points;
      
      const currentDiff = Math.abs(overloadedSprint.totalPoints - underloadedSprint.totalPoints);
      const newDiff = Math.abs(newOverloadedTotal - newUnderloadedTotal);
      
      if (newDiff < currentDiff && newUnderloadedTotal <= targetCapacity * 1.1) {
        let canMove = true;
        if (underloadedSprint.number > overloadedSprint.number && 
            item.goLiveDate && item.goLiveDate instanceof Date) {
          canMove = underloadedSprint.endDate <= item.goLiveDate;
        }
        
        if (canMove) {
          overloadedSprint.items.splice(i, 1);
          overloadedSprint.totalPoints = newOverloadedTotal;
          underloadedSprint.items.push(item);
          underloadedSprint.totalPoints = newUnderloadedTotal;
          moved = true;
          break;
        }
      }
    }
    
    if (!moved) break;
  }
}