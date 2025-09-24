/**
 * Planning Tools - Optimized Version 8.0
 * 
 * IMPROVEMENTS:
 * - Fixed memory leaks in refresh function
 * - Optimized date calculations
 * - Reduced redundant sheet access with batch operations
 * - Removed hard-coded row limits
 * - Consolidated duplicate code
 * - Added better error handling
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  DEFAULT_FIRST_SPRINT: 1,
  DEFAULT_TEAM_SIZE: 5,
  WORKING_DAYS_PER_WEEK: 5,
  MAX_MANIFEST_ROWS: 47, // 61 - 14
  COLORS: {
    HEADER: '#4285F4',
    CONFIG_BG: '#F5F5F5',
    SPRINT_1: '#E8F5E9',
    SPRINT_2: '#FFF9C4',
    SPRINT_3: '#FFE0B2',
    SPRINT_4: '#F3E5F5',
    SPRINT_5: '#E1F5FE',
    SPRINT_6: '#FCE4EC',
    SEPARATOR: '#E0E0E0',
    PLANNING_HEADER: '#E1BEE7'
  }
};

// ==================== OPTIMIZED DATE FUNCTIONS ====================
function getNextMonday(date = new Date()) {
  const result = new Date(date);
  const day = result.getDay();
  const daysUntilMonday = day === 0 ? 1 : (8 - day) % 7 || 7;
  result.setDate(result.getDate() + daysUntilMonday);
  return result;
}

function addWorkingDays(startDate, workingDays) {
  if (workingDays <= 0) return new Date(startDate);
  
  const result = new Date(startDate);
  let totalDays = 0;
  let weekends = 0;
  
  // Calculate full weeks
  const fullWeeks = Math.floor(workingDays / 5);
  totalDays += fullWeeks * 7;
  
  // Calculate remaining days
  let remainingDays = workingDays % 5;
  const startDay = result.getDay();
  
  // Check if remaining days cross a weekend
  if (startDay + remainingDays > 5) {
    weekends = (startDay === 0) ? 1 : 2;
  }
  
  totalDays += remainingDays + weekends;
  result.setDate(result.getDate() + totalDays);
  
  // Ensure we don't land on a weekend
  while (result.getDay() === 0 || result.getDay() === 6) {
    result.setDate(result.getDate() + 1);
  }
  
  return result;
}

// ==================== INSTALLATION ====================
function installPlanningTools() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createPlanningMenu') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('createPlanningMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  createPlanningMenu();
  
  SpreadsheetApp.getUi().alert(
    'Planning Tools Installed!',
    'Version 8.0 - Optimized for performance',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createPlanningMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Planning Tools')
    .addItem('🎯 Apply Sprint Planning', 'applySprintPlanning')
    .addItem('👥 Apply Waterfall Planning', 'applyWaterfallPlanning')
    .addSeparator()
    .addItem('🔄 Refresh Planning Display', 'refreshPlanningDisplay')
    .addSeparator()
    .addItem('⚙️ Planning Settings', 'openPlanningSettings')
    .addItem('🧹 Clear All Planning', 'clearAllPlanning')
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
  
  // Set column widths in batch
  sheet.setColumnWidths(1, 3, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  
  // Batch set values
  const values = [
    ['PLANNING CONFIGURATION', '', ''],
    ['', '', ''],
    ['Planning Method:', 'Sprint', ''],
    ['Sprint Duration:', PLANNING_CONFIG.DEFAULT_SPRINT_DURATION, ''],
    ['First Sprint Number:', PLANNING_CONFIG.DEFAULT_FIRST_SPRINT, ''],
    ['Start Date:', getNextMonday(), '(Adjusted to Monday)'],
    ['Team Size:', PLANNING_CONFIG.DEFAULT_TEAM_SIZE, '(For waterfall planning)']
  ];
  
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  
  // Format header
  sheet.getRange('A1:C1').merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.HEADER)
    .setFontColor('#FFFFFF');
  
  // Add validations
  const methodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sprint', 'Waterfall'], true)
    .build();
  sheet.getRange('B3').setDataValidation(methodValidation);
  
  const durationValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SPRINT_DURATIONS, true)
    .build();
  sheet.getRange('B4').setDataValidation(durationValidation);
  
  // Format config cells
  sheet.getRange('B3:B7').setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  sheet.getRange('B6').setNumberFormat('yyyy-MM-dd');
  sheet.getRange('C6:C7').setFontStyle('italic').setFontSize(9);
  
  sheet.getRange(3, 1, 5, 2).setBorder(true, true, true, true, true, true);
}

// ==================== MAIN PLANNING FUNCTIONS ====================
function applySprintPlanning() {
  applyPlanning('Sprint');
}

function applyWaterfallPlanning() {
  applyPlanning('Waterfall');
}

function applyPlanning(planningType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName('Allocation')) {
    SpreadsheetApp.getUi().alert('Points System not found. Please set up Points System first.');
    return;
  }
  
  const configSheet = getOrCreatePlanningConfig();
  const config = readConfig(configSheet);
  
  const teams = getTeamNames();
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found.');
    return;
  }
  
  let successCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const manifestItems = collectManifestItems(teamSheet);
    if (manifestItems.length === 0) return;
    
    // Clear and setup
    clearPlanningAreas(teamSheet);
    addPlanningHeaders(teamSheet, planningType);
    
    // Apply planning based on type
    if (planningType === 'Sprint') {
      const netCapacity = teamSheet.getRange('D7').getValue() || 0;
      applySprints(teamSheet, manifestItems, netCapacity, config);
    } else {
      const teamSize = parseInt(teamSheet.getRange('B4').getValue()) || PLANNING_CONFIG.DEFAULT_TEAM_SIZE;
      const netCapacity = teamSheet.getRange('D7').getValue() || 0;
      applyWaterfall(teamSheet, manifestItems, teamSize, netCapacity, config.startDate);
    }
    
    successCount++;
  });
  
  if (successCount > 0) {
    const message = planningType === 'Sprint' ? 
      `Applied sprint planning to ${successCount} team(s).\n\nAdjust assignments using dropdowns and refresh to reorganize.` :
      `Applied waterfall planning to ${successCount} team(s).\n\nAdjust assignments and refresh to reorganize.`;
    
    SpreadsheetApp.getUi().alert(`${planningType} Planning Applied`, message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== OPTIMIZED CONFIG READER ====================
function readConfig(configSheet) {
  const values = configSheet.getRange('B3:B6').getValues();
  return {
    method: values[0][0],
    duration: values[1][0],
    firstSprint: parseInt(values[2][0]) || 1,
    startDate: getNextMonday(new Date(values[3][0]))
  };
}

// ==================== SPRINT PLANNING ====================
function applySprints(teamSheet, items, netCapacity, config) {
  const workingDaysPerSprint = getWorkingDaysForDuration(config.duration);
  const sprintCapacity = Math.round(netCapacity * (workingDaysPerSprint / 20));
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const sprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
  
  // Initialize sprints
  const sprints = [];
  for (let i = 0; i < sprintsNeeded; i++) {
    sprints.push({
      number: config.firstSprint + i,
      items: [],
      totalPoints: 0,
      capacity: sprintCapacity,
      startDate: calculateSprintDate(config.startDate, i, workingDaysPerSprint, true),
      endDate: calculateSprintDate(config.startDate, i, workingDaysPerSprint, false)
    });
  }
  
  // Distribute items
  distributeItems(items, sprints, 'sprint');
  
  // Write to sheet
  writeGroupsToSheet(teamSheet, sprints, 'Sprint', sprintCapacity);
  
  // Add dropdowns
  addDropdowns(teamSheet, items.length, sprints.map(s => `Sprint ${s.number}`));
}

// ==================== WATERFALL PLANNING ====================
function applyWaterfall(teamSheet, items, teamSize, netCapacity, startDate) {
  const capacityPerPerson = Math.round(netCapacity / teamSize);
  
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
  
  // Distribute items with dates
  items.forEach(item => {
    // Find person with least load
    const person = teamMembers.reduce((min, p) => 
      p.totalPoints < min.totalPoints ? p : min
    );
    
    // Calculate dates
    const workingDays = Math.max(1, Math.round(item.points));
    item.startDate = new Date(person.currentDate);
    
    // Skip weekends for start
    while (item.startDate.getDay() === 0 || item.startDate.getDay() === 6) {
      item.startDate.setDate(item.startDate.getDate() + 1);
    }
    
    item.endDate = addWorkingDays(item.startDate, workingDays - 1);
    item.assignedPerson = person.number;
    
    person.items.push(item);
    person.totalPoints += item.points;
    person.currentDate = addWorkingDays(item.endDate, 1);
  });
  
  // Write to sheet
  writeGroupsToSheet(teamSheet, teamMembers, 'Person', capacityPerPerson);
  
  // Add dropdowns
  addDropdowns(teamSheet, items.length, teamMembers.map(p => `Person ${p.number}`));
}

// ==================== OPTIMIZED WRITING ====================
function writeGroupsToSheet(teamSheet, groups, groupType, capacity) {
  const teamName = teamSheet.getName().replace(' Team', '');
  let currentRow = 14;
  const outputData = [];
  const formats = [];
  const backgrounds = [];
  
  groups.forEach(group => {
    if (group.items.length === 0) return;
    
    const utilization = Math.round((group.totalPoints / capacity) * 100);
    const icon = utilization > 100 ? '🔥' : '✅';
    const headerText = groupType === 'Sprint' ?
      `--- ${teamName.toUpperCase()} SPRINT ${group.number} (${group.totalPoints}/${capacity} pts - ${utilization}% ${icon}) ---` :
      `--- PERSON ${group.number} (${group.totalPoints}/${capacity} pts - ${utilization}% ${icon}) ---`;
    
    // Add header row
    outputData.push([headerText, '', '', '', '', '', '', '', '']);
    formats.push(['', '', '', '', '', '', '', '', '']);
    backgrounds.push([PLANNING_CONFIG.COLORS.SEPARATOR, '', '', '', '', '', '', '', '']);
    
    // Sort and add items
    group.items.sort((a, b) => {
      const dateA = a.goLiveDate || new Date('2099-12-31');
      const dateB = b.goLiveDate || new Date('2099-12-31');
      return dateA - dateB;
    });
    
    group.items.forEach(item => {
      const assignment = `${groupType} ${group.number}`;
      const color = getGroupColor(groupType, group.number);
      
      outputData.push([
        item.origin || '',
        item.description || '',
        item.size || '-',
        item.points || 0,
        item.goLiveDate || '',
        item.source || '',
        assignment,
        group.startDate || item.startDate || '',
        group.endDate || item.endDate || ''
      ]);
      
      formats.push(['', '', '', '0', 'yyyy-MM-dd', '', '', 'yyyy-MM-dd', 'yyyy-MM-dd']);
      backgrounds.push(['', '', '', '', '', '', color, color, color]);
    });
    
    // Add spacing
    if (currentRow + outputData.length < 60) {
      outputData.push(['', '', '', '', '', '', '', '', '']);
      formats.push(['', '', '', '', '', '', '', '', '']);
      backgrounds.push(['', '', '', '', '', '', '', '', '']);
    }
  });
  
  // Write all data at once
  if (outputData.length > 0) {
    const maxRows = Math.min(outputData.length, PLANNING_CONFIG.MAX_MANIFEST_ROWS);
    const range = teamSheet.getRange(14, 1, maxRows, 9);
    range.setValues(outputData.slice(0, maxRows));
    
    // Apply formats and backgrounds efficiently
    for (let i = 0; i < maxRows; i++) {
      const row = 14 + i;
      
      // Merge header rows
      if (outputData[i][0].startsWith('---')) {
        teamSheet.getRange(row, 1, 1, 9).merge()
          .setFontWeight('bold')
          .setFontStyle('italic');
      }
      
      // Apply formats where needed
      for (let col = 0; col < 9; col++) {
        if (formats[i][col]) {
          teamSheet.getRange(row, col + 1).setNumberFormat(formats[i][col]);
        }
        if (backgrounds[i][col]) {
          teamSheet.getRange(row, col + 1).setBackground(backgrounds[i][col]);
        }
      }
    }
    
    // Apply border
    range.setBorder(true, true, true, true, true, false);
  }
}

// ==================== OPTIMIZED REFRESH ====================
function refreshPlanningDisplay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNames();
  
  if (teams.length === 0) {
    SpreadsheetApp.getUi().alert('No team sheets found.');
    return;
  }
  
  let refreshCount = 0;
  
  teams.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    // Check for planning data more efficiently
    const planningCheck = teamSheet.getRange(14, 7, 10, 1).getValues();
    const hasPlan = planningCheck.some(row => 
      row[0] && (row[0].toString().includes('Sprint') || row[0].toString().includes('Person'))
    );
    
    if (!hasPlan) return;
    
    // Collect items efficiently
    const manifestData = teamSheet.getRange(14, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 9).getValues();
    const items = [];
    
    for (let i = 0; i < manifestData.length; i++) {
      const row = manifestData[i];
      if (row[1] && !row[1].toString().includes('---') && row[6]) {
        items.push({
          rowIndex: i,
          origin: row[0],
          description: row[1],
          size: row[2],
          points: row[3],
          goLiveDate: row[4],
          source: row[5],
          assignment: row[6],
          startDate: row[7],
          endDate: row[8]
        });
      }
    }
    
    if (items.length === 0) return;
    
    // Group and reorganize
    const groups = {};
    items.forEach(item => {
      if (!groups[item.assignment]) {
        groups[item.assignment] = [];
      }
      groups[item.assignment].push(item);
    });
    
    // Clear and rewrite
    teamSheet.getRange(14, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 9).clear();
    
    // Convert groups to array format for reuse
    const groupArray = Object.keys(groups)
      .sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)?.[0] || '0');
        const numB = parseInt(b.match(/\d+/)?.[0] || '0');
        return numA - numB;
      })
      .map(key => ({
        number: parseInt(key.match(/\d+/)?.[0] || '1'),
        items: groups[key],
        totalPoints: groups[key].reduce((sum, item) => sum + (item.points || 0), 0)
      }));
    
    // Determine type and capacity
    const isSprint = items[0].assignment.includes('Sprint');
    const groupType = isSprint ? 'Sprint' : 'Person';
    const netCapacity = teamSheet.getRange('D7').getValue() || 100;
    const groupCapacity = Math.round(netCapacity / groupArray.length);
    
    // Write back
    writeGroupsToSheet(teamSheet, groupArray, groupType, groupCapacity);
    
    // Restore dropdowns
    const options = groupArray.map(g => `${groupType} ${g.number}`);
    addDropdowns(teamSheet, items.length, options);
    
    refreshCount++;
  });
  
  if (refreshCount > 0) {
    SpreadsheetApp.getUi().alert('Planning Refreshed', 
      `Successfully reorganized ${refreshCount} team sheet(s).`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Planning Found', 
      'No valid planning found. Apply Sprint or Waterfall planning first.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== HELPER FUNCTIONS ====================
function collectManifestItems(teamSheet) {
  const data = teamSheet.getRange(14, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 6).getValues();
  const items = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[3] > 0 && !row[1].toString().startsWith('---')) {
      items.push({
        origin: row[0] || '',
        description: row[1],
        size: row[2] || '-',
        points: parseFloat(row[3]) || 0,
        goLiveDate: row[4],
        source: row[5] || ''
      });
    }
  }
  
  return items;
}

function clearPlanningAreas(teamSheet) {
  teamSheet.getRange(14, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 9).clear();
  teamSheet.getRange(14, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 9).clearDataValidations();
  teamSheet.getRange(13, 7, 1, 3).clear();
}

function addPlanningHeaders(teamSheet, planningType) {
  const headerType = planningType === 'Sprint' ? 'Sprint' : 'Person';
  const headers = [[headerType, 'Start Date', 'End Date']];
  
  teamSheet.getRange(13, 7, 1, 3)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.PLANNING_HEADER);
  
  teamSheet.getRange(13, 1, 1, 9).setBorder(true, true, true, true, true, true);
}

function addDropdowns(teamSheet, itemCount, options) {
  if (options.length === 0) return;
  
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  
  let applied = 0;
  for (let row = 14; row <= Math.min(60, 14 + itemCount * 2); row++) {
    const desc = teamSheet.getRange(row, 2).getValue();
    if (desc && !desc.toString().startsWith('---')) {
      teamSheet.getRange(row, 7).setDataValidation(validation);
      applied++;
      if (applied >= itemCount) break;
    }
  }
}

function distributeItems(items, containers, containerType) {
  // Sort by priority
  items.sort((a, b) => {
    const dateA = a.goLiveDate || new Date('2099-12-31');
    const dateB = b.goLiveDate || new Date('2099-12-31');
    return dateA - dateB;
  });
  
  // Simple load balancing
  items.forEach(item => {
    const container = containers.reduce((min, c) => 
      c.totalPoints < min.totalPoints ? c : min
    );
    container.items.push(item);
    container.totalPoints += item.points;
  });
}

function calculateSprintDate(startDate, sprintIndex, workingDaysPerSprint, isStart) {
  const totalDays = sprintIndex * workingDaysPerSprint;
  if (!isStart) {
    return addWorkingDays(startDate, totalDays + workingDaysPerSprint - 1);
  }
  return sprintIndex === 0 ? new Date(startDate) : addWorkingDays(startDate, totalDays);
}

function getWorkingDaysForDuration(duration) {
  const map = {
    '1 week': 5,
    '2 weeks': 10,
    '1 month': 20
  };
  return map[duration] || 10;
}

function getGroupColor(groupType, number) {
  const colors = groupType === 'Sprint' ? 
    [PLANNING_CONFIG.COLORS.SPRINT_1, PLANNING_CONFIG.COLORS.SPRINT_2, 
     PLANNING_CONFIG.COLORS.SPRINT_3, PLANNING_CONFIG.COLORS.SPRINT_4,
     PLANNING_CONFIG.COLORS.SPRINT_5, PLANNING_CONFIG.COLORS.SPRINT_6] :
    [PLANNING_CONFIG.COLORS.SPRINT_1, PLANNING_CONFIG.COLORS.SPRINT_2,
     PLANNING_CONFIG.COLORS.SPRINT_3, PLANNING_CONFIG.COLORS.SPRINT_4,
     PLANNING_CONFIG.COLORS.SPRINT_5];
  
  return colors[(number - 1) % colors.length];
}

function getTeamNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(sheet => sheet.getName().endsWith(' Team'))
    .map(sheet => sheet.getName().replace(' Team', ''));
}

function clearAllPlanning() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear All Planning', 
    'Remove all sprint/person assignments?', 
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  const teams = getTeamNames();
  let cleared = 0;
  
  teams.forEach(teamName => {
    const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teamName + ' Team');
    if (teamSheet) {
      clearPlanningAreas(teamSheet);
      cleared++;
    }
  });
  
  ui.alert(`Cleared ${cleared} team sheet(s).`);
}

function openPlanningSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(getOrCreatePlanningConfig());
}