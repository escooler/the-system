/**
 * Planning Tools - Version 9.0
 * Added Assignee field to sprint planning for Jira export preparation
 */

// ==================== CONSTANTS ====================
const PLANNING_CONFIG = {
  CONFIG_SHEET_NAME: 'Planning Config',
  DEFAULT_SPRINT_DURATION: '2 weeks',
  SPRINT_DURATIONS: ['1 week', '2 weeks', '1 month'],
  DEFAULT_FIRST_SPRINT: 1,
  WORKING_DAYS_PER_WEEK: 5,
  MANIFEST_START_ROW: 16,
  MAX_MANIFEST_ROWS: 47,
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

// ==================== DATE FUNCTIONS ====================
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
  
  const fullWeeks = Math.floor(workingDays / 5);
  totalDays += fullWeeks * 7;
  
  let remainingDays = workingDays % 5;
  const startDay = result.getDay();
  
  if (startDay + remainingDays > 5) {
    totalDays += (startDay === 0) ? 1 : 2;
  }
  
  totalDays += remainingDays;
  result.setDate(result.getDate() + totalDays);
  
  while (result.getDay() === 0 || result.getDay() === 6) {
    result.setDate(result.getDate() + 1);
  }
  
  return result;
}

// ==================== INSTALLATION ====================
function installPlanningTools() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
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
    'Version 9.0 - Sprint planning now includes assignee field',
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
  sheet.setColumnWidths(1, 3, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  
  const values = [
    ['PLANNING CONFIGURATION', '', ''],
    ['', '', ''],
    ['Planning Method:', 'Sprint', ''],
    ['Sprint Duration:', PLANNING_CONFIG.DEFAULT_SPRINT_DURATION, ''],
    ['First Sprint Number:', PLANNING_CONFIG.DEFAULT_FIRST_SPRINT, ''],
    ['Start Date:', getNextMonday(), '(Adjusted to Monday)']
  ];
  
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  
  sheet.getRange('A1:C1').merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.HEADER)
    .setFontColor('#FFFFFF');
  
  const methodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sprint', 'Waterfall'], true)
    .build();
  sheet.getRange('B3').setDataValidation(methodValidation);
  
  const durationValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLANNING_CONFIG.SPRINT_DURATIONS, true)
    .build();
  sheet.getRange('B4').setDataValidation(durationValidation);
  
  sheet.getRange('B3:B6').setBackground(PLANNING_CONFIG.COLORS.CONFIG_BG);
  sheet.getRange('B6').setNumberFormat('yyyy-MM-dd');
  sheet.getRange('C6').setFontStyle('italic').setFontSize(9);
  sheet.getRange(3, 1, 4, 2).setBorder(true, true, true, true, true, true);
}

function readConfig(configSheet) {
  const values = configSheet.getRange('B3:B6').getValues();
  return {
    method: values[0][0],
    duration: values[1][0],
    firstSprint: parseInt(values[2][0]) || 1,
    startDate: getNextMonday(new Date(values[3][0]))
  };
}

// ==================== TEAM HELPER FUNCTIONS ====================
function getTeamSizeFromSheet(teamSheet) {
  return parseInt(teamSheet.getRange('B4').getValue()) || 5;
}

function getTeamMemberNames(teamSheet) {
  const teamSize = getTeamSizeFromSheet(teamSheet);
  const teamMembers = [];
  const memberData = teamSheet.getRange('G5:G14').getValues();
  
  for (let i = 0; i < teamSize; i++) {
    const customName = memberData[i] ? memberData[i][0] : '';
    if (customName && customName.toString().trim() !== '' && 
        !customName.toString().trim().startsWith('Team Member')) {
      teamMembers.push(customName.toString().trim());
    } else {
      teamMembers.push(`Team Member ${i + 1}`);
    }
  }
  
  return teamMembers;
}

function getStakeholder(origin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // First check if it's a workstream - get actual owner name from workstream sheet
  const wsSheet = ss.getSheetByName(origin);
  if (wsSheet) {
    const ownerName = wsSheet.getRange('G2').getValue();
    if (ownerName && ownerName.toString().trim() !== '') {
      return ownerName.toString().trim();
    }
  }
  
  // Fallback to role name if no owner specified
  const stakeholderMap = {
    'SoMe': 'SoMe Owner',
    'PUA': 'PUA Owner', 
    'ASO': 'ASO Owner',
    'Portal': 'Portal Owner',
    'Creative': 'Creative Director',
    'Content': 'Content Lead',
    'Performance': 'Performance Lead'
  };
  
  return stakeholderMap[origin] || 'PMM';
}

function getTeamNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sheet => sheet.getName().endsWith(' Team'))
    .map(sheet => sheet.getName().replace(' Team', ''));
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
    
    clearPlanningAreas(teamSheet);
    addPlanningHeaders(teamSheet, planningType);
    
    if (planningType === 'Sprint') {
      const netCapacity = teamSheet.getRange('D7').getValue() || 0;
      applySprints(teamSheet, manifestItems, netCapacity, config);
    } else {
      const teamSize = getTeamSizeFromSheet(teamSheet);
      const netCapacity = teamSheet.getRange('D7').getValue() || 0;
      applyWaterfall(teamSheet, manifestItems, teamSize, netCapacity, config.startDate);
    }
    
    successCount++;
  });
  
  if (successCount > 0) {
    const message = planningType === 'Sprint' ? 
      `Applied sprint planning to ${successCount} team(s).\n\nUse assignee dropdowns to assign work to team members.\nLeave blank for no assignee (useful for Jira export).` :
      `Applied waterfall planning to ${successCount} team(s).\n\nTeam members are pre-assigned based on capacity.\nAdjust assignments using dropdowns as needed.`;
    
    SpreadsheetApp.getUi().alert(`${planningType} Planning Applied`, message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function applySprints(teamSheet, items, netCapacity, config) {
  const workingDaysPerSprint = getWorkingDaysForDuration(config.duration);
  const sprintCapacity = Math.round(netCapacity * (workingDaysPerSprint / 20));
  const totalPoints = items.reduce((sum, item) => sum + item.points, 0);
  const sprintsNeeded = Math.max(2, Math.ceil(totalPoints / sprintCapacity));
  
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
  
  items.forEach(item => {
    item.stakeholder = getStakeholder(item.origin);
    item.assigneeName = ''; // Start with no assignee for sprint planning
  });
  
  distributeItems(items, sprints, 'sprint');
  writeGroupsToSheet(teamSheet, sprints, 'Sprint', sprintCapacity);
  
  // Add both sprint and assignee dropdowns
  const teamMembers = getTeamMemberNames(teamSheet);
  addSprintAndAssigneeDropdowns(teamSheet, items.length, sprints.map(s => `Sprint ${s.number}`), teamMembers);
}

function applyWaterfall(teamSheet, items, teamSize, netCapacity, startDate) {
  const capacityPerPerson = Math.round(netCapacity / teamSize);
  const teamMemberNames = getTeamMemberNames(teamSheet);
  
  const teamMembers = [];
  for (let i = 0; i < teamSize; i++) {
    teamMembers.push({
      number: i + 1,
      name: teamMemberNames[i],
      items: [],
      totalPoints: 0,
      capacity: capacityPerPerson,
      currentDate: new Date(startDate)
    });
  }
  
  items.forEach(item => {
    const person = teamMembers.reduce((min, p) => 
      p.totalPoints < min.totalPoints ? p : min
    );
    
    const workingDays = Math.max(1, Math.round(item.points));
    item.startDate = new Date(person.currentDate);
    
    while (item.startDate.getDay() === 0 || item.startDate.getDay() === 6) {
      item.startDate.setDate(item.startDate.getDate() + 1);
    }
    
    item.endDate = addWorkingDays(item.startDate, workingDays - 1);
    item.assignedPerson = person.number;
    item.assigneeName = person.name;
    item.stakeholder = getStakeholder(item.origin);
    
    person.items.push(item);
    person.totalPoints += item.points;
    person.currentDate = addWorkingDays(item.endDate, 1);
  });
  
  writeGroupsToSheet(teamSheet, teamMembers, 'Person', capacityPerPerson);
  addAssigneeDropdowns(teamSheet, items.length, teamMemberNames);
}

// ==================== HEADER AND WRITING FUNCTIONS ====================
function addPlanningHeaders(teamSheet, planningType) {
  // Sprint planning has Sprint + Assignee columns
  // Waterfall planning has only Assignee column (no redundant "Person")
  const headers = planningType === 'Sprint' ?
    [['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source', 'Sprint', 'Assignee', 'Stakeholder', 'Start Date', 'End Date']] :
    [['Origin', 'Description', 'T-Shirt Size', 'Points', 'Go Live Date', 'Source', 'Assignee', 'Stakeholder', 'Start Date', 'End Date']];
  
  const numColumns = planningType === 'Sprint' ? 11 : 10;
  
  teamSheet.getRange(15, 1, 1, numColumns)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground(PLANNING_CONFIG.COLORS.PLANNING_HEADER)
    .setBorder(true, true, true, true, true, true);
}

function writeGroupsToSheet(teamSheet, groups, groupType, capacity) {
  const teamName = teamSheet.getName().replace(' Team', '');
  const outputData = [];
  const formats = [];
  const backgrounds = [];
  const isSprint = groupType === 'Sprint';
  const numColumns = isSprint ? 11 : 10;
  
  groups.forEach(group => {
    if (group.items.length === 0) return;
    
    const utilization = Math.round((group.totalPoints / capacity) * 100);
    const icon = utilization > 100 ? '🔥' : '✅';
    const headerText = isSprint ?
      `--- ${teamName.toUpperCase()} SPRINT ${group.number} (${group.totalPoints}/${capacity} pts - ${utilization}% ${icon}) ---` :
      `--- ${group.name} (${group.totalPoints}/${capacity} pts - ${utilization}% ${icon}) ---`;
    
    // Create empty row for header
    const headerRow = new Array(numColumns).fill('');
    headerRow[0] = headerText;
    outputData.push(headerRow);
    formats.push(new Array(numColumns).fill(''));
    backgrounds.push([PLANNING_CONFIG.COLORS.SEPARATOR, ...new Array(numColumns - 1).fill('')]);
    
    group.items.sort((a, b) => {
      const dateA = a.goLiveDate || new Date('2099-12-31');
      const dateB = b.goLiveDate || new Date('2099-12-31');
      return dateA - dateB;
    });
    
    group.items.forEach(item => {
      const stakeholder = item.stakeholder || getStakeholder(item.origin);
      const color = getGroupColor(groupType, group.number);
      
      let itemRow;
      let formatRow;
      let bgRow;
      
      if (isSprint) {
        // Sprint planning: has Sprint column + Assignee column
        itemRow = [
          item.origin || '',
          item.description || '',
          item.size || '-',
          item.points || 0,
          item.goLiveDate || '',
          item.source || '',
          `Sprint ${group.number}`,
          item.assigneeName || '',
          stakeholder,
          group.startDate || item.startDate || '',
          group.endDate || item.endDate || ''
        ];
        formatRow = ['', '', '', '0', 'yyyy-MM-dd', '', '', '', '', 'yyyy-MM-dd', 'yyyy-MM-dd'];
        bgRow = ['', '', '', '', '', '', color, '', '', color, color];
      } else {
        // Waterfall planning: only Assignee column (no redundant Person column)
        itemRow = [
          item.origin || '',
          item.description || '',
          item.size || '-',
          item.points || 0,
          item.goLiveDate || '',
          item.source || '',
          item.assigneeName || '',
          stakeholder,
          item.startDate || '',
          item.endDate || ''
        ];
        formatRow = ['', '', '', '0', 'yyyy-MM-dd', '', '', '', 'yyyy-MM-dd', 'yyyy-MM-dd'];
        bgRow = ['', '', '', '', '', '', color, '', color, color];
      }
      
      outputData.push(itemRow);
      formats.push(formatRow);
      backgrounds.push(bgRow);
    });
    
    if (outputData.length < PLANNING_CONFIG.MAX_MANIFEST_ROWS - 2) {
      outputData.push(new Array(numColumns).fill(''));
      formats.push(new Array(numColumns).fill(''));
      backgrounds.push(new Array(numColumns).fill(''));
    }
  });
  
  if (outputData.length > 0) {
    const maxRows = Math.min(outputData.length, PLANNING_CONFIG.MAX_MANIFEST_ROWS);
    const range = teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 1, maxRows, numColumns);
    range.setValues(outputData.slice(0, maxRows));
    
    for (let i = 0; i < maxRows; i++) {
      const row = PLANNING_CONFIG.MANIFEST_START_ROW + i;
      
      if (outputData[i][0].startsWith('---')) {
        teamSheet.getRange(row, 1, 1, numColumns).merge()
          .setFontWeight('bold')
          .setFontStyle('italic');
      }
      
      for (let col = 0; col < numColumns; col++) {
        if (formats[i][col]) {
          teamSheet.getRange(row, col + 1).setNumberFormat(formats[i][col]);
        }
        if (backgrounds[i][col]) {
          teamSheet.getRange(row, col + 1).setBackground(backgrounds[i][col]);
        }
      }
    }
    
    range.setBorder(true, true, true, true, true, false);
  }
}

function clearPlanningAreas(teamSheet) {
  // Clear both 10 and 11 column ranges to handle both sprint and waterfall layouts
  teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 11).clear();
  teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 11).clearDataValidations();
}

// ==================== DROPDOWN FUNCTIONS ====================
function addSprintAndAssigneeDropdowns(teamSheet, itemCount, sprintOptions, assigneeOptions) {
  if (sprintOptions.length === 0 || assigneeOptions.length === 0) return;
  
  const sprintValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(sprintOptions, true)
    .setAllowInvalid(false)
    .build();
  
  const assigneeOptionsWithNone = ['None', ...assigneeOptions];
  const assigneeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(assigneeOptionsWithNone, true)
    .setAllowInvalid(false)
    .build();
  
  let applied = 0;
  const startRow = PLANNING_CONFIG.MANIFEST_START_ROW;
  const endRow = Math.min(startRow + PLANNING_CONFIG.MAX_MANIFEST_ROWS, startRow + itemCount * 2);
  
  for (let row = startRow; row <= endRow; row++) {
    const desc = teamSheet.getRange(row, 2).getValue();
    if (desc && !desc.toString().startsWith('---')) {
      teamSheet.getRange(row, 7).setDataValidation(sprintValidation); // Sprint column
      teamSheet.getRange(row, 8).setDataValidation(assigneeValidation); // Assignee column
      applied++;
      if (applied >= itemCount) break;
    }
  }
}

function addAssigneeDropdowns(teamSheet, itemCount, assigneeOptions) {
  if (assigneeOptions.length === 0) return;
  
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(assigneeOptions, true)
    .setAllowInvalid(false)
    .build();
  
  let applied = 0;
  const startRow = PLANNING_CONFIG.MANIFEST_START_ROW;
  const endRow = Math.min(startRow + PLANNING_CONFIG.MAX_MANIFEST_ROWS, startRow + itemCount * 2);
  
  for (let row = startRow; row <= endRow; row++) {
    const desc = teamSheet.getRange(row, 2).getValue();
    if (desc && !desc.toString().startsWith('---')) {
      teamSheet.getRange(row, 7).setDataValidation(validation); // Assignee is now column 7 for waterfall
      applied++;
      if (applied >= itemCount) break;
    }
  }
}

// ==================== HELPER FUNCTIONS ====================
function collectManifestItems(teamSheet) {
  const data = teamSheet.getRange(16, 1, 47, 6).getValues();
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

function distributeItems(items, containers, containerType) {
  items.sort((a, b) => {
    const dateA = a.goLiveDate || new Date('2099-12-31');
    const dateB = b.goLiveDate || new Date('2099-12-31');
    return dateA - dateB;
  });
  
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

// ==================== REFRESH AND CLEANUP FUNCTIONS ====================
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
    
    // Check for planning data - could be 10 or 11 columns
    const planningCheck = teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 7, 10, 1).getValues();
    const hasPlan = planningCheck.some(row => row[0] && row[0].toString().trim() !== '');
    
    if (!hasPlan) return;
    
    // Try to detect if it's sprint or waterfall by checking column 7
    const firstDataRow = teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 7, 1, 1).getValue();
    const isSprint = firstDataRow && firstDataRow.toString().includes('Sprint');
    const numColumns = isSprint ? 11 : 10;
    
    const manifestData = teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, numColumns).getValues();
    const items = [];
    
    for (let i = 0; i < manifestData.length; i++) {
      const row = manifestData[i];
      if (row[1] && !row[1].toString().includes('---')) {
        if (isSprint && row[6]) {
          // Sprint planning: has Sprint + Assignee
          items.push({
            rowIndex: i,
            origin: row[0],
            description: row[1],
            size: row[2],
            points: row[3],
            goLiveDate: row[4],
            source: row[5],
            assignment: row[6],
            assigneeName: row[7] === 'None' ? '' : row[7],
            stakeholder: row[8],
            startDate: row[9],
            endDate: row[10]
          });
        } else if (!isSprint && row[6]) {
          // Waterfall planning: only Assignee (no Person column)
          items.push({
            rowIndex: i,
            origin: row[0],
            description: row[1],
            size: row[2],
            points: row[3],
            goLiveDate: row[4],
            source: row[5],
            assignment: row[6], // This is the assignee name for grouping
            assigneeName: row[6],
            stakeholder: row[7],
            startDate: row[8],
            endDate: row[9]
          });
        }
      }
    }
    
    if (items.length === 0) return;
    
    const groups = {};
    items.forEach(item => {
      if (!groups[item.assignment]) {
        groups[item.assignment] = [];
      }
      groups[item.assignment].push(item);
    });
    
    teamSheet.getRange(PLANNING_CONFIG.MANIFEST_START_ROW, 1, PLANNING_CONFIG.MAX_MANIFEST_ROWS, 11).clear();
    
    const groupArray = Object.keys(groups)
      .sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)?.[0] || '0');
        const numB = parseInt(b.match(/\d+/)?.[0] || '0');
        return numA - numB;
      })
      .map(key => ({
        number: parseInt(key.match(/\d+/)?.[0] || '1'),
        name: key,
        items: groups[key],
        totalPoints: groups[key].reduce((sum, item) => sum + (item.points || 0), 0)
      }));
    
    const groupType = isSprint ? 'Sprint' : 'Person';
    const netCapacity = teamSheet.getRange('D7').getValue() || 100;
    const groupCapacity = Math.round(netCapacity / groupArray.length);
    
    writeGroupsToSheet(teamSheet, groupArray, groupType, groupCapacity);
    
    const teamMembers = getTeamMemberNames(teamSheet);
    if (isSprint) {
      const sprintOptions = groupArray.map(g => g.name);
      addSprintAndAssigneeDropdowns(teamSheet, items.length, sprintOptions, teamMembers);
    } else {
      addAssigneeDropdowns(teamSheet, items.length, teamMembers);
    }
    
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
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getOrCreatePlanningConfig());
}