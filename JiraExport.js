/**
 * Jira Export System - Version 1.0
 * Standalone module for exporting Points System planning to Jira
 */

const JIRA_VERSION = "1.0";

const JIRA_CONFIG = {
  CONFIG_SHEET: 'Jira Config',
  TSHIRT_TO_POINTS: {
    'XS': 1,
    'S': 3,
    'M': 5,
    'L': 13,
    'XL': 21
  }
};

// ==================== INSTALLATION ====================
function installJiraExport() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createJiraExportMenu') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('createJiraExportMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  createJiraExportMenu();
  
  SpreadsheetApp.getUi().alert(
    'Jira Export Installed! 🎉',
    `Version ${JIRA_VERSION}\n\nThe Jira Export menu is now available.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createJiraExportMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Jira Export')
    .addItem('📤 Export to Jira', 'showExportDialog')
    .addSeparator()
    .addItem('📋 View Last Export Log', 'viewExportLog')
    .addSeparator()
    .addItem('⚙️ Jira Configuration', 'openJiraConfig')
    .addItem('🔍 Find Story Points Field', 'findStoryPointsField')
    .addItem('ℹ️ About Jira Export', 'showAboutJiraExport')
    .addToUi();
}

// ==================== CONFIGURATION ====================
function getOrCreateJiraConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(JIRA_CONFIG.CONFIG_SHEET);
  
  if (!configSheet) {
    configSheet = ss.insertSheet(JIRA_CONFIG.CONFIG_SHEET);
    setupJiraConfigSheet(configSheet);
  }
  
  return configSheet;
}

function setupJiraConfigSheet(sheet) {
  sheet.clear();
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 400);
  
  const values = [
    ['JIRA CONFIGURATION', '', ''],
    ['', '', ''],
    ['Jira Instance URL:', 'https://yourcompany.atlassian.net', ''],
    ['API Token:', 'YOUR_API_TOKEN_HERE', ''],
    ['Project Key:', 'MKTG', ''],
    ['Default Issue Type:', 'Story', ''],
    ['', '', ''],
    ['CUSTOM FIELD IDs', '', 'Leave blank to skip field'],
    ['Story Points Field ID:', '', 'Optional - e.g. customfield_10016'],
    ['Go Live Date Field ID:', '', 'Optional - e.g. customfield_10050'],
    ['Start Date Field ID:', '', 'Optional - e.g. customfield_10052'],
    ['Due Date Field:', 'duedate', 'Use "duedate" or custom field ID'],
    ['', '', ''],
    ['Jira User Email:', 'your.email@company.com', ''],
    ['', '', ''],
    ['NOTES:', '', ''],
    ['To find custom field IDs:', '', ''],
    ['1. Go to Jira Settings → Issues → Custom fields', '', ''],
    ['2. Click on a field name', '', ''],
    ['3. Look in the URL for the ID number', '', ''],
    ['', '', ''],
    ['Stakeholder field is not currently supported', '', '']
  ];
  
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  
  sheet.getRange('A1:C1').merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');
  
  sheet.getRange('A8:C8').merge()
    .setFontWeight('bold')
    .setBackground('#E8F0FE');
  
  sheet.getRange('B3:B6').setBackground('#FFF9C4');
  sheet.getRange('B9:B12').setBackground('#E8F5E9');
  sheet.getRange('B14').setBackground('#FFF9C4');
  sheet.getRange('C8:C12').setFontStyle('italic').setFontSize(9);
}

function readJiraConfig() {
  const configSheet = getOrCreateJiraConfig();
  
  return {
    url: configSheet.getRange('B3').getValue().toString().trim(),
    apiToken: configSheet.getRange('B4').getValue().toString().trim(),
    projectKey: configSheet.getRange('B5').getValue().toString().trim(),
    issueType: configSheet.getRange('B6').getValue().toString().trim(),
    customFields: {
      storyPoints: configSheet.getRange('B9').getValue().toString().trim(),
      goLiveDate: configSheet.getRange('B10').getValue().toString().trim(),
      startDate: configSheet.getRange('B11').getValue().toString().trim(),
      dueDate: configSheet.getRange('B12').getValue().toString().trim()
    },
    email: configSheet.getRange('B14').getValue().toString().trim()
  };
}

function openJiraConfig() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getOrCreateJiraConfig());
}

// ==================== MAIN EXPORT FLOW ====================
function showExportDialog() {
  const ui = SpreadsheetApp.getUi();
  
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Allocation')) {
    ui.alert('Points System Not Found', 
      'Please set up the Points System first before exporting to Jira.',
      ui.ButtonSet.OK);
    return;
  }
  
  const scopeResult = ui.alert(
    'Export Scope Selection',
    'What would you like to export?\n\n' +
    '• YES = Export All Teams\n' +
    '• NO = Select Specific Teams or Workstreams\n' +
    '• CANCEL = Cancel Export',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (scopeResult === ui.Button.CANCEL) return;
  
  let exportScope;
  if (scopeResult === ui.Button.YES) {
    exportScope = { type: 'all' };
  } else {
    exportScope = showScopeSelectionDialog();
    if (!exportScope) return;
  }
  
  performExport(exportScope);
}

function showScopeSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  const teams = getTeamNamesForExport();
  const workstreams = getWorkstreamNamesForExport();
  
  const result = ui.alert(
    'Select Export Scope',
    'Export specific:\n\n' +
    '• YES = Select Teams\n' +
    '• NO = Select Workstreams',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (result === ui.Button.CANCEL) return null;
  
  if (result === ui.Button.YES) {
    const teamList = teams.join(', ');
    const response = ui.prompt(
      'Select Teams',
      `Enter team names separated by commas:\n\nAvailable: ${teamList}\n\nExample: Creative, Performance`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.CANCEL) return null;
    
    const selectedTeams = response.getResponseText()
      .split(',')
      .map(t => t.trim())
      .filter(t => teams.includes(t));
    
    if (selectedTeams.length === 0) {
      ui.alert('No valid teams selected');
      return null;
    }
    
    return { type: 'teams', teams: selectedTeams };
  } else {
    const wsList = workstreams.join(', ');
    const response = ui.prompt(
      'Select Workstreams',
      `Enter workstream names separated by commas:\n\nAvailable: ${wsList}\n\nExample: SoMe, PUA`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.CANCEL) return null;
    
    const selectedWs = response.getResponseText()
      .split(',')
      .map(w => w.trim())
      .filter(w => workstreams.includes(w));
    
    if (selectedWs.length === 0) {
      ui.alert('No valid workstreams selected');
      return null;
    }
    
    return { type: 'workstreams', workstreams: selectedWs };
  }
}

function performExport(scope) {
  const ui = SpreadsheetApp.getUi();
  
  const config = readJiraConfig();
  if (!config.url || !config.apiToken || !config.projectKey) {
    ui.alert('Configuration Incomplete',
      'Please configure Jira settings first (Jira Export → Jira Configuration)',
      ui.ButtonSet.OK);
    return;
  }
  
  const connectionTest = testJiraConnection(config);
  if (!connectionTest.success) {
    const fallbackChoice = ui.alert(
      'Jira Connection Failed',
      `Cannot connect to Jira API.\n\nError: ${connectionTest.error}\n\n` +
      'Would you like to export to CSV instead?',
      ui.ButtonSet.YES_NO
    );
    
    if (fallbackChoice === ui.Button.YES) {
      exportToCSV(scope, config);
      return;
    } else {
      return;
    }
  }
  
  const exportData = collectExportData(scope);
  
  if (!exportData.items || exportData.items.length === 0) {
    ui.alert('No Data to Export',
      'No work items found for the selected scope. ' +
      'Make sure teams have planning applied and work assigned.',
      ui.ButtonSet.OK);
    return;
  }
  
  const confirmMsg = buildConfirmationMessage(exportData, scope);
  const confirm = ui.alert('Ready to Export', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  const result = executeJiraExport(exportData, config);
  
  showExportResults(result);
}

// ==================== DATA COLLECTION ====================
function collectExportData(scope) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  const month = allocSheet.getRange('C3').getValue();
  const year = allocSheet.getRange('E3').getValue();
  
  const data = {
    month: month,
    year: year,
    epics: [],
    sprints: [],
    items: [],
    scope: scope
  };
  
  let workstreamsToProcess = [];
  if (scope.type === 'all' || scope.type === 'teams') {
    workstreamsToProcess = getWorkstreamNamesForExport();
  } else {
    workstreamsToProcess = scope.workstreams;
  }
  
  // Create Epics for workstreams
  workstreamsToProcess.forEach(wsName => {
    const wsSheet = ss.getSheetByName(wsName);
    if (!wsSheet) return;
    
    const epic = {
      name: `${month} ${year} - ${wsName}`,
      workstream: wsName,
      owner: getWorkstreamOwnerJiraUsername(wsSheet),
      allocation: wsSheet.getRange('B2').getValue() || 0,
      description: buildEpicDescription(wsSheet, wsName),
      type: 'workstream'
    };
    
    data.epics.push(epic);
  });
  
  let teamsToProcess = [];
  if (scope.type === 'all') {
    teamsToProcess = getTeamNamesForExport();
  } else if (scope.type === 'teams') {
    teamsToProcess = scope.teams;
  } else {
    teamsToProcess = getTeamNamesForExport();
  }
  
  // Check if any team has team-initiated work
  let hasAnyTeamWork = false;
  teamsToProcess.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const teamData = teamSheet.getRange(64, 1, 30, 6).getValues();
    if (teamData.some(row => row[1] && row[3] > 0)) {
      hasAnyTeamWork = true;
    }
  });
  
  // Create single shared Epic for all team-initiated work
  if (hasAnyTeamWork) {
    const teamEpic = {
      name: `${month} ${year} - Team Initiatives`,
      workstream: 'Team Initiatives',
      owner: '',
      allocation: 0,
      description: `Team-initiated work across all teams for ${month} ${year}`,
      type: 'team'
    };
    
    data.epics.push(teamEpic);
  }
  
  teamsToProcess.forEach(teamName => {
    const teamSheet = ss.getSheetByName(teamName + ' Team');
    if (!teamSheet) return;
    
    const teamItems = collectTeamItems(teamSheet, teamName, scope, workstreamsToProcess);
    data.items.push(...teamItems);
    
    const sprints = collectTeamSprints(teamSheet, teamName, teamItems);
    data.sprints.push(...sprints);
  });
  
  return data;
}

function collectTeamItems(teamSheet, teamName, scope, workstreamsToProcess) {
  const items = [];
  
  const headers = teamSheet.getRange(15, 1, 1, 11).getValues()[0];
  const hasPlanning = headers.some(h => h && (h.includes('Sprint') || h.includes('Assignee') || h.includes('Stakeholder')));
  
  if (!hasPlanning) return items;
  
  const isSprint = headers[6] && headers[6].toString().includes('Sprint');
  const numColumns = isSprint ? 11 : 10;
  
  const data = teamSheet.getRange(16, 1, 47, numColumns).getValues();
  
  // Get team member mapping for this team
  const memberMapping = getTeamMemberJiraMapping(teamSheet);
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    if (!row[1] || row[1].toString().startsWith('---')) continue;
    
    const origin = row[0];
    
    if (scope.type === 'workstreams' && !workstreamsToProcess.includes(origin)) {
      continue;
    }
    
    let item;
    if (isSprint) {
      const assigneeName = row[7] === 'None' ? '' : row[7];
      item = {
        team: teamName,
        origin: origin,
        description: row[1],
        size: row[2],
        points: row[3] || 0,
        goLiveDate: row[4],
        source: row[5],
        sprint: row[6],
        assignee: assigneeName ? (memberMapping[assigneeName] || '') : '',
        assigneeName: assigneeName,
        stakeholder: row[8],
        startDate: row[9],
        endDate: row[10],
        storyPoints: convertTShirtToPoints(row[2])
      };
    } else {
      const assigneeName = row[6];
      item = {
        team: teamName,
        origin: origin,
        description: row[1],
        size: row[2],
        points: row[3] || 0,
        goLiveDate: row[4],
        source: row[5],
        sprint: null,
        assignee: assigneeName ? (memberMapping[assigneeName] || '') : '',
        assigneeName: assigneeName,
        stakeholder: row[7],
        startDate: row[8],
        endDate: row[9],
        storyPoints: convertTShirtToPoints(row[2])
      };
    }
    
    items.push(item);
  }
  
  return items;
}

function getTeamMemberJiraMapping(teamSheet) {
  const mapping = {};
  
  // Get team member names (column G) and Jira usernames (column I)
  const memberNames = teamSheet.getRange('G5:G14').getValues();
  const jiraUsernames = teamSheet.getRange('I5:I14').getValues();
  
  for (let i = 0; i < memberNames.length; i++) {
    const name = memberNames[i][0];
    const jiraId = jiraUsernames[i][0];
    
    if (name && jiraId && jiraId.toString().trim() !== '') {
      mapping[name.toString().trim()] = jiraId.toString().trim();
    }
  }
  
  return mapping;
}

function collectTeamSprints(teamSheet, teamName, teamItems) {
  const sprints = [];
  const sprintMap = {};
  
  teamItems.forEach(item => {
    if (!item.sprint) return;
    
    if (!sprintMap[item.sprint]) {
      sprintMap[item.sprint] = {
        name: `${teamName} - ${item.sprint}`,
        team: teamName,
        sprintNumber: item.sprint,
        items: [],
        startDate: item.startDate,
        endDate: item.endDate
      };
    }
    
    sprintMap[item.sprint].items.push(item);
    
    if (item.startDate && (!sprintMap[item.sprint].startDate || item.startDate < sprintMap[item.sprint].startDate)) {
      sprintMap[item.sprint].startDate = item.startDate;
    }
    if (item.endDate && (!sprintMap[item.sprint].endDate || item.endDate > sprintMap[item.sprint].endDate)) {
      sprintMap[item.sprint].endDate = item.endDate;
    }
  });
  
  Object.values(sprintMap).forEach(sprint => {
    sprints.push(sprint);
  });
  
  return sprints;
}

function buildEpicDescription(wsSheet, wsName) {
  const allocation = wsSheet.getRange('B2').getValue() || 0;
  const spent = wsSheet.getRange('B3').getValue() || 0;
  
  let description = `Workstream: ${wsName}\n`;
  description += `Total Allocation: ${allocation} points\n`;
  description += `Planned Assets: ${spent} points\n\n`;
  
  const pmmPriorities = [];
  for (let row = 21; row <= 35; row++) {
    const priority = wsSheet.getRange(row, 1).getValue();
    const percent = wsSheet.getRange(row, 3).getValue();
    if (priority && percent) {
      pmmPriorities.push(`- ${priority} (${Math.round(percent * 100)}%)`);
    }
  }
  
  if (pmmPriorities.length > 0) {
    description += 'Strategic Priorities:\n' + pmmPriorities.join('\n') + '\n\n';
  }
  
  const wsPriorities = [];
  for (let row = 7; row <= 16; row++) {
    const priority = wsSheet.getRange(row, 1).getValue();
    const percent = wsSheet.getRange(row, 3).getValue();
    if (priority && percent) {
      wsPriorities.push(`- ${priority} (${Math.round(percent * 100)}%)`);
    }
  }
  
  if (wsPriorities.length > 0) {
    description += 'Workstream Priorities:\n' + wsPriorities.join('\n');
  }
  
  return description;
}

// ==================== JIRA API ====================
function testJiraConnection(config) {
  if (!config.url.includes('http')) {
    return { success: false, error: 'Invalid Jira URL' };
  }
  
  if (config.apiToken === 'YOUR_API_TOKEN_HERE' || !config.apiToken) {
    return { success: false, error: 'API Token not configured' };
  }
  
  if (!config.projectKey || config.projectKey.length === 0) {
    return { success: false, error: 'Project Key not configured' };
  }
  
  if (!config.email || config.email === 'your.email@company.com') {
    return { success: false, error: 'Email not configured' };
  }
  
  try {
    const url = `${config.url}/rest/api/3/myself`;
    const headers = getJiraHeaders(config);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return { success: true };
    } else {
      return { 
        success: false, 
        error: `HTTP ${response.getResponseCode()}: ${response.getContentText()}` 
      };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getJiraHeaders(config) {
  const authString = Utilities.base64Encode(`${config.email}:${config.apiToken}`);
  
  return {
    'Authorization': `Basic ${authString}`,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };
}

function executeJiraExport(data, config) {
  const result = {
    success: true,
    epicsCreated: 0,
    sprintsCreated: 0,
    storiesCreated: 0,
    failed: 0,
    errors: [],
    timestamp: new Date(),
    details: {
      epics: [],
      sprints: [],
      stories: []
    }
  };
  
  try {
    const epicMap = {};
    for (const epic of data.epics) {
      try {
        const epicKey = createOrFindEpic(epic, config);
        epicMap[epic.workstream] = epicKey;
        result.epicsCreated++;
        result.details.epics.push({ name: epic.name, key: epicKey });
      } catch (e) {
        result.errors.push(`Epic ${epic.name}: ${e.toString()}`);
        result.failed++;
      }
    }
    
    const sprintMap = {};
    if (data.sprints.length > 0) {
      const boardId = getBoardId(config);
      
      if (boardId) {
        for (const sprint of data.sprints) {
          try {
            const sprintId = createSprint(sprint, boardId, config);
            sprintMap[sprint.name] = sprintId;
            result.sprintsCreated++;
            result.details.sprints.push({ name: sprint.name, id: sprintId });
          } catch (e) {
            result.errors.push(`Sprint ${sprint.name}: ${e.toString()}`);
          }
        }
      } else {
        result.errors.push('Could not find board ID - sprints not created');
      }
    }
    
    for (const item of data.items) {
      try {
        // Check if origin is a workstream or a team
        let epicKey = epicMap[item.origin];
        
        // If no epic found for origin, it's team-initiated work
        // Use the shared Team Initiatives epic
        if (!epicKey) {
          epicKey = epicMap['Team Initiatives'];
        }
        
        const sprintId = item.sprint ? sprintMap[`${item.team} - ${item.sprint}`] : null;
        
        const storyKey = createStory(item, epicKey, sprintId, config);
        result.storiesCreated++;
        result.details.stories.push({ summary: item.description, key: storyKey });
      } catch (e) {
        result.errors.push(`Story "${item.description}": ${e.toString()}`);
        result.failed++;
      }
    }
    
    if (result.storiesCreated === 0 && data.items.length > 0) {
      result.success = false;
    } else if (result.failed > result.storiesCreated) {
      result.success = false;
    }
    
  } catch (e) {
    result.success = false;
    result.errors.push(`Fatal error: ${e.toString()}`);
  }
  
  saveExportLog(result, data);
  
  return result;
}

function createOrFindEpic(epic, config) {
  const searchUrl = `${config.url}/rest/api/3/search`;
  const jql = `project = "${config.projectKey}" AND issuetype = Epic AND summary ~ "${epic.name}"`;
  
  const searchPayload = {
    jql: jql,
    maxResults: 1,
    fields: ['key']
  };
  
  try {
    const searchResponse = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      headers: getJiraHeaders(config),
      payload: JSON.stringify(searchPayload),
      muteHttpExceptions: true
    });
    
    if (searchResponse.getResponseCode() === 200) {
      const searchData = JSON.parse(searchResponse.getContentText());
      if (searchData.issues && searchData.issues.length > 0) {
        return searchData.issues[0].key;
      }
    }
  } catch (e) {
    // Continue to create new epic
  }
  
  const url = `${config.url}/rest/api/3/issue`;
  
  const payload = {
    fields: {
      project: { key: config.projectKey },
      summary: epic.name,
      description: {
        type: 'doc',
        version: 1,
        content: [{
          type: 'paragraph',
          content: [{
            type: 'text',
            text: epic.description
          }]
        }]
      },
      issuetype: { name: 'Epic' }
    }
  };
  
  if (epic.owner) {
    payload.fields.assignee = { accountId: epic.owner };
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: getJiraHeaders(config),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 201) {
    const data = JSON.parse(response.getContentText());
    return data.key;
  } else {
    throw new Error(`Failed to create Epic: ${response.getContentText()}`);
  }
}

function getBoardId(config) {
  const url = `${config.url}/rest/agile/1.0/board?projectKeyOrId=${config.projectKey}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: getJiraHeaders(config),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.values && data.values.length > 0) {
        return data.values[0].id;
      }
    }
  } catch (e) {
    // Return null if can't get board
  }
  
  return null;
}

function createSprint(sprint, boardId, config) {
  const url = `${config.url}/rest/agile/1.0/sprint`;
  
  const payload = {
    name: sprint.name,
    originBoardId: boardId,
    startDate: formatDateForJira(sprint.startDate),
    endDate: formatDateForJira(sprint.endDate)
  };
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: getJiraHeaders(config),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 201) {
    const data = JSON.parse(response.getContentText());
    return data.id;
  } else {
    throw new Error(`Failed to create Sprint: ${response.getContentText()}`);
  }
}

function createStory(item, epicKey, sprintId, config) {
  const url = `${config.url}/rest/api/3/issue`;
  
  const payload = {
    fields: {
      project: { key: config.projectKey },
      summary: item.description,
      issuetype: { name: config.issueType || 'Story' }
    }
  };
  
  if (config.customFields.storyPoints && item.storyPoints > 0) {
    payload.fields[config.customFields.storyPoints] = item.storyPoints;
  }
  
  if (epicKey) {
    payload.fields.parent = { key: epicKey };
  }
  
  // Assignee - now properly mapped to Jira account ID
  if (item.assignee && item.assignee.trim() !== '') {
    payload.fields.assignee = { accountId: item.assignee };
  }
  
  if (config.customFields.goLiveDate && item.goLiveDate) {
    payload.fields[config.customFields.goLiveDate] = formatDateForJira(item.goLiveDate);
  }
  
  if (config.customFields.startDate && item.startDate) {
    payload.fields[config.customFields.startDate] = formatDateForJira(item.startDate);
  }
  
  if (item.endDate) {
    if (config.customFields.dueDate === 'duedate') {
      payload.fields.duedate = formatDateForJira(item.endDate);
    } else if (config.customFields.dueDate) {
      payload.fields[config.customFields.dueDate] = formatDateForJira(item.endDate);
    }
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: getJiraHeaders(config),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 201) {
    const data = JSON.parse(response.getContentText());
    const issueKey = data.key;
    
    if (sprintId) {
      try {
        addIssueToSprint(issueKey, sprintId, config);
      } catch (e) {
        // Non-fatal error
      }
    }
    
    return issueKey;
  } else {
    throw new Error(`Failed to create Story: ${response.getContentText()}`);
  }
}

function addIssueToSprint(issueKey, sprintId, config) {
  const url = `${config.url}/rest/agile/1.0/sprint/${sprintId}/issue`;
  
  const payload = {
    issues: [issueKey]
  };
  
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: getJiraHeaders(config),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function formatDateForJira(date) {
  if (!date) return null;
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return date.toString();
}

// ==================== HELPER FUNCTIONS ====================
function getTeamNamesForExport() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sheet => sheet.getName().endsWith(' Team'))
    .map(sheet => sheet.getName().replace(' Team', ''));
}

function getWorkstreamNamesForExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allocSheet = ss.getSheetByName('Allocation');
  const workstreams = [];
  
  let row = 10;
  let wsName = allocSheet.getRange(row, 1).getValue();
  while (wsName && wsName !== 'TOTAL' && row < 20) {
    workstreams.push(wsName);
    row++;
    wsName = allocSheet.getRange(row, 1).getValue();
  }
  
  return workstreams;
}

function getWorkstreamOwnerJiraUsername(wsSheet) {
  const ownerName = wsSheet.getRange('G2').getValue();
  const jiraUsername = wsSheet.getRange('G4').getValue();
  
  if (jiraUsername && jiraUsername.toString().trim() !== '') {
    return jiraUsername.toString().trim();
  }
  
  if (ownerName && ownerName.toString().trim() !== '') {
    const teams = getTeamNamesForExport();
    for (const teamName of teams) {
      const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teamName + ' Team');
      if (!teamSheet) continue;
      
      const memberNames = teamSheet.getRange('G5:G14').getValues();
      const jiraUsernames = teamSheet.getRange('I5:I14').getValues();
      
      for (let i = 0; i < memberNames.length; i++) {
        if (memberNames[i][0] === ownerName && jiraUsernames[i][0]) {
          return jiraUsernames[i][0].toString().trim();
        }
      }
    }
  }
  
  return '';
}

function convertTShirtToPoints(size) {
  if (!size) return 0;
  return JIRA_CONFIG.TSHIRT_TO_POINTS[size.toString().trim()] || 0;
}

function buildConfirmationMessage(data, scope) {
  let msg = 'Export Summary:\n\n';
  
  if (scope.type === 'all') {
    msg += `Scope: All Teams\n`;
  } else if (scope.type === 'teams') {
    msg += `Scope: Teams (${scope.teams.join(', ')})\n`;
  } else {
    msg += `Scope: Workstreams (${scope.workstreams.join(', ')})\n`;
  }
  
  const workstreamEpics = data.epics.filter(e => e.type === 'workstream').length;
  const hasTeamEpic = data.epics.some(e => e.type === 'team');
  
  msg += `\nWill create:\n`;
  msg += `• ${data.epics.length} Epic(s)`;
  if (hasTeamEpic) {
    msg += ` (${workstreamEpics} workstream, 1 team initiatives)`;
  }
  msg += `\n`;
  msg += `• ${data.sprints.length} Sprint(s)\n`;
  msg += `• ${data.items.length} Story/Stories\n`;
  msg += `\nEstimated time: ~${Math.ceil(data.items.length / 20)}-${Math.ceil(data.items.length / 10)} minutes\n`;
  msg += `\nContinue with export?`;
  
  return msg;
}

// ==================== CSV EXPORT ====================
function exportToCSV(scope, config) {
  const data = collectExportData(scope);
  
  if (!data.items || data.items.length === 0) {
    SpreadsheetApp.getUi().alert('No data to export');
    return;
  }
  
  let csv = 'Epic Name,Summary,Story Points,Assignee,Go Live Date,Start Date,Due Date,Sprint\n';
  
  data.items.forEach(item => {
    const epicName = `${data.month} ${data.year} - ${item.origin}`;
    const row = [
      epicName,
      item.description,
      item.storyPoints,
      item.assignee || '',
      formatDateForCSV(item.goLiveDate),
      formatDateForCSV(item.startDate),
      formatDateForCSV(item.endDate),
      item.sprint || ''
    ];
    
    csv += row.map(field => `"${field}"`).join(',') + '\n';
  });
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const csvSheet = ss.insertSheet(`Jira_Export_${timestamp}`);
  
  const rows = csv.split('\n').map(row => {
    return row.split(',').map(cell => cell.replace(/^"|"$/g, ''));
  });
  
  csvSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  csvSheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#E8F0FE');
  
  SpreadsheetApp.getUi().alert(
    'Exported to CSV Sheet',
    `Created sheet: ${csvSheet.getName()}\n\n` +
    `Total items: ${data.items.length}\n\n` +
    'You can now copy this data and import it to Jira manually.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function formatDateForCSV(date) {
  if (!date) return '';
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return date.toString();
}

// ==================== EXPORT LOG ====================
function saveExportLog(result, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Jira Export Log');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('Jira Export Log');
    logSheet.getRange(1, 1, 1, 7).setValues([
      ['Timestamp', 'Scope', 'Epics', 'Sprints', 'Stories', 'Failed', 'Status']
    ]).setFontWeight('bold').setBackground('#E8F0FE');
  }
  
  const lastRow = logSheet.getLastRow() + 1;
  const scopeDesc = data.scope.type === 'all' ? 'All Teams' :
                    data.scope.type === 'teams' ? data.scope.teams.join(', ') :
                    data.scope.workstreams.join(', ');
  
  const status = result.success ? '✅ Success' : 
                 result.storiesCreated > 0 ? '⚠️ Partial' : 
                 '❌ Failed';
  
  logSheet.getRange(lastRow, 1, 1, 7).setValues([[
    result.timestamp,
    scopeDesc,
    result.epicsCreated,
    result.sprintsCreated,
    result.storiesCreated,
    result.failed,
    status
  ]]);
  
  if (result.errors.length > 0) {
    logSheet.getRange(lastRow, 8).setValue(result.errors.join('\n\n'));
  }
}

function viewExportLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Jira Export Log');
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No export log found. Run an export first.');
    return;
  }
  
  ss.setActiveSheet(logSheet);
}

// ==================== UI FUNCTIONS ====================
function showExportResults(result) {
  const ui = SpreadsheetApp.getUi();
  
  if (result.success || result.storiesCreated > 0) {
    let msg = result.success ? 'Export completed successfully! 🎉\n\n' : 'Export completed with some errors\n\n';
    msg += `Created:\n`;
    msg += `• ${result.epicsCreated} Epic(s)\n`;
    msg += `• ${result.sprintsCreated} Sprint(s)\n`;
    msg += `• ${result.storiesCreated} Story/Stories\n`;
    
    if (result.failed > 0) {
      msg += `\n⚠️ Failed: ${result.failed} item(s)\n\n`;
      
      const fieldErrors = result.errors.filter(e => 
        e.includes('customfield') || e.includes('cannot be set')
      );
      
      if (fieldErrors.length > 0) {
        msg += 'Tip: Field errors detected.\n';
        msg += '• Go to Jira Config and clear any custom field IDs that are causing errors\n';
        msg += '• Or use "Find Story Points Field" to get the correct field ID\n\n';
      }
    }
    
    msg += `\nView details in "Jira Export Log" sheet.`;
    
    ui.alert(result.success ? 'Export Complete' : 'Partial Export', msg, ui.ButtonSet.OK);
  } else {
    let errorMsg = `The export encountered errors:\n\n`;
    
    const firstErrors = result.errors.slice(0, 3);
    errorMsg += firstErrors.join('\n\n');
    
    if (result.errors.length > 3) {
      errorMsg += `\n\n...and ${result.errors.length - 3} more errors`;
    }
    
    errorMsg += `\n\nCheck "Jira Export Log" for full details.`;
    
    ui.alert('Export Failed', errorMsg, ui.ButtonSet.OK);
  }
}

function showAboutJiraExport() {
  const ui = SpreadsheetApp.getUi();
  
  const msg = `Jira Export System\nVersion ${JIRA_VERSION}\n\n` +
    'Features:\n' +
    '• Export workstreams as Epics\n' +
    '• Create Sprints from sprint planning\n' +
    '• Create Stories with full field mapping\n' +
    '• Selective export (teams or workstreams)\n' +
    '• CSV fallback when API unavailable\n' +
    '• Configurable custom fields\n\n' +
    'Field Mappings:\n' +
    '• T-Shirt Size → Story Points (optional)\n' +
    '• Origin → Epic Link\n' +
    '• Assignee → Jira Assignee\n' +
    '• Go Live Date → Custom Field (optional)\n' +
    '• Start/End Dates → Date Fields (optional)\n\n' +
    'Setup Tips:\n' +
    '• Leave custom fields blank to skip them\n' +
    '• Find field IDs in Jira admin settings\n' +
    '• Test with a small export first\n\n' +
    '© 2024 Marketing Team';
  
  ui.alert('About Jira Export', msg, ui.ButtonSet.OK);
}

// ==================== FIELD DISCOVERY HELPER ====================
function findStoryPointsField() {
  const ui = SpreadsheetApp.getUi();
  const config = readJiraConfig();
  
  if (!config.url || !config.apiToken || !config.email) {
    ui.alert('Please configure Jira connection settings first');
    return;
  }
  
  ui.alert('Finding Story Points Field',
    'This will search for the Story Points custom field in your Jira instance.\n\n' +
    'This may take a moment...',
    ui.ButtonSet.OK);
  
  try {
    const url = `${config.url}/rest/api/3/field`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: getJiraHeaders(config),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const fields = JSON.parse(response.getContentText());
      
      const storyPointsFields = fields.filter(f => 
        f.name && (
          f.name.toLowerCase().includes('story points') ||
          f.name.toLowerCase().includes('story point') ||
          f.key === 'customfield_10016'
        )
      );
      
      if (storyPointsFields.length > 0) {
        const fieldInfo = storyPointsFields.map(f => 
          `${f.name}: ${f.id || f.key}`
        ).join('\n');
        
        ui.alert('Story Points Field(s) Found!',
          `Found the following field(s):\n\n${fieldInfo}\n\n` +
          'Copy the ID (customfield_XXXXX) to your Jira Config sheet, cell B9.',
          ui.ButtonSet.OK);
      } else {
        ui.alert('Story Points Field Not Found',
          'Could not find a Story Points field.\n\n' +
          'Options:\n' +
          '• Leave Story Points blank in config (tickets will be created without points)\n' +
          '• Ask your Jira admin to add Story Points field\n' +
          '• Manually find the field ID in Jira settings',
          ui.ButtonSet.OK);
      }
    }
  } catch (e) {
    ui.alert('Error', `Failed to search fields: ${e.toString()}`, ui.ButtonSet.OK);
  }
}