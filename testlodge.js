// Main execution function for test runs
function testlodgeChangelog() {
  console.info("Starting execution of testlodgeChangelog function.");
  const startTime = new Date();

  const data = gatherTestRunAndCaseData();
  const rowsCount = writeTestDataToSheet(data);

  const endTime = new Date();
  const executionTime = (endTime - startTime) / 1000; // Convert milliseconds to seconds
  const minutes = Math.floor(executionTime / 60);
  const seconds = Math.floor(executionTime % 60);
  console.info(`Execution completed in ${minutes} minutes and ${seconds} seconds.`);

  return rowsCount;
}

function getTestlodgeProjectsData() {
  try {
    const projects = fetchAllTestLodgeProjects();
    const projectsData = [];
    let totalSuites = 0;
    let totalSteps = 0;

    const squadMap = new Map();
    PHM.Spreadsheet.getRangeValues(TESTLODGE_PROJECT_IDS_SQUADS).forEach(row => {
      const [projectId, squadName] = row;
      squadMap.set(projectId.toString(), squadByTestLodgeProject(projectId.toString()));
    });

    for (const project of projects) {
      const suites = fetchTestSuites(project.id);
      const squad = squadMap.get(project.id.toString()) || '';

      totalSuites += suites.length;
      let projectTotalSteps = 0;
      let coreSteps = 0;
      let automatedSteps = 0;
      let coreAndAutomatedSteps = 0;
      for (const suite of suites) {
        const steps = fetchTestSteps(project.id, suite.id);

        projectTotalSteps += steps.length;

        for (const step of steps) {
          const stepTitle = String(step.title || '').toLowerCase();
          if (stepTitle.includes('[core]')) {
            coreSteps++;
          }
          if (stepTitle.includes('[automatizado]')) {
            automatedSteps++;
          }
          if (stepTitle.includes('[core]') && stepTitle.includes('[automatizado]')) {
            coreAndAutomatedSteps++;
          }
        }
      }

      totalSteps += projectTotalSteps;

      projectsData.push({
        id: project.id,
        name: project.name,
        created_at: PHM.DateUtils.formatDate(project.created_at),
        squad: squad,
        total_test_cases: projectTotalSteps,
        core_test_cases: coreSteps,
        automated_test_cases: automatedSteps,
        core_and_automated_test_cases: coreAndAutomatedSteps
      });
    }

    return projectsData;
  } catch (error) {
    throw new Error(`Failed to fetch projects data: ${error.message}`);
  }
}

function writeTestlodgeProjectsDataOnSheet(projectsData) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TESTLODGE_SHEET_NAME);
  const startRow = 2;

  sheet.getRange('A2:G').clearContent();

  const data = projectsData.map(project => [
    project.id,
    project.name,
    project.created_at,
    project.squad,
    project.total_test_cases,
    project.core_test_cases,
    project.automated_test_cases,
    project.core_and_automated_test_cases
  ]);

  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
}

function fetchTestSuites(projectId) {
  let allSuites = [];
  let page = 1;
  let hasMorePages = true;

  while (hasMorePages) {
    try {
      const response = makeTestLodgeApiRequest(`/projects/${projectId}/suites.json`, { page: page, per_page: 100 });

      if (!response || !response.pagination || response.pagination.total_entries === 0) {
        return allSuites;
      }

      const newSuites = response.suites || [];
      allSuites = allSuites.concat(newSuites);

      hasMorePages = response.pagination.next_page !== null;
      page++;
    } catch (error) {
      throw error;
    }
  }

  return allSuites;
}

function fetchTestSteps(projectId, suiteId) {
  let allSteps = [];
  let page = 1;
  let hasMorePages = true;

  while (hasMorePages) {
    try {
      const response = makeTestLodgeApiRequest(`/projects/${projectId}/suites/${suiteId}/steps.json`, { page: page, per_page: 100 });

      if (!response || !response.pagination || response.pagination.total_entries === 0) {
        return allSteps;
      }

      const newSteps = response.steps || [];
      allSteps = allSteps.concat(newSteps);

      hasMorePages = response.pagination.next_page !== null;
      page++;
    } catch (error) {
      throw error;
    }
  }

  return allSteps;
}

function fetchAllTestLodgeProjects() {
  let allProjects = [];
  let page = 1;
  let hasMorePages = true;

  while (hasMorePages) {
    const response = makeTestLodgeApiRequest('/projects.json', { page: page, per_page: 100 });

    if (!response || !response.pagination || response.pagination.total_entries === 0) {
      return allProjects;
    }

    allProjects = allProjects.concat(response.projects || []);
    hasMorePages = response.pagination.next_page !== null;
    page++;
  }

  return allProjects;
}

function testlodgeData() {
  console.info("Starting execution of testlodgeData function.");
  const startTime = new Date();
  const projectsData = getTestlodgeProjectsData();
  writeTestlodgeProjectsDataOnSheet(projectsData);
  const endTime = new Date();
  const executionTime = (endTime - startTime) / 1000; // Convert milliseconds to seconds
  const minutes = Math.floor(executionTime / 60);
  const seconds = Math.floor(executionTime % 60);
  console.info(`Execution completed in ${minutes} minutes and ${seconds} seconds.`);
}

// Function to fetch all test runs for a project
function fetchTestRunsForProject(projectId) {
  let allRuns = [];
  let page = 1;
  let hasMorePages = true;

  while (hasMorePages) {
    const response = makeTestLodgeApiRequest(`/projects/${projectId}/runs.json`, { page: page, per_page: 100 });

    if (!response || !response.pagination || response.pagination.total_entries === 0) {
      return allRuns;
    }

    allRuns = allRuns.concat(response.runs || []);
    hasMorePages = response.pagination.next_page !== null;
    page++;
  }

  return allRuns;
}

// Function to fetch all users and create a dictionary mapping user IDs to their name and email
function fetchUserDictionary() {
  const users = fetchAllUsers();
  const userDict = {};

  for (const user of users) {
    userDict[user.id] = PHM.Utilities.updateUsername(user.email,USER_DICTIONARY) ||`${user.firstname} ${user.lastname}*`;
  }
  return userDict;
}

// Main function to gather all test run and test case data
function gatherTestRunAndCaseData() {
  const projects = fetchAllTestLodgeProjects();
  const data = [];

  const configSheet = PHM.Spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  const datespan = configSheet.getRange(CHGLOG_DATESPAN_RANGE).getValue();

  const today = new Date();
  const startDate = new Date(today.getTime() - (datespan * 24 * 60 * 60 * 1000));

  const userDict = fetchUserDictionary();


  // Grab the squad data from the range
  const squadData = PHM.Spreadsheet.getRangeValues(TESTLODGE_PROJECT_IDS_SQUADS);
  const squadMap = new Map();

  // Build a map of projectId -> squadName
  squadData.forEach(row => {
    const [projectId, squadName] = row;
    squadMap.set(projectId.toString(), squadName);
  });

  for (const project of projects) {
    const suites = fetchTestSuites(project.id);
    const squad = squadMap.get(project.id.toString()) || '';


    const lastUpdaterDict = {};

    for (const suite of suites) {
      const steps = fetchTestSteps(project.id, suite.id);

      for (const step of steps) {
        const createdDate = new Date(step.created_at);
        const user = userDict[step.last_saved_by_id];

        if (createdDate >= startDate) {
          const createdRow = [
            'testlodge-' + step.id,
            PHM.DateUtils.formatDate(createdDate),
            'testlodge',
            `${project.id} | ${project.name}`,
            squad,
            user,
            `Caso de teste ${step.step_number} criado`,
            `Caso de teste: ${step.step_number} | ${step.title}`
          ];
          data.push(createdRow);
        }

        if (step.updated_at) {
          const updatedDate = new Date(step.updated_at);

          if (updatedDate >= startDate) {
            const updatedRow = [
              'testlodge-' + step.id,
              PHM.DateUtils.formatDate(updatedDate),
              'testlodge',
              `${project.id} | ${project.name}`,
              squad,
              user,
              `Caso de teste ${step.step_number} atualizado`,
              `Caso de teste: ${step.step_number} | ${step.title}`
            ];
            data.push(updatedRow);
          }
        }
        //last updater
        if (step.last_saved_by_id && user) {
          lastUpdaterDict[project.id] = user;
        }
      }
    }

    const testRuns = fetchTestRunsForProject(project.id);

    for (const run of testRuns) {
      const runDate = new Date(run.created_at);

      if (runDate >= startDate) {
        const lastUpdater = lastUpdaterDict[project.id] || '';

        const row = [
          'testlodge-' + run.id,
          PHM.DateUtils.formatDate(runDate),
          `testlodge`,
          `${project.id} | ${project.name}`,
          squad,
          lastUpdater,
          `Regressão realizada: ${project.name}`,
          `${project.name} | ${run.name} | Passaram: ${run.passed_number}, Incompletos: ${run.incomplete_number}, Ignorados: ${run.skipped_number}, Falharam: ${run.failed_number} | ${PHM.DateUtils.formatDate(runDate, true)}`
        ];
        data.push(row);
      }
    }
  }

  return data;
}

// Function to write test run data to the spreadsheet
function writeTestDataToSheet(data) {
  const sheet = PHM.Spreadsheet.getSheetByName(TESTLODGE_CHG_SHEET_NAME);
  const range = sheet.getRange(2, 1, data.length, data[0].length);

  const headerRow = [
    "Test Run ID",
    "Test Run Date",
    "Tool",
    "Project ID | Name",
    "Squad",
    "Author",
    "Action",
    "Detail"
  ];

  sheet.getRange(2, 1, sheet.getLastRow(), data[0].length).clearContent();

  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  range.setValues(data);
  range.sort([{ column: 3, ascending: false }])

  return data.length;
}

// Function to fetch all users from the TestLodge API
function fetchAllUsers() {
  const users = [];
  let page = 1;
  let hasMorePages = true;

  while (hasMorePages) {
    const data = makeTestLodgeApiRequest('users.json', { page: page });

    users.push(...data.users);

    hasMorePages = data.pagination.next_page !== null;
    page++;
  }
  return users;
}


function makeTestLodgeApiRequest(endpoint, params = {}) {
  let url = `${TESTLODGE_API_URL}${endpoint}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${TESTLODGE_API_EMAIL}:${TESTLODGE_API_TOKEN}`),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (Object.keys(params).length > 0) {
    url += '?' + Object.keys(params).map(key => `${key}=${params[key]}`).join('&');
  }

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function squadByTestLodgeProject(projectId) {
  const squadMapping = PHM.Spreadsheet.getRangeValues(TESTLODGE_PROJECT_IDS_SQUADS);

  // Verify if the data was loaded correctly
  if (!squadMapping || squadMapping.length === 0) {
    throw new Error("TestLodge squad mapping data is empty or could not be retrieved.");
  }

  // Iterate over the range to find the corresponding squad
  for (let i = 0; i < squadMapping.length; i++) {
    const [project, squad] = squadMapping[i];
    if (project == projectId) {
      return squad;
    }
  }
  // Return null if the project is not found
  return null;
}