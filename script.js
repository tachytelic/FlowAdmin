const AZURE_CLIENT_ID = '38442c9b-62e6-44a9-a756-effd91ef7b82'; // Replace with your own
const REDIRECT_URI = location.origin + location.pathname;       // Defaults to current page


document.addEventListener("DOMContentLoaded", function () {
  // Initialize Privacy Notice
  const closePrivacyNoticeButton = document.getElementById("closePrivacyNotice");
  const privacyNotice = document.getElementById("privacyNotice");
  if (closePrivacyNoticeButton) {
    closePrivacyNoticeButton.addEventListener("click", function () {
      privacyNotice.style.display = "none";
    });
  }

  // Hide notice if previously dismissed
  if (localStorage.getItem("privacyNoticeDismissed") === "true") {
    privacyNotice.style.display = "none";
  }

  if (closePrivacyNoticeButton) {
    closePrivacyNoticeButton.addEventListener("click", function () {
      privacyNotice.style.display = "none";
      localStorage.setItem("privacyNoticeDismissed", "true");
    });
  }

  // Retrieve and set the last used environment ID
  const savedEnvironmentId = localStorage.getItem("environmentId");
  if (savedEnvironmentId) {
    const environmentDropdown = document.getElementById("environmentDropdown");
    if (environmentDropdown) {
      environmentDropdown.value = savedEnvironmentId;
    }
  }

  // Save environment ID on "Load Flows" button click
  const loadFlowsButton = document.getElementById("loadFlowsButton");
  if (loadFlowsButton) {
    loadFlowsButton.addEventListener("click", function () {
      const environmentDropdown = document.getElementById("environmentDropdown");
      if (environmentDropdown) {
        localStorage.setItem("environmentId", environmentDropdown.value);
      }
    });
  }
});

$(document).ready(function () {
  // MSAL configuration
  const msalConfig = {
    auth: {
      clientId: AZURE_CLIENT_ID, 
      redirectUri: REDIRECT_URI, 
    },
    cache: {
      cacheLocation: 'sessionStorage', // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);

  // Request object for login and token acquisition
  const loginRequest = {
    scopes: ['https://service.flow.microsoft.com/.default'],
  };

  const tokenRequest = {
    scopes: ['https://service.flow.microsoft.com/.default'],
  };

  let environmentId = '';
  let editorVersion = 'true';
  let flowsData = [];

  function initialize() {
    const signInButton = document.getElementById('signInButton');
    if (signInButton) {
      signInButton.addEventListener('click', signIn);
    } else {
      console.error('Button with ID "signInButton" not found.');
    }

    const loadFlowsButton = document.getElementById('loadFlowsButton');
    if (loadFlowsButton) {
      loadFlowsButton.addEventListener('click', loadFlows);
    } else {
      console.error('Button with ID "loadFlowsButton" not found.');
    }

    const editorSelection = document.getElementById('editorSelection');
    if (editorSelection) {
      editorSelection.addEventListener('change', updateEditorVersion);
    }

    const loadAllRunHistoryButton = document.getElementById('loadAllRunHistoryButton');
    if (loadAllRunHistoryButton) {
      loadAllRunHistoryButton.addEventListener("click", fetchAllRunHistories);
    }
  }

  function updateEditorVersion() {
    const editorSelection = document.getElementById('editorSelection');
    editorVersion = editorSelection.value; // 'true' or 'false'

    // If the flows have been loaded, update the grid
    if (flowsData && flowsData.length > 0) {
      displayFlows({ value: flowsData });
    }
  }

  initialize();

  async function signIn() {
    const statusMessage = document.getElementById('statusMessage');
    try {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      msalInstance.setActiveAccount(loginResponse.account);

      // Hide the sign-in button and display sections for authenticated users
      document.getElementById('signInButton').style.display = 'none';
      document.getElementById('environmentSection').style.display = 'block';

      // Show loading message, hide dropdown
      const loadingMsg = document.getElementById('environmentLoadingMessage');
      const dropdown = document.getElementById('environmentDropdown');
      if (loadingMsg) loadingMsg.style.display = 'block';
      if (dropdown) dropdown.style.display = 'none';

      // Fetch and populate environments using working function
      const environments = await fetchUserEnvironments();
      populateEnvironmentDropdown(environments);

      // Hide loading message, show dropdown
      if (loadingMsg) loadingMsg.style.display = 'none';
      if (dropdown) dropdown.style.display = 'block';

      statusMessage.innerHTML = '<p>Signed in successfully!</p>';
    } catch (error) {
      console.error('Error during sign-in:', error);
      alert('Error during sign-in: ' + error.message);
    }
  }

  // Fetch environments using Flow API (user-specific)
  async function fetchUserEnvironments() {
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        throw new Error('No active account! Please sign in again.');
      }

      const tokenRequest = {
        scopes: ['https://service.flow.microsoft.com/.default'],
      };

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      // Fetch all environments the user has access to using Flow API
      const apiUrl = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/';

      const apiResponse = await fetch(apiUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (apiResponse.ok) {
        const data = await apiResponse.json();
        const userEnvironments = data.value || [];
        return userEnvironments;
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error ? errorData.error.message : 'Unknown error';
        throw new Error(`Error fetching environments: ${errorMessage}`);
      }
    } catch (error) {
      if (error instanceof msal.InteractionRequiredAuthError) {
        // Fallback to interactive method if silent acquisition fails
        const response = await msalInstance.acquireTokenPopup({
          scopes: ['https://service.flow.microsoft.com/.default'],
          account: msalInstance.getActiveAccount(),
        });
        // Retry fetching environments
        return await fetchUserEnvironments();
      } else {
        throw error;
      }
    }
  }

  // Populate the dropdown with environments
  function populateEnvironmentDropdown(environments) {
    const dropdown = document.getElementById("environmentDropdown");
    if (!dropdown) {
      alert("Could not find the environment dropdown in the page. Please check your HTML.");
      return;
    }
    dropdown.innerHTML = "";
    environments.forEach(env => {
      const option = document.createElement("option");
      option.value = env.name; // environment ID
      option.textContent = `${env.properties.displayName} (${env.name})`;
      dropdown.appendChild(option);
    });
  }

  async function loadFlows() {
    const environmentDropdown = document.getElementById('environmentDropdown');
    environmentId = environmentDropdown.value.trim();

    const editorSelection = document.getElementById('editorSelection');
    editorVersion = editorSelection.value; // 'true' or 'false'

    const statusMessage = document.getElementById('statusMessage');
    const gridContainer = $('#gridContainer');

    if (!environmentId) {
      alert('Please select an Environment.');
      return;
    }

    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      statusMessage.innerHTML = '<p>Loading flows...</p>';
      gridContainer.hide(); // Hide the grid while loading

      // Construct the API URL with the $select parameter
      const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
        environmentId
      )}/flows?api-version=2016-11-01?$top=100`;

      // Fetch flows for the selected environment
      const apiResponse = await fetch(apiUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (apiResponse.ok) {
        const data = await apiResponse.json();
        flowsData = data.value; // store flows data
        displayFlows({ value: flowsData });
        loadAllRunHistoryButton.style.display = 'block';

        // Dynamically update the success message with the number of flows
        const flowCount = flowsData.length;
        statusMessage.innerHTML = `<p>${flowCount} Flow${flowCount !== 1 ? 's' : ''} Loaded Successfully!</p>`;
        const historyButton = document.getElementById('loadAllRunHistoryButton');
        if (historyButton) {
          historyButton.style.display = 'block';
        }
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error
          ? errorData.error.message
          : 'Unknown error';
        statusMessage.innerHTML = `<p class="error">Error loading flows: ${errorMessage}</p>`;
      }
    } catch (error) {
      if (error instanceof msal.InteractionRequiredAuthError) {
        // Fallback to interactive method if silent acquisition fails
        try {
          const response = await msalInstance.acquireTokenPopup({
            ...tokenRequest,
            account: msalInstance.getActiveAccount(),
          });
          const accessToken = response.accessToken;
          // Retry loading flows
          await loadFlows();
        } catch (popupError) {
          statusMessage.innerHTML = `<p class="error">Error during interactive authentication: ${popupError.message}</p>`;
          console.error('Authentication Error:', popupError);
        }
      } else {
        statusMessage.innerHTML = `<p class="error">Error loading flows: ${error.message}</p>`;
        console.error('Error:', error);
      }
    }
  }

  // Define deleteFlow function before displayFlows
  async function deleteFlow(flowData) {
    const confirmation = confirm(
      `Are you sure you want to delete the flow "${flowData.displayName}"?`
    );
    if (!confirmation) {
      return;
    }

    const statusMessage = document.getElementById('statusMessage');
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
        environmentId
      )}/flows/${encodeURIComponent(flowData.name)}?api-version=2016-11-01`;

      const apiResponse = await fetch(apiUrl, {
        method: 'DELETE',
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (apiResponse.ok) {
        alert(`Flow "${flowData.displayName}" deleted successfully.`);
        // Remove the deleted flow from flowsData
        flowsData = flowsData.filter(
          (flow) => flow.name !== flowData.name
        );
        // Update the grid
        displayFlows({ value: flowsData });
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error
          ? errorData.error.message
          : 'Unknown error';
        alert(`Error deleting flow: ${errorMessage}`);
      }
    } catch (error) {
      console.error('Error deleting flow:', error);
      alert(`Error deleting flow: ${error.message}`);
    }
  }

  // Define exportFlow function before displayFlows
  async function exportFlow(flowData) {
    const statusMessage = document.getElementById('statusMessage');
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
        environmentId
      )}/flows/${encodeURIComponent(
        flowData.name
      )}/exportToArmTemplate?api-version=2016-11-01`;

      const apiResponse = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({}),
      });

      if (apiResponse.ok) {
        const jsonData = await apiResponse.json();
        const fileName = `${flowData.displayName}.json`;

        // Create a blob and trigger download
        const blob = new Blob([JSON.stringify(jsonData, null, 2)], {
          type: 'application/json',
        });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error
          ? errorData.error.message
          : 'Unknown error';
        alert(`Error exporting flow: ${errorMessage}`);
      }
    } catch (error) {
      console.error('Error exporting flow:', error);
      alert(`Error exporting flow: ${error.message}`);
    }
  }

  async function toggleFlowState(flowData) {
    const action = flowData.state === 'Started' ? 'stop' : 'start';
    const confirmAction = action === 'stop' ? 'Turn off' : 'Turn on';
    const confirmation = confirm(
      `Are you sure you want to ${confirmAction} the flow "${flowData.displayName}"?`
    );
    if (!confirmation) {
      return;
    }

    const statusMessage = document.getElementById('statusMessage');
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
        environmentId
      )}/flows/${encodeURIComponent(flowData.name)}/${action}?api-version=2016-11-01`;

      const apiResponse = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (apiResponse.ok) {
        const newState = action === 'stop' ? 'Stopped' : 'Started';
        alert(`Flow "${flowData.displayName}" is now ${newState}.`);
        // Update the flow's state in flowsData
        for (let flow of flowsData) {
          if (flow.name === flowData.name) {
            flow.properties.state = newState;
            flow.state = newState;
            break;
          }
        }
        // Update the grid
        displayFlows({ value: flowsData });
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error
          ? errorData.error.message
          : 'Unknown error';
        alert(`Error changing flow state: ${errorMessage}`);
      }
    } catch (error) {
      console.error('Error changing flow state:', error);
      alert(`Error changing flow state: ${error.message}`);
    }
  }


  function displayFlows(data) {
    const gridContainer = $('#gridContainer');
    if (gridContainer.length === 0) {
      console.error('Element with ID "gridContainer" not found.');
      return;
    }
  
    const flows = data.value;
  
    if (flows && flows.length > 0) {
      const gridData = flows.map((flow) => {
        const name = flow.name;
        const properties = flow.properties || {};
        const displayName = properties.displayName || 'N/A';
        const state = properties.state || 'Unknown';
      const definitionSummary = properties.definitionSummary || {};
      const triggers = definitionSummary.triggers || [];
      const actions = definitionSummary.actions || [];
    
        // Extract from first trigger (if available)
      const trigger = triggers.length > 0 ? triggers[0] : {};
      const triggerType = trigger.type || '';
      const triggerKind = trigger.kind || '';
      const operationId = trigger.metadata?.operationMetadataId || '';
      const actionCount = actions.length;		
    
        const editLink = `https://make.powerautomate.com/environments/${encodeURIComponent(
          environmentId
        )}/flows/shared/${encodeURIComponent(name)}?v3=${editorVersion}`; // Construct URL with v3 parameter
  
    return {
     displayName,
     name,
     state,
     editLink,
     triggerType,
     triggerKind,
     operationId,
     actionCount,
     ...properties,
    };
      });
  
      if (gridContainer.data('dxDataGrid')) {
        gridContainer.dxDataGrid('dispose');
        gridContainer.empty();
      }
  
      gridContainer.show();
  
      gridContainer.dxDataGrid({
        dataSource: gridData,
        keyExpr: 'name',
        columns: [
          { dataField: 'displayName', caption: 'Display Name', allowSearch: true },
          { dataField: 'triggerType', caption: 'Trigger Type', allowSearch: false },
          { dataField: 'triggerKind', caption: 'Trigger Kind', allowSearch: false },
          { dataField: 'operationId', caption: 'Operation ID', allowSearch: false },
          { dataField: 'actionCount', caption: 'Action Count', dataType: 'number', allowSearch: false },
          { dataField: 'failureAlertSubscribed', caption: 'Failure Alert Subscribed', visible: false, allowSearch: false },
          { dataField: 'name', caption: 'ID', width: 250, visible: false },
          { dataField: 'userType', caption: 'User Type', visible: false, allowSearch: false },
          {
            dataField: 'createdTime',
            caption: 'Created Time',
            dataType: 'date',
            format: 'yyyy-MM-dd HH:mm:ss',
            sortOrder: 'desc',
            sortIndex: 0,
          },
          {
            dataField: 'lastModifiedTime',
            caption: 'Last Modified Time',
            dataType: 'date',
            format: 'yyyy-MM-dd HH:mm:ss',
          },
          { dataField: 'isManaged', caption: 'Is Managed', visible: false, allowSearch: false },
          {
            dataField: 'runHistoryLink',
            caption: 'Run History',
            allowSorting: false,
            allowFiltering: false,
            visible: false,
            cellTemplate: function (container, options) {
              $('<a>')
                .attr('href', options.data.runHistoryLink)
                .attr('target', '_blank')
                .text('Run History')
                .appendTo(container);
            },
          },
          {
            dataField: 'state',
            allowSearch: false,
            caption: 'State',
            cellTemplate: function (container, options) {
              const stateText = options.data.state;
              const span = $('<span>').text(stateText);
              if (stateText === 'Stopped') {
                span.css('color', 'red');
              }
              span.appendTo(container);
            },
          },
          {
              dataField: 'actions',
              caption: 'Actions',
              allowSorting: false,
              allowFiltering: false,
              cellTemplate: function (container, options) {
                // Edit Button
                $('<button>')
                  .text('Edit')
                  .addClass('btn btn-primary btn-sm')
                  .on('click', function () {
                    window.open(options.data.editLink, '_blank');
                  })
                  .appendTo(container);
  
                // Delete Button
                $('<button>')
                  .text('Delete')
                  .addClass('btn btn-danger btn-sm')
                  .css('margin-left', '5px')
                  .on('click', function () {
                    deleteFlow(options.data);
                  })
                  .appendTo(container);
  
                // Export Button
                $('<button>')
                  .text('Export')
                  .addClass('btn btn-secondary btn-sm')
                  .css('margin-left', '5px')
                  .on('click', function () {
                    exportFlow(options.data);
                  })
                  .appendTo(container);
  
                // Turn On/Off Button
                const toggleText =
                  options.data.state === 'Started' ? 'Turn off' : 'Turn on';
                $('<button>')
                  .text(toggleText)
                  .addClass('btn btn-secondary btn-sm')
                  .css('margin-left', '5px')
                  .on('click', function () {
                    toggleFlowState(options.data);
                  })
                  .appendTo(container);
              },
            },
            ,
        ],
        rowAlternationEnabled: true,
        columnAutoWidth: true,
        showBorders: true,
        paging: {
          pageSize: 10,
        },
        pager: {
          showPageSizeSelector: true,
          allowedPageSizes: [5, 10, 20, 50, 100],
          showInfo: true,
        },
        filterRow: {
          visible: true,
        },
        headerFilter: {
          visible: true,
        },
        groupPanel: {
          visible: false,
        },
        masterDetail: {
          enabled: true,
          template: async function (detailElement, detailInfo) {
            const flowData = detailInfo.data;
  
            // Create a container for the detail grid
            const detailGrid = $('<div>').appendTo(detailElement);
  
            // Show a loading indicator
            detailGrid.text('Loading run history...');
  
            try {
              // Fetch run history data for the selected flow
              const runHistoryData = await fetchRunHistory(flowData.name);
  
              // Initialize the detail grid
              detailGrid.dxDataGrid({
                dataSource: runHistoryData,
                columns: [
                  {
                    dataField: 'startTime',
                    caption: 'Start Time',
                    dataType: 'date',
                    format: 'yyyy-MM-dd HH:mm:ss',
                  },
                  {
                    dataField: 'endTime',
                    caption: 'End Time',
                    dataType: 'date',
                    format: 'yyyy-MM-dd HH:mm:ss',
                  },
                  {
                    dataField: 'duration',
                    caption: 'Duration (seconds)',
                    dataType: 'number',
                  },
                  {
                    dataField: 'status',
                    caption: 'Status',
                    cellTemplate: function (container, options) {
                      const flowRunLink = `https://make.powerautomate.com/environments/${encodeURIComponent(
                        environmentId
                      )}/flows/${encodeURIComponent(
                        flowData.name
                      )}/runs/${encodeURIComponent(options.data.runId)}`;
                      const statusText = options.data.status;
                      const link = $('<a>')
                        .attr('href', flowRunLink)
                        .attr('target', '_blank')
                        .text(statusText);
  
                      if (statusText === 'Failed') {
                        link.css('color', 'red'); // Highlight failed status in red
                      }
  
                      link.appendTo(container);
                    },
                  },
                ],
                columnAutoWidth: true,
                showBorders: true,
              });
            } catch (error) {
              console.error('Error fetching run history:', error);
              detailGrid.text('Error loading run history.');
            }
          },
        },
      });
    } else {
      gridContainer.html('<p>No flows found.</p>');
      gridContainer.show();
    }
  }
  

  // Function to fetch run history for a flow
  async function fetchRunHistory(flowId) {
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }

      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });

      const accessToken = response.accessToken;

      // Construct the API URL
      const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
        environmentId
      )}/flows/${encodeURIComponent(
        flowId
      )}/runs?api-version=2016-11-01`;

      const apiResponse = await fetch(apiUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (apiResponse.ok) {
        const data = await apiResponse.json();
        const runHistoryData = data.value.map((run) => {
          const properties = run.properties || {};
          const startTime = properties.startTime
            ? new Date(properties.startTime)
            : null;
          const endTime = properties.endTime
            ? new Date(properties.endTime)
            : null;
          const status = properties.status || 'Unknown';
          const runId = run.name; // Get the run ID

          // Calculate duration in seconds
          const duration =
            startTime && endTime
              ? (endTime - startTime) / 1000
              : null;

          return {
            startTime,
            endTime,
            duration,
            status,
            runId,
          };
        });
        return runHistoryData;
      } else {
        const errorData = await apiResponse.json();
        const errorMessage = errorData.error
          ? errorData.error.message
          : 'Unknown error';
        throw new Error(errorMessage);
      }
    } catch (error) {
      console.error('Error in fetchRunHistory:', error);
      throw error;
    }
  }

  // Master detail template function
  function masterDetailTemplate(detailElement, detailInfo) {
    const flowData = detailInfo.data;

    // Create a container for the detail grid
    const detailGrid = $('<div>').appendTo(detailElement);

    // Show a loading indicator
    detailGrid.text('Loading run history...');

    // Fetch run history data
    fetchRunHistory(flowData.name)
      .then((runHistoryData) => {
        // Initialize the detail grid
        detailGrid.dxDataGrid({
          dataSource: runHistoryData,
          columns: [
            {
              dataField: 'startTime',
              caption: 'Start Time',
              dataType: 'date',
              format: 'yyyy-MM-dd HH:mm:ss',
            },
            {
              dataField: 'endTime',
              caption: 'End Time',
              dataType: 'date',
              format: 'yyyy-MM-dd HH:mm:ss',
            },
            {
              dataField: 'duration',
              caption: 'Duration (seconds)',
              dataType: 'number',
              format: {
                type: 'fixedPoint',
                precision: 2,
              },
            },
            {
              dataField: 'status',
              caption: 'Status',
              cellTemplate: function (container, options) {
                const flowRunLink = `https://make.powerautomate.com/environments/${encodeURIComponent(
                  environmentId
                )}/flows/${encodeURIComponent(
                  flowData.name
                )}/runs/${encodeURIComponent(options.data.runId)}`;
                const statusText = options.data.status;
                const link = $('<a>')
                  .attr('href', flowRunLink)
                  .attr('target', '_blank')
                  .text(statusText);

                if (statusText === 'Failed') {
                  link.css('color', 'red'); // Highlight failed status in red
                }

                link.appendTo(container);
              },
            },
          ],
          columnAutoWidth: true,
          showBorders: true,
        });
      })
      .catch((error) => {
        console.error('Error fetching run history:', error);
        detailGrid.text('Error loading run history.');
      });
  }


  async function fetchAllRunHistories() {
    const statusMessage = document.getElementById('statusMessage');
    statusMessage.innerHTML = '<p>Loading run history...</p>';
  
    try {
      const activeAccount = msalInstance.getActiveAccount();
      if (!activeAccount) {
        alert('No active account! Please sign in again.');
        return;
      }
  
      const response = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: activeAccount,
      });
  
      const accessToken = response.accessToken;
  
      let allRunHistory = [];
      const totalFlows = flowsData.length;
  
      // Testing variable to limit flows (set to 50 for testing, toggle as needed)
      const TESTING_FLOW_LIMIT = 50; // Change this as needed
      const isTesting = false; // Toggle this for full loading or testing
      const flowLimit = isTesting ? TESTING_FLOW_LIMIT : totalFlows;
  
      // Function to process each flow
      async function fetchFlowRunHistory(flow) {
        const apiUrl = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(
          environmentId
        )}/flows/${encodeURIComponent(
          flow.name
        )}/runs?api-version=2016-11-01`;
  
        const apiResponse = await fetch(apiUrl, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        });
  
        if (apiResponse.ok) {
          const data = await apiResponse.json();
          const runHistoryData = data.value.map((run) => {
            const properties = run.properties || {};
            const startTime = properties.startTime
              ? new Date(properties.startTime)
              : null;
            const endTime = properties.endTime
              ? new Date(properties.endTime)
              : null;
            const duration =
              startTime && endTime
                ? (endTime - startTime) / 1000
                : null;
            return {
              flowName: flow.properties.displayName,
              startTime,
              endTime,
              duration,
              status: properties.status || 'Unknown',
            };
          });
          return runHistoryData;
        } else {
          console.warn(`Failed to fetch run history for flow ${flow.name}`);
          return [];
        }
      }
  
      // Batched requests with concurrency control
      const BATCH_SIZE = 10;
      for (let i = 0; i < flowLimit; i += BATCH_SIZE) {
        const batch = flowsData.slice(i, i + BATCH_SIZE);
  
        // Fetch run histories for the batch
        const batchResults = await Promise.all(
          batch.map((flow) => fetchFlowRunHistory(flow))
        );
  
        // Update progress
        allRunHistory = allRunHistory.concat(...batchResults);
        const processedCount = Math.min(i + BATCH_SIZE, flowLimit);
        statusMessage.innerHTML = `<p>Loaded ${processedCount} of ${flowLimit} flows...</p>`;
      }
  
      // Display the grid with all collected run history
      displayAllRunHistory(allRunHistory);
      statusMessage.innerHTML = `<p>All run history loaded for ${flowLimit} flows.</p>`;
    } catch (error) {
      console.error('Error fetching all run histories:', error);
      statusMessage.innerHTML = `<p class="error">Error: ${error.message}</p>`;
    }
  }
  

  function displayAllRunHistory(data) {
    const gridContainer = $('#allRunHistoryGridContainer');
  
    // Make the container visible
    gridContainer[0].style.display = 'block';
  
    if (gridContainer.data('dxDataGrid')) {
      gridContainer.dxDataGrid('dispose');
      gridContainer.empty();
    }
  
    gridContainer.dxDataGrid({
      dataSource: data,
      columns: [
        { dataField: 'flowName', caption: 'Flow Name', allowSearch: true },
        {
          dataField: 'startTime',
          caption: 'Start Time',
          dataType: 'date',
          format: 'yyyy-MM-dd HH:mm:ss',
          sortOrder: 'desc', // Default sort order (descending)
          sortIndex: 0, // This column is the primary sort column
        },
        {
          dataField: 'endTime',
          caption: 'End Time',
          dataType: 'date',
          format: 'yyyy-MM-dd HH:mm:ss',
        },
        {
          dataField: 'duration',
          caption: 'Duration (seconds)',
          dataType: 'number',
          filterOperations: ['<', '>', '=', 'between'],
        },
        { dataField: 'status', caption: 'Status' },
      ],
      filterRow: { visible: true },
      headerFilter: { visible: true },
      columnAutoWidth: true,
      rowAlternationEnabled: true,
      showBorders: true,
      paging: {
        pageSize: 20, // Matches the other grid
      },
      pager: {
        showPageSizeSelector: true,
        allowedPageSizes: [5, 10, 20, 50, 100], // Matches the other grid
        showInfo: true,
      },
      groupPanel: { visible: true }, // Enable the group panel
      grouping: {
        autoExpandAll: true, // Automatically expand all groups by default
      },
    });
  }
  
    

});


