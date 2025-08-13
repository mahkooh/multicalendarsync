/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { CalendarSyncManager } from './calendarSync.js';
import { MultiAccountUI } from './multiAccountUI.js';

let syncManager;
let multiAccountUI;

Office.onReady((info) => {
  try {
    console.log('🚀 Office.onReady called with info:', info);
    console.log('📊 Office.js diagnostics:');
    console.log('  - Office object type:', typeof Office);
    console.log('  - Office.context type:', typeof Office?.context);
    console.log('  - Office.context.auth type:', typeof Office?.context?.auth);
    console.log('  - Host detected:', info?.host);
    console.log('  - Host type comparison:', info?.host === Office.HostType.Outlook);
    console.log('  - Platform:', info?.platform);
    
    // More permissive host check
    if (info.host === Office.HostType.Outlook || info.host === 'Outlook' || window.Office) {
      console.log('✅ Host validated as Outlook-compatible');
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      
      // Set default date to August 12, 2025 for testing
      const defaultDate = new Date('2025-08-12');
      document.getElementById("sync-date").value = defaultDate.toISOString().split('T')[0];
      
      // Initialize the calendar sync manager
      syncManager = new CalendarSyncManager();
      initializeApp();
    } else {
      console.warn('⚠️ Unsupported or unrecognized host application. Host info:', info);
      // Show diagnostic info but try to continue
      document.getElementById("sideload-msg").innerHTML = `
        <div style="color: orange; padding: 10px; border: 1px solid orange; border-radius: 5px;">
          <h3>🔍 Host Detection Diagnostic</h3>
          <p><strong>Detected host:</strong> ${info?.host || 'Unknown'}</p>
          <p><strong>Expected:</strong> Outlook</p>
          <p><strong>Platform:</strong> ${info?.platform || 'Unknown'}</p>
          <p><strong>Office.js available:</strong> ${!!window.Office ? 'Yes' : 'No'}</p>
          <p><strong>Status:</strong> Attempting to continue with limited functionality...</p>
          <button onclick="this.parentElement.style.display='none'; document.getElementById('app-body').style.display='flex';">Continue Anyway</button>
        </div>
      `;
      
      // Try to initialize anyway for testing
      setTimeout(() => {
        try {
          syncManager = new CalendarSyncManager();
          initializeApp();
        } catch (initError) {
          console.error('❌ Delayed initialization failed:', initError);
        }
      }, 1000);
    }
  } catch (error) {
    console.error('❌ Office.onReady error:', error);
    document.getElementById("sideload-msg").innerHTML = `
      <div style="color: red; padding: 10px; border: 1px solid red; border-radius: 5px;">
        <h3>🚨 Initialization Error</h3>
        <p><strong>Error:</strong> ${error.message}</p>
        <p><strong>Suggestion:</strong> Please ensure this add-in is loaded within Outlook</p>
        <p><strong>Current URL:</strong> ${window.location.href}</p>
        <button onclick="location.reload();">Refresh Page</button>
        <button onclick="this.parentElement.style.display='none'; document.getElementById('app-body').style.display='flex';">Try Anyway</button>
      </div>
    `;
  }
});

// Fallback initialization if Office.onReady doesn't work
window.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (!syncManager) {
      console.log('🔄 Fallback initialization triggered - Office.onReady may not have fired');
      console.log('🔍 Office.js status check:');
      console.log('  - window.Office available:', !!window.Office);
      console.log('  - Current URL:', window.location.href);
      console.log('  - User Agent:', navigator.userAgent);
      
      if (window.Office) {
        console.log('✅ Office.js detected, attempting manual initialization');
        try {
          document.getElementById("sideload-msg").style.display = "none";
          document.getElementById("app-body").style.display = "flex";
          
          const defaultDate = new Date('2025-08-12');
          document.getElementById("sync-date").value = defaultDate.toISOString().split('T')[0];
          
          syncManager = new CalendarSyncManager();
          initializeApp();
        } catch (error) {
          console.error('❌ Fallback initialization failed:', error);
          showContextGuidance();
        }
      } else {
        console.warn('⚠️ Office.js not available - showing context guidance');
        showContextGuidance();
      }
    }
  }, 2000); // Wait 2 seconds for Office.js to load
});

function showContextGuidance() {
  document.getElementById("sideload-msg").innerHTML = `
    <div style="color: #d63031; padding: 15px; border: 1px solid #d63031; border-radius: 8px; margin: 10px;">
      <h3>🚨 Office.js Context Required</h3>
      <p><strong>Issue:</strong> This add-in needs to run within Microsoft Outlook to access Office.js APIs.</p>
      
      <h4>📋 How to test this add-in properly:</h4>
      <ol style="text-align: left; margin: 10px 0;">
        <li><strong>Deploy the manifest:</strong> Upload <code>manifest.xml</code> to Microsoft 365 Admin Center</li>
        <li><strong>Open Outlook:</strong> Use Outlook desktop app or Outlook on the web</li>
        <li><strong>Find the add-in:</strong> Look for "Sync Calendars" button in the ribbon</li>
        <li><strong>Click to open:</strong> The add-in will load in a proper Office.js context</li>
      </ol>
      
      <h4>🔧 For development testing:</h4>
      <ul style="text-align: left; margin: 10px 0;">
        <li>Use <code>npm run dev-server</code> and sideload in Outlook</li>
        <li>Or use Office Add-in Development Tools</li>
        <li>Direct browser testing won't have full Office.js capabilities</li>
      </ul>
      
      <p><strong>Current context:</strong> ${window.location.href}</p>
      <button onclick="testMockMode();" style="background: #0984e3; color: white; border: none; padding: 8px 16px; border-radius: 4px; margin: 5px;">
        Test with Mock Data
      </button>
      <button onclick="location.reload();" style="background: #636e72; color: white; border: none; padding: 8px 16px; border-radius: 4px; margin: 5px;">
        Refresh
      </button>
    </div>
  `;
}

function testMockMode() {
  console.log('🧪 Enabling mock mode for testing without Office.js');
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  const defaultDate = new Date('2025-08-12');
  document.getElementById("sync-date").value = defaultDate.toISOString().split('T')[0];
  
  // Initialize with mock-only mode
  syncManager = new CalendarSyncManager();
  initializeApp();
}

async function initializeApp() {
  try {
    console.log('� Starting Multi-Tenant Calendar Sync initialization...');
    
    // Initialize the multi-account UI
    if (window.multiTenantManager) {
      multiAccountUI = new MultiAccountUI(window.multiTenantManager);
      window.multiAccountUI = multiAccountUI; // Make globally available
      multiAccountUI.initialize();
      
      // Hide the old sections and show the multi-account interface
      const statusSection = document.getElementById('status-section');
      const calendarSection = document.getElementById('calendar-section');
      if (statusSection) statusSection.style.display = 'none';
      if (calendarSection) calendarSection.style.display = 'none';
      
      console.log('✅ Multi-account interface initialized');
    } else {
      console.warn('⚠️ Multi-tenant manager not available, falling back to simple mode');
      // Fall back to simple authentication
      if (window.calendarAuth) {
        window.calendarAuth.renderAuthenticationOptions();
      }
    }
    
    // Set up event listeners
    setupEventListeners();
    
    // Initialize the calendar sync manager
    await syncManager.initialize();
    
    // Update the display
    updateSyncStatus('ready', 'Ready to sync');
    
    console.log('✅ Multi-tenant calendar sync initialized successfully');
    
  } catch (error) {
    console.error('❌ App initialization failed:', error);
    updateSyncStatus('error', `Initialization failed: ${error.message}`);
    showContextGuidance();
  }
}

function setupEventListeners() {
  console.log('⚙️ Setting up event listeners...');
  
  // Diagnostic button
  const diagnosticBtn = document.getElementById('run-diagnostic');
  if (diagnosticBtn) {
    diagnosticBtn.addEventListener('click', runCalendarDiagnostic);
  }
  
  // Sync button
  const syncBtn = document.getElementById('run-sync');
  if (syncBtn) {
    syncBtn.addEventListener('click', runSync);
  }
  
  // Settings button  
  const settingsBtn = document.getElementById('configure');
  if (settingsBtn) {
    settingsBtn.addEventListener('click', showSettings);
  }
  
  console.log('✅ Event listeners set up');
}

async function runCalendarDiagnostic() {
  console.log('🔍 Running calendar discovery diagnostic...');
  
  const diagnosticBtn = document.getElementById('run-diagnostic');
  const resultsDiv = document.getElementById('diagnostic-results');
  
  if (diagnosticBtn) {
    diagnosticBtn.disabled = true;
    diagnosticBtn.innerHTML = '<span class="ms-Button-label">🔄 Checking...</span>';
  }
  
  try {
    const diagnostic = new CalendarDiscoveryDiagnostic();
    const results = await diagnostic.runFullDiagnostic();
    
    // Update the results area
    if (resultsDiv) {
      const calendarCount = results.graphApiResults.calendarsFound.length;
      const authSuccess = results.graphApiResults.accessTokenObtained;
      
      resultsDiv.innerHTML = `
        <div class="diagnostic-summary">
          <h4>📊 Discovery Results:</h4>
          <p><strong>Graph API Auth:</strong> ${authSuccess ? '✅ Working' : '❌ Failed'}</p>
          <p><strong>Calendars Found:</strong> ${calendarCount} calendar(s)</p>
          ${calendarCount > 0 ? `
            <div class="calendar-list">
              <strong>Found calendars:</strong>
              <ul>
                ${results.graphApiResults.calendarsFound.map(cal => 
                  `<li>📅 ${cal.name} ${cal.isDefaultCalendar ? '(Default)' : ''}</li>`
                ).join('')}
              </ul>
            </div>
          ` : `
            <p><em>This means Graph API can only see calendars from your current Office 365/Exchange account.</em></p>
            <p><em>Connected accounts (Gmail, other Exchange) won't appear here.</em></p>
          `}
          <button id="add-external-calendar" class="ms-Button" style="margin-top: 10px;">
            <span class="ms-Button-label">➕ Add External Calendar</span>
          </button>
        </div>
      `;
      
      // Add event listener for external calendar button
      const addExternalBtn = document.getElementById('add-external-calendar');
      if (addExternalBtn) {
        addExternalBtn.addEventListener('click', showAddExternalCalendar);
      }
    }
    
  } catch (error) {
    console.error('❌ Diagnostic failed:', error);
    if (resultsDiv) {
      resultsDiv.innerHTML = `
        <div class="error-message">
          <p><strong>❌ Diagnostic Failed:</strong> ${error.message}</p>
          <p><em>This likely means you're not in a proper Office.js context or don't have the necessary permissions.</em></p>
        </div>
      `;
    }
  } finally {
    if (diagnosticBtn) {
      diagnosticBtn.disabled = false;
      diagnosticBtn.innerHTML = '<span class="ms-Button-label">🔍 Check My Calendars</span>';
    }
  }
}

function showAddExternalCalendar() {
  const resultsDiv = document.getElementById('diagnostic-results');
  if (!resultsDiv) return;
  
  resultsDiv.innerHTML += `
    <div class="add-calendar-section" style="margin-top: 15px; padding: 10px; border: 1px solid #ccc; border-radius: 5px;">
      <h4>➕ Add External Calendar</h4>
      <p>To sync calendars from other accounts, choose an option:</p>
      
      <div class="calendar-options">
        <button class="calendar-type-btn" onclick="addGoogleCalendar()">
          📧 Google Calendar
        </button>
        <button class="calendar-type-btn" onclick="addOutlookCalendar()">
          📮 Outlook.com/Hotmail
        </button>
        <button class="calendar-type-btn" onclick="addExchangeCalendar()">
          🏢 Exchange Server
        </button>
        <button class="calendar-type-btn" onclick="addICSCalendar()">
          🔗 ICS/WebCal URL
        </button>
      </div>
      
      <div id="calendar-form-container" style="margin-top: 10px;">
        <!-- Calendar form will be inserted here -->
      </div>
    </div>
  `;
}

async function loadCalendars() {
  try {
    const calendars = await syncManager.getAvailableCalendars();
    displayCalendars(calendars);
  } catch (error) {
    console.error('Failed to load calendars:', error);
    document.getElementById('calendar-list').innerHTML = 
      '<div class="error">Failed to load calendars. Please check permissions.</div>';
  }
}

function displayCalendars(calendars) {
  const calendarList = document.getElementById('calendar-list');
  
  if (calendars.length === 0) {
    calendarList.innerHTML = '<div class="no-activity">No calendars found</div>';
    return;
  }
  
  calendarList.innerHTML = calendars.map(calendar => `
    <div class="calendar-item">
      <div>
        <div class="calendar-name">${calendar.name}</div>
        <div class="calendar-status">${calendar.type} • ${calendar.itemCount} items</div>
      </div>
      <div class="calendar-toggle">
        <div class="toggle-switch ${calendar.syncEnabled ? 'active' : ''}" 
             data-calendar-id="${calendar.id}"
             onclick="toggleCalendarSync('${calendar.id}')">
        </div>
      </div>
    </div>
  `).join('');
}

async function handleSyncNow() {
  try {
    // Get the selected date
    const selectedDate = document.getElementById("sync-date").value;
    if (!selectedDate) {
      updateStatus('Error', 'Please select a date to sync');
      return;
    }
    
    const syncDate = new Date(selectedDate);
    updateStatus('Syncing', `Syncing ${syncDate.toLocaleDateString()}...`);
    updateSyncButton(true);
    
    // Pass the selected date to the sync manager
    const result = await syncManager.performSync(syncDate);
    
    updateStatus('Active', `Sync completed for ${syncDate.toLocaleDateString()}. ${result.blocksCreated} blocks created, ${result.blocksRemoved} removed.`);
    addActivityLog(`Sync completed for ${syncDate.toLocaleDateString()}: ${result.blocksCreated} created, ${result.blocksRemoved} removed`);
    
  } catch (error) {
    console.error('Sync failed:', error);
    updateStatus('Error', `Sync failed: ${error.message}`);
    addActivityLog(`Sync failed: ${error.message}`);
  } finally {
    updateSyncButton(false);
  }
}

async function handleConfigure() {
  // TODO: Open configuration dialog
  console.log('Configure clicked');
}

async function toggleCalendarSync(calendarId) {
  try {
    const enabled = await syncManager.toggleCalendarSync(calendarId);
    
    // Update the toggle switch appearance
    const toggle = document.querySelector(`[data-calendar-id="${calendarId}"]`);
    if (toggle) {
      toggle.classList.toggle('active', enabled);
    }
    
    addActivityLog(`Calendar sync ${enabled ? 'enabled' : 'disabled'} for calendar`);
    
  } catch (error) {
    console.error('Failed to toggle calendar sync:', error);
    addActivityLog(`Failed to toggle calendar sync: ${error.message}`);
  }
}

async function updateSyncStatus() {
  try {
    const status = await syncManager.getSyncStatus();
    updateStatus(status.state, status.message);
    
    if (status.lastSync) {
      document.getElementById('last-sync').textContent = 
        `Last sync: ${new Date(status.lastSync).toLocaleString()}`;
    }
    
  } catch (error) {
    console.error('Failed to get sync status:', error);
    updateStatus('Error', 'Unable to check sync status');
  }
}

function updateStatus(state, message) {
  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  
  // Remove all state classes
  statusDot.classList.remove('active', 'syncing');
  
  // Add appropriate class
  switch (state.toLowerCase()) {
    case 'active':
      statusDot.classList.add('active');
      break;
    case 'syncing':
      statusDot.classList.add('syncing');
      break;
    default:
      // Default red color for error/stopped
      break;
  }
  
  statusText.textContent = message;
}

function updateSyncButton(isDisabled) {
  const button = document.getElementById('sync-now');
  const label = button.querySelector('.ms-Button-label');
  
  button.style.opacity = isDisabled ? '0.6' : '1';
  button.style.pointerEvents = isDisabled ? 'none' : 'auto';
  label.textContent = isDisabled ? 'Syncing...' : 'Sync Now';
}

function addActivityLog(message) {
  const activityLog = document.getElementById('activity-log');
  const noActivity = activityLog.querySelector('.no-activity');
  
  if (noActivity) {
    noActivity.remove();
  }
  
  const activityItem = document.createElement('div');
  activityItem.className = 'activity-item';
  activityItem.innerHTML = `
    <div class="activity-time">${new Date().toLocaleTimeString()}</div>
    <div>${message}</div>
  `;
  
  activityLog.insertBefore(activityItem, activityLog.firstChild);
  
  // Keep only the last 10 items
  const items = activityLog.querySelectorAll('.activity-item');
  if (items.length > 10) {
    items[items.length - 1].remove();
  }
}

// Make toggle function globally available
window.toggleCalendarSync = toggleCalendarSync;

// Global functions for calendar addition (called from onclick handlers)
window.addGoogleCalendar = function() {
  showCalendarForm('google', 'Google Calendar', {
    fields: [
      { name: 'email', label: 'Google Account Email', type: 'email', required: true },
      { name: 'calendarId', label: 'Calendar ID (optional - leave blank for primary)', type: 'text', required: false }
    ],
    instructions: 'You\'ll need to authenticate with Google Calendar API. We\'ll redirect you to Google for permission.'
  });
};

window.addOutlookCalendar = function() {
  showCalendarForm('outlook', 'Outlook.com/Hotmail', {
    fields: [
      { name: 'email', label: 'Outlook.com Email', type: 'email', required: true },
      { name: 'calendarName', label: 'Calendar Name (optional)', type: 'text', required: false }
    ],
    instructions: 'We\'ll authenticate with Microsoft Graph API for your personal Outlook.com account.'
  });
};

window.addExchangeCalendar = function() {
  showCalendarForm('exchange', 'Exchange Server', {
    fields: [
      { name: 'email', label: 'Exchange Email', type: 'email', required: true },
      { name: 'serverUrl', label: 'Exchange Server URL', type: 'url', required: true, placeholder: 'https://mail.company.com/ews/exchange.asmx' },
      { name: 'domain', label: 'Domain (optional)', type: 'text', required: false }
    ],
    instructions: 'We\'ll connect to your Exchange server using EWS (Exchange Web Services).'
  });
};

window.addICSCalendar = function() {
  showCalendarForm('ics', 'ICS/WebCal Calendar', {
    fields: [
      { name: 'name', label: 'Calendar Name', type: 'text', required: true },
      { name: 'url', label: 'Calendar URL', type: 'url', required: true, placeholder: 'https://calendar.google.com/calendar/ical/...ics' },
      { name: 'refreshInterval', label: 'Refresh Interval (minutes)', type: 'number', required: false, value: '60' }
    ],
    instructions: 'Enter the ICS or WebCal URL. This is read-only and will be refreshed periodically.'
  });
};

function showCalendarForm(type, typeName, config) {
  const container = document.getElementById('calendar-form-container');
  if (!container) return;
  
  const fieldsHtml = config.fields.map(field => `
    <div class="form-field" style="margin: 10px 0;">
      <label for="${field.name}" style="display: block; font-weight: bold; margin-bottom: 5px;">
        ${field.label} ${field.required ? '*' : ''}
      </label>
      <input 
        type="${field.type}" 
        id="${field.name}" 
        name="${field.name}"
        placeholder="${field.placeholder || ''}"
        value="${field.value || ''}"
        ${field.required ? 'required' : ''}
        style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"
      />
    </div>
  `).join('');
  
  container.innerHTML = `
    <div class="calendar-form">
      <h5>📝 Add ${typeName}</h5>
      <p style="font-size: 0.9em; color: #666;">${config.instructions}</p>
      
      <form id="add-calendar-form">
        <input type="hidden" name="type" value="${type}" />
        ${fieldsHtml}
        
        <div class="form-actions" style="margin-top: 15px;">
          <button type="submit" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">🔗 Connect Calendar</span>
          </button>
          <button type="button" onclick="cancelCalendarForm()" class="ms-Button" style="margin-left: 10px;">
            <span class="ms-Button-label">Cancel</span>
          </button>
        </div>
      </form>
    </div>
  `;
  
  // Add form submit handler
  const form = document.getElementById('add-calendar-form');
  if (form) {
    form.addEventListener('submit', handleCalendarFormSubmit);
  }
}

window.cancelCalendarForm = function() {
  const container = document.getElementById('calendar-form-container');
  if (container) {
    container.innerHTML = '';
  }
};

async function handleCalendarFormSubmit(event) {
  event.preventDefault();
  
  const formData = new FormData(event.target);
  const calendarData = {};
  
  for (let [key, value] of formData.entries()) {
    calendarData[key] = value;
  }
  
  console.log('📝 Calendar form submitted:', calendarData);
  
  try {
    // Show loading state
    const submitBtn = event.target.querySelector('button[type="submit"]');
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="ms-Button-label">🔄 Connecting...</span>';
    }
    
    // Process the calendar addition based on type
    await addExternalCalendarConnection(calendarData);
    
    // Success
    const container = document.getElementById('calendar-form-container');
    if (container) {
      container.innerHTML = `
        <div class="success-message" style="color: green; padding: 10px; border: 1px solid green; border-radius: 4px;">
          ✅ Calendar "${calendarData.name || calendarData.email}" has been added successfully!
          <button onclick="runCalendarDiagnostic()" style="margin-left: 10px;">Refresh List</button>
        </div>
      `;
    }
    
  } catch (error) {
    console.error('❌ Failed to add calendar:', error);
    
    const container = document.getElementById('calendar-form-container');
    if (container) {
      container.innerHTML = `
        <div class="error-message" style="color: red; padding: 10px; border: 1px solid red; border-radius: 4px;">
          ❌ Failed to add calendar: ${error.message}
          <button onclick="showAddExternalCalendar()" style="margin-left: 10px;">Try Again</button>
        </div>
      `;
    }
  }
}

async function addExternalCalendarConnection(calendarData) {
  console.log(`🔗 Adding ${calendarData.type} calendar:`, calendarData);
  
  // For now, just store the connection info - we'll implement actual connections later
  const calendarInfo = {
    id: `${calendarData.type}-${Date.now()}`,
    type: calendarData.type,
    name: calendarData.name || `${calendarData.type} (${calendarData.email})`,
    email: calendarData.email,
    ...calendarData,
    status: 'pending_implementation',
    added: new Date().toISOString()
  };
  
  // Store in localStorage for now
  const existingCalendars = JSON.parse(localStorage.getItem('externalCalendars') || '[]');
  existingCalendars.push(calendarInfo);
  localStorage.setItem('externalCalendars', JSON.stringify(existingCalendars));
  
  console.log(`📅 ${calendarData.type} Calendar connection prepared:`, calendarInfo);
  
  return calendarInfo;
}
