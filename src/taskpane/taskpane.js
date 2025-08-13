/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { CalendarSyncManager } from './calendarSync.js';

let syncManager;

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
    // Set up event listeners
    document.getElementById("sync-now").onclick = handleSyncNow;
    document.getElementById("configure").onclick = handleConfigure;
    
    // Initialize the sync manager
    await syncManager.initialize();
    
    // Load calendars and status
    await loadCalendars();
    await updateSyncStatus();
    
    // Set up periodic status updates
    setInterval(updateSyncStatus, 30000); // Update every 30 seconds
    
  } catch (error) {
    console.error('Failed to initialize app:', error);
    updateStatus('Error', `Initialization failed: ${error.message}`);
  }
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
