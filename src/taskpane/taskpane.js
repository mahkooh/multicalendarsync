/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { CalendarSyncManager } from './calendarSync.js';

let syncManager;

Office.onReady((info) => {
  try {
    console.log('Office.onReady called with info:', info);
    
    // More permissive host check
    if (info.host === Office.HostType.Outlook || info.host === 'Outlook' || window.Office) {
      console.log('Host validated as Outlook');
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      
      // Set default date to August 12, 2025 for testing
      const defaultDate = new Date('2025-08-12');
      document.getElementById("sync-date").value = defaultDate.toISOString().split('T')[0];
      
      // Initialize the calendar sync manager
      syncManager = new CalendarSyncManager();
      initializeApp();
    } else {
      console.error('Unsupported host application. Host info:', info);
      // Show error but don't fail completely
      document.getElementById("sideload-msg").innerHTML = `
        <div style="color: red; padding: 10px;">
          <h3>Host Detection Issue</h3>
          <p>Detected host: ${info.host}</p>
          <p>Expected: Outlook</p>
          <p>Trying to continue anyway...</p>
        </div>
      `;
      
      // Try to initialize anyway
      syncManager = new CalendarSyncManager();
      initializeApp();
    }
  } catch (error) {
    console.error('Office.onReady error:', error);
    document.getElementById("sideload-msg").innerHTML = `
      <div style="color: red; padding: 10px;">
        <h3>Initialization Error</h3>
        <p>${error.message}</p>
        <p>Please try refreshing the page</p>
      </div>
    `;
  }
});

// Fallback initialization if Office.onReady doesn't work
window.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (!syncManager && window.Office) {
      console.log('Fallback initialization triggered');
      try {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        
        const defaultDate = new Date('2025-08-12');
        document.getElementById("sync-date").value = defaultDate.toISOString().split('T')[0];
        
        syncManager = new CalendarSyncManager();
        initializeApp();
      } catch (error) {
        console.error('Fallback initialization failed:', error);
      }
    }
  }, 2000); // Wait 2 seconds for Office.js to load
});

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
