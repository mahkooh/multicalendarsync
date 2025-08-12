/*
 * Calendar Synchronization Manager
 * Handles the core logic for synchronizing busy time across multiple calendars
 */

/* global Office */

export class CalendarSyncManager {
  constructor() {
    this.syncEnabled = false;
    this.calendars = new Map();
    this.syncInterval = null;
    this.lastSync = null;
    this.syncInProgress = false;
    
    // Configuration
    this.config = {
      syncIntervalMinutes: 15, // Auto-sync every 15 minutes
      busyBlockSubject: '[Auto-Sync] Busy',
      busyBlockCategory: 'Auto-Sync',
      lookAheadDays: 0, // Sync only current day (0 days ahead)
      lookBehindDays: 0   // Sync only current day (0 days behind)
    };
  }

  async initialize() {
    try {
      console.log('Initializing Calendar Sync Manager...');
      
      // Check if Office is available
      if (!window.Office) {
        throw new Error('Office.js is not loaded');
      }
      
      // Check if we're in a supported context
      if (!Office.context || !Office.context.mailbox) {
        console.warn('Limited Office context detected, some features may not work');
      }
      
      console.log('Office context available:', !!Office.context);
      console.log('Mailbox available:', !!Office.context?.mailbox);
      
      // Request necessary permissions
      await this.requestPermissions();
      
      // Load configuration from local storage
      this.loadConfiguration();
      
      // Discover available calendars
      await this.discoverCalendars();
      
      // Start auto-sync if enabled
      if (this.syncEnabled) {
        this.startAutoSync();
      }
      
      console.log('Calendar Sync Manager initialized successfully');
      
    } catch (error) {
      console.error('Failed to initialize Calendar Sync Manager:', error);
      throw error;
    }
  }

  async requestPermissions() {
    return new Promise((resolve, reject) => {
      try {
        // Check if we have Office context and mailbox
        if (!Office.context || !Office.context.mailbox) {
          console.warn('Limited Office context, proceeding with basic initialization');
          resolve();
          return;
        }

        // Get diagnostic info
        const diagnostics = Office.context.mailbox.diagnostics;
        const hostName = diagnostics?.hostName;
        console.log('Host diagnostics:', {
          hostName: hostName,
          hostVersion: diagnostics?.hostVersion,
          platform: diagnostics?.platform
        });

        // More permissive host checking
        if (hostName === 'Outlook' || 
            hostName === 'OutlookWebApp' || 
            hostName === 'OutlookIOS' || 
            hostName === 'OutlookAndroid' ||
            !hostName) { // Allow undefined hostName as fallback
          console.log('Host validation passed for:', hostName || 'unknown host');
          resolve();
        } else {
          console.warn('Unknown host detected, proceeding anyway:', hostName);
          resolve(); // Don't reject, just proceed
        }
      } catch (error) {
        console.error('Error in requestPermissions:', error);
        // Don't fail completely, just proceed
        resolve();
      }
    });
  }

  loadConfiguration() {
    try {
      const savedConfig = localStorage.getItem('calendarSync_config');
      if (savedConfig) {
        const parsed = JSON.parse(savedConfig);
        this.config = { ...this.config, ...parsed };
      }
      
      const syncEnabled = localStorage.getItem('calendarSync_enabled');
      this.syncEnabled = syncEnabled === 'true';
      
    } catch (error) {
      console.warn('Failed to load configuration:', error);
      // Use defaults
    }
  }

  saveConfiguration() {
    try {
      localStorage.setItem('calendarSync_config', JSON.stringify(this.config));
      localStorage.setItem('calendarSync_enabled', this.syncEnabled.toString());
    } catch (error) {
      console.warn('Failed to save configuration:', error);
    }
  }

  async discoverCalendars() {
    try {
      // In a real implementation, this would use Microsoft Graph API
      // For now, we'll simulate multiple calendars
      
      const mockCalendars = [
        {
          id: 'calendar-1',
          name: 'Company A Calendar',
          type: 'Exchange',
          syncEnabled: true,
          itemCount: 12
        },
        {
          id: 'calendar-2',
          name: 'Company B Calendar',
          type: 'Office 365',
          syncEnabled: true,
          itemCount: 8
        },
        {
          id: 'calendar-3',
          name: 'Company C Calendar',
          type: 'Exchange',
          syncEnabled: false,
          itemCount: 5
        },
        {
          id: 'calendar-4',
          name: 'Personal Calendar',
          type: 'Outlook.com',
          syncEnabled: true,
          itemCount: 15
        },
        {
          id: 'calendar-5',
          name: 'Project Calendar',
          type: 'SharePoint',
          syncEnabled: false,
          itemCount: 3
        }
      ];

      this.calendars.clear();
      mockCalendars.forEach(cal => {
        this.calendars.set(cal.id, cal);
      });

      console.log(`Discovered ${this.calendars.size} calendars`);
      
    } catch (error) {
      console.error('Failed to discover calendars:', error);
      throw error;
    }
  }

  async getAvailableCalendars() {
    return Array.from(this.calendars.values());
  }

  async toggleCalendarSync(calendarId) {
    const calendar = this.calendars.get(calendarId);
    if (!calendar) {
      throw new Error('Calendar not found');
    }

    calendar.syncEnabled = !calendar.syncEnabled;
    this.saveConfiguration();
    
    console.log(`Calendar sync ${calendar.syncEnabled ? 'enabled' : 'disabled'} for ${calendar.name}`);
    
    return calendar.syncEnabled;
  }

  async performSync(targetDate = null) {
    if (this.syncInProgress) {
      throw new Error('Sync already in progress');
    }

    this.syncInProgress = true;
    
    try {
      // Use target date or default to current date range
      const syncDate = targetDate || new Date();
      console.log(`Starting calendar synchronization for ${syncDate.toLocaleDateString()}...`);
      
      // Calculate date range for single day sync
      const startDate = new Date(syncDate);
      startDate.setHours(0, 0, 0, 0); // Start of day
      
      const endDate = new Date(syncDate);
      endDate.setHours(23, 59, 59, 999); // End of day
      
      // Get all enabled calendars
      const enabledCalendars = Array.from(this.calendars.values())
        .filter(cal => cal.syncEnabled);

      if (enabledCalendars.length < 2) {
        throw new Error('At least 2 calendars must be enabled for synchronization');
      }

      // Get busy times from all calendars
      const busyTimes = await this.getBusyTimesFromCalendars(enabledCalendars);
      
      // Remove existing sync blocks
      const removedBlocks = await this.removeExistingSyncBlocks(enabledCalendars);
      
      // Create new sync blocks
      const createdBlocks = await this.createSyncBlocks(enabledCalendars, busyTimes);
      
      this.lastSync = new Date();
      
      console.log(`Sync completed: ${createdBlocks} blocks created, ${removedBlocks} blocks removed`);
      
      return {
        blocksCreated: createdBlocks,
        blocksRemoved: removedBlocks,
        syncTime: this.lastSync
      };
      
    } catch (error) {
      console.error('Sync failed:', error);
      throw error;
    } finally {
      this.syncInProgress = false;
    }
  }

  async getBusyTimesFromCalendars(calendars) {
    const busyTimes = new Map();
    
    for (const calendar of calendars) {
      try {
        // In a real implementation, this would query the actual calendar
        // For now, we'll simulate busy times
        const calendarBusyTimes = await this.simulateGetBusyTimes(calendar);
        busyTimes.set(calendar.id, calendarBusyTimes);
        
      } catch (error) {
        console.warn(`Failed to get busy times for ${calendar.name}:`, error);
        busyTimes.set(calendar.id, []);
      }
    }
    
    return busyTimes;
  }

  async simulateGetBusyTimes(calendar) {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Generate some mock busy times for demonstration
    const now = new Date();
    const busyTimes = [];
    
    // Add a few random busy slots over the next few days
    for (let i = 0; i < Math.floor(Math.random() * 5) + 1; i++) {
      const start = new Date(now.getTime() + Math.random() * 7 * 24 * 60 * 60 * 1000);
      const duration = (Math.random() * 3 + 0.5) * 60 * 60 * 1000; // 30min to 3.5 hours
      const end = new Date(start.getTime() + duration);
      
      busyTimes.push({
        start: start,
        end: end,
        subject: `Meeting in ${calendar.name}`,
        isPrivate: false
      });
    }
    
    return busyTimes;
  }

  async removeExistingSyncBlocks(calendars) {
    let removedCount = 0;
    
    for (const calendar of calendars) {
      try {
        // In a real implementation, this would find and delete sync blocks
        // For now, we'll simulate the removal
        const removed = await this.simulateRemoveSyncBlocks(calendar);
        removedCount += removed;
        
      } catch (error) {
        console.warn(`Failed to remove sync blocks from ${calendar.name}:`, error);
      }
    }
    
    return removedCount;
  }

  async simulateRemoveSyncBlocks(calendar) {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 50));
    
    // Simulate removing some blocks
    return Math.floor(Math.random() * 3);
  }

  async createSyncBlocks(calendars, busyTimes) {
    let createdCount = 0;
    
    // For each calendar, create busy blocks for times from other calendars
    for (const targetCalendar of calendars) {
      try {
        const blocksToCreate = [];
        
        // Collect busy times from all other calendars
        for (const [sourceCalendarId, times] of busyTimes) {
          if (sourceCalendarId !== targetCalendar.id) {
            blocksToCreate.push(...times);
          }
        }
        
        // Create the busy blocks
        const created = await this.simulateCreateSyncBlocks(targetCalendar, blocksToCreate);
        createdCount += created;
        
      } catch (error) {
        console.warn(`Failed to create sync blocks in ${targetCalendar.name}:`, error);
      }
    }
    
    return createdCount;
  }

  async simulateCreateSyncBlocks(calendar, busyTimes) {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // In a real implementation, this would create actual calendar events with:
    // - Subject: this.config.busyBlockSubject
    // - ShowAs: "Busy"
    // - IsPrivate: true (so they don't show in the user's view)
    // - Category: this.config.busyBlockCategory (for easy identification)
    
    console.log(`Would create ${busyTimes.length} sync blocks in ${calendar.name}`);
    
    return busyTimes.length;
  }

  startAutoSync() {
    if (this.syncInterval) {
      clearInterval(this.syncInterval);
    }
    
    const intervalMs = this.config.syncIntervalMinutes * 60 * 1000;
    this.syncInterval = setInterval(() => {
      this.performSync().catch(error => {
        console.error('Auto-sync failed:', error);
      });
    }, intervalMs);
    
    console.log(`Auto-sync started with ${this.config.syncIntervalMinutes} minute interval`);
  }

  stopAutoSync() {
    if (this.syncInterval) {
      clearInterval(this.syncInterval);
      this.syncInterval = null;
      console.log('Auto-sync stopped');
    }
  }

  async getSyncStatus() {
    if (this.syncInProgress) {
      return {
        state: 'Syncing',
        message: 'Synchronization in progress...',
        lastSync: this.lastSync
      };
    }
    
    if (!this.syncEnabled) {
      return {
        state: 'Stopped',
        message: 'Synchronization disabled',
        lastSync: this.lastSync
      };
    }
    
    const enabledCount = Array.from(this.calendars.values())
      .filter(cal => cal.syncEnabled).length;
      
    if (enabledCount < 2) {
      return {
        state: 'Error',
        message: 'Need at least 2 enabled calendars',
        lastSync: this.lastSync
      };
    }
    
    return {
      state: 'Active',
      message: `Monitoring ${enabledCount} calendars`,
      lastSync: this.lastSync
    };
  }

  // Configuration methods
  setSyncInterval(minutes) {
    this.config.syncIntervalMinutes = minutes;
    this.saveConfiguration();
    
    if (this.syncInterval) {
      this.startAutoSync(); // Restart with new interval
    }
  }

  setBusyBlockSubject(subject) {
    this.config.busyBlockSubject = subject;
    this.saveConfiguration();
  }

  setLookAheadDays(days) {
    this.config.lookAheadDays = days;
    this.saveConfiguration();
  }

  getConfiguration() {
    return { ...this.config };
  }
}
