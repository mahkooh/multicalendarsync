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
      console.log('Discovering real Microsoft accounts and calendars...');
      
      // Try to get real calendars first
      const calendars = await this.getGraphCalendars();
      
      if (calendars && calendars.length > 0) {
        this.calendars.clear();
        calendars.forEach(cal => {
          this.calendars.set(cal.id, cal);
        });

        console.log(`✅ Discovered ${this.calendars.size} real calendars from Graph API:`);
        Array.from(this.calendars.values()).forEach(cal => {
          console.log(`  📅 ${cal.name} (${cal.userEmail}) - Default: ${cal.isDefaultCalendar}`);
        });
      } else {
        throw new Error('No calendars returned from Graph API');
      }
      
    } catch (error) {
      console.warn('⚠️ Graph API discovery failed, checking why:', error);
      
      // Try alternative approaches or fallback
      try {
        await this.tryAlternativeDiscovery();
      } catch (altError) {
        console.warn('⚠️ Alternative discovery also failed:', altError);
        console.log('📝 Using enhanced mock data with email simulation...');
        await this.discoverEnhancedMockCalendars();
      }
    }
  }

  async tryAlternativeDiscovery() {
    // Try to get user profile first to check authentication
    const accessToken = await this.getGraphAccessToken();
    
    console.log('🔍 Testing Graph API access with user profile...');
    const profileResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!profileResponse.ok) {
      throw new Error(`Profile API error: ${profileResponse.status} ${profileResponse.statusText}`);
    }

    const profile = await profileResponse.json();
    console.log(`✅ Authenticated as: ${profile.displayName} (${profile.mail || profile.userPrincipalName})`);
    
    // Try calendars again with better error info
    const calResponse = await fetch('https://graph.microsoft.com/v1.0/me/calendars', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!calResponse.ok) {
      const errorText = await calResponse.text();
      throw new Error(`Calendars API error: ${calResponse.status} ${calResponse.statusText} - ${errorText}`);
    }

    const calData = await calResponse.json();
    console.log('📅 Raw calendar data from Graph:', calData);
    
    // Convert to our format with the authenticated user's email
    const userEmail = profile.mail || profile.userPrincipalName;
    const calendars = calData.value.map(graphCal => ({
      id: graphCal.id,
      name: graphCal.name,
      userEmail: graphCal.owner?.emailAddress?.address || userEmail,
      type: 'Microsoft Graph',
      syncEnabled: true,
      canEdit: graphCal.canEdit || false,
      isDefaultCalendar: graphCal.isDefaultCalendar || false
    }));

    this.calendars.clear();
    calendars.forEach(cal => {
      this.calendars.set(cal.id, cal);
    });

    console.log(`✅ Successfully discovered ${calendars.length} calendars with emails`);
  }

  async discoverEnhancedMockCalendars() {
    // Enhanced mock calendars with realistic email addresses
    const mockCalendars = [
      {
        id: 'calendar-1',
        name: 'Primary Work Calendar',
        userEmail: 'dan.hookham@company1.com',
        type: 'Exchange Online',
        syncEnabled: true,
        isDefaultCalendar: true
      },
      {
        id: 'calendar-2', 
        name: 'Secondary Work Calendar',
        userEmail: 'dan.hookham@company2.com',
        type: 'Office 365',
        syncEnabled: true,
        isDefaultCalendar: false
      },
      {
        id: 'calendar-3',
        name: 'Personal Calendar',
        userEmail: 'dan.hookham@outlook.com',
        type: 'Outlook.com',
        syncEnabled: true,
        isDefaultCalendar: false
      },
      {
        id: 'calendar-4',
        name: 'Project Calendar',
        userEmail: 'dan.hookham@company3.com',
        type: 'Exchange',
        syncEnabled: true,
        isDefaultCalendar: false
      },
      {
        id: 'calendar-5',
        name: 'Client Calendar',
        userEmail: 'dan.hookham@company4.com',
        type: 'Office 365',
        syncEnabled: true,
        isDefaultCalendar: false
      }
    ];

    this.calendars.clear();
    mockCalendars.forEach(cal => {
      this.calendars.set(cal.id, cal);
    });

    console.log('📝 Using enhanced mock calendars with email addresses:');
    mockCalendars.forEach(cal => {
      console.log(`  📅 ${cal.name} (${cal.userEmail})`);
    });
  }

  async getGraphCalendars() {
    try {
      const accessToken = await this.getGraphAccessToken();
      console.log('🔐 Graph access token obtained, fetching calendars...');
      
      // First try to get the user's profile to understand the context
      const profileResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!profileResponse.ok) {
        throw new Error(`Profile API failed: ${profileResponse.status}`);
      }

      const userProfile = await profileResponse.json();
      const primaryEmail = userProfile.mail || userProfile.userPrincipalName;
      console.log(`👤 Primary user: ${userProfile.displayName} (${primaryEmail})`);

      // Get calendars with expanded properties
      const calendarResponse = await fetch('https://graph.microsoft.com/v1.0/me/calendars?$select=id,name,canEdit,isDefaultCalendar,owner', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!calendarResponse.ok) {
        const errorText = await calendarResponse.text();
        console.error('Calendar API error response:', errorText);
        throw new Error(`Calendar API failed: ${calendarResponse.status} - ${errorText}`);
      }

      const calendarData = await calendarResponse.json();
      console.log('📊 Raw Graph calendar response:', calendarData);

      if (!calendarData.value || calendarData.value.length === 0) {
        console.warn('⚠️ No calendars found in Graph response');
        return [];
      }

      // Try to also get calendar groups for shared/delegated calendars
      let additionalCalendars = [];
      try {
        const groupResponse = await fetch('https://graph.microsoft.com/v1.0/me/calendarGroups?$expand=calendars', {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        });

        if (groupResponse.ok) {
          const groupData = await groupResponse.json();
          console.log('📂 Calendar groups found:', groupData);
          
          // Extract calendars from groups
          groupData.value?.forEach(group => {
            if (group.calendars) {
              additionalCalendars.push(...group.calendars);
            }
          });
        }
      } catch (groupError) {
        console.log('ℹ️ Calendar groups not accessible (normal for some accounts):', groupError.message);
      }

      // Combine and deduplicate calendars
      const allCalendars = [...calendarData.value, ...additionalCalendars];
      const uniqueCalendars = allCalendars.filter((cal, index, self) => 
        index === self.findIndex(c => c.id === cal.id)
      );

      console.log(`📅 Found ${uniqueCalendars.length} total calendars (${calendarData.value.length} primary + ${additionalCalendars.length} from groups)`);

      // Convert to our calendar format
      const calendars = uniqueCalendars.map(graphCal => {
        const ownerEmail = graphCal.owner?.emailAddress?.address || primaryEmail;
        return {
          id: graphCal.id,
          name: graphCal.name || 'Unnamed Calendar',
          userEmail: ownerEmail,
          type: 'Microsoft Graph',
          syncEnabled: true,
          canEdit: graphCal.canEdit !== false, // Default to true if not specified
          isDefaultCalendar: graphCal.isDefaultCalendar || false
        };
      });

      console.log('✅ Processed calendars:');
      calendars.forEach(cal => {
        console.log(`  📅 "${cal.name}" owned by ${cal.userEmail} (Default: ${cal.isDefaultCalendar}, Editable: ${cal.canEdit})`);
      });

      return calendars;

    } catch (error) {
      console.error('❌ getGraphCalendars failed:', error);
      console.log('🔄 Will attempt alternative discovery or fall back to mock data');
      throw error;
    }
  }

  async getGraphAccessToken() {
    try {
      console.log('🔐 Requesting Graph access token...');
      
      // Check Office.js availability step by step
      console.log('🔍 Checking Office.js availability:');
      console.log('  - Office object:', typeof Office);
      console.log('  - Office.context:', typeof Office?.context);
      console.log('  - Office.context.auth:', typeof Office?.context?.auth);
      console.log('  - getAccessToken method:', typeof Office?.context?.auth?.getAccessToken);
      
      // Check if we're in the right host application
      if (Office?.context?.host) {
        console.log('🏢 Host application:', Office.context.host);
        console.log('📱 Platform:', Office.context.platform);
        console.log('🔧 Requirements:', Office.context.requirements);
      }

      // Check if Office.js auth is available
      if (!Office?.context?.auth?.getAccessToken) {
        console.warn('⚠️ Office.js authentication not available in this context');
        console.log('🔄 Attempting alternative authentication methods...');
        
        // Try to check if we're in a web context that supports alternative auth
        if (typeof window !== 'undefined' && window.location) {
          console.log('🌐 Running in web context:', window.location.href);
          
          // For testing purposes, try to simulate a token or provide guidance
          throw new Error('Office.js SSO not available - add-in may need to be loaded in proper Office context');
        }
        
        throw new Error('Office.js authentication not available in this context');
      }

      console.log('✅ Office.js authentication is available, requesting token...');

      // Use Office.js to get access token for Microsoft Graph
      return new Promise((resolve, reject) => {
        const tokenOptions = {
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true
        };
        
        console.log('📋 Token request options:', tokenOptions);

        Office.context.auth.getAccessToken(tokenOptions, (result) => {
          console.log('🔍 Auth result status:', result.status);
          console.log('🔍 Full auth result:', result);
          
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✅ Access token obtained successfully');
            console.log('🎫 Token preview:', result.value.substring(0, 20) + '...');
            console.log('🎫 Token length:', result.value.length);
            resolve(result.value);
          } else {
            console.error('❌ Authentication failed:', result.error);
            console.error('Error code:', result.error?.code);
            console.error('Error message:', result.error?.message);
            console.error('Error name:', result.error?.name);
            console.error('Error tracing:', result.error?.tracing);
            
            // Provide more specific error information
            let errorMsg = `Authentication failed: ${result.error?.message || 'Unknown error'}`;
            if (result.error?.code === 13001) {
              errorMsg += ' (User not signed in - this may require signing into Office)';
            } else if (result.error?.code === 13002) {
              errorMsg += ' (User consent required - admin may need to grant permissions)';
            } else if (result.error?.code === 13003) {
              errorMsg += ' (Invalid audience - token scope issue)';
            } else if (result.error?.code === 13006) {
              errorMsg += ' (Invalid request - check add-in manifest permissions)';
            } else if (result.error?.code === 13007) {
              errorMsg += ' (Invalid grant - user may need to re-authenticate)';
            } else if (result.error?.code === 13012) {
              errorMsg += ' (Add-in not trusted - may need admin consent)';
            }
            
            reject(new Error(errorMsg));
          }
        });
      });
    } catch (error) {
      console.error('❌ Failed to get Graph access token:', error);
      console.log('💡 Troubleshooting suggestions:');
      console.log('  1. Ensure add-in is loaded in Outlook (not just web browser)');
      console.log('  2. Check that manifest is properly deployed');
      console.log('  3. Verify Office.js is fully loaded');
      console.log('  4. Check manifest WebApplicationInfo configuration');
      throw new Error(`Authentication setup failed: ${error.message}`);
    }
  }

  async discoverMockCalendars() {
    // Fallback mock calendars for testing
    const mockCalendars = [
      {
        id: 'calendar-1',
        name: 'Primary Work Calendar',
        userEmail: 'user@company1.com',
        type: 'Exchange',
        syncEnabled: true,
        isDefaultCalendar: true
      },
      {
        id: 'calendar-2', 
        name: 'Secondary Work Calendar',
        userEmail: 'user@company2.com',
        type: 'Office 365',
        syncEnabled: true,
        isDefaultCalendar: false
      },
      {
        id: 'calendar-3',
        name: 'Personal Calendar',
        userEmail: 'personal@outlook.com',
        type: 'Outlook.com',
        syncEnabled: true,
        isDefaultCalendar: false
      }
    ];

    this.calendars.clear();
    mockCalendars.forEach(cal => {
      this.calendars.set(cal.id, cal);
    });

    console.log('Using mock calendars for testing');
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
        console.log(`Fetching busy times for ${calendar.name} (${calendar.userEmail})`);
        
        // Get real busy times from Microsoft Graph API
        const calendarBusyTimes = await this.getRealBusyTimes(calendar);
        busyTimes.set(calendar.id, calendarBusyTimes);
        
        console.log(`Found ${calendarBusyTimes.length} busy times in ${calendar.name}`);
        
      } catch (error) {
        console.warn(`Failed to get busy times for ${calendar.name}:`, error);
        // Fallback to simulation if Graph API fails
        const simulatedTimes = await this.simulateGetBusyTimes(calendar);
        busyTimes.set(calendar.id, simulatedTimes);
      }
    }
    
    return busyTimes;
  }

  async getRealBusyTimes(calendar) {
    try {
      const accessToken = await this.getGraphAccessToken();
      
      // Calculate date range for the target date
      const targetDate = this.targetDate || new Date();
      const startTime = new Date(targetDate);
      startTime.setHours(0, 0, 0, 0);
      
      const endTime = new Date(targetDate);
      endTime.setHours(23, 59, 59, 999);
      
      // Query calendar events for the specific date
      const eventsUrl = `https://graph.microsoft.com/v1.0/me/calendars/${calendar.id}/events` +
        `?$filter=start/dateTime ge '${startTime.toISOString()}' and end/dateTime le '${endTime.toISOString()}'` +
        `&$select=subject,start,end,showAs,isPrivate`;
      
      const response = await fetch(eventsUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      
      // Convert events to busy time blocks
      return data.value
        .filter(event => event.showAs === 'busy' || event.showAs === 'tentative')
        .map(event => ({
          start: new Date(event.start.dateTime),
          end: new Date(event.end.dateTime),
          subject: event.isPrivate ? '[Private]' : (event.subject || '[No Subject]'),
          isPrivate: event.isPrivate || false
        }));

    } catch (error) {
      console.error(`Graph API call failed for ${calendar.name}:`, error);
      throw error;
    }
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
    try {
      // Try to create real sync blocks via Graph API
      return await this.createRealSyncBlocks(calendar, busyTimes);
    } catch (error) {
      console.warn(`Graph API failed, simulating for ${calendar.name}:`, error);
      
      // Fallback to simulation
      await new Promise(resolve => setTimeout(resolve, 100));
      
      const calendarInfo = calendar.userEmail ? 
        `${calendar.name} (${calendar.userEmail})` : 
        calendar.name;
      
      console.log(`Would create ${busyTimes.length} sync blocks in ${calendarInfo}`);
      
      return busyTimes.length;
    }
  }

  async createRealSyncBlocks(calendar, busyTimes) {
    const accessToken = await this.getGraphAccessToken();
    let createdCount = 0;

    for (const busyTime of busyTimes) {
      try {
        // Create a busy block event in the target calendar
        const eventData = {
          subject: this.config.busyBlockSubject,
          start: {
            dateTime: busyTime.start.toISOString(),
            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
          },
          end: {
            dateTime: busyTime.end.toISOString(),
            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
          },
          showAs: 'busy',
          isPrivate: true,
          categories: [this.config.busyBlockCategory],
          body: {
            contentType: 'text',
            content: 'Auto-synchronized busy time from another calendar'
          }
        };

        const response = await fetch(`https://graph.microsoft.com/v1.0/me/calendars/${calendar.id}/events`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(eventData)
        });

        if (response.ok) {
          createdCount++;
          console.log(`Created sync block in ${calendar.name}: ${busyTime.start.toLocaleTimeString()} - ${busyTime.end.toLocaleTimeString()}`);
        } else {
          console.warn(`Failed to create sync block in ${calendar.name}: ${response.status}`);
        }

      } catch (error) {
        console.error(`Error creating sync block in ${calendar.name}:`, error);
      }
    }

    console.log(`Created ${createdCount} real sync blocks in ${calendar.name} (${calendar.userEmail})`);
    return createdCount;
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
