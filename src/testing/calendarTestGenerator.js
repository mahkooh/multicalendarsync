// Calendar Sync Test Data Generator and Simulator
// This file helps test the calendar synchronization logic with realistic scenarios

class CalendarTestDataGenerator {
    constructor() {
        this.testScenarios = [];
        this.mockCalendars = new Map();
    }

    // Generate a day's worth of test appointments
    generateDayScenario(date = new Date()) {
        const scenarios = [
            {
                name: "Typical Business Day",
                calendars: {
                    "work-company1": [
                        { start: "09:00", end: "10:30", title: "Team Standup", location: "Conference Room A" },
                        { start: "11:00", end: "12:00", title: "Project Review", location: "Online" },
                        { start: "14:00", end: "15:30", title: "Client Presentation", location: "Boardroom" },
                        { start: "16:00", end: "17:00", title: "1-on-1 with Manager", location: "Office" }
                    ],
                    "work-company2": [],
                    "personal": [
                        { start: "12:00", end: "13:00", title: "Lunch with Friend", location: "Restaurant" },
                        { start: "18:00", end: "19:30", title: "Gym Session", location: "Fitness Center" }
                    ]
                }
            },
            {
                name: "Meeting Heavy Day",
                calendars: {
                    "work-company1": [
                        { start: "08:30", end: "09:30", title: "Early Team Meeting", location: "Conference Room" },
                        { start: "10:00", end: "11:00", title: "Vendor Call", location: "Phone" },
                        { start: "13:00", end: "14:00", title: "Product Demo", location: "Demo Room" },
                        { start: "15:00", end: "16:00", title: "Strategy Session", location: "Boardroom" }
                    ],
                    "work-company2": [
                        { start: "09:30", end: "10:00", title: "Quick Sync", location: "Online" },
                        { start: "14:30", end: "15:30", title: "Technical Review", location: "Lab" }
                    ],
                    "personal": [
                        { start: "19:00", end: "20:00", title: "Family Dinner", location: "Home" }
                    ]
                }
            },
            {
                name: "Conflict Resolution Day",
                calendars: {
                    "work-company1": [
                        { start: "10:00", end: "11:00", title: "Important Meeting", location: "HQ" },
                        { start: "14:00", end: "15:00", title: "Client Call", location: "Phone" }
                    ],
                    "work-company2": [
                        { start: "10:30", end: "11:30", title: "Overlapping Meeting", location: "Branch Office" },
                        { start: "14:15", end: "14:45", title: "Quick Check-in", location: "Online" }
                    ],
                    "personal": []
                }
            },
            {
                name: "August 12th 2025 Real Schedule",
                calendars: {
                    "work-company1": [
                        { start: "07:30", end: "08:30", title: "Morning Team Standup", location: "Conference Room B" },
                        { start: "10:00", end: "11:30", title: "Strategic Planning Session", location: "Boardroom" },
                        { start: "13:30", end: "14:30", title: "Client Presentation", location: "Online Meeting" },
                        { start: "16:00", end: "17:00", title: "Project Review", location: "Office" }
                    ],
                    "work-company2": [
                        { start: "08:00", end: "09:00", title: "Cross-functional Meeting", location: "Branch Office" },
                        { start: "11:00", end: "12:00", title: "Technical Deep Dive", location: "Lab" },
                        { start: "15:00", end: "16:00", title: "Stakeholder Update", location: "Conference Call" }
                    ],
                    "work-company3": [
                        { start: "09:30", end: "10:30", title: "Vendor Negotiation", location: "Meeting Room 3" },
                        { start: "14:00", end: "15:00", title: "Product Demo", location: "Demo Center" },
                        { start: "17:30", end: "18:30", title: "International Team Call", location: "Online" }
                    ],
                    "personal": [
                        { start: "07:00", end: "07:30", title: "Morning Workout", location: "Home Gym" },
                        { start: "12:00", end: "13:00", title: "Lunch Meeting", location: "Local Café" },
                        { start: "18:30", end: "19:30", title: "Family Time", location: "Home" },
                        { start: "20:00", end: "21:00", title: "Evening Walk", location: "Park" }
                    ],
                    "work-consulting": [
                        { start: "08:30", end: "09:30", title: "Client Check-in Call", location: "Phone" },
                        { start: "12:30", end: "13:30", title: "Project Status Review", location: "Client Office" },
                        { start: "19:00", end: "20:00", title: "Evening Consultation", location: "Online Meeting" }
                    ]
                }
            }
        ];

        return scenarios;
    }

    // Convert time strings to Date objects for a given date
    parseTimeSlot(timeStr, baseDate) {
        const [hours, minutes] = timeStr.split(':').map(Number);
        const date = new Date(baseDate);
        date.setHours(hours, minutes, 0, 0);
        return date;
    }

    // Create mock calendar events for testing
    createMockCalendar(calendarId, events, date = new Date()) {
        const mockEvents = events.map((event, index) => ({
            id: `${calendarId}-event-${index}`,
            subject: event.title,
            start: {
                dateTime: this.parseTimeSlot(event.start, date).toISOString(),
                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
            },
            end: {
                dateTime: this.parseTimeSlot(event.end, date).toISOString(),
                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
            },
            location: {
                displayName: event.location
            },
            showAs: "busy",
            sensitivity: "normal"
        }));

        this.mockCalendars.set(calendarId, mockEvents);
        return mockEvents;
    }

    // Test the sync logic with mock data
    async testSyncScenario(scenarioName, targetDate = new Date()) {
        console.log(`\n🧪 Testing Scenario: ${scenarioName}`);
        console.log(`📅 Date: ${targetDate.toDateString()}\n`);

        const scenarios = this.generateDayScenario(targetDate);
        const scenario = scenarios.find(s => s.name === scenarioName);

        if (!scenario) {
            console.error(`❌ Scenario "${scenarioName}" not found`);
            return false;
        }

        // Create mock calendars with test data
        const mockCalendars = {};
        for (const [calendarId, events] of Object.entries(scenario.calendars)) {
            mockCalendars[calendarId] = this.createMockCalendar(calendarId, events, targetDate);
        }

        // Simulate the sync process
        console.log("📊 Original Calendar States:");
        this.displayCalendarState(mockCalendars);

        // Perform sync simulation
        const syncResults = await this.simulateSync(mockCalendars);

        console.log("\n✅ After Sync:");
        this.displayCalendarState(syncResults.calendars);

        console.log("\n📈 Sync Summary:");
        console.log(`- Busy blocks created: ${syncResults.busyBlocksCreated}`);
        console.log(`- Conflicts detected: ${syncResults.conflicts.length}`);
        console.log(`- Calendars updated: ${syncResults.calendarsUpdated.length}`);

        if (syncResults.conflicts.length > 0) {
            console.log("\n⚠️ Conflicts Found:");
            syncResults.conflicts.forEach(conflict => {
                console.log(`  - ${conflict.time}: ${conflict.description}`);
            });
        }

        return syncResults;
    }

    // Simulate the calendar sync logic
    async simulateSync(mockCalendars) {
        const results = {
            calendars: { ...mockCalendars },
            busyBlocksCreated: 0,
            conflicts: [],
            calendarsUpdated: []
        };

        // Collect all busy times from all calendars
        const allBusyTimes = [];
        for (const [calendarId, events] of Object.entries(mockCalendars)) {
            events.forEach(event => {
                allBusyTimes.push({
                    calendarId,
                    start: new Date(event.start.dateTime),
                    end: new Date(event.end.dateTime),
                    originalEvent: event
                });
            });
        }

        // Sort by start time
        allBusyTimes.sort((a, b) => a.start - b.start);

        // For each calendar, add busy blocks for events from other calendars
        for (const [targetCalendarId, events] of Object.entries(mockCalendars)) {
            const otherCalendarEvents = allBusyTimes.filter(bt => bt.calendarId !== targetCalendarId);
            
            for (const busyTime of otherCalendarEvents) {
                // Check if this time slot conflicts with existing events
                const hasConflict = events.some(existingEvent => {
                    const existingStart = new Date(existingEvent.start.dateTime);
                    const existingEnd = new Date(existingEvent.end.dateTime);
                    return (busyTime.start < existingEnd && busyTime.end > existingStart);
                });

                if (hasConflict) {
                    results.conflicts.push({
                        time: `${busyTime.start.toLocaleTimeString()} - ${busyTime.end.toLocaleTimeString()}`,
                        description: `Conflict in ${targetCalendarId} with event from ${busyTime.calendarId}`
                    });
                } else {
                    // Create a busy block
                    const busyBlock = {
                        id: `sync-busy-${Date.now()}-${Math.random()}`,
                        subject: "Busy",
                        start: {
                            dateTime: busyTime.start.toISOString(),
                            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
                        },
                        end: {
                            dateTime: busyTime.end.toISOString(),
                            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
                        },
                        showAs: "busy",
                        sensitivity: "private",
                        isQuickMeetingEmpty: true,
                        location: { displayName: "" },
                        body: { content: "Automatically created busy time block" }
                    };

                    results.calendars[targetCalendarId].push(busyBlock);
                    results.busyBlocksCreated++;
                    
                    if (!results.calendarsUpdated.includes(targetCalendarId)) {
                        results.calendarsUpdated.push(targetCalendarId);
                    }
                }
            }
        }

        return results;
    }

    // Display calendar state in a readable format
    displayCalendarState(calendars) {
        for (const [calendarId, events] of Object.entries(calendars)) {
            console.log(`\n📋 ${calendarId.toUpperCase()}:`);
            if (events.length === 0) {
                console.log("   (No events)");
            } else {
                const sortedEvents = events.sort((a, b) => 
                    new Date(a.start.dateTime) - new Date(b.start.dateTime)
                );
                
                sortedEvents.forEach(event => {
                    const start = new Date(event.start.dateTime);
                    const end = new Date(event.end.dateTime);
                    const timeStr = `${start.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})} - ${end.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
                    const icon = event.sensitivity === 'private' ? '🔒' : '📅';
                    console.log(`   ${icon} ${timeStr}: ${event.subject}`);
                });
            }
        }
    }

    // Run comprehensive test suite
    async runTestSuite() {
        console.log("🚀 Starting Calendar Sync Test Suite\n");
        
        const testDate = new Date();
        const scenarios = ["Typical Business Day", "Meeting Heavy Day", "Conflict Resolution Day"];
        
        for (const scenario of scenarios) {
            await this.testSyncScenario(scenario, testDate);
            console.log("\n" + "=".repeat(60));
        }
        
        console.log("\n✅ Test Suite Complete!");
    }
}

// Export for use in testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CalendarTestDataGenerator;
}

// Browser/testing usage
if (typeof window !== 'undefined') {
    window.CalendarTestDataGenerator = CalendarTestDataGenerator;
}

// Example usage for immediate testing
if (typeof window === 'undefined' && require.main === module) {
    const tester = new CalendarTestDataGenerator();
    tester.runTestSuite();
}
