# Calendar Sync Testing with Real Data Scenarios

## 🧪 **Functional Testing Plan**

### **Test Scenario 1: Single Day Sync**
**Goal**: Verify busy time sync across 3 calendars for one day

#### Test Setup
```
Calendar A (Primary - Company 1): 
- 9:00 AM - 10:30 AM: Team Meeting
- 2:00 PM - 3:00 PM: Client Call

Calendar B (Company 2): 
- Currently empty

Calendar C (Company 3):
- Currently empty
```

#### Expected Result
After sync, Calendar B and C should show:
- 9:00 AM - 10:30 AM: Busy (Private)
- 2:00 PM - 3:00 PM: Busy (Private)

---

## 🔧 **Create Test Data Generator**

Let's create a tool to generate realistic calendar scenarios:
