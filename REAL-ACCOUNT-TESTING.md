# 🔐 Real Account Testing Guide
## MultiCalendar Sync with 5 Real Microsoft Accounts

### 📋 **Pre-Testing Checklist**

#### **Account Setup Requirements**
- [ ] 5 Microsoft accounts with Outlook/Exchange access
- [ ] Each account has Microsoft Graph API permissions
- [ ] Accounts represent different organizations/contexts
- [ ] Test appointments scheduled for August 12th, 2025

**📅 Date Range Note:** The app has been configured to sync **one day at a time** for precise control. You can select any specific date to sync, including August 12th, 2025 for testing!

#### **Safety Precautions**
- [ ] **Backup existing calendars** before testing
- [ ] **Use test appointments** (not critical meetings)
- [ ] **Monitor sync behavior** during initial tests
- [ ] **Have rollback plan** ready

---

## 🧪 **Real Account Testing Strategy**

### **Phase 1: Account Authentication Setup**

#### **Step 1: Configure Account Access**
1. **Primary Account** (Main Outlook)
   - Your primary work/personal account
   - This will be the "master" account for sync operations

2. **Secondary Accounts** (Companies 1-4)
   - Company 1: [Your Company 1 Account]
   - Company 2: [Your Company 2 Account] 
   - Company 3: [Your Company 3 Account]
   - Company 4: [Your Company 4 Account]

#### **Step 2: Sideload the Add-in**
```xml
<!-- Use the corrected manifest.xml from your dist/ folder -->
1. Open Outlook Web App or Desktop
2. Go to Get Add-ins > My add-ins > Upload from file
3. Upload: dist/manifest.xml
4. Verify add-in appears in ribbon
```

### **Phase 2: Create Test Appointments**

#### **August 12th, 2025 Test Schedule**
Create these appointments in your real accounts:

**🏢 Company 1 Account:**
```
07:30-08:30: Team Morning Standup (Conference Room B)
10:00-11:30: Strategic Planning Session (Boardroom)
13:30-14:30: Client Presentation (Online)
16:00-17:00: Project Review (Office)
```

**🏭 Company 2 Account:**
```
08:00-09:00: Cross-functional Meeting (Branch Office)
11:00-12:00: Technical Deep Dive (Lab)
15:00-16:00: Stakeholder Update (Conference Call)
```

**🏪 Company 3 Account:**
```
09:30-10:30: Vendor Negotiation (Meeting Room 3)
14:00-15:00: Product Demo (Demo Center)
17:30-18:30: International Team Call (Online)
```

**💼 Company 4 Account:**
```
08:30-09:30: Client Check-in Call (Phone)
12:30-13:30: Project Status Review (Client Office)
19:00-20:00: Evening Consultation (Online)
```

**👤 Personal Account:**
```
07:00-07:30: Morning Workout (Home Gym)
12:00-13:00: Lunch Meeting (Local Café)
18:30-19:30: Family Time (Home)
20:00-21:00: Evening Walk (Park)
```

---

## 🚀 **Testing Execution Plan**

### **Phase 3: Sync Testing**

#### **Test 1: Single Day Sync**
1. **Open the add-in** in your primary account
2. **Select August 12, 2025** from the date picker
3. **Authenticate all 5 accounts** using the add-in interface
4. **Click "Sync Selected Date"**
5. **Verify busy blocks** appear in all calendars for that specific day only

#### **Test 2: Conflict Detection**
1. **Check for overlapping appointments** (e.g., 10:00-11:30 overlap)
2. **Verify conflict warnings** in the add-in
3. **Confirm busy blocks** respect privacy settings

#### **Test 3: Real-time Updates**
1. **Add a new appointment** in one account
2. **Run sync again**
3. **Verify new busy block** appears in other calendars

### **Phase 4: Validation Checklist**

#### **✅ Success Criteria**
- [ ] All 5 accounts successfully authenticated
- [ ] Busy times from one calendar appear as "Busy" in others
- [ ] Private events show as busy without revealing details
- [ ] No duplicate appointments created
- [ ] Original appointments remain unchanged
- [ ] Sync completes within reasonable time (< 2 minutes)

#### **⚠️ Issues to Watch For**
- [ ] Authentication failures
- [ ] Duplicate busy blocks
- [ ] Missing busy blocks
- [ ] Privacy leaks (showing private event details)
- [ ] Sync errors or timeouts
- [ ] Calendar corruption

---

## 🔧 **Troubleshooting Guide**

### **Common Issues & Solutions**

#### **Authentication Problems**
```javascript
// If auth fails, check:
1. Account permissions in Azure AD
2. Microsoft Graph API access
3. Add-in manifest permissions
4. Network connectivity
```

#### **Sync Conflicts**
```javascript
// If conflicts occur:
1. Check appointment overlap times
2. Verify time zone settings
3. Review privacy settings
4. Check for calendar permissions
```

#### **Missing Busy Blocks**
```javascript
// If busy blocks don't appear:
1. Verify sync completed successfully
2. Check target calendar permissions
3. Refresh calendar view
4. Re-run sync operation
```

---

## 📊 **Testing Results Template**

### **Account Authentication Results**
| Account | Status | Auth Time | Issues |
|---------|--------|-----------|---------|
| Primary | ✅/❌ | __:__ | _______ |
| Company 1 | ✅/❌ | __:__ | _______ |
| Company 2 | ✅/❌ | __:__ | _______ |
| Company 3 | ✅/❌ | __:__ | _______ |
| Company 4 | ✅/❌ | __:__ | _______ |

### **Sync Operation Results**
| Test | Expected | Actual | Status | Notes |
|------|----------|--------|---------|-------|
| Initial Sync | 14 busy blocks | __ | ✅/❌ | _____ |
| Conflict Detection | 3 conflicts | __ | ✅/❌ | _____ |
| Privacy Protection | Private events hidden | __ | ✅/❌ | _____ |
| Real-time Update | New block appears | __ | ✅/❌ | _____ |

### **Performance Metrics**
- **Total Sync Time**: _____ seconds
- **Appointments Processed**: _____ events
- **Busy Blocks Created**: _____ blocks
- **Conflicts Detected**: _____ conflicts
- **Errors Encountered**: _____ errors

---

## 🔒 **Security & Privacy Notes**

### **Data Protection**
- ✅ Add-in only reads calendar availability (free/busy)
- ✅ Private event details are not exposed
- ✅ No calendar data is stored externally
- ✅ All operations use Microsoft Graph API security

### **Rollback Procedures**
If testing causes issues:
1. **Delete generated busy blocks** manually
2. **Disconnect add-in** from accounts
3. **Restore from backup** if necessary
4. **Review error logs** for debugging

---

## 📝 **Next Steps After Testing**

### **If Tests Pass:**
1. **Document successful configuration**
2. **Create user guide** for other users
3. **Consider production deployment**
4. **Monitor ongoing performance**

### **If Tests Fail:**
1. **Review error logs** and debug issues
2. **Adjust sync logic** as needed
3. **Re-test with corrected code**
4. **Update manifest** if permissions needed

### **Production Readiness:**
1. **Submit to Microsoft AppSource** (optional)
2. **Deploy to organization** via admin center
3. **Train end users** on the tool
4. **Monitor usage** and feedback

---

*Remember: Always test in a controlled environment before using with critical calendar data!*
