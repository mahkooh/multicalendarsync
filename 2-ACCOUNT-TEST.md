# 🚀 Quick Start: Testing with 2 Accounts

## ✅ **Ready to Test!**
Your manifest is now valid and ready for deployment. Let's start with a simple 2-account test.

### **Phase 1: Sideload the Add-in**

#### **Step 1: Open Outlook Web App**
1. Go to https://outlook.office.com
2. Sign in with your **primary account**
3. Navigate to **Settings** (gear icon) → **View all Outlook settings**
4. Go to **General** → **Manage add-ins**
5. Click **+ Add a custom add-in** → **Add from file**
6. Upload: `dist/manifest.xml` from your project folder

#### **Step 2: Verify Installation**
- Look for the **"Calendar Sync"** button in your Outlook ribbon
- The button should appear when you're viewing calendar or composing appointments

### **Phase 2: Set Up Test Accounts**

#### **Account Setup**
- **Account 1**: Your primary Outlook account (already logged in)
- **Account 2**: Your secondary company/personal account

#### **Create Test Appointments for August 12th, 2025**

**Account 1 (Primary):**
```
09:00-10:30: Team Meeting
14:00-15:30: Client Call
16:00-17:00: Project Review
```

**Account 2 (Secondary):**
```
10:00-11:00: Department Sync
13:30-14:30: Vendor Meeting
17:30-18:30: Training Session
```

### **Phase 3: Test the Sync**

#### **Step 1: Open the Add-in**
1. Click the **"Calendar Sync"** button in Outlook
2. The task pane should open on the right side

#### **Step 2: Authenticate Both Accounts**
1. Click **"Add Account"** for Account 1 (should auto-detect current account)
2. Click **"Add Account"** for Account 2 (will prompt for login)
3. Grant calendar permissions when prompted

#### **Step 3: Run Sync Test**
1. Set date to **August 12, 2025**
2. Click **"Sync Calendars"**
3. Watch for sync progress and completion

### **Phase 4: Verify Results**

#### **Expected Outcome:**
- **Account 1** should show:
  - Original: Team Meeting, Client Call, Project Review
  - **New Busy Blocks**: Department Sync (busy), Vendor Meeting (busy), Training Session (busy)

- **Account 2** should show:
  - Original: Department Sync, Vendor Meeting, Training Session  
  - **New Busy Blocks**: Team Meeting (busy), Client Call (busy), Project Review (busy)

#### **Success Indicators:**
✅ Both accounts authenticated successfully
✅ Sync completed without errors  
✅ Busy blocks appeared in both calendars
✅ Original appointments remained unchanged
✅ Private event details not exposed

### **Phase 5: Conflict Testing**

#### **Create an Overlap:**
1. Add a meeting in Account 1: **10:15-10:45: Quick Standup**
2. Re-run the sync
3. Check for conflict warnings in the add-in interface

### **🔧 Troubleshooting Quick Fixes**

#### **If Add-in Doesn't Appear:**
- Refresh Outlook
- Check if manifest uploaded correctly
- Try uploading again

#### **If Authentication Fails:**
- Ensure both accounts have calendar access
- Check browser pop-up blockers
- Try incognito/private browsing mode

#### **If Sync Doesn't Work:**
- Check network connectivity
- Verify calendar permissions
- Look for error messages in the add-in interface

### **📊 Test Results Template**

**2-Account Test Results:**
- [ ] Add-in successfully sideloaded
- [ ] Account 1 authenticated: ✅/❌
- [ ] Account 2 authenticated: ✅/❌  
- [ ] Sync completed: ✅/❌
- [ ] Busy blocks created in Account 1: __ blocks
- [ ] Busy blocks created in Account 2: __ blocks
- [ ] Conflicts detected: __ conflicts
- [ ] Sync time: __ seconds

**Issues Found:**
- ___________________________
- ___________________________

### **🎯 Next Steps**
Once 2-account testing is successful:
1. ✅ Add your 3rd account
2. ✅ Add your 4th account  
3. ✅ Add your 5th account
4. ✅ Test the full August 12th scenario
5. ✅ Deploy to Microsoft 365 Admin Center

---
*Start simple, then scale up! This 2-account test will validate your core functionality.*
