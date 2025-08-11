# Testing MultiCalendar Sync Add-in

## 🧪 Method 1: Sideload in Outlook Web (Easiest)

### Step 1: Access Outlook Web
1. Go to https://outlook.office.com
2. Sign in with your Microsoft 365 account
3. Make sure you have calendar access

### Step 2: Sideload the Add-in
1. Click the **Settings** gear icon (top right)
2. Search for "Get Add-ins" or go to **View all Outlook settings**
3. Navigate to **General** → **Manage add-ins**
4. Click **+ Add a custom add-in** → **Add from file**
5. Upload the `manifest-test.xml` file from this directory
6. Click **Install** when prompted

### Step 3: Test the Add-in
1. Go to your **Calendar** view
2. Look for "MultiCalendar Sync" in the ribbon or add-ins panel
3. Click to open the add-in panel
4. Verify the UI loads correctly

---

## 🧪 Method 2: Use Office Add-in Development Tools

### Install Office Add-in Debugger
```bash
npm install -g office-addin-debugging
```

### Start Local Testing Server
```bash
npm run dev-server
```

### Test Locally First
```bash
npm run start
```

---

## 🧪 Method 3: Validate Manifest First

### Online Validation
1. Go to https://appsource.microsoft.com/marketplace/office
2. Use Microsoft's manifest validation tool
3. Upload our `manifest-test.xml` file

### Command Line Validation
```bash
office-addin-manifest validate manifest-test.xml
```

---

## 🔍 What to Test

### ✅ Basic Functionality Checklist
- [ ] Add-in loads without errors
- [ ] UI displays correctly in the task pane
- [ ] Calendar sync controls are visible
- [ ] Status indicators work
- [ ] Activity log displays messages
- [ ] Error handling works gracefully

### ✅ Calendar Integration Checklist
- [ ] Add-in can read calendar permissions
- [ ] Calendar list populates (even if simulated)
- [ ] Sync button responds to clicks
- [ ] Status updates appear in real-time

### ✅ Technical Validation
- [ ] All HTTPS URLs load correctly
- [ ] No console errors in browser dev tools
- [ ] Responsive design works in the task pane
- [ ] Icons and assets load properly

---

## 🐛 Common Issues and Solutions

### Issue: "Add-in won't load"
**Check:**
- Manifest XML syntax is valid
- All URLs return HTTP 200
- HTTPS is enabled for all resources
- CORS is properly configured

### Issue: "Permission errors"
**Check:**
- Manifest requests correct permissions
- User has calendar access in their account
- Add-in is properly scoped for mailbox access

### Issue: "UI doesn't display"
**Check:**
- HTML file loads correctly
- CSS and JS files are accessible
- Browser developer console for errors

---

## 🔧 Debug Commands

### Test All URLs
```bash
# Test main taskpane
curl -I https://mahkooh.github.io/multicalendarsync/dist/taskpane.html

# Test commands page
curl -I https://mahkooh.github.io/multicalendarsync/dist/commands.html

# Test assets
curl -I https://mahkooh.github.io/multicalendarsync/assets/icon-64.png
```

### Validate Manifest Structure
```bash
# Check XML syntax
xml val manifest-test.xml

# Check required elements
grep -E "(Id|DisplayName|ProviderName)" manifest-test.xml
```

---

## 📋 Testing Results Template

After testing, document results:

```
### Test Date: [DATE]
### Test Environment: Outlook Web / Desktop
### User Account: [YOUR EMAIL]

**Basic Loading:**
- [ ] Manifest loads successfully
- [ ] Task pane opens
- [ ] UI renders correctly

**Functionality:**
- [ ] Calendar sync panel displays
- [ ] Controls are interactive
- [ ] Status updates work

**Issues Found:**
- List any problems encountered
- Note error messages or console logs

**Next Steps:**
- Priority fixes needed
- Additional testing required
```

---

## 🚀 Ready for Organization Deployment?

Once testing passes:
1. Document successful test results
2. Fix any issues found
3. Re-test after fixes
4. Then proceed with admin center deployment

**Start with sideloading in Outlook Web - it's the fastest way to validate the add-in!**
