# Production Deployment Instructions

## 🚨 Current Manifest Status

Your current `manifest.json` has **localhost URLs** that need to be updated for production deployment:

### Current (Development) URLs:
- ❌ `https://localhost:3000/taskpane.html`
- ❌ `https://localhost:3000/commands.html`  
- ❌ `https://localhost:3000/commands.js`

### Need to Update To (Production):
- ✅ `https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/taskpane.html`
- ✅ `https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/commands.html`
- ✅ `https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/commands.js`

## 🔧 Quick Fix: Update Your Manifest

### Step 1: Replace URLs in manifest.json

Find these lines in your `manifest.json` and update them:

```json
// FIND THIS:
"page": "https://localhost:3000/taskpane.html"

// REPLACE WITH:
"page": "https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/taskpane.html"
```

```json
// FIND THIS:
"page": "https://localhost:3000/commands.html",
"script": "https://localhost:3000/commands.js"

// REPLACE WITH:
"page": "https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/commands.html",
"script": "https://YOUR-GITHUB-USERNAME.github.io/multicalendar-sync/commands.js"
```

### Step 2: Update Valid Domains

```json
// FIND THIS:
"validDomains": [
    "contoso.com"
]

// REPLACE WITH:
"validDomains": [
    "YOUR-GITHUB-USERNAME.github.io"
]
```

## 📋 Complete Deployment Checklist

### Before Uploading to Admin Center:

1. **Host your files**:
   - [ ] Push code to GitHub
   - [ ] Enable GitHub Pages
   - [ ] Verify files are accessible

2. **Update manifest.json**:
   - [ ] Replace all localhost URLs with GitHub Pages URLs
   - [ ] Update validDomains
   - [ ] Test that all URLs are accessible

3. **Validate manifest**:
   ```bash
   npm run validate
   ```

4. **Upload to Microsoft 365 Admin Center**:
   - [ ] Go to admin.microsoft.com
   - [ ] Settings > Integrated apps
   - [ ] Upload custom apps
   - [ ] Select your updated manifest.json

## 🚀 Quick Commands

### Update URLs in Manifest (PowerShell):
```powershell
# Replace YOUR-GITHUB-USERNAME with your actual GitHub username
$githubUsername = "YOUR-GITHUB-USERNAME"
$manifestPath = "manifest.json"

(Get-Content $manifestPath) -replace "https://localhost:3000", "https://$githubUsername.github.io/multicalendar-sync" | Set-Content $manifestPath
(Get-Content $manifestPath) -replace "contoso\.com", "$githubUsername.github.io" | Set-Content $manifestPath
```

### Validate Updated Manifest:
```bash
npm run validate
```

## ⚠️ Important Notes

1. **All URLs must be HTTPS** - GitHub Pages provides this automatically
2. **Files must be publicly accessible** - Don't use private repositories
3. **Wait for propagation** - GitHub Pages can take 5-10 minutes to update
4. **Test before deploying** - Sideload the updated manifest first

## 🎯 Next Steps

1. **Choose your GitHub username** for hosting
2. **Update the manifest URLs** as shown above
3. **Push to GitHub and enable Pages**
4. **Upload to your organization** via Admin Center

The add-in code is ready - it just needs the correct production URLs!
