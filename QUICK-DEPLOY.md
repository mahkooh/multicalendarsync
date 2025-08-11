# Quick Deployment Script for Organization Admin

## 🚀 5-Minute Deployment to Your Organization

### Option 1: GitHub Pages (Fastest - Free)

```bash
# 1. Initialize git repository
git init
git add .
git commit -m "MultiCalendar Sync - Initial version"

# 2. Create GitHub repository (replace 'yourusername' with your GitHub username)
# Go to github.com and create a new repository called 'multicalendar-sync'

# 3. Push to GitHub
git remote add origin https://github.com/yourusername/multicalendar-sync.git
git branch -M main
git push -u origin main

# 4. Enable GitHub Pages
# - Go to your repository on GitHub
# - Click Settings > Pages
# - Source: Deploy from a branch
# - Branch: main / (root)
# - Click Save
# 
# Your add-in will be available at:
# https://yourusername.github.io/multicalendar-sync
```

### Option 2: Azure Static Web Apps

```bash
# 1. Install Azure CLI (if not already installed)
# Download from: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli

# 2. Login to Azure
az login

# 3. Create resource group (if needed)
az group create --name "multicalendar-rg" --location "East US"

# 4. Build the application
npm run build

# 5. Create static web app
az staticwebapp create \
  --name "multicalendar-sync" \
  --resource-group "multicalendar-rg" \
  --source "." \
  --location "East US 2" \
  --branch "main"

# Your add-in will be available at:
# https://multicalendar-sync.azurestaticapps.net
```

## 📝 Update Manifest for Production

After deployment, update these URLs in `manifest.json`:

```json
{
  "runtimes": [
    {
      "code": {
        "page": "https://YOUR-DOMAIN/taskpane.html"
      }
    }
  ]
}
```

Replace `YOUR-DOMAIN` with:
- GitHub Pages: `yourusername.github.io/multicalendar-sync`
- Azure: `multicalendar-sync.azurestaticapps.net`
- Your hosting: `your-domain.com`

## 🏢 Upload to Microsoft 365 Admin Center

1. **Go to**: [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. **Navigate**: Settings > Integrated apps
3. **Click**: "Upload custom apps"
4. **Select**: "Upload manifest file"
5. **Upload**: Your updated `manifest.json`
6. **Configure**:
   - Deploy to: "Everyone" or specific groups
   - Auto-deploy: Enable
   - Default state: Enabled
7. **Click**: "Deploy"

## ✅ Verification Checklist

Before uploading to Admin Center:

- [ ] Add-in builds successfully (`npm run build`)
- [ ] All URLs in manifest are HTTPS
- [ ] Production URLs are accessible
- [ ] Manifest validates (`npm run validate`)
- [ ] Test sideloading works with production manifest

## 📞 Need Help?

If you run into issues:

1. **Manifest validation**: Run `npm run validate`
2. **URL accessibility**: Test all URLs in manifest
3. **Admin permissions**: Ensure Global or Exchange Admin role
4. **Propagation time**: Allow up to 24 hours for deployment

Your users will see "MultiCalendar Sync" in their Outlook add-ins within 24 hours of deployment!
