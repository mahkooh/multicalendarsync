# Organization Deployment Guide - MultiCalendar Sync

## 🏢 Deploying to Your Organization (Admin Guide)

Since you're the organization administrator, you can deploy MultiCalendar Sync directly to your users without going through Microsoft AppSource certification.

## 📋 Prerequisites

### What You Need
- ✅ **Microsoft 365 Admin Center access**
- ✅ **Global Administrator or Exchange Administrator role**
- ✅ **Web hosting for the add-in files** (Azure, AWS, or any HTTPS server)
- ✅ **Domain with SSL certificate** (for HTTPS)

### Current Status
- ✅ Add-in code is complete and tested
- ✅ Manifest.json is configured
- ⏳ Need to deploy to production hosting
- ⏳ Need to upload to organization catalog

## 🚀 Step-by-Step Deployment

### Step 1: Deploy Add-in to Production Server

#### Option A: Azure Static Web Apps (Recommended)
```bash
# Build production version
npm run build

# Deploy to Azure (requires Azure CLI)
az staticwebapp create \
  --name "multicalendar-sync" \
  --resource-group "your-resource-group" \
  --source "dist" \
  --location "East US 2"
```

#### Option B: Any Web Hosting Service
1. **Build the production files**:
   ```bash
   npm run build
   ```

2. **Upload the `dist` folder** to your web server
3. **Ensure HTTPS is enabled**
4. **Note your domain** (e.g., `https://multicalendar.yourcompany.com`)

### Step 2: Update Manifest for Production

Update the URLs in `manifest.json`:

```json
{
  "runtimes": [
    {
      "code": {
        "page": "https://your-domain.com/taskpane.html"
      }
    }
  ]
}
```

### Step 3: Upload to Microsoft 365 Admin Center

#### Access Admin Center
1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings** > **Integrated apps**
3. Click **Upload custom apps**

#### Upload Your Add-in
1. **Choose upload type**: "Upload manifest file"
2. **Upload manifest.json** (your updated production version)
3. **Configure deployment**:
   - **Deploy to**: "Specific users/groups" or "Everyone"
   - **Auto-deploy**: Enable for automatic installation
   - **Default state**: Enabled

#### Deployment Settings
- **Name**: MultiCalendar Sync
- **Publisher**: Your organization name
- **Description**: Synchronize busy time across multiple calendars
- **Privacy policy**: Create at `https://your-domain.com/privacy`
- **Terms of use**: Create at `https://your-domain.com/terms`

### Step 4: Configure Permissions (Optional)

If your add-in needs additional Graph API permissions:

1. Go to **Azure Portal** > **App registrations**
2. Register a new application for your add-in
3. Configure the required permissions:
   - `Calendars.ReadWrite`
   - `Calendars.ReadWrite.Shared`
   - `User.Read`
4. Update your add-in code to use the registered app ID

## 🎯 Quick Deployment (5 Minutes)

### For Immediate Testing
If you want to get this working immediately for testing:

1. **Use GitHub Pages** (free hosting):
   ```bash
   # Push your code to GitHub
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/yourusername/multicalendar.git
   git push -u origin main
   
   # Enable GitHub Pages in repository settings
   # Your add-in will be available at: https://yourusername.github.io/multicalendar
   ```

2. **Update manifest.json** with GitHub Pages URL
3. **Upload to Admin Center** using the GitHub Pages manifest

## 📱 User Installation Process

### Automatic Installation (Recommended)
When you deploy with "auto-deploy" enabled:
- ✅ **Appears automatically** in users' Outlook
- ✅ **No user action required**
- ✅ **Available within 24 hours**

### Manual Installation
If users need to install manually:
1. **Outlook Desktop**: File > Get Add-ins > My Organization
2. **Outlook Web**: Settings gear > View all settings > Mail > Manage apps
3. **Find "MultiCalendar Sync"** and click Install

## 🔒 Security & Compliance

### Data Protection
- ✅ **All data processing happens locally** in user's browser
- ✅ **No data stored on external servers**
- ✅ **Only calendar busy/free status accessed**
- ✅ **Meeting details remain private**

### Admin Controls
- **Usage monitoring**: View adoption metrics in Admin Center
- **Permission management**: Control which users can access
- **Remove/update**: Easy deployment management
- **Audit logs**: Track installation and usage

## 🎛️ Admin Configuration Options

### Deployment Scope
- **Everyone**: All users in organization
- **Specific groups**: Selected security groups
- **Specific users**: Individual user selection
- **Pilot groups**: Test with small group first

### Installation Method
- **Auto-deploy**: Installs automatically
- **Available to install**: Users can choose to install
- **Admin pre-install**: Admin installs for users

## 🚨 Troubleshooting

### Common Issues
1. **Manifest validation errors**: Check all URLs are HTTPS and accessible
2. **Permissions denied**: Ensure you have Global Admin or Exchange Admin role
3. **Add-in not appearing**: Wait up to 24 hours for propagation
4. **SSL certificate errors**: Ensure hosting has valid SSL certificate

### Testing Before Deployment
1. **Validate manifest**:
   ```bash
   npm run validate
   ```
2. **Test with sideloading** first
3. **Deploy to test group** before organization-wide rollout

## 📞 Support & Maintenance

### User Support
- **Documentation**: Provide user guide for calendar sync features
- **Help desk**: Train support team on add-in functionality
- **Feedback collection**: Set up channel for user feedback

### Ongoing Maintenance
- **Monitor usage**: Check Admin Center analytics
- **Update as needed**: Deploy updates through same process
- **Performance monitoring**: Watch for any issues

## 🎉 Next Steps

1. **Choose hosting option** (Azure, GitHub Pages, or your existing hosting)
2. **Build and deploy** production files
3. **Update manifest.json** with production URLs
4. **Upload to Admin Center**
5. **Test with pilot group**
6. **Deploy organization-wide**

Would you like me to help you with any specific step? I can:
- **Set up Azure Static Web Apps deployment**
- **Configure the production manifest**
- **Create the required privacy/terms pages**
- **Walk through the Admin Center upload process**
