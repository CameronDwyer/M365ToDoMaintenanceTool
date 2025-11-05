# Microsoft 365 To-Do Maintenance Tool

A C# console application that connects to your Microsoft 365 tenant to safely delete completed tasks from Microsoft To-Do, improving application performance.

## Prerequisites

- .NET 8.0 SDK installed
- Active Microsoft 365 subscription with To-Do access
- Azure AD app registration (see setup instructions below)

## Azure AD App Registration Setup

Before running the application, you need to create an Azure AD app registration:

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Fill in the details:
   - **Name**: `Microsoft To-Do Maintenance Tool` (or your preferred name)
   - **Supported account types**: Select "Accounts in this organizational directory only" (or "Accounts in any organizational directory" if multi-tenant)
   - **Redirect URI**: Select "Public client/native (mobile & desktop)" and enter `http://localhost`
5. Click **Register**

### Step 2: Copy Client ID

1. After registration, you'll see the **Overview** page
2. Copy the **Application (client) ID** - you'll need this for `appsettings.json`
3. (Optional) Copy the **Directory (tenant) ID** if you want to restrict to your specific tenant

### Step 3: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add the following permissions:
   - `Tasks.ReadWrite` - Read and write user's tasks
   - `User.Read` - Sign in and read user profile
5. Click **Add permissions**
6. (Optional) Click **Grant admin consent** if you have admin rights, or ask your admin to grant consent

### Step 4: Configure Authentication

1. Go to **Authentication** in your app registration
2. Under **Platform configurations**, ensure you have:
   - Platform: **Mobile and desktop applications**
   - Redirect URI: `http://localhost`
3. Under **Advanced settings** → **Allow public client flows**: Set to **Yes**
4. Click **Save**

## Configuration

1. Open `appsettings.json`
2. Replace `YOUR_CLIENT_ID_HERE` with your Azure AD app's Client ID:

```json
{
  "AzureAd": {
    "ClientId": "YOUR_CLIENT_ID_HERE",
    "TenantId": "common",
    "Scopes": [
      "Tasks.ReadWrite",
      "User.Read"
    ]
  },
  "ToDoSettings": {
    "TargetListName": "Tasks",
    "DryRun": false
  }
}
```

### Configuration Options

- **ClientId**: Your Azure AD app registration Client ID (required)
- **TenantId**: Use `common` for any Microsoft account, or your specific tenant ID
- **Scopes**: Microsoft Graph API permissions required
- **TargetListName**: The name of the task list to clean (default: "Tasks" which is the standard "My Tasks" list)
- **DryRun**: Set to `true` for testing without actual deletion

## Building the Application

```powershell
dotnet restore
dotnet build
```

## Running the Application

```powershell
dotnet run
```

Or build and run the executable:

```powershell
dotnet build -c Release
cd bin\Release\net8.0
.\ToDoMaintenance.exe
```

## How It Works

The application authenticates via browser (OAuth 2.0), retrieves your task lists, filters for completed tasks only, shows you a preview with counts, and asks for confirmation before deletion. It processes deletions in batches of 20 with automatic retry handling for API throttling. Only completed tasks are deleted - active tasks are never touched. Set `DryRun: true` in config to test safely without actual deletion.

## Troubleshooting

**Authentication fails**: Verify Client ID in `appsettings.json`, ensure "Allow public client flows" is enabled in Azure AD, and check that API permissions (`Tasks.ReadWrite`, `User.Read`) are granted.

**List not found**: The default list name is "Tasks" - check your Microsoft To-Do app for the exact list name and update `TargetListName` in config if different.

## Notes
- Works with M365 Work and School accounts only
- No data is stored locally - all operations are against Microsoft's API
- Uses delegated permissions (operates in your user context)

## Disclaimer and License

**IMPORTANT: USE AT YOUR OWN RISK**

This tool permanently deletes completed tasks from Microsoft To-Do. Once deleted, tasks cannot be recovered.

Key points:
- This software is provided as-is without warranty
- You are solely responsible for any data loss
- Always review the task list before confirming deletion
- Test with DryRun mode first
- Ensure you have proper backups if needed

By using this software, you accept full responsibility for any consequences.

---


