# Enable Apps for SharePoint Server

If the **Add It** button is disabled and SharePoint shows:

```text
Sorry, apps are turned off. If you know who runs the server, tell them to enable apps.
```

the SPFx package is not the problem. The SharePoint farm Apps platform is not enabled.

These steps must be done by a SharePoint farm administrator on the SharePoint server.

## Required Farm Services

In Central Administration, make sure these service applications exist and are started:

- App Management Service Application
- Microsoft SharePoint Foundation Subscription Settings Service Application

Also confirm these services are running under **Manage services on server**:

- App Management Service
- Microsoft SharePoint Foundation Subscription Settings Service

## If Configure App URLs Shows Subscription Settings Error

If **Configure App URLs** says:

```text
The Subscription Settings service and corresponding application and proxy needs to be running in order to make changes to these settings.
```

run the setup from the SharePoint Management Shell as a farm administrator. First check the service instances:

```powershell
Get-SPServiceInstance | Where-Object {$_.TypeName -like "*Subscription Settings*"}
Get-SPServiceInstance | Where-Object {$_.TypeName -like "*App Management*"}
```

Start them if they are not online:

```powershell
Get-SPServiceInstance | Where-Object {$_.TypeName -like "*Subscription Settings*"} | Start-SPServiceInstance
Get-SPServiceInstance | Where-Object {$_.TypeName -like "*App Management*"} | Start-SPServiceInstance
```

Then create the missing service applications/proxies if they do not already exist.

## Configure App URLs

Open Central Administration:

1. Go to **Apps**.
2. Open **Configure App URLs**.
3. Set an app domain and prefix.

Example:

```text
App domain: apps.spse26h.local
App prefix: apps
```

The app domain must resolve in DNS to the SharePoint web front-end server.

## PowerShell Example

Run this in the SharePoint Management Shell as farm admin. Adjust the managed account and app domain to match your farm.

```powershell
$account = Get-SPManagedAccount "DOMAIN\spfarm"

$appPool = New-SPServiceApplicationPool `
  -Name "SharePoint Apps Service Application Pool" `
  -Account $account

$appMgmt = New-SPAppManagementServiceApplication `
  -Name "App Management Service Application" `
  -ApplicationPool $appPool

New-SPAppManagementServiceApplicationProxy `
  -Name "App Management Service Application Proxy" `
  -ServiceApplication $appMgmt

$settings = New-SPSubscriptionSettingsServiceApplication `
  -ApplicationPool $appPool `
  -Name "Subscription Settings Service Application" `
  -DatabaseName "SubscriptionSettingsServiceDB"

New-SPSubscriptionSettingsServiceApplicationProxy `
  -ServiceApplication $settings

Set-SPAppDomain "apps.spse26h.local"
Set-SPAppSiteSubscriptionName -Name "apps" -Confirm:$false
```

Then run IIS reset during a maintenance window:

```powershell
iisreset
```

## After Apps Are Enabled

1. Go back to the App Catalog.
2. Upload or replace `sharepoint\solution\gm-approval-dashboard.sppkg`.
3. Deploy/trust the package.
4. Open the target site.
5. Site Contents -> add `gm-approval-dashboard-client-side-solution`.
6. Add the web part to the page and publish it.
