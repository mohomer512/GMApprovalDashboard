# GM Approval Dashboard Deployment

Use this process for production deployment. Users do not need `gulp serve`.

## Build the package

```powershell
gulp bundle --ship
gulp package-solution --ship
```

The package is created here:

```text
C:\GMApprovalDashboard\sharepoint\solution\gm-approval-dashboard.sppkg
```

## Deploy to SharePoint Server

1. Open the SharePoint App Catalog site.
2. Upload `sharepoint\solution\gm-approval-dashboard.sppkg` to the `Apps for SharePoint` library.
3. When SharePoint asks, deploy/trust the solution.
4. Open the target site, for example `http://spse26h/GM`.
5. Go to Site Contents, add the app `gm-approval-dashboard-client-side-solution`.
6. Edit the page where the dashboard should appear.
7. Add the `General Manager Approval Dashboard` web part.
8. Publish the page.

## Update Later

For every production update:

1. Increase the version in `config/package-solution.json`.
2. Run the two build commands again.
3. Upload the new `.sppkg` to the App Catalog and replace the old package.
4. Trust/deploy the updated package.

Because `includeClientSideAssets` is `true`, the JavaScript, CSS, and logo are included inside the `.sppkg`.
