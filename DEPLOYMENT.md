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

## SharePoint prerequisites

Before using PDF creation, verify these fields and permissions:

- `GM Requests` has a plain-text multiple-lines field with internal name `PDFFileJson`, configured with enough capacity for the schema-version-3 document metadata and append-only conversation history, and hyperlink fields named `PDFFileUrl` and `SignedPDFUrl`.
- If `Status` is a Choice field, include `Pending Secretary Information` and `Pending HOD Information` in addition to the existing workflow statuses.
- List attachments are enabled on `GM Requests`.
- `GM Approval Documents` has `RequestNo` (text) and `RequestItemId` (number). If `DocumentType` is a Choice field, include `Request PDF`, `Office Manager PDF`, `Workflow PDF`, `HOD PDF`, and `Secretary PDF`.
- Office Managers have permission to create folders and upload/replace files in `GM Approval Documents`, and to update/attach files to `GM Requests`.
- Members of `GM Secretaries` have permission to read and update `GM Requests` while replying to `Pending Secretary Information` requests.
- Members of `GM HODs` (or the legacy `GM HOD` group) have permission to read and update `GM Requests` while replying to `Pending HOD Information` requests.

The web part creates one library folder per request reference. It stores the original supporting PDF separately and saves the Office Manager workflow document as the deterministic `<RequestNo>.pdf`; optional HOD and Secretary documents use `HOD-<RequestNo>.pdf` and `Secretary-<RequestNo>.pdf`. `PDFFileJson` uses schema version 3 with `documents` and `conversation` collections; existing schema-version-1 and schema-version-2 records are read and upgraded without discarding their document metadata or prior role comments.

The file-sync workflow prefers an exact `<RequestNo>.pdf` file, then a document tagged `Workflow PDF`, and then the existing `Request PDF`. `Pending GM Signature` uses the `Pending` share folder. `GM Signed Pending Office Manager Confirmation`, `Approved by GM`, `Pending Secretary Information`, and `Pending HOD Information` use the `Signed` share folder. `Closed` and `Rejected` use `Archive`.

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
