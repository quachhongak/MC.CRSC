# CRSC Prompts deployment

## Build the production package

Use Node.js 18.17.1 through 18.x with the SPFx 1.18.2 toolchain:

```powershell
npm install
npx gulp bundle --ship
npx gulp package-solution --ship
```

The production package is `sharepoint/solution/crsc-prompts.sppkg` (the package-solution configuration names this output). Static assets are included in the package; no external CDN is required.

## Site collection app catalog

This solution is configured for site collection installation (`skipFeatureDeployment: false`). A SharePoint administrator must enable a site collection app catalog for the target site if one is not already enabled. For example, with the SharePoint Online Management Shell and appropriate administrator access:

```powershell
Connect-SPOService -Url https://<tenant>-admin.sharepoint.com
Set-SPOSite -Identity https://<tenant>.sharepoint.com/sites/<site> -SiteCollectionAppCatalogEnabled $true
```

## Upload and install

1. Open the target site's **Apps for SharePoint** library.
2. Upload `crsc-prompts.sppkg`. If prompted to deploy the package, confirm it.
3. On the target site, add the **CRSC-Prompts** app from **Add an app**.
4. Edit a page, add the **CRSC-Prompts** web part, and publish the page.

## Configure

Open the web part properties, select the prompt list, and choose 50, 100, or 200 items per page. The configured list needs the internal fields listed in [README.md](./README.md). The web part displays only prompts whose `Status` choice is `Approved`. Users need read access to the site and list.

## Update and rollback

To update, build and upload the new `.sppkg` to the same app catalog, then deploy the updated package. To roll back, upload a prior known-good package version or remove the app from the site and restore the prior package according to the site's normal change-management process. Keep a copy of each deployed package.

## Limitations

- The prompt list must be on the same SharePoint site as the web part.
- The property picker displays up to 200 visible lists.
- Prompt list fields must use the internal names documented in the README.
- No Microsoft Graph permissions, cross-site access, caching, or telemetry are configured.
