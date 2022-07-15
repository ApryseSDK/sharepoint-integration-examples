# SharePoint-extension
- change the `sharepointSiteUrl` to the your sharepoint site url, ex: https://5s4vrg.sharepoint.com/sites/5s4vrg
- change the `page` to your SharePoint page that has webviewer
- run `gulp serve` and enable debug script.

To deploy the extension to SharePoint, follow the steps below:
- use `gulp bundle --ship` to bundle the extension
- use `gulp package-solution --ship` to create the package file
Login to the SharePoint admin center, and navigate to __Manage Apss__ page, then upload the custom app (package file) and enable it.
In your SharePoint site, add the extension by clicking the setting button on top right and click "Add an app". Select the extension app you just enabled.
Go back to the document library, and it should show *Open in PDFTron* option when you right click a document.
After clicking, it will redirect to the page specified, whether it's page that contains webviewer web part or a static page with webviewer.


# SharePoint-web-part
## Developing web part
- follow the guide on PDFTron website to map a network drive to the sharepoint master page gallery.
- Create a folder called `pdftron` in the mapped network drive, move the 'lib' folder, which contains `ui` and `core` folders, to the `pdftron` folder.
- Change the `path` in Webviewer options to `'/_catalogs/masterpage/pdftron/lib'`.
- Change the `uiPath` in Webviewer options to `./ui/index.aspx`. (This is important, otherwise the webviewer will be unavailable to get the UI interface).
- To view the web part in browser, use `gulp serve`.

## Deploy the app
Before deploying the app, we need to bundle it and create a solution package.
- Use `gulp bundle --ship` to bundle the web part with static files.
- Use `gulp package-solution --ship` to create the solution package file.
- To deploy the app, log in as the sharepoint admin in __More features in the SharePoint admin center__. On the __Manage Apps__ page, upload the custom app and enable it.
- Then in the SharePoint site, add the custom app by clicking the setting button on top right and click "Add an app". Select the app you just uploaded.
- Create a page and add the web part to the page.

# Alternative: Using webviewer with SharePoint in document library
Since MacOS users aren't available to map the their network drive to the SharePoint Master page gallery, it is possible to integrate Weviewer with Sharepoint using the document library. 

To enable custom page in Sharepoint, start with connecting to the SharePoint Online Management Shell:
- On Windows, you can use powershell.
Use `Connect-SPOService -Url https://{your-tenant-id}-admin.sharepoint.com`.
Then use `Set-SPOsite https://{your-tenant-id}/sites/{site-name} -DenyAddAndCustomizePages 0` to enable customize pages.

- On Mac, you can use PnP powershell.
Install Pnp powershell: https://www.c-sharpcorner.com/article/how-to-run-pnp-powershell-in-macos/
After installation, use `Connect-PnPOnlin -Url https://{your-tenant-id}-admin.sharepoint.com -Interactive` to login.
Then use `Set-PnPSite -Identity https://{your-tenant-id}.sharepoint.com/sites/${site-name} -NoScriptSite $false` to enable custom script in SharePoint.

After configuring, upload the sharepoint-static to the document library. Navigate into the sharepoint-static folder in SharePoint, and click the index.aspx. This should open a web page that contains webviewer.

To integrate with sharepoint-extension, just replace the url with the url of your page in `window.open(...)`.