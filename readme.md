# SharePoint-extension#
- change the `sharepointSiteUrl` to the your sharepoint site url, ex: https://5s4vrg.sharepoint.com/sites/5s4vrg
- change the `page` to your SharePoint page that has webviewer
- run `gulp serve` and enable debug script.


# SharePoint-web-part#
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