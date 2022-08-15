import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import axios from 'axios';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

// These variables should be configured by the user's site
const sharepointSiteUrl = process.env.SHAREPOINT_SITE_URL;
const sitePage = process.env.SITE_PAGE

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPDFTronCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'PDFTronCommandSet';

export default class PDFTronCommandSet extends BaseListViewCommandSet<IPDFTronCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PDFTronCommandSet');
    Log.info('absolute_url', this.context.pageContext.web.toString());
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('OPEN_IN_PDFTRON');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const fileRef = event.selectedRows[0].getValueByName('FileRef');
    const fileName = event.selectedRows[0].getValueByName('FileLeafRef');
    const spItemUrl = event.selectedRows[0].getValueByName('.spItemUrl');
    const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
    const folderName = fileRef.match(`(?<=${serverRelativeUrl}\/)(.*?)(?=\/${fileName})`)[1]; 
    const {displayName, email} = this.context.pageContext.user;

    switch (event.itemId) {
      case 'OPEN_IN_PDFTRON':
        axios.get(spItemUrl).then(({data}) => {
          const downloadUrl = data['@content.downloadUrl'];
          const url = new URL(downloadUrl);
          const urlParams =  new URLSearchParams(url.search);
          const uniqueId = urlParams.get('UniqueId');
          const tempAuth = urlParams.get('tempauth');
          window.open(`${sharepointSiteUrl}/SitePages/${sitePage}?filename=${fileName}&foldername=${folderName}&username=${displayName}&email=${email}&uniqueId=${uniqueId}&tempAuth=${tempAuth}`);
          
          // If you're using sharepoint-static, just copy the URL of your page and put into the following

          // let staticPageUrl = `https://{your-tenant-id}.sharepoint.com/sites/{site-name}/Shared%20Documents/{your-static-folder-path}`;
          // window.open(`${staticPageUrl}?filename=${fileName}&foldername=${folderName}&username=${displayName}&email=${email}&uniqueId=${uniqueId}&tempAuth=${tempAuth}`);
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
