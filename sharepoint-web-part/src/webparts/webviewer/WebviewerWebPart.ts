import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import WebViewer, { Core, UI, WebViewerInstance } from '@pdftron/webviewer';

import * as strings from 'WebviewerWebPartStrings';

import GraphConsumer from './components/GraphConsumer';
import DataverseQueries from './components/DataverseQueries';

export interface IWebviewerWebPartProps {
  description: string;
}

export default class WebviewerWebPart extends BaseClientSideWebPart<IWebviewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _mode: string;
  private _graphConsumer: GraphConsumer;
  private _dataverseQueries: DataverseQueries
  private _recordID: string = 'ceed5dc6-357a-ed11-81ad-00224826545c';

  public validateQueryParam(urlParams: URLSearchParams): boolean {
    const necessaryParams: string[] = ['filename'];
    let result: boolean = true;
	console.log("Test ID:" + urlParams.get('uniqueID'));
	this._recordID = urlParams.get('uniqueID');
    necessaryParams.forEach(paramKey => {
      if (!urlParams.get(paramKey)) {
        // result = false;
      }
    });
    return result;
  }

  public render(): void {
    this.domElement.style.height = '1000px';
    WebViewer({
      // We suggest to use the method of uploading static files to the Documents folder in your sharepoint site
      // The provided path below is a template, it may varies in your site
      path: `https://${process.env.TENANT_ID}.sharepoint.com/sites/${process.env.SITE_NAME}/Shared Documents/${process.env.WEBVIEWER_LIB_FOLDER_PATH}`,
      // You'll need to indicate the entry point of webviewer ui. In sharepoint, it will be with .aspx extension.
      uiPath: './ui/index.aspx',
    }, this.domElement)
    .then(async instance => {
	const { documentViewer, annotationManager } = instance.Core;
	this._graphConsumer = new GraphConsumer(this.context);
	// this._dataverseQueries = new DataverseQueries(this.context);
	// await this._dataverseQueries.Login();
	await this._graphConsumer.GetCurrentUser();
	await this._graphConsumer.ListUsers();

	const userData: UI.MentionsManager.UserData[] = this._graphConsumer.users.map(s =>
		({ value: s.displayName, email: s.mail }));
	console.info(userData);
	instance.UI.mentions.setUserData(userData);
	instance.Core.annotationManager.setCurrentUser(this._graphConsumer.currentUser.displayName);

      const { Feature } = instance.UI;
      instance.UI.enableFeatures([Feature.FilePicker]);
      const urlParams: URLSearchParams = new URLSearchParams(window.location.search);
      const validateQueryParamResult: boolean = this.validateQueryParam(urlParams);

	  await this._graphConsumer.TestAPI(this._recordID);
	  await this._graphConsumer.GetFileAPI(this._recordID);

      this._createSavedModal(instance);
      this._createSaveFileButton(instance);
      if (validateQueryParamResult) {
        this._mode = "sharepoint-file";
        const uniqueId: string = urlParams.get("uniqueId");

        const tempAuth: string = urlParams.get("tempAuth");
        const filename: string = urlParams.get("filename");
        const newPathnameArray: string[] = window.location.pathname.split('/').slice(0, 3);
        const newPathname: string = newPathnameArray.join('/');
        const domain: string = window.location.origin;
        const domainUrl: string = `${domain}${newPathname}`;
        const docUrl: string = `${domainUrl}/_layouts/15/download.aspx?UniqueId=${uniqueId}&Translate=false&tempauth=${tempAuth}&ApiVersion=2.0`;
        // instance.UI.loadDocument(docUrl, {filename});

		// Testing...
		if (this._graphConsumer.testFile !== null)
		{
			instance.UI.loadDocument(this._graphConsumer.testFile);

			documentViewer.addEventListener('annotationsLoaded', async () => {
				console.log("Test" + this._graphConsumer.xfdf);
				await annotationManager.importAnnotations(this._graphConsumer.xfdf);
			})
			
		}
      } else {
        this._mode = "local-file";
		if (this._graphConsumer.testFile !== null)
		{
			instance.UI.loadDocument(this._graphConsumer.testFile);

			documentViewer.addEventListener('annotationsLoaded', async () => {
				console.log("Test" + this._graphConsumer.xfdf);
				await annotationManager.importAnnotations(this._graphConsumer.xfdf);
			})
			
		}
      }

	  instance.UI.mentions.on('mentionChanged', async (mentions, action) => {
		if (action === 'add') {
		  // a new mention was just added to a comment
		  console.log("Mention Added");
		//   await this._graphConsumer.PostSendEmail();
		}
  
		if (action === 'modify') {
		  // the mentioned names in a comment didn't change, but the surrounding text was changed
		  console.log("Mention Modified");
		}
  
		if (action === 'delete') {
		  // a mention was just deleted from a comment
		  console.log("Mention Deleted");
		}
  
		console.log(mentions);
		// [
		// {
		// value: 'John Doe',
		// email: 'johndoe@gmail.com',
		// annotId: '...', // the annotation to which the mention belongs to
		// },
		// {
		// value: 'Jane Doe',
		// email: 'janedoe@gmail.com'
		// annotId: '...',
		// },
		// ]
	  })
    })
    .catch(err => console.error(err));
  }

  private _createSaveFileButton(instance: WebViewerInstance): void {
    const me: WebviewerWebPart = this;
    instance.UI.setHeaderItems(function(header: UI.Header) {
      const saveFileButton: unknown = {
        type: 'actionButton',
        dataElement: 'saveFileButton',
        title: 'Save file to sharepoint',
        img: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="none"/><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>',
        onClick: async function() {
          instance.UI.openElements(['loadingModal']);
          if (me._mode === 'sharepoint-file') {
            const searchparams: URLSearchParams = new URLSearchParams(window.location.search);
            const folderName: string = searchparams.get('foldername');
            const fileName: string = searchparams.get('filename');
            await me.saveFile(instance, folderName, fileName);
          } else if (me._mode === 'local-file') {
            const fileName: string = await instance.Core.documentViewer.getDocument().getFilename();
            const folderName: string = encodeURIComponent(process.env.FOLDER_URL);

            await me.saveFile(instance, folderName, fileName);
          }
          instance.UI.closeElements(['loadingModal']);
          instance.UI.openElements(['savedModal']);
        }
      };
      header.get('viewControlsButton').insertBefore(saveFileButton);
    })
  }

  /* 
    The purpose of this function is to get the request digest (client-side token) for us to go through the authorization
    when uploading the file.
  */
  private async _getFormDigestValue(): Promise<string> {
    try {
      const resp: Response = await fetch(`${window.location.origin}/sites/${process.env.SITE_NAME}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json; odata=verbose'
        },
      });
      
      interface IDigestResponseJson {
        d: {
          GetContextWebInformation: {
            FormDigestValue: string
          }
        }
      }
      const respJson: IDigestResponseJson = await resp.json();
      return respJson.d.GetContextWebInformation.FormDigestValue;
    } catch(error) {
      console.error(error);
    }
  }

  public async saveFile(instance: WebViewerInstance, folderUrl: string, fileName: string): Promise<void> {
	console.log("Save File");
    const annotationManager: Core.AnnotationManager = instance.Core.annotationManager;
    const xfdfString: string = await annotationManager.exportAnnotations({ widgets: false});
	console.log("XFDF Length:" + xfdfString.length);
    const fileData: ArrayBuffer = await instance.Core.documentViewer.getDocument().getFileData({ xfdfString });
    const digest: string = await this._getFormDigestValue();
    const fileArray: Uint8Array= new Uint8Array(fileData);
    const file: File = new File([fileArray], fileName, {
      type: 'application/pdf'
    });

	await this._graphConsumer.SaveXFDF(xfdfString, this._recordID);

    await fetch(`${window.location.origin}/sites/${process.env.SITE_NAME}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${fileName}', overwrite=true)`, {
      method: 'POST',
      body: file,
      headers: {
        'accept': 'application/json; odata=verbose',
        'X-RequestDigest': digest,
        'Content-Length': fileData.byteLength.toString()
      }
    });
  }

  private _createSavedModal(instance: WebViewerInstance): void {
    const divInput: HTMLElement = document.createElement('div');
    divInput.innerText = 'File saved successfully';

    interface IModal { 
      dataElement: string;
      disableBackdropClick?: boolean; 
      disableEscapeKeyDown?: boolean; 
      render: UI.renderCustomModal; 
      header: unknown; 
      body: unknown; 
      footer: unknown; 
    }

    const modal: IModal = {
      dataElement: 'savedModal',
      body: {
        className: 'myCustomModal-body',
        style: {
          'text-align': 'center'
        },
        children: [divInput]
      },
      header: null,
      footer: null,
      render: null
    }
    instance.UI.addCustomModal(modal);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
