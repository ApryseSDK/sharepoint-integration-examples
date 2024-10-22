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

export interface IWebviewerWebPartProps {
  description: string;
}

export default class WebviewerWebPart extends BaseClientSideWebPart<IWebviewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _mode: string;
  private _graphConsumer: GraphConsumer;

  public validateQueryParam(urlParams: URLSearchParams): boolean {
    const necessaryParams: string[] = ['filename'];
    let result: boolean = true;
    necessaryParams.forEach(paramKey => {
      if (!urlParams.get(paramKey)) {
        result = false;
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
	this._graphConsumer = new GraphConsumer(this.context);
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
      this._createSavedModal(instance);
      this._createSaveFileButton(instance);
      if (validateQueryParamResult) {
        this._mode = "sharepoint-file";
        const filename: string = urlParams.get("filename");
		const folderName: string = urlParams.get("foldername");
		const docURL: string = `${window.location.origin}/sites/${process.env.SITE_NAME}/_api/web/GetFolderByServerRelativeUrl('${folderName}')/Files(url='${filename}')/$value`;
		
		instance.UI.loadDocument(docURL, {filename});
      } else {
        this._mode = "local-file";
      }
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
    const annotationManager: Core.AnnotationManager = instance.Core.annotationManager;
    const xfdfString: string = await annotationManager.exportAnnotations();
    const fileData: ArrayBuffer = await instance.Core.documentViewer.getDocument().getFileData({ xfdfString });
    const digest: string = await this._getFormDigestValue();
    const fileArray: Uint8Array= new Uint8Array(fileData);
    const file: File = new File([fileArray], fileName, {
      type: 'application/pdf'
    });
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
