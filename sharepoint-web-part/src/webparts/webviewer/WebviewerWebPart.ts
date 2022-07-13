import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import WebViewer from '@pdftron/webviewer';

import * as strings from 'WebviewerWebPartStrings';

export interface IWebviewerWebPartProps {
  description: string;
}

export default class WebviewerWebPart extends BaseClientSideWebPart<IWebviewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  // specify the sharepoint site here
  private _sharepointSiteUrl: string = '';
  
  public validateQueryParam(urlParams: URLSearchParams): boolean {
    const necessaryParams: string[] = ['uniqueId', 'tempAuth', 'filename'];
    let result = true;
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
      path: '/_catalogs/masterpage/pdftron/lib',
      uiPath: './ui/index.aspx'
    }, this.domElement)
    .then(instance => {
      const urlParams: URLSearchParams = new URLSearchParams(window.location.search);
      const validateQueryParamResult: boolean = this.validateQueryParam(urlParams);
      if (validateQueryParamResult) {
        const uniqueId: string = urlParams.get("uniqueId");
        const tempAuth: string = urlParams.get("tempAuth");
        const filename: string = urlParams.get("filename");
        const newPathnameArray: string[] = window.location.pathname.split('/').slice(0, 3);
        const newPathname: string = newPathnameArray.join('/');
        const domain: string = window.location.origin;
        const domainUrl: string = `${domain}${newPathname}`;
        const docUrl: string = `${domainUrl}/_layouts/15/download.aspx?UniqueId=${uniqueId}&Translate=false&tempauth=${tempAuth}&ApiVersion=2.0`;
        instance.UI.loadDocument(docUrl, {filename});
      } else {
        alert('Please open the webviewer with proper document queries.')
      }
    })
    .catch(err => console.log(err));
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
