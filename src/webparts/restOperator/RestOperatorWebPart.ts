import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './RestOperator.module.scss';
import * as strings from 'restOperatorStrings';
import { IRestOperatorWebPartProps } from './IRestOperatorWebPartProps';

import { SPHttpClient, SPHttpClientConfigurations, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import { IODataUser, IODataWeb } from '@microsoft/sp-odata-types';

export default class RestOperatorWebPart extends BaseClientSideWebPart<IRestOperatorWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.row}">
        <div class="${styles.column}">
          <span class="${styles.title}">
            Welcome to SharePoint!
          </span>
          <p class="${styles.subtitle}">
            Customize SharePoint experiences using Web Parts.
          </p>
          <p class="${styles.description}">
            ${escape(this.properties.description)}
          </p>
          <a class="ms-Button ${styles.button}" href="https://github.com/SharePoint/sp-dev-docs/wiki">
            <span class="ms-Button-label">
              Learn more
            </span>
          </a>
        </div>
      </div>`;

      const spHttpClient: SPHttpClient = this.context.spHttpClient;
      const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

      // spHttpClient.get(currentWebUrl + `/_api/web`, SPHttpClientConfigurations.v1).then((response: SPHttpClientResponse) => {
        
      //     response.json().then((web: IODataWeb) => {

      //         console.log(web.Url);
      //     });
      // });

      // const spHttpClient: SPHttpClient = this.context.spHttpClient;

      // spHttpClient.get(`/_api/web/currentuser`, SPHttpClientConfigurations.v1).then((response: SPHttpClientResponse) => {
        
      //     response.json().then((user: IODataUser) => {

      //         console.log(user.LoginName);
      //     });
      // });

      const spFlags : ISPHttpClientConfiguration = {
          defaultSameOriginCredentials: true,
          defaultODataVersion: ODataVersion.v3,
          requestDigest: true
      };

      const clientConfigODataV3: SPHttpClientConfiguration  = new SPHttpClientConfiguration(spFlags)

      spHttpClient.get(`/_api/search/query?querytext='sharepoint'`, clientConfigODataV3).then((response: SPHttpClientResponse) => {
        
          response.json().then((user: IODataUser) => {

              console.log(user.LoginName);
          });
      });
      
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
