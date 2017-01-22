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



//import { SPHttpClient, SPHttpClientConfigurations, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

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

    // Here, 'this' refers to my SPFx webpart which inherits from the BaseClientSideWebPart class.
    // Since I am calling this method from inside the class, I have access to 'this'.
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    //Since the SP Search REST API works with ODataVersion 3, we have to create a new ISPHttpClientConfiguration object with defaultODataVersion = ODataVersion.v3
    const spSearchConfig: ISPHttpClientConfiguration = {
      defaultODataVersion: ODataVersion.v3
    };

    //Override the default ODataVersion.v4 flag with the ODataVersion.v3
    const clientConfigODataV3: SPHttpClientConfiguration = SPHttpClientConfigurations.v1.overrideWith(spSearchConfig);

    //Make the REST call
    spHttpClient.get(`${currentWebUrl}/_api/search/query?querytext='sharepoint'`, clientConfigODataV3).then((response: SPHttpClientResponse) => {

      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
      });
    });

    //GET current web info
    spHttpClient.get(`${currentWebUrl}/_api/web`, SPHttpClientConfigurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((web: IODataWeb) => {

        console.log(web.Url);
      });
    });

    //GET current user information from the User Information List
    spHttpClient.get(`${currentWebUrl}/_api/web/currentuser`, SPHttpClientConfigurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((user: IODataUser) => {

        console.log(user.LoginName);
      });
    });

    //GET current user information from the User Profile Service
    spHttpClient.get(`${currentWebUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClientConfigurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((userProfileProps: any) => {

        console.log(userProfileProps);
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
