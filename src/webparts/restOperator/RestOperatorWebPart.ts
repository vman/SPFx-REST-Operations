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

import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration, ISPHttpClientOptions, ISPHttpClientBatchOptions, SPHttpClientBatch, ISPHttpClientBatchCreationOptions } from '@microsoft/sp-http';
import { IODataUser, IODataWeb } from '@microsoft/sp-odata-types';

export default class RestOperatorWebPart extends BaseClientSideWebPart<IRestOperatorWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

    this._makeSPHttpClientGETRequest();

    this._makeSPHttpClientPOSTRequest();

    this._makeSPHttpClientBatchRequest();
  }

  private _makeSPHttpClientPOSTRequest(): void {
    // Here, 'this' refers to my SPFx webpart which inherits from the BaseClientSideWebPart class.
    // Since I am calling this method from inside the class, I have access to 'this'.
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    const currentTime: string = new Date().toString();
    const spOpts: ISPHttpClientOptions = {
      body: `{ Title: 'Developer Workbench ${currentTime}', BaseTemplate: 100 }`
    };


    spHttpClient.post(`${currentWebUrl}/_api/web/lists`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        // Access properties of the response object. 
        console.log(`Status code: ${response.status}`);
        console.log(`Status text: ${response.statusText}`);

        //response.json() returns a promise so you get access to the json in the resolve callback.
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);
        });
      });
  }

  private _makeSPHttpClientGETRequest(): void {

    // Here, 'this' refers to my SPFx webpart which inherits from the BaseClientSideWebPart class.
    // Since I am calling this method from inside the class, I have access to 'this'.
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    //Since the SP Search REST API works with ODataVersion 3, we have to create a new ISPHttpClientConfiguration object with defaultODataVersion = ODataVersion.v3
    const spSearchConfig: ISPHttpClientConfiguration = {
      defaultODataVersion: ODataVersion.v3
    };

    //Override the default ODataVersion.v4 flag with the ODataVersion.v3
    const clientConfigODataV3: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith(spSearchConfig);

    //Make the REST call
    spHttpClient.get(`${currentWebUrl}/_api/search/query?querytext='sharepoint'`, clientConfigODataV3).then((response: SPHttpClientResponse) => {

      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
      });
    });

    //GET current web info
    spHttpClient.get(`${currentWebUrl}/_api/web`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((web: IODataWeb) => {

        console.log(web.Url);
      });
    });

    //GET current user information from the User Information List
    spHttpClient.get(`${currentWebUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((user: IODataUser) => {

        console.log(user.LoginName);
      });
    });

    //GET current user information from the User Profile Service
    spHttpClient.get(`${currentWebUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {

      response.json().then((userProfileProps: any) => {

        console.log(userProfileProps);
      });
    });

  }

  private _makeSPHttpClientBatchRequest(): void {

    // Here, 'this' refers to my SPFx webpart which inherits from the BaseClientSideWebPart class.
    // Since I am calling this method from inside the class, I have access to 'this'.
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    const spBatchCreationOpts: ISPHttpClientBatchCreationOptions = { webUrl: currentWebUrl };

    const spBatch: SPHttpClientBatch = spHttpClient.beginBatch(spBatchCreationOpts);

    // Queue a request to get current user's userprofile properties
    const getMyProperties: Promise<SPHttpClientResponse> = spBatch.get(`${currentWebUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClientBatch.configurations.v1);

    // Queue a request to get the title of the current web
    const getWebTitle: Promise<SPHttpClientResponse> = spBatch.get(`${currentWebUrl}/_api/web/title`, SPHttpClientBatch.configurations.v1);

    // Queue a request to create a list in the current web.
    const currentTime: string = new Date().toString();
    const batchOps: ISPHttpClientBatchOptions = {
      body: `{ Title: 'List created with SPFx batching at ${currentTime}', BaseTemplate: 100 }`
    };
    const createList: Promise<SPHttpClientResponse> = spBatch.post(`${currentWebUrl}/_api/web/lists`, SPHttpClientBatch.configurations.v1, batchOps);


    spBatch.execute().then(() => {

      getMyProperties.then((response: SPHttpClientResponse) => {
        response.json().then((props: any) => {
          console.log(props);
        });
      });

      getWebTitle.then((response: SPHttpClientResponse) => {

        response.json().then((webTitle: string) => {

          console.log(webTitle);
        });
      });

      createList.then((response: SPHttpClientResponse) => {

        response.json().then((responseJSON: any) => {

          console.log(responseJSON);
        });
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
