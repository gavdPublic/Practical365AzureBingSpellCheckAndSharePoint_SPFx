import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

// This are the JS Libraries to make HTTP calls
import { 
  HttpClient, 
  SPHttpClient, 
  HttpClientConfiguration, 
  HttpClientResponse, 
  ODataVersion, 
  IHttpClientConfiguration, 
  IHttpClientOptions, 
  ISPHttpClientOptions 
} from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Prac365BingSpellCheckWpWebPart.module.scss';
import * as strings from 'Prac365BingSpellCheckWpWebPartStrings';

export interface IPrac365BingSpellCheckWpWebPartProps {
  description: string;
}

export default class Prac365BingSpellCheckWpWebPart extends BaseClientSideWebPart<IPrac365BingSpellCheckWpWebPartProps> {

  // This is the Azure Function URL and, if necessary, authentication code
  protected AzureFunctionUrl: 
  string = "https://[domain].azurewebsites.net/api/[FuntionName]HttpTriggerFunction";

  protected runFunction(): void {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");

    // Gather the information from the form fields
    var stringToCheck: string = (<HTMLInputElement>document.getElementById("txtStringToCheck")).value;
    var myMarket: string = (<HTMLInputElement>document.getElementById("txtMarket")).value;
    var myMode: string = (<HTMLInputElement>document.getElementById("txtMode")).value;

    // Just log some information for debugging purposses
    console.log(`StringToCheck: '${stringToCheck}' - Market: '${myMarket}' - Mode: '${myMode}'`);

    // This are the options for the HTTP call
    const callOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ strcheck: '${stringToCheck}', market: '${myMarket}', mode: '${myMode}' }`
    };

    // Create the responce object...
    let responseText: string = "";
    let responceMessage: HTMLElement = document.getElementById("responseContainer");
    responceMessage.innerText = "";  // Cleanup the interface

    // And make a POST request to the Function
    this.context.httpClient.post(this.AzureFunctionUrl, HttpClient.configurations.v1, callOptions).then((response: HttpClientResponse) => {
       response.json().then((responseJSON: JSON) => {
          responseText = JSON.stringify(responseJSON);
          if (response.ok) {
            responceMessage.style.color = "aqua";
          }

          responceMessage.innerText = responseText;
        })
        .catch ((response: any) => {
          let errorMessage: string = `Error calling ${this.AzureFunctionUrl} = ${response.message}`;
          responceMessage.style.color = "yellow";
          responceMessage.innerText = errorMessage;
        });
    });
  }

  // This is the interface
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.prac365BingSpellCheckWp}">
      <div class="${styles.container}">
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span class="ms-font-xl ms-fontColor-white">Check the spelling of a text</span>
            <div class="${styles.controlRow}">
              <span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Text To Check:</span>
              <input type="text" id="txtStringToCheck"></input>
            </div>
            <div class="${styles.controlRow}">
              <span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Market:</span>
              <input type="text" id="txtMarket"></input>
            </div>
            <div class="${styles.controlRow}">
              <span class="ms-font-l ms-fontColor-white ${styles.controlLabel}">Mode:</span>
              <input type="text" id="txtMode"></input>
            </div>
            <div class="${styles.buttonRow}"></div>
            <button id="btnRunFunction" class="${styles.button}">Check Spelling</button>
            <div id="responseContainer" class="${styles.resultRow}"></div>
          </div>
        </div>
      </div>
    </div>`;

    // The Event Handler for the Button  
    document.getElementById("btnRunFunction").onclick = this.runFunction.bind(this);
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
