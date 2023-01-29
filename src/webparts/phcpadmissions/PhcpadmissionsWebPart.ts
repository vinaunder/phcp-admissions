import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PhcpadmissionsWebPartStrings';
import Phcpadmissions from './components/Phcpadmissions';
import { IPhcpadmissionsProps } from './components/IPhcpadmissionsProps';

import IAdmissionItem from './models/IAdmissionItem';

import { SPHttpClient} from '@microsoft/sp-http';

export interface IPhcpadmissionsWebPartProps {
  description: string;
  list: string;
  qtd: number;
}

export default class PhcpadmissionsWebPart extends BaseClientSideWebPart<IPhcpadmissionsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _itens:IAdmissionItem[] = [];  

  private async loadDetails(): Promise<void> {
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
    this._itens = [];
    const response = await spHttpClient.get(
      `${currentWebUrl}/_api/web/lists/getbytitle('${this.properties.list}')/items?$orderby=Created desc&$top=${this.properties.qtd}`,
      SPHttpClient.configurations.v1);

    const currentPageDetails = await response.json();

    if(currentPageDetails?.value){
      currentPageDetails.value.forEach((item: IAdmissionItem) => {
        this._itens.push({
          Id: item.Id,
          Title: item.Title,
          message: item.message,
          Created: item.Created
        });
      });
    }
    
  }

  public async  render(): Promise<void> {
    await this.loadDetails();
    const element: React.ReactElement<IPhcpadmissionsProps> = React.createElement(
      Phcpadmissions,
      {
        webparttitle: this.properties.description,
        itens:this._itens,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                }),
                PropertyPaneTextField('list', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneSlider('qtd', {
                  label: strings.ListNameFieldLabel,
                  min: 1,
                  max: 12,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
