import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TransmittalEditDocumentWebPartStrings';
import TransmittalEditDocument from './components/TransmittalEditDocument';
import { ITransmittalEditDocumentProps } from './components/ITransmittalEditDocumentProps';


export default class TransmittalEditDocumentWebPart extends BaseClientSideWebPart<ITransmittalEditDocumentProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITransmittalEditDocumentProps> = React.createElement(
      TransmittalEditDocument,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        hubUrl: this.properties.hubUrl,
        hubsite: this.properties.hubsite,
        redirectUrl: this.properties.redirectUrl,
        project: this.properties.project,
        notificationPreference: this.properties.notificationPreference,
        emailNotification: this.properties.emailNotification,
        userMessageSettings: this.properties.userMessageSettings,
        documentIndexList: this.properties.documentIndexList,
        businessUnit: this.properties.businessUnit,
        department: this.properties.department,
        category: this.properties.category,
        subCategory: this.properties.subCategory,
        publisheddocumentLibrary: this.properties.publisheddocumentLibrary,
        documentIdSettings: this.properties.documentIdSettings,
        documentIdSequenceSettings: this.properties.documentIdSequenceSettings,
        sourceDocumentLibrary: this.properties.sourceDocumentLibrary,
        revisionHistoryPage: this.properties.revisionHistoryPage,
        siteAddress: this.properties.siteAddress,
        sourceDocumentViewLibrary: this.properties.sourceDocumentViewLibrary,
        documentRevisionLogList: this.properties.documentRevisionLogList,
        revisionLevelList: this.properties.revisionLevelList,
        revisionSettingsList: this.properties.revisionSettingsList,
        projectInformationListName: this.properties.projectInformationListName,
        backUrl: this.properties.backUrl,
        transmittalHistory: this.properties.transmittalHistory,
        revokePage: this.properties.revokePage,
        legalEntity: this.properties.legalEntity,
        permissionMatrix: this.properties.permissionMatrix,
        departmentList: this.properties.departmentList,
        accessGroupDetailsList: this.properties.accessGroupDetailsList,
        businessUnitList: this.properties.businessUnitList,
        requestList: this.properties.requestList,
        webpartHeader: this.properties.webpartHeader,
        QDMSUrl: this.properties.QDMSUrl
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
              groupName: "Webpart Property",
              groupFields: [
                PropertyPaneTextField('webpartHeader', {
                  label: 'webpartHeader'
                }),
              ]
            },
            {
              groupName: "Current Site",
              groupFields: [
                PropertyPaneTextField('documentIndexList', {
                  label: 'Document Index List'
                }),
                PropertyPaneTextField('sourceDocumentLibrary', {
                  label: 'Source Document Library'
                }),
                PropertyPaneTextField('publisheddocumentLibrary', {
                  label: 'Published Document Library'
                }),
                PropertyPaneTextField('documentIdSettings', {
                  label: 'documentIdSettings'
                }),
                PropertyPaneTextField('documentIdSequenceSettings', {
                  label: 'documentIdSequenceSettings'
                }),
                PropertyPaneTextField('documentRevisionLogList', {
                  label: 'Document RevisionLog List'
                }),
                PropertyPaneTextField('legalEntity', {
                  label: 'Legal Entity List'
                }),
                PropertyPaneTextField('permissionMatrix', {
                  label: 'Permission Matrix'
                }),
                PropertyPaneTextField('accessGroupDetailsList', {
                  label: 'AccessGroupDetailsList'
                }),
                PropertyPaneTextField('requestList', {
                  label: 'requestList'
                }),
              ]
            },
            {
              groupName: "HubSite",
              groupFields: [
                PropertyPaneTextField('hubUrl', {
                  label: 'HubUrl'
                }),
                PropertyPaneTextField('hubsite', {
                  label: 'hubsite'
                }),
                PropertyPaneTextField('businessUnit', {
                  label: 'businessUnit'
                }),
                PropertyPaneTextField('department', {
                  label: 'department'
                }),
                PropertyPaneTextField('category', {
                  label: 'category'
                }),
                PropertyPaneTextField('subCategory', {
                  label: 'subCategory'
                }),
                PropertyPaneTextField('emailNotification', {
                  label: 'emailNotification'
                }),
                PropertyPaneTextField('notificationPreference', {
                  label: 'notificationPreference'
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: 'userMessageSettings'
                }),
                PropertyPaneTextField('QDMSUrl', {
                  label: 'QDMSUrl'
                }),
              ]
            },
            {
              groupName: "Pages",
              groupFields: [
                PropertyPaneTextField('revisionHistoryPage', {
                  label: 'RevisionHistoryPage'
                }),
                PropertyPaneTextField('transmittalHistory', {
                  label: 'TransmittalHistoryPage'
                }),
                PropertyPaneTextField('revokePage', {
                  label: 'revokePage'
                }),
              ]
            },
            {
              groupName: "LA Params",
              groupFields: [

                PropertyPaneTextField('sourceDocumentViewLibrary', {
                  label: 'Source Document View Library'
                }),
              ]
            },
            {
              groupName: "Project",
              groupFields: [
                PropertyPaneToggle('project', {
                  label: 'Project',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('revisionLevelList', {
                  label: 'Revision Level List'
                }),
                PropertyPaneTextField('revisionSettingsList', {
                  label: 'Revision Settings List'
                }),
                PropertyPaneTextField('projectInformationListName', {
                  label: 'Project Information ListName'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
