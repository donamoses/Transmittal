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

import * as strings from 'TransmittalApproveDocumentWebPartStrings';
import TransmittalApproveDocument from './components/TransmittalApproveDocument';
import { ITransmittalApproveDocumentProps } from './components/ITransmittalApproveDocumentProps';


export default class TransmittalApproveDocumentWebPart extends BaseClientSideWebPart<ITransmittalApproveDocumentProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITransmittalApproveDocumentProps> = React.createElement(
      TransmittalApproveDocument,
      {
        context: this.context,
        description: this.properties.description,
        project: this.properties.project,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        hubUrl: this.properties.hubUrl,
        notificationPreference: this.properties.notificationPreference,
        emailNotification: this.properties.emailNotification,
        userMessageSettings: this.properties.userMessageSettings,
        workflowHeaderList: this.properties.workflowHeaderList,
        documentIndexList: this.properties.documentIndexList,
        workflowDetailsList: this.properties.workflowDetailsList,
        sourceDocument: this.properties.sourceDocument,
        publishedDocument: this.properties.publishedDocument,
        documentRevisionLogList: this.properties.documentRevisionLogList,
        transmittalCodeSettingsList: this.properties.transmittalCodeSettingsList,
        workflowTasksList: this.properties.workflowTasksList,
        PermissionMatrixSettings: this.properties.PermissionMatrixSettings,
        departmentList: this.properties.departmentList,
        sourceDocumentLibrary: this.properties.sourceDocumentLibrary,
        siteAddress: this.properties.siteAddress,
        accessGroupDetailsList: this.properties.accessGroupDetailsList,
        hubsite: this.properties.hubsite,
        projectInformationListName: this.properties.projectInformationListName,
        businessUnit: this.properties.businessUnit,
        requestList: this.properties.requestList,
        webpartHeader: this.properties.webpartHeader
      }
    );

    ReactDom.render(element, this.domElement);
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
              groupName: "Hub Site",
              groupFields: [
                PropertyPaneTextField('hubUrl', {
                  label: 'HubUrl'
                }),
                PropertyPaneTextField('hubsite', {
                  label: 'hubsite'
                }),
                PropertyPaneTextField('notificationPreference', {
                  label: 'Notification Preference'
                }),
                PropertyPaneTextField('emailNotification', {
                  label: 'Email Notification'
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: 'User Message Settings'
                }),
                PropertyPaneTextField('PermissionMatrixSettings', {
                  label: 'Permission Matrix Settings List'
                }),
                PropertyPaneTextField('workflowTasksList', {
                  label: 'Workflow Tasks List'
                }),
                PropertyPaneTextField('departmentList', {
                  label: 'Department List'
                }),
                PropertyPaneTextField('businessUnit', {
                  label: 'Business Unit'
                }),
                PropertyPaneTextField('requestList', {
                  label: 'requestList'
                }),
                PropertyPaneTextField('accessGroupDetailsList', {
                  label: 'AccessGroupDetailsList'
                }),
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('documentRevisionLogList', {
                  label: 'Document RevisionLog List'
                }),
                PropertyPaneTextField('documentIndexList', {
                  label: 'Document Index List'
                }),
                PropertyPaneTextField('sourceDocument', {
                  label: 'Source Document Library'
                }),

                PropertyPaneTextField('workflowHeaderList', {
                  label: 'WorkflowHeaderList'
                }),
                PropertyPaneTextField('workflowDetailsList', {
                  label: 'Workflow Details List'
                }),
                PropertyPaneTextField('publishedDocument', {
                  label: 'Published Document Library'
                })

              ]
            },
            {
              groupName: "LA Params",
              groupFields: [

                PropertyPaneTextField('sourceDocumentLibrary', {
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
                PropertyPaneTextField('transmittalCodeSettingsList', {
                  label: 'Transmittal Code Settings List'
                }),
                PropertyPaneTextField('projectInformationListName', {
                  label: 'projectInformationListName'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
