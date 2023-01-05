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

import * as strings from 'TransmittalSendRequestWebPartStrings';
import TransmittalSendRequest from './components/TransmittalSendRequest';
import { ITransmittalSendRequestProps } from './components/ITransmittalSendRequestProps';

export interface ITransmittalSendRequestWebPartProps {
  description: string;
}

export default class TransmittalSendRequestWebPart extends BaseClientSideWebPart<ITransmittalSendRequestProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITransmittalSendRequestProps> = React.createElement(
      TransmittalSendRequest,
      {
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        hubUrl: this.properties.hubUrl,
        redirectUrl: this.properties.redirectUrl,
        project: this.properties.project,
        notificationPreference: this.properties.notificationPreference,
        emailNotification: this.properties.emailNotification,
        userMessageSettings: this.properties.userMessageSettings,
        workflowHeaderList: this.properties.workflowHeaderList,
        documentIndexList: this.properties.documentIndexList,
        workflowDetailsList: this.properties.workflowDetailsList,
        sourceDocumentLibrary: this.properties.sourceDocumentLibrary,
        documentRevisionLogList: this.properties.documentRevisionLogList,
        transmittalCodeSettingsList: this.properties.transmittalCodeSettingsList,
        workflowTasksList: this.properties.workflowTasksList,
        revisionLevelList: this.properties.revisionLevelList,
        taskDelegationSettings: this.properties.taskDelegationSettings,
        revisionHistoryPage: this.properties.revisionHistoryPage,
        documentApprovalPage: this.properties.documentApprovalPage,
        documentReviewPage: this.properties.documentReviewPage,
        accessGroups: this.properties.accessGroups,
        departmentList: this.properties.departmentList,
        accessGroupDetailsList: this.properties.accessGroupDetailsList,
        hubsite: this.properties.hubsite,
        projectInformationListName: this.properties.projectInformationListName,
        businessUnitList: this.properties.businessUnitList,
        webpartHeader: this.properties.webpartHeader,
        siteAddress: this.properties.siteAddress,
        requestList: this.properties.requestList,
        sourceDocumentLibraryView: this.properties.sourceDocumentLibraryView,
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('documentIndexList', {
                  label: 'Document Index List'
                }),
                PropertyPaneTextField('sourceDocumentLibrary', {
                  label: 'Source Document Library'
                }),
                PropertyPaneTextField('workflowHeaderList', {
                  label: 'WorkflowHeaderList'
                }),
                PropertyPaneTextField('workflowDetailsList', {
                  label: 'Workflow Details List'
                }),
                PropertyPaneTextField('documentRevisionLogList', {
                  label: 'Document RevisionLog List'
                }),
                PropertyPaneTextField('sourceDocumentLibraryView', {
                  label: 'SourceDocument Library View'
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
                PropertyPaneTextField('accessGroups', {
                  label: 'Access Groups List'
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
                PropertyPaneTextField('workflowTasksList', {
                  label: 'Workflow Tasks List'
                }),
                PropertyPaneTextField('taskDelegationSettings', {
                  label: 'Task Delegation Settings'
                }),
                PropertyPaneTextField('accessGroupDetailsList', {
                  label: 'Access Group Details List'
                }),
                PropertyPaneTextField('departmentList', {
                  label: 'Department List'
                }),
                PropertyPaneTextField('businessUnitList', {
                  label: 'Business Unit List'
                }),
                PropertyPaneTextField('requestList', {
                  label: 'RequestList'
                }),
              ]
            },
            {
              groupName: "Pages",
              groupFields: [
                PropertyPaneTextField('documentReviewPage', {
                  label: 'Document Review Page'
                }),
                PropertyPaneTextField('documentApprovalPage', {
                  label: 'Document Approval Page'
                }),
                PropertyPaneTextField('revisionHistoryPage', {
                  label: 'Revision History Page'
                }),
              ]
            },
            {
              groupName: "LA Params",
              groupFields: [

                PropertyPaneTextField('sourceDocumentLibraryView', {
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
