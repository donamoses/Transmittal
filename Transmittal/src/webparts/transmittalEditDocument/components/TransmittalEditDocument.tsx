import * as React from 'react';
import styles from './TransmittalEditDocument.module.scss';
import { ITransmittalEditDocumentProps, ITransmittalEditDocumentState } from './ITransmittalEditDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import SimpleReactValidator from 'simple-react-validator';
import { BaseService } from '../services';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IChoiceGroupOption, IChoiceGroupStyles, IconButton, IDropdownOption, IIconProps, ITooltipHostStyles, Label, MessageBar, Pivot, PivotItem, PrimaryButton, ProgressIndicator, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Iframe from 'react-iframe';
import * as moment from 'moment';
import * as _ from 'lodash';
import { IHttpClientOptions, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
export default class TransmittalEditDocument extends React.Component<ITransmittalEditDocumentProps, ITransmittalEditDocumentState, {}> {
  private validator: SimpleReactValidator;
  private _Service: BaseService;
  // private reqWeb;
  private currentEmail: any;
  private currentId: any;
  private currentUser: any;
  private getSelectedReviewers: any[] = [];
  private revisionHistoryUrl: string;
  private revokeUrl: string;
  private today: any;
  private directPublish: string;
  private createDocument: string;
  private editDocument: string;
  private revokeExpiry: string;
  private sourceDocumentID: any;
  private sourceDocumentLibraryId: any;
  private mode: any;
  private documentIndexID: any;
  private revokeExpiryError: string;
  private documentNameExtension: any;
  // private valid = "ok";
  // private accessForTitleRename;
  // private accessForRevoke;
  private noAccess: string;
  // private departmentExists;
  private postUrl: string;
  // private siteUrl;
  // private indexUrl;
  private isDocument: string;
  private myfile: any;
  private permissionpostUrl: string;
  public constructor(props: ITransmittalEditDocumentProps) {
    super(props);
    this.state = {
      hideProject: "none",
      hideDoc: "",
      hidePublish: "none",
      hideExpiry: "",
      projectEditDocumentView: "none",
      createDocumentProject: "none",
      createDocumentView: "",
      qdmsEditDocumentView: "none",
      revokeExpiryView: "none",
      hideCreate: "",
      messageBar: "none",
      hidebutton: "",
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      cancelConfirmMsg: "none",
      confirmDialog: true,
      accessDeniedMessageBar: "none",
      title: "",
      businessUnitID: null,
      departmentId: null,
      categoryId: null,
      subCategoryKey: "",
      legalEntityId: null,
      createDocument: false,
      replaceDocumentCheckbox: false,
      templateDocuments: "",
      directPublishCheck: false,
      approvalDate: "",
      publishOptionKey: "",
      reviewersName: "",
      expiryCheck: "",
      expiryDate: null,
      expiryLeadPeriod: "",
      criticalDocument: false,
      templateDocument: false,
      titleReadonly: true,
      saveDisable: false,
      legalEntityOption: [],
      businessUnitOption: [],
      departmentOption: [],
      approvalDateEdit: null,
      categoryOption: [],
      businessUnitCode: "",
      departmentCode: "",
      subCategoryArray: [],
      subCategoryId: null,
      reviewers: [],
      approver: null,
      approverEmail: "",
      approverName: "",
      owner: "",
      ownerEmail: "",
      ownerName: "",
      templateId: "",
      publishOption: "Native",
      incrementSequenceNumber: "",
      documentid: "",
      documentName: "",
      businessUnit: "",
      category: "",
      subCategory: "",
      department: "",
      newDocumentId: "",
      sourceDocumentId: "",
      templateKey: "",
      dcc: null,
      dccEmail: "",
      dccName: "",
      revisionCoding: "",
      revisionLevel: "",
      transmittalCheck: false,
      externalDocument: false,
      revisionCodingId: null,
      revisionLevelId: null,
      revisionLevelArray: [],
      revisionSettingsArray: [],
      categoryCode: "",
      projectName: "",
      projectNumber: "",
      currentRevision: "",
      previousRevisionItemID: null,
      revisionItemID: "",
      newRevision: "",
      hideloader: true,
      legalEntity: "",
      updateDisable: false,
      hideLoading: true,
      workflowStatus: "",
      leadmsg: "none",
      invalidQueryParam: "",
      isdocx: "none",
      nodocx: "",
      insertdocument: "none",
      loaderDisplay: "",
      checkdirect: "none",
      hideDirect: "none",
      validApprover: "none",
      createDocumentCheckBoxDiv: "",
      replaceDocument: "none",
      hideSelectTemplate: "none",
      validDocType: "none",
      checkrename: "none",
      subContractorNumber: "",
      customerNumber: "",
      linkToDoc: "",
      hideCreateLoading: "none",
      norefresh: "none",
      upload: false,
      template: false,
      hideupload: "none",
      sourceId: "",
      hidetemplate: "none",
      hidesource: "none",
      uploadOrTemplateRadioBtn: "",
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl, window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._selectedOwner = this._selectedOwner.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    this._templatechange = this._templatechange.bind(this);
    this._onDirectPublishChecked = this._onDirectPublishChecked.bind(this);
    this._onApprovalDatePickerChange = this._onApprovalDatePickerChange.bind(this);
    this._publishOptionChange = this._publishOptionChange.bind(this);
    this._onExpDatePickerChange = this._onExpDatePickerChange.bind(this);
    this._expLeadPeriodChange = this._expLeadPeriodChange.bind(this);
    this._onCriticalChecked = this._onCriticalChecked.bind(this);
    this._onTemplateChecked = this._onTemplateChecked.bind(this);
    this._project = this._project.bind(this);
    this._sourcechange = this._sourcechange.bind(this);
    this._revisionCodingChange = this._revisionCodingChange.bind(this);
    this._updateDocumentIndex = this._updateDocumentIndex.bind(this);
    this._selectedDCC = this._selectedDCC.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._generateNewRevision = this._generateNewRevision.bind(this);
    this._updateSourceDocument = this._updateSourceDocument.bind(this);
    this._updateDocument = this._updateDocument.bind(this);
    this._updatePublishDocument = this._updatePublishDocument.bind(this);
    this._onUpdateClick = this._onUpdateClick.bind(this);
    this._bindDataEditProject = this._bindDataEditProject.bind(this);
    this._updateWithoutDocument = this._updateWithoutDocument.bind(this);
    this._add = this._add.bind(this);
    this._checkRename = this._checkRename.bind(this);
    this._onReplaceDocumentChecked = this._onReplaceDocumentChecked.bind(this);
    this._subContractorNumberChange = this._subContractorNumberChange.bind(this);
    this._CustomerNumberChange = this._CustomerNumberChange.bind(this);
  }
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: { required: "This field is mandatory" }
    });
  }
  public async componentDidMount() {
    //Get Current User
    const user = await this._Service.getCurrentUser();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    this.currentUser = user.Title;
    //Get Today
    this.today = new Date();
    this.setState({ approvalDate: this.today });
    //for getting  sourcedoument library ID
    this._Service.getLibrary(this.props.siteUrl, this.props.sourceDocumentViewLibrary).then((list: any) => {
      this.sourceDocumentLibraryId = list.Id;

    })
    this._queryParamGetting();
  }
  //Search Query
  private async _queryParamGetting() {
    this.setState({ accessDeniedMessageBar: "none", createDocumentView: "none", createDocumentProject: "none", qdmsEditDocumentView: "none", projectEditDocumentView: "none", revokeExpiryView: "none", });
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentindexid = params.get('did');
    let mode = params.get('mode');
    this.documentIndexID = documentindexid;
    //Project
    if (this.props.project) {
      //Edit
      if (documentindexid != "" && documentindexid != null) {
        this._Service.getIndexdata(this.props.siteUrl, this.props.documentIndexList, Number(documentindexid))
          .then(DocumentStatus => {
            this.sourceDocumentID = DocumentStatus.SourceDocumentID;
            if ((DocumentStatus.WorkflowStatus != "Under Review" && DocumentStatus.WorkflowStatus != "Under Approval" && DocumentStatus.TransmittalStatus != "Ongoing")) {
              if (DocumentStatus.DocumentStatus == "Active") {
                this.setState({ accessDeniedMessageBar: "none", qdmsEditDocumentView: "none", projectEditDocumentView: "none" });
                //Permission handiling 
                this.setState({
                  qdmsEditDocumentView: "none", projectEditDocumentView: "", accessDeniedMessageBar: "none", loaderDisplay: "none"
                });
                this._bindDataEditProject(this.documentIndexID);

                // this._checkPermission('Project_EditDocument');
              }
              else {
                this.setState({
                  qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "", loaderDisplay: "none",
                  statusMessage: { isShowMessage: true, message: "Document is not active right now", messageType: 1 },
                });
                setTimeout(() => {
                  this.setState({ accessDeniedMessageBar: 'none', });
                  window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                }, 10000);
              }
            }
            else {
              this.setState({
                qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "", loaderDisplay: "none",
                statusMessage: { isShowMessage: true, message: "Document is already gone in a workflow", messageType: 1 },
              });
              setTimeout(() => {
                this.setState({ accessDeniedMessageBar: 'none', });
                window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
              }, 10000);
            }
          });
      }
      else {
        this.setState({
          qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "",
          statusMessage: { isShowMessage: true, message: this.state.invalidQueryParam, messageType: 4 },
        });
        setTimeout(() => {
          this.setState({ accessDeniedMessageBar: 'none', });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
    }
    else {
      if (documentindexid != "" && documentindexid != null) {
        this._Service.getIndexdataa(this.props.siteUrl, this.props.documentIndexList, Number(documentindexid))
          .then(DocumentStatus => {
            this.sourceDocumentID = DocumentStatus.SourceDocumentID;
            if ((DocumentStatus.WorkflowStatus != "Under Review" && DocumentStatus.WorkflowStatus != "Under Approval")) {
              if (DocumentStatus.DocumentStatus == "Active") {
                this.setState({ accessDeniedMessageBar: "none", qdmsEditDocumentView: "none", projectEditDocumentView: "none" });
                //Permission handiling 
                this.setState({
                  qdmsEditDocumentView: "", projectEditDocumentView: "none", accessDeniedMessageBar: "none", loaderDisplay: "none"
                });
                this._bindDataEditQdms(this.documentIndexID);
                // this._accessGroups('QDMS_EditDocument');
              }
              else {
                this.setState({
                  qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "", loaderDisplay: "none",
                  statusMessage: { isShowMessage: true, message: "Document is not active right now", messageType: 1 },
                });
                setTimeout(() => {
                  this.setState({ accessDeniedMessageBar: 'none', });
                  window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                }, 10000);
              }
            }
            else {
              this.setState({
                qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "",
                statusMessage: { isShowMessage: true, message: "Document is already gone in a workflow", messageType: 1 },
              });
              setTimeout(() => {
                this.setState({ accessDeniedMessageBar: 'none', });
                window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
              }, 10000);
            }
          });
      }
      else {
        this.setState({
          qdmsEditDocumentView: "none", projectEditDocumentView: "none", accessDeniedMessageBar: "",
          statusMessage: { isShowMessage: true, message: this.state.invalidQueryParam, messageType: 1 },
        });
        setTimeout(() => {
          this.setState({ accessDeniedMessageBar: 'none', });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
    }
    this._userMessageSettings();
  }
  // Bind data in qdms
  public async _bindDataEditQdms(documentindexid: any) {
    this._checkRename('QDMS_RenameDocument');
    const indexItems = await this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, parseInt(documentindexid));

    console.log("dataForEdit", indexItems);
    let tempReviewers: any[] = [];
    let temReviewersID: any[] = [];
    let items;
    let expand;
    let subcategoryArray = [];
    let sorted_subcategory: any[];
    if (documentindexid != "" && documentindexid != null) {
      items = "Title,Owner/Title,Owner/ID,Owner/EMail,SubCategoryID,WorkflowStatus,SourceDocument,SubCategory,Approver/Title,Approver/ID,ApprovedDate,BusinessUnit,BusinessUnitID,Category,CategoryID,DepartmentName,DepartmentID,DocumentID,DocumentName,ExpiryDate,Reviewers/ID,Reviewers/Title,ExpiryLeadPeriod,CategoryID,CriticalDocument,Template,PublishFormat,ApprovedDate,DirectPublish,CreateDocument,LegalEntity";
      expand = "Owner,Approver,Reviewers";
      this._Service.getIndexdatabind(this.props.siteUrl, this.props.documentIndexList, documentindexid, items, expand)
        .then(async dataForEdit => {
          console.log("dataForEdit", dataForEdit);
          this.setState({
            title: dataForEdit.Title,
            documentid: dataForEdit.DocumentID,
            documentName: dataForEdit.DocumentName,
            businessUnit: dataForEdit.BusinessUnit,
            department: dataForEdit.DepartmentName,
            category: dataForEdit.Category,
            ownerName: dataForEdit.Owner.Title,
            expiryLeadPeriod: dataForEdit.ExpiryLeadPeriod,
            owner: dataForEdit.Owner.ID,
            ownerEmail: dataForEdit.Owner.EMail,
            legalEntity: dataForEdit.LegalEntity,
            subCategory: dataForEdit.SubCategory,
            businessUnitID: dataForEdit.BusinessUnitID,
            departmentId: dataForEdit.DepartmentID
          });
          if (indexItems.ApproverId != null) {
            this.setState({
              approver: dataForEdit.Approver.ID,
              approverName: dataForEdit.Approver.Title
            });
          }
          if (dataForEdit.SourceDocument != null) {
            this.setState({
              linkToDoc: dataForEdit.SourceDocument.Url,
            });
          }
          for (var k in dataForEdit.Reviewers) {
            temReviewersID.push(dataForEdit.Reviewers[k].ID);
            this.setState({
              reviewers: temReviewersID,
            });
            tempReviewers.push(dataForEdit.Reviewers[k].Title);
          }

          if (indexItems.SubCategoryID != null) {
            this.setState({
              subCategoryId: parseInt(dataForEdit.SubCategoryID)
            });
          }
          if (dataForEdit.ExpiryDate != null) {
            let date = new Date(dataForEdit.ExpiryDate);
            this.setState({ expiryDate: date, expiryCheck: true, hideExpiry: "" });
          }
          if (dataForEdit.CriticalDocument == true) {
            this.setState({ criticalDocument: true });
          }
          if (dataForEdit.CreateDocument == true) {
            this.setState({ createDocument: true, hideCreate: "", createDocumentCheckBoxDiv: "none", replaceDocument: "", hidePublish: "none", hideDoc: "none" });
            this.isDocument = "Yes";
          }
          if (dataForEdit.CreateDocument == false) {
            this._checkdirectPublish('QDMS_DirectPublish');

          }
          if (dataForEdit.Template == true) {
            this.setState({ templateDocument: true });
          }
          if (dataForEdit.DirectPublish == true) {
            let date = new Date(dataForEdit.ApprovedDate);
            this.setState({ directPublishCheck: true, hidePublish: "none", publishOptionKey: dataForEdit.PublishFormat, approvalDateEdit: date });
          }
          this.setState({
            reviewersName: tempReviewers,
          });

        });
    }

  }
  // Bind data from project
  public _bindDataEditProject(documentindexid: any) {
    this._checkRename('Project_RenameDocument');
    this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, documentindexid)
      .then(async indexItems => {
        console.log("dataForEdit", indexItems);
        let tempReviewers: any[] = [];
        let temReviewersID: any[] = [];
        let items;
        let expand;
        let subcategoryArray: any[] = [];
        let sorted_subcategory: any[];
        if (indexItems.CategoryID != null) {
          await this._Service.gethubListItems(this.props.hubUrl, this.props.subCategory)
            .then(subcategory => {
              for (let i = 0; i < subcategory.length; i++) {
                if (subcategory[i].CategoryId == indexItems.CategoryID) {
                  let subcategorydata = {
                    key: subcategory[i].ID,
                    text: subcategory[i].SubCategory,
                  };
                  subcategoryArray.push(subcategorydata);
                }
              }
              sorted_subcategory = _.orderBy(subcategoryArray, 'text', ['asc']);
              this.setState({
                subCategoryArray: sorted_subcategory
              });
            });
        }
        if (documentindexid != "" && documentindexid != null && this.mode != "expiry") {
          items = "Title,SubCategoryID,SubCategory,BusinessUnit,Category,DepartmentName,SourceDocument,CustomerDocumentNo,SubcontractorDocumentNo,Owner/Title,Owner/ID,Owner/EMail,Approver/Title,Approver/ID,ApprovedDate,WorkflowStatus,DocumentID,DocumentName,ExpiryDate,Reviewers/ID,Reviewers/Title,ExpiryLeadPeriod,CategoryID,CriticalDocument,CreateDocument,Template,PublishFormat,ApprovedDate,RevisionCoding/ID,RevisionCoding/Title,RevisionLevel/ID,RevisionLevel/Title,DocumentController/ID,DocumentController/Title,TransmittalDocument,ExternalDocument,DirectPublish";
          expand = "Owner,Approver,Reviewers,RevisionCoding,RevisionLevel,DocumentController";
          this._Service.getIndexdatabind(this.props.siteUrl, this.props.documentIndexList, documentindexid, items, expand)
            .then(dataForEdit => {
              console.log("dataForEdit", dataForEdit);
              this.setState({
                title: dataForEdit.Title,
                documentid: dataForEdit.DocumentID,
                businessUnit: dataForEdit.BusinessUnit,
                department: dataForEdit.DepartmentName,
                category: dataForEdit.Category,
                ownerName: dataForEdit.Owner.Title,
                approverName: dataForEdit.Approver.Title,
                expiryLeadPeriod: dataForEdit.ExpiryLeadPeriod,
                owner: dataForEdit.Owner.ID,
                workflowStatus: dataForEdit.WorkflowStatus,
                subCategory: dataForEdit.SubCategory,
                documentName: dataForEdit.DocumentName,
                subContractorNumber: dataForEdit.SubcontractorDocumentNo,
                customerNumber: dataForEdit.CustomerDocumentNo,

              });
              for (var k in dataForEdit.Reviewers) {
                temReviewersID.push(dataForEdit.Reviewers[k].ID);
                this.setState({
                  reviewers: temReviewersID,
                });
                tempReviewers.push(dataForEdit.Reviewers[k].Title);
              }
              if (dataForEdit.SourceDocument != null) {
                this.setState({
                  linkToDoc: dataForEdit.SourceDocument.Url,
                });
              }
              if (dataForEdit.ExpiryDate != null) {
                let date = new Date(dataForEdit.ExpiryDate);
                this.setState({ expiryDate: date, expiryCheck: true, hideExpiry: "", });
              }
              if (dataForEdit.CriticalDocument == true) {
                this.setState({ criticalDocument: true });
              }
              if (dataForEdit.CreateDocument == true) {
                this.isDocument = "Yes";
                this.setState({ createDocument: true, hideCreate: "", createDocumentCheckBoxDiv: "none", replaceDocument: "", hidePublish: "none", hideDoc: "none" });
              }
              if (dataForEdit.CreateDocument == false) {

                this._checkdirectPublish('Project_DirectPublish');
              }
              if (dataForEdit.Template == true) {
                this.setState({ templateDocument: true });
              }
              if (dataForEdit.DirectPublish == true) {
                let date = new Date(dataForEdit.ApprovedDate);
                this.setState({ directPublishCheck: true, hidePublish: "none", publishOptionKey: dataForEdit.PublishFormat, approvalDateEdit: date });
              }
              if (indexItems.RevisionCodingId != null) {
                this.setState({
                  revisionCodingId: dataForEdit.RevisionCoding.ID
                });
              }
              if (indexItems.SubCategoryID != null) {
                this.setState({
                  subCategoryId: parseInt(dataForEdit.SubCategoryID)
                });
              }
              if (indexItems.RevisionLevelId != null) {
                this.setState({
                  revisionLevelId: dataForEdit.RevisionLevel.ID
                });
              }
              this.setState({
                reviewersName: tempReviewers,
              });

              this._project();
              if (dataForEdit.ExternalDocument == true) {
                this.setState({ externalDocument: true });
              }
              if (dataForEdit.TransmittalDocument == true) {
                this.setState({ transmittalCheck: true });
              }
              if (indexItems.DocumentControllerId != null) {
                this.setState({
                  dcc: dataForEdit.DocumentController.ID,
                  dccName: dataForEdit.DocumentController.Title
                });
              }
              if (indexItems.ApproverId != null) {
                this.setState({
                  approver: dataForEdit.Approver.ID,
                  approverName: dataForEdit.Approver.Title
                });
              }

            });
        }

      });
  }
  //Bind data on Project
  public async _project() {
    let revisionLevelArray = [];
    let sorted_RevisionLevel = [];
    let revisionSettingsArray = [];
    let sorted_RevisionSettings = [];
    //Get Revision Level
    const revisionLevelItem: any = await this._Service.getDrpdwnListItems(this.props.siteUrl, this.props.revisionLevelList)
    for (let i = 0; i < revisionLevelItem.length; i++) {
      let revisionLevelItemdata = {
        key: revisionLevelItem[i].ID,
        text: revisionLevelItem[i].Title
      };
      revisionLevelArray.push(revisionLevelItemdata);
    }
    sorted_RevisionLevel = _.orderBy(revisionLevelArray, 'text', ['asc']);
    //Get RevisionSettings
    const revisionSettingsItem: any = await this._Service.getDrpdwnListItems(this.props.siteUrl, this.props.revisionSettingsList)
    for (let i = 0; i < revisionSettingsItem.length; i++) {
      let revisionSettingsItemdata = {
        key: revisionSettingsItem[i].ID,
        text: revisionSettingsItem[i].Title
      };
      revisionSettingsArray.push(revisionSettingsItemdata);
    }
    sorted_RevisionSettings = _.orderBy(revisionSettingsArray, 'text', ['asc']);
    //Get Project Information
    const projectInformation = await this._Service.getListItems(this.props.siteUrl, this.props.projectInformationListName);
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
        if (projectInformation[k].Key == "ProjectName") {
          this.setState({
            projectName: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key == "ProjectNumber") {
          this.setState({
            projectNumber: projectInformation[k].Title,
          });
        }
      }
    }
    this.setState({
      revisionSettingsArray: sorted_RevisionSettings,
      revisionLevelArray: sorted_RevisionLevel
    });
  }
  // Check permission to rename
  public async _checkRename(type: any) {
    this.setState({ checkrename: "" });
    const laUrl = await this._Service.getrename(this.props.hubUrl, this.props.requestList)
    console.log("Posturl", laUrl[0].PostUrl);
    this.permissionpostUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.permissionpostUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': type,
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.currentEmail

    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] == "Valid") {
        this.setState({
          titleReadonly: false,
          checkrename: "none"
        });
      }
      else {
        this.setState({
          titleReadonly: true,
          checkrename: "none"
        });
      }
    }

    else { }
  }
  // On direct publih checked
  public async _checkdirectPublish(type: any) {
    const laUrl = await this._Service.getdirectpublish(this.props.hubUrl, this.props.requestList)
    console.log("Posturl", laUrl[0].PostUrl);
    this.permissionpostUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.permissionpostUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': type,
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.currentEmail

    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] == "Valid") {
        this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "none" });
      }
      else {
        this.setState({ checkdirect: "none", hideDirect: "none", hidePublish: "none" });
      }
    }
    else { }
  }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.gethubUserMessageListItems(this.props.hubUrl, this.props.userMessageSettings);
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "CreateDocumentSuccess") {
        var successmsg = userMessageSettings[i].Message;
        this.createDocument = replaceString(successmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "DirectPublishSuccess") {
        var publishmsg = userMessageSettings[i].Message;
        this.directPublish = replaceString(publishmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "EditDocumentSuccess") {
        var editmsg = userMessageSettings[i].Message;
        this.editDocument = replaceString(editmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "DocumentRevokeSuccess") {
        var revokemsg = userMessageSettings[i].Message;
        this.revokeExpiry = replaceString(revokemsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "DocumentRevokeError") {
        var revokeErrormsg = userMessageSettings[i].Message;
        this.revokeExpiryError = replaceString(revokeErrormsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "NoAccess") {
        this.noAccess = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidQueryParams") {
        this.setState({
          invalidQueryParam: userMessageSettings[i].Message,
        });
      }
    }
  }
  //Title Change
  public _titleChange = (ev: React.FormEvent<HTMLInputElement>, title?: string) => {
    this.setState({ title: title || '', saveDisable: false });
  }
  //Owner Change
  public _selectedOwner = (items: any[]) => {
    let ownerEmail;
    let ownerName;
    let getSelectedOwner = [];
    for (let item in items) {
      ownerEmail = items[item].secondaryText,
        ownerName = items[item].text,
        getSelectedOwner.push(items[item].id);
    }
    this.setState({ owner: getSelectedOwner[0], ownerEmail: ownerEmail, ownerName: ownerName, saveDisable: false });
  }
  //DCC Change
  public _selectedDCC = (items: any[]) => {
    let dccEmail;
    let dccName;
    let getSelectedDCC = [];
    for (let item in items) {
      dccEmail = items[item].secondaryText,
        dccName = items[item].text,
        getSelectedDCC.push(items[item].id);
    }
    this.setState({
      dcc: getSelectedDCC[0],
      dccEmail: dccEmail,
      dccName: dccName
    });
  }
  //Reviewer Change
  public _selectedReviewers = (items: any[]) => {
    this.getSelectedReviewers = [];
    for (let item in items) {
      this.getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: this.getSelectedReviewers });
  }
  //Approver Change
  public _selectedApprover = async (items: any[]) => {
    let approverEmail;
    let approverName;
    let getSelectedApprover = [];
    if (this.props.project) {
      for (let item in items) {
        approverEmail = items[item].secondaryText,
          approverName = items[item].text,
          getSelectedApprover.push(items[item].id);
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
    }
    else {
      this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });
      if (this.state.businessUnitID != null) {
        const businessUnit = await this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.businessUnit);
        for (let i = 0; i < businessUnit.length; i++) {
          if (businessUnit[i].ID == this.state.businessUnitID) {
            const approve = await this._Service.getByEmail(businessUnit[i].Approver.EMail);
            approverEmail = businessUnit[i].Approver.EMail;
            approverName = businessUnit[i].Approver.Title;
            getSelectedApprover.push(approve.Id);
          }
        }
      }
      else {
        const departments = await this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.department);
        for (let i = 0; i < departments.length; i++) {
          if (departments[i].ID == this.state.departmentId) {
            const deptapprove = await this._Service.getByEmail(departments[i].Approver.EMail);
            approverEmail = departments[i].Approver.EMail;
            approverName = departments[i].Approver.Title;
            getSelectedApprover.push(deptapprove.Id);
          }
        }
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
      setTimeout(() => {
        this.setState({ validApprover: "none" });
      }, 5000);
    }
  }
  //Expiry Date Change
  public _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ expiryDate: date });
  }
  //Expiry Lead Period Change
  public _expLeadPeriodChange = (ev: React.FormEvent<HTMLInputElement>, expiryLeadPeriod?: string) => {
    let LeadPeriodformat = /^[0-9]*$/;
    if (expiryLeadPeriod.match(LeadPeriodformat)) {
      if (Number(expiryLeadPeriod) < 101) {
        this.setState({ expiryLeadPeriod: expiryLeadPeriod || '', leadmsg: "none" });
      }
      else {
        this.setState({ leadmsg: "" });
      }
    }
    else {
      this.setState({ leadmsg: "" });
    }
  }
  // On replace document checked
  public _onReplaceDocumentChecked = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    let upload;
    this.setState({ validDocType: "none" });
    if (isChecked) {
      this.setState({
        hideSelectTemplate: "none",
        hideDoc: "none",
        hideupload: "",
        replaceDocumentCheckbox: true,
      });
    }
    else {
      this.setState({
        hideSelectTemplate: "none",
        hideDoc: "none",
        hideupload: "none",
        replaceDocumentCheckbox: false,
      });
      if (this.props.project) {
        upload = "#editproject";
      }
      else {
        upload = "#editqdms";
      }
      (document.querySelector(upload) as HTMLInputElement).value = null;
    }
  }
  private onUploadOrTemplateRadioBtnChange = async (ev: React.FormEvent<HTMLInputElement>, option?: IChoiceGroupOption) => {
    let publishedDocumentArray = [];
    let sorted_PublishedDocument: any[];
    this.setState({
      uploadOrTemplateRadioBtn: option.key,
      createDocument: true
    });
    if (option.key === "Upload") {
      this.setState({ upload: true, hideupload: "", template: false, hidesource: "none", hidetemplate: "none" });
    }
    if (option.key === "Template") {
      let publishedDocumentArray = [];
      let sorted_PublishedDocument: any[];
      let qdms = window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl;
      // this.QDMSUrl = window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.QDMSUrl;
      console.log("site :" + this.props.siteUrl);
      console.log("qdms :" + qdms);
      if (!this.props.project) {
        if (this.props.siteUrl === qdms) {
          this.setState({ hidesource: "none" })
        }

        else {
          this.setState({ hidesource: "" })
        }
        this.setState({ template: true, upload: false, hideupload: "none", hidetemplate: "" });
        let publishedDocument: any[] = await this._Service.getLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary);
        for (let i = 0; i < publishedDocument.length; i++) {
          if (publishedDocument[i].Template === true) {
            let publishedDocumentdata = {
              key: publishedDocument[i].ID,
              text: publishedDocument[i].DocumentName,
            };
            publishedDocumentArray.push(publishedDocumentdata);
          }
        }
        sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
        this.setState({ templateDocuments: sorted_PublishedDocument });
      }
      else {
        this.setState({ template: true, hidesource: "", upload: false, hideupload: "none", hidetemplate: "" });
      }

    }
  }
  // On upload document change
  public _add(e: any) {
    this.setState({ insertdocument: "none", validDocType: "none" });
    this.myfile = e.target.value;
    let upload;
    let type;
    let doctype;
    this.isDocument = "Yes";
    if (this.props.project) {
      upload = "#editproject";
    }
    else {
      upload = "#editqdms";
    }
    let myfile = (document.querySelector(upload) as HTMLInputElement).files[0];
    console.log(myfile);
    this.isDocument = "Yes";
    var splitted = myfile.name.split(".");
    type = splitted[splitted.length - 1];
    if (this.state.replaceDocumentCheckbox == true) {
      var docsplitted = this.state.documentName.split(".");
      doctype = docsplitted[docsplitted.length - 1];
      if (doctype != type) {
        this.setState({ validDocType: "" });
        (document.querySelector(upload) as HTMLInputElement).value = null;
      }
    }
    if (type == "docx") {
      this.setState({ isdocx: "", nodocx: "none" });
    }
    else {
      this.setState({ isdocx: "none", nodocx: "" });
    }
  }
  public async _sourcechange(option: { key: any; text: any }) {
    this.setState({ hidetemplate: "", sourceId: option.key });
    let publishedDocumentArray = [];
    let sorted_PublishedDocument: any[];
    // this.QDMSUrl = Web(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.QDMSUrl);
    //alert(this.QDMSUrl);
    if (option.key == "QDMS") {
      let publishedDocument: any[] = await this._Service.getqdmsLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template == true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
    }
    else {
      let publishedDocument: any[] = await this._Service.getLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary);
      for (let i = 0; i < publishedDocument.length; i++) {
        if (publishedDocument[i].Template == true) {
          let publishedDocumentdata = {
            key: publishedDocument[i].ID,
            text: publishedDocument[i].DocumentName,
          };
          publishedDocumentArray.push(publishedDocumentdata);
        }
      }
      sorted_PublishedDocument = _.orderBy(publishedDocumentArray, 'text', ['asc']);
    }
    this.setState({ templateDocuments: sorted_PublishedDocument, sourceId: option.key });
  }
  //Template change
  public async _templatechange(option: { key: any; text: any }) {
    this.setState({ insertdocument: "none" });
    this.setState({ templateId: option.key, templateKey: option.text });
    let type: any;
    let publishName: any;
    this.isDocument = "Yes";
    if (this.state.sourceId == "QDMS") {
      await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary)
        .then(publishdoc => {
          console.log(publishdoc);
          for (let i = 0; i < publishdoc.length; i++) {
            if (publishdoc[i].Id == this.state.templateId) {

              publishName = publishdoc[i].LinkFilename;
            }
          }
          var split = publishName.split(".", 2);
          type = split[1];
          if (type == "docx") {
            this.setState({ isdocx: "", nodocx: "none" });
          }
          else {
            this.setState({ isdocx: "none", nodocx: "" });
          }
        });
    }
    else {
      await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary)
        .then(publishdoc => {
          console.log(publishdoc);
          for (let i = 0; i < publishdoc.length; i++) {
            if (publishdoc[i].Id == this.state.templateId) {
              publishName = publishdoc[i].LinkFilename;
            }
          }
          var split = publishName.split(".", 2);
          type = split[1];
          if (type == "docx") {
            this.setState({ isdocx: "", nodocx: "none" });
          }
          else {
            this.setState({ isdocx: "none", nodocx: "" });
          }
        });
    }
  }
  //Direct Publish change
  private _onDirectPublishChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ hidePublish: "", directPublishCheck: true, approvalDate: new Date() });
    }
    else if (!isChecked) {
      this.setState({ hidePublish: "none", directPublishCheck: false, approvalDate: new Date(), publishOption: "" });
    }
  }
  //Approval Date Change
  public _onApprovalDatePickerChange = (date?: Date): void => {
    this.setState({
      approvalDate: date,
      approvalDateEdit: date
    });
  }
  //PublishOption Change
  public _publishOptionChange(option: { key: any; text: any }) {
    this.setState({ publishOption: option.key, saveDisable: false });
  }
  //Critical Document Change
  public _onCriticalChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ criticalDocument: true }); }
    else if (!isChecked) { this.setState({ criticalDocument: false }); }
  }
  // Template Change
  public _onTemplateChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ templateDocument: true }); }
    else if (!isChecked) { this.setState({ templateDocument: false }); }
  }
  //Revision Settings Change
  public _revisionCodingChange(option: { key: any; text: any }) {
    this.setState({ revisionCodingId: option.key, revisionCoding: option.text });
  }
  // on subcontractor number change
  public _subContractorNumberChange = (ev: React.FormEvent<HTMLInputElement>, subContractorNumber?: string) => {
    this.setState({ subContractorNumber: subContractorNumber || '' });
  }
  public _CustomerNumberChange = (ev: React.FormEvent<HTMLInputElement>, customerNumber?: string) => {
    this.setState({ customerNumber: customerNumber || '' });
  }
  // Transmittal Change
  public _onTransmittalChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ transmittalCheck: true }); }
    else if (!isChecked) { this.setState({ transmittalCheck: false }); }
  }
  // External Document Change
  public _onExternalDocumentChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ externalDocument: true }); }
    else if (!isChecked) { this.setState({ externalDocument: false }); }
  }
  // On update click
  public async _onUpdateClick() {
    if (this.state.createDocument == true && this.isDocument == "Yes" || this.state.createDocument == false) {
      if (this.state.expiryCheck == true) {
        if (this.props.project) {
          //Validation without direct publish
          if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
            && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')
            && this.validator.fieldValid('DocumentController') && this.validator.fieldValid('Revision')) {
            this.setState({ updateDisable: true });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish')
            && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
            && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')
            && this.validator.fieldValid('DocumentController') && this.validator.fieldValid('Revision')) {
            this.setState({ updateDisable: true, hideloader: false });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
        else {
          //Validation without direct publish
          if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
            this.setState({ updateDisable: true });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')) {
            this.setState({ updateDisable: true, hideloader: false });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
      }
      else {
        if (this.props.project) {
          //Validation without direct publish
          if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
            && this.validator.fieldValid('DocumentController') && this.validator.fieldValid('Revision')) {
            this.setState({ updateDisable: true });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish')
            && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
            && this.validator.fieldValid('DocumentController') && this.validator.fieldValid('Revision')) {
            this.setState({ updateDisable: true, hideloader: false });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
        else {
          //Validation without direct publish
          if (this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
            this.setState({ updateDisable: true });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if ((this.state.directPublishCheck == true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')) {
            this.setState({ updateDisable: true, hideloader: false });
            await this._updateDocument();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
      }
    }
    else {
      this.setState({ insertdocument: "" });
    }
  }
  // On update document
  public async _updateDocument() {
    this.setState({
      hideCreateLoading: " ",
      norefresh: " "
    });
    this._userMessageSettings();
    let documentNameExtension: any;
    let sourceDocumentId: any;
    let upload: any;
    let documentIdname: any;
    if (this.props.project) {
      upload = "#editproject";
    }
    else {
      upload = "#editqdms";
    }
    // With document
    if (this.state.createDocument == true) {
      await this._updateDocumentIndex();
      // Get file from form
      console.log((document.querySelector(upload) as HTMLInputElement).files.length)
      if ((document.querySelector(upload) as HTMLInputElement).files.length != 0) {
        let myfile = (document.querySelector(upload) as HTMLInputElement).files[0];
        console.log(myfile);
        var splitted = myfile.name.split(".");
        if (this.state.replaceDocumentCheckbox == true) {
          if (this.state.titleReadonly == true) {
            documentNameExtension = this.state.documentName;
          }
          else {
            documentNameExtension = this.state.documentid + " " + this.state.title + '.' + splitted[splitted.length - 1];
          }
        }
        else {
          if (this.state.titleReadonly == true) {
            documentNameExtension = this.state.documentName + '.' + splitted[splitted.length - 1];
          }
          else {
            documentNameExtension = this.state.documentid + " " + this.state.title + '.' + splitted[splitted.length - 1];
          }
        }
        documentIdname = this.state.documentid + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        // alert(this.documentNameExtension);
        if (myfile.size) {
          // add file to source library
          const fileUploaded = await this._Service.uploadDocument(this.props.sourceDocumentLibrary, documentIdname, myfile);
          if (fileUploaded) {
            console.log("File Uploaded");
            const item = await fileUploaded.file.getItem();
            console.log(item);
            sourceDocumentId = item["ID"];
            // if(splitted[1] == "pdf"||splitted[1] == "Pdf"||splitted[1] == "PDF"){
            //   documenturl = item["ServerRedirectedEmbedUrl"];
            // }
            // else{
            // docServerUrl = item["ServerRedirectedEmbedUrl"];
            // splitdocUrl = docServerUrl.split("&", 2);
            // documenturl = splitdocUrl[0];
            // }
            this.sourceDocumentID = sourceDocumentId;
            this.setState({ sourceDocumentId: sourceDocumentId });
            // update metadata
            await this._updateSourceDocument();
            if (item) {
              let revision;
              if (this.props.project) {
                revision = "-";
              }
              else {
                revision = "0";
              }
              this._updatePublishDocument();
              if (this.state.replaceDocumentCheckbox == true) {
                const indexdata1 = {
                  SourceDocument: {
                    Description: this.documentNameExtension,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  DocumentName: this.documentNameExtension,
                  WorkflowStatus: "Draft"
                }
                this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata1, parseInt(this.documentIndexID))
                this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                setTimeout(() => {
                  window.location.replace(this.props.siteUrl);
                }, 5000);
              }
              else {
                const logdata1 = {
                  Title: this.state.documentid,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: revision,
                  DocumentIndexId: this.documentIndexID
                }
                await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata1)

                // update document index
                if (this.state.directPublishCheck == false) {
                  const indexdata = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    },
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.documentIndexID))

                }
                else {
                  const indexdata2 = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    ApprovedDate: this.state.approvalDate,
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata2, parseInt(this.documentIndexID))
                }
                await this._triggerPermission(sourceDocumentId);
                if (this.state.directPublishCheck == true) {
                  this.setState({ hideLoading: false, hideCreateLoading: "none" });
                  await this._publish();
                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.props.siteUrl);
                  }, 5000);
                }
              }
            }
          }
        }
      }
      else if (this.state.templateId != "") {
        let publishName: any;
        let extension: any;
        let newDocumentName: any;

        // Get template
        if (this.state.sourceId == "QDMS") {
          let publishdoc = await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
          console.log(publishdoc);
          for (let i = 0; i < publishdoc.length; i++) {
            if (publishdoc[i].Id == this.state.templateId) {
              publishName = publishdoc[i].LinkFilename;
            }
          }
          var split = publishName.split(".", 2);
          extension = split[1];
          if (publishdoc) {
            // Add template document to source document
            newDocumentName = this.state.documentName + "." + extension;
            this.documentNameExtension = newDocumentName;
            documentIdname = this.state.documentid + '.' + extension;
            let qdmsURl = this.props.QDMSUrl + "/" + this.props.publisheddocumentLibrary + "/" + publishName;
            await this._Service.getqdmsdocument(qdmsURl)
              .then((templateData: any) => {
                return this._Service.uploadDocument(this.props.sourceDocumentLibrary, documentIdname, templateData)
              }).then(async (fileUploaded: any) => {
                console.log("File Uploaded");
                fileUploaded.file.getItem()
                  .then(async (item: any) => {
                    console.log(item);
                    sourceDocumentId = item["ID"];
                    this.sourceDocumentID = sourceDocumentId;
                    this.setState({ sourceDocumentId: sourceDocumentId });
                    await this._updateSourceDocument();
                  });
                let revision;
                if (this.props.project) {
                  revision = "-";
                }
                else {
                  revision = "0";
                }
                this._updatePublishDocument();
                const logdata2 = {
                  Title: this.state.documentid,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: revision,
                  DocumentIndexId: this.documentIndexID
                }
                this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata2)
                // update document index
                if (this.state.directPublishCheck == false) {
                  const indexdata3 = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata3, parseInt(this.documentIndexID))
                }
                else {
                  const indexdata4 = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    ApprovedDate: this.state.approvalDate,
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    },
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata4, parseInt(this.documentIndexID))
                }
                await this._triggerPermission(sourceDocumentId);
                if (this.state.directPublishCheck == true) {
                  this.setState({ hideLoading: false, hideCreateLoading: "none" });
                  await this._publish();
                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.props.siteUrl);
                  }, 5000);
                }

              });
          }
        }
        else {
          let publishdoc = await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary)
          console.log(publishdoc);
          for (let i = 0; i < publishdoc.length; i++) {
            if (publishdoc[i].Id == this.state.templateId) {
              publishName = publishdoc[i].LinkFilename;
            }
          }
          var split = publishName.split(".", 2);
          extension = split[1];
          if (publishdoc) {
            // Add template document to source document
            newDocumentName = this.state.documentName + "." + extension;
            this.documentNameExtension = newDocumentName;
            documentIdname = this.state.documentid + '.' + extension;
            let siteUrl = this.props.siteUrl + "/" + this.props.publisheddocumentLibrary + "/" + publishName;
            this._Service.getDocument(siteUrl)
              .then((templateData: any) => {
                return this._Service.uploadDocument(this.props.sourceDocumentLibrary, documentIdname, templateData)
              }).then(async (fileUploaded: any) => {
                console.log("File Uploaded");
                fileUploaded.file.getItem().then(async (item: any) => {
                  console.log(item);
                  sourceDocumentId = item["ID"];
                  this.sourceDocumentID = sourceDocumentId;
                  this.setState({ sourceDocumentId: sourceDocumentId });
                  await this._updateSourceDocument();
                });
                let revision;
                if (this.props.project) {
                  revision = "-";
                }
                else {
                  revision = "0";
                }
                this._updatePublishDocument();
                const logdata3 = {
                  Title: this.state.documentid,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: revision,
                  DocumentIndexId: this.documentIndexID
                }
                this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata3)
                // update document index
                if (this.state.directPublishCheck == false) {
                  const indexdata5 = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata5, parseInt(this.documentIndexID))

                }
                else {
                  const indexdata6 = {
                    SourceDocumentID: parseInt(this.state.sourceDocumentId),
                    ApprovedDate: this.state.approvalDate,
                    DocumentName: this.documentNameExtension,
                    SourceDocument: {
                      Description: this.documentNameExtension,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    RevokeExpiry: {
                      Description: "Revoke",
                      Url: this.revokeUrl
                    }
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata6, parseInt(this.documentIndexID))

                }
                await this._triggerPermission(sourceDocumentId);
                if (this.state.directPublishCheck == true) {
                  this.setState({ hideLoading: false, hideCreateLoading: "none" });
                  await this._publish();
                }
                else {
                  this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
                  setTimeout(() => {
                    window.location.replace(this.props.siteUrl);
                  }, 5000);
                }

              });
          }
        }
      }
      else {
        this._updateWithoutDocument();
      }
    }
    else {
      await this._updateDocumentIndex();
      this.setState({ hideCreateLoading: "none", norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 } });
      setTimeout(() => {
        window.location.replace(this.props.siteUrl);
      }, 5000);
    }
  }
  // Update Document Index
  public _updateDocumentIndex() {
    // Without Expiry date
    if ((this.state.expiryCheck == false) || (this.state.expiryCheck == "")) {
      //DMS
      if (this.props.project) {
        const indexdata8 = {
          Title: this.state.title,
          DocumentName: this.state.documentid + " " + this.state.title,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          OwnerId: this.state.owner,
          ApproverId: this.state.approver,
          CreateDocument: this.state.createDocument,
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          PublishFormat: this.state.publishOption,
          DirectPublish: this.state.directPublishCheck,
          ApprovedDate: this.state.approvalDateEdit,
          ReviewersId: this.state.reviewers,
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata8, parseInt(this.documentIndexID))
          .then(afteradd => {
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.documentIndexID + "&mode=expiry";
          });
      }
      //QDMS
      else {
        const indexdata9 = {
          Title: this.state.title,
          DocumentName: this.state.documentid + " " + this.state.title,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          OwnerId: this.state.owner,
          ApproverId: this.state.approver,
          CreateDocument: this.state.createDocument,
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          PublishFormat: this.state.publishOption,
          DirectPublish: this.state.directPublishCheck,
          ApprovedDate: this.state.approvalDateEdit,
          ReviewersId: this.state.reviewers,
        }
        this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata9, parseInt(this.documentIndexID))
          .then(afteradd => {
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.documentIndexID + "&mode=expiry";
          });
      }
    }
    else {
      // DMS
      if (this.props.project) {
        const indexdata7 = {
          Title: this.state.title,
          DocumentName: this.state.documentid + " " + this.state.title,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          OwnerId: this.state.owner,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          ExpiryDate: this.state.expiryDate,
          ApproverId: this.state.approver,
          CreateDocument: this.state.createDocument,
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          PublishFormat: this.state.publishOption,
          DirectPublish: this.state.directPublishCheck,
          ApprovedDate: this.state.approvalDateEdit,
          ReviewersId: this.state.reviewers,
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata7, parseInt(this.documentIndexID)).then(afteradd => {
          this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID + "";
          this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.documentIndexID + "&mode=expiry";
        });
      }
      //QDMS
      else {
        const indexdata10 = {
          Title: this.state.title,
          DocumentName: this.state.documentid + " " + this.state.title,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          OwnerId: this.state.owner,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          ExpiryDate: this.state.expiryDate,
          ApproverId: this.state.approver,
          CreateDocument: this.state.createDocument,
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          PublishFormat: this.state.publishOption,
          DirectPublish: this.state.directPublishCheck,
          ApprovedDate: this.state.approvalDateEdit,
          ReviewersId: this.state.reviewers,
        }
        this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata10, parseInt(this.documentIndexID))
          .then(afteradd => {
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.documentIndexID + "&mode=expiry";
          });
      }
    }
  }
  // Update Source Document
  public _updateSourceDocument() {
    // Without Expiry date
    if (this.state.expiryCheck == false) {
      //DMS
      if (this.props.project) {
        const libdata1 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.documentNameExtension,
          BusinessUnit: this.state.businessUnit,
          Category: this.state.category,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          Revision: "-",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          DocumentIndexId: this.documentIndexID,
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          },
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          CriticalDocument: this.state.criticalDocument,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, libdata1, this.sourceDocumentID)

      }
      //QDMS
      else {
        const libdata2 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.documentNameExtension,
          BusinessUnit: this.state.businessUnit,
          Category: this.state.category,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          Revision: "0",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          DocumentIndexId: this.documentIndexID,
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          CriticalDocument: this.state.criticalDocument,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          }
        }
        this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, libdata2, this.sourceDocumentID)

      }
    }
    // With Expiry Date
    else {
      // DMS
      if (this.props.project) {
        const libdata3 = {
          DocumentID: this.state.documentid,
          Title: this.state.title,
          ReviewersId: this.state.reviewers,
          DocumentName: this.documentNameExtension,
          BusinessUnit: this.state.businessUnit,
          Category: this.state.category,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          Revision: "-",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          DocumentIndexId: this.documentIndexID,
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          },
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          CriticalDocument: this.state.criticalDocument,
          ExpiryDate: this.state.expiryDate,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, libdata3, this.sourceDocumentID)

      }
      // QDMS
      else {
        const libdata4 = {
          DocumentID: this.state.documentid,
          Title: this.state.title,
          ReviewersId: this.state.reviewers,
          DocumentName: this.documentNameExtension,
          BusinessUnit: this.state.businessUnit,
          Category: this.state.category,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          ExpiryDate: this.state.expiryDate,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          Revision: "0",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          DocumentIndexId: this.documentIndexID,
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          CriticalDocument: this.state.criticalDocument,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          }
        }
        this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, libdata4, this.sourceDocumentID);

      }
    }
  }
  // Update Publish Document
  public async _updatePublishDocument() {
    const publishDocumentID: any[] = await this._Service.getpublishlibrary(this.props.siteUrl, this.props.publisheddocumentLibrary, this.documentIndexID);
    console.log("publishDocumentID", publishDocumentID);
    if (publishDocumentID.length > 0) {
      // Without Expiry date
      if (this.state.expiryCheck == false) {
        //DMS
        if (this.props.project) {
          for (var k in publishDocumentID) {
            const publishdata1 = {
              Title: this.state.title,
              DocumentName: this.documentNameExtension,
              SubCategory: this.state.subCategory,
              OwnerId: this.state.owner,
              ApproverId: this.state.approver,
              Template: this.state.templateDocument,
              PublishFormat: this.state.publishOption,
              ReviewersId: this.state.reviewers,
              TransmittalDocument: this.state.transmittalCheck,
              ExternalDocument: this.state.externalDocument,
              RevisionCodingId: this.state.revisionCodingId,
              RevisionLevel: this.state.revisionLevelId,
              DocumentControllerId: this.state.dcc,
            }
            this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata1, publishDocumentID[k].ID)

          }
        }
        else {
          for (var s in publishDocumentID) {
            const publishdata2 = {
              Title: this.state.title,
              DocumentName: this.documentNameExtension,
              SubCategory: this.state.subCategory,
              OwnerId: this.state.owner,
              ApproverId: this.state.approver,
              Template: this.state.templateDocument,
              PublishFormat: this.state.publishOption,
              ReviewersId: this.state.reviewers,
            }
            this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata2, publishDocumentID[s].ID)

          }
        }
      }
      // With Expiry date
      else {
        if (this.props.project) {
          for (var g in publishDocumentID) {
            const publishdata3 = {
              Title: this.state.title,
              DocumentName: this.documentNameExtension,
              SubCategory: this.state.subCategory,
              OwnerId: this.state.owner,
              ExpiryLeadPeriod: this.state.expiryLeadPeriod,
              ExpiryDate: this.state.expiryDate,
              ApproverId: this.state.approver,
              Template: this.state.templateDocument,
              PublishFormat: this.state.publishOption,
              ReviewersId: this.state.reviewers,
              TransmittalDocument: this.state.transmittalCheck,
              ExternalDocument: this.state.externalDocument,
              RevisionCodingId: this.state.revisionCodingId,
              RevisionLevel: this.state.revisionLevelId,
              DocumentControllerId: this.state.dcc,
            }
            this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata3, publishDocumentID[g].ID)

          }
        }
        else {
          for (var j in publishDocumentID) {
            const publishdata4 = {
              Title: this.state.title,
              DocumentName: this.documentNameExtension,
              SubCategory: this.state.subCategory,
              OwnerId: this.state.owner,
              ExpiryLeadPeriod: this.state.expiryLeadPeriod,
              ExpiryDate: this.state.expiryDate,
              ApproverId: this.state.approver,
              Template: this.state.templateDocument,
              PublishFormat: this.state.publishOption,
              ReviewersId: this.state.reviewers,
            }
            this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata4, publishDocumentID[j].ID);

          }
        }
      }
    }
  }
  // Update without document
  public async _updateWithoutDocument() {
    let sourceUrl: any;
    let extensionSplit: any;
    const sourceLink: any = await this._Service.getListSourceItem(this.props.siteUrl, this.props.documentIndexList, parseInt(this.documentIndexID))
    console.log(sourceLink);
    sourceUrl = sourceLink.SourceDocument.Url;
    var split = sourceLink.SourceDocument.Description.split(".", 2);
    extensionSplit = split[1];
    if (sourceLink) {
      if (this.props.project) {
        if (this.state.directPublishCheck == false) {
          this.setState({
            approvalDate: null,
          });
        }

        //  if(diupdate){
        const sourcedata = {
          Title: this.state.title,
          DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
          OwnerId: this.state.owner,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          ExpiryDate: this.state.expiryDate,
          ApproverId: this.state.approver,
          Template: this.state.templateDocument,
          PublishFormat: this.state.publishOption,
          ReviewersId: this.state.reviewers,
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata, this.sourceDocumentID)
          .then(async results => {
            const publishDocumentID: any[] = await this._Service.getpublishlibrary(this.props.siteUrl, this.props.publisheddocumentLibrary, this.documentIndexID);
            console.log("publishDocumentID", publishDocumentID);
            for (var k in publishDocumentID) {
              const publishdata5 = {
                Title: this.state.title,
                DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
                OwnerId: this.state.owner,
                ExpiryLeadPeriod: this.state.expiryLeadPeriod,
                ExpiryDate: this.state.expiryDate,
                ApproverId: this.state.approver,
                Template: this.state.templateDocument,
                PublishFormat: this.state.publishOption,
                ReviewersId: this.state.reviewers,
                TransmittalDocument: this.state.transmittalCheck,
                ExternalDocument: this.state.externalDocument,
                RevisionCodingId: this.state.revisionCodingId,
                RevisionLevel: this.state.revisionLevelId,
                DocumentControllerId: this.state.dcc,
                SubcontractorDocumentNo: this.state.subContractorNumber,
                CustomerDocumentNo: this.state.customerNumber
              }
              this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata5, publishDocumentID[k].ID);

            }
          });
        const indexdata11 = {
          DocumentName: this.state.documentid + " " + this.state.title + "." + extensionSplit,
          SourceDocument: {
            Description: this.state.documentid + " " + this.state.title + "." + extensionSplit,
            Url: sourceUrl
          },
        }
        const diupdate = await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata11, parseInt(this.documentIndexID))
        // }
        this.setState({
          statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 },
          messageBar: "",
          hideCreateLoading: "none", norefresh: "none"
        });
        setTimeout(() => {
          window.location.replace(this.props.siteUrl);
        }, 5000);

      }
      else {
        if (this.state.directPublishCheck == false) {
          this.setState({
            approvalDate: null,
          });
        }

        // if(afterIndexUpdate) {
        const sourcedata2 = {
          Title: this.state.title,
          DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
          OwnerId: this.state.owner,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          ExpiryDate: this.state.expiryDate,
          ApproverId: this.state.approver,
          Template: this.state.templateDocument,
          PublishFormat: this.state.publishOption,
          ReviewersId: this.state.reviewers,

        }
        const results = await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata2, this.sourceDocumentID)
        if (results) {
          const publishDocumentID: any[] = await this._Service.getpublishlibrary(this.props.siteUrl, this.props.publisheddocumentLibrary, this.documentIndexID);
          console.log("publishDocumentID", publishDocumentID);
          for (var k in publishDocumentID) {
            const publishdata6 = {
              Title: this.state.title,
              DocumentName: this.state.documentid + this.state.title + "." + extensionSplit,
              OwnerId: this.state.owner,
              ExpiryLeadPeriod: this.state.expiryLeadPeriod,
              ExpiryDate: this.state.expiryDate,
              ApproverId: this.state.approver,
              Template: this.state.templateDocument,
              PublishFormat: this.state.publishOption,
              ReviewersId: this.state.reviewers
            }
            await this._Service.updateLibraryItem(this.props.siteUrl, this.props.publisheddocumentLibrary, publishdata6, publishDocumentID[k].ID)

          }
          const indexdata12 = {
            DocumentName: this.state.documentid + " " + this.state.title + "." + extensionSplit,
            SourceDocument: {
              Description: this.state.documentid + " " + this.state.title + "." + extensionSplit,
              Url: sourceUrl
            },
          }
          const afterIndexUpdate = await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata12, parseInt(this.documentIndexID))

        }
        // }

        this.setState({
          statusMessage: { isShowMessage: true, message: this.editDocument, messageType: 4 },
          messageBar: "",
          hideCreateLoading: "none", norefresh: "none"
        });
        setTimeout(() => {
          window.location.replace(this.props.siteUrl);
        }, 5000);

      }
    }
  }
  // Document permission
  protected async _triggerPermission(sourceDocumentID: any) {
    const laUrl = await this._Service.gettriggerPermission(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  //Document Published
  protected async _publish() {
    if (this.props.project) {
      await this._generateNewRevision();
    }
    else {
      await this._revisionCoding();
    }
    const laUrl = await this._Service.getpublish(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'Status': 'Published',
      'PublishFormat': this.state.publishOption,
      'SourceDocumentID': this.state.sourceDocumentId,
      'SiteURL': siteUrl,
      'PublishedDate': this.today,
      'DocumentName': this.state.documentName,
      'Revision': this.state.newRevision,
      'SourceDocumentLibrary': this.props.sourceDocumentViewLibrary,
      'WorkflowStatus': "Published",
      'RevisionUrl': this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.state.newDocumentId,
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      this._publishUpdate(responseJSON.PublishDocID);
    }
    else { }
  }
  // To generate new revision
  public _generateNewRevision = async () => {
    let currentRevision = '-'; // set the current revisionsettings ID in state variable.
    this.setState({
      previousRevisionItemID: this.state.revisionCodingId // set this value with previous revision settings id from Project document index item.
    });
    // Reading current revision coding details from RevisionSettings.
    const revisionItem: any = await this._Service.getRevisionListItems(this.props.siteUrl, this.props.revisionSettingsList, parseInt(this.state.revisionCodingId));
    let startPrefix = '-';
    let newRevision = '';
    let pattern = revisionItem.Pattern;
    let endWith = '0';
    let minN = revisionItem.MinN;
    let maxN = '0';
    let isAutoIncrement = revisionItem.AutoIncrement == 'TRUE';
    let firstChar = currentRevision.substring(0, 1);
    let currentNumber = currentRevision.substring(1, currentRevision.length);
    let startWith = revisionItem.StartWith;

    if (revisionItem.EndWith != null)
      endWith = revisionItem.EndWith;

    if (revisionItem.MaxN != null)
      maxN = revisionItem.MaxN;

    if (revisionItem.StartPrefix != null)
      startPrefix = revisionItem.StartPrefix.toString();

    //splitting pattern values
    let incrementValue = 1;
    let isAlphaIncrement = pattern.split('+')[0] == 'A';
    let isNumericIncrement = pattern.split('+')[0] == 'N';
    if (pattern.split('+').length == 2) {
      incrementValue = Number(pattern.split('+')[1]);
    }
    //Resetting current revision as blank if current revisionsetting id is different.
    if (this.state.revisionItemID != this.state.previousRevisionItemID) {
      currentRevision = '-';
    }
    try {
      //Getting first revision value.
      if (currentRevision == '-') {
        if (!isAutoIncrement) // Not an auto increment pattern, splitting the pattern with command reading the first value.
        {
          newRevision = pattern.split(',')[0];
        }
        else {
          if (startPrefix != '-' && startPrefix.split(',').length > 0)  //Auto increment   with startPrefix eg. A1,A2, A3 etc., then handling both single and multple startPrefix
          {
            startPrefix = startPrefix.split(',')[0];
          }
          if (startWith != null) // 
          {
            newRevision = startWith; //assigning startWith as newRevision for the first time.
          }
          else {
            newRevision = startPrefix + '' + minN;
          }
          if (startWith == null && startPrefix == '-') // Assigning minN if startWith and StartPrefix are null.
          {
            newRevision = minN;
          }
        }
      }
      else if (!isAutoIncrement) // currentRevision is not blank, so splitting pattern string for non- auto - increment pattern.
      {
        let patternArray = pattern.split(',');
        newRevision = patternArray[0]; // if array value exceeds , resetting revision.
        /* let prevRevision = patternArray[0];
         for(let i= 0;i < patternArray.length; i++)
         {
           if(i > 0 && String(currentRevision) == String(patternArray[i]))
           {
             prevRevision = String(patternArray[i-1]);
             break;
           }
         }
         console.log('prevRevision:' + prevRevision);*/
        console.log('currentRevision:' + currentRevision);
        for (let i = 0; i < patternArray.length; i++) {
          {
            //B,C,D,C,E,G
            if (String(currentRevision) == patternArray[i] && (i + 1) < patternArray.length) {
              newRevision = patternArray[i + 1];
              break;
            }
          }
        }
      }
      else if (isAutoIncrement)// current revision is not blank and auto increment pattern .
      {
        if (startWith != null && String(currentRevision) == String(startWith)) // Revision code with startWith  and startWith already set as Revision
        {
          if (startPrefix == '-') // second revision without startPrefix / '-' no StartPrefix
          {
            newRevision = minN;
          }
          else // 
          {
            newRevision = startPrefix + minN;
          }
        }
        // For all other cases
        else if (startPrefix != '-') // Handling revisions with startPrefix here first char will be alpha
        {
          if (startPrefix.split(',').length == 1) // Single startPrefix eg. A1,A2,A3 etc with startPrefix 'A' and patter N+1
          {
            if (this.isNotANumber(minN)) // Alpha increment.
            {
              newRevision = startPrefix + this.nextChar(firstChar, incrementValue);
            }
            else  // number increment.
            {
              newRevision = startPrefix + (Number(currentNumber) + Number(incrementValue)).toString();
            }
          }
          else // startPrefix with multiple values
          {
            if (maxN != '0') {
              if (this.isNotANumber(currentRevision)) //MaxN set and not a number.
              {
                if (Number(currentNumber) < Number(maxN)) // alpha type revision
                {
                  newRevision = firstChar + (Number(currentNumber) + Number(incrementValue)).toString();
                }
                else if (Number(currentNumber) == Number(maxN)) {
                  // if current number part is same as maxN, get the next StartPrefix value from startPrefix.split(',')
                  let startPrefixArray = startPrefix.split(',');
                  for (let i = 0; i < startPrefixArray.length; i++) {
                    if (firstChar == startPrefixArray[i] && (i + 1) < startPrefixArray.length) {
                      firstChar = startPrefixArray[i + 1];
                      break;
                    }
                  }
                  if (firstChar == " ") // " " will denote a number
                  {
                    newRevision = minN;
                  }
                  else {
                    newRevision = firstChar + minN;
                  }
                }
              }
              else  // current revion number itself is a number and with multiple StartPrefix
              {
                if (Number(currentRevision) < Number(maxN)) {
                  newRevision = (Number(currentRevision) + Number(incrementValue)).toString(); // current revision s not an alpha 
                }
                else if (Number(currentRevision) == Number(maxN)) {
                  {
                    if (!this.isNotANumber(currentRevision)) // for setting a default value after the last item
                    {
                      firstChar = " ";
                    }
                    // if current number part is same as maxN, get the next StartPrefix value from startPrefix.split(',')
                    let startPrefixArray = startPrefix.split(',');
                    for (let i = 0; i < startPrefixArray.length; i++) {
                      if (firstChar == startPrefixArray[i] && (i + 1) < startPrefixArray.length) {
                        firstChar = startPrefixArray[i + 1];
                        break;
                      }
                    }
                    if (firstChar == " ") // Assigning number for blank array.
                    {
                      newRevision = minN;
                    }
                    else {
                      newRevision = firstChar + minN;
                    }
                  }
                }
              }
            }
          }
        }
        if (newRevision == '' && startPrefix == '-' && endWith == '0') // No StartPrefix and No EndWith
        {
          if (isAlphaIncrement) // Alpha increment.
          {
            newRevision = this.nextChar(firstChar, incrementValue);
          }
          else {
            newRevision = (Number(currentRevision) + Number(incrementValue)).toString();
          }
        }
        else if (startPrefix == '-' && endWith != '0') // No StartPrefix and with EndWith 
        {
          // cases A to E  then 0,1, 2,3 etc,
          if (currentRevision == endWith) {
            newRevision = minN;
          }
          else// if(currentRevision != '0')
          {
            if (this.isNotANumber(currentRevision)) // Alpha increment.
            {
              newRevision = this.nextChar(firstChar, incrementValue);
            }
            else // (currentRevision == startWith && endWith != null) // always alpha increment "X,,B"
            {
              newRevision = (Number(currentRevision) + Number(incrementValue)).toString();
            }
          }
        }
      }
      if (newRevision.indexOf('undefined') > -1 || newRevision == '') // Assigning with zero if array value exceeds.
      {
        newRevision = '0';
      }
    }
    catch {
      newRevision = '-1'; // check with -1 for error value
    }
    this.setState({
      newRevision: newRevision,
      currentRevision: newRevision
    });
    console.log('new revision :' + newRevision);
  }
  // Creating next alpha char.
  private nextChar(currentChar: any, increment: any) {
    if (currentChar == 'Z')
      return 'A';
    else
      return String.fromCharCode(currentChar.charCodeAt(0) + increment);
  }
  // Check for number and alpha
  private isNotANumber(checkChar: any) {
    return isNaN(checkChar);
  }
  // QDMS revision coding
  public _revisionCoding = async () => {
    let revision = parseInt("0");
    let rev = revision + 1;
    this.setState({ newRevision: rev.toString() });

  }
  // Published Document Metadata update
  public async _publishUpdate(publishid: any) {
    const indexdata13 = {
      PublishFormat: this.state.publishOption,
      WorkflowStatus: "Published",
      Revision: this.state.newRevision,
      ApprovedDate: new Date()
    }
    await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata13, parseInt(this.documentIndexID))
    if (this.state.owner != this.currentId) {
      this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
    }
    const logdata4 = {
      Title: this.state.documentid,
      Status: "Published",
      LogDate: this.today,
      Revision: this.state.newRevision,
      DocumentIndexId: this.documentIndexID
    }
    await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata4)
    this.setState({ hideLoading: true, norefresh: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.directPublish, messageType: 4 } });
    setTimeout(() => {
      window.location.replace(this.props.siteUrl);
    }, 5000);
  }
  //Send Mail
  public _sendMail = async (emailuser: any, type: any, name: any) => {
    let formatday = moment(this.today).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    console.log(this.state.criticalDocument);
    const notificationPreference: any[] = await this._Service.getnotification(this.props.hubUrl, this.props.notificationPreference, emailuser);
    console.log(notificationPreference[0].Preference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        mailSend = "Yes";
      }
      else {
        mailSend = "No";
      }
    }
    else if (this.state.criticalDocument == true) {
      mailSend = "Yes";
    }
    if (mailSend == "Yes") {
      const emailNotification: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.emailNotification);
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title == type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }
      }
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
      let replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject,
          "body": {
            "contentType": "Text",
            "content": replaceBody
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": emailuser
              }
            }
          ],
        }
      };
      //Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient("3")
        .then((client: MSGraphClientV3): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody);
        });
    }
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',
  };
  private modalProps = {
    isBlocking: true,
  };
  // Back button of version history & revision history in edit form
  public _back = () => {
    window.location.replace(this.props.siteUrl + "/SitePages/EditDocument.aspx?did=" + this.documentIndexID);
  }

  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
    window.location.replace(this.props.siteUrl);
  }
  //Not Cancel
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  //Cancel Document
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });
  }
  // Format date
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  }
  public render(): React.ReactElement<ITransmittalEditDocumentProps> {
    const publishOptions: IDropdownOption[] = [
      { key: 'PDF', text: 'PDF' },
      { key: 'Native', text: 'Native' },
    ];
    const publishOption: IDropdownOption[] = [
      { key: 'Native', text: 'Native' },
    ];
    const Source: IDropdownOption[] = [
      { key: 'QDMS', text: 'Quality' },
      { key: 'Current Site', text: 'Current Site' }
    ];
    const back: IIconProps = { iconName: 'ChromeBack' };
    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const Go: IIconProps = { iconName: "CaretRightSolid8" };
    const uploadOrTemplateRadioBtnOptions:
      IChoiceGroupOption[] = [
        { key: 'Upload', text: 'Upload existing file' },
        { key: 'Template', text: 'Create document existing template', styles: { field: { marginLeft: '18em' } } },
      ];
    const choiceGroupStyles: Partial<IChoiceGroupStyles> = { root: { display: 'flex' }, flexContainer: { display: "flex", justifyContent: 'space-between' } };

    return (
      <section className={`${styles.transmittalEditDocument}`}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        {/* Edit Document QDMS */}
        <div style={{ display: this.state.qdmsEditDocumentView }} >
          <div className={styles.border}>
            <div className={styles.alignCenter}>{this.props.webpartHeader}</div>
            <Pivot aria-label="Links of Tab Style Pivot Example" >
              <PivotItem headerText="Document Info" >
                <div>
                  <Label>Document ID: {this.state.documentid}</Label>
                </div>
                <div>
                  <TextField required id="t1"
                    label="Title"
                    onChange={this._titleChange}
                    value={this.state.title} readOnly={this.state.titleReadonly} ></TextField>
                  <div style={{ color: "#dc3545", fontWeight: "bold", display: this.state.checkrename }}>Checking your permission to rename.Please wait...</div>
                  <div style={{ color: "#dc3545" }}>
                    {this.validator.message("Name", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthrgt}>
                    <TextField
                      label="Department"
                      value={this.state.department} readOnly></TextField>
                  </div>
                  <div className={styles.wdthlft}>
                    <TextField
                      label="Business Unit"
                      value={this.state.businessUnit} readOnly></TextField>
                  </div>
                </div>

                <div className={styles.divrow}>
                  <div className={styles.wdthrgt}>
                    <TextField
                      label="Category"
                      value={this.state.category} readOnly></TextField>
                  </div>
                  <div className={styles.wdthlft}>
                    <TextField
                      label="Sub Category"
                      value={this.state.subCategory} readOnly></TextField>
                    {/* <Dropdown id="t2" label="Sub Category"
                    placeholder="Select an option"
                    selectedKey={this.state.subCategoryId}
                    options={this.state.subCategoryArray}
                    onChanged={this._subCategoryChange} />  */}
                  </div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthrgt}>
                    <TextField
                      label="Legal Entity"
                      value={this.state.legalEntity} readOnly></TextField>
                  </div>
                  <div className={styles.wdthlft}>
                    <PeoplePicker
                      context={this.props.context as any}
                      titleText="Owner"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users    
                      showtooltip={true}
                      required={false}
                      disabled={false}
                      ensureUser={true}
                      onChange={(items) => this._selectedOwner(items)}
                      defaultSelectedUsers={[this.state.ownerName]}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Owner", this.state.owner, "required")}{" "}
                    </div>
                  </div>
                </div>
                <div >
                  <PeoplePicker
                    context={this.props.context as any}
                    titleText="Reviewer(s)"
                    personSelectionLimit={20}
                    groupName={""} // Leave this blank in case you want to filter from all users
                    showtooltip={true}
                    required={true}
                    disabled={false}
                    ensureUser={true}
                    showHiddenInUI={false}
                    onChange={(items) => this._selectedReviewers(items)}
                    defaultSelectedUsers={this.state.reviewersName}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
                <div>
                  <PeoplePicker
                    context={this.props.context as any}
                    titleText="Approver"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    ensureUser={true}
                    onChange={(items) => this._selectedApprover(items)}
                    showHiddenInUI={false}
                    defaultSelectedUsers={[this.state.approverName]}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
                <div style={{ display: this.state.validApprover, color: "#dc3545" }}>Not able to change approver</div>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Approver", this.state.approver, "required")}{" "}
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthrgt} >
                    <DatePicker label="Expiry Date"
                      value={this.state.expiryDate}
                      onSelectDate={this._onExpDatePickerChange}
                      placeholder="Select a date..."
                      ariaLabel="Select a date" minDate={new Date()}
                      formatDate={this._onFormatDate} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("expiryDate", this.state.expiryDate, "required")}{""}</div>
                  </div>
                  {/* </div> */}
                  {/* <div style={{ display: this.state.hideExpiry }}> */}
                  <div className={styles.wdthlft} >
                    <TextField id="ExpiryLeadPeriod" name="ExpiryLeadPeriod"
                      label="Expiry Lead Period(in days)" onChange={this._expLeadPeriodChange}
                      value={this.state.expiryLeadPeriod}>
                    </TextField>
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required")}{""}</div>
                    <div style={{ color: "#dc3545", display: this.state.leadmsg }}>
                      Enter only numbers less than 101
                    </div></div>
                  {/* </div> */}
                </div>
                <div style={{ display: this.state.hideCreate }}>
                  <div className={styles.divrow}>
                    <div style={{ display: this.state.replaceDocument }}>
                      <Label >Document :
                        <a href={this.state.linkToDoc} target="_blank">
                          {this.state.documentName}</a>
                      </Label>
                    </div>
                  </div>
                  <div className={styles.divrow}>
                    <div className={styles.wdthfrst} style={{ marginTop: "20px" }}>
                      <div style={{ display: this.state.createDocumentCheckBoxDiv }}>
                        <TooltipHost
                          content="Check if the template or attachment is added"
                          //id={tooltipId}
                          calloutProps={calloutProps}
                          styles={hostStyles}>
                          {/* <Checkbox label="Create Document ? " boxSide="start"
                            onChange={this._onCreateDocChecked}
                            checked={this.state.createDocument} /> */}
                          <Label> Create Document :</Label>
                        </TooltipHost>
                      </div>
                      <div style={{ display: this.state.replaceDocument }}>
                        <TooltipHost
                          content="Check if the template or attachment is added"
                          //id={tooltipId}
                          calloutProps={calloutProps}
                          styles={hostStyles}>
                          <Checkbox label="Replace Document ? " boxSide="start"
                            onChange={this._onReplaceDocumentChecked}
                            checked={this.state.replaceDocumentCheckbox} />
                        </TooltipHost>
                      </div>
                    </div>
                    <div className={styles.wdthmid} style={{ display: this.state.hideDoc }}>
                      <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                        onChange={this.onUploadOrTemplateRadioBtnChange}
                        options={uploadOrTemplateRadioBtnOptions} styles={choiceGroupStyles}
                      /></div>

                  </div>
                  <div className={styles.divrow} style={{ display: this.state.hideupload, marginTop: "10px" }}>
                    <div className={styles.wdthfrst}><Label>Upload Document:</Label></div>
                    <div className={styles.wdthmid}>  <input type="file" name="myFile" id="editqdms" onChange={this._add}></input></div>
                    <div style={{ display: this.state.validDocType, color: "#dc3545" }}>Please select valid Document </div>
                    <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
                  </div>
                  <div className={styles.divrow} >
                    <div className={styles.wdthrgt} style={{ display: this.state.hidesource }}>
                      <Dropdown id="t7"
                        label="Source"
                        placeholder="Select an option"
                        selectedKey={this.state.sourceId}
                        options={Source}
                        onChanged={this._sourcechange} /></div>
                    <div className={styles.wdthlft} style={{ display: this.state.hidetemplate }}>
                      <Dropdown id="t7"
                        label="Select a Template"
                        placeholder="Select an option"
                        selectedKey={this.state.templateId}
                        options={this.state.templateDocuments}
                        onChanged={this._templatechange} /></div>
                  </div>
                  {/* <div style={{ display: this.state.hideDoc }} >
                      <div className={styles.wdthmid} style={{ marginTop: "10px" }}> <Label>Upload Document: </Label></div>
                      <div className={styles.wdthlst} style={{ marginTop: "10px" }}> <input type="file" name="myFile" id="editqdms" onChange={(e) => this._add(e)} ref={ref => this.myfile = ref}></input></div>
                      <div style={{ display: this.state.validDocType, color: "#dc3545" }}>Please select valid Document </div>
                      <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
                    </div>
                    <div className={styles.wdthlst} style={{ display: this.state.hideSelectTemplate }} >
                      <Dropdown id="t7"
                        label="Select a Template"
                        placeholder="Select an option"
                        selectedKey={this.state.templateId}
                        options={this.state.templateDocuments}
                        onChanged={this._templatechange} />
                    </div> */}
                  {/* </div > */}
                  <div className={styles.divrow}>
                    <div className={styles.wdthfrst} style={{ display: this.state.hideDirect }}>
                      <TooltipHost
                        content="The document to published library without sending it for review/approval"
                        //id={tooltipId}
                        calloutProps={calloutProps}
                        styles={hostStyles}>
                        <Checkbox label="Direct Publish?" boxSide="start" onChange={this._onDirectPublishChecked} checked={this.state.directPublishCheck} />
                      </TooltipHost></div>
                    <div className={styles.wdthmid} style={{ display: this.state.checkdirect }}><Spinner label={'Please Wait...'} /></div>
                    <div className={styles.wdthmid} style={{ display: this.state.hidePublish }}>
                      <DatePicker label="Published Date"
                        style={{ width: '200px' }}
                        value={this.state.approvalDateEdit}
                        onSelectDate={this._onApprovalDatePickerChange}
                        placeholder="Select a date..."
                        ariaLabel="Select a date" minDate={new Date()} maxDate={new Date()}
                        formatDate={this._onFormatDate} /></div>
                    <div className={styles.wdthlst} style={{ display: this.state.hidePublish }}>
                      <div style={{ display: this.state.isdocx }}>
                        <Dropdown id="t2" required={true}
                          label="Publish Option"
                          selectedKey={this.state.publishOption}
                          placeholder="Select an option"
                          options={publishOptions}
                          onChanged={this._publishOptionChange} /></div>
                      <div style={{ display: this.state.nodocx }}>
                        <Dropdown id="t2" required={true}
                          label="Publish Option"
                          selectedKey={this.state.publishOption}
                          placeholder="Select an option"
                          options={publishOption}
                          onChanged={this._publishOptionChange} /></div>
                      <div style={{ color: "#dc3545" }}>
                        {this.validator.message("publish", this.state.publishOption, "required")}{""}</div>
                    </div>
                  </div>
                </div>
                {/* <div><Label>Do you want to replace the document</Label>
                <IconButton iconProps={Go} title="Update Document" ariaLabel="Update Document" onClick={() => this._updatefileDocument()} /></div> */}

                <div className={styles.divrow}>
                  <div className={styles.wdthfrst} style={{ marginTop: "10px" }}> <TooltipHost
                    content="Is the document Critical?"
                    calloutProps={calloutProps}
                    styles={hostStyles}>
                    <Checkbox label="Critical Document? " boxSide="start" onChange={this._onCriticalChecked} checked={this.state.criticalDocument} />
                  </TooltipHost></div>
                  <div className={styles.wdthmid} style={{ marginTop: "10px" }}> <TooltipHost
                    content="Is the document a template?"
                    //id={tooltipId}
                    calloutProps={calloutProps}
                    styles={hostStyles}>
                    <Checkbox label="Template document? " boxSide="start" onChange={this._onTemplateChecked} checked={this.state.templateDocument} />
                  </TooltipHost></div>
                </div>
                <div style={{ display: this.state.messageBar }}>
                  {/* Show Message bar for Notification*/}
                  {this.state.statusMessage.isShowMessage ?
                    <MessageBar
                      messageBarType={this.state.statusMessage.messageType}
                      isMultiline={false}
                      dismissButtonAriaLabel="Close"
                    >{this.state.statusMessage.message}</MessageBar>
                    : ''}
                </div>
                <div className={styles.mt}>
                  <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
                </div>
                <div className={styles.mt}>
                  <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Updating...'} /></div>
                </div>
                <div className={styles.mt}>
                  <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                    <Label>***PLEASE DON'T REFRESH***</Label>
                  </div>
                </div>
                <DialogFooter>

                  <div className={styles.rgtalign}>
                    <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                  </div>
                  <div style={{ display: this.state.hidebutton }} >
                    <div className={styles.rgtalign} >
                      <PrimaryButton id="b2" className={styles.btn} disabled={this.state.updateDisable} onClick={this._onUpdateClick} >Update</PrimaryButton >
                      <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
                    </div>
                  </div>
                </DialogFooter>
                {/* {/ {/ Cancel Dialog Box /} /} */}

                <div>
                  <Dialog
                    hidden={this.state.confirmDialog}
                    dialogContentProps={this.dialogContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>

                <br />

                {/* editDocument div close */}

              </PivotItem>

              <PivotItem headerText="Version History">
                <div>
                  <IconButton iconProps={back} title="Back" onClick={this._back} />
                  <Iframe id="iframeModal" url={this.props.siteUrl + "/_layouts/15/Versions.aspx?list={" + this.sourceDocumentLibraryId + "}&ID=" + this.sourceDocumentID + "&IsDlg=0"}
                    width={"100%"} frameBorder={0}
                    height={"500rem"} />
                  {/* <iframe src={"https://ccsdev01.sharepoint.com/sites/TrialTest/_layouts/15/Versions.aspx?list=%7Bda53146b-3f5c-4321-926e-c3c2adbff323%7D&ID=1&IsDlg=0"} style={{overflow: "hidden",width:"100%",border:"white"}}></iframe>*/}
                </div>
              </PivotItem>
              <PivotItem headerText="RevisionHistory">
                <div>
                  <IconButton iconProps={back} title="Back" onClick={this._back} />
                  <Iframe id="iframeModal" url={this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID}
                    width={"100%"} frameBorder={0}
                    height={"500rem"} />
                </div>
              </PivotItem>
            </Pivot>
          </div>
        </div>
        {/* Edit document view for Project */}
        <div style={{ display: this.state.projectEditDocumentView }} >
          <div className={styles.border}>
            <div className={styles.alignCenter}>{this.props.webpartHeader + ":" + this.state.documentid}</div>
            <Pivot aria-label="Links of Tab Style Pivot Example"  >
              <PivotItem headerText="Document Info " >
                <div>
                  <TextField required id="t1"
                    label="Title"
                    onChange={this._titleChange}
                    value={this.state.title} readOnly={this.state.titleReadonly} ></TextField>
                  <div style={{ color: "#dc3545", fontWeight: "bold", display: this.state.checkrename }}>Checking your permission to rename.Please wait...</div>
                  <div style={{ color: "#dc3545" }}>
                    {this.validator.message("Name", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthfrst}>
                    <TextField
                      label="Department"
                      value={this.state.department} readOnly></TextField>
                  </div>
                  <div className={styles.wdthmid}>
                    <TextField
                      label="Category"
                      value={this.state.category} readOnly></TextField>
                  </div>
                  <div className={styles.wdthlst}>
                    <PeoplePicker
                      context={this.props.context as any}
                      titleText="Owner"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users    
                      showtooltip={true}
                      required={false}
                      disabled={false}
                      ensureUser={true}
                      onChange={(items) => this._selectedOwner(items)}
                      defaultSelectedUsers={[this.state.ownerName]}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Owner", this.state.owner, "required")}{" "}
                    </div>
                  </div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthfrst}>
                    <PeoplePicker
                      context={this.props.context as any}
                      titleText="Document Controller"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users    
                      showtooltip={true}
                      required={false}
                      disabled={false}
                      ensureUser={true}
                      onChange={(items) => this._selectedDCC(items)}
                      defaultSelectedUsers={[this.state.dccName]}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("DocumentController", this.state.dcc, "required")}{" "}</div>
                  </div>
                  <div className={styles.wdthmid} >
                    <PeoplePicker
                      context={this.props.context as any}
                      titleText="Reviewer(s)"
                      personSelectionLimit={20}
                      groupName={""} // Leave this blank in case you want to filter from all users
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      ensureUser={true}
                      showHiddenInUI={false}
                      onChange={(items) => this._selectedReviewers(items)}
                      defaultSelectedUsers={this.state.reviewersName}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                  </div>
                  <div className={styles.wdthlst}>
                    <PeoplePicker
                      context={this.props.context as any}
                      titleText="Approver"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users    
                      showtooltip={true}
                      required={false}
                      disabled={false}
                      ensureUser={true}
                      onChange={(items) => this._selectedApprover(items)}
                      showHiddenInUI={false}
                      defaultSelectedUsers={[this.state.approverName]}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />

                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
                  </div>
                </div>
                <div style={{ display: this.state.hideCreate }}>
                  <div className={styles.divrow}>
                    <div style={{ display: this.state.replaceDocument }}>
                      <Label >Document :
                        <a href={this.state.linkToDoc} target="_blank">
                          {this.state.documentName}</a>
                      </Label>
                    </div>
                  </div>
                  {this.state.hideDoc != "none" && <div>
                    <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                      onChange={this.onUploadOrTemplateRadioBtnChange}
                      options={uploadOrTemplateRadioBtnOptions} styles={choiceGroupStyles}
                    /></div>}
                  <div className={styles.divrow}>
                    <div className={styles.wdthfrst} >

                      {this.state.replaceDocument != "none" && <div>
                        <TooltipHost
                          content="Check if the template or attachment is added"
                          //id={tooltipId}
                          calloutProps={calloutProps}
                          styles={hostStyles}>
                          <Checkbox label="Replace Document ? " boxSide="start"
                            onChange={this._onReplaceDocumentChecked}
                            checked={this.state.replaceDocumentCheckbox} />
                        </TooltipHost>
                      </div>}
                    </div>

                  </div>
                  <div className={styles.divrow} style={{ display: this.state.hideupload, marginTop: "10px" }}>
                    <div className={styles.wdthfrst}> <Label>Upload Document:</Label></div>
                    <div className={styles.wdthmid}>  <input type="file" name="myFile" id="editproject" onChange={this._add}></input></div>
                    <div style={{ display: this.state.validDocType, color: "#dc3545" }}>Please select valid Document </div>
                    <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
                  </div>
                  <div className={styles.divrow} >
                    <div className={styles.wdthrgt} style={{ display: this.state.hidesource }}>
                      <Dropdown id="t7"
                        label="Source"
                        placeholder="Select an option"
                        selectedKey={this.state.sourceId}
                        options={Source}
                        onChanged={this._sourcechange} /></div>
                    <div className={styles.wdthlft} style={{ display: this.state.hidetemplate }}>
                      <Dropdown id="t7"
                        label="Select a Template"
                        placeholder="Select an option"
                        selectedKey={this.state.templateId}
                        options={this.state.templateDocuments}
                        onChanged={this._templatechange} /></div>
                  </div>
                  <div className={styles.divrow}>
                    <div className={styles.wdthfrst} style={{ marginTop: '3em' }}>
                      <TooltipHost
                        content="Is the document a template?"
                        //id={tooltipId}
                        calloutProps={calloutProps}
                        styles={hostStyles}>
                        <Checkbox label="Template Document? " boxSide="start" onChange={this._onTemplateChecked} checked={this.state.templateDocument} />
                      </TooltipHost>

                    </div>
                    <div className={styles.wdthmid} style={{ display: this.state.hideDirect, marginTop: '3em' }}>
                      <TooltipHost
                        content="The document to published library without sending it for review/approval"
                        //id={tooltipId}
                        calloutProps={calloutProps}
                        styles={hostStyles}>
                        <Checkbox label="Direct Publish?" boxSide="start" onChange={this._onDirectPublishChecked} checked={this.state.directPublishCheck} />
                      </TooltipHost>
                    </div>
                    <div className={styles.wdthlst} style={{ display: this.state.hidePublish }}>
                      <div style={{ display: this.state.isdocx }}>
                        <Dropdown id="t2" required={true}
                          label="Publish Option"
                          selectedKey={this.state.publishOption}
                          placeholder="Select an option"
                          options={publishOptions}
                          onChanged={this._publishOptionChange} /></div>
                      <div style={{ display: this.state.nodocx }}>
                        <Dropdown id="t2" required={true}
                          label="Publish Option"
                          selectedKey={this.state.publishOption}
                          placeholder="Select an option"
                          options={publishOption}
                          onChanged={this._publishOptionChange} /></div>
                      <div style={{ color: "#dc3545" }}>
                        {this.validator.message("publish", this.state.publishOption, "required")}{""}</div>
                    </div>
                  </div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthfrst}>
                    <Dropdown id="t2" required={true}
                      label="Revision Coding"
                      selectedKey={this.state.revisionCodingId}
                      placeholder="Select an option"
                      options={this.state.revisionSettingsArray}
                      onChanged={this._revisionCodingChange} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Revision", this.state.revisionCodingId, "required")}{" "}</div>
                  </div>
                  <div className={styles.wdthmid}>
                    <TextField label="Sub-Contractor Document Number" onChange={this._subContractorNumberChange} value={this.state.subContractorNumber}></TextField>
                  </div>
                  <div className={styles.wdthlst}>
                    <TextField label="Customer Document Number" onChange={this._CustomerNumberChange} value={this.state.customerNumber}></TextField>
                  </div>
                </div>

                <div className={styles.divrow}>
                  <div className={styles.wdthfrst} style={{ display: "flex" }}>
                    <div style={{ width: "13em" }}><DatePicker label="Expiry Date"
                      value={this.state.expiryDate}
                      onSelectDate={this._onExpDatePickerChange}
                      placeholder="Select a date..."
                      ariaLabel="Select a date" minDate={this.today}
                      formatDate={this._onFormatDate} /></div>
                    <div style={{ marginLeft: "1em", width: "13em" }}>
                      <TextField id="ExpiryLeadPeriod" name="ExpiryLeadPeriod"
                        label="Expiry Reminder(Days)" onChange={this._expLeadPeriodChange}
                        value={this.state.expiryLeadPeriod}>
                      </TextField>
                    </div>
                  </div>
                  <div className={styles.wdthmid} style={{ display: "flex" }}>
                    <div style={{ marginTop: "3em" }}>
                      <TooltipHost
                        content="Check if the document is for transmittal"
                        //id={tooltipId}
                        calloutProps={calloutProps}
                        styles={hostStyles}>
                        <Checkbox label="Transmittal Document ? " boxSide="start"
                          onChange={this._onTransmittalChecked}
                          checked={this.state.transmittalCheck} />
                      </TooltipHost>
                    </div>
                    <div style={{ marginTop: "3em", marginLeft: "1em" }}>
                      <TooltipHost
                        content="Check if the document is a subcontractor document"
                        //id={tooltipId}
                        calloutProps={calloutProps}
                        styles={hostStyles}>
                        <Checkbox label="External Document ? " boxSide="start"
                          onChange={this._onExternalDocumentChecked}
                          checked={this.state.externalDocument} />
                      </TooltipHost>
                    </div>
                  </div>
                </div>
                <div className={styles.divrow}>
                  <div className={styles.wdthfrst} style={{ display: "flex" }}>
                    <div style={{ width: "13em" }}> </div>
                    <div style={{ marginLeft: "1em", width: "13em" }}>
                    </div>
                    <div style={{ color: "#dc3545", display: this.state.leadmsg }}>
                      Enter only numbers less than 100
                    </div>
                  </div>
                </div>
                <div style={{ display: this.state.messageBar }}>
                  {/* Show Message bar for Notification*/}
                  {this.state.statusMessage.isShowMessage ?
                    <MessageBar
                      messageBarType={this.state.statusMessage.messageType}
                      isMultiline={false}
                      dismissButtonAriaLabel="Close"
                    >{this.state.statusMessage.message}</MessageBar>
                    : ''}
                </div>
                <div className={styles.mt}>
                  <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
                </div>
                <div className={styles.mt}>
                  <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Updating...'} /></div>
                </div>
                <div className={styles.mt}>
                  <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                    <Label>***PLEASE DON'T REFRESH***</Label>
                  </div>
                </div>
                <div className={styles.mandatory}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                <DialogFooter>
                  <div className={styles.rgtalign} >
                    <PrimaryButton id="b2" className={styles.btn} disabled={this.state.updateDisable} onClick={this._onUpdateClick} >Update</PrimaryButton >
                    <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
                  </div>
                </DialogFooter>
                {/* {/ {/ Cancel Dialog Box /} /} */}

                <div>
                  <Dialog
                    hidden={this.state.confirmDialog}
                    dialogContentProps={this.dialogContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>

                <br />
                {/* editDocument div close */}
              </PivotItem>
              <PivotItem headerText="Version History">
                <div>
                  <IconButton iconProps={back} title="Back" onClick={this._back} />
                  <Iframe id="iframeModal" url={this.props.siteUrl + "/_layouts/15/Versions.aspx?list={" + this.sourceDocumentLibraryId + "}&ID=" + this.sourceDocumentID + "&IsDlg=0"}
                    width={"100%"} frameBorder={0}
                    height={"500rem"} />
                  {/* <iframe src={"https://ccsdev01.sharepoint.com/sites/TrialTest/_layouts/15/Versions.aspx?list=%7Bda53146b-3f5c-4321-926e-c3c2adbff323%7D&ID=1&IsDlg=0"} style={{overflow: "hidden",width:"100%",border:"white"}}></iframe>*/}
                </div>
              </PivotItem>
              <PivotItem headerText="RevisionHistory">
                <div>
                  <IconButton iconProps={back} title="Back" onClick={this._back} />
                  <Iframe id="iframeModal" url={this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID}
                    width={"100%"} frameBorder={0}
                    height={"500rem"} />
                </div>
              </PivotItem>
              <PivotItem headerText="TransmittalHistory">
                <div>
                  <IconButton iconProps={back} title="Back" onClick={this._back} />
                  <Iframe id="iframeModal" url={this.props.siteUrl + "/SitePages/" + this.props.transmittalHistory + ".aspx?did=" + this.documentIndexID}
                    width={"100%"} frameBorder={0}
                    height={"500rem"} />
                </div>
              </PivotItem>
            </Pivot>
          </div>
        </div>
        <div style={{ display: this.state.accessDeniedMessageBar }}>
          {/* Show Message bar for Notification*/}
          {this.state.statusMessage.isShowMessage ?
            <MessageBar
              messageBarType={this.state.statusMessage.messageType}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >{this.state.statusMessage.message}</MessageBar>
            : ''}
        </div>
      </section>
    );
  }
}