import * as React from 'react';
import styles from './TransmittalSendRequest.module.scss';
import { ITransmittalSendRequestProps, ITransmittalSendRequestState } from './ITransmittalSendRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { BaseService } from '../services';
import SimpleReactValidator from 'simple-react-validator';
import { Checkbox, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, ITooltipHostStyles, Label, Link, MessageBar, PrimaryButton, ProgressIndicator, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as moment from 'moment';
import * as _ from 'lodash';
import { IHttpClientOptions, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
export default class TransmittalSendRequest extends React.Component<ITransmittalSendRequestProps, ITransmittalSendRequestState, {}> {
  private _Service: BaseService;
  private validator: SimpleReactValidator;
  private documentIndexID: any;
  private getSelectedReviewers: any[];
  private invalidUser: string;
  private currentEmail: string;
  private currentId: any;
  private today: any;
  // private time;
  private workflowStatus: string;
  private sourceDocumentID: any;
  private newheaderid: any;
  private newDetailItemID: any;
  private dccReview: string;
  private underApproval: any;
  private underReview: any;
  private invalidSendRequestLink: string;
  // private getSelectedReviewers = [];
  // private valid;
  private noDocument: string;
  private taskDelegate = "No";
  private taskDelegateDccReview: string;
  private taskDelegateUnderApproval: string;
  private taskDelegateUnderReview: string;
  // private departmentExists;
  private postUrl: string;
  private postUrlForUnderReview: string;
  // private permissionpostUrl;
  private postUrlForAdaptive: string;
  private TaskID: any;
  public constructor(props: ITransmittalSendRequestProps) {
    super(props);
    this.state = {

      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      documentID: "",
      linkToDoc: "",
      documentName: "",
      revision: "",
      ownerName: "",
      currentUser: "",
      hideProject: true,
      revisionLevel: [],
      revisionLevelvalue: "",
      dcc: "",
      reviewer: "",
      dueDate: "",
      approver: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      saveDisable: "",
      requestSend: 'none',
      statusKey: "",
      access: "none",
      accessDeniedMsgBar: "none",
      reviewers: [],
      ownerId: "",
      delegatedToId: "",
      delegateToIdInSubSite: "",
      delegateForIdInSubSite: "",
      reviewerEmail: "",
      reviewerName: "",
      delegatedFromId: "",
      detailIdForReviewer: "",
      approverEmail: "",
      approverName: "",
      hubSiteUserId: 0,
      detailIdForApprover: "",
      criticalDocument: "",
      dccReviewerName: "",
      dccReviewerEmail: "",
      dccReviewer: "",
      revisionLevelArray: [],
      revisionCoding: "",
      currentUserReviewer: [],
      projectName: "",
      projectNumber: "",
      acceptanceCodeId: "",
      transmittalRevision: "",
      reviewersName: [],
      hideLoading: true,
      sameRevision: false,
      loaderDisplay: "",
      businessUnitID: null,
      departmentId: null,
      validApprover: "none"
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    // this.componentDidMount = this.componentDidMount.bind(this);
    // this._userMessageSettings = this._userMessageSettings.bind(this);
    // this._queryParamGetting = this._queryParamGetting.bind(this);
    // this._accessGroups = this._accessGroups.bind(this);
    // this._checkWorkflowStatus = this._checkWorkflowStatus.bind(this);
    // this._openRevisionHistory = this._openRevisionHistory.bind(this);
    // this._bindSendRequestForm = this._bindSendRequestForm.bind(this);
    // this._project = this._project.bind(this);
    // this._revisionLevelChanged = this._revisionLevelChanged.bind(this);
    // this._dccReviewerChange = this._dccReviewerChange.bind(this);
    // this._reviewerChange = this._reviewerChange.bind(this);
    // this._approverChange = this._approverChange.bind(this);
    // this._submitSendRequest = this._submitSendRequest.bind(this);
    // this._dccReview = this._dccReview.bind(this);
    // this._underApprove = this._underApprove.bind(this);
    // this._underReview = this._underReview.bind(this);
    // this._underProjectApprove = this._underProjectApprove.bind(this);
    // this._underProjectReview = this._underProjectReview.bind(this);
    // this._onSameRevisionChecked = this._onSameRevisionChecked.bind(this);
    // this._adaptiveCard = this._adaptiveCard.bind(this);
    // this._LaUrlGettingAdaptive = this._LaUrlGettingAdaptive.bind(this);
  }
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "This field is mandatory"
      }
    });
  }
  //Page Load
  public async componentDidMount() {
    // this.reqWeb = Web(window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    // this.redirectUrl = this.props.redirectUrl;
    this.setState({ access: "none", accessDeniedMsgBar: "none" });
    // Get User Messages
    await this._userMessageSettings();
    //Get Current User
    const user = await this._Service.getCurrentUser();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    //Get Parameter from URL
    this._queryParamGetting();

    if (this.props.project) {
      this.setState({ hideProject: false });
    }

    let currentUserReviewer = [];
    currentUserReviewer.push(this.currentId);
    //Get Today
    this.today = new Date();
    this.setState({
      currentUser: user.Title,
      currentUserReviewer: currentUserReviewer
    });


  }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.gethubUserMessageListItems(this.props.hubUrl, this.props.userMessageSettings);
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "InvalidSendRequestUser") {
        this.invalidUser = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "InvalidSendRequestLink") {
        this.invalidSendRequestLink = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "NoDocument") {
        this.noDocument = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "WorkflowStatusError") {
        this.workflowStatus = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "DccReview") {
        var DccReview = userMessageSettings[i].Message;
        this.dccReview = replaceString(DccReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "UnderApproval") {
        var UnderApproval = userMessageSettings[i].Message;
        this.underApproval = replaceString(UnderApproval, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "UnderReview") {
        var UnderReview = userMessageSettings[i].Message;
        this.underReview = replaceString(UnderReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "TaskDelegateDccReview") {
        var TaskDelegateDccReview = userMessageSettings[i].Message;
        this.taskDelegateDccReview = replaceString(TaskDelegateDccReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "TaskDelegateUnderApproval") {
        var TaskDelegateUnderApproval = userMessageSettings[i].Message;
        this.taskDelegateUnderApproval = replaceString(TaskDelegateUnderApproval, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "TaskDelegateUnderReview") {
        var TaskDelegateUnderReview = userMessageSettings[i].Message;
        this.taskDelegateUnderReview = replaceString(TaskDelegateUnderReview, '[DocumentName]', this.state.documentName);

      }
    }

  }
  //Get Parameter from URL
  private async _queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentindexid = params.get('did');

    if (documentindexid != "" && documentindexid != null) {
      this.documentIndexID = parseInt(documentindexid);
      //Get Access
      this.setState({ access: "none", accessDeniedMsgBar: "none" });
      if (this.props.project) {
        await this._checkWorkflowStatus();
        // this._checkPermission('Project_SendRequest');
      }
      else {
        // await this._accessGroups();
        await this._checkWorkflowStatus();
      }
    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidSendRequestLink, messageType: 1 },
      });
      setTimeout(() => {
        window.location.replace(this.props.siteUrl);
      }, 10000);
    }
  }
  //Workflow Status Checking
  private async _checkWorkflowStatus() {
    const documentIndexItem: any = await this._Service.getIndexData(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
    if (documentIndexItem.WorkflowStatus == "Under Review" || documentIndexItem.WorkflowStatus == "Under Approval") {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.workflowStatus, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else if (documentIndexItem.DocumentStatus != "Active") {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: "Document is not currently Active", messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else if (documentIndexItem.SourceDocument == null) {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        access: "none",
        statusMessage: { isShowMessage: true, message: this.noDocument, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
    else {
      this.setState({ access: "", accessDeniedMsgBar: "none", loaderDisplay: "none" });
      await this._bindSendRequestForm();
    }
    if (this.props.project) {
      this.setState({ hideProject: false });
      await this._project();
    }
  }
  //Bind Send Request Form
  public async _bindSendRequestForm() {
    await this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID).then(async indexItems => {
      console.log("dataForEdit", indexItems);
      let documentID;
      let documentName;
      let ownerName;
      let ownerId;
      let revision;
      let linkToDocument;
      let criticalDocument;
      let approverName;
      let approverId;
      let approverEmail;
      let temReviewersID = [];
      let tempReviewers = [];
      let businessUnitID;
      let departmentId;
      //Get Document Index
      const documentIndexItem: any = await this._Service.getIndexDataId(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
      console.log(documentIndexItem);
      documentID = documentIndexItem.DocumentID;
      documentName = documentIndexItem.DocumentName;
      ownerName = documentIndexItem.Owner.Title;
      ownerId = documentIndexItem.Owner.ID;
      revision = documentIndexItem.Revision;
      linkToDocument = documentIndexItem.SourceDocument.Url;
      // this.SourceDocumentID = DocumentIndexItem.SourceDocumentID;
      criticalDocument = documentIndexItem.CriticalDocument;
      approverName = documentIndexItem.Approver.Title;
      approverId = documentIndexItem.Approver.ID;
      approverEmail = documentIndexItem.Approver.EMail;
      businessUnitID = documentIndexItem.BusinessUnitID;
      departmentId = documentIndexItem.DepartmentID;
      for (var k in documentIndexItem.Reviewers) {
        temReviewersID.push(documentIndexItem.Reviewers[k].ID);
        this.setState({
          reviewers: temReviewersID,
        });
        tempReviewers.push(documentIndexItem.Reviewers[k].Title);
      }
      if (indexItems.ApproverId != null) {
        this.setState({
          approver: documentIndexItem.Approver.ID,
          approverName: documentIndexItem.Approver.Title,
          approverEmail: documentIndexItem.Approver.EMail
        });
      }
      this.setState({
        documentID: documentID,
        documentName: documentName,
        ownerName: ownerName,
        ownerId: ownerId,
        revision: revision,
        linkToDoc: linkToDocument,
        criticalDocument: criticalDocument,
        approver: approverId,
        approverName: approverName,
        reviewersName: tempReviewers,
        businessUnitID: businessUnitID,
        departmentId: departmentId
      });
      const sourceDocumentItem: any = await this._Service.getSourceLibraryItems(this.props.siteUrl, this.props.sourceDocumentLibrary, this.documentIndexID);
      console.log(sourceDocumentItem);
      this.sourceDocumentID = sourceDocumentItem[0].ID;
      await this._userMessageSettings();
    });
  }
  // bind data for project
  public async _project() {
    await this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID).then(async (indexItems: any) => {
      console.log("dataForEdit", indexItems);
      let revisionLevelArray = [];
      let sorted_RevisionLevel = [];
      let revisionCoding;
      let transmittalRevision;
      let acceptanceCodeId;
      let documentControllerName;
      let documentControllerId;
      const documentIndexItem: any = await this._Service.getIndexProjectData(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
      console.log(documentIndexItem.RevisionCodingId);
      revisionCoding = documentIndexItem.RevisionCodingId;
      acceptanceCodeId = documentIndexItem.AcceptanceCodeId;
      transmittalRevision = documentIndexItem.TransmittalRevision;
      if (indexItems.DocumentControllerId != null) {
        this.setState({
          dccReviewer: documentIndexItem.DocumentController.ID,
          dccReviewerName: documentIndexItem.DocumentController.Title,
          dccReviewerEmail: documentIndexItem.DocumentController.EMail
        });
      }
      if (indexItems.RevisionLevelId != null) {
        this.setState({
          revisionLevelvalue: documentIndexItem.RevisionLevelId
        });
      }
      const revisionLevelItem: any = await this._Service.getRevisionLevelData(this.props.siteUrl, this.props.revisionLevelList);
      console.log(revisionLevelItem);
      for (let i = 0; i < revisionLevelItem.length; i++) {
        let revisionLevelItemdata = {
          key: revisionLevelItem[i].ID,
          text: revisionLevelItem[i].Title
        };
        revisionLevelArray.push(revisionLevelItemdata);
      }
      console.log(revisionLevelArray);
      sorted_RevisionLevel = _.orderBy(revisionLevelArray, 'text', ['asc']);
      this.setState({
        revisionLevelArray: sorted_RevisionLevel,
        revisionCoding: revisionCoding,
        acceptanceCodeId: acceptanceCodeId,
        transmittalRevision: transmittalRevision,
      });
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
    });
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.documentIndexID);
  }
  //Same Revision Checked
  public _onSameRevisionChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) { this.setState({ sameRevision: true }); }
    else if (!isChecked) { this.setState({ sameRevision: false }); }
  }
  // on dccreviewer change
  public _dccReviewerChange = (items: any[]) => {
    this.setState({ saveDisable: "" });
    let dccreviewerEmail;
    let dccreviewerName;
    console.log(items);
    let getSelecteddccreviewer = [];
    for (let item in items) {
      dccreviewerEmail = items[item].secondaryText,
        dccreviewerName = items[item].text,
        getSelecteddccreviewer.push(items[item].id);
    }
    this.setState({
      dccReviewer: getSelecteddccreviewer[0],
      dccReviewerEmail: dccreviewerEmail,
      dccReviewerName: dccreviewerName
    });
  }
  // on reviewer change
  public _reviewerChange = (items: any[]) => {
    this.setState({ saveDisable: "" });
    console.log(items);
    this.getSelectedReviewers = [];
    for (let item in items) {
      this.getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: this.getSelectedReviewers });
    console.log(this.getSelectedReviewers);
  }
  // on approver change
  public _approverChange = async (items: any[]) => {
    this.setState({ saveDisable: "" });
    let approverEmail;
    let approverName;
    console.log(items);
    let getSelectedApprover = [];
    if (this.props.project) {
      for (let item in items) {
        approverEmail = items[item].secondaryText,
          approverName = items[item].text,
          getSelectedApprover.push(items[item].id);
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName });
    }
    else {
      this.setState({ validApprover: "", approver: null, approverEmail: "", approverName: "", });
      if (this.state.businessUnitID != null) {
        const businessUnit = await this._Service.getApproverData(this.props.hubUrl, this.props.businessUnitList);
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
        const departments = await this._Service.getApproverData(this.props.hubUrl, this.props.departmentList);
        for (let i = 0; i < departments.length; i++) {
          if (departments[i].ID == this.state.departmentId) {
            const deptapprove = await this._Service.getByEmail(departments[i].Approver.EMail);
            approverEmail = departments[i].Approver.EMail;
            approverName = departments[i].Approver.Title;
            getSelectedApprover.push(deptapprove.Id);
          }
        }
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName });
      setTimeout(() => {
        this.setState({ validApprover: "none" });
      }, 5000);
    }
  }
  // on expirydate change
  private _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ saveDisable: "" });
    this.setState({ dueDate: date });
  }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, saveDisable: "" });
  }
  // on submit send request
  private _submitSendRequest = async () => {
    //this.setState({ saveDisable: true, hideLoading: false });
    let sorted_previousHeaderItems = [];
    let previousHeaderItem = 0;
    let dcc = "dcc";
    const previousHeaderItems = await this._Service.getpreviousheader(this.props.siteUrl, this.props.workflowHeaderList, Number(this.documentIndexID));
    if (previousHeaderItems.length != 0) {
      sorted_previousHeaderItems = _.orderBy(previousHeaderItems, 'ID', ['desc']);
      previousHeaderItem = sorted_previousHeaderItems[0].ID;
    }
    if (this.props.project) {
      if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate") && this.validator.fieldValid("DocumentController")) {
        if (this.state.dccReviewer != "" && this.state.dccReviewer != undefined) {
          this.setState({ saveDisable: "none", hideLoading: false });
          this._dccReview(previousHeaderItem);
        }
        else if (this.state.reviewers.length == 0) {
          this.setState({ saveDisable: "none", hideLoading: false });
          this._underProjectApprove(previousHeaderItem);
        }
        else {
          this.setState({ saveDisable: "none", hideLoading: false });
          this._underProjectReview(previousHeaderItem);
        }
        this.validator.hideMessages();
        this.setState({ requestSend: "" });
        setTimeout(() => this.setState({ requestSend: 'none' }), 3000);
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate")) {
        if (this.state.reviewers.length == 0) {
          this.setState({ saveDisable: "none", hideLoading: false });
          this._underApprove(previousHeaderItem);
        }
        else {
          this.setState({ saveDisable: "none", hideLoading: false });
          this._underReview(previousHeaderItem);
        }
        this.validator.hideMessages();
        this.setState({ requestSend: "" });
        setTimeout(() => this.setState({ requestSend: 'none', saveDisable: "none", }), 3000);
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  //  request to dcc review
  public async _dccReview(previousHeaderItem: any) {
    this._LAUrlGettingForUnderReview();
    // this._LaUrlGettingAdaptive();
    let headerdata = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: this.state.reviewers,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "DCC Review",
      DocumentControllerId: this.state.dccReviewer,

      RevisionCodingId: this.state.revisionCoding,
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString(),
      ApproveInSameRevision: this.state.sameRevision
    }
    const header = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata);
    if (header) {
      this.newheaderid = header.data.ID;
      let log1 = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      const log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, log1);
      //Task delegation getting user id from hubsite
      const user = await this._Service.getByhubEmail(this.state.dccReviewerEmail);
      if (user) {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });
        //Task delegation 
        const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubUrl, this.props.taskDelegationSettings, user.Id);
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let toDate = moment(taskDelegation[0].ToDate).toDate();
          let fromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
          fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
          if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
            this.taskDelegate = "Yes";
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
            //detail list adding an item for approval
            const DelegatedTo = await this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail);
            if (DelegatedTo) {
              this.setState({
                delegateToIdInSubSite: DelegatedTo.Id,
              });
              let deleateforID: any;
              this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail).then(async (DelegatedFor: any) => {
                this.setState({
                  delegateForIdInSubSite: DelegatedFor.Id,
                });
                let detaildata1 = {
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "DCC Review",
                  Title: this.state.documentName,
                  ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.dccReviewer),
                  DueDate: this.state.dueDate,
                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                  ResponseStatus: "Under Review",
                  SourceDocument: {
                    Description: this.state.documentName,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  OwnerId: this.state.ownerId,
                }
                const details = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata1)
                if (details) {
                  this.setState({ detailIdForApprover: details.data.ID });
                  this.newDetailItemID = details.data.ID;
                  let updatedetaildata1 = {
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + "&wf=dcc"
                    }
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetaildata1, details.data.ID);
                  let updatedelegateuser = {
                    DocumentControllerId: this.state.delegateToIdInSubSite,
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, updatedelegateuser, this.documentIndexID);
                  await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, updatedelegateuser, this.sourceDocumentID);
                  //MY tasks list updation
                  let taskdata1 = {
                    Title: "Document Controller Review '" + this.state.documentName + "'",
                    Description: "DCC Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                    Workflow: "DCC Review",
                    // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                    DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + "&wf=dcc"
                    }
                  }
                  const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata1);
                  if (task) {
                    this.TaskID = task.data.ID;
                    let taskdata2 = {
                      TaskID: task.data.ID,
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata2, details.data.ID);
                    //notification preference checking                                 
                    this._sendmail(this.state.approverEmail, "DocDCCReview", this.state.approverName);
                    // await this._adaptiveCard("DCC Review", this.state.approverEmail, this.state.approverName, "Project", task.data.ID)
                  }//taskID
                }//r

              })//DelegatedFor
            }//DelegatedTo
          }//duedate checking
          else {
            let detaildata2 = {
              HeaderIDId: Number(this.newheaderid),
              Workflow: "DCC Review",
              Title: this.state.documentName,
              ResponsibleId: this.state.dccReviewer,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Review",
              SourceDocument: {
                Description: this.state.documentName,
                Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
              },
              OwnerId: this.state.ownerId,
            }
            const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata2);
            if (detail) {
              this.setState({ detailIdForApprover: detail.data.ID });
              this.newDetailItemID = detail.data.ID;
              let updatedetaildata2 = {
                Link: {
                  Description: this.state.documentName + "-Review",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + "&wf=dcc"
                },
              }
              const updatedetail = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetaildata2, detail.data.ID);
              // my task updation
              let taskdata3 = {
                Title: "Document Controller Review '" + this.state.documentName + "'",
                Description: "DCC Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                DueDate: this.state.dueDate,
                StartDate: this.today,
                AssignedToId: user.Id,
                Workflow: "DCC Review",
                Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                Source: (this.props.project ? "Project" : "QDMS"),
                Link: {
                  Description: this.state.documentName + "-Review",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + "&wf=dcc"
                },

              }
              const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata3);
              if (task) {
                this.TaskID = task.data.ID;
                let taskdata4 = {
                  TaskID: task.data.ID,
                }
                let taskdetail = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata4, detail.data.ID);
                //notification preference checking                                 
                this._sendmail(this.state.dccReviewerEmail, "DocDCCReview", this.state.dccReviewerName);
                // await this._adaptiveCard("DCC Review", this.state.dccReviewerEmail, this.state.dccReviewerName, "Project", task.data.ID)
                let dccreviewer = {
                  DocumentControllerId: this.state.dccReviewer
                }
                await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, dccreviewer, this.documentIndexID);
                await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, dccreviewer, this.sourceDocumentID);
              }//taskID
            }//r 
          }

        }
        else {
          let detaildata3 = {
            HeaderIDId: Number(this.newheaderid),
            Workflow: "DCC Review",
            Title: this.state.documentName,
            ResponsibleId: this.state.dccReviewer,
            DueDate: this.state.dueDate,
            ResponseStatus: "Under Review",
            SourceDocument: {
              Description: this.state.documentName,
              Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            },
            OwnerId: this.state.ownerId,
          }
          const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata3);
          if (detail) {
            this.setState({ detailIdForApprover: detail.data.ID });
            this.newDetailItemID = detail.data.ID;
            let updatedetail1 = {
              Link: {
                Description: this.state.documentName + "-Review",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + "&wf=dcc"
              },
            }
            await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetail1, detail.data.ID);
            let taskdata5 = {
              Title: "Document Controller Review '" + this.state.documentName + "'",
              Description: "DCC Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
              DueDate: this.state.dueDate,
              StartDate: this.today,
              AssignedToId: user.Id,
              Workflow: "DCC Review",
              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
              Source: (this.props.project ? "Project" : "QDMS"),
              Link: {
                Description: this.state.documentName + "-Review",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + "&wf=dcc"
              },

            }
            // my task updation
            const task = await this._Service.createNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata5);
            if (task) {
              let taskdetail = {
                TaskID: task.data.ID,
              }
              this.TaskID = task.data.ID;
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdetail, detail.data.ID);
              //notification preference checking                                 
              await this._sendmail(this.state.dccReviewerEmail, "DocDCCReview", this.state.dccReviewerName);
              // await this._adaptiveCard("DCC Review", this.state.dccReviewerEmail, this.state.dccReviewerName, "Project", task.data.ID)
              let dccid = {
                DocumentControllerId: this.state.dccReviewer
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, dccid, this.documentIndexID);
              await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, dccid, this.sourceDocumentID);
            }//taskID
          }//r
        }//else no delegation
      }
      let indexdata = {
        WorkflowStatus: "Under Review",
        Workflow: "DCC Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
        WorkflowDueDate: this.state.dueDate
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, this.documentIndexID);
      let sourcedata = {
        WorkflowStatus: "Under Review",
        Workflow: "DCC Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata, this.sourceDocumentID);
      await this._triggerDocumentUnderReview(this.sourceDocumentID, "DCC Review");
      let logdata = {
        Title: this.state.documentID,
        Status: "Under Review",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        Workflow: "DCC Review",
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata)
      this.setState({
        comments: "",
        statusKey: "",
        approverEmail: "",
        approverName: "",
        approver: "",
      });
      if (this.taskDelegate == "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateDccReview, messageType: 4 },
        });
      }
      else {
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.dccReview, messageType: 4 },
        });
      }
      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);

    }//newheaderid
  }
  // la for under review permission
  private _LAUrlGettingForUnderReview = async () => {
    const laUrl = await this._Service.gettriggerUnderReviewPermission(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrlForUnderReview = laUrl[0].PostUrl;
  }

  // La for under approval permission
  private async _LAUrlGetting() {
    const laUrl = await this._Service.gettriggerUnderApprovalPermission(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;

  }
  //Send Mail
  public _sendmail = async (emailuser: any, type: any, name: any) => {
    console.log(name);
    let mailSend = "No";
    let Subject;
    let Body;
    let link;
    console.log(this.state.criticalDocument);
    const notificationPreference: any[] = await this._Service.getnotification(this.props.hubUrl, this.props.notificationPreference, emailuser);
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
      //console.log("Send mail for critical document");
      mailSend = "Yes";
    }

    if (mailSend == "Yes") {
      const emailNotification: any[] = await this._Service.getemail(this.props.hubUrl, this.props.emailNotification, type);
      console.log(emailNotification);
      Subject = emailNotification[0].Subject;
      Body = emailNotification[0].Body;
      if (type == "DocApproval") {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID}>Link</a>`;

      }
      else if (type == "DocDCCReview") {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID + "&wf=dcc"} >Link</a>`;
      }
      else {
        link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + this.newDetailItemID}>Link</a>`;
      }
      console.log(link);
      //Replacing the email body with current values
      let dueDateformail = moment(this.state.dueDate).format("DD/MM/YYYY");
      console.log(dueDateformail);
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      console.log(replacedSubject);
      let replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', dueDateformail);
      console.log(replacedSubjectWithDueDate);
      console.log('Body' + Body);
      console.log('name:' + name);
      let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
      console.log(replaceRequester);
      let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
      console.log(replaceBody);
      let replacelink = replaceString(replaceBody, '[Link]', link);
      console.log(replacelink);
      let var1: any[] = replacelink.split('/');
      console.log(var1)
      let FinalBody = replacelink;
      console.log(FinalBody)
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubjectWithDueDate,
          "body": {
            "contentType": "HTML",
            "content": FinalBody

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
  // set permission for approver
  protected async _triggerPermission(sourceDocumentID: any) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': "Under Approval"
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  // set permission for reviewer
  protected async _triggerDocumentUnderReview(sourceDocumentID: any, type: any) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    console.log(siteUrl);
    const postURL = this.postUrlForUnderReview;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': "Under Review",
      'Workflow': type
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //  qdms request to review
  public async _underReview(previousHeaderItem: any) {
    this._LAUrlGettingForUnderReview();
    this._LaUrlGettingAdaptive();
    let detaildata14 = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: this.state.reviewers,
      ApproverId: this.state.approver,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }
    const header = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowHeaderList, detaildata14);
    if (header) {
      this.newheaderid = header.data.ID;
      let logdata = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      const log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata);
      //for reviewers if exist
      for (var k = 0; k < this.state.reviewers.length; k++) {
        console.log(this.state.reviewers[k]);
        const user = await this._Service.getByUserId(this.state.reviewers[k]);
        if (user) {
          console.log(user);
          await this._Service.getByhubEmail(user.Email)
            .then(async (hubsieUser: any) => {
              console.log(hubsieUser.Id);
              //Task delegation 
              const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubUrl, this.props.taskDelegationSettings, hubsieUser.Id);
              console.log(taskDelegation);
              if (taskDelegation.length > 0) {
                let duedate = moment(this.state.dueDate).toDate();
                let toDate = moment(taskDelegation[0].ToDate).toDate();
                let fromDate = moment(taskDelegation[0].FromDate).toDate();
                duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                  this.taskDelegate = "Yes";
                  this.setState({
                    approverEmail: taskDelegation[0].DelegatedTo.EMail,
                    approverName: taskDelegation[0].DelegatedTo.Title,
                    delegatedToId: taskDelegation[0].DelegatedTo.ID,
                    delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                  });
                  //Get Delegated To ID
                  const DelegatedTo = this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail)
                    .then(async (DelegatedTo: any) => {
                      this.setState({
                        delegateToIdInSubSite: DelegatedTo.Id,
                      });
                      //Get Delegated For ID
                      const DelegatedFor = await this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail);
                      if (DelegatedFor) {
                        this.setState({
                          delegateForIdInSubSite: DelegatedFor.Id,
                        });
                        //detail list adding an item for reviewers
                        let index = this.state.reviewers.indexOf(DelegatedFor.Id);
                        console.log(index);
                        this.state.reviewers[index] = DelegatedTo.Id;
                        console.log(this.state.reviewers);
                        let detaildata13 = {
                          HeaderIDId: Number(this.newheaderid),
                          Workflow: "Review",
                          Title: this.state.documentName,
                          ResponsibleId: (this.state.delegatedToId != "" ? DelegatedTo.Id : user.Id),
                          DueDate: this.state.dueDate,
                          DelegatedFromId: (this.state.delegatedToId != "" ? DelegatedFor.Id : parseInt("")),
                          ResponseStatus: "Under Review",
                          SourceDocument: {
                            Description: this.state.documentName,
                            Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                          },
                          OwnerId: this.state.ownerId
                        }
                        const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata13);
                        if (detail) {
                          this.setState({ detailIdForApprover: detail.data.ID });
                          this.newDetailItemID = detail.data.ID;
                          let taskdata14 = {
                            Link: {
                              Description: this.state.documentName + "-Review",
                              Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                            }
                          }
                          await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata14, detail.data.ID);
                          //Update link

                          //MY tasks list updation with delegated from
                          let taskdata13 = {
                            Title: "Review '" + this.state.documentName + "'",
                            Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                            DueDate: this.state.dueDate,
                            StartDate: this.today,
                            AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                            Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                            DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                            Source: (this.props.project ? "Project" : "QDMS"),
                            DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                            Workflow: "Review",
                            Link: {
                              Description: this.state.documentName + "-Review",
                              Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                            }
                          }
                          const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata13);
                          if (task) {
                            this.TaskID = task.data.ID;
                            let taskdata12 = {
                              TaskID: task.data.ID,
                            }
                            await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata12, detail.data.ID);
                            await this._sendmail(DelegatedTo.Email, "DocReview", DelegatedTo.Title);
                            // await this._adaptiveCard("Review",DelegatedTo.Email,DelegatedTo.Title,"General",task.data.ID);
                          }//taskID
                        }//r
                      }//Delegated For
                    });//Delegated To
                }
                else {
                  //detail list adding an item for reviewers
                  let detaildata12 = {
                    HeaderIDId: Number(this.newheaderid),
                    Workflow: "Review",
                    Title: this.state.documentName,
                    ResponsibleId: user.Id,
                    DueDate: this.state.dueDate,
                    ResponseStatus: "Under Review",
                    SourceDocument: {
                      Description: this.state.documentName,
                      Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                    },
                    OwnerId: this.state.ownerId,
                  }
                  const details = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata12);
                  if (details) {
                    this.setState({ detailIdForApprover: details.data.ID });
                    this.newDetailItemID = details.data.ID;
                    let updatedetaildata4 = {
                      Link: {
                        Description: this.state.documentName + "-Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                      }
                    }
                    const updatedetail = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetaildata4, details.data.ID);
                    //MY tasks list updation with delegated from
                    let taskdata11 = {
                      Title: "Review '" + this.state.documentName + "'",
                      Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                      DueDate: this.state.dueDate,
                      StartDate: this.today,
                      AssignedToId: hubsieUser.Id,
                      Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                      Source: (this.props.project ? "Project" : "QDMS"),
                      Workflow: "Review",
                      Link: {
                        Description: this.state.documentName + "-Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                      }
                    }
                    const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata11);
                    if (task) {
                      let updatetaskdetail = {
                        TaskID: task.data.ID,
                      }
                      const updatetask = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatetaskdetail, details.data.ID);
                      if (updatetask) {
                        await this._sendmail(user.Email, "DocReview", user.Title);
                        // await  this._adaptiveCard("Review",user.Email,user.Title,"General",task.data.ID);
                      }
                    }//taskId
                  }//r
                }//else

              }//IF
              //If no task delegation
              else {
                //detail list adding an item for reviewers
                let detaildata11 = {
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Review",
                  Title: this.state.documentName,
                  ResponsibleId: user.Id,
                  DueDate: this.state.dueDate,
                  ResponseStatus: "Under Review",
                  SourceDocument: {
                    Description: this.state.documentName,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  OwnerId: this.state.ownerId,
                }
                const details = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata11);
                if (details) {
                  this.setState({ detailIdForApprover: details.data.ID });
                  this.newDetailItemID = details.data.ID;
                  let updatedetaildata3 = {
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                    }
                  }
                  const updatedetail = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetaildata3, details.data.ID);
                  //MY tasks list updation with delegated from
                  let taskdata10 = {
                    Title: "Review '" + this.state.documentName + "'",
                    Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: hubsieUser.Id,
                    Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    Workflow: "Review",
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + details.data.ID + ""
                    }
                  }
                  const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata10);
                  if (task) {
                    this.TaskID = task.data.ID;
                    let taskdata9 = {
                      TaskID: task.data.ID,
                    }
                    const updatetask = await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata9, details.data.ID);
                    if (updatetask) {
                      await this._sendmail(user.Email, "DocReview", user.Title);
                      // await  this._adaptiveCard("Review",user.Email,user.Title,"General",task.data.ID);
                    }
                  }//taskId
                }//r
              }//else
            });//hubsiteuser
        }//user
      }
      let indexdata2 = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
        WorkflowDueDate: this.state.dueDate
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata2, this.documentIndexID);
      let sourcedata = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata, this.sourceDocumentID);
      let headeritem1 = {
        ReviewersId: { results: this.state.reviewers }
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headeritem1, parseInt(this.newheaderid));
      await this._triggerDocumentUnderReview(this.sourceDocumentID, "Review");
      let logitem1 = {
        Title: this.state.documentID,
        Status: "Under Review",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Review",
        DueDate: this.state.dueDate,
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem1);
      if (this.taskDelegate == "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateUnderReview, messageType: 4 },
        });
      }
      else {
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.underReview, messageType: 4 },
        });
      }
      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);


    }
  }
  //  project request to review
  public async _underProjectReview(previousHeaderItem: any) {
    this._LAUrlGettingForUnderReview();
    this._LaUrlGettingAdaptive();
    let headerdata2 = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: this.state.reviewers,
      ApproverId: this.state.approver,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString(),
      RevisionCodingId: this.state.revisionCoding,
      TransmittalRevision: this.state.transmittalRevision,
      AcceptanceCodeId: this.state.acceptanceCodeId,
      ApproveInSameRevision: this.state.sameRevision
    }
    const header = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata2);
    if (header) {
      this.newheaderid = header.data.ID;
      let logdata1 = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata1);
      //for reviewers if exist
      for (var k = 0; k < this.state.reviewers.length; k++) {
        console.log(this.state.reviewers[k]);
        const user = await this._Service.getByUserId(this.state.reviewers[k]);
        if (user) {
          console.log(user);
          const hubsieUser = await this._Service.getByhubEmail(user.Email);
          if (hubsieUser) {
            console.log(hubsieUser.Id);
            const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubUrl, this.props.taskDelegationSettings, hubsieUser.Id);
            console.log(taskDelegation);
            //Check if Task Delegation
            if (taskDelegation.length > 0) {
              let duedate = moment(this.state.dueDate).toDate();
              let toDate = moment(taskDelegation[0].ToDate).toDate();
              let fromDate = moment(taskDelegation[0].FromDate).toDate();
              duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
              toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
              fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
              if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                this.taskDelegate = "Yes";
                this.setState({
                  approverEmail: taskDelegation[0].DelegatedTo.EMail,
                  approverName: taskDelegation[0].DelegatedTo.Title,
                  delegatedToId: taskDelegation[0].DelegatedTo.ID,
                  delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                });
                //Get Delegated To ID
                const DelegatedTo = await this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail);
                if (DelegatedTo) {
                  this.setState({
                    delegateToIdInSubSite: DelegatedTo.Id,
                  });
                  //Get Delegated For ID
                  const DelegatedFor = await this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail);
                  if (DelegatedFor) {
                    this.setState({
                      delegateForIdInSubSite: DelegatedFor.Id,
                    });
                    let index = this.state.reviewers.indexOf(DelegatedFor.Id);
                    console.log(index);
                    this.state.reviewers[index] = DelegatedTo.Id;
                    console.log(this.state.reviewers);
                    //detail list adding an item for reviewers
                    let detaildata21 = {
                      HeaderIDId: Number(this.newheaderid),
                      Workflow: "Review",
                      Title: this.state.documentName,
                      ResponsibleId: (this.state.delegatedToId != "" ? DelegatedTo.Id : user.Id),
                      DueDate: this.state.dueDate,
                      DelegatedFromId: (this.state.delegatedToId != "" ? DelegatedFor.Id : parseInt("")),
                      ResponseStatus: "Under Review",
                      SourceDocument: {
                        Description: this.state.documentName,
                        Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                      },
                      OwnerId: this.state.ownerId
                    }
                    const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata21);
                    if (detail) {
                      this.setState({ detailIdForApprover: detail.data.ID });
                      this.newDetailItemID = detail.data.ID;
                      let detaildata22 = {
                        Link: {
                          Description: this.state.documentName + "-Review",
                          Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                        },
                      }
                      //Update link
                      this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata22, detail.data.ID);
                      //MY tasks list updation with delegated from
                      let taskdata21 = {
                        Title: "Review '" + this.state.documentName + "'",
                        Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                        DueDate: this.state.dueDate,
                        StartDate: this.today,
                        AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                        Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                        DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                        Source: (this.props.project ? "Project" : "QDMS"),
                        DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                        Workflow: "Review",
                        Link: {
                          Description: this.state.documentName + "-Review",
                          Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                        }
                      }
                      const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata21)
                      if (task) {
                        this.TaskID = task.data.ID;
                        let taskdata22 = {
                          TaskID: task.data.ID,
                        }
                        this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata22, detail.data.ID);
                        //notification preference checking                                 
                        await this._sendmail(DelegatedTo.Email, "DocReview", DelegatedTo.Title);
                        // await this._adaptiveCard("Review",DelegatedTo.Email,DelegatedTo.Title,"Project",task.data.ID);
                      }//taskID
                    }//r
                  }//Delegated For
                }//Delegated To
              }
              else {
                //detail list adding an item for reviewers
                let detaildata23 = {
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Review",
                  Title: this.state.documentName,
                  ResponsibleId: user.Id,
                  DueDate: this.state.dueDate,
                  ResponseStatus: "Under Review",
                  SourceDocument: {
                    Description: this.state.documentName,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  OwnerId: this.state.ownerId
                }
                const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata23);
                if (detail) {
                  this.setState({ detailIdForApprover: detail.data.ID });
                  this.newDetailItemID = detail.data.ID;
                  let detaildata24 = {
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                    }
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata24, detail.data.ID);
                  //MY tasks list updation with delegated from
                  let taskdata23 = {
                    Title: "Review '" + this.state.documentName + "'",
                    Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: hubsieUser.Id,
                    Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    Workflow: "Review",
                    Link: {
                      Description: this.state.documentName + "-Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                    }
                  }
                  const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata23);
                  if (task) {
                    this.TaskID = task.data.ID;
                    let taskdata24 = {
                      TaskID: task.data.ID,
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata24, detail.data.ID);
                    //notification preference checking                                 
                    await this._sendmail(user.Email, "DocReview", user.Title);
                    // await this._adaptiveCard("Review", user.Email, user.Title, "Project", task.data.ID);
                  }//taskId
                }//r
              }//else

            }//IF
            //If no task delegation
            else {
              //detail list adding an item for reviewers
              let detaildata25 = {
                HeaderIDId: Number(this.newheaderid),
                Workflow: "Review",
                Title: this.state.documentName,
                ResponsibleId: user.Id,
                DueDate: this.state.dueDate,
                ResponseStatus: "Under Review",
                SourceDocument: {
                  Description: this.state.documentName,
                  Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                },
                OwnerId: this.state.ownerId
              }
              const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata25);
              if (detail) {
                this.setState({ detailIdForApprover: detail.data.ID });
                this.newDetailItemID = detail.data.ID;
                let detaildata26 = {
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                  }
                }
                await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata26, detail.data.ID);
                //MY tasks list updation with delegated from
                let taskdata25 = {
                  Title: "Review '" + this.state.documentName + "'",
                  Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: hubsieUser.Id,
                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Workflow: "Review",
                  Link: {
                    Description: this.state.documentName + "-Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                  }
                }
                const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata25);
                if (task) {
                  this.TaskID = task.data.ID;
                  let taskdata26 = {
                    TaskID: task.data.ID
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata26, detail.data.ID);
                  //notification preference checking                                 
                  await this._sendmail(user.Email, "DocReview", user.Title);
                  // await this._adaptiveCard("Review", user.Email, user.Title, "Project", task.data.ID);
                }//taskId
              }//r
            }//else
          }//hubsiteuser
        }//user
      }
      let indexdata3 = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        WorkflowDueDate: this.state.dueDate,
        ReviewersId: { results: this.state.reviewers }
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata3, this.documentIndexID);
      let sourcedata = {
        WorkflowStatus: "Under Review",
        Workflow: "Review",
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata, this.sourceDocumentID);
      let headerdata3 = {
        ReviewersId: { results: this.state.reviewers }
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata3, parseInt(this.newheaderid));
      await this._triggerDocumentUnderReview(this.sourceDocumentID, "Review");
      let logdata2 = {
        Title: this.state.documentID,
        Status: "Under Review",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Review",
        DueDate: this.state.dueDate
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata2)
        .then(msg => {
          this.setState({

            comments: "",
            statusKey: "",
            approverEmail: "",
            approverName: "",
            approver: "",
            delegateForIdInSubSite: ""
          });
          if (this.taskDelegate == "Yes") {
            this.setState({
              hideLoading: true,
              statusMessage: { isShowMessage: true, message: this.taskDelegateUnderReview, messageType: 4 },
            });
          }
          else {
            this.setState({
              hideLoading: true,
              statusMessage: { isShowMessage: true, message: this.underReview, messageType: 4 },
            });
          }
          setTimeout(() => {
            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
          }, 10000);
        });//msg

    }
  }
  //  qdms request to approve
  public async _underApprove(previousHeaderItem: any) {
    this._LAUrlGetting();
    this._LaUrlGettingAdaptive();
    let headerdata1 = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.currentUserReviewer },
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }
    const header = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata1);
    if (header) {
      this.newheaderid = header.data.ID;
      let logitem = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      const log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem);
      let detaildata15 = {
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
        SourceDocument: {
          Description: this.state.documentName,
          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
        },
        OwnerId: this.state.ownerId
      }
      const detailadd = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata15);
      if (detailadd) {
        this.setState({ detailIdForApprover: detailadd.data.ID });
        this.newDetailItemID = detailadd.data.ID;
        let detailtask = {
          Link: {
            Description: this.state.documentName + "-Review",
            Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailadd.data.ID + ""
          }
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detailtask, detailadd.data.ID);
      }
      //Task delegation getting user id from hubsite
      const user = await this._Service.getByhubEmail(this.state.approverEmail);
      if (user) {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });

        //Task delegation 
        const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubUrl, this.props.taskDelegationSettings, user.Id);
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let toDate = moment(taskDelegation[0].ToDate).toDate();
          let fromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
          fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
          if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
            this.taskDelegate = "Yes";
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
            //detail list adding an item for approval
            const DelegatedTo = await this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail);
            if (DelegatedTo) {
              this.setState({
                delegateToIdInSubSite: DelegatedTo.Id,
              });
              const DelegatedFor = await this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail);
              if (DelegatedFor) {
                this.setState({
                  delegateForIdInSubSite: DelegatedFor.Id,
                });
                let detaildata16 = {
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Approval",
                  Title: this.state.documentName,
                  ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                  DueDate: this.state.dueDate,
                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                  ResponseStatus: "Under Approval",
                  SourceDocument: {
                    Description: this.state.documentName,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  OwnerId: this.state.ownerId
                }
                const detailsAdd = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata16);
                if (detailsAdd) {
                  this.setState({ detailIdForApprover: detailsAdd.data.ID });
                  this.newDetailItemID = detailsAdd.data.ID;
                  let detaildata17 = {
                    Link: {
                      Description: this.state.documentName + "-Approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
                    }
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata17, detailsAdd.data.ID);
                  let approverdata2 = {
                    ApproverId: this.state.delegateToIdInSubSite,
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, approverdata2, this.documentIndexID);
                  await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, approverdata2, this.sourceDocumentID);
                  await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, approverdata2, this.newheaderid);
                  //MY tasks list updation
                  let taskdata16 = {
                    Title: "Approve '" + this.state.documentName + "'",
                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: (this.state.delegatedToId),
                    Workflow: "Approval",
                    // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                    DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                    Link: {
                      Description: this.state.documentName + "-Approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detailsAdd.data.ID + ""
                    },

                  }
                  const task = await this._Service.createNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata16);
                  if (task) {
                    let taskdata17 = {
                      TaskID: task.data.ID,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata17, detailsAdd.data.ID);
                    //notification preference checking                                 
                    await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
                    // await this._adaptiveCard("Approval",this.state.approverEmail,this.state.approverName,"General",task.data.ID);

                  }//taskID
                }//r

              }//DelegatedFor
            }//DelegatedTo
          }//duedate checking
          else {
            let detaildata18 = {
              HeaderIDId: Number(this.newheaderid),
              Workflow: "Approval",
              Title: this.state.documentName,
              ResponsibleId: this.state.approver,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Approval",
              SourceDocument: {
                Description: this.state.documentName,
                Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
              },
              OwnerId: this.state.ownerId,
            }
            const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata18);
            if (detail) {
              this.setState({ detailIdForApprover: detail.data.ID });
              this.newDetailItemID = detail.data.ID;
              let detaildata19 = {
                Link: {
                  Description: this.state.documentName + "-Approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                }
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata19, detail.data.ID);
              let approverdata3 = {
                ApproverId: this.state.approver
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, approverdata3, this.documentIndexID);
              await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, approverdata3, this.sourceDocumentID);
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, approverdata3, this.newheaderid);
              //MY tasks list updation
              let taskdata20 = {
                Title: "Approve '" + this.state.documentName + "'",
                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                DueDate: this.state.dueDate,
                StartDate: this.today,
                AssignedToId: user.Id,
                Workflow: "Approval",
                Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                Source: (this.props.project ? "Project" : "QDMS"),
                Link: {
                  Description: this.state.documentName + "-Approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                },
              }
              const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata20);
              if (task) {
                let taskdata19 = {
                  TaskID: task.data.ID,
                }
                await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata19, detail.data.ID);
                //notification preference checking                                 
                await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
                // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General", task.data.ID);


              }//taskID
            }//r
          }

        }
        else {
          let detaildata20 = {
            HeaderIDId: Number(this.newheaderid),
            Workflow: "Approval",
            Title: this.state.documentName,
            ResponsibleId: this.state.approver,
            DueDate: this.state.dueDate,
            ResponseStatus: "Under Approval",
            SourceDocument: {
              Description: this.state.documentName,
              Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            },
            OwnerId: this.state.ownerId
          }
          const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata20);
          if (detail) {
            this.setState({ detailIdForApprover: detail.data.ID });
            this.newDetailItemID = detail.data.ID;
            let detaildata21 = {
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              }
            }
            await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata21, detail.data.ID);
            let approverdata4 = {
              ApproverId: this.state.approver
            }
            await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, approverdata4, this.documentIndexID);
            await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, approverdata4, this.sourceDocumentID);
            await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, approverdata4, this.newheaderid);

            //MY tasks list updation
            let taskdata15 = {
              Title: "Approve '" + this.state.documentName + "'",
              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
              DueDate: this.state.dueDate,
              StartDate: this.today,
              AssignedToId: user.Id,
              Workflow: "Approval",
              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
              Source: (this.props.project ? "Project" : "QDMS"),
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              }
            }
            const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata15);
            if (task) {
              this.TaskID = task.data.ID
              let taskdata18 = {
                TaskID: task.data.ID,
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdata18, detail.data.ID);
              //notification preference checking                                 
              await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
              // await this._adaptiveCard("Approval",this.state.approverEmail,this.state.approverName,"General",task.data.ID);

            }//taskID
          }//r
        }//else no delegation
      }
      let indexdata = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: { results: this.state.currentUserReviewer },
        WorkflowDueDate: this.state.dueDate
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, this.documentIndexID);
      let sourcedata = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: { results: this.state.currentUserReviewer },
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourcedata, this.sourceDocumentID);
      let logdata = {
        Title: this.state.documentID,
        Status: "Under Approval",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Approval",
        DueDate: this.state.dueDate
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata);
      await this._triggerPermission(this.sourceDocumentID);
      this.setState({
        comments: "",
        statusKey: "",
        approverEmail: "",
        approverName: "",
        approver: "",
      });
      if (this.taskDelegate == "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateUnderApproval, messageType: 4 },
        });
      }
      else {
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 },
        });
      }

      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
      //msg
    }//newheaderid
  }
  //  project request to approve
  public async _underProjectApprove(previousHeaderItem: any) {
    this._LAUrlGetting();
    // this._LaUrlGettingAdaptive();
    let headerdata = {
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.currentUserReviewer },
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString(),
      RevisionCodingId: this.state.revisionCoding,
      TransmittalRevision: this.state.transmittalRevision,
      AcceptanceCodeId: this.state.acceptanceCodeId,
      ApproveInSameRevision: this.state.sameRevision
    }
    const header = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata);
    if (header) {
      this.newheaderid = header.data.ID;
      let logdata = {
        Title: this.state.documentID,
        Status: "Workflow Initiated",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        DueDate: this.state.dueDate,
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata);
      let detaildata4 = {
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
        SourceDocument: {
          Description: this.state.documentName,
          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
        },
        OwnerId: this.state.ownerId,
      }
      const detail = this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata4)
        .then(async (detail: any) => {
          this.setState({ detailIdForApprover: detail.data.ID });
          this.newDetailItemID = detail.data.ID;
          let detaildata5 = {
            Link: {
              Description: this.state.documentName + "-Review",
              Url: this.props.siteUrl + "/SitePages/" + this.props.documentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
            }
          }
          await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata5, detail.data.ID);
        })
      //Task delegation getting user id from hubsite
      const user = await this._Service.getByhubEmail(this.state.approverEmail);
      if (user) {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });

        //Task delegation 
        const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubUrl, this.props.taskDelegationSettings, user.Id);
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let toDate = moment(taskDelegation[0].ToDate).toDate();
          let fromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
          fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
          if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
            this.taskDelegate = "Yes";
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
            //detail list adding an item for approval
            const DelegatedTo = await this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail);
            if (DelegatedTo) {
              this.setState({
                delegateToIdInSubSite: DelegatedTo.Id,
              });
              const DelegatedFor = await this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail);
              if (DelegatedFor) {
                this.setState({
                  delegateForIdInSubSite: DelegatedFor.Id,
                });
                let detaildata = {
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Approval",
                  Title: this.state.documentName,
                  ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                  DueDate: this.state.dueDate,
                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                  ResponseStatus: "Under Approval",
                  SourceDocument: {
                    Description: this.state.documentName,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  OwnerId: this.state.ownerId,
                }
                const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata);
                if (detail) {
                  this.setState({ detailIdForApprover: detail.data.ID });
                  this.newDetailItemID = detail.data.ID;
                  let detaildata6 = {
                    Link: {
                      Description: this.state.documentName + "-Approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                    },
                  }
                  this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata6, detail.data.ID);
                  let delegateApprover = {
                    ApproverId: this.state.delegateToIdInSubSite
                  }
                  await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, delegateApprover, this.documentIndexID);
                  await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, delegateApprover, this.sourceDocumentID);
                  await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, delegateApprover, this.newheaderid);
                  //MY tasks list updation
                  let taskdata6 = {
                    Title: "Approve '" + this.state.documentName + "'",
                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                    Workflow: "Approval",
                    // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                    DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                    Link: {
                      Description: this.state.documentName + "-Approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                    },
                  }
                  const task = await this._Service.createNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata6)
                  if (task) {
                    this.TaskID = task.data.ID;
                    let taskdetail1 = {
                      TaskID: task.data.ID,
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, taskdetail1, detail.data.ID);
                    //notification preference checking                                 
                    await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
                    // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "Project", task.data.ID);
                  }//taskID
                }//r

              }//DelegatedFor
            }//DelegatedTo
          }//duedate checking
          else {
            let detaildata7 = {
              HeaderIDId: Number(this.newheaderid),
              Workflow: "Approval",
              Title: this.state.documentName,
              ResponsibleId: this.state.approver,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Approval",
              SourceDocument: {
                Description: this.state.documentName,
                Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
              },
              OwnerId: this.state.ownerId,
            }
            const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata7);
            if (detail) {
              this.setState({ detailIdForApprover: detail.data.ID });
              this.newDetailItemID = detail.data.ID;
              let detaildata8 = {
                Link: {
                  Description: this.state.documentName + "-Approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                }
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata8, detail.data.ID);
              let approverdata = {
                ApproverId: this.state.approver
              }
              await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, approverdata, this.documentIndexID);
              await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, approverdata, this.sourceDocumentID);
              await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, approverdata, this.newheaderid);
              //MY tasks list updation
              let taskdata7 = {
                Title: "Approve '" + this.state.documentName + "'",
                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                DueDate: this.state.dueDate,
                StartDate: this.today,
                AssignedToId: user.Id,
                Workflow: "Approval",
                Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                Source: (this.props.project ? "Project" : "QDMS"),
                Link: {
                  Description: this.state.documentName + "-Approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
                },
              }
              const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata7);
              if (task) {
                this.TaskID = task.data.ID;
                let updatedetail2 = {
                  TaskID: task.data.ID,
                }
                this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetail2, detail.data.ID);
                //notification preference checking                                 
                await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
                // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "Project", task.data.ID);
              }//taskID
            }//r
          }

        }
        else {
          let detaildata9 = {
            HeaderIDId: Number(this.newheaderid),
            Workflow: "Approval",
            Title: this.state.documentName,
            ResponsibleId: this.state.approver,
            DueDate: this.state.dueDate,
            ResponseStatus: "Under Approval",
            SourceDocument: {
              Description: this.state.documentName,
              Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
            },
            OwnerId: this.state.ownerId,
          }
          const detail = await this._Service.createNewItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata9);
          if (detail) {
            this.setState({ detailIdForApprover: detail.data.ID });
            this.newDetailItemID = detail.data.ID;
            let detaildata10 = {
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              }
            }
            await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata10, detail.data.ID)
            let approverdata1 = {
              ApproverId: this.state.approver,
            }
            await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, approverdata1, this.documentIndexID);
            await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, approverdata1, this.sourceDocumentID);
            await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, approverdata1, this.newheaderid);
            //MY tasks list updation
            let taskdata8 = {
              Title: "Approve '" + this.state.documentName + "'",
              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
              DueDate: this.state.dueDate,
              StartDate: this.today,
              AssignedToId: user.Id,
              Workflow: "Approval",
              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
              Source: (this.props.project ? "Project" : "QDMS"),
              Link: {
                Description: this.state.documentName + "-Approve",
                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + detail.data.ID + ""
              },

            }
            const task = await this._Service.createhubNewItem(this.props.hubUrl, this.props.workflowTasksList, taskdata8);
            if (task) {
              this.TaskID = task.data.ID;
              let updatedetail3 = {
                TaskID: task.data.ID,
              }
              this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, updatedetail3, detail.data.ID);
              //notification preference checking                                 
              await this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName);
              // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "Project", task.data.ID);
            }//taskID
          }//r
        }//else no delegation
      }
      let updateindex1 = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        WorkflowDueDate: this.state.dueDate,
        ReviewersId: { results: this.state.currentUserReviewer },
      }
      this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, updateindex1, this.documentIndexID);
      let updatesource1 = {
        WorkflowStatus: "Under Approval",
        Workflow: "Approval",
        ReviewersId: { results: this.state.currentUserReviewer },
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, updatesource1, this.sourceDocumentID);
      let logitem1 = {
        Title: this.state.documentID,
        Status: "Under Approval",
        LogDate: this.today,
        WorkflowID: this.newheaderid,
        Revision: this.state.revision,
        DocumentIndexId: this.documentIndexID,
        Workflow: "Approval",
        DueDate: this.state.dueDate,
      }
      await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem1);
      await this._triggerPermission(this.sourceDocumentID);
      this.setState({
        comments: "",
        statusKey: "",
        approverEmail: "",
        approverName: "",
        approver: "",
      });
      if (this.taskDelegate == "Yes") {
        this.setState({
          hideLoading: true,
          statusMessage: { isShowMessage: true, message: this.taskDelegateUnderApproval, messageType: 4 },
        });
      }
      else {
        this.setState({
          hideLoading: true,
          saveDisable: "none",
          statusMessage: { isShowMessage: true, message: this.underApproval, messageType: 4 },
        });
      }

      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);

    }//newheaderid
  }

  //Adaptive Card
  private _LaUrlGettingAdaptive = async () => {
    const laUrl: any[] = await this._Service.getListItems(this.props.hubUrl, this.props.requestList);
    console.log("Posturl" + laUrl);
    for (let i = 0; i < laUrl.length; i++) {
      if (laUrl[i].Title == "Adaptive _Card") {
        this.postUrlForAdaptive = laUrl[i].PostUrl;
      }
    }
  }
  // on cancel
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });


  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      statusKey: "",
      comments: "",
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
    //this.validator.hideMessages();
    // window.location.replace(this.RedirectUrl);
  }
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',
    //subText: '<b>Do you want to cancel? </b> ',
  };
  private modalProps = {
    isBlocking: true,
  };
  // on format date
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  };
  public render(): React.ReactElement<ITransmittalSendRequestProps> {
    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    return (
      <section className={`${styles.transmittalSendRequest}`}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div style={{ display: this.state.access }}>
          <div className={styles.border}>
            <div className={styles.alignCenter}> {this.props.webpartHeader}</div>
            <br></br>

            <div className={styles.flex}>
              <div className={styles.width}><Label >Document ID : {this.state.documentID}</Label></div>
              <div ><Link onClick={this._openRevisionHistory} target="_blank" underline>Revision History</Link></div>
            </div>
            <div hidden={this.state.hideProject}>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Project Name : {this.state.projectName} </Label></div>
                <div><Label >Project Number : {this.state.projectNumber}</Label></div>
              </div>
            </div>
            <div className={styles.flex}>
              <div className={styles.width}><Label >Document : <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></Label></div>
              <div ><Label >Revision : {this.state.revision}</Label></div>
            </div>
            <div className={styles.flex}>
              <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
              <div><Label >Requester : {this.state.currentUser}</Label></div>
            </div>
            <div hidden={this.state.hideProject}>
              <div>
                <TooltipHost
                  content="Check if the document need to approve in same revision"
                  //id={tooltipId}
                  calloutProps={calloutProps}
                  styles={hostStyles}>
                  <Checkbox label="Approve in same revision ? " boxSide="end"
                    onChange={this._onSameRevisionChecked}
                    checked={this.state.sameRevision} />
                </TooltipHost>
              </div>

              {/* <div className={styles.width} style={{ paddingRight: '15px' }}>
                  <Dropdown
                    placeholder="Select Option"
                    label="Revision Level"
                    style={{ marginBottom: '10px', backgroundColor: "white", height: '34px' }}
                    options={this.state.revisionLevelArray}
                    onChanged={this._revisionLevelChanged}
                    selectedKey={this.state.revisionLevelvalue}
                    required />
                  <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionLevel", this.state.revisionLevelvalue, "required")}{" "}</div>
                </div> */}
              <div>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Document Controller"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._dccReviewerChange(items)}
                  defaultSelectedUsers={[this.state.dccReviewerName]}
                  showHiddenInUI={false}
                  // isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("DocumentController", this.state.dccReviewer, "required")}{" "}</div>
              </div>


            </div>
            <div><PeoplePicker
              context={this.props.context as any}
              titleText="Reviewer(s)"
              personSelectionLimit={20}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              disabled={false}
              ensureUser={true}
              onChange={(items) => this._reviewerChange(items)}
              defaultSelectedUsers={this.state.reviewersName}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            /></div>

            <div className={styles.flex}>
              <div className={styles.width}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Approver *"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._approverChange(items)}
                  defaultSelectedUsers={[this.state.approverName]}
                  showHiddenInUI={false}
                  required={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
                <div style={{ display: this.state.validApprover, color: "#dc3545" }}>Not able to change approver</div>
                <div style={{ color: "#dc3545" }}>{this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
              </div>
              <div className={styles.width} style={{ paddingLeft: '15px', marginTop: '2px' }}>
                <DatePicker label="Due Date:" id="DueDate"
                  onSelectDate={this._onExpDatePickerChange}
                  placeholder="Select a date..."
                  isRequired={true}
                  value={this.state.dueDate}
                  minDate={new Date()}
                  formatDate={this._onFormatDate}
                // className={controlClass.control}
                // onSelectDate={this._onDatePickerChange}                 
                /><div style={{ color: "#dc3545" }}>{this.validator.message("DueDate", this.state.dueDate, "required")}{" "}</div>
              </div>
            </div>

            <div className={styles.mt}>
              < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
            <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.mt}>
              <div hidden={this.state.hideLoading}><Spinner label={'Document is Sending...'} /></div>
            </div>
            <DialogFooter>

              <div className={styles.rgtalign}>
                <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
              </div>
              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._submitSendRequest} style={{ display: this.state.saveDisable }}>Submit</PrimaryButton >
                <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
              </div>
            </DialogFooter>

            {/* {/ Cancel Dialog Box /} */}
            <div style={{ display: this.state.cancelConfirmMsg }}>
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
            </div>
          </div>
        </div>
        <div style={{ display: this.state.accessDeniedMsgBar }}>

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
