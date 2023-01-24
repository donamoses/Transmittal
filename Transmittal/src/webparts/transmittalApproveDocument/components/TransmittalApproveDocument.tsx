import * as React from 'react';
import styles from './TransmittalApproveDocument.module.scss';
import { ITransmittalApproveDocumentProps, ITransmittalApproveDocumentState } from './ITransmittalApproveDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import SimpleReactValidator from 'simple-react-validator';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IconButton, IDropdownOption, IIconProps, Label, Link, MessageBar, PrimaryButton, ProgressIndicator, Spinner, TextField } from '@fluentui/react';
import { Accordion, AccordionItem, AccordionItemButton, AccordionItemHeading, AccordionItemPanel } from 'react-accessible-accordion';
import * as moment from 'moment';
import { BaseService } from '../services';
import { IHttpClientOptions, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as _ from 'lodash';
export default class TransmittalApproveDocument extends React.Component<ITransmittalApproveDocumentProps, ITransmittalApproveDocumentState, {}> {
  private validator: SimpleReactValidator;
  private _Service: BaseService;
  private workflowHeaderID: any;
  private documentIndexID: any;
  private sourceDocumentID: any;
  private workflowDetailID: any;
  private currentEmail: string;
  // private reqWeb;
  private documentApprovedSuccess: string;
  private documentSavedAsDraft: string;
  private documentRejectSuccess: string;
  private documentReturnSuccess: string;
  private today = new Date();
  private revisionLogId: any;
  private currentrevision: any;
  private invalidApprovalLink: string;
  private invalidUser: string;
  // private redirectUrlSuccess;
  private valid: string;
  private approverEmail: string;
  private departmentExists: string;
  private postUrl: string;
  // private siteUrl;
  // private permissionpostUrl;
  public constructor(props: ITransmittalApproveDocumentProps) {
    super(props);
    this.state = {
      publishOptionKey: "",
      requester: "",
      linkToDoc: "",
      requesterComments: "",
      dueDate: "",
      dccComments: "",
      dcc: null,
      dccEmail: "",
      dccName: "",
      hideProject: true,
      publishOption: "",
      status: "",
      statusKey: "",
      approveDocument: 'none',
      hideLoading: true,
      documentID: "",
      documentName: "",
      revision: "",
      ownerName: "",
      ownerEmail: "",
      requesterName: "",
      requesterEmail: "",
      requestedDate: "",
      requesterComment: "",
      reviewerData: [],
      access: "none",
      accessDeniedMsgBar: "none",
      hidepublish: true,
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      comments: "",
      criticalDocument: "",
      approverName: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      savedisable: "",
      taskID: "",
      dccreviewerData: [],
      revisionLevel: "",
      acceptanceCodearray: [],
      acceptanceCode: "",
      hideacceptance: true,
      externalDocument: "",
      hidetransmittalrevision: true,
      transmittalRevision: "",
      publishcheck: "",
      projectName: "",
      projectNumber: "",
      currentRevision: "",
      previousRevisionItemID: null,
      revisionItemID: "",
      newRevision: "",
      sameRevision: "",
      hideButton: false,
      reviewersTableDiv: "none",
      isdocx: "none",
      nodocx: "",
      loaderDisplay: "",
      dccTableDiv: "none",
      commentValid: "none",
      commentRequired: false
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    this._accessGroups = this._accessGroups.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this._bindApprovalForm = this._bindApprovalForm.bind(this);
    this._project = this._project.bind(this);
    this._drpdwnPublishFormat = this._drpdwnPublishFormat.bind(this);
    this._status = this._status.bind(this);
    this._commentsChange = this._commentsChange.bind(this);
    this._saveAsDraft = this._saveAsDraft.bind(this);
    this._docSave = this._docSave.bind(this);
    this._publish = this._publish.bind(this);
    this._returnDoc = this._returnDoc.bind(this);
    this._sendMail = this._sendMail.bind(this);
    this._onCancel = this._onCancel.bind(this);
    this._acceptanceChanged = this._acceptanceChanged.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._publishUpdate = this._publishUpdate.bind(this);
    this._generateNewRevision = this._generateNewRevision.bind(this);
    this._checkCurrentUser = this._checkCurrentUser.bind(this);
    this._LAUrlGetting = this._LAUrlGetting.bind(this);

  }
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });

  }
  //Page Load
  public async componentDidMount() {
    this.setState({ loaderDisplay: "" });
    // Get User Messages
    await this._userMessageSettings();
    //Get Current User
    this._Service.getCurrentUser()
      .then(async (user: any) => {
        let userEmail = user.Email;
        this.currentEmail = userEmail;
      });
    //Get Today
    this.today = new Date();
    //Get Parameter from URL
    await this._queryParamGetting();
    //Get Approver
    const headerItem: any = await this._Service.getheaderdata(this.props.siteUrl, this.props.workflowHeaderList, this.workflowHeaderID);
    this.approverEmail = headerItem.Approver.EMail;
    this.documentIndexID = headerItem.DocumentIndexID;

    if (this.valid === "ok") {
      //Get Access
      if (this.props.project) {
        await this._checkCurrentUser();

      }
      else {
        await this._accessGroups();

      }

    }
    this._LAUrlGetting();
  }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.gethubUserMessageListItems(this.props.hubUrl, this.props.userMessageSettings);
    console.log(userMessageSettings);
    for (let i in userMessageSettings) {
      if (userMessageSettings[i].Title === "ApproveSubmitSuccess") {
        let successmsg = userMessageSettings[i].Message;
        this.documentApprovedSuccess = replaceString(successmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title === "ApproveDraftSuccess") {
        this.documentSavedAsDraft = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title === "InvalidApprovalLink") {
        this.invalidApprovalLink = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title === "InvalidApproverUser") {
        this.invalidUser = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title === "ApproveRejectSuccess") {
        let rejectmsg = userMessageSettings[i].Message;
        this.documentRejectSuccess = replaceString(rejectmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title === "ApproveReturnSuccess") {
        let returnmsg = userMessageSettings[i].Message;
        this.documentReturnSuccess = replaceString(returnmsg, '[DocumentName]', this.state.documentName);
      }
    }

  }
  //Get Parameter from URL
  private async _queryParamGetting() {
    await this._userMessageSettings();
    //Query getting...
    let params = new URLSearchParams(window.location.search);

    let headerid = params.get('hid');
    let detailid = params.get('dtlid');
    if (headerid !== "" && headerid !== null && detailid !== "" && detailid !== null) {
      this.workflowHeaderID = parseInt(headerid);
      this.workflowDetailID = parseInt(detailid);
      this.valid = "ok";

    }
    else {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidApprovalLink, messageType: 1 },
      });
      // this.setState({ accessDeniedMsgBar: 'none', });
      setTimeout(() => {
        window.location.replace(this.props.siteUrl);
      }, 10000);
    }
  }
  //Check Current User is approver
  public async _checkCurrentUser() {
    if (this.currentEmail === this.approverEmail) {
      this.setState({ access: "", accessDeniedMsgBar: "none", loaderDisplay: "none" });
      if (this.props.project) {
        this.setState({ hideProject: false });
        await this._project();
      }
      await this._bindApprovalForm();
    }
    else {
      this.setState({
        loaderDisplay: "none",
        access: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 1 }
      });
    }
  }
  // Check Access
  private async _accessGroups() {
    let AccessGroup = [];
    let ok = "No";
    if (this.props.project) {
      AccessGroup = await this._Service.getprojectaccessgroup(this.props.hubUrl, this.props.PermissionMatrixSettings);
    }
    else {
      AccessGroup = await this._Service.getqdmsaccessgroup(this.props.hubUrl, this.props.PermissionMatrixSettings);
    }

    let AccessGroupItems: any[] = AccessGroup[0].AccessGroups.split(',');
    console.log("AccessGroupItems", AccessGroupItems);
    const DocumentIndexItem: any = await this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
    console.log("DocumentIndexItem", DocumentIndexItem);
    //cheching if department selected
    if (DocumentIndexItem.DepartmentID !== null) {
      this.departmentExists = "Exists";
      let deptid = parseInt(DocumentIndexItem.DepartmentID);
      const departmentItem: any = await this._Service.gethubItemById(this.props.hubUrl, this.props.departmentList, deptid);
      //let AG = DepartmentItem[0].AccessGroups;
      console.log("departmentItem", departmentItem);
      let accessGroupvar = departmentItem.AccessGroups;
      const accessGroupItem: any = await this._Service.gethubListItems(this.props.hubUrl, this.props.accessGroupDetailsList);
      let accessGroupID;
      console.log(accessGroupItem.length);
      for (let a = 0; a < accessGroupItem.length; a++) {
        if (accessGroupItem[a].Title === accessGroupvar) {
          accessGroupID = accessGroupItem[a].GroupID;
          this.GetGroupMembers(this.props.context, accessGroupID);
        }
      }
    }
    //if no department
    else {
      //alert("with bussinessUnit");
      if (DocumentIndexItem.BusinessUnitID !== null) {
        this.departmentExists === "Exists";
        let bussinessUnitID = parseInt(DocumentIndexItem.BusinessUnitID);
        const bussinessUnitItem: any = await this._Service.gethubItemById(this.props.hubUrl, this.props.businessUnit, bussinessUnitID);
        console.log("departmentItem", bussinessUnitItem);
        let accessGroupvar = bussinessUnitItem.AccessGroups;
        // alert(accessGroupvar);
        const accessGroupItem: any = await this._Service.gethubListItems(this.props.hubUrl, this.props.accessGroupDetailsList);
        let accessGroupID;
        console.log(accessGroupItem.length);
        for (let a = 0; a < accessGroupItem.length; a++) {
          if (accessGroupItem[a].Title === accessGroupvar) {
            accessGroupID = accessGroupItem[a].GroupID;
            this.GetGroupMembers(this.props.context, accessGroupID);
          }
        }
      }
    }
  }
  // LA url getting
  private _LAUrlGetting = async () => {
    const laUrl = await this._Service.gettriggerUnderApprovalPermission(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
  }
  // Getting group members
  public async GetGroupMembers(context: WebPartContext, groupId: string): Promise<any[]> {
    let users: string[] = [];
    try {
      let client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
      let response = await client
        .api(`/groups/${groupId}/members`)
        .version('v1.0')
        .select(['mail', 'displayName'])
        .get();
      response.value.map((item: any) => {
        users.push(item);
      });
    } catch (error) {
      console.log('MSGraphService.GetGroupMembers Error: ', error);
    }
    console.log('MSGraphService.GetGroupMembers: ', users, "GroupID:", groupId);
    //checking current users 
    if (users.length > 0) {
      this._checkingCurrent(users);
    }
    else if (this.departmentExists === "Exists") {
      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.props.siteUrl);
      }, 10000);
    }

    return;
  }
  // Checking current user email
  private async _checkingCurrent(userEmail: any) {
    for (let k in userEmail) {
      if (this.currentEmail === userEmail[k].mail) {
        this.setState({ access: "none", accessDeniedMsgBar: "none" });
        this.valid = "Yes";
        await this._checkCurrentUser();

        break;
      }
    }
    if (this.valid !== "Yes") {

      this.setState({
        loaderDisplay: "none",
        accessDeniedMsgBar: "", access: "none",
        statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.props.siteUrl);
      }, 10000);
    }
  }
  //Bind Approval Form
  public async _bindApprovalForm() {
    let approverId;
    let approverName;
    let requesterName;
    let requesterEmail;
    let requestedDate;
    let requesterComment;
    let dueDate;
    let documentID;
    let documentName;
    let ownerName;
    let ownerEmail;
    let revision;
    let linkToDocument;
    let approverComment;
    let reviewerArr: any[] = [];
    let reviewDate;
    let criticalDocument;
    let taskID;
    let status;
    let publishOption;
    let type;
    //Get Header Item
    const headerItem: any = await this._Service.getheaderbinddata(this.props.siteUrl, this.props.workflowHeaderList, this.workflowHeaderID);
    approverId = headerItem.Approver.ID;
    approverName = headerItem.Approver.Title;

    this.documentIndexID = headerItem.DocumentIndexID;
    requesterName = headerItem.Requester.Title;
    requesterEmail = headerItem.Requester.EMail;
    if (headerItem.RequestedDate !== null) {
      let reqdate = new Date(headerItem.RequestedDate);
      requestedDate = moment(reqdate).format('DD-MM-YYYY HH:mm');
    }
    requesterComment = headerItem.RequesterComment;
    let duedate = new Date(headerItem.DueDate);
    dueDate = moment(duedate).format('DD-MM-YYYY');
    publishOption = headerItem.PublishFormat;
    //Get Document Index
    const documentIndexItem: any = await this._Service.getindexbinddata(this.props.siteUrl, this.props.documentIndexList, this.documentIndexID);
    console.log(documentIndexItem);
    documentID = documentIndexItem.DocumentID;
    documentName = documentIndexItem.DocumentName;
    ownerName = documentIndexItem.Owner.Title;
    ownerEmail = documentIndexItem.Owner.EMail;
    revision = documentIndexItem.Revision;
    linkToDocument = documentIndexItem.SourceDocument.Url;
    criticalDocument = documentIndexItem.CriticalDocument;
    this.sourceDocumentID = documentIndexItem.SourceDocumentID;
    //Get Workflow Details
    const detailItem: any[] = await this._Service.getdetailbinddata(this.props.siteUrl, this.props.workflowDetailsList, this.workflowHeaderID);
    for (let k in detailItem) {
      if (detailItem[k].Workflow === 'Review') {
        let rewdate = new Date(detailItem[k].ResponseDate);
        reviewDate = moment(rewdate).format('DD-MMM-YYYY HH:mm');
        reviewerArr.push({
          ResponseDate: reviewDate,
          Reviewer: detailItem[k].Responsible.Title,
          ResponsibleComment: detailItem[k].ResponsibleComment
        });
      }
      else if (detailItem[k].Workflow === 'Approval') {
        approverComment = detailItem[k].ResponsibleComment;
        taskID = detailItem[k].TaskID;
        if (detailItem[k].ResponseStatus === "Published") {
          status = "Approved";
          this.setState({
            hidepublish: false,
            savedisable: "none",
            hideButton: true,
            statusKey: status,
          });

        }
        else {
          status = detailItem[k].ResponseStatus;
        }
        if (detailItem[k].ResponseStatus !== "Under Approval") {
          this.setState({ savedisable: "none", hideButton: true });
        }
        if (detailItem[k].ResponseStatus === "Under Approval") {
          this.setState({ statusKey: "" });
        }
      }

    }
    if (reviewerArr.length > 0) {
      this.setState({
        reviewersTableDiv: ""
      });
    }
    else {
      this.setState({
        reviewersTableDiv: "none",
      });
    }
    let split = documentName.split(".", 2);
    type = split[1];
    if (type === "docx") {
      this.setState({ isdocx: "", nodocx: "none" });
    }
    else {
      this.setState({ isdocx: "none", nodocx: "" });
    }
    this.setState({
      documentID: documentID,
      documentName: documentName,
      linkToDoc: linkToDocument,
      revision: revision,
      ownerName: ownerName,
      ownerEmail: ownerEmail,
      dueDate: dueDate,
      requesterName: requesterName,
      requesterEmail: requesterEmail,
      requestedDate: requestedDate,
      requesterComment: requesterComment,
      reviewerData: reviewerArr,
      comments: approverComment,
      criticalDocument: criticalDocument,
      approverName: approverName,
      taskID: taskID,
      //statusKey: status,
      publishOptionKey: publishOption

    });
    await this._userMessageSettings();
  }
  //Bind datas for project
  public async _project() {
    let reviewDate: any;
    let dccReviewerArr: any[] = [];
    let acceptanceArray: any[] = [];
    let sorted_Acceptance: any[] = [];
    let projectName: any;
    let projectNumber: any;

    const headerItem: any = await this._Service.getprojectheaderbinddata(this.props.siteUrl, this.props.workflowHeaderList, this.workflowHeaderID)
    let dcc = headerItem.DocumentController.ID;
    let dccName = headerItem.DocumentController.Title;
    let dccEmail = headerItem.DocumentController.EMail;
    let documentIndexId = headerItem.DocumentIndexID;
    let acceptanceCode = headerItem.AcceptanceCodeId;
    let RevisionCodingId = headerItem.RevisionCodingId;
    let ApproveInSameRevision = headerItem.ApproveInSameRevision;
    const documentIndexItem: any = await this._Service.getprojectindexbinddata(this.props.siteUrl, this.props.documentIndexList, documentIndexId)
    let externalDocument = documentIndexItem.ExternalDocument;
    let transmittalDocument = documentIndexItem.TransmittalDocument;
    let transmittalRevision = documentIndexItem.TransmittalRevision;

    const projectInformation = await this._Service.getListItems(this.props.siteUrl, this.props.projectInformationListName);
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (let k in projectInformation) {
        if (projectInformation[k].Key === "ProjectName") {
          this.setState({
            projectName: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key === "ProjectNumber") {
          this.setState({
            projectNumber: projectInformation[k].Title,
          });
        }
      }
    }
    if (dcc !== null) {
      const detailItem: any[] = await this._Service.getprojectdetailbinddata(this.props.siteUrl, this.props.workflowDetailsList, this.workflowHeaderID)
      for (let l in detailItem) {
        if (detailItem[l].Workflow === 'DCC Review') {
          let rewdate = new Date(detailItem[l].ResponseDate);
          reviewDate = moment(rewdate).format('DD-MM-YYYY HH:mm');
          dccReviewerArr.push({
            ResponseDate: reviewDate,
            Reviewer: detailItem[l].Responsible.Title,
            DCCResponsibleComment: detailItem[l].ResponsibleComment
          });
        }
      }
    }
    if (externalDocument === true) {
      this.setState({ hideacceptance: false });
      const transmittalcodeitems: any[] = await this._Service.getListItems(this.props.siteUrl, this.props.transmittalCodeSettingsList);

      for (let i = 0; i < transmittalcodeitems.length; i++) {
        if (transmittalcodeitems[i].AcceptanceCode === true) {
          let transmittalcodedata = {
            key: transmittalcodeitems[i].ID,
            text: transmittalcodeitems[i].Description
          };
          acceptanceArray.push(transmittalcodedata);
        }
      }
      console.log(acceptanceArray);
      sorted_Acceptance = _.orderBy(acceptanceArray, 'text', ['asc']);

    }
    if (transmittalDocument === true) {
      this.setState({ hidetransmittalrevision: false });
    }
    if (dccReviewerArr.length > 0) {
      this.setState({
        dccTableDiv: ""
      });
    }
    else {
      this.setState({
        dccTableDiv: "none",
      });
    }
    this.setState({
      dccreviewerData: dccReviewerArr,
      acceptanceCodearray: sorted_Acceptance,
      externalDocument: externalDocument,
      transmittalRevision: transmittalRevision,
      acceptanceCode: acceptanceCode,
      revisionItemID: RevisionCodingId,
      sameRevision: ApproveInSameRevision,
      dcc: dcc,
      dccName: dccName,
      dccEmail: dccEmail
    });
  }
  //Status Change
  public _status(option: { key: any; text: any }) {
    //console.log(option.key);
    if (option.key === 'Approved') {
      this.setState({ hidepublish: false, commentRequired: false, commentValid: "none" });
    }
    else {
      this.setState({ hidepublish: true, commentRequired: true, commentValid: "" });
    }
    this.setState({ statusKey: option.key, status: option.text });
  }
  //Publish Change
  public _drpdwnPublishFormat(option: { key: any; text: any }) {
    //console.log(option.key);
    this.setState({ publishOptionKey: option.key, publishOption: option.text });
  }
  public async _acceptanceChanged(option: { key: any; text: any }) {
    console.log(option.key);
    this.setState({ acceptanceCode: option.key });
  }
  //Comment Change
  public _commentsChange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, });
  }
  //Save as Draft
  public _saveAsDraft = async () => {
    const commentadd = {
      ResponsibleComment: this.state.comments
    }
    await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, commentadd, this.workflowDetailID);
    this.setState({
      statusMessage: { isShowMessage: true, message: this.documentSavedAsDraft, messageType: 4 }
    });
    setTimeout(() => {
      window.location.replace(this.props.siteUrl);
    }, 5000);
  }
  //Data Save
  private _docSave = async () => {
    await this._Service.getlogitem(this.props.siteUrl, this.props.documentRevisionLogList, this.workflowHeaderID).then(ifyes => {
      this.revisionLogId = ifyes[0].ID;
    });

    if (this.state.hidepublish === false) {
      if (this.validator.fieldValid("publish") && (this.state.statusKey !== "")) {
        this.validator.hideMessages();
        this.setState({ hideLoading: false, savedisable: "none" });
        const publishdata = {
          PublishFormat: this.state.publishOption
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, publishdata, this.workflowHeaderID);

        this._publish();
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();

      }
    }
    else {
      if ((this.state.statusKey !== "") && this.validator.fieldValid("comments")) {
        this.validator.hideMessages();
        const detaildata = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.today
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata, this.workflowDetailID)
        await this._returnDoc().then((afterReturn: any) => {
          this.setState({ approveDocument: "" });
          setTimeout(() => this.setState({ approveDocument: 'none', hideLoading: true }), 3000);
          this.setState({ savedisable: "none" });
        });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }

  }
  //Document Published
  protected async _publish() {
    if (this.props.project) {
      if (this.state.sameRevision === "Yes") {
        this.setState({ newRevision: this.state.revision });
      }
      else if (this.state.revisionItemID === null) {
        this._revisionCoding();
      }
      else {
        await this._generateNewRevision();
      }
    }
    else {
      this._revisionCoding();
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
      'SourceDocumentID': this.sourceDocumentID,
      'SiteURL': siteUrl,
      'DocumentName': this.state.documentName,
      'PublishedDate': this.today,
      'SourceDocumentLibrary': this.props.sourceDocumentLibrary,
      'PublishFormat': this.state.publishOption,
      'WorkflowStatus': "Published",
      'Revision': this.state.newRevision,
      'RevisionUrl': this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexID,
      'AcceptanceCode': parseInt(this.state.acceptanceCode)
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

      this._publishUpdate();
    }
    else {
    }
  }
  public _revisionCoding = async () => {
    let intrev;
    if (this.state.revision === "-") {
      intrev = 0;
    }
    else {
      intrev = this.state.revision;
    }
    if (this.props.project) {
      let revision = parseInt(intrev);
      let rev = revision + 1;
      this.currentrevision = rev.toString();
      this.setState({ newRevision: this.currentrevision });
    }
    else {
      let revision = parseInt(intrev);
      let rev = revision + 1;
      this.currentrevision = rev.toString();
      this.setState({ newRevision: this.currentrevision });
    }
  }
  public _generateNewRevision = async () => {
    let currentRevision = this.state.revision; // set the current revisionsettings ID in state variable.
    this.setState({
      previousRevisionItemID: this.state.revisionItemID // set this value with previous revision settings id from Project document index item.
    });

    // Reading current revision coding details from RevisionSettings.
    const revisionItem: any = await this._Service.getRevisionListItems(this.props.hubUrl, this.props.RevisionSettings, this.state.revisionItemID);
    console.log(revisionItem);
    let startPrefix = '-';
    let newRevision = '';
    let pattern = revisionItem.Pattern;
    let endWith = '0';
    let minN = revisionItem.MinN;
    let maxN = '0';
    let isAutoIncrement = revisionItem.AutoIncrement === 'TRUE';
    let firstChar = currentRevision.substring(0, 1);
    let currentNumber = currentRevision.substring(1, currentRevision.length);
    let startWith = revisionItem.StartWith;

    if (revisionItem.EndWith !== null)
      endWith = revisionItem.EndWith;

    if (revisionItem.MaxN !== null)
      maxN = revisionItem.MaxN;

    if (revisionItem.StartPrefix !== null)
      startPrefix = revisionItem.StartPrefix.toString();

    //splitting pattern values
    let incrementValue = 1;
    let isAlphaIncrement = pattern.split('+')[0] === 'A';
    let isNumericIncrement = pattern.split('+')[0] === 'N';
    if (pattern.split('+').length === 2) {
      incrementValue = Number(pattern.split('+')[1]);
    }
    //Resetting current revision as blank if current revisionsetting id is different.
    if (this.state.revisionItemID !== this.state.previousRevisionItemID) {
      currentRevision = '-';
    }
    try {
      //Getting first revision value.
      if (currentRevision === '-') {
        if (!isAutoIncrement) // Not an auto increment pattern, splitting the pattern with command reading the first value.
        {
          newRevision = pattern.split(',')[0];
        }
        else {
          if (startPrefix !== '-' && startPrefix.split(',').length > 0)  //Auto increment   with startPrefix eg. A1,A2, A3 etc., then handling both single and multple startPrefix
          {
            startPrefix = startPrefix.split(',')[0];
          }
          if (startWith !== null) // 
          {
            newRevision = startWith; //assigning startWith as newRevision for the first time.
          }
          else {
            newRevision = startPrefix + '' + minN;
          }
          if (startWith === null && startPrefix === '-') // Assigning minN if startWith and StartPrefix are null.
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
           if(i > 0 && String(currentRevision) === String(patternArray[i]))
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
            if (String(currentRevision) === patternArray[i] && (i + 1) < patternArray.length) {
              newRevision = patternArray[i + 1];
              break;
            }
          }
        }
      }
      else if (isAutoIncrement)// current revision is not blank and auto increment pattern .
      {
        if (startWith !== null && String(currentRevision) === String(startWith)) // Revision code with startWith  and startWith already set as Revision
        {
          if (startPrefix === '-') // second revision without startPrefix / '-' no StartPrefix
          {
            newRevision = minN;
          }
          else // 
          {
            newRevision = startPrefix + minN;
          }
        }
        // For all other cases
        else if (startPrefix !== '-') // Handling revisions with startPrefix here first char will be alpha
        {
          if (startPrefix.split(',').length === 1) // Single startPrefix eg. A1,A2,A3 etc with startPrefix 'A' and patter N+1
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
            if (maxN !== '0') {
              if (this.isNotANumber(currentRevision)) //MaxN set and not a number.
              {
                if (Number(currentNumber) < Number(maxN)) // alpha type revision
                {
                  newRevision = firstChar + (Number(currentNumber) + Number(incrementValue)).toString();
                }
                else if (Number(currentNumber) === Number(maxN)) {
                  // if current number part is same as maxN, get the next StartPrefix value from startPrefix.split(',')
                  let startPrefixArray = startPrefix.split(',');
                  for (let i = 0; i < startPrefixArray.length; i++) {
                    if (firstChar === startPrefixArray[i] && (i + 1) < startPrefixArray.length) {
                      firstChar = startPrefixArray[i + 1];
                      break;
                    }
                  }
                  if (firstChar === " ") // " " will denote a number
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
                else if (Number(currentRevision) === Number(maxN)) {
                  {
                    if (!this.isNotANumber(currentRevision)) // for setting a default value after the last item
                    {
                      firstChar = " ";
                    }
                    // if current number part is same as maxN, get the next StartPrefix value from startPrefix.split(',')
                    let startPrefixArray = startPrefix.split(',');
                    for (let i = 0; i < startPrefixArray.length; i++) {
                      if (firstChar === startPrefixArray[i] && (i + 1) < startPrefixArray.length) {
                        firstChar = startPrefixArray[i + 1];
                        break;
                      }
                    }
                    if (firstChar === " ") // Assigning number for blank array.
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
        if (newRevision === '' && startPrefix === '-' && endWith === '0') // No StartPrefix and No EndWith
        {
          if (isAlphaIncrement) // Alpha increment.
          {
            newRevision = this.nextChar(firstChar, incrementValue);
          }
          else {
            newRevision = (Number(currentRevision) + Number(incrementValue)).toString();
          }
        }
        else if (startPrefix === '-' && endWith !== '0') // No StartPrefix and with EndWith 
        {
          // cases A to E  then 0,1, 2,3 etc,
          if (currentRevision === endWith) {
            newRevision = minN;
          }
          else// if(currentRevision !== '0')
          {
            if (this.isNotANumber(currentRevision)) // Alpha increment.
            {
              newRevision = this.nextChar(firstChar, incrementValue);
            }
            else // (currentRevision === startWith && endWith !== null) // always alpha increment "X,,B"
            {
              newRevision = (Number(currentRevision) + Number(incrementValue)).toString();
            }
          }
        }
      }
      if (newRevision.indexOf('undefined') > -1 || newRevision === '') // Assigning with zero if array value exceeds.
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
  // Craeting next alpha char.
  private nextChar(currentChar: any, increment: any) {
    if (currentChar === 'Z')
      return 'A';
    else
      return String.fromCharCode(currentChar.charCodeAt(0) + increment);
  }
  // Check for number and alpha
  private isNotANumber(checkChar: any) {
    return isNaN(checkChar);
  }
  // Published Update
  public async _publishUpdate() {
    if (this.state.hideProject === true) {
      const indexdata = {
        PublishFormat: this.state.publishOption,
        WorkflowStatus: "Published",
        Revision: this.state.newRevision
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, this.documentIndexID);
      const headerdata = {
        ApprovedDate: this.today,
        WorkflowStatus: "Published",
        PublishFormat: this.state.publishOption,
        Revision: this.state.newRevision,
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata, this.workflowHeaderID);
      const detaildata = {
        ResponsibleComment: this.state.comments,
        ResponseStatus: "Published",
        ResponseDate: this.today
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata, this.workflowDetailID);

    }
    else {
      const indexdata1 = {
        PublishFormat: this.state.publishOption,
        WorkflowStatus: "Published",
        Revision: this.state.newRevision,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode),
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata1, this.documentIndexID);
      const headerdata1 = {
        ApprovedDate: this.today,
        WorkflowStatus: "Published",
        PublishFormat: this.state.publishOption,
        Revision: this.state.newRevision,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata1, this.workflowHeaderID);
      const detaildata1 = {
        ResponsibleComment: this.state.comments,
        ResponseStatus: "Published",
        ResponseDate: this.today,
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata1, this.workflowDetailID);
    }
    if (this.state.ownerEmail === this.state.requesterEmail) {
      this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
    }
    else {
      this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
      this._sendMail(this.state.requesterEmail, "DocPublish", this.state.requesterName);
    }
    if (this.props.project) {
      if (this.state.ownerEmail === this.state.dccEmail) { }
      else if (this.state.requesterEmail === this.state.dccEmail) { }
      else {
        this._sendMail(this.state.dccEmail, "DocPublish", this.state.dccName);
      }
    }
    const logdata = {
      Status: "Published",
      Workflow: "Approval"
    }
    let a = await this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata, this.revisionLogId);
    await this._Service.deletehubItemById(this.props.hubUrl, this.props.workflowTasksList, parseInt(this.state.taskID));
    this.setState({ approveDocument: "", savedisable: "none", hideLoading: true });
    setTimeout(() => this.setState({ approveDocument: 'none', }), 3000);
    this.setState({ hideLoading: true, statusMessage: { isShowMessage: true, message: this.documentApprovedSuccess, messageType: 4 } });
    setTimeout(() => {
      window.location.replace(this.props.siteUrl);
    }, 5000);

  }
  //Send Mail
  public _sendMail = async (emailuser: any, type: any, name: any) => {

    let formatday = moment(this.today).format('DD/MM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    let link;
    console.log(this.state.criticalDocument);

    const notificationPreference: any[] = await this._Service.getnotification(this.props.hubUrl, this.props.notificationPreference, emailuser);
    console.log(notificationPreference[0].Preference);
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference === "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference === "Send mail for critical document" && this.state.criticalDocument === true) {
        mailSend = "Yes";

      }
      else {
        mailSend = "No";
      }
    }
    else if (this.state.criticalDocument === true) {
      //console.log("Send mail for critical document");
      mailSend = "Yes";
    }
    if (mailSend === "Yes") {
      const emailNotification: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.emailNotification);
      console.log(emailNotification);
      for (let k in emailNotification) {
        if (emailNotification[k].Title === type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }

      }

      link = `<a href=${this.state.linkToDoc}>` + this.state.documentName + `</a>`;
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
      let replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);
      let replacelink = replaceString(replaceBody, '[DocumentLink]', link);
      let FinalBody = replacelink;
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject,
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
  //Document Return
  public async _returnDoc() {
    let message;
    let logstatus;
    if (this.props.project) {
      const indexdata2 = {
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata2, this.documentIndexID);
      const headerdata2 = {
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata2, this.workflowHeaderID);
      const sourcedata = {
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocument, sourcedata, this.sourceDocumentID);
    }
    else {
      const indexdata3 = {
        WorkflowStatus: this.state.status
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata3, this.documentIndexID);
      const headerdata3 = {
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status
      }
      await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, headerdata3, this.workflowHeaderID);
      const sourcedata1 = {
        WorkflowStatus: this.state.status
      }
      await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocument, sourcedata1, this.sourceDocumentID);
    }
    await this.triggerDocumentReview(this.sourceDocumentID, this.state.status)
    if (this.state.status === "Returned with comments") {
      message = this.documentReturnSuccess;
      logstatus = "Returned with comments";
      if (this.state.ownerEmail === this.state.requesterEmail) {
        this._sendMail(this.state.ownerEmail, "DocReturn", this.state.ownerName);
      }
      else {
        this._sendMail(this.state.ownerEmail, "DocReturn", this.state.ownerName);
        this._sendMail(this.state.requesterEmail, "DocReturn", this.state.requesterName);
      }


    }
    else {
      message = this.documentRejectSuccess;
      logstatus = "Rejected";

      if (this.state.ownerEmail === this.state.requesterEmail) {
        this._sendMail(this.state.ownerEmail, "DocRejected", this.state.ownerName);
      }
      else {
        this._sendMail(this.state.ownerEmail, "DocRejected", this.state.ownerName);
        this._sendMail(this.state.requesterEmail, "DocRejected", this.state.requesterName);
      }
    }
    const logdata = {
      Status: logstatus,
      Workflow: "Approval"
    }
    let a = await this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata, this.revisionLogId);
    await this._Service.deletehubItemById(this.props.hubUrl, this.props.workflowTasksList, parseInt(this.state.taskID));

    this.setState({
      hideLoading: true,
      statusMessage: { isShowMessage: true, message: message, messageType: 4 }
    });
    setTimeout(() => {
      window.location.replace(this.props.siteUrl);
    }, 5000);
  }
  // Document Review trigger
  protected async triggerDocumentReview(sourceDocumentID: any, status: any) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': status
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  //Cancel Document
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
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexID);
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
  public render(): React.ReactElement<ITransmittalApproveDocumentProps> {
    const status: IDropdownOption[] = [
      { key: 'Approved', text: 'Approved' },
      { key: 'Returned with comments', text: 'Returned with comments' },
      { key: 'Rejected', text: 'Rejected' },
    ];
    const publishOptions: IDropdownOption[] = [
      { key: 'PDF', text: 'PDF' },
      { key: 'Native', text: 'Native' },
    ];
    const publishOption: IDropdownOption[] = [
      { key: 'Native', text: 'Native' },
    ];
    const DownIcon: IIconProps = { iconName: 'ChevronDown' };
    return (
      <section className={`${styles.transmittalApproveDocument}`}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div style={{ display: this.state.access }}>

          <div className={styles.border}>
            <div className={styles.alignCenter}> {this.props.webpartHeader}</div>
            <br></br>
            <div hidden={this.state.hideProject}>
              <div className={styles.header}>
                <div className={styles.divMetadataCol1}>
                  <h3 >Project Details</h3>
                  <a href={this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexID} style={{ marginLeft: "79%", marginTop: "17px" }} target="_blank">Revision History</a>
                </div>
              </div>
              <div className={styles.divMetadata}>
                <div className={styles.divMetadataCol1}>
                  <Label className={styles.bold}>Project Name : </Label><div className={styles.divLabel}>{this.state.projectName}</div>
                </div>
                <div className={styles.divMetadataCol3}>
                  <Label className={styles.bold}>Project Number :</Label><div className={styles.divLabel}> {this.state.projectNumber}</div>
                </div>
              </div>
            </div>
            <div className={styles.header}>
              <div className={styles.divMetadataCol1}>
                <h3 >Document Details</h3>
              </div>
            </div>
            <div className={styles.divMetadata}>
              <div className={styles.divMetadataCol1}>
                <Label >Document ID : </Label><div className={styles.divLabel}>{this.state.documentID}</div>
              </div>
              <div className={styles.divMetadataCol3}>
                <Label >Revision :</Label><div className={styles.divLabel}> {this.state.revision}</div>
              </div>

            </div>
            <div className={styles.divRow}>
              <Label >Document :</Label><div className={styles.divLabel}>  <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></div>
            </div>
            <div className={styles.header}>
              <h3 className="ExampleCard-title title-222">Workflow Details</h3>
            </div>
            <div className={styles.divMetadata}>
              <div className={styles.divMetadataCol1}>
                <Label className={styles.bold}>Owner : </Label><div className={styles.divLabel}> {this.state.ownerName}</div>
              </div>
              <div className={styles.divMetadataCol2}><Label className={styles.bold}>Due Date :</Label> <div className={styles.divLabel}> {this.state.dueDate}</div></div>
              <div className={styles.divMetadataCol3}><Label className={styles.bold}>Requested Date :</Label><div className={styles.divLabel}>{this.state.requestedDate} </div></div>
            </div>
            <div className={styles.divMetadata}>
              <div style={{ display: "flex" }}>
                <Label className={styles.bold}>Requester :</Label> <div className={styles.divLabel}>{this.state.requesterName}</div>
              </div>
              <div style={{ marginLeft: "16.9em", display: "flex" }}><Label className={styles.bold}>Requester Comment: </Label><div className={styles.divLabel}>{this.state.requesterComment}</div></div>
            </div>
            <div >
              <br></br>
              <div hidden={this.state.hideProject} >
                <div style={{ display: this.state.dccTableDiv }}>
                  <Accordion allowZeroExpanded className={styles.Accordion}>
                    <AccordionItem >
                      <AccordionItemHeading>
                        <AccordionItemButton className={styles.AccordionItemButton}>
                          <Label className={styles.pleft}><IconButton iconProps={DownIcon} />Document Controller Review Details</Label>
                        </AccordionItemButton>
                      </AccordionItemHeading>
                      <AccordionItemPanel>
                        <div style={{ display: (this.state.dccreviewerData.length === 0 ? 'none' : 'block') }}>
                          <table className={styles.tableClass}   >
                            <tr className={styles.tr}>
                              <th className={styles.th}>Document Controller</th>
                              <th className={styles.th}>Document Controller Date</th>
                              <th className={styles.th}>Document Controller Comment</th>
                            </tr>
                            <tbody className={styles.tbody}>
                              {this.state.dccreviewerData.map((item) => {
                                return (<tr className={styles.tr}>
                                  <td className={styles.th}>{item.Reviewer}</td>
                                  <td className={styles.th}>{item.ResponseDate}</td>
                                  <td className={styles.th}>{item.DCCResponsibleComment}</td>
                                </tr>);
                              })
                              }
                            </tbody>
                          </table>
                        </div>
                      </AccordionItemPanel>
                    </AccordionItem>
                  </Accordion>
                </div>
              </div>
              <br></br>
              <div style={{ display: this.state.reviewersTableDiv }}>
                <Accordion allowZeroExpanded className={styles.Accordion}>
                  <AccordionItem >
                    <AccordionItemHeading>
                      <AccordionItemButton className={styles.AccordionItemButton}>
                        <Label className={styles.pleft}><IconButton iconProps={DownIcon} />Review Details</Label>
                      </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <div style={{ display: (this.state.reviewerData.length === 0 ? 'none' : 'block') }}>
                        <table className={styles.tableClass}   >
                          <tr className={styles.tr}>
                            <th className={styles.th}>Reviewer</th>
                            <th className={styles.th}>Review Date</th>
                            <th className={styles.th}>Review Comment</th>
                          </tr>
                          <tbody className={styles.tbody}>
                            {this.state.reviewerData.map((item) => {
                              return (<tr className={styles.tr}>
                                <td className={styles.th}>{item.Reviewer}</td>
                                <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                                <td className={styles.th}>{item.ResponsibleComment}</td>
                              </tr>);
                            })
                            }
                          </tbody>
                        </table>
                      </div>
                    </AccordionItemPanel>
                  </AccordionItem>
                </Accordion>

              </div>
            </div>
            <div className={styles.header}>
              <h3 className="ExampleCard-title title-222"></h3>
            </div>
            <div className={styles.divMetadata}>
              <div style={{ width: "100%", }}>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={status}
                  onChanged={this._status}
                  selectedKey={this.state.statusKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("Docstatus", this.state.statusKey, "required")}{" "}</div>
              </div>
              <div hidden={this.state.hidepublish} style={{ width: "100%", marginLeft: "10px" }}>
                <div style={{ display: this.state.isdocx }}>
                  <Dropdown id="t2" required={true}
                    label="Publish Option"
                    selectedKey={this.state.publishOption}
                    defaultSelectedKey={this.state.publishOptionKey}
                    placeholder="Select an option"
                    options={publishOptions}
                    onChanged={this._drpdwnPublishFormat} /></div>
                <div style={{ display: this.state.nodocx }}>
                  <Dropdown id="t2" required={true}
                    label="Publish Option"
                    selectedKey={this.state.publishOption}
                    placeholder="Select an option"
                    options={publishOption}
                    onChanged={this._drpdwnPublishFormat} /></div>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("publish", this.state.publishOption, "required")}{""}</div></div>
              <div hidden={this.state.hideProject} style={{ width: "100%", marginLeft: "10px" }}>
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.acceptanceCodearray}
                    onChanged={this._acceptanceChanged}
                    selectedKey={this.state.acceptanceCode}
                  /></div>

              </div>
            </div>
            <div >
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentsChange} multiline required={this.state.commentRequired} autoAdjustHeight></TextField></div>
              <div style={{ display: this.state.commentValid }}>
                <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div></div>
              <div> {this.state.statusMessage.isShowMessage ?
                <MessageBar
                  messageBarType={this.state.statusMessage.messageType}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.statusMessage.message}</MessageBar>
                : ''} </div>
              <div className={styles.mt}>
                <div hidden={this.state.hideLoading}>
                  <Spinner label={"Publishing... "} />
                </div>
              </div>
              <div className={styles.mt}>
                <div hidden={this.state.hideLoading} style={{ color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                  <Label>***PLEASE DON'T REFRESH***</Label>
                </div>
              </div>
              <div className={styles.divRow}>
                <div style={{ fontStyle: "italic", fontSize: "12px", position: "absolute" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>

                <div className={styles.rgtalign} >
                  <PrimaryButton id="b2" className={styles.btn} onClick={this._saveAsDraft} style={{ display: this.state.savedisable }}>Save as Draft</PrimaryButton >
                  <PrimaryButton id="b2" className={styles.btn} onClick={this._docSave} style={{ display: this.state.savedisable }}>Submit</PrimaryButton >
                  <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
                </div>
              </div>

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
              <br />
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
