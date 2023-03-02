import * as React from 'react';
import styles from './TransmittalCreateDocument.module.scss';
import { ITransmittalCreateDocumentProps, ITransmittalCreateDocumentState } from './ITransmittalCreateDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ActionButton, Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IChoiceGroupOption, IChoiceGroupStyles, IDropdownOption, IIconProps, ITooltipHostStyles, Label, MessageBar, PrimaryButton, Spinner, TextField, TooltipHost } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { BaseService } from '../services';
import SimpleReactValidator from 'simple-react-validator';
import { Web } from '@pnp/sp/webs';
import * as _ from 'lodash';
import replaceString from 'replace-string';
import * as moment from 'moment';
import { IHubSiteWebData } from "@pnp/sp/hubsites";
import { IHttpClientOptions, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';

export default class TransmittalCreateDocument extends React.Component<ITransmittalCreateDocumentProps, ITransmittalCreateDocumentState, {}> {
  private _Service: BaseService;
  private validator: SimpleReactValidator;
  private reqWeb: any;
  private directPublish: any;
  private createDocument: any;
  private noAccess: any;
  private getSelectedReviewers: any[] = [];
  private isDocument: any;
  private myfile: any;
  private documentNameExtension: any;
  private Timeout = 5000;
  private Count: any;
  private currentCount: any;
  private revokeUrl: any;
  private postUrl: any;
  private documentIndexID: any;
  private revisionHistoryUrl: any;
  private permissionpostUrl: any;
  public constructor(props: ITransmittalCreateDocumentProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      title: "",
      createDocumentView: "none",
      createDocumentProject: "none",
      accessDeniedMessageBar: "none",
      businessUnitID: null,
      departmentId: null,
      categoryId: null,
      legalEntityOption: [],
      businessUnitOption: [],
      departmentOption: [],
      categoryOption: [],
      saveDisable: false,
      businessUnit: "",
      businessUnitCode: "",
      departmentCode: "",
      department: "",
      subCategoryArray: [],
      subCategoryId: null,
      category: "",
      subCategory: "",
      categoryCode: "",
      legalEntityId: null,
      legalEntity: "",
      reviewers: [],
      approver: null,
      approverEmail: "",
      approverName: "",
      owner: "",
      ownerEmail: "",
      ownerName: "",
      dcc: null,
      dccEmail: "",
      dccName: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      revisionCodingId: null,
      revisionLevelId: null,
      revisionLevelArray: [],
      revisionSettingsArray: [],
      revisionCoding: "",
      revisionLevel: "",
      transmittalCheck: false,
      externalDocument: false,
      hideDoc: "",
      createDocument: false,
      templateDocuments: "",
      templateId: "",
      hidePublish: "none",
      hideExpiry: "",
      expiryCheck: false,
      expiryDate: null,
      expiryLeadPeriod: "",
      directPublishCheck: false,
      templateKey: "",
      approvalDate: "",
      publishOption: "",
      approvalDateEdit: null,
      leadmsg: "none",
      criticalDocument: false,
      templateDocument: false,
      hideLoading: true,
      hideloader: true,
      incrementSequenceNumber: "",
      documentid: "",
      documentName: "",
      projectName: "",
      projectNumber: "",
      newDocumentId: "",
      sourceDocumentId: "",
      newRevision: "",
      messageBar: "none",
      previousRevisionItemID: null,
      revisionItemID: "",
      currentRevision: "",
      isdocx: "none",
      nodocx: "",
      insertdocument: "none",
      loaderDisplay: "",
      uploadDocumentError: "none",
      checkdirect: "none",
      hideDirect: "none",
      validApprover: "none",
      hideCreateLoading: "none",
      subContractorNumber: "",
      customerNumber: "",
      documentCount: "",
      bulkDocumentIndex: "none",
      counter: "",
      norefresh: "none",
      IsLot: false,
      upload: false,
      template: false,
      hideupload: "none",
      sourceId: "",
      hidetemplate: "none",
      hidesource: "none",
      uploadOrTemplateRadioBtn: "",
      CurrentUserEmail: "",
      currentUserId: null,
      currentUser: "",
      today: null
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + "/" + this.props.QDMSUrl, window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    this._bindData = this._bindData.bind(this);
    this._businessUnitChange = this._businessUnitChange.bind(this);
    this._departmentChange = this._departmentChange.bind(this);
    this._categoryChange = this._categoryChange.bind(this);
    this._subCategoryChange = this._subCategoryChange.bind(this);
    this._selectedOwner = this._selectedOwner.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    this._sourcechange = this._sourcechange.bind(this);
    this._templatechange = this._templatechange.bind(this);
    this._onDirectPublishChecked = this._onDirectPublishChecked.bind(this);
    this._onApprovalDatePickerChange = this._onApprovalDatePickerChange.bind(this);
    this._publishOptionChange = this._publishOptionChange.bind(this);
    this._onExpDatePickerChange = this._onExpDatePickerChange.bind(this);
    this._expLeadPeriodChange = this._expLeadPeriodChange.bind(this);
    this._onCriticalChecked = this._onCriticalChecked.bind(this);
    this._onTemplateChecked = this._onTemplateChecked.bind(this);
    this._onCreateDocument = this._onCreateDocument.bind(this);
    this._documentidgeneration = this._documentidgeneration.bind(this);
    this._incrementSequenceNumber = this._incrementSequenceNumber.bind(this);
    this._documentCreation = this._documentCreation.bind(this);
    this._addSourceDocument = this._addSourceDocument.bind(this);
    this._project = this._project.bind(this);
    this._revisionLevelChange = this._revisionLevelChange.bind(this);
    this._revisionCodingChange = this._revisionCodingChange.bind(this);
    this._createDocumentIndex = this._createDocumentIndex.bind(this);
    this._selectedDCC = this._selectedDCC.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._generateNewRevision = this._generateNewRevision.bind(this);
    this._add = this._add.bind(this);
    this._checkdirectPublish = this._checkdirectPublish.bind(this);
    this._subContractorNumberChange = this._subContractorNumberChange.bind(this);
    this._CustomerNumberChange = this._CustomerNumberChange.bind(this);
    this._documentCount = this._documentCount.bind(this);
    this._onCreateIndex = this._onCreateIndex.bind(this);
    this._createMultipleIndex = this._createMultipleIndex.bind(this);

    this.onUploadOrTemplateRadioBtnChange = this.onUploadOrTemplateRadioBtnChange.bind(this);
  }
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: { required: "This field is mandatory" }
    });
  }
  // On load
  public async componentDidMount() {
    //Get Current User
    this._Service.getCurrentUser()
      .then(async (user: any) => {
        console.log("user ", user)
        this.setState({
          currentUserId: user.Id,
          currentUser: user.Title,
          CurrentUserEmail: user.Email
        });
      });

    //Get Today
    this.setState({ approvalDate: new Date(), today: new Date() });
    if (this.props.project) {
      this.setState({ createDocumentView: "none", createDocumentProject: "", loaderDisplay: "none" });
      this._bindData();
      this._project();
      this._checkdirectPublish('Project_DirectPublish');
    }
    else {
      this._checkdirectPublish('QDMS_DirectPublish');
      this.setState({ createDocumentView: "", createDocumentProject: "none", loaderDisplay: "none" });
    }
  }
  //Bind dropdown in create
  public async _bindData() {
    let businessUnitArray = [];
    let sorted_BusinessUnit: any[];
    let departmentArray = [];
    let sorted_Department: any[];
    let categoryArray = [];
    let sorted_Category: any[];
    let legalEntityArray = [];
    let sorted_LegalEntity: any[];
    //Get Business Unit
    const businessUnit: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.businessUnit);
    for (let i = 0; i < businessUnit.length; i++) {
      let businessUnitdata = {
        key: businessUnit[i].ID,
        text: businessUnit[i].BusinessUnitName,
      };
      businessUnitArray.push(businessUnitdata);
    }
    sorted_BusinessUnit = _.orderBy(businessUnitArray, 'text', ['asc']);
    //Get Department
    const department: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.department);
    for (let i = 0; i < department.length; i++) {
      let departmentdata = {
        key: department[i].ID,
        text: department[i].Department,
      };
      departmentArray.push(departmentdata);
    }
    sorted_Department = _.orderBy(departmentArray, 'text', ['asc']);
    //Get Category
    const category: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.category);
    let categorydata;
    for (let i = 0; i < category.length; i++) {
      if (this.props.project) {
        if (category[i].Project === true) {
          categorydata = {
            key: category[i].ID,
            text: category[i].Category,
          };
          categoryArray.push(categorydata);
        }
      }
      else {
        if (category[i].QDMS === true) {
          categorydata = {
            key: category[i].ID,
            text: category[i].Category,
          };
          categoryArray.push(categorydata);
        }
      }
    }
    sorted_Category = _.orderBy(categoryArray, 'text', ['asc']);
    //Get Legal Entity
    if (!this.props.project) {
      const legalEntity: any = await this._Service.gethubListItems(this.props.hubUrl, this.props.legalEntity);
      // const legalEntity: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.legalEntity).items.select("Title,ID").get();
      for (let i = 0; i < legalEntity.length; i++) {
        let legalEntityItemdata = {
          key: legalEntity[i].ID,
          text: legalEntity[i].Title
        };
        legalEntityArray.push(legalEntityItemdata);
      }
      sorted_LegalEntity = _.orderBy(legalEntityArray, 'text', ['asc']);
    }
    this.setState({
      businessUnitOption: sorted_BusinessUnit,
      departmentOption: sorted_Department,
      categoryOption: sorted_Category,
      legalEntityOption: sorted_LegalEntity,
      owner: this.state.currentUserId,
      ownerEmail: this.state.CurrentUserEmail,
      ownerName: this.state.currentUser
    });
    this._userMessageSettings();
  }
  //Bind data on Project
  public async _project() {
    let revisionLevelArray = [];
    let sorted_RevisionLevel = [];
    let revisionSettingsArray = [];
    let sorted_RevisionSettings = [];
    let businessCode;
    let BusinessUnitName;
    let BusinessUnitID;
    //Get Revision Level
    const revisionLevelItem: any = await this._Service.getDrpdwnListItems(this.props.siteUrl, this.props.revisionLevelList);
    // const revisionLevelItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.revisionLevelList).items.select("Title,ID").get();
    for (let i = 0; i < revisionLevelItem.length; i++) {
      let revisionLevelItemdata = {
        key: revisionLevelItem[i].ID,
        text: revisionLevelItem[i].Title
      };
      revisionLevelArray.push(revisionLevelItemdata);
    }
    sorted_RevisionLevel = _.orderBy(revisionLevelArray, 'text', ['asc']);
    //Get RevisionSettings
    const revisionSettingsItem: any = await this._Service.getDrpdwnListItems(this.props.siteUrl, this.props.revisionSettingsList);
    // const revisionSettingsItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.revisionSettingsList).items.select("Title,ID").get();
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
    // const projectInformation = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.projectInformationListName).items.get();
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
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
        if (projectInformation[k].Key === "BusinessUnit") {
          const businessUnit = await this._Service.gethubListItems(this.props.hubUrl, this.props.businessUnit);
          console.log(businessUnit);
          for (let i = 0; i < businessUnit.length; i++) {
            if (businessUnit[i].Title === projectInformation[k].Title) {
              console.log(businessUnit[i]);
              businessCode = businessUnit[i].Title;
              BusinessUnitName = businessUnit[i].BusinessUnitName;
              BusinessUnitID = businessUnit[i].ID;

            }
          }
          // this.setState({ businessUnitID: BusinessUnitID, businessUnitCode: businessCode, businessUnit: BusinessUnitName });

        }
      }
    }
    this.setState({
      revisionSettingsArray: sorted_RevisionSettings,
      revisionLevelArray: sorted_RevisionLevel
    });
  }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this._Service.gethubUserMessageListItems(this.props.hubUrl, this.props.userMessageSettings);
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title === "CreateDocumentSuccess") {
        var successmsg = userMessageSettings[i].Message;
        this.createDocument = replaceString(successmsg, '[DocumentName]', this.state.documentName);
      }
      if (userMessageSettings[i].Title === "DirectPublishSuccess") {
        var publishmsg = userMessageSettings[i].Message;
        this.directPublish = replaceString(publishmsg, '[DocumentName]', this.state.documentName);
      }
      if (userMessageSettings[i].Title === "NoAccess") {
        this.noAccess = userMessageSettings[i].Message;

      }
    }
  }
  public _createMultipleIndex() {
    this.setState({ bulkDocumentIndex: " ", createDocumentProject: "none" });
  }
  // Direct publish change
  public async _checkdirectPublish(type: any) {
    const laUrl = await this._Service.getdirectpublish(this.props.hubUrl, this.props.requestList);
    console.log("Posturl", laUrl[0].PostUrl);
    this.permissionpostUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.permissionpostUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': type,
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.state.CurrentUserEmail

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
      if (responseJSON['Status'] === "Valid") {
        this.setState({ checkdirect: "none", hideDirect: "", hidePublish: "none" });
      }
      else {
        this.setState({ checkdirect: "none", hideDirect: "none", hidePublish: "none" });
      }
    }
    else { }
  }
  //Title Change
  public _titleChange = (ev: React.FormEvent<HTMLInputElement>, title?: string) => {
    this.setState({ title: title || '', saveDisable: false });
  }
  // BusinessUnit Change
  public async _businessUnitChange(option: { key: any; text: any }) {
    let getApprover = [];
    let approverEmail;
    let approverName;
    const businessUnit = await this._Service.gethubItemById(this.props.hubUrl, this.props.businessUnit, option.key);
    let businessCode = businessUnit.Title;
    this.setState({ businessUnitID: option.key, businessUnitCode: businessCode, businessUnit: option.text, saveDisable: false });
    if (this.props.project) { }
    else {
      const businessUnits = await this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.businessUnit);
      console.log(businessUnits);
      for (let i = 0; i < businessUnits.length; i++) {
        if (businessUnits[i].ID === option.key) {
          const approve = await this._Service.getByEmail(businessUnits[i].Approver.EMail);
          console.log("approve:" + approve);
          approverEmail = businessUnits[i].Approver.EMail;
          approverName = businessUnits[i].Approver.Title;
          getApprover.push(approve.Id);
        }
      }
      this.setState({ approver: getApprover[0], approverEmail: approverEmail, approverName: approverName });
    }
  }
  //Department Change
  public async _departmentChange(option: { key: any; text: any }) {
    let getApprover = [];
    let approverEmail;
    let approverName;
    const department = await this._Service.gethubItemById(this.props.hubUrl, this.props.department, option.key);
    let departmentCode = department.Title;
    this.setState({ departmentId: option.key, departmentCode: departmentCode, department: option.text, saveDisable: false });
    if (this.props.project) { }
    else {
      if (this.state.businessUnitCode === "") {
        const departments = await this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.department);
        for (let i = 0; i < departments.length; i++) {
          if (departments[i].ID === option.key) {
            const deptapprove = await this._Service.getByEmail(departments[i].Approver.EMail);
            approverEmail = departments[i].Approver.EMail;
            approverName = departments[i].Approver.Title;
            getApprover.push(deptapprove.Id);
          }
        }
        this.setState({ approver: getApprover[0], approverEmail: approverEmail, approverName: approverName });
      }
    }
  }
  //Category Change
  public async _categoryChange(option: { key: any; text: any }) {
    let subcategoryArray: any[] = [];
    let sorted_subcategory: any[];
    let category = await this._Service.gethubItemById(this.props.hubUrl, this.props.category, option.key);
    let categoryCode = category.Title;
    await this._Service.gethubListItems(this.props.hubUrl, this.props.subCategory).then((subcategory: any) => {
      for (let i = 0; i < subcategory.length; i++) {
        if (subcategory[i].CategoryId === option.key) {
          let subcategorydata = {
            key: subcategory[i].ID,
            text: subcategory[i].SubCategory,
          };
          subcategoryArray.push(subcategorydata);
        }
      }
      sorted_subcategory = _.orderBy(subcategoryArray, 'text', ['asc']);
      this.setState({
        categoryId: option.key,
        subCategoryArray: sorted_subcategory,
        category: option.text,
        categoryCode: categoryCode,
        saveDisable: false
      });
    });
  }
  //SubCategory Change
  public _subCategoryChange(option: { key: any; text: any }) {
    this.setState({ subCategoryId: option.key, subCategory: option.text });
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
    let getSelectedApprover: any[] = [];
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
      if (this.state.businessUnitCode !== "") {
        this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.businessUnit).then(async (businessUnit: any) => {
          for (let i = 0; i < businessUnit.length; i++) {
            if (businessUnit[i].ID === this.state.businessUnitID) {
              const approve = await this._Service.getByEmail(businessUnit[i].Approver.EMail);
              approverEmail = businessUnit[i].Approver.EMail;
              approverName = businessUnit[i].Approver.Title;
              getSelectedApprover.push(approve.Id);
            }
          }
        })
      }
      else {
        this._Service.getBusinessUnitItem(this.props.hubUrl, this.props.department).then(async (departments: any) => {
          for (let i = 0; i < departments.length; i++) {
            if (departments[i].ID === this.state.departmentId) {
              const deptapprove = await this._Service.getByEmail(departments[i].Approver.EMail);
              approverEmail = departments[i].Approver.EMail;
              approverName = departments[i].Approver.Title;
              getSelectedApprover.push(deptapprove.Id);
            }
          }
        });
      }
      this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName, saveDisable: false });
      setTimeout(() => {
        this.setState({ validApprover: "none" });
      }, 5000);
    }


  }
  //Revision Settings Change
  public _revisionCodingChange(option: { key: any; text: any }) {
    this.setState({ revisionCodingId: option.key, revisionCoding: option.text });
  }
  //Revision Level Change
  public _revisionLevelChange(option: { key: any; text: any }) {
    this.setState({ revisionLevelId: option.key, revisionLevel: option.text });
  }
  // on subcontractor number change
  public _subContractorNumberChange = (ev: React.FormEvent<HTMLInputElement>, subContractorNumber?: string) => {
    this.setState({ subContractorNumber: subContractorNumber || '' });
  }
  public _CustomerNumberChange = (ev: React.FormEvent<HTMLInputElement>, customerNumber?: string) => {
    this.setState({ customerNumber: customerNumber || '' });
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
  // On format date
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
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
  // On upload
  public _add(e: any) {
    this.setState({ insertdocument: "none" });
    this.myfile = e.target.value;
    let upload;
    let type;
    let myfile;
    this.isDocument = "Yes";
    if (this.props.project) {
      upload = "#addproject";
      myfile = (document.querySelector("#addproject") as HTMLInputElement).files[0];
    }
    else {
      upload = "#addqdms";
      myfile = (document.querySelector("#addqdms") as HTMLInputElement).files[0];
    }
    //let myfile = (document.querySelector(upload) as HTMLInputElement).files[0];
    console.log(myfile);
    this.isDocument = "Yes";
    var splitted = myfile.name.split(".");
    // let docsplit =splitted.slice(0, -1).join('.')+"."+splitted[splitted.length - 1];
    // alert(docsplit);
    type = splitted[splitted.length - 1];
    if (type === "docx") {
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
    if (option.key === "QDMS") {
      let publishedDocument: any[] = await this._Service.getqdmsLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary);
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
    }
    else {
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
    if (this.state.sourceId === "QDMS") {
      await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id === this.state.templateId) {

            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type === "docx") {
          this.setState({ isdocx: "", nodocx: "none" });
        }
        else {
          this.setState({ isdocx: "none", nodocx: "" });
        }
      });
    }
    else {
      await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
        console.log(publishdoc);
        for (let i = 0; i < publishdoc.length; i++) {
          if (publishdoc[i].Id === this.state.templateId) {
            publishName = publishdoc[i].LinkFilename;
          }
        }
        var split = publishName.split(".", 2);
        type = split[1];
        if (type === "docx") {
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
      // if (this.props.project) {
      //   this.setState({ checkdirect: "", });
      //   this._checkdirectPublish('Project_DirectPublish');
      // }
      // else {
      //   this.setState({ checkdirect: "", });
      //   this._checkdirectPublish('QDMS_DirectPublish');
      // }
      this.setState({ hidePublish: "", directPublishCheck: true, approvalDate: new Date() });
    }
    else if (!isChecked) { this.setState({ hidePublish: "none", directPublishCheck: false, approvalDate: new Date(), publishOption: "" }); }
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
  //On create button click
  public async _onCreateDocument() {
    if (this.state.createDocument === true && this.isDocument === "Yes" || this.state.createDocument === false) {
      if (this.state.expiryDate !== null || this.state.expiryDate !== undefined) {
        if (this.props.project) {
          //Validation without direct publish
          if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
            // && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')
            // && this.validator.fieldValid('Revision') 
            && this.validator.fieldValid('DocumentController')) {
            this.setState({
              saveDisable: true, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner')
            && this.validator.fieldValid('Approver')
            // && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod') 
            // && this.validator.fieldValid('Revision') 
            && this.validator.fieldValid('DocumentController')) {
            this.setState({
              saveDisable: true, hideloader: false, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
        else {
          //Validation without direct publish
          if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('legalentity')
            // && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')
          ) {
            this.setState({
              saveDisable: true, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('legalentity')
            //  && this.validator.fieldValid('expiryDate') && this.validator.fieldValid('ExpiryLeadPeriod')
          ) {
            this.setState({
              saveDisable: true, hideloader: false, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
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
          if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner')
            && this.validator.fieldValid('Approver')
            // && this.validator.fieldValid('Revision') 
            && this.validator.fieldValid('DocumentController')) {
            this.setState({
              saveDisable: true, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep')
            && (this.state.directPublishCheck === true) && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner')
            && this.validator.fieldValid('Approver')
            // && this.validator.fieldValid('Revision') 
            && this.validator.fieldValid('DocumentController')) {
            this.setState({
              saveDisable: true, hideloader: false, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          else {
            this.validator.showMessages();
            this.forceUpdate();
          }
        }
        else {
          //Validation without direct publish
          if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver') && this.validator.fieldValid('legalentity')) {
            this.setState({
              saveDisable: true, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
            this.validator.hideMessages();
          }
          //Validation with direct publish
          else if (this.validator.fieldValid('Title') && this.validator.fieldValid('category')
            && this.validator.fieldValid('BU/Dep') && (this.state.directPublishCheck === true)
            && this.validator.fieldValid('publish') && this.validator.fieldValid('Owner')
            && this.validator.fieldValid('Approver') && this.validator.fieldValid('legalentity')) {
            this.setState({
              saveDisable: true, hideloader: false, hideCreateLoading: " ",
              norefresh: " "
            });
            await this._documentidgeneration();
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
  //Documentid generation
  public async _documentidgeneration() {
    let prefix;
    let separator;
    let sequenceNumber;
    let idcode;
    let counter;
    var incrementstring;
    let increment;
    let documentid;
    let isValue = "false";
    let settingsid;
    let documentname;
    // Get document id settings
    const documentIdSettings = await this._Service.getListItems(this.props.siteUrl, this.props.documentIdSettings);
    console.log("documentIdSettings", documentIdSettings);
    prefix = "EMEC";
    separator = documentIdSettings[0].Separator;
    sequenceNumber = documentIdSettings[0].SequenceDigit;
    // Project id prefix
    if (this.props.project) {
      if (this.state.businessUnitCode !== "") {
        idcode = this.state.businessUnitCode + separator + this.state.projectNumber + separator + this.state.categoryCode;
      }
      else {
        idcode = this.state.departmentCode + separator + this.state.projectNumber + separator + this.state.categoryCode;
      }
    }
    // Qdms id prefix
    else {
      // if (this.state.businessUnitCode !== "") {
      //   idcode = this.state.legalEntity;
      // }
      // else {
      idcode = this.state.legalEntity;
      // }
    }
    if (documentIdSettings) {
      // Get sequence of id
      const documentIdSequenceSettings = await this._Service.getListItems(this.props.siteUrl, this.props.documentIdSequenceSettings);
      console.log("documentIdSequenceSettings", documentIdSequenceSettings);
      for (var k in documentIdSequenceSettings) {
        if (documentIdSequenceSettings[k].Title === idcode) {
          counter = documentIdSequenceSettings[k].Sequence;
          settingsid = documentIdSequenceSettings[k].ID;
          isValue = "true";
        }
      }
      if (documentIdSequenceSettings) {
        // No sequence
        if (isValue === "false") {
          increment = 1;
          incrementstring = increment.toString();
          const idseq = {
            Title: idcode,
            Sequence: incrementstring
          }
          const addidseq = await this._Service.createNewItem(this.props.siteUrl, this.props.documentIdSequenceSettings, idseq);
          if (addidseq) {
            await this._incrementSequenceNumber(incrementstring, sequenceNumber);
            if (this.props.project) {
              documentid = idcode + separator + this.state.incrementSequenceNumber;
            }
            else {
              if (this.state.departmentCode !== "") {
                documentid = this.state.legalEntity + separator + this.state.departmentCode + separator + this.state.incrementSequenceNumber;
              }
              else {
                documentid = this.state.legalEntity + separator + this.state.businessUnitCode + separator + this.state.incrementSequenceNumber;
              }
              // documentid = idcode+ separator + this.state.incrementSequenceNumber;
            }
            documentname = documentid + " " + this.state.title;

            this.setState({ documentid: documentid, documentName: documentname });
            await this._documentCreation();
          }
        }
        // Has sequence
        else {
          increment = parseInt(counter) + 1;
          incrementstring = increment.toString();
          let updateCounter = {
            Title: idcode,
            Sequence: incrementstring
          }
          const afterCounter = await this._Service.updateItem(this.props.siteUrl, this.props.documentIdSequenceSettings, updateCounter, settingsid);
          if (afterCounter) {
            await this._incrementSequenceNumber(incrementstring, sequenceNumber);
            if (this.props.project) {
              documentid = idcode + separator + this.state.incrementSequenceNumber;
            }
            else {
              if (this.state.departmentCode !== "") {
                documentid = this.state.legalEntity + separator + this.state.departmentCode + separator + this.state.incrementSequenceNumber;
              }
              else {
                documentid = this.state.legalEntity + separator + this.state.businessUnitCode + separator + this.state.incrementSequenceNumber;
              }
              // documentid = idcode+ separator + this.state.incrementSequenceNumber;
            }
            if (this.props.project) {
              if (this.state.bulkDocumentIndex === " ") {
                documentname = documentid + " " + this.state.title + " " + this.state.counter;
              }
              else {
                documentname = documentid + " " + this.state.title;
              }
            }
            else {
              documentname = documentid + " " + this.state.title;
            }

            this.setState({ documentid: documentid, documentName: documentname });

            await this._documentCreation();
          }
        }
      }
    }
  }
  // Append sequence to the count
  public _incrementSequenceNumber(incrementvalue: any, sequenceNumber: any) {
    var incrementSequenceNumber = incrementvalue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }
  // Create item with id
  public async _documentCreation() {
    await this._userMessageSettings();
    let documentNameExtension: any;
    let sourceDocumentId: any;
    let docServerUrl: any;
    let splitdocUrl: any;
    let documenturl: any;
    let upload: any;
    let docinsertname: any;
    if (this.props.project) {
      upload = "#addproject";
    }
    else {
      upload = "#addqdms";
    }
    // With document
    if (this.state.createDocument === true) {
      // Create document index item
      await this._createDocumentIndex();
      // Get file from form
      if ((document.querySelector(upload) as HTMLInputElement).files[0] != null) {
        let myfile = (document.querySelector(upload) as HTMLInputElement).files[0];
        console.log(myfile);
        var splitted = myfile.name.split(".");
        documentNameExtension = this.state.documentName + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        docinsertname = this.state.documentid + '.' + splitted[splitted.length - 1];
        if (myfile.size) {
          // add file to source library
          // const fileUploaded = await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/").files.add(docinsertname, myfile, true);
          const fileUploaded = await this._Service.uploadDocument(this.props.sourceDocumentLibrary, docinsertname, myfile);
          if (fileUploaded) {
            console.log("File Uploaded");
            const item = await fileUploaded.file.getItem();
            console.log(item);
            sourceDocumentId = item["ID"];
            // if(splitted[1] === "pdf"||splitted[1] === "Pdf"||splitted[1] === "PDF"){
            //   documenturl = this.props.siteUrl + "/" + this.props.sourceDocumentLibrary  + "/" + documentNameExtension;
            //   // documenturl = item["ServerRedirectedEmbedUrl"];
            // }
            // else{
            // docServerUrl = item["ServerRedirectedEmbedUrl"];
            // splitdocUrl = docServerUrl.split("&", 2);
            // documenturl = splitdocUrl[0];
            // }
            this.setState({ sourceDocumentId: sourceDocumentId });
            // update metadata
            await this._addSourceDocument();
            if (item) {
              let revision;
              if (this.props.project) {
                revision = "-";
              }
              else {
                revision = "0";
              }
              const logdata = {
                Title: this.state.documentid,
                Status: "Document Created",
                LogDate: this.state.today,
                Revision: revision,
                DocumentIndexId: parseInt(this.state.newDocumentId),
              }
              const log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata);
              // update document index
              if (this.state.directPublishCheck === false) {
                let diupdatedata = {
                  SourceDocumentID: parseInt(this.state.sourceDocumentId),
                  DocumentName: this.documentNameExtension,
                  SourceDocument: {
                    Description: this.documentNameExtension,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },
                  RevokeExpiry: {
                    Description: "Revoke",
                    Url: this.revokeUrl
                  }
                }
                const diupdate = await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, diupdatedata, parseInt(this.state.newDocumentId));

              }
              else {
                let indexdata = {
                  SourceDocumentID: parseInt(this.state.sourceDocumentId),
                  ApprovedDate: this.state.approvalDate,
                  DocumentName: this.documentNameExtension,
                  SourceDocument: {
                    Description: this.documentNameExtension,
                    Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                  },

                  RevokeExpiry: {
                    Description: "Revoke",
                    Url: this.revokeUrl
                  },
                }
                const diupdate = await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.state.newDocumentId));

              }
              await this._triggerPermission(sourceDocumentId);
              if (this.state.directPublishCheck === true) {
                this.setState({ hideLoading: false, hideCreateLoading: "none" });
                await this._publish();
              }
              else {
                this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                setTimeout(() => {
                  window.location.replace(this.props.siteUrl);
                }, 5000);
              }
            }
          }
        }
      }
      else if (this.state.templateId !== "") {
        let publishName: any;
        let extension: any;
        let newDocumentName;
        // Get template
        if (this.state.sourceId === "QDMS") {
          await this._Service.getqdmsselectLibraryItems(this.props.QDMSUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
            console.log(publishdoc);
            for (let i = 0; i < publishdoc.length; i++) {
              if (publishdoc[i].Id === this.state.templateId) {
                publishName = publishdoc[i].LinkFilename;
              }
            }
            var split = publishName.split(".", 2);
            extension = split[1];
          }).then(async cpysrc => {
            // Add template document to source document
            newDocumentName = this.state.documentName + "." + extension;
            this.documentNameExtension = newDocumentName;
            docinsertname = this.state.documentid + '.' + extension;

            let qdmsURl = this.props.QDMSUrl + "/" + this.props.publisheddocumentLibrary + "/" + this.state.category + "/" + publishName;
            await this._Service.getqdmsdocument(qdmsURl)
              .then((templateData: any) => {
                return this._Service.uploadDocument(this.props.sourceDocumentLibrary, docinsertname, templateData);
              }).then((fileUploaded: any) => {
                console.log("File Uploaded");
                fileUploaded.file.getItem()
                  .then(async (item: any) => {
                    console.log(item);
                    sourceDocumentId = item["ID"];

                    this.setState({ sourceDocumentId: sourceDocumentId });
                    await this._addSourceDocument();
                  }).then(async (updateDocumentIndex: any) => {
                    let revision;
                    if (this.props.project) {
                      revision = "-";
                    }
                    else {
                      revision = "0";
                    }
                    let logdata = {
                      Title: this.state.documentid,
                      Status: "Document Created",
                      LogDate: this.state.today,
                      Revision: revision,
                      DocumentIndexId: parseInt(this.state.newDocumentId),
                    }
                    let log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logdata)
                    if (this.state.directPublishCheck === false) {
                      let indexdata = {
                        SourceDocumentID: parseInt(this.state.sourceDocumentId),
                        DocumentName: this.documentNameExtension,
                        SourceDocument: {
                          Description: this.documentNameExtension,
                          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        },
                        RevokeExpiry: {
                          Description: "Revoke",
                          Url: this.revokeUrl
                        },
                      }
                      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.state.newDocumentId));

                    }
                    else {
                      let indexdata = {
                        SourceDocumentID: parseInt(this.state.sourceDocumentId),
                        DocumentName: this.documentNameExtension,
                        ApprovedDate: this.state.approvalDate,

                        SourceDocument: {
                          Description: this.documentNameExtension,
                          Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                        },
                        RevokeExpiry: {
                          Description: "Revoke",
                          Url: this.revokeUrl
                        },
                      }
                      await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.state.newDocumentId));

                    }
                    await this._triggerPermission(sourceDocumentId);
                    if (this.state.directPublishCheck === true) {
                      this.setState({ hideLoading: false, hideCreateLoading: "none" });
                      await this._publish();
                    }
                    else {
                      this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                      setTimeout(() => {
                        window.location.replace(this.props.siteUrl);
                      }, 5000);
                    }
                  });
              });
          });
        }
        else {
          await this._Service.getselectLibraryItems(this.props.siteUrl, this.props.publisheddocumentLibrary).then((publishdoc: any) => {
            console.log(publishdoc);
            for (let i = 0; i < publishdoc.length; i++) {
              if (publishdoc[i].Id === this.state.templateId) {
                publishName = publishdoc[i].LinkFilename;
              }
            }
            var split = publishName.split(".", 2);
            extension = split[1];
          }).then((cpysrc: any) => {
            // Add template document to source document
            newDocumentName = this.state.documentName + "." + extension;
            this.documentNameExtension = newDocumentName;
            docinsertname = this.state.documentid + '.' + extension;
            let siteUrl = this.props.siteUrl + "/" + this.props.publisheddocumentLibrary + "/" + publishName;
            this._Service.getDocument(siteUrl)
              .then((templateData: any) => {
                return this._Service.uploadDocument(this.props.sourceDocumentLibrary, docinsertname, templateData)
              }).then((fileUploaded: any) => {
                console.log("File Uploaded");
                fileUploaded.file.getItem().then(async (item: any) => {
                  console.log(item);
                  sourceDocumentId = item["ID"];
                  // if(extension === "pdf"||extension === "Pdf"||extension === "PDF"){
                  //   documenturl = item["ServerRedirectedEmbedUrl"];
                  // }
                  // else{
                  // docServerUrl = item["ServerRedirectedEmbedUrl"];
                  // splitdocUrl = docServerUrl.split("&", 2);
                  // documenturl = splitdocUrl[0];
                  // }
                  this.setState({ sourceDocumentId: sourceDocumentId });
                  await this._addSourceDocument();
                }).then(async (updateDocumentIndex: any) => {
                  let revision;
                  if (this.props.project) {
                    revision = "-";
                  }
                  else {
                    revision = "0";
                  }
                  let revlog = {
                    Title: this.state.documentid,
                    Status: "Document Created",
                    LogDate: this.state.today,
                    Revision: revision,
                    DocumentIndexId: parseInt(this.state.newDocumentId),
                  }
                  const log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, revlog);
                  if (this.state.directPublishCheck === false) {
                    let indexdata = {
                      SourceDocumentID: parseInt(this.state.sourceDocumentId),
                      DocumentName: this.documentNameExtension,
                      SourceDocument: {
                        Description: this.documentNameExtension,
                        Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                      },
                      RevokeExpiry: {
                        Description: "Revoke",
                        Url: this.revokeUrl
                      }
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.state.newDocumentId));
                  }
                  else {
                    let indexdata = {
                      SourceDocumentID: parseInt(this.state.sourceDocumentId),
                      DocumentName: this.documentNameExtension,
                      ApprovedDate: this.state.approvalDate,
                      SourceDocument: {
                        Description: this.documentNameExtension,
                        Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                      },
                      RevokeExpiry: {
                        Description: "Revoke",
                        Url: this.revokeUrl
                      },
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, indexdata, parseInt(this.state.newDocumentId));

                  }
                  await this._triggerPermission(sourceDocumentId);
                  if (this.state.directPublishCheck === true) {
                    this.setState({ hideLoading: false, hideCreateLoading: "none" });
                    await this._publish();
                  }
                  else {
                    this.setState({ hideCreateLoading: "none", norefresh: "none", statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 } });
                    setTimeout(() => {
                      window.location.replace(this.props.siteUrl);
                    }, 5000);
                  }
                });
              });
          });
        }
      }
      else { }
    }
    // without document
    else {
      await this._createDocumentIndex();

      this.setState({ statusMessage: { isShowMessage: true, message: this.createDocument, messageType: 4 }, norefresh: "none", hideCreateLoading: "none" });
      if (this.state.bulkDocumentIndex === " ") {
        if (this.Count === this.currentCount) {
          // alert("current count reached")
          window.location.replace(this.props.siteUrl);
        }
      }
      else {
        setTimeout(() => {
          window.location.replace(this.props.siteUrl);
        }, this.Timeout);
      }
    }
  }
  // Set permission for document
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
  public _documentCount = (ev: React.FormEvent<HTMLInputElement>, documentCount?: string) => {
    this.setState({ documentCount: documentCount || '' });

  }
  // Create Document Index
  public async _createDocumentIndex() {
    let documentIndexId;
    // Without Expiry date
    // if (this.state.expiryCheck === false) {
    if (this.state.expiryDate === null || this.state.expiryDate === undefined) {
      // DMS
      if (this.props.project) {
        let index1 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.state.documentName,
          BusinessUnitID: this.state.businessUnitID,
          BusinessUnit: this.state.businessUnit,
          CategoryID: this.state.categoryId,
          Category: this.state.category,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          Revision: "-",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          CreateDocument: this.state.createDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          DepartmentID: this.state.departmentId,
          PublishFormat: this.state.publishOption,
          TransmittalDocument: this.state.transmittalCheck,
          DirectPublish: this.state.directPublishCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber,
          IsLot: this.state.IsLot
        }
        await this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, index1)
          .then(async newdocid => {
            console.log(newdocid);
            documentIndexId = newdocid.data.ID;
            this.documentIndexID = newdocid.data.ID;
            this.setState({ newDocumentId: documentIndexId });
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
          });
      }
      // QDMS
      else {
        let index2 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.state.documentName,
          BusinessUnitID: this.state.businessUnitID,
          BusinessUnit: this.state.businessUnit,
          CategoryID: this.state.categoryId,
          Category: this.state.category,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          Revision: "0",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          CreateDocument: this.state.createDocument,
          DirectPublish: this.state.directPublishCheck,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          DepartmentID: this.state.departmentId,
          PublishFormat: this.state.publishOption,
          LegalEntity: this.state.legalEntity
        }
        await this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, index2)
          .then(async newdocid => {
            console.log(newdocid);
            this.documentIndexID = newdocid.data.ID;
            documentIndexId = newdocid.data.ID;
            this.setState({ newDocumentId: documentIndexId });
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
          });
      }
    }
    // With Expiry date
    else {
      // DMS
      if (this.props.project) {
        let index3 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.state.documentName,
          BusinessUnitID: this.state.businessUnitID,
          BusinessUnit: this.state.businessUnit,
          CategoryID: this.state.categoryId,
          Category: this.state.category,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          ExpiryDate: this.state.expiryDate,
          DirectPublish: this.state.directPublishCheck,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          Revision: "-",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          CreateDocument: this.state.createDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          DepartmentID: this.state.departmentId,
          PublishFormat: this.state.publishOption,
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber,
          IsLot: this.state.IsLot
        }
        await this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, index3)
          .then(async newdocid => {
            console.log(newdocid);
            this.documentIndexID = newdocid.data.ID;
            documentIndexId = newdocid.data.ID;
            this.setState({ newDocumentId: documentIndexId });
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
          });
      }
      // QDMS
      else {
        let index4 = {
          Title: this.state.title,
          DocumentID: this.state.documentid,
          ReviewersId: this.state.reviewers,
          DocumentName: this.state.documentName,
          BusinessUnitID: this.state.businessUnitID,
          BusinessUnit: this.state.businessUnit,
          CategoryID: this.state.categoryId,
          Category: this.state.category,
          SubCategoryID: this.state.subCategoryId,
          SubCategory: this.state.subCategory,
          ApproverId: this.state.approver,
          ExpiryDate: this.state.expiryDate,
          DirectPublish: this.state.directPublishCheck,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          Revision: "0",
          WorkflowStatus: "Draft",
          DocumentStatus: "Active",
          Template: this.state.templateDocument,
          CriticalDocument: this.state.criticalDocument,
          CreateDocument: this.state.createDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          DepartmentID: this.state.departmentId,
          PublishFormat: this.state.publishOption,
          LegalEntity: this.state.legalEntity
        }
        await this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, index4)
          .then(async newdocid => {
            console.log(newdocid);
            this.documentIndexID = newdocid.data.ID;
            documentIndexId = newdocid.data.ID;
            this.setState({ newDocumentId: documentIndexId });
            this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + newdocid.data.ID + "";
            this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + newdocid.data.ID + "&mode=expiry";
          });
      }
    }
  }
  // Add Source Document metadata
  public async _addSourceDocument() {
    // Without Expiry Date
    // if (this.state.expiryCheck === false) {
    if (this.state.expiryDate === null || this.state.expiryDate === undefined) {
      // DMS
      if (this.props.project) {
        let sourceitem1 = {
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
          DocumentIndexId: parseInt(this.state.newDocumentId),
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          },
          CriticalDocument: this.state.criticalDocument,
          TransmittalDocument: this.state.transmittalCheck,
          ExternalDocument: this.state.externalDocument,
          RevisionCodingId: this.state.revisionCodingId,
          RevisionLevelId: this.state.revisionLevelId,
          DocumentControllerId: this.state.dcc,
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourceitem1, parseInt(this.state.sourceDocumentId));
      }
      // QDMS
      else {
        let sourceitem2 = {
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
          DocumentIndexId: parseInt(this.state.newDocumentId),
          PublishFormat: this.state.publishOption,
          CriticalDocument: this.state.criticalDocument,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          }
        }
        await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourceitem2, parseInt(this.state.sourceDocumentId));

      }
    }
    // With Expiry Date
    else {
      // DMS
      if (this.props.project) {
        let sourceitem3 = {
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
          ExpiryDate: this.state.expiryDate,
          ExpiryLeadPeriod: this.state.expiryLeadPeriod,
          DocumentIndexId: parseInt(this.state.newDocumentId),
          PublishFormat: this.state.publishOption,
          CriticalDocument: this.state.criticalDocument,
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
          SubcontractorDocumentNo: this.state.subContractorNumber,
          CustomerDocumentNo: this.state.customerNumber
        }
        await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourceitem3, parseInt(this.state.sourceDocumentId));

      }
      // QDMS
      else {
        let sourceitem4 = {
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
          CriticalDocument: this.state.criticalDocument,
          DocumentIndexId: parseInt(this.state.newDocumentId),
          PublishFormat: this.state.publishOption,
          Template: this.state.templateDocument,
          OwnerId: this.state.owner,
          DepartmentName: this.state.department,
          RevisionHistory: {
            Description: "Revision History",
            Url: this.revisionHistoryUrl
          }
        }
        await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocumentLibrary, sourceitem4, parseInt(this.state.sourceDocumentId));

      }
    }
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
      'PublishedDate': this.state.today,
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
      this._publishUpdate();
    }
    else { }
  }
  // Published Document Metadata update
  public async _publishUpdate() {
    let updateindexdata = {
      PublishFormat: this.state.publishOption,
      WorkflowStatus: "Published",
      Revision: this.state.newRevision,
      ApprovedDate: new Date()
    }
    let updateindex = await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, updateindexdata, this.documentIndexID)


    if (this.state.owner !== this.state.currentUserId) {
      this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
    }
    if (this.props.project) {
      if (this.state.dcc !== this.state.owner) {
        this._sendMail(this.state.dccEmail, "DocPublish", this.state.dccName);
      }
    }
    let logitem = {
      Title: this.state.documentid,
      Status: "Published",
      LogDate: this.state.today,
      Revision: this.state.newRevision,
      DocumentIndexId: this.documentIndexID,
    }
    let log = await this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLogList, logitem)

    this.setState({ hideLoading: true, norefresh: "none", hideCreateLoading: "none", messageBar: "", statusMessage: { isShowMessage: true, message: this.directPublish, messageType: 4 } });
    setTimeout(() => {
      window.location.replace(this.props.siteUrl);
    }, 5000);
  }
  // Generate new revision
  public _generateNewRevision = async () => {
    let currentRevision = '-'; // set the current revisionsettings ID in state variable.
    this.setState({
      previousRevisionItemID: this.state.revisionCodingId // set this value with previous revision settings id from Project document index item.
    });
    // Reading current revision coding details from RevisionSettings.
    const revisionItem: any = await this._Service.getRevisionListItems(this.props.siteUrl, "RevisionSettings", parseInt(this.state.revisionCodingId));
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
  // qdms revision
  public _revisionCoding = async () => {
    let revision = parseInt("0");
    let rev = revision + 1;
    this.setState({ newRevision: rev.toString() });

  }
  //Send Mail
  public _sendMail = async (emailuser: any, type: any, name: any) => {
    let formatday = moment(this.state.today).format('DD/MMM/YYYY');
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
      mailSend = "Yes";
    }
    if (mailSend === "Yes") {
      const emailNotification: any[] = await this._Service.gethubListItems(this.props.hubUrl, this.props.emailNotification);
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title === type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }
      }
      let linkValue = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2";
      // link = `<a href=${window.location.protocol + "//" + window.location.hostname+this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.newDocumentId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"}>`+this.state.documentName+`</a>`;
      link = `<a href=${linkValue}>` + this.state.documentName + `</a>`;
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
  public async _onCreateIndex() {

    let counter = parseInt(this.state.documentCount);
    this.Count = counter;
    console.log(counter);
    this.Timeout = counter * 10000;
    if (this.validator.fieldValid('Title') && this.validator.fieldValid('category') && this.validator.fieldValid('BU/Dep') && this.validator.fieldValid('Owner') && this.validator.fieldValid('Approver')
      && this.validator.fieldValid('Revision') && this.validator.fieldValid('DocumentController') && this.validator.fieldValid('documentCount')) {
      this.setState({
        hideCreateLoading: " ",
        IsLot: true
      });
      for (let i = 1; i <= counter; i++) {
        let count = i;
        this.currentCount = count;
        this.setState({ counter: count.toString() });
        await this._documentidgeneration();
      }
      this.validator.hideMessages();
    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }

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
  //access denied msgbar close button click
  private _closeButton = () => {
    window.location.replace(this.props.siteUrl);
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
  };
  private modalProps = {
    isBlocking: true,
  };
  public render(): React.ReactElement<ITransmittalCreateDocumentProps> {
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
    const AddIcon: IIconProps = { iconName: 'Add' };
    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const uploadOrTemplateRadioBtnOptionQdms:
      IChoiceGroupOption[] = [
        { key: 'Upload', text: 'Upload existing file' },
        { key: 'Template', text: 'Create document existing template', styles: { field: { marginLeft: '3em' } } },
      ];
    const uploadOrTemplateRadioBtnOptions:
      IChoiceGroupOption[] = [
        { key: 'Upload', text: 'Upload existing file' },
        { key: 'Template', text: 'Create document existing template', styles: { field: { marginLeft: '18em' } } },
      ];
    const choiceGroupStyles: Partial<IChoiceGroupStyles> = { root: { display: 'flex' }, flexContainer: { display: "flex", justifyContent: 'space-between' } };

    return (
      <section className={`${styles.transmittalCreateDocument}`}>
        {/* Create Document QDMS */}
        <div style={{ display: this.state.createDocumentView }} >
          <div className={styles.border}>
            <div className={styles.alignCenter}>{this.props.webpartHeader}</div>
            <div>
              <TextField required id="t1"
                label="Title"
                onChange={this._titleChange}
                value={this.state.title} ></TextField>
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthfrst}>
                <Dropdown id="t3" label="Department"
                  selectedKey={this.state.departmentId}
                  placeholder="Select an option"
                  options={this.state.departmentOption}
                  onChanged={this._departmentChange} />
                <div style={{ color: "#dc3545", textAlign: "center" }}>
                  {this.validator.message("BU/Dep", this.state.departmentId, "required")}{""}
                </div>
              </div>
              <div className={styles.wdthmid}>
                <Dropdown id="t2" required={true} label="Category"
                  placeholder="Select an option"
                  selectedKey={this.state.categoryId}
                  options={this.state.categoryOption}
                  onChanged={this._categoryChange} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("category", this.state.categoryId, "required")}{" "}
                </div>
              </div>
              <div className={styles.wdthlst}>
                <Dropdown id="t2" label="Sub Category"
                  placeholder="Select an option"
                  selectedKey={this.state.subCategoryId}
                  options={this.state.subCategoryArray}
                  onChanged={this._subCategoryChange} /> </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthfrst}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Owner"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={this._selectedOwner}
                  defaultSelectedUsers={[this.state.ownerName]}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Owner", this.state.owner, "required")}{" "}</div>
              </div>
              <div className={styles.wdthmid}>
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
                  onChange={this._selectedApprover}
                  showHiddenInUI={false}
                  defaultSelectedUsers={[this.state.approverName]}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
                <div style={{ display: this.state.validApprover, color: "#dc3545" }}>Not able to change approver</div>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
              </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt} style={{ display: this.state.hideDoc }}>
                <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                  onChange={this.onUploadOrTemplateRadioBtnChange}
                  options={uploadOrTemplateRadioBtnOptionQdms} styles={choiceGroupStyles}
                /></div>
              <div className={styles.wdthlst} >
                <div style={{ display: this.state.hidetemplate }}>
                  <Dropdown id="t7"
                    label="Select a Template"
                    placeholder="Select an option"
                    selectedKey={this.state.templateId}
                    options={this.state.templateDocuments}
                    onChanged={this._templatechange} /></div>
                <div style={{ display: this.state.hideupload, marginTop: "2em" }}>
                  <input type="file" name="myFile" id="addqdms" onChange={this._add}></input>
                </div>
                <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
              </div>
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
            <div className={styles.divrow}>
              <div className={styles.wdthfrst} style={{ display: "flex" }}>
                <div style={{ width: "13em" }}> <DatePicker label="Expiry Date"
                  value={this.state.expiryDate}
                  onSelectDate={this._onExpDatePickerChange}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  minDate={new Date()}
                  formatDate={this._onFormatDate} /></div>
                {/* <div style={{ color: "#dc3545" }}>
                  {this.validator.message("expiryDate", this.state.expiryDate, "required")}{""}</div> */}
                <div style={{ marginLeft: "1em", width: "13em" }}>
                  <TextField id="ExpiryLeadPeriod" name="ExpiryLeadPeriod"
                    label="Expiry Reminder(Days)" onChange={this._expLeadPeriodChange}
                    value={this.state.expiryLeadPeriod}>
                  </TextField></div>
              </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthfrst} style={{ display: "flex" }}>
                <div style={{ width: "13em" }}> </div>
                <div style={{ marginLeft: "1em", width: "13em" }}>
                </div>
                {/* <div style={{ color: "#dc3545" }}>
                  {this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required")}{""}</div> */}
                <div style={{ color: "#dc3545", display: this.state.leadmsg }}>
                  Enter only numbers less than 100
                </div>
              </div>
            </div>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.mt}>
              <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Creating...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                <Label>***PLEASE DON'T REFRESH***</Label>
              </div>
            </div>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.mt}>
              <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Creating...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                <Label>***PLEASE DON'T REFRESH***</Label>
              </div>
            </div>
            <div className={styles.mandatory}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
            <DialogFooter>
              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onCreateDocument}>Submit</PrimaryButton >
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
            <br />
          </div>










        </div>
        {/* Create Document Project*/}
        <div style={{ display: this.state.createDocumentProject }} >
          <div className={styles.border}>
            <div >
              <div className={styles.alignCenter}>{this.props.webpartHeader}
                <div className={styles.rgtalign}>
                  {/* <IconButton iconProps={AddIcon} title="Create Multiple Index" label='Add Multiple' ariaLabel="Create Multiple Index" onClick={() => this._createMultipleIndex()} /> */}
                  {/* <ActionButton iconProps={AddIcon} onClick={() => this._createMultipleIndex()}>Create Multiple Documents </ActionButton> */}
                </div>
              </div>

            </div>
            <div>
              <TextField required id="t1"
                label="Title"
                onChange={this._titleChange}
                value={this.state.title} ></TextField>
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthfrst}>
                <Dropdown id="t3" label="Department"
                  selectedKey={this.state.departmentId}
                  placeholder="Select an option" required
                  options={this.state.departmentOption}
                  onChanged={this._departmentChange} />
                <div style={{ color: "#dc3545", textAlign: "center" }}>
                  {this.validator.message("BU/Dep", this.state.businessUnitID || this.state.departmentId, "required")}{""}
                </div>
              </div>
              <div className={styles.wdthmid}>
                <Dropdown id="t2" required={true} label="Category"
                  placeholder="Select an option"
                  selectedKey={this.state.categoryId}
                  options={this.state.categoryOption}
                  onChanged={this._categoryChange} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("category", this.state.categoryId, "required")}{" "}</div>
              </div>
              <div className={styles.wdthlst}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Owner"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._selectedOwner(items)}
                  defaultSelectedUsers={[this.props.context.pageContext.user.email]}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Owner", this.state.owner, "required")}{" "}</div>
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
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  showHiddenInUI={false}
                  defaultSelectedUsers={[this.state.dccName]}
                  onChange={(items) => this._selectedDCC(items)}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("DocumentController", this.state.dcc, "required")}{" "}</div>
              </div>
              <div className={styles.wdthmid}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Reviewer(s)"
                  personSelectionLimit={20}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  ensureUser={true}
                  showHiddenInUI={false}
                  onChange={(items) => this._selectedReviewers(items)}
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
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._selectedApprover(items)}
                  showHiddenInUI={false}
                  // defaultSelectedUsers={[this.state.approverName]}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />

                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
              </div>
            </div>
            <div className={styles.divrow}>
              <div style={{ display: this.state.hideDoc }}>
                <ChoiceGroup selectedKey={this.state.uploadOrTemplateRadioBtn}
                  onChange={this.onUploadOrTemplateRadioBtnChange}
                  options={uploadOrTemplateRadioBtnOptions} styles={choiceGroupStyles}
                /></div>

            </div>
            <div className={styles.divrow} style={{ display: this.state.hideupload, marginTop: "10px" }}>
              <div className={styles.wdthfrst}> <Label>Upload Document:</Label></div>
              <div className={styles.wdthmid}> <input type="file" name="myFile" id="addproject" onChange={this._add}></input></div>
              <div style={{ display: this.state.insertdocument, color: "#dc3545" }}>Please select valid Document or Please uncheck Create Document</div>
            </div>
            <div className={styles.divrow} style={{ display: this.state.hidetemplate }}>
              <div className={styles.wdthrgt}>
                <Dropdown id="t7"
                  label="Source"
                  placeholder="Select an option"
                  selectedKey={this.state.sourceId}
                  options={Source}
                  onChanged={this._sourcechange} /></div>
              <div className={styles.wdthlft}>
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
            <div className={styles.divrow} >
              <div className={styles.wdthfrst}>
                <Dropdown id="t2"
                  label="Revision Coding"
                  selectedKey={this.state.revisionCodingId}
                  placeholder="Select an option"
                  options={this.state.revisionSettingsArray}
                  onChanged={this._revisionCodingChange} />

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
                <div style={{ width: "13em" }}> <DatePicker label="Expiry Date"
                  value={this.state.expiryDate}
                  onSelectDate={this._onExpDatePickerChange}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  minDate={new Date()}
                  formatDate={this._onFormatDate} /></div>
                {/* <div style={{ color: "#dc3545" }}>
                  {this.validator.message("expiryDate", this.state.expiryDate, "required")}{""}</div> */}
                <div style={{ marginLeft: "1em", width: "13em" }}>
                  <TextField id="ExpiryLeadPeriod" name="ExpiryLeadPeriod"
                    label="Expiry Reminder(Days)" onChange={this._expLeadPeriodChange}
                    value={this.state.expiryLeadPeriod}>
                  </TextField></div>
                {/* <div style={{ color: "#dc3545" }}>
                  {this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required")}{""}</div> */}

              </div>

              <div className={styles.wdthmid} style={{ display: "flex" }}>
                <div style={{ marginTop: "3em" }}>
                  <TooltipHost
                    content="Check if the document is for transmittal"
                    //id={tooltipId}
                    calloutProps={calloutProps}
                    styles={hostStyles}>
                    <Checkbox label="Transmittal Document " boxSide="start"
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
                    <Checkbox label="External Document " boxSide="start"
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
                {/* <div style={{ color: "#dc3545" }}>
                  {this.validator.message("ExpiryLeadPeriod", this.state.expiryLeadPeriod, "required")}{""}</div> */}
                <div style={{ color: "#dc3545", display: this.state.leadmsg }}>
                  Enter only numbers less than 100
                </div>
              </div>
            </div>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.mt}>
              <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Creating...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                <Label>***PLEASE DON'T REFRESH***</Label>
              </div>
            </div>


            <div className={styles.mandatory}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
            <DialogFooter>
              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._onCreateDocument}>Submit</PrimaryButton >
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
            <br />
          </div>
        </div>


        {/* Create Multiple Document */}

        <div style={{ display: this.state.bulkDocumentIndex }} >
          <div className={styles.border}>
            <div className={styles.alignCenter}>{this.props.webpartHeader}</div>
            <div className={styles.divrow}>
              <div ><Label>Please enter document count to generate multiple indices :</Label></div>
              <div style={{ marginLeft: "20px" }}>
                <TextField required type='number' min={1}
                  onChange={this._documentCount}
                  value={this.state.documentCount} ></TextField>
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("documentCount", this.state.documentCount, "required")}{" "}</div>
              </div>
            </div>
            <div>
              <TextField required id="t1"
                label="Title"
                onChange={this._titleChange}
                value={this.state.title} ></TextField>
              <div style={{ color: "#dc3545" }}>
                {this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Dropdown id="t3" label="Business Unit"
                  selectedKey={this.state.businessUnitID}
                  placeholder="Select an option"
                  options={this.state.businessUnitOption}
                  onChanged={this._businessUnitChange}
                //  disabled
                />
              </div>
              <div className={styles.wdthlft}>
                <Dropdown id="t3" label="Department"
                  selectedKey={this.state.departmentId}
                  placeholder="Select an option"
                  options={this.state.departmentOption}
                  onChanged={this._departmentChange} />
              </div>
            </div>
            <div style={{ color: "#dc3545", textAlign: "center" }}>
              {this.validator.message("BU/Dep", this.state.businessUnitID || this.state.departmentId, "required")}{""}
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}><Dropdown id="t2" required={true} label="Category"
                placeholder="Select an option"
                selectedKey={this.state.categoryId}
                options={this.state.categoryOption}
                onChanged={this._categoryChange} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("category", this.state.categoryId, "required")}{" "}</div> </div>
              <div className={styles.wdthlft}>
                <Dropdown id="t2" label="Sub Category"
                  placeholder="Select an option"
                  selectedKey={this.state.subCategoryId}
                  options={this.state.subCategoryArray}
                  onChanged={this._subCategoryChange} /> </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}><PeoplePicker
                context={this.props.context as any}
                titleText="Owner"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}
                required={true}
                disabled={false}
                ensureUser={true}
                onChange={(items) => this._selectedOwner(items)}
                defaultSelectedUsers={[this.state.ownerName]}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Owner", this.state.owner, "required")}{" "}</div>
              </div>
              <div className={styles.wdthlft}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Document Controller"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={true}
                  disabled={false}
                  ensureUser={true}
                  showHiddenInUI={false}
                  defaultSelectedUsers={[this.state.dccName]}
                  onChange={(items) => this._selectedDCC(items)}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("DocumentController", this.state.dcc, "required")}{" "}</div>
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
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
            </div>
            <div className={styles.divrow} >
              <div className={styles.wdthrgt}>
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
                  // defaultSelectedUsers={[this.state.approverName]}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />

                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
              </div>
              <div className={styles.wdthlft}>
                <Dropdown id="t2"
                  label="Revision Coding"
                  selectedKey={this.state.revisionCodingId}
                  placeholder="Select an option"
                  options={this.state.revisionSettingsArray}
                  onChanged={this._revisionCodingChange} />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("Revision", this.state.revisionCodingId, "required")}{" "}</div>
              </div>

            </div>

            <div className={styles.divrow}>
              <div className={styles.wdthfrst} style={{ marginTop: "45px" }}>
                <div>
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
              </div>
              <div className={styles.wdthmid} style={{ marginTop: "45px" }}>
                <div>
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
            <div style={{ fontWeight: "bolder", color: "Red", marginTop: "50px" }}>
              NOTE : If you need to send for Transmittal Please check the boxes</div>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.mt}>
              <div hidden={this.state.hideLoading}><Spinner label={'Publishing...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.hideCreateLoading }}><Spinner label={'Creating...'} /></div>
            </div>
            <div className={styles.mt}>
              <div style={{ display: this.state.norefresh, color: "Red", fontWeight: "bolder", textAlign: "center" }}>
                <Label>***PLEASE DON'T REFRESH***</Label>
              </div>
            </div>
            <DialogFooter>

              <div className={styles.rgtalign}>
                <div className={styles.mandatory}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
              </div>
              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._onCreateIndex}>Create</PrimaryButton >
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
            <br />
          </div>
        </div >

        {/* Access Denied message bar*/}
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
      </section >
    );
  }
}
