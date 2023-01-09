import * as React from 'react';
import styles from './TransmittalReviewDocument.module.scss';
import { ITransmittalReviewDocumentProps, ITransmittalReviewDocumentState } from './ITransmittalReviewDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IconButton, IDropdownOption, IIconProps, Label, Link, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, TextField } from '@fluentui/react';
import { Accordion, AccordionItem, AccordionItemButton, AccordionItemHeading, AccordionItemPanel } from 'react-accessible-accordion';
import * as moment from 'moment';
import * as strings from 'TransmittalReviewDocumentWebPartStrings';
import SimpleReactValidator from 'simple-react-validator';
import { BaseService } from '../services';
import { IHttpClientOptions, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
export default class TransmittalReviewDocument extends React.Component<ITransmittalReviewDocumentProps, ITransmittalReviewDocumentState, {}> {
  private validator: SimpleReactValidator;
  private _Service: BaseService;
  private headerId: number;
  private documentIndexId: any;
  private status: string;
  // private reqWeb ;
  private documentReviewedSuccess: string;
  // private documentSavedAsDraft;
  private detailID: number;
  private sourceDocumentID: number;
  private taskID: number;
  private newDetailItemID: number;
  private revisionLogID: number;
  private RevisionHistoryUrl: string;
  // private RedirectUrl;
  // private valid = "ok";
  // private noAccess;
  private currentDate = new Date();
  private workFlow: string;
  // private departmentExist;
  private postUrl: string;
  // private postUrlForUnderReview;
  // private postUrlForPermission;
  private dueDateWithoutConversion: any;
  // private postUrlForAdaptive;
  constructor(props: ITransmittalReviewDocumentProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      currentUser: null,
      status: "",
      statusKey: "",
      comments: "",
      reviewerItems: [],
      access: "none",
      accessDeniedMsgBar: "none",
      documentIndexItems: [],
      documentID: "",
      linkToDoc: "",
      documentName: "",
      revision: "",
      owner: "",
      requestor: "",
      requestorComment: "",
      dueDate: "",
      DueDate: null,
      requestorDate: "",
      workflowStatus: "",
      hideReviewersTable: "none",
      detailListID: null,
      cancelConfirmMsg: "none",
      confirmDialog: true,
      approverEmail: "",
      requestorEmail: "",
      documentControllerEmail: "",
      notificationPreference: "",
      headerListItem: [],
      approverName: "",
      approverId: "",
      ownerEmail: "",
      ownerID: "",
      reviewPending: "No",
      criticalDocument: false,
      currentUserEmail: "",
      userMessageSettings: [],
      invalidMessage: "",
      pageLoadItems: [],
      buttonHidden: "",
      detailIdForApprover: "",
      hubSiteUserId: "",
      delegatedToId: "",
      delegatedFromId: "",
      divForDCC: "none",
      divForReview: "none",
      ifDccComment: "none",
      dcc: "",
      dccComment: "",
      dccCompletionDate: "",
      revisionLogID: "",
      delegateToIdInSubSite: "",
      delegateForIdInSubSite: "",
      noAccess: "",
      invalidQueryParam: "",
      projectName: "",
      projectNumber: "",
      hideproject: true,
      reviewers: [],
      dccReviewItems: [],
      currentReviewComment: "none",
      currentReviewItems: [],
      loaderDisplay: "",
      documentControllerName: "",
      commentvalid: "none",
      commentrequired: false
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + this.props.hubSiteUrl);
    // this._drpdwnStatus = this._drpdwnStatus.bind(this);
    // this._onPageLoadDataBind = this._onPageLoadDataBind.bind(this);
    // this._currentUser = this._currentUser.bind(this);
    // this._loadPreviousReturnWithComments = this._loadPreviousReturnWithComments.bind(this);
    // this._docReviewSaveAsDraft = this._docReviewSaveAsDraft.bind(this);
    // this._docReviewSubmit = this._docReviewSubmit.bind(this);
    // this._cancel = this._cancel.bind(this);
    // this._confirmNoCancel = this._confirmNoCancel.bind(this);
    // this._confirmYesCancel = this._confirmYesCancel.bind(this);
    // this._sendAnEmailUsingMSGraph = this._sendAnEmailUsingMSGraph.bind(this);
    // this._checkingReviewStatus = this._checkingReviewStatus.bind(this);
    // this._returnWithComments = this._returnWithComments.bind(this);
    // this._userMessageSettings = this._userMessageSettings.bind(this);
    // this._queryParamGetting = this._queryParamGetting.bind(this);
    // this._documentIndexListBind = this._documentIndexListBind.bind(this);
    // this._docDCCReviewSubmit = this._docDCCReviewSubmit.bind(this);
    // this._revisionLogChecking = this._revisionLogChecking.bind(this);
    // this._accessGroups = this._accessGroups.bind(this);
    // this._projectInformation = this._projectInformation.bind(this);
    // this._checkingCurrent = this._checkingCurrent.bind(this);
    // this.GetGroupMembers = this.GetGroupMembers.bind(this);
    // this._gettingGroupID = this._gettingGroupID.bind(this);
    // this._LAUrlGetting = this._LAUrlGetting.bind(this);
    // this._LAUrlGettingForUnderReview = this._LAUrlGettingForUnderReview.bind(this);
    // this.triggerDocumentReview = this.triggerDocumentReview.bind(this);
    // this._LAUrlGettingForPermission = this._LAUrlGettingForPermission.bind(this);
    // this.triggerProjectPermissionFlow = this.triggerProjectPermissionFlow.bind(this);
    // this._sendAnEmailUsingMSGraphTest = this._sendAnEmailUsingMSGraphTest.bind(this);
  }
  public componentWillMount = async () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });
  }
  //Dropdown Status binding
  public _drpdwnStatus(option: { key: any; text: any }) {
    // alert(option.key);
    this.setState({ statusKey: option.key, status: option.text });
    if (option.key == "Returned with comments") {
      this.setState({ commentvalid: "", commentrequired: true });
    }
    else {
      this.setState({ commentvalid: "none", commentrequired: false });
    }
  }
  //Comment Box
  private _commentBoxChange = (ev: React.FormEvent<HTMLInputElement>, Comment?: string) => {
    this.setState({ comments: Comment || '' });
  }
  //submit
  private _docReviewSubmit = async () => {
    this._revisionLogChecking();
    console.log(this.revisionLogID);
    let reviewStatus: string;
    let count = 0;
    var today = new Date();
    let date = today.toLocaleString();
    let cancelCount = 0;
    //checking validation
    if (this.state.status == "Reviewed") {
      if (this.validator.fieldValid("status")) {
        const detaildata1 = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.currentDate,
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata1, this.detailID)
          .then(async deleteTask => {
            if (this.taskID != null) {
              let list = this._Service.deletehubItemById(this.props.hubSiteUrl, this.props.workflowTaskListName, this.taskID);
            }
          }).then(detailLIstUpdate => {
            this._Service.getdetailresponsible(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
              .then(async ResponseStatus => {
                if (ResponseStatus.length > 0) { //checking all reviewers response status
                  for (var k in ResponseStatus) {
                    if (ResponseStatus[k].ResponseStatus == "Reviewed") {
                      count++;
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Returned with comments") {
                      reviewStatus = "Returned with comments";
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Cancelled") {
                      cancelCount++;
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Under Review") {
                      this.setState({
                        reviewPending: "Yes",
                      });
                    }
                  }
                  //all reviewers reviewed
                  if (ResponseStatus.length == count || (ResponseStatus.length == add(count, cancelCount))) {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata1 = {
                      WorkflowStatus: "Under Approval",
                      Workflow: "Approval",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata1, this.headerId);
                    const indexdata1 = {
                      WorkflowStatus: "Under Approval",//docIndex
                      Workflow: "Approval"
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, indexdata1, this.documentIndexId);
                    //Updating DocumentRevisionlog 
                    const logdata1 = {
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLog, logdata1, this.state.revisionLogID);
                    const logdata2 = {
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }
                    this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata2);
                    //upadting source library without version change.            
                    let bodyArray = [
                      { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
                    ];
                    this._Service.updateLibraryItemwithoutversion(this.props.siteUrl, this.props.sourceDocument, this.sourceDocumentID, bodyArray);
                    //Task delegation getting user id from hubsite
                    this._Service.getByEmail(this.state.approverEmail).then(async user => {
                      console.log('User Id: ', user.Id);
                      this.setState({
                        hubSiteUserId: user.Id,
                      });
                      //Task delegation 
                      const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubSiteUrl, this.props.taskDelegationSettingsListName, user.Id);
                      if (taskDelegation.length > 0) {
                        let duedate = moment(this.dueDateWithoutConversion).toDate();
                        let toDate = moment(taskDelegation[0].ToDate).toDate();
                        let fromDate = moment(taskDelegation[0].FromDate).toDate();
                        duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                        toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                        fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                        if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                          this.setState({
                            approverEmail: taskDelegation[0].DelegatedTo.EMail,
                            approverName: taskDelegation[0].DelegatedTo.Title,
                            delegatedToId: taskDelegation[0].DelegatedTo.ID,
                            delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                          });
                          //duedate checking

                          //detail list adding an item for approval
                          this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail).then(async DelegatedTo => {
                            this.setState({
                              delegateToIdInSubSite: DelegatedTo.Id,
                            });
                            this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail).then(async DelegatedFor => {
                              this.setState({
                                delegateForIdInSubSite: DelegatedFor.Id,
                              });
                              const detaildata2 = {
                                HeaderIDId: Number(this.headerId),
                                Workflow: "Approval",
                                Title: this.state.documentName,
                                ResponsibleId: DelegatedTo.Id,
                                DueDate: this.state.DueDate,
                                DelegatedFromId: this.state.approverId,
                                ResponseStatus: "Under Approval",
                                SourceDocument: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName,
                                  Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                },
                                OwnerId: this.state.ownerID,
                              }
                              this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata2)
                                .then(async r => {
                                  this.setState({ detailIdForApprover: r.data.ID });
                                  this.newDetailItemID = r.data.ID;
                                  const detaildata3 = {
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    }
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata3, r.data.ID);
                                  const Approverdata1 = {
                                    ApproverId: DelegatedTo.Id
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, Approverdata1, this.headerId);
                                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, Approverdata1, this.documentIndexId);
                                  //upadting source library without version change.            
                                  this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocument, Approverdata1, this.sourceDocumentID);
                                  //MY tasks list updation
                                  const taskdata1 = {
                                    Title: "Approve '" + this.state.documentName + "'",
                                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                    DueDate: this.state.DueDate,
                                    StartDate: this.currentDate,
                                    AssignedToId: taskDelegation[0].DelegatedTo.ID,
                                    Workflow: "Approval",
                                    Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                                    DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                    Source: (this.props.project ? "Project" : "QDMS"),
                                    DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    },
                                  }
                                  await this._Service.createNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata1)
                                    .then(taskId => {
                                      this.taskID = taskId.data.ID;
                                      const taskdata2 = {
                                        TaskID: taskId.data.ID
                                      }
                                      this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, taskdata2, r.data.ID)
                                        .then(async aftermail => {
                                          this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");
                                          //this._adaptiveCard("Approval");
                                          if (!this.props.project) {
                                            // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                          }
                                          this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                          //Email pending  emailbody to approver                 
                                          this.validator.hideMessages();
                                          this.setState({
                                            comments: "",
                                            statusKey: "",
                                            approverEmail: "",
                                            approverName: "",
                                            approverId: "",
                                            buttonHidden: "none"
                                          });

                                        }).then(redirect => {
                                          setTimeout(() => {
                                            this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                                            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                            //this.RedirectUrl;
                                          }, 10000);

                                        });//aftermai
                                      //notification preference checking  

                                    });//taskID
                                });//r

                            });//DelegatedFor
                          });//DelegatedTo
                        }
                        else {
                          const headerdata2 = {
                            HeaderIDId: Number(this.headerId),
                            Workflow: "Approval",
                            Title: this.state.documentName,
                            ResponsibleId: this.state.approverId,
                            OwnerId: this.state.ownerID,
                            DueDate: this.state.DueDate,
                            ResponseStatus: "Under Approval",
                            SourceDocument: {
                              "__metadata": { type: "SP.FieldUrlValue" },
                              Description: this.state.documentName,
                              Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                            }
                          }
                          this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, headerdata2)
                            .then(async r => {
                              this.setState({ detailIdForApprover: r.data.ID });
                              this.newDetailItemID = r.data.ID;
                              const detaildata4 = {
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                }
                              }
                              this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata4, r.data.ID);

                              //MY tasks list updation
                              const taskdata3 = {
                                Title: "Approve '" + this.state.documentName + "'",
                                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                DueDate: this.state.DueDate,
                                StartDate: this.currentDate,
                                AssignedToId: user.Id,
                                Workflow: "Approval",
                                Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                                Source: (this.props.project ? "Project" : "QDMS"),
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                }
                              }
                              await this._Service.createhubNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata3)
                                .then(async taskId => {
                                  const taskdata4 = {
                                    TaskID: taskId.data.ID
                                  }
                                  await this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, taskdata4, r.data.ID)
                                    .then(aftermail => {
                                      this.validator.hideMessages();
                                      this.setState({
                                        statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                        comments: "",
                                        statusKey: "",
                                        approverEmail: "",
                                        approverName: "",
                                        approverId: "",
                                        buttonHidden: "none"
                                      });
                                      //Email pending  emailbody to approver  
                                      this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");

                                    }).then(redirect => {
                                      setTimeout(() => {
                                        this.setState({});
                                        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                        //this.RedirectUrl;
                                      }, 10000);

                                    });//aftermai
                                  //notification preference checking  
                                  this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                  if (!this.props.project) {
                                    // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                  }
                                });//taskID
                            });//r
                        }//else no delegation

                      }

                      else {
                        const detaildata5 = {
                          HeaderIDId: Number(this.headerId),
                          Workflow: "Approval",
                          Title: this.state.documentName,
                          ResponsibleId: this.state.approverId,
                          DueDate: this.state.DueDate,
                          OwnerId: Number(this.state.ownerID),
                          ResponseStatus: "Under Approval",
                          SourceDocument: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: this.state.documentName,
                            Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                          }
                        }
                        this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata5)
                          .then(async r => {
                            this.setState({ detailIdForApprover: r.data.ID });
                            this.newDetailItemID = r.data.ID;
                            const detaildata6 = {
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName + "-- Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              }
                            }
                            this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata6, r.data.ID);

                            //MY tasks list updation
                            const taskdata5 = {
                              Title: "Approve '" + this.state.documentName + "'",
                              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                              DueDate: this.state.DueDate,
                              StartDate: this.currentDate,
                              AssignedToId: user.Id,
                              Workflow: "Approval",
                              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                              Source: (this.props.project ? "Project" : "QDMS"),
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName + "-- Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              }
                            }
                            await this._Service.createhubNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata5)
                              .then(async taskId => {
                                this.taskID = taskId.data.ID;
                                const taskdata6 = {
                                  TaskID: taskId.data.ID
                                }
                                this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, taskdata6, r.data.ID)
                                  .then(async aftermail => {
                                    this.validator.hideMessages();
                                    this.setState({
                                      statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                      comments: "",
                                      statusKey: "",
                                      approverEmail: "",
                                      approverName: "",
                                      approverId: "",
                                      buttonHidden: "none"
                                    });
                                    //Email pending  emailbody to approver  
                                    this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");


                                  }).then(redirect => {
                                    setTimeout(() => {
                                      this.setState({});
                                      window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                      //this.RedirectUrl;
                                    }, 10000);

                                  });//aftermai
                                //notification preference checking  
                                this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                if (!this.props.project) {
                                  // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                }

                              });//taskID
                          });//r
                      }//else no delegation

                    }).catch(reject => console.error('Error getting Id of user by Email ', reject));
                  }
                  //any of the reviewer returned with comments
                  else if (reviewStatus == "Returned with comments" && this.state.reviewPending == "No") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata3 = {
                      WorkflowStatus: "Returned with comments",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata3, this.headerId);
                    const indexdata2 = {
                      WorkflowStatus: "Returned with comments",
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, indexdata2, this.documentIndexId);

                    //Updationg DocumentRevisionlog   
                    const logdata3 = {
                      Status: "Returned with comments",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLog, logdata3, this.revisionLogID)
                      .then(afterHeaderStatusUpdate => {
                        this.triggerDocumentReview(this.sourceDocumentID, "Returned with comments");
                        this._returnWithComments();
                        //mail to document controller if any one reviewer return with comments.
                        if (this.props.project) { this._sendAnEmailUsingMSGraph(this.state.documentControllerEmail, "DocReturn", this.state.documentControllerName, this.newDetailItemID); }
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none",
                        });
                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        }, 10000);

                      });
                  }
                  //if any review process pending
                  else if (this.state.reviewPending == "Yes") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata4 = {
                      WorkflowStatus: "Under Review"
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata4, this.headerId)
                      .then(async after => {
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none"
                        });

                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                          // this.RedirectUrl;
                        }, 10000);

                      });
                  }
                }
              });
          });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      if (this.validator.fieldValid("status") && this.validator.fieldValid("reviewercomment")) {
        const detaildata7 = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.currentDate
        }
        this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata7, this.detailID)
          .then(async deleteTask => {
            if (this.taskID != null) {
              this._Service.deletehubItemById(this.props.hubSiteUrl, this.props.workflowTaskListName, this.taskID);
            }
          }).then(detailLIstUpdate => {
            this._Service.getdetailresponsible(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
              .then(async ResponseStatus => {
                if (ResponseStatus.length > 0) { //checking all reviewers response status
                  for (var k in ResponseStatus) {
                    if (ResponseStatus[k].ResponseStatus == "Reviewed") {
                      count++;
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Returned with comments") {
                      reviewStatus = "Returned with comments";
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Cancelled") {
                      cancelCount++;
                    }
                    else if (ResponseStatus[k].ResponseStatus == "Under Review") {
                      this.setState({
                        reviewPending: "Yes",
                      });
                    }
                  }
                  //all reviewers reviewed
                  if (ResponseStatus.length == count || (ResponseStatus.length == add(count, cancelCount))) {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata5 = {
                      WorkflowStatus: "Under Approval",
                      Workflow: "Approval",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata5, this.headerId);
                    const indexdata3 = {
                      WorkflowStatus: "Under Approval",//docIndex
                      Workflow: "Approval",
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, indexdata3, this.documentIndexId);
                    //Updationg DocumentRevisionlog 
                    const logdata4 = {
                      Status: "Reviewed",
                      LogDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLog, logdata4, this.state.revisionLogID)
                    const logdata5 = {
                      Status: "Under Approval",
                      LogDate: this.currentDate,
                      WorkflowID: this.headerId,
                      DocumentIndexId: this.documentIndexId,
                      DueDate: this.state.DueDate,
                      Workflow: "Approval",
                      Revision: this.state.revision,
                      Title: this.state.documentID,
                    }
                    this._Service.createNewItem(this.props.siteUrl, this.props.documentRevisionLog, logdata5);
                    //upadting source library without version change.            
                    let bodyArray = [
                      { "FieldName": "WorkflowStatus", "FieldValue": "Under Approval" }, { "FieldName": "Workflow", "FieldValue": "Approval" }
                    ];
                    this._Service.updateLibraryItemwithoutversion(this.props.siteUrl, this.props.sourceDocument, this.sourceDocumentID, bodyArray);
                    //Task delegation getting user id from hubsite
                    this._Service.getByhubEmail(this.state.approverEmail).then(async user => {
                      console.log('User Id: ', user.Id);
                      this.setState({
                        hubSiteUserId: user.Id,
                      });
                      //Task delegation 
                      const taskDelegation: any[] = await this._Service.gettaskdelegation(this.props.hubSiteUrl, this.props.taskDelegationSettingsListName, user.Id);
                      console.log(taskDelegation);
                      if (taskDelegation.length > 0) {
                        let duedate = moment(this.dueDateWithoutConversion).toDate();
                        let toDate = moment(taskDelegation[0].ToDate).toDate();
                        let fromDate = moment(taskDelegation[0].FromDate).toDate();
                        duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                        toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                        fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                        if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                          this.setState({
                            approverEmail: taskDelegation[0].DelegatedTo.EMail,
                            approverName: taskDelegation[0].DelegatedTo.Title,
                            delegatedToId: taskDelegation[0].DelegatedTo.ID,
                            delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                          });
                          //duedate checking

                          //detail list adding an item for approval
                          this._Service.getByEmail(taskDelegation[0].DelegatedTo.EMail).then(async DelegatedTo => {
                            this.setState({
                              delegateToIdInSubSite: DelegatedTo.Id,
                            });
                            this._Service.getByEmail(taskDelegation[0].DelegatedFor.EMail).then(async DelegatedFor => {
                              this.setState({
                                delegateForIdInSubSite: DelegatedFor.Id,
                              });
                              const detaildata8 = {
                                HeaderIDId: Number(this.headerId),
                                Workflow: "Approval",
                                Title: this.state.documentName,
                                ResponsibleId: DelegatedTo.Id,
                                DueDate: this.state.DueDate,
                                DelegatedFromId: this.state.approverId,
                                ResponseStatus: "Under Approval",
                                SourceDocument: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName,
                                  Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                                },
                                OwnerId: this.state.ownerID,
                              }
                              this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata8)
                                .then(async r => {
                                  this.setState({ detailIdForApprover: r.data.ID });
                                  this.newDetailItemID = r.data.ID;
                                  const detaildata9 = {
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    }
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata9, r.data.ID);
                                  const Approverdata3 = {
                                    ApproverId: DelegatedTo.Id
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, Approverdata3, this.headerId);
                                  this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, Approverdata3, this.documentIndexId);

                                  //upadting source library without version change.            
                                  await this._Service.updateLibraryItem(this.props.siteUrl, this.props.sourceDocument, Approverdata3, this.sourceDocumentID);
                                  //MY tasks list updation
                                  const taskdata7 = {
                                    Title: "Approve '" + this.state.documentName + "'",
                                    Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                    DueDate: this.state.DueDate,
                                    StartDate: this.currentDate,
                                    AssignedToId: taskDelegation[0].DelegatedTo.ID,
                                    Workflow: "Approval",
                                    Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                                    DelegatedOn: (this.state.delegatedToId !== "" ? this.currentDate : " "),
                                    Source: (this.props.project ? "Project" : "QDMS"),
                                    DelegatedFromId: taskDelegation[0].DelegatedFor.ID,
                                    Link: {
                                      "__metadata": { type: "SP.FieldUrlValue" },
                                      Description: this.state.documentName + "-- Approve",
                                      Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                    }
                                  }
                                  await this._Service.createhubNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata7)
                                    .then(taskId => {
                                      this.taskID = taskId.data.ID;
                                      const detaildata10 = {
                                        TaskID: taskId.data.ID,
                                      }
                                      this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata10, r.data.ID)
                                        .then(async aftermail => {
                                          this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");
                                          //this._adaptiveCard("Approval");
                                          if (!this.props.project) {
                                            // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                          }
                                          this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                          //Email pending  emailbody to approver                 
                                          this.validator.hideMessages();
                                          this.setState({
                                            comments: "",
                                            statusKey: "",
                                            approverEmail: "",
                                            approverName: "",
                                            approverId: "",
                                            buttonHidden: "none"
                                          });

                                        }).then(redirect => {
                                          setTimeout(() => {
                                            this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                                            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                            //this.RedirectUrl;
                                          }, 10000);

                                        });//aftermai
                                      //notification preference checking  

                                    });//taskID
                                });//r

                            });//DelegatedFor
                          });//DelegatedTo
                        }
                        else {
                          const headerdata6 = {
                            HeaderIDId: Number(this.headerId),
                            Workflow: "Approval",
                            Title: this.state.documentName,
                            ResponsibleId: this.state.approverId,
                            OwnerId: this.state.ownerID,
                            DueDate: this.state.DueDate,
                            ResponseStatus: "Under Approval",
                            SourceDocument: {
                              "__metadata": { type: "SP.FieldUrlValue" },
                              Description: this.state.documentName,
                              Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                            }
                          }
                          this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, headerdata6)
                            .then(async r => {
                              this.setState({ detailIdForApprover: r.data.ID });
                              this.newDetailItemID = r.data.ID;
                              const detaildata11 = {
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                }
                              }
                              this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata11, r.data.ID);

                              //MY tasks list updation
                              const taskdata8 = {
                                Title: "Approve '" + this.state.documentName + "'",
                                Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                                DueDate: this.state.DueDate,
                                StartDate: this.currentDate,
                                AssignedToId: user.Id,
                                Workflow: "Approval",
                                Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                                Source: (this.props.project ? "Project" : "QDMS"),
                                Link: {
                                  "__metadata": { type: "SP.FieldUrlValue" },
                                  Description: this.state.documentName + "-- Approve",
                                  Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                                },

                              }
                              await this._Service.createhubNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata8)
                                .then(async taskId => {
                                  const taskdata9 = {
                                    TaskID: taskId.data.ID
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, taskdata9, r.data.ID)
                                    .then(aftermail => {
                                      this.validator.hideMessages();
                                      this.setState({
                                        statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                        comments: "",
                                        statusKey: "",
                                        approverEmail: "",
                                        approverName: "",
                                        approverId: "",
                                        buttonHidden: "none"
                                      });
                                      //Email pending  emailbody to approver  
                                      this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");

                                    }).then(redirect => {
                                      setTimeout(() => {
                                        this.setState({});
                                        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                        //this.RedirectUrl;
                                      }, 10000);

                                    });//aftermai
                                  //notification preference checking  
                                  this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                  if (!this.props.project) {
                                    // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                  }
                                });//taskID
                            });//r
                        }//else no delegation

                      }

                      else {
                        const detaildata12 = {
                          HeaderIDId: Number(this.headerId),
                          Workflow: "Approval",
                          Title: this.state.documentName,
                          ResponsibleId: this.state.approverId,
                          DueDate: this.state.DueDate,
                          OwnerId: Number(this.state.ownerID),
                          ResponseStatus: "Under Approval",
                          SourceDocument: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: this.state.documentName,
                            Url: this.props.siteUrl + "/SourceDocuments/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                          },
                        }
                        this._Service.createNewItem(this.props.siteUrl, this.props.workFlowDetail, detaildata12)
                          .then(async r => {
                            this.setState({ detailIdForApprover: r.data.ID });
                            this.newDetailItemID = r.data.ID;
                            const detaildata13 = {
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName + "-- Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              }
                            }
                            this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata13, r.data.ID);
                            //MY tasks list updation
                            const taskdata10 = {
                              Title: "Approve '" + this.state.documentName + "'",
                              Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.requestor + "' on '" + this.state.requestorDate + "'",
                              DueDate: this.state.DueDate,
                              StartDate: this.currentDate,
                              AssignedToId: user.Id,
                              Workflow: "Approval",
                              Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                              Source: (this.props.project ? "Project" : "QDMS"),
                              Link: {
                                "__metadata": { type: "SP.FieldUrlValue" },
                                Description: this.state.documentName + "-- Approve",
                                Url: this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + r.data.ID + ""
                              }
                            }
                            await this._Service.createhubNewItem(this.props.hubSiteUrl, this.props.workflowTaskListName, taskdata10)
                              .then(async taskId => {
                                this.taskID = taskId.data.ID;
                                const detaildata14 = {
                                  TaskID: taskId.data.ID
                                }
                                this._Service.updateItem(this.props.siteUrl, this.props.workFlowDetail, detaildata14, r.data.ID)
                                  .then(async aftermail => {
                                    this.validator.hideMessages();
                                    this.setState({
                                      statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                                      comments: "",
                                      statusKey: "",
                                      approverEmail: "",
                                      approverName: "",
                                      approverId: "",
                                      buttonHidden: "none"
                                    });
                                    //Email pending  emailbody to approver  
                                    this.triggerDocumentReview(this.sourceDocumentID, "Under Approval");


                                  }).then(redirect => {
                                    setTimeout(() => {
                                      this.setState({});
                                      window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                                      //this.RedirectUrl;
                                    }, 10000);

                                  });//aftermai
                                //notification preference checking  
                                this._sendAnEmailUsingMSGraph(this.state.approverEmail, "DocApproval", this.state.approverName, this.newDetailItemID);
                                if (!this.props.project) {
                                  // await this._adaptiveCard("Approval", this.state.approverEmail, this.state.approverName, "General");
                                }

                              });//taskID
                          });//r
                      }//else no delegation

                    }).catch(reject => console.error('Error getting Id of user by Email ', reject));
                  }
                  //any of the reviewer returned with comments
                  else if (reviewStatus == "Returned with comments" && this.state.reviewPending == "No") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata7 = {
                      WorkflowStatus: "Returned with comments",
                      ReviewedDate: this.currentDate,
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata7, this.headerId);
                    const indexdata4 = {
                      WorkflowStatus: "Returned with comments"
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.documentIndex, indexdata4, this.documentIndexId);
                    const logdata6 = {
                      Status: "Returned with comments",
                      LogDate: this.currentDate,
                    }
                    //Updationg DocumentRevisionlog                
                    this._Service.updateItem(this.props.siteUrl, this.props.documentRevisionLog, logdata6, this.revisionLogID)
                      .then(afterHeaderStatusUpdate => {
                        this.triggerDocumentReview(this.sourceDocumentID, "Returned with comments");
                        this._returnWithComments();
                        //mail to document controller if any one reviewer return with comments.
                        if (this.props.project) { this._sendAnEmailUsingMSGraph(this.state.documentControllerEmail, "DocReturn", this.state.documentControllerName, this.newDetailItemID); }
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none",
                        });
                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                        }, 10000);

                      });
                  }
                  //if any review process pending
                  else if (this.state.reviewPending == "Yes") {
                    this.setState({
                      buttonHidden: "none",
                    });
                    const headerdata8 = {
                      WorkflowStatus: "Under Review"
                    }
                    this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderListName, headerdata8, this.headerId)
                      .then(async after => {
                        this.validator.hideMessages();
                        this.setState({
                          statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 },
                          comments: "",
                          statusKey: "",
                          buttonHidden: "none"
                        });

                      }).then(redirect => {
                        setTimeout(() => {
                          this.setState({ statusMessage: { isShowMessage: true, message: this.documentReviewedSuccess, messageType: 4 }, });
                          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
                          // this.RedirectUrl;
                        }, 10000);

                      });
                  }
                }
              });
          });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  private _revisionLogChecking() {
    var today = new Date();
    let date = today.toLocaleString();
    //Updationg DocumentRevisionlog
    if (this.props.project && this.workFlow == "dcc") {
      this._Service.getdccreviewlog(this.props.siteUrl, this.props.documentRevisionLog, this.headerId, this.documentIndexId)
        .then(ifyes => {
          if (ifyes.length > 0) {
            this.revisionLogID = ifyes[0].ID;
            console.log(ifyes[0].ID);
            this.setState({
              revisionLogID: ifyes[0].ID,
            });
          }
        });
    }
    else {
      this._Service.getreviewlog(this.props.siteUrl, this.props.documentRevisionLog, this.headerId, this.documentIndexId)
        .then(ifyes => {
          if (ifyes.length > 0) {
            this.revisionLogID = ifyes[0].ID;
            console.log(ifyes[0].ID);
            this.setState({
              revisionLogID: ifyes[0].ID,
            });
          }

        });
    }
  }
  protected async triggerDocumentReview(sourceDocumentID: any, ResponseStatus: any) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;

    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID,
      'WorkflowStatus': ResponseStatus
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  private async _sendAnEmailUsingMSGraph(email: any, type: any, name: any, detailID: any): Promise<void> {
    let Subject;
    let Body;
    let link;
    let tableHeader;
    let tableFooter;
    let tableBody = "";
    let finalBody;
    let DocumentLink;
    //console.log(queryVar);
    const notificationPreference: any[] = await this._Service.getnotification(this.props.hubSiteUrl, this.props.notificationPrefListName, email);
    // console.log(notificationPreference);
    if (notificationPreference.length > 0) {
      this.setState({
        notificationPreference: notificationPreference[0].Preference,
      });
    }
    else if (this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await this._Service.getemail(this.props.hubSiteUrl, this.props.emailNotificationSettings, type);
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;

    if (type == "DocApproval") {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentApprovalSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;
      //for binding current reviewers comments in table
      if (this.props.project) {
        await this._Service.getdetail(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
              });
              currentReviewersItems.map((item: any) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;v">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>
 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
          });
      }
      else {
        await this._Service.getdetaildata(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
              });
              currentReviewersItems.map((item: any) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>
 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
          });
      }
    }
    else if (type == "DocReview") {
      link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/" + this.props.documentReviewSitePage + ".aspx?hid=" + this.headerId + "&dtlid=" + detailID}>Link</a>`;
      if (this.props.project) {
        await this._Service.getdetails(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
              });
              currentReviewersItems.map((item: any) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>
 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
          });
      }
    }
    //returned with comments mail body
    else if (type == "DocReturn") {
      if (this.props.project) {
        await this._Service.getdetail(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
                linkToDoc: currentReviewersItems[0].SourceDocument.Url,
              });
              currentReviewersItems.map((item: any) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td><td>" + item.Workflow + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>
 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
            DocumentLink = `<a href=${this.state.linkToDoc}>Document Link </a>`;
          });
      }
      else {
        await this._Service.getdetaildata(this.props.siteUrl, this.props.workFlowDetail, this.headerId)
          .then(currentReviewersItems => {
            console.log("currentReviewersItems", currentReviewersItems);
            if (currentReviewersItems.length > 0) {
              console.log("currentReviewersItems", currentReviewersItems);
              this.setState({
                currentReviewComment: "",
                currentReviewItems: currentReviewersItems,
                linkToDoc: currentReviewersItems[0].SourceDocument.Url,
              });
              currentReviewersItems.map((item: any) => {
                tableBody += "<tr><td>" + item.Responsible.Title + "</td><td>" + moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a') + "</td><td>" + item.ResponseStatus + "</td><td>" + item.ResponsibleComment + "</td></tr>";
              });
            }
          }).then(after => {
            tableHeader = `<table style=" border: 1px solid #ddd;width: 100%;border-collapse: collapse;text-align: center;">
   <tr  style="background-color: #002d71;     color: white;text-align: center;">
   <th >Reviewer</th>
   <th >Review Date</th>
   <th >Response Status</th>
   <th >Review Comment</th>
   <th >Workflow</th>

 </tr>
 <tbody style ="width: 100%;
 border-collapse: collapse;border: 2px solid #ddd";text-align: center;>`;

            tableFooter = `</tbody>
 </table>`;
            finalBody = tableHeader + tableBody + tableFooter;
            DocumentLink = `<a href=${this.state.linkToDoc}>Click here </a>`;
          });
      }
    }

    //Replacing the email body with current values
    let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
    let replacedSubjectWithDueDate = replaceString(replacedSubject, '[DueDate]', this.state.dueDate);
    let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);
    let replacelink = replaceString(replaceBody, '[Link]', link);
    let var1: any[] = replacelink.split('/');
    let FinalBody = replacelink;
    if (this.state.notificationPreference == "Send all emails") {
      this.status = "Yes";
      //console.log("Send mail for all");                 
    }
    else if (this.state.notificationPreference == "Send mail for critical document" && this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      this.status = "Yes";
    }
    else {
      this.setState({
        statusMessage: { isShowMessage: true, message: strings.DocumentReviewedMsgBar, messageType: 4 },
        comments: "",
        statusKey: "",
      });
    }
    //mail sending
    if (this.status == "Yes") {
      //Check if TextField value is empty or not  
      if (email) {
        //Create Body for Email  
        let emailPostBody: any = {
          "message": {
            "subject": replacedSubjectWithDueDate,
            "body": {
              "contentType": "HTML",
              "content": FinalBody + "<br></br>" + (type == "DocReturn" ? DocumentLink : "") + "<br></br>" + finalBody
            },
            "toRecipients": [
              {
                "emailAddress": {
                  "address": email
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
              .post(emailPostBody, (error: any, response: any, rawResponse?: any) => {
              });
          });
      }
    }
  }
  //Cancel button click
  private _cancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });
    this.validator.hideMessages();
  }
  //confirm cancel button click
  private _confirmYesCancel = () => {
    window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });

    this.validator.hideMessages();

  }
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
  }
  //access denied msgbar close button click
  private _closeButton = () => {
    window.location.replace(this.props.redirectUrl);
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
  public render(): React.ReactElement<ITransmittalReviewDocumentProps> {
    const DownIcon: IIconProps = { iconName: 'ChevronDown' };
    const Status: IDropdownOption[] = [
      { key: 'Reviewed', text: 'Reviewed' },
      { key: 'Returned with comments', text: 'Returned with comments' },
    ];
    return (
      <section className={`${styles.transmittalReviewDocument}`}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div style={{ display: this.state.access }}>
          {/*For Review Webpart view */}
          <div style={{ border: "1px solid black" }}>
            <div style={{ display: this.state.divForReview, padding: "9px 14px 0px 13px" }}>
              <div className={styles.title}> {escape(this.props.webPartName)} </div>
              <div hidden={this.state.hideproject}>
                <div className={styles.flex} style={{ marginTop: "10px" }}>
                  <div className={styles.width}><Label >Project Name : {this.state.projectName} </Label></div>
                  <div className={styles.width}><Label >Project Number : {this.state.projectNumber}</Label></div>
                </div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width} style={{ fontWeight: "bold" }}>Document ID :{this.state.documentID}</div>
                <div className={styles.width}>
                  <Link underline href={this.RevisionHistoryUrl} target="_blank" > Revision History </Link>
                </div>
              </div>
              <div className={styles.width}>
                <Label >Document: <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></Label>
              </div>
              <div className={styles.innerRow}>
                <Label>Revision: {this.state.revision}</Label>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}>  <Label>Owner :{this.state.owner}</Label> </div>
                <div className={styles.width}><Label>Due Date :{this.state.dueDate}</Label> </div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}> <Label>Requester :{this.state.requestor}</Label></div>
                <div><Label>Requested Date :{this.state.requestorDate}</Label> </div>
              </div>
              <div className={styles.innerRow}>
                <Label>Requester Comment: </Label>{this.state.requestorComment}
              </div>
              <div className={styles.innerRow} style={{ display: this.state.hideReviewersTable }}>
                <Accordion allowZeroExpanded className={styles.Accordion}>
                  <AccordionItem >
                    <AccordionItemHeading>
                      <AccordionItemButton className={styles.AccordionItemButton}>
                        <Label className={styles.pleft}><IconButton iconProps={DownIcon} />Previous Review Details</Label>
                      </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <table className={styles.tableClass}>
                        <tr className={styles.tr}>
                          <th className={styles.th}>Reviewer</th>
                          <th className={styles.th}>Review Date</th>
                          <th className={styles.th}>Review Comment</th>
                        </tr>
                        <tbody className={styles.tbody}>
                          {this.state.reviewerItems.map((item, key) => {
                            return (<tr className={styles.tr}>
                              <td className={styles.th}>{item.Responsible.Title}</td>
                              <td className={styles.th}>{moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                              <td className={styles.th}>{item.ResponsibleComment}</td>
                            </tr>);
                          })}

                        </tbody>
                      </table>
                    </AccordionItemPanel>
                  </AccordionItem>
                </Accordion>

              </div>
              <div className={styles.innerRow} style={{ display: this.state.ifDccComment }}>
                <Accordion allowZeroExpanded className={styles.Accordion}>
                  <AccordionItem >
                    <AccordionItemHeading>
                      <AccordionItemButton className={styles.AccordionItemButton}>
                        <Label className={styles.pleft}><IconButton iconProps={DownIcon} />Document Controller Review Details</Label>
                      </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <table className={styles.tableClass}>
                        <tr className={styles.tr}>
                          <th className={styles.th}>Document Controller</th>
                          <th className={styles.th}>DCC Date</th>
                          <th className={styles.th}>DCC Comment</th>
                        </tr>
                        <tbody className={styles.tbody}>
                          {this.state.dccReviewItems.map((item, key) => {
                            return (<tr className={styles.tr}>
                              <td className={styles.th}>{item.Responsible.Title}</td>
                              <td className={styles.th}>{moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                              <td className={styles.th}>{item.ResponsibleComment}</td>
                            </tr>);
                          })}
                        </tbody>
                      </table>
                    </AccordionItemPanel>
                  </AccordionItem>
                </Accordion>
              </div>
              <div className={styles.innerRow} style={{ display: this.state.currentReviewComment }}>
                <Accordion allowZeroExpanded className={styles.Accordion}>
                  <AccordionItem >
                    <AccordionItemHeading>
                      <AccordionItemButton className={styles.AccordionItemButton}>
                        <Label className={styles.pleft}><IconButton iconProps={DownIcon} />Reviewers Details</Label>
                      </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <table className={styles.tableClass}>
                        <tr className={styles.tr}>
                          <th className={styles.th}>Reviewer</th>
                          <th className={styles.th}>Review Date</th>
                          <th className={styles.th}>Response Status</th>
                          <th className={styles.th}>Review Comment</th>
                        </tr>
                        <tbody className={styles.tbody}>
                          {this.state.currentReviewItems.map((item, key) => {
                            return (<tr className={styles.tr}>
                              <td className={styles.th}>{item.Responsible.Title}</td>
                              <td className={styles.th}>{(item.ResponseDate == null) ? "Not Reviewed Yet" : moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                              <td className={styles.th}>{item.ResponseStatus}</td>
                              <td className={styles.th}>{item.ResponsibleComment}</td>
                            </tr>);
                          })}
                        </tbody>
                      </table>
                    </AccordionItemPanel>
                  </AccordionItem>
                </Accordion>
              </div>
              <div>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={Status}
                  onChanged={this._drpdwnStatus}
                  selectedKey={this.state.statusKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("status", this.state.statusKey, "required")}{" "}</div>
              </div>
              <TextField label="Comments" id="Comments" value={this.state.comments} onChange={this._commentBoxChange} multiline autoAdjustHeight required={this.state.commentrequired} />
              <div style={{ display: this.state.commentvalid }}>
                <div style={{ color: "#dc3545" }}>{this.validator.message("reviewercomment", this.state.comments, "required")}{" "}</div></div>

              <DialogFooter>
                {/* Show Message bar for Notification*/}
                {this.state.statusMessage.isShowMessage ?
                  <MessageBar
                    messageBarType={this.state.statusMessage.messageType}
                    isMultiline={false}
                    dismissButtonAriaLabel="Close"
                  >{this.state.statusMessage.message}</MessageBar>
                  : ''}
                <table style={{ float: "right", rowGap: "0px" }}>
                  <tr>
                    <td style={{ display: "flex", padding: "0 0 0 33rem" }}>
                      <Label style={{ color: "red", fontSize: "23px" }}>*</Label>
                      <label style={{ fontStyle: "italic", fontSize: "12px" }}>fields are mandatory </label>
                    </td>

                    <PrimaryButton id="b1" style={{ float: "right", borderRadius: "10px", border: "1px solid gray" }} onClick={this._cancel}>Cancel</PrimaryButton>
                    <div style={{ display: this.state.buttonHidden }}>
                      <PrimaryButton id="b2" style={{ float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }} onClick={this._docReviewSubmit}>Submit</PrimaryButton>
                      <PrimaryButton id="b2" style={{ float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }} onClick={this._docReviewSaveAsDraft}>Save as Draft</PrimaryButton>
                    </div>
                  </tr>
                </table>
              </DialogFooter>
              {/* Cancel Dialog Box */}
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
          {/*For DCC Webpart view */}
          <div style={{ border: "1px solid black" }}>
            <div style={{ display: this.state.divForDCC, padding: "9px 14px 0px 13px" }}>
              <div className={styles.digitalReview}>
                <div className={styles.title}> {escape(this.props.webPartName)} </div>
                <div hidden={this.state.hideproject}>
                  <div className={styles.flex}>
                    <div className={styles.width}><Label >Project Name : {this.state.projectName} </Label></div>
                    <div className={styles.width}><Label >Project Number : {this.state.projectNumber}</Label></div>
                  </div>
                </div>
                <div></div>
                <div className={styles.flex} style={{ marginTop: "10px" }}>
                  <div className={styles.width} style={{ fontWeight: "bold", }}>Document ID :{this.state.documentID}</div>
                  <div className={styles.width}>
                    <Link underline href={this.RevisionHistoryUrl} target="_blank" > Revision History </Link>
                  </div>
                </div>

                <div className={styles.innerRow1}>
                  <Label>Document: <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></Label>
                </div>

                <div className={styles.innerRow}>
                  <Label>Revision: {this.state.revision}</Label>
                </div>
                <div className={styles.flex}>
                  <div className={styles.width}>  <Label>Owner :{this.state.owner}</Label> </div>
                  <div className={styles.width}><Label>Due Date :{this.state.dueDate}</Label> </div>
                </div>
                <div className={styles.flex}>
                  <div className={styles.width}> <Label>Requester :{this.state.requestor}</Label></div>
                  <div><Label>Requested Date :{this.state.requestorDate}</Label> </div>
                </div>
                <div className={styles.innerRow}>
                  <Label>Requester Comment: </Label>{this.state.requestorComment}
                </div>
                {/* <div className={styles.innerRow} style={{ display: this.state.currentReviewComment }}>
                      <table className={styles.tableClass}>
                        <tr className={styles.tr}>
                          <th className={styles.th}>Reviewer</th>
                          <th className={styles.th}>Review Date</th>
                          <th className={styles.th}>Review Comment</th>
                        </tr>
                        <tbody className={styles.tbody}>
                        {this.state.currentReviewItems.map((item, key) => {
                            return (<tr className={styles.tr}>
                            <td className={styles.th}>{item.Responsible.Title}</td>
                              <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                              <td className={styles.th}>{item.ResponsibleComment}</td>
                            </tr>);
                             })}
                        </tbody>
                      </table>
                    </div> */}
                <div className={styles.innerRow} style={{ display: this.state.hideReviewersTable }}>
                  <table className={styles.tableClass}>
                    <tr className={styles.tr}>
                      <th className={styles.th}>Reviewer</th>
                      <th className={styles.th}>Review Date</th>
                      <th className={styles.th}>Review Comment</th>
                    </tr>
                    <tbody className={styles.tbody}>
                      {this.state.reviewerItems.map((item, key) => {
                        return (<tr className={styles.tr}>
                          <td className={styles.th}>{item.Responsible.Title}</td>
                          <td className={styles.th}>{moment(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                          <td className={styles.th}>{item.ResponsibleComment}</td>
                        </tr>);
                      })}

                    </tbody>
                  </table>
                </div>
                <div>
                  <Dropdown
                    placeholder="Select Status"
                    label="Status"
                    options={Status}
                    onChanged={this._drpdwnStatus}
                    selectedKey={this.state.statusKey}
                    required />
                  <div style={{ color: "#dc3545" }}>{this.validator.message("status", this.state.statusKey, "required")}{" "}</div>
                </div>
                <TextField label="Comments" id="Comments" value={this.state.comments} onChange={this._commentBoxChange} multiline autoAdjustHeight required={this.state.commentrequired} />
                <div style={{ display: this.state.commentvalid }}>  <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div></div>
              </div>
              <DialogFooter>
                {/* Show Message bar for Notification*/}
                {this.state.statusMessage.isShowMessage ?
                  <MessageBar
                    messageBarType={this.state.statusMessage.messageType}
                    isMultiline={false}
                    dismissButtonAriaLabel="Close"
                  >{this.state.statusMessage.message}</MessageBar>
                  : ''}
                <table style={{ float: "right", rowGap: "0px" }}>
                  <tr>
                    <td style={{ display: "flex", padding: "0 0 0 33rem" }}>
                      <Label style={{ color: "red", fontSize: "23px" }}>*</Label>
                      <label style={{ fontStyle: "italic", fontSize: "12px" }}>fields are mandatory </label>
                    </td>

                    <PrimaryButton id="b1" style={{ float: "right", borderRadius: "10px", border: "1px solid gray" }} onClick={this._cancel}>Cancel</PrimaryButton>
                    <div style={{ display: this.state.buttonHidden }}>
                      <PrimaryButton id="b2" style={{ float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }} onClick={this._docDCCReviewSubmit}>Submit</PrimaryButton>
                      <PrimaryButton id="b2" style={{ float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }} onClick={this._docReviewSaveAsDraft}>Save as Draft</PrimaryButton>
                    </div>
                  </tr>
                </table>
              </DialogFooter>
              {/* Cancel Dialog Box */}
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
        </div>
        <div style={{ display: this.state.accessDeniedMsgBar }}>
          <MessageBar messageBarType={MessageBarType.error} onDismiss={this._closeButton} isMultiline={false}> {this.state.invalidMessage}</MessageBar>
        </div>
      </section>
    );
  }
}
