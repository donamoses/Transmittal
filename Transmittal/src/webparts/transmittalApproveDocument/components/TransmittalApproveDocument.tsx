import * as React from 'react';
import styles from './TransmittalApproveDocument.module.scss';
import { ITransmittalApproveDocumentProps, ITransmittalApproveDocumentState } from './ITransmittalApproveDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import SimpleReactValidator from 'simple-react-validator';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IconButton, IDropdownOption, IIconProps, Label, Link, MessageBar, PrimaryButton, ProgressIndicator, Spinner, TextField } from '@fluentui/react';
import { Accordion, AccordionItem, AccordionItemButton, AccordionItemHeading, AccordionItemPanel } from 'react-accessible-accordion';
import * as moment from 'moment';
import { BaseService } from '../services';
export default class TransmittalApproveDocument extends React.Component<ITransmittalApproveDocumentProps, ITransmittalApproveDocumentState, {}> {
  private validator: SimpleReactValidator;
  private _Service: BaseService;
  private workflowHeaderID: any;
  private documentIndexID: any;
  // private sourceDocumentID;
  private workflowDetailID: any;
  // private currentEmail;
  // private reqWeb;
  // private documentApprovedSuccess;
  private documentSavedAsDraft: string;
  // private documentRejectSuccess;
  // private documentReturnSuccess;
  private today = new Date();
  private revisionLogId: any;
  // private currentrevision;
  // private invalidApprovalLink;
  // private invalidUser;
  // private redirectUrlSuccess;
  // private redirectUrlError;
  // private valid;
  // private approverEmail;
  // private departmentExists;
  // private postUrl;
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
    // this._queryParamGetting = this._queryParamGetting.bind(this);
    // this._userMessageSettings = this._userMessageSettings.bind(this);
    // this._accessGroups = this._accessGroups.bind(this);
    // this._openRevisionHistory = this._openRevisionHistory.bind(this);
    // this._bindApprovalForm = this._bindApprovalForm.bind(this);
    // this._project = this._project.bind(this);
    // this._drpdwnPublishFormat = this._drpdwnPublishFormat.bind(this);
    // this._status = this._status.bind(this);
    // this._commentsChange = this._commentsChange.bind(this);
    // this._saveAsDraft = this._saveAsDraft.bind(this);
    // this._docSave = this._docSave.bind(this);
    // this._publish = this._publish.bind(this);
    // this._returnDoc = this._returnDoc.bind(this);
    // this._sendMail = this._sendMail.bind(this);
    // this._onCancel = this._onCancel.bind(this);
    // this._acceptanceChanged = this._acceptanceChanged.bind(this);
    // this._revisionCoding = this._revisionCoding.bind(this);
    // this._publishUpdate = this._publishUpdate.bind(this);
    // this._generateNewRevision = this._generateNewRevision.bind(this);
    // this._checkCurrentUser = this._checkCurrentUser.bind(this);
    // this._LAUrlGetting = this._LAUrlGetting.bind(this);
    // this._checkPermission = this._checkPermission.bind(this);
  }
  //Status Change
  public _status(option: { key: any; text: any }) {
    //console.log(option.key);
    if (option.key == 'Approved') {
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
    let commentadd = {
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

    if (this.state.hidepublish == false) {
      if (this.validator.fieldValid("publish") && (this.state.statusKey != "")) {
        this.validator.hideMessages();
        this.setState({ hideLoading: false, savedisable: "none" });
        let publishdata = {
          PublishFormat: this.state.publishOption
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workflowHeaderList, publishdata, this.workflowHeaderID);

        // this._publish();
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();

      }
    }
    else {
      if ((this.state.statusKey != "") && this.validator.fieldValid("comments")) {
        this.validator.hideMessages();
        let detaildata = {
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.today
        }
        await this._Service.updateItem(this.props.siteUrl, this.props.workflowDetailsList, detaildata, this.workflowDetailID)
        // await this._returnDoc().then((afterReturn: any) => {
        //   this.setState({ approveDocument: "" });
        //   setTimeout(() => this.setState({ approveDocument: 'none', hideLoading: true }), 3000);
        //   this.setState({ savedisable: "none" });
        // });
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
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
            <div style={{ display: "flex" }}>
              <div className={styles.width}><Label >Document ID : {this.state.documentID}</Label></div>
              <div><Link onClick={this._openRevisionHistory} target="_blank" underline>Revision History</Link></div>
            </div>

            <div >
              <Label >Document : <a href={this.state.linkToDoc} target="_blank">{this.state.documentName}</a></Label>
              <div hidden={this.state.hideProject}>
                <div className={styles.flex} >
                  <div className={styles.width}><Label >Project Name : {this.state.projectName}</Label></div>
                  <div ><Label>Project Number : {this.state.projectNumber} </Label></div>
                </div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Revision : {this.state.revision}</Label></div>
                {/* <div hidden={this.state.hideProject}><Label>Revision Level : {this.state.revisionLevel} </Label></div> */}
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
                <div><Label >Due Date : {this.state.dueDate}</Label></div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label>Requester : {this.state.requesterName} </Label></div>
                <div><Label >Requested Date : {this.state.requestedDate} </Label></div>
              </div>
              <div className={styles.flex}>
                <div><Label> Requester Comment : </Label>{this.state.requesterComment}</div>
              </div>
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
                        <div style={{ display: (this.state.dccreviewerData.length == 0 ? 'none' : 'block') }}>
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
                      <div style={{ display: (this.state.reviewerData.length == 0 ? 'none' : 'block') }}>
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
            <div >
              {/* <div className={styles.mt}>
                <div hidden={this.state.hideProject}>
                  <div className={styles.flex} >
                    <div className={styles.width}><Label >Transmittal Revision : {this.state.transmittalRevision}</Label></div>
                    <div ><Checkbox label="Publish For Transmittal " onChange={this._onPublishTransmittal} /></div>
                  </div>
                </div>
              </div> */}
              <div className={styles.mt}>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={status}
                  onChanged={this._status}
                  selectedKey={this.state.statusKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("Docstatus", this.state.statusKey, "required")}{" "}</div>
              </div>
              {/* <div style={{ color: "#dc3545" }}>{this.validator.message("Docstatus", this.state.statusKey, "required")}{" "}</div> */}
              <div className={styles.mt} hidden={this.state.hidepublish}>
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
              <div className={styles.mt} hidden={this.state.hideProject} >
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.acceptanceCodearray}
                    onChanged={this._acceptanceChanged}
                    selectedKey={this.state.acceptanceCode}
                  /></div></div>
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
              <DialogFooter>

                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <PrimaryButton id="b2" className={styles.btn} onClick={this._saveAsDraft} style={{ display: this.state.savedisable }}>Save as Draft</PrimaryButton >
                  <PrimaryButton id="b2" className={styles.btn} onClick={this._docSave} style={{ display: this.state.savedisable }}>Submit</PrimaryButton >
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
