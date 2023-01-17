import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import { add } from 'lodash';

export class BaseService {
    private _sp: SPFI;
    private sphub: SPFI;

    constructor(context: WebPartContext, huburl: string) {
        this._sp = getSP(context);
        this.sphub = new SPFI(huburl).using(SPFx(context));
    }

    public getListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubUserMessageListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items
            .select("Title,Message").filter("PageName eq 'Review'")()
    }
    public getLibraryItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items();
    }
    public getCurrentUser(): Promise<any> {
        return this._sp.web.currentUser();
    }
    public createNewItem(url: string, listname: string, data: any): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/Lists/" + listname).items.add(data);
    }
    public createhubNewItem(url: string, listname: string, data: any): Promise<any> {
        console.log(data);
        return this.sphub.web.getList(url + "/Lists/" + listname).items.add(data);
    }
    public updateItem(url: string, listname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public updatehubItem(url: string, listname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this.sphub.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public updateLibraryItem(url: string, libraryname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/" + libraryname).items.getById(id).update(data);
    }
    public uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        return this._sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
    }
    public getDocument(Url: string, publisheddocumentLibrary: string, publishName: string): Promise<any> {
        return this._sp.web.getFileByServerRelativePath(Url + "/" + publisheddocumentLibrary + "/" + publishName).getBuffer()
    }
    public getDrpdwnListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("Title,ID")()
    }
    public getRevisionListItems(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)
            .select("ID,StartPrefix,Pattern,StartWith,EndWith,MinN,MaxN,AutoIncrement")()
    }
    public getByEmail(email: string): Promise<any> {
        return this._sp.web.siteUsers.getByEmail(email)()
    }
    public getByhubEmail(email: string): Promise<any> {
        return this.sphub.web.siteUsers.getByEmail(email)()
    }
    public getByUserId(id: any): Promise<any> {
        return this.sphub.web.siteUsers.getById(id)()
    }
    public getHubsiteData(): Promise<any> {
        return this._sp.web.hubSiteData()
    }
    public getItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public gethubItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public getApproverData(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")()
    }
    public getIndexData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID)
            .select("DepartmentID,BusinessUnitID")();
    }
    public getIndexDataId(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID)
            .select("DocumentID,DocumentName,DepartmentID,BusinessUnitID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Owner,Approver,Reviewers")();
    }
    public getIndexProjectData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID)
            .select("RevisionCodingId,RevisionLevelId,TransmittalRevision,AcceptanceCodeId,DocumentController/ID,DocumentController/Title,DocumentController/EMail").expand("DocumentController")();
    }
    public getRevisionLevelData(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("ID,Title")()
    }
    public getSourceLibraryItems(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items
            .filter('DocumentIndexId eq ' + ID)()
    }
    public getpreviousheader(url: string, listname: string, IndexID: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items
            .select("ID").filter("DocumentIndex eq '" + IndexID + "' and(WorkflowStatus eq 'Returned with comments')")();
    }
    public gettriggerUnderReviewPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.filter("Title eq 'EMEC_DocumentPermission_UnderReview'")();
    }
    public gettriggerUnderApprovalPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.filter("Title eq 'EMEC_DocumentPermission_UnderApproval'")();
    }

    public getnotification(url: string, listname: string, emailuser: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference")()
    }
    public getemail(url: string, listname: string, type: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.filter("Title eq '" + type + "'")();
    }
    public gettaskdelegation(url: string, listname: string, Id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + Id + "' and(Status eq 'Active')")();
    }
    public deletehubItemById(url: string, listname: string, id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.getById(id).delete();
    }
    public getdetailresponsible(url: string, listname: string, Id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items
            .select("ResponseStatus")
            .filter("HeaderID eq " + Id + " and (Workflow eq 'Review')")()
    }
    public updateLibraryItemwithoutversion(url: string, libraryname: string, id: number, data: any,): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/" + libraryname).items.getById(id).validateUpdateListItem(data);
    }
    public getdccreviewlog(url: string, listname: string, headerId: number, documentIndexId: number,): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.filter("WorkflowID eq '" + headerId + "' and (DocumentIndexId eq '" + documentIndexId + "') and (Workflow eq 'DCC Review') and (Status eq 'Under Review')")()

    }
    public getreviewlog(url: string, listname: string, headerId: number, documentIndexId: number,): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.filter("WorkflowID eq '" + headerId + "' and (DocumentIndexId eq '" + documentIndexId + "') and (Workflow eq 'Review') and (Status eq 'Under Review')")()

    }
    public getdetail(url: string, listname: string, headerId: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
            .expand("Responsible").filter("HeaderID eq '" + headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review')")()
    }
    public getdetaildata(url: string, listname: string, headerId: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
            .expand("Responsible").filter("HeaderID eq '" + headerId + "' and (Workflow eq 'Review')")()
    }
    public getdetails(url: string, listname: string, headerId: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,ResponseDate,Workflow")
            .expand("Responsible").filter("HeaderID eq '" + headerId + "' and (Workflow eq 'Review' or Workflow eq 'DCC Review') and (ResponseStatus ne 'Under Review') ")()
    }
    public getdetaildatas(url: string, listname: string, detailID: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,Responsible/EMail,ResponsibleComment,ResponseStatus,TaskID")
            .expand("Responsible").filter("ID eq '" + detailID + "'")()
    }
    public getheaderdetails(url: string, listname: string, headerId: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Reviewers").getById(headerId)();
    }
    public getheaderdata(url: string, listname: string, headerId: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Reviewers").getById(headerId)();
    }
    public getprojectaccessgroup(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items
            .select("AccessGroups,AccessFields").filter("Title eq 'Project_SendReviewWF'")();
    }
    public getqdmsaccessgroup(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items
            .select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendReviewWF'")();
    }
    public getheaderdatas(url: string, listname: string, headerId: any, headerItems: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(headerId).select(headerItems).expand("Owner,Approver,Requester")();
    }
    public getprojectheaderdatas(url: string, listname: string, headerId: any, headerItems: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(headerId).select(headerItems)
            .expand("Owner,Approver,Requester,DocumentController")();
    }
    public getreviewComment(url: string, listname: string, previousReviewHeader: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate")
            .expand("Responsible").filter("HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'Review')  ")();
    }
    public getdccreviewComment(url: string, listname: string, previousReviewHeader: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.select("Responsible/ID,Responsible/Title,ResponsibleComment,ResponseDate,ResponsibleComment,ResponseDate").expand("Responsible")
            .filter("HeaderID eq '" + previousReviewHeader + "' and (Workflow eq 'DCC Review')  ")();
    }
    public indexbind(url: string, listname: string, documentIndexId: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(documentIndexId).select("CriticalDocument,DocumentName,SourceDocument")()
    }
} 