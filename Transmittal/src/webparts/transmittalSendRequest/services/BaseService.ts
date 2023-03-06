import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";

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
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'SendRequest'")()
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
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).select("ID,StartPrefix,Pattern,StartWith,EndWith,MinN,MaxN,AutoIncrement")()
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
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")()
    }
    public getIndexData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID).select("WorkflowStatus,SourceDocument,DocumentStatus")();
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
        return this._sp.web.getList(url + "/" + listname).items.filter('DocumentIndexId eq ' + ID)()
    }
    public getpreviousheader(url: string, listname: string, IndexID: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("ID").filter("DocumentIndex eq '" + IndexID + "' and(WorkflowStatus eq 'Returned with comments')")();
    }
    public gettriggerUnderReviewPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPermission_UnderReview'")()
    }
    public gettriggerUnderApprovalPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPermission_UnderApproval'")()
    }
    public getdirectpublish(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_PermissionWebpart'")()
    }
    public getnotification(url: string, listname: string, emailuser: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference")()
    }
    public getemail(url: string, listname: string, type: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq '" + type + "'")();
    }
    public gettaskdelegation(url: string, listname: string, Id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + Id + "' and(Status eq 'Active')")();
    }
} 