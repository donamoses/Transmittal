import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";

export class BaseService {
    private _sp: SPFI;
    private spQdms: SPFI;
    private sphub: SPFI;

    constructor(context: WebPartContext, qdmsURL: string, huburl: string) {
        this._sp = getSP(context);
        this.spQdms = new SPFI(qdmsURL).using(SPFx(context));
        this.sphub = new SPFI(huburl).using(SPFx(context));
    }

    public getListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubUserMessageListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'DocumentIndex'")()
    }
    public getLibraryItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items();
    }
    public getqdmsLibraryItems(url: string, listname: string): Promise<any> {
        return this.spQdms.web.getList(url + "/" + listname).items();
    }
    public getselectLibraryItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.select("LinkFilename,ID,Template,DocumentName")();
    }
    public getqdmsselectLibraryItems(url: string, listname: string): Promise<any> {
        return this.spQdms.web.getList(url + "/" + listname).items.select("LinkFilename,ID")();
    }

    public getCurrentUser() {
        return this._sp.web.currentUser();
    }
    public createNewItem(url: string, listname: string, data: any): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/Lists/" + listname).items.add(data);
    }

    public updateItem(url: string, listname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }

    public uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        return this._sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
    }
    public getDocument(Url: string): Promise<any> {
        return this._sp.web.getFileByServerRelativePath(Url).getBuffer()
    }
    public getqdmsdocument(Url: string): Promise<any> {
        return this.spQdms.web.getFileByServerRelativePath(Url).getBuffer()
    }
    public updateLibraryItem(url: string, libraryname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/" + libraryname).items.getById(id).update(data);
    }
    public getDrpdwnListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("Title,ID")()
    }
    public getRevisionListItems(url: string, listname: string, id: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).select("ID,StartPrefix,Pattern,StartWith,EndWith,MinN,MaxN,AutoIncrement")()
    }
    public getByEmail(email: string): Promise<any> {
        return this._sp.web.siteUsers.getByEmail(email)()
    }
    public getHubsiteData(): Promise<any> {
        return this._sp.web.hubSiteData()
    }
    public getItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public gethubItemById(url: string, listname: string, id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public getBusinessUnitItem(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")()
    }
    public gettriggerPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname)
            .items.filter("Title eq 'EMEC_DocumentPermission-Create Document'")()
    }
    public getpublish(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPublish'")()
    }
    public getdirectpublish(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_PermissionWebpart'")()
    }
    public getnotification(url: string, listname: string, emailuser: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference")()
    }
    public getpublishlibrary(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/" + listname)
            .items.select("ID").filter("DocumentIndex/ID eq '" + id + "'")();
    }
    public getListSourceItem(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(id).select("SourceDocument")();
    }
    public getIndexdata(url: string, listname: string, documentindexid: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(documentindexid)
            .select("DocumentStatus,SourceDocumentID,TransmittalStatus,WorkflowStatus")()
    }
    public getIndexdataa(url: string, listname: string, documentindexid: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname)
            .items.getById(documentindexid)
            .select("DocumentStatus,SourceDocumentID,WorkflowStatus")()
    }
} 