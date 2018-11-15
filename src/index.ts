export interface IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    appContextSite?: SP.AppContextSite;
    appContextSiteUrl?: string;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}

export class JsomContext implements IJsomContext {
    public url: string;
    public clientContext: SP.ClientContext;
    public appContextSite: SP.AppContextSite;
    public appContextSiteUrl: string;
    public web: SP.Web;
    public site: SP.Site;
    public rootWeb: SP.Web;
    public lists: SP.ListCollection;
    public propBag: SP.FieldStringValues;

    /**
     * Constructor
     * 
     * @param {string} url URL (defaults to current)
     * @param {string} appContextSiteUrl App context site URL
     */
    public constructor(url?: string, appContextSiteUrl?: string) {
        this.url = url;
        this.appContextSiteUrl = appContextSiteUrl;
        this.clientContext = url ? new SP.ClientContext(url) : SP.ClientContext.get_current();
    }

    public async load(): Promise<JsomContext> {
        try {
            this.web = this.clientContext.get_web();
            this.site = this.clientContext.get_site();
            if (this.appContextSiteUrl) {
                this.appContextSite = new SP.AppContextSite(this.clientContext, this.appContextSiteUrl);
                this.web = this.appContextSite.get_web();
                this.site = this.appContextSite.get_site();
            }
            this.rootWeb = this.site.get_rootWeb();
            this.lists = this.web.get_lists();
            this.propBag = this.web.get_allProperties();
            return this;
        } catch (err) {
            throw `Failed to load context for url ${this.url}`;
        }
    }
}

/**
 * Creates a JSOM context object
 * 
 * @param {string} url URL (defaults to current)
 * @param {string} appContextSiteUrl App context site URL
 */
export async function CreateJsomContext(url?: string, appContextSiteUrl?: string): Promise<JsomContext> {
    let _ = new JsomContext(url, appContextSiteUrl);
    let jsomCtx = await _.load();
    return jsomCtx;
}

export interface IJsomLoadObject {
    clientObject: any;
    exps?: string;
}

export function ExecuteJsomQuery(ctx: JsomContext, load: Array<IJsomLoadObject> = []) {
    return new Promise<{ sender, args }>((resolve, reject) => {
        try {
            load.forEach(l => {
                if (l.exps) {
                    ctx.clientContext.load(l.clientObject, l.exps);
                } else {
                    ctx.clientContext.load(l.clientObject);
                }
            });
            ctx.clientContext.executeQueryAsync((sender, args) => {
                resolve({ sender, args });
            }, (sender, args) => {
                reject({ sender, args });
            })
        } catch (error) {
            reject(error);
        }
    });
}

