import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}

async function _getClientContext(url: string) {
    await SPComponentLoader.loadScript('/_layouts/15/SP.js', { globalExportsName: 'SP' });
    const clientContext = new SP.ClientContext(url);
    return clientContext;
}

export class JsomContext implements IJsomContext {
    public url: string;
    public clientContext: SP.ClientContext;
    public web: SP.Web;
    public site: SP.Site;
    public rootWeb: SP.Web;
    public lists: SP.ListCollection;
    public propBag: SP.FieldStringValues;

    constructor(url: string) {
        this.url = url;
    }

    public async load(): Promise<JsomContext> {
        try {
            this.clientContext = await _getClientContext(this.url);
            this.web = this.clientContext.get_web();
            this.site = this.clientContext.get_site();
            this.rootWeb = this.clientContext.get_site().get_rootWeb();
            this.lists = this.web.get_lists();
            this.propBag = this.web.get_allProperties();
            return this;
        } catch (err) {
            throw `Failed to load context for url ${this.url}`;
        }
    }
}

export async function CreateJsomContext(url: string): Promise<JsomContext> {
    let _ = new JsomContext(url);
    let jsomCtx = await _.load();
    return jsomCtx;
}

export function ExecuteJsomQuery(ctx: JsomContext, load: Array<{ clientObject: any, exps?: string }> = []) {
    return new Promise<{ sender, args }>((resolve, reject) => {
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
    });
}

