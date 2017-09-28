export interface IJsomContext {
    url: string;
    isLoaded: boolean;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}

function __getClientContext(url: string) {
    return new Promise<SP.ClientContext>((resolve, reject) => {
        SP.SOD.executeFunc("sp.js", "SP.ClientContext", () => {
            const clientContext = new SP.ClientContext(url);
            resolve(clientContext);
        });
    });
}

export class JsomContext implements IJsomContext {
    public url: string;
    public isLoaded: boolean;
    public clientContext: SP.ClientContext;
    public web: SP.Web;
    public lists: SP.ListCollection;
    public propBag: SP.FieldStringValues;

    constructor(url: string) {
        this.isLoaded = false;
        this.url = url;
    }

    public async load(): Promise<IJsomContext> {
        this.clientContext = await __getClientContext(this.url);
        this.web = this.clientContext.get_web();
        this.lists = this.web.get_lists();
        this.propBag = this.web.get_allProperties();
        return this;
    }
}

export function ExecuteJsomQuery(ctx: IJsomContext, clientObjectsToLoad: SP.ClientObject[] = []) {
    return new Promise<{ sender, args }>((resolve, reject) => {
        if (ctx.isLoaded) {
            clientObjectsToLoad.forEach(clientObj => ctx.clientContext.load(clientObj));
            ctx.clientContext.executeQueryAsync((sender, args) => {
                resolve({ sender, args });
            }, (sender, args) => {
                reject({ sender, args });
            })
        } else {
            reject();
        }
    });
}

