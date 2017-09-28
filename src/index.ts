export interface IJsomContext {
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

const JsomContext = {
    Context: class Context implements IJsomContext {
        private url: string;
        public isLoaded: boolean;
        clientContext: SP.ClientContext;
        web: SP.Web;
        lists: SP.ListCollection;
        propBag: SP.FieldStringValues;

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
    },
    ExecuteQuery: (ctx: IJsomContext, clientObjectsToLoad: SP.ClientObject[] = []) => {
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
}

export default JsomContext;
