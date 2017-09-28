/// <reference types="sharepoint" />
export interface IJsomContext {
    url: string;
    isLoaded: boolean;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}
declare const JsomContext: {
    Context: {
        new (url: string): {
            url: string;
            isLoaded: boolean;
            clientContext: SP.ClientContext;
            web: SP.Web;
            lists: SP.ListCollection;
            propBag: SP.FieldStringValues;
            load(): Promise<IJsomContext>;
        };
    };
    ExecuteQuery: (ctx: IJsomContext, clientObjectsToLoad?: SP.ClientObject[]) => Promise<{
        sender: any;
        args: any;
    }>;
};
export default JsomContext;
