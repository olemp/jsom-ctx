/// <reference types="sharepoint" />
export interface IJsomContext {
    url: string;
    isLoaded: boolean;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}
export declare class JsomContext implements IJsomContext {
    url: string;
    isLoaded: boolean;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
    constructor(url: string);
    load(): Promise<IJsomContext>;
}
export declare function ExecuteQuery(ctx: IJsomContext, clientObjectsToLoad?: SP.ClientObject[]): Promise<{
    sender: any;
    args: any;
}>;
