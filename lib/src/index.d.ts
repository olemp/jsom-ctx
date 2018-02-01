/// <reference types="sharepoint" />
export interface IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}
export declare class JsomContext implements IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    web: SP.Web;
    site: SP.Site;
    rootWeb: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
    constructor(url: string);
    load(): Promise<JsomContext>;
}
export declare function CreateJsomContext(url: string): Promise<JsomContext>;
export declare function ExecuteJsomQuery(ctx: JsomContext, load?: Array<{
    clientObject: any;
    exps?: string;
}>): Promise<{
    sender: any;
    args: any;
}>;
