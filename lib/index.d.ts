/// <reference types="sharepoint" />
export interface IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    appContextSite?: SP.AppContextSite;
    appContextSiteUrl?: string;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}
export declare class JsomContext implements IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    appContextSite: SP.AppContextSite;
    appContextSiteUrl: string;
    web: SP.Web;
    site: SP.Site;
    rootWeb: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
    /**
     * Constructor
     *
     * @param {string} url URL
     * @param {string} appContextSiteUrl App context site URL
     */
    constructor(url: string, appContextSiteUrl?: string);
    load(): Promise<JsomContext>;
}
/**
 * Creates a JSOM context object
 *
 * @param {string} url URL
 * @param {string} appContextSiteUrl App context site URL
 */
export declare function CreateJsomContext(url: string, appContextSiteUrl?: string): Promise<JsomContext>;
export declare function ExecuteJsomQuery(ctx: JsomContext, load?: Array<{
    clientObject: any;
    exps?: string;
}>): Promise<{
    sender: any;
    args: any;
}>;
