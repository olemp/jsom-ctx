"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
function __getClientContext(url) {
    return new Promise((resolve, reject) => {
        SP.SOD.executeFunc("sp.js", "SP.ClientContext", () => {
            const clientContext = new SP.ClientContext(url);
            resolve(clientContext);
        });
    });
}
class JsomContext {
    constructor(url) {
        this.url = url;
    }
    load() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                this.clientContext = yield __getClientContext(this.url);
                this.web = this.clientContext.get_web();
                this.site = this.clientContext.get_site();
                this.rootWeb = this.clientContext.get_site().get_rootWeb();
                this.lists = this.web.get_lists();
                this.propBag = this.web.get_allProperties();
                return this;
            }
            catch (err) {
                throw `Failed to load context for url ${this.url}`;
            }
        });
    }
}
exports.JsomContext = JsomContext;
function CreateJsomContext(url) {
    return __awaiter(this, void 0, void 0, function* () {
        let _ = new JsomContext(url);
        let jsomCtx = yield _.load();
        return jsomCtx;
    });
}
exports.CreateJsomContext = CreateJsomContext;
function ExecuteJsomQuery(ctx, load = []) {
    return new Promise((resolve, reject) => {
        load.forEach(l => {
            if (l.exps) {
                ctx.clientContext.load(l.clientObject, l.exps);
            }
            else {
                ctx.clientContext.load(l.clientObject);
            }
        });
        ctx.clientContext.executeQueryAsync((sender, args) => {
            resolve({ sender, args });
        }, (sender, args) => {
            reject({ sender, args });
        });
    });
}
exports.ExecuteJsomQuery = ExecuteJsomQuery;
