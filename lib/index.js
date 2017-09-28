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
const JsomContext = {
    Context: class Context {
        constructor(url) {
            this.isLoaded = false;
            this.url = url;
        }
        load() {
            return __awaiter(this, void 0, void 0, function* () {
                this.clientContext = yield __getClientContext(this.url);
                this.web = this.clientContext.get_web();
                this.lists = this.web.get_lists();
                this.propBag = this.web.get_allProperties();
                return this;
            });
        }
    },
    ExecuteQuery: (ctx, clientObjectsToLoad = []) => {
        return new Promise((resolve, reject) => {
            if (ctx.isLoaded) {
                clientObjectsToLoad.forEach(clientObj => ctx.clientContext.load(clientObj));
                ctx.clientContext.executeQueryAsync((sender, args) => {
                    resolve({ sender, args });
                }, (sender, args) => {
                    reject({ sender, args });
                });
            }
            else {
                reject();
            }
        });
    }
};
exports.default = JsomContext;
