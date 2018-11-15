"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var JsomContext = /** @class */ (function () {
    /**
     * Constructor
     *
     * @param {string} url URL (defaults to current)
     * @param {string} appContextSiteUrl App context site URL
     */
    function JsomContext(url, appContextSiteUrl) {
        this.url = url;
        this.appContextSiteUrl = appContextSiteUrl;
        this.clientContext = url ? new SP.ClientContext(url) : SP.ClientContext.get_current();
    }
    JsomContext.prototype.load = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
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
                    return [2 /*return*/, this];
                }
                catch (err) {
                    throw "Failed to load context for url " + this.url;
                }
                return [2 /*return*/];
            });
        });
    };
    return JsomContext;
}());
exports.JsomContext = JsomContext;
/**
 * Creates a JSOM context object
 *
 * @param {string} url URL (defaults to current)
 * @param {string} appContextSiteUrl App context site URL
 */
function CreateJsomContext(url, appContextSiteUrl) {
    return __awaiter(this, void 0, void 0, function () {
        var _, jsomCtx;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _ = new JsomContext(url, appContextSiteUrl);
                    return [4 /*yield*/, _.load()];
                case 1:
                    jsomCtx = _a.sent();
                    return [2 /*return*/, jsomCtx];
            }
        });
    });
}
exports.CreateJsomContext = CreateJsomContext;
function ExecuteJsomQuery(ctx, load) {
    if (load === void 0) { load = []; }
    return new Promise(function (resolve, reject) {
        try {
            load.forEach(function (l) {
                if (l.exps) {
                    ctx.clientContext.load(l.clientObject, l.exps);
                }
                else {
                    ctx.clientContext.load(l.clientObject);
                }
            });
            ctx.clientContext.executeQueryAsync(function (sender, args) {
                resolve({ sender: sender, args: args });
            }, function (sender, args) {
                reject({ sender: sender, args: args });
            });
        }
        catch (error) {
            reject(error);
        }
    });
}
exports.ExecuteJsomQuery = ExecuteJsomQuery;
//# sourceMappingURL=index.js.map