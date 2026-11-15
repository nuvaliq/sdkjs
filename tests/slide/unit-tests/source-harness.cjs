'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '../../..');
function source(file) { return fs.readFileSync(path.join(root, file), 'utf8'); }
function method(text, name, prefix = 'asc_docs_api.prototype.') {
    const start = text.indexOf(prefix + name + ' =');
    if (start < 0) throw new Error('Missing method: ' + name);
    const fn = text.indexOf('function', start), open = text.indexOf('{', fn);
    let depth = 0, quote = '', comment = '';
    for (let i = open; i < text.length; i++) {
        const c = text[i], n = text[i + 1];
        if (comment === '//') { if (c === '\n') comment = ''; continue; }
        if (comment === '/*') { if (c === '*' && n === '/') { comment = ''; i++; } continue; }
        if (quote) { if (c === '\\') i++; else if (c === quote) quote = ''; continue; }
        if (c === '/' && (n === '/' || n === '*')) { comment = c + n; i++; continue; }
        if (c === '"' || c === "'") { quote = c; continue; }
        if (c === '{') depth++;
        if (c === '}' && --depth === 0) return text.slice(fn, i + 1);
    }
    throw new Error('Unclosed method: ' + name);
}
function docInfoClass() {
    const text = source('common/apiCommon.js');
    const body = text.slice(text.indexOf('function asc_CDocInfo()'), text.indexOf('prot.extendWithWopiParams'));
    const tokenEnd = text.indexOf('prot.get_OfflineApp', text.indexOf('prot.get_Token ='));
    return vm.runInNewContext(body + text.slice(text.indexOf('prot.get_Token ='), tokenEnd) + '\nasc_CDocInfo;', {JSON});
}
function connection() {
    let options;
    const common = {
        ConnectionState: {None: 0}, c_oEditorId: {Spreadsheet: 1, Presentation: 2},
        c_oCloseCode: {}, c_oAscServerCommandErrors: {}, c_oAscForceSaveTypes: {},
        getBaseUrlPathname: () => '/web-apps/apps/presentationeditor/main/',
        getSocketIO: () => value => { options = value; return {auth: value.auth, on() {}, io: {on() {}}}; }
    };
    const Asc = {editor: {isPdfEditor: () => false}};
    vm.runInNewContext(source('common/docscoapi.js'), {window: {Asc, AscCommon: common}, Asc, AscCommon: common, performance: {now: () => 0}});
    const wrapper = new common.CDocsCoApi();
    const api = wrapper._CoAuthoringApi;
    function init(token, type, editor = 2) {
        const info = {
            get_Mode: () => 'view', get_Permissions: () => ({edit: false}), get_Lang: () => 'en',
            get_Token: () => token, get_Encrypted: () => false, get_IsWebOpening: () => false,
            get_IsAnonymousUser: () => false, asc_getCoEditingMode: () => undefined
        };
        if (type !== undefined) info.get_TokenType = () => type;
        const user = {asc_getId: () => 'test', asc_getUserName: () => 'Test', asc_getFirstName: () => '', asc_getLastName: () => ''};
        api.init(user, 'test-document', null, 'service-token', editor, 0, info);
        return options;
    }
    return {api, wrapper, init};
}
module.exports = {source, method, docInfoClass, connection};
