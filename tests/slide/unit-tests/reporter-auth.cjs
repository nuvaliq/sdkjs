'use strict';
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {source, method, docInfoClass, connection} = require('./source-harness.cjs');
const vm = require('node:vm');

test('document token type defaults to open and survives cloning', () => {
    const Info = docInfoClass(), info = new Info();
    assert.equal(info.get_TokenType(), 'open');
    info.put_Token('session-value'); info.asc_putTokenType('session');
    const copy = info.clone();
    assert.equal(copy.asc_getTokenType(), 'session'); assert.equal(copy.get_Token(), 'session-value');
    assert.throws(() => info.put_TokenType('invalid'), {name: 'TypeError'});
    const exports = source('common/apiCommon.js');
    for (const name of ['get_TokenType', 'put_TokenType', 'asc_getTokenType', 'asc_putTokenType']) assert.ok(exports.includes('["' + name + '"]'));
});

test('first connection routes each token to its declared slot and clears a previous slot', () => {
    const {api, init} = connection();
    for (const [type, token] of [['session', 'session-value'], ['open', 'open-value'], ['open', null], ['session', 'another-session']]) {
        const auth = init(token, type).auth, command = api.getAuthCommand();
        const outer = type === 'session' ? 'session' : 'token', inner = type === 'session' ? 'jwtSession' : 'jwtOpen';
        assert.equal(auth[outer], token); assert.equal(auth.data[inner], token); assert.equal(command[inner], token);
        assert.equal(auth[type === 'session' ? 'token' : 'session'], undefined);
        assert.equal(auth.data[type === 'session' ? 'jwtOpen' : 'jwtSession'], undefined);
        assert.equal(auth.data.token, 'service-token'); assert.equal(auth.data.mode, 'view');
    }
});

test('invalid types and empty session tokens fail before socket creation', () => {
    const {init} = connection();
    assert.throws(() => init('value', 'other'), {name: 'TypeError'});
    assert.throws(() => init('', 'session'), {name: 'TypeError'});
    assert.throws(() => init(null, 'session'), {name: 'TypeError'});
});

test('legacy document, spreadsheet and presentation connections keep open authentication', () => {
    for (const editor of [0, 1, 2]) {
        const {init} = connection();
        assert.equal(init('open-value', undefined, editor).auth.data.jwtOpen, 'open-value');
    }
});

test('refresh updates both auth representations without rebuilding locks or service data', () => {
    for (const type of ['open', 'session']) {
        const {api, init} = connection();
        const auth = init('initial-value', type).auth, data = auth.data, block = data.block, openCmd = data.openCmd;
        const locks = {one: {state: 2, blockValue: 'locked-block'}};
        api._locks = locks;
        const original = api.getAuthCommand;
        api.getAuthCommand = () => { throw new Error('must not rebuild authentication'); };
        api._onRefreshToken('new-session');
        assert.equal(auth.token, undefined); assert.equal(auth.session, 'new-session');
        assert.equal(data.jwtOpen, undefined); assert.equal(data.jwtSession, 'new-session');
        assert.equal(auth.data, data); assert.equal(data.block, block); assert.equal(data.openCmd, openCmd);
        assert.equal(data.token, 'service-token'); assert.equal(data.mode, 'view'); assert.equal(api._locks, locks);
        api.getAuthCommand = original;
        assert.equal(api.getAuthCommand().jwtSession, 'new-session');
    }
});

test('token synchronization tolerates absent browser objects and keeps native token setters', () => {
    const {api} = connection();
    api._syncAuthTokens(); api.socketio = {}; api._syncAuthTokens(); assert.equal(api.socketio.auth, undefined);
    api.socketio = {auth: {}}; api._syncAuthTokens(); assert.equal(api.socketio.auth.data, undefined);
    const calls = [];
    api.socketio = {io: {setOpenToken: v => calls.push(['open', v]), setSessionToken: v => calls.push(['session', v])}};
    api._onRefreshToken('updated');
    assert.deepEqual(calls, [['open', undefined], ['session', 'updated']]);
});

test('wrapper returns fresh typed tokens and keeps the legacy getter', () => {
    const {api, wrapper, init} = connection();
    assert.equal(wrapper.getAuthToken().token, null);
    init('open-value', 'open'); wrapper._onlineWork = true;
    assert.equal(wrapper.getAuthToken().type, 'open'); assert.equal(wrapper.getAuthToken().token, 'open-value');
    api.jwtSession = 'session-value';
    const selected = wrapper.getAuthToken(); assert.equal(selected.type, 'session'); assert.equal(selected.token, 'session-value');
    selected.token = 'mutated'; assert.equal(wrapper.getAuthToken().token, 'session-value');
    assert.equal(wrapper.get_jwt(), 'session-value');
});

test('presentation exposes typed authentication through its public method', () => {
    const text = source('slide/api.js');
    const get = vm.runInNewContext('(' + method(text, 'asc_getReporterAuth') + ')');
    const {api, wrapper, init} = connection();
    init('opening', 'open'); wrapper._onlineWork = true;
    assert.equal(get.call({CoAuthoringApi: wrapper}).type, 'open');
    api.jwtSession = 'session';
    assert.equal(get.call({CoAuthoringApi: wrapper}).token, 'session');
    assert.ok(text.includes("prototype['asc_getReporterAuth']"));
});
