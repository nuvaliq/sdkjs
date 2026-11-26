'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const vm = require('node:vm');
const {source, method} = require('../../slide/unit-tests/source-harness.cjs');

const normalizedSources = new Map();
function loadMethod(file, prefix, name, globals = {}) {
    let text = normalizedSources.get(file);
    if (!text) {
        text = source(file).replace(/([\w.]+)[ \t]+=/g, '$1 =');
        normalizedSources.set(file, text);
    }
    return vm.runInNewContext('(' + method(text, name, prefix) + ')', {c_oEditorId: AscCommon.c_oEditorId, ...globals});
}

const loadCancellationMethods = [
    '_abortPendingDocumentLoad', 'startFontLoadAction', 'endFontLoadAction',
    'startImageLoadAction', 'endImageLoadAction', 'startApplyChangesAction',
    'endApplyChangesAction', 'endOpenSilentMode', '_openVersionHistoryEndCallback',
    '_closeDocumentLoad', 'endServerIdWaitAction', 'isPluginQueueSuspended',
    '_pausePluginQueueForRefresh', '_finishPluginQueuePause', '_runAbortScope'
];

function attachCancellationMethods(api, globals) {
    for (const name of [...loadCancellationMethods, 'endOpenAction', 'endDisconnectAction', 'clearWaitNotIdle']) {
        api[name] = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    }
}

const noop = () => {};
const AscCommon = {
    c_oEditorId: {Spreadsheet: 'spreadsheet'},
    History: {Clear: noop, private_ClearRecalcData: noop},
    g_oTableId: {Clear: noop},
    g_oIdCounter: {Clear: noop},
    pptx_content_loader: {},
    c_oAscAdvancedOptionsAction: {None: 0},
    CollaborativeEditing: {Clear: noop},
    PasteElementsId: {}
};

test('Visio waits for the reopened document even when the image list is empty', () => {
    const visio = name => loadMethod('visio/api.js', 'VisioEditorApi.prototype.', name, {AscCommon});
    const serverReady = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'asyncServerIdEndLoaded');
    const loadImages = loadMethod('common/GlobalLoaders.js', 'this.', 'LoadImagesWithCallback');
    let readyCount = 0;
    const api = {
        isDocumentLoadComplete: true,
        ServerImagesWaitComplete: true,
        _openDocumentEndCallback: visio('_openDocumentEndCallback'),
        ImageLoader: {},
        WordControl: {
            m_oLogicDocument: {toCShapes: noop, Recalculate: noop, DrawingDocument: {OnEndRecalculate: noop}},
            GoToPage: noop,
            InitControl: noop
        },
        Document: {getCurrentPage: () => 0},
        turnOffSpecialModes: noop,
        setOpenedAt: noop,
        onDocumentContentReady() {
            this.isDocumentLoadComplete = true;
            readyCount++;
        }
    };

    attachCancellationMethods(api, {AscCommon});
    visio('asc_CloseFile').call(api);
    serverReady.call(api);
    assert.equal(readyCount, 0, 'server response cannot make the previous Visio file ready');
    assert.equal(api.ServerImagesWaitComplete, false);

    loadImages.call({Api: api}, [], visio('asyncImagesDocumentEndLoaded'));
    assert.equal(readyCount, 1);
});

test('PDF waits for the renderer to open the new file after reconnecting', () => {
    const events = new Map();
    const Asc = {editor: {asc_registerCallback: noop}};
    AscCommon.CViewer = function() {
        this.registerEvent = (name, callback) => events.set(name, callback);
    };
    const pdf = name => loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', name, {AscCommon, Asc});
    const serverReady = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'asyncServerIdEndLoaded');
    let readyCount = 0;
    const api = {
        isPdfViewer: true,
        isLiveViewer: () => false,
        WordControl: {m_oLogicDocument: {}, m_oDrawingDocument: {}},
        externalChartCollector: {init: noop},
        _openDocumentEndCallback: pdf('_openDocumentEndCallback'),
        onDocumentContentReady() {
            this.isDocumentLoadComplete = true;
            readyCount++;
        },
        sendStandartTextures: noop,
        sendMathToMenu: noop,
        initBroadcastChannelListeners: noop,
        _applyPreOpenLocks: noop,
        turnOffSpecialModes: noop,
        sendEvent: noop,
        setOpenedAt: noop
    };
    pdf('initDocumentRenderer').call(api);
    api.DocumentRenderer.isDocumentContentReady = true;
    api.isDocumentLoadComplete = true;

    attachCancellationMethods(api, {AscCommon});
    pdf('asc_CloseFile').call(api);
    // Fonts and images can finish while a password-protected file still awaits a password.
    api.ServerImagesWaitComplete = true;
    serverReady.call(api);
    assert.equal(readyCount, 0, 'stale renderer state cannot make a password prompt ready');
    assert.equal(api.DocumentRenderer.isDocumentContentReady, false);

    events.get('onFileOpened')();
    assert.equal(readyCount, 1);
});

test('closing the PDF renderer clears content readiness even with an invalid file', () => {
    const close = loadMethod('pdf/src/viewer.js', 'this.', 'close');
    const viewer = {
        openSessionId: 0, getPDFDoc: () => null,
        cancelPendingLoad: loadMethod('pdf/src/viewer.js', 'this.', 'cancelPendingLoad', {clearTimeout}),
        file: {isValid: () => false}, isDocumentContentReady: true, isDocumentReady: true
    };
    close.call(viewer);
    assert.equal(viewer.isDocumentContentReady, false);
    assert.equal(viewer.isDocumentReady, false);
});

function createBaseRefreshHarness() {
    const timers = new Map();
    let nextTimerId = 1;
    let globalLockCount = 0;
    const events = [];
    const c_oAscAsyncActionType = {Empty: 0, BlockInteraction: 1, Information: 2};
    const c_oAscAsyncAction = {Disconnect: 10, RefreshFile: 11, Open: 12, ApplyChanges: 13, LoadDocumentFonts: 14, LoadDocumentImages: 15, LoadFont: 16, LoadImage: 17};
    const c_oAscRestrictionType = {None: 0, View: 1};
    const c_oCloseCode = {quiet: 1, sessionIdle: 2, updateVersion: 3, noCache: 4, restore: 5, drop: 6};
    const Asc = {
        c_oAscAsyncActionType,
        c_oAscAsyncAction,
        c_oAscRestrictionType,
        c_oAscError: {Level: {Critical: -1, NoCritical: 0}, ID: {Unknown: 0, UpdateVersion: 1, Drop: 99}}
    };
    const coApis = [];
    const baseCommon = {
        ...AscCommon,
        c_oCloseCode,
        ConnectionState: {WaitAuth: 1, ClosedCoAuth: 3, ClosedAll: 4},
        getDisconnectErrorCode: (_loaded, code) =>
            code === c_oCloseCode.updateVersion ? Asc.c_oAscError.ID.UpdateVersion : Asc.c_oAscError.ID.Drop,
        getEnableDownloadByCloseCode: () => false,
        CollaborativeEditing: {
            Clear: noop,
            Set_GlobalLock(value) { globalLockCount = value ? globalLockCount + 1 : Math.max(0, globalLockCount - 1); }
        },
        CDocsCoApi: function() {
            this.connectCalls = 0;
            this.disconnectCalls = [];
            this.connect = () => { this.connectCalls++; };
            this.disconnect = code => { this.disconnectCalls.push(code); };
            this.get_state = () => 0;
            coApis.push(this);
        }
    };
    const globals = {
        Asc,
        AscCommon: baseCommon,
        c_oAscAsyncActionType,
        c_oAscError: Asc.c_oAscError,
        c_oCloseCode,
        setTimeout: fn => {
            const id = nextTimerId++;
            timers.set(id, fn);
            return id;
        },
        clearTimeout: id => { timers.delete(id); }
    };
    const loadBase = name => loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    const api = {
        documentIsWopi: false,
        DocInfo: {get_Id: () => 'doc-1'},
        isDocumentLoadComplete: true,
        isOnLoadLicense: true,
        isApplyChangesOnOpenEnabled: true,
        ServerIdWaitComplete: true,
        ServerImagesWaitComplete: false,
        FontLoadWaitComplete: false,
        DocumentRenderer: null,
        ServerIdWaitAction: null,
        IsLongActionCurrent: 0,
        LongActionCallbacks: [],
        LongActionCallbacksParams: [],
        IsActionRestrictionCurrent: 0,
        IsActionRestrictionPrev: null,
        isDisconnectActionActive: false,
        isOpenActionActive: false,
        isRefreshInProgress: false,
        waitNotIdleTimeout: null,
        restrictions: c_oAscRestrictionType.None,
        lastWork: 100,
        getLastWork() { return this.lastWork; },
        getRestrictionSettings: () => null,
        asc_setRestriction(value) { this.restrictions = value; },
        asc_setDocInfo(info) { this.DocInfo = info; },
        asc_CloseFile() { this.isDocumentLoadComplete = false; },
        isDocumentModified: () => false,
        asc_checkNeedCallback: () => true,
        setViewModeDisconnect: noop,
        setOpenedAt: noop,
        _openDocumentEndCallback: noop,
        sendEvent(name, ...args) { events.push([name, ...args]); },
        CoAuthoringApi: null
    };
    for (const name of [
        ...loadCancellationMethods,
        'sync_StartAction', 'sync_EndAction', 'isActionWithRestriction', 'isLongActionBase',
        'incrementCounterLongAction', 'decrementCounterLongAction',
        'incrementCounterActionRestriction', 'decrementCounterActionRestriction',
        'clearWaitNotIdle', 'endDisconnectAction', 'endOpenAction', 'waitNotIdle',
        'canRefreshFile', 'onRefreshFile',
        'asc_refreshFile', 'reopenFileWithReconnection', 'asyncServerIdEndLoaded'
    ]) {
        api[name] = loadBase(name);
    }
    const handlerGlobals = {...globals, t: api};
    const onDisconnect = loadMethod('common/apiBase.js', 'this.CoAuthoringApi.', 'onDisconnect', handlerGlobals);
    const onServerVersion = loadMethod('common/apiBase.js', 'this.CoAuthoringApi.', 'onServerVersion', handlerGlobals);
    const onFirstLoadChangesEnd = loadMethod('common/apiBase.js', 'this.CoAuthoringApi.', 'onFirstLoadChangesEnd', handlerGlobals);
    api._coAuthoringInit = function() {
        this.CoAuthoringApi.onDisconnect = onDisconnect;
        this.CoAuthoringApi.onServerVersion = onServerVersion;
        this.CoAuthoringApi.onFirstLoadChangesEnd = onFirstLoadChangesEnd;
    };
    api.CoAuthoringApi = new baseCommon.CDocsCoApi();
    api._coAuthoringInit();
    return {
        api, events, timers, coApis, Asc, common: baseCommon, ConnectionState: baseCommon.ConnectionState,
        c_oAscAsyncActionType, c_oAscAsyncAction, c_oAscRestrictionType, c_oCloseCode,
        getGlobalLock: () => globalLockCount > 0,
        getGlobalLockCount: () => globalLockCount,
        flushTimers() {
            for (const [id, fn] of [...timers]) {
                timers.delete(id);
                fn();
            }
        }
    };
}

test('refreshFile ends an active Disconnect before starting RefreshFile and restores restrictions on auth', () => {
    const harness = createBaseRefreshHarness();
    const {api, events, c_oAscAsyncActionType, c_oAscAsyncAction, c_oAscRestrictionType} = harness;

    api.CoAuthoringApi.onDisconnect(null, undefined);
    assert.equal(api.isDisconnectActionActive, true);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.restrictions, c_oAscRestrictionType.View);
    assert.equal(harness.getGlobalLock(), true);

    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    assert.equal(api.isDisconnectActionActive, false);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.restrictions, c_oAscRestrictionType.View);
    assert.equal(harness.getGlobalLock(), true);
    assert.deepEqual(events, [
        ['asc_onStartAction', c_oAscAsyncActionType.Information, c_oAscAsyncAction.Disconnect],
        ['asc_onEndAction', c_oAscAsyncActionType.Information, c_oAscAsyncAction.Disconnect],
        ['asc_onStartAction', c_oAscAsyncActionType.Information, c_oAscAsyncAction.RefreshFile]
    ]);

    // onServerVersion without an active Disconnect must not prematurely clear RefreshFile restrictions.
    api.CoAuthoringApi.onServerVersion('9.4.0', 1);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.restrictions, c_oAscRestrictionType.View);
    assert.equal(harness.getGlobalLock(), true);

    api.asyncServerIdEndLoaded();
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, c_oAscRestrictionType.None);
    assert.equal(harness.getGlobalLock(), false);
    assert.deepEqual(events.at(-1), ['asc_onEndAction', c_oAscAsyncActionType.Information, c_oAscAsyncAction.RefreshFile]);
});

test('refreshFile infinity-loop branch clears Disconnect and resets ServerIdWaitAction for subsequent refreshes', () => {
    const harness = createBaseRefreshHarness();
    const {api, events, c_oAscAsyncActionType, c_oAscAsyncAction, c_oAscRestrictionType} = harness;

    api.isDocumentLoadComplete = false;
    api.CoAuthoringApi.onDisconnect(null, undefined);
    api.ServerIdWaitAction = [c_oAscAsyncActionType.Information, c_oAscAsyncAction.RefreshFile, c_oAscRestrictionType.View];
    api.sync_StartAction.apply(api, api.ServerIdWaitAction);
    assert.equal(api.IsActionRestrictionCurrent, 2);

    api.asc_refreshFile({get_Id: () => 'doc-1'});
    assert.equal(api.isDisconnectActionActive, false);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, c_oAscRestrictionType.None);
    assert.equal(events.at(-1)[0], 'asc_onDocumentUpdateVersion');

    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    assert.ok(api.ServerIdWaitAction);
    assert.equal(api.IsActionRestrictionCurrent, 1);
});

test('sessionIdle waitNotIdle reconnects after activity, is cancelled by refresh/reopen, and ignores stale callbacks', () => {
    // 1. Without refreshFile, activity on the second tick reconnects the active CoAuthoringApi.
    const wake = createBaseRefreshHarness();
    const wakeCoApi = wake.api.CoAuthoringApi;
    wake.api.CoAuthoringApi.onDisconnect(null, wake.c_oCloseCode.sessionIdle);
    assert.equal(wake.timers.size, 1);
    wake.flushTimers();
    assert.equal(wake.timers.size, 1, 'first tick records baseline lastWork and reschedules');
    assert.equal(wakeCoApi.connectCalls, 0);
    wake.api.lastWork = 200;
    wake.flushTimers();
    assert.equal(wakeCoApi.connectCalls, 1);

    // 2. asc_refreshFile cancels a rescheduled timer after the first tick.
    const cancel = createBaseRefreshHarness();
    const cancelOldCoApi = cancel.api.CoAuthoringApi;
    cancel.api.CoAuthoringApi.onDisconnect(null, cancel.c_oCloseCode.sessionIdle);
    cancel.flushTimers();
    assert.equal(cancel.timers.size, 1);
    cancel.api.asc_refreshFile({get_Id: () => 'doc-2'});
    assert.equal(cancel.timers.size, 0);
    assert.equal(cancel.api.waitNotIdleTimeout, null);
    assert.notEqual(cancel.api.CoAuthoringApi, cancelOldCoApi);
    assert.equal(cancel.coApis.length, 2);

    // 3. Direct reopenFileWithReconnection ends Disconnect and clears the timer; a stale callback is blocked even if Disconnect restarts on the new connection.
    const reopen = createBaseRefreshHarness();
    const reopenOldCoApi = reopen.api.CoAuthoringApi;
    reopen.api.CoAuthoringApi.onDisconnect(null, reopen.c_oCloseCode.sessionIdle);
    reopen.flushTimers();
    const staleTimerFn = [...reopen.timers.values()][0];
    assert.equal(typeof staleTimerFn, 'function');
    reopen.api.reopenFileWithReconnection({get_Id: () => 'doc-2'});
    assert.equal(reopen.timers.size, 0);
    assert.equal(reopen.api.waitNotIdleTimeout, null);
    assert.equal(reopen.api.isDisconnectActionActive, false);
    assert.equal(reopen.api.IsActionRestrictionCurrent, 0);
    assert.equal(reopen.api.restrictions, reopen.c_oAscRestrictionType.None);
    assert.equal(reopen.getGlobalLock(), false);
    assert.ok(reopen.events.some(e => e[0] === 'asc_onEndAction' && e[2] === reopen.c_oAscAsyncAction.Disconnect));
    assert.notEqual(reopen.api.CoAuthoringApi, reopenOldCoApi);
    // Restart Disconnect on the new CoAuthoringApi so isDisconnectActionActive is true again; coApi mismatch must still block connect.
    reopen.api.CoAuthoringApi.onDisconnect(null, undefined);
    assert.equal(reopen.api.isDisconnectActionActive, true);
    reopen.api.lastWork = 200;
    staleTimerFn();
    assert.equal(reopenOldCoApi.connectCalls, 0);
    assert.equal(reopen.api.CoAuthoringApi.connectCalls, 0);

    // 4. onServerVersion ends active Disconnect, cancels waitNotIdle, and blocks a stale callback on the same CoAuthoringApi.
    const reconnected = createBaseRefreshHarness();
    const sameCoApi = reconnected.api.CoAuthoringApi;
    reconnected.api.CoAuthoringApi.onDisconnect(null, reconnected.c_oCloseCode.sessionIdle);
    reconnected.flushTimers();
    const sameCoApiStaleTimer = [...reconnected.timers.values()][0];
    assert.equal(typeof sameCoApiStaleTimer, 'function');
    reconnected.api.CoAuthoringApi.onServerVersion('9.4.0', 1);
    assert.equal(reconnected.api.isDisconnectActionActive, false);
    assert.equal(reconnected.api.IsActionRestrictionCurrent, 0);
    assert.equal(reconnected.api.restrictions, reconnected.c_oAscRestrictionType.None);
    assert.equal(reconnected.getGlobalLock(), false);
    assert.equal(reconnected.timers.size, 0);
    assert.equal(reconnected.api.waitNotIdleTimeout, null);
    assert.ok(reconnected.events.some(e => e[0] === 'asc_onEndAction' && e[2] === reconnected.c_oAscAsyncAction.Disconnect));
    reconnected.api.lastWork = 200;
    sameCoApiStaleTimer();
    assert.equal(sameCoApi.connectCalls, 0);
});

test('onDisconnect with allowRefresh and repeated onRefreshFile keep restriction counter balanced', () => {
    const harness = createBaseRefreshHarness();
    const {api, events, c_oAscAsyncActionType, c_oAscAsyncAction, c_oAscRestrictionType, c_oCloseCode} = harness;

    api.CoAuthoringApi.onDisconnect(null, undefined);
    assert.equal(api.isDisconnectActionActive, true);
    assert.equal(api.IsActionRestrictionCurrent, 1);

    // Non-idle close code ends Disconnect and starts RefreshFile once.
    api.CoAuthoringApi.onDisconnect(null, c_oCloseCode.updateVersion);
    assert.equal(api.isDisconnectActionActive, false);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.ServerIdWaitAction[1], c_oAscAsyncAction.RefreshFile);

    // Repeated allowRefresh or documentOpen("updateversion") during an in-flight RefreshFile must not double-lock.
    api.CoAuthoringApi.onDisconnect(null, c_oCloseCode.updateVersion);
    api.onRefreshFile();
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(
        events.filter(e => e[0] === 'asc_onStartAction' && e[2] === c_oAscAsyncAction.RefreshFile).length,
        1
    );

    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    api.CoAuthoringApi.onServerVersion('9.4.0', 1);
    api.asyncServerIdEndLoaded();
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, c_oAscRestrictionType.None);
    assert.equal(harness.getGlobalLock(), false);
    assert.equal(api.ServerIdWaitAction, null);

    // Switching from another 2-element ServerIdWaitAction (e.g. Open from asc_showRevision) to RefreshFile ends the previous action first.
    events.length = 0;
    api.ServerIdWaitAction = [c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open];
    api.sync_StartAction.apply(api, api.ServerIdWaitAction);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(api.IsActionRestrictionCurrent, 0);
    api.onRefreshFile();
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.ServerIdWaitAction[1], c_oAscAsyncAction.RefreshFile);
    const openEndIdx = events.findIndex(e => e[0] === 'asc_onEndAction' && e[1] === c_oAscAsyncActionType.BlockInteraction && e[2] === c_oAscAsyncAction.Open);
    const refreshStartIdx = events.findIndex(e => e[0] === 'asc_onStartAction' && e[1] === c_oAscAsyncActionType.Information && e[2] === c_oAscAsyncAction.RefreshFile);
    assert.ok(openEndIdx >= 0 && refreshStartIdx > openEndIdx, 'Open ends before RefreshFile starts');
    api.asyncServerIdEndLoaded();
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, c_oAscRestrictionType.None);

    // asc_refreshFile also replaces a pending 2-element Open action with RefreshFile and protects against ClosedAll readiness.
    events.length = 0;
    let openEndCalls = 0;
    api._openDocumentEndCallback = () => { openEndCalls++; };
    api.isDocumentLoadComplete = false;
    api.ServerIdWaitComplete = false;
    api.ServerIdWaitAction = [c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open];
    api.sync_StartAction.apply(api, api.ServerIdWaitAction);
    assert.equal(api.IsLongActionCurrent, 1);
    api.asc_refreshFile({get_Id: () => 'doc-3', asc_getOptions: () => ({})});
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.ServerIdWaitAction[1], c_oAscAsyncAction.RefreshFile);
    const ascOpenEndIdx = events.findIndex(e => e[0] === 'asc_onEndAction' && e[1] === c_oAscAsyncActionType.BlockInteraction && e[2] === c_oAscAsyncAction.Open);
    const ascRefreshStartIdx = events.findIndex(e => e[0] === 'asc_onStartAction' && e[1] === c_oAscAsyncActionType.Information && e[2] === c_oAscAsyncAction.RefreshFile);
    assert.ok(ascOpenEndIdx >= 0 && ascRefreshStartIdx > ascOpenEndIdx, 'asc_refreshFile ends Open before starting RefreshFile');
    api.CoAuthoringApi.get_state = () => harness.ConnectionState.ClosedAll;
    api.CoAuthoringApi.onDisconnect('drop', c_oCloseCode.drop);
    assert.equal(openEndCalls, 0);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsActionRestrictionCurrent, 0);
});

test('non-refresh onDisconnect ends Disconnect and waitNotIdle, repeat waitNotIdle replaces timer, and sessionIdle respects restrictions', () => {
    // 1. Terminal close code (drop) ends Disconnect and cancels waitNotIdle without calling onRefreshFile.
    const dropped = createBaseRefreshHarness();
    dropped.api.CoAuthoringApi.onDisconnect(null, dropped.c_oCloseCode.sessionIdle);
    assert.equal(dropped.api.isDisconnectActionActive, true);
    assert.equal(dropped.api.IsActionRestrictionCurrent, 1);
    assert.equal(dropped.timers.size, 1);

    dropped.api.CoAuthoringApi.onDisconnect(null, dropped.c_oCloseCode.drop);
    assert.equal(dropped.api.isDisconnectActionActive, false);
    assert.equal(dropped.api.IsActionRestrictionCurrent, 0);
    assert.equal(dropped.api.waitNotIdleTimeout, null);
    assert.equal(dropped.timers.size, 0);
    assert.equal(dropped.api.ServerIdWaitAction, null);
    assert.ok(dropped.events.some(e => e[0] === 'asc_onEndAction' && e[2] === dropped.c_oAscAsyncAction.Disconnect));
    assert.equal(dropped.events.at(-1)[0], 'asc_onError');

    // 2. allowRefresh code on a modified document also takes the non-refresh branch and still ends Disconnect.
    const modified = createBaseRefreshHarness();
    modified.api.isDocumentModified = () => true;
    modified.api.CoAuthoringApi.onDisconnect(null, undefined);
    assert.equal(modified.api.isDisconnectActionActive, true);
    assert.equal(modified.api.IsActionRestrictionCurrent, 1);

    modified.api.CoAuthoringApi.onDisconnect(null, modified.c_oCloseCode.updateVersion);
    assert.equal(modified.api.isDisconnectActionActive, false);
    assert.equal(modified.api.IsActionRestrictionCurrent, 0);
    assert.equal(modified.api.ServerIdWaitAction, null);
    assert.equal(modified.events.at(-1)[0], 'asc_onError');

    // 3. Repeating waitNotIdle while a timer is active cancels the previous timer instead of leaking it.
    const repeat = createBaseRefreshHarness();
    let firstFired = 0;
    let secondFired = 0;
    repeat.api.waitNotIdle(undefined, () => { firstFired++; });
    assert.equal(repeat.timers.size, 1);
    repeat.api.waitNotIdle(undefined, () => { secondFired++; });
    assert.equal(repeat.timers.size, 1);
    repeat.flushTimers();
    repeat.api.lastWork = 200;
    repeat.flushTimers();
    assert.equal(firstFired, 0);
    assert.equal(secondFired, 1);

    // 4. sessionIdle during an active RefreshFile restriction must not start Disconnect or schedule waitNotIdle.
    const restricted = createBaseRefreshHarness();
    restricted.api.onRefreshFile();
    assert.equal(restricted.api.IsActionRestrictionCurrent, 1);
    restricted.api.CoAuthoringApi.onDisconnect(null, restricted.c_oCloseCode.sessionIdle);
    assert.equal(restricted.api.isDisconnectActionActive, false);
    assert.equal(restricted.api.IsActionRestrictionCurrent, 1);
    assert.equal(restricted.timers.size, 0);
});

test('ClosedAll disconnect during refreshFile does not emit readiness before critical error and keeps restrictions balanced', () => {
    // 1. Terminal ClosedAll during refreshFile must end RefreshFile before setViewModeDisconnect and Critical error, without firing readiness.
    const failed = createBaseRefreshHarness();
    let openEndCalls = 0;
    failed.api.isViewMode = false;
    failed.api.asc_setRestriction = function(value) {
        this.restrictions = value;
        this.sendEvent('asc_onChangeRestrictions', value);
    };
    failed.api.asc_setViewMode = function(isViewMode) {
        this.isViewMode = Boolean(isViewMode);
        if (isViewMode) {
            failed.common.CollaborativeEditing.Set_GlobalLock(true);
        }
    };
    failed.api.setViewModeDisconnect = function(enableDownload) {
        this.sendEvent('asc_onCoAuthoringDisconnect', enableDownload);
        this.asc_setViewMode(true);
    };
    failed.api._openDocumentEndCallback = function() {
        openEndCalls++;
        if (this.ServerIdWaitComplete) {
            this.isDocumentLoadComplete = true;
            this.sendEvent('asc_onDocumentContentReady');
        }
    };
    failed.api.asc_refreshFile({get_Id: () => 'doc-2'});
    assert.equal(failed.api.ServerIdWaitComplete, false);
    assert.equal(failed.api.isDocumentLoadComplete, false);
    assert.equal(failed.api.ServerIdWaitAction[1], failed.c_oAscAsyncAction.RefreshFile);

    failed.api.CoAuthoringApi.get_state = () => failed.ConnectionState.ClosedAll;
    failed.events.length = 0;
    failed.api.CoAuthoringApi.onDisconnect('reconnect_failed', failed.c_oCloseCode.drop);

    assert.equal(openEndCalls, 0);
    assert.equal(failed.api.ServerIdWaitComplete, false);
    assert.equal(failed.api.isDocumentLoadComplete, false);
    assert.equal(failed.api.ServerIdWaitAction, null);
    assert.equal(failed.api.IsActionRestrictionCurrent, 0);
    assert.equal(failed.api.restrictions, failed.c_oAscRestrictionType.None);
    assert.equal(failed.api.isViewMode, true);
    assert.equal(failed.getGlobalLock(), true);
    const endIndex = failed.events.findIndex(e => e[0] === 'asc_onEndAction' && e[2] === failed.c_oAscAsyncAction.RefreshFile);
    const restrictionIndex = failed.events.findIndex(e => e[0] === 'asc_onChangeRestrictions' && e[1] === failed.c_oAscRestrictionType.None);
    const disconnectIndex = failed.events.findIndex(e => e[0] === 'asc_onCoAuthoringDisconnect');
    const errIndex = failed.events.findIndex(e => e[0] === 'asc_onError' && e[2] === failed.Asc.c_oAscError.Level.Critical);
    assert.ok(endIndex >= 0, 'RefreshFile action ends');
    assert.ok(restrictionIndex >= 0 && restrictionIndex < disconnectIndex, 'restrictions restore before disconnect view lock');
    assert.ok(disconnectIndex > endIndex, 'RefreshFile action ends before setViewModeDisconnect');
    assert.ok(errIndex > disconnectIndex, 'critical error is emitted after disconnect');
    assert.equal(failed.events.some(e => e[0] === 'asc_onDocumentContentReady'), false);

    // 2. allowRefresh ClosedAll during refreshFile keeps RefreshFile active and balanced without emitting readiness.
    const retry = createBaseRefreshHarness();
    let retryOpenEndCalls = 0;
    retry.api._openDocumentEndCallback = () => { retryOpenEndCalls++; };
    retry.api.asc_refreshFile({get_Id: () => 'doc-2'});
    retry.api.CoAuthoringApi.get_state = () => retry.ConnectionState.ClosedAll;
    retry.api.CoAuthoringApi.onDisconnect(null, retry.c_oCloseCode.updateVersion);
    assert.equal(retryOpenEndCalls, 0);
    assert.equal(retry.api.ServerIdWaitComplete, false);
    assert.equal(retry.api.isDocumentLoadComplete, false);
    assert.equal(retry.api.ServerIdWaitAction[1], retry.c_oAscAsyncAction.RefreshFile);
    assert.equal(retry.api.IsActionRestrictionCurrent, 1);

    // 3. sessionIdle ClosedAll during refreshFile ends RefreshFile without readiness, starts Disconnect with waitNotIdle, and clears Disconnect on first auth.
    const idle = createBaseRefreshHarness();
    let idleOpenEndCalls = 0;
    let disconnectEndedBeforeReady = false;
    idle.api._openDocumentEndCallback = function() {
        idleOpenEndCalls++;
        disconnectEndedBeforeReady = !this.isDisconnectActionActive &&
            this.IsActionRestrictionCurrent === 0 &&
            idle.events.some(e => e[0] === 'asc_onEndAction' && e[2] === idle.c_oAscAsyncAction.Disconnect);
    };
    idle.api.asc_refreshFile({get_Id: () => 'doc-2'});
    const idleCoApi = idle.api.CoAuthoringApi;
    idle.api.CoAuthoringApi.get_state = () => idle.ConnectionState.ClosedAll;
    idle.api.CoAuthoringApi.onDisconnect(null, idle.c_oCloseCode.sessionIdle);
    assert.equal(idleOpenEndCalls, 0);
    assert.equal(idle.api.ServerIdWaitAction, null);
    assert.equal(idle.api.isDisconnectActionActive, true);
    assert.equal(idle.api.IsActionRestrictionCurrent, 1);
    assert.equal(idle.timers.size, 1);
    idle.flushTimers();
    idle.api.lastWork = 200;
    idle.flushTimers();
    assert.equal(idleCoApi.connectCalls, 1);
    idle.api.CoAuthoringApi.onFirstLoadChangesEnd(123);
    assert.equal(idleOpenEndCalls, 1);
    assert.equal(disconnectEndedBeforeReady, true, 'Disconnect ends before readiness callback runs');
    assert.equal(idle.api.ServerIdWaitComplete, true);
    assert.equal(idle.api.isDisconnectActionActive, false);
    assert.equal(idle.api.IsActionRestrictionCurrent, 0);
    assert.equal(idle.api.restrictions, idle.c_oAscRestrictionType.None);
    assert.equal(idle.getGlobalLock(), false);

    // 4. Non-refresh initial load ClosedAll preserves upstream asyncServerIdEndLoaded behavior.
    const initial = createBaseRefreshHarness();
    let initialOpenEndCalls = 0;
    initial.api.isDocumentLoadComplete = false;
    initial.api.ServerIdWaitComplete = false;
    initial.api.ServerIdWaitAction = null;
    initial.api._openDocumentEndCallback = () => { initialOpenEndCalls++; };
    initial.api.CoAuthoringApi.get_state = () => initial.ConnectionState.ClosedAll;
    initial.api.CoAuthoringApi.onDisconnect('drop', initial.c_oCloseCode.drop);
    assert.equal(initialOpenEndCalls, 1);
    assert.equal(initial.api.ServerIdWaitComplete, true);

    // 5. Terminal drop in ClosedCoAuth (non-ClosedAll) during refreshFile ends RefreshFile before disconnect and Critical error.
    const dropped = createBaseRefreshHarness();
    let droppedOpenEndCalls = 0;
    dropped.api.isViewMode = false;
    dropped.api.asc_setViewMode = function(isViewMode) {
        this.isViewMode = Boolean(isViewMode);
        if (isViewMode) {
            dropped.common.CollaborativeEditing.Set_GlobalLock(true);
        }
    };
    dropped.api.setViewModeDisconnect = function(enableDownload) {
        this.sendEvent('asc_onCoAuthoringDisconnect', enableDownload);
        this.asc_setViewMode(true);
    };
    dropped.api._openDocumentEndCallback = () => { droppedOpenEndCalls++; };
    dropped.api.asc_refreshFile({get_Id: () => 'doc-2'});
    dropped.api.CoAuthoringApi.get_state = () => dropped.ConnectionState.ClosedCoAuth;
    dropped.events.length = 0;
    dropped.api.CoAuthoringApi.onDisconnect('drop', dropped.c_oCloseCode.drop);
    assert.equal(droppedOpenEndCalls, 0);
    assert.equal(dropped.api.ServerIdWaitAction, null);
    assert.equal(dropped.api.IsActionRestrictionCurrent, 0);
    assert.equal(dropped.api.restrictions, dropped.c_oAscRestrictionType.None);
    assert.equal(dropped.api.isViewMode, true);
    assert.equal(dropped.getGlobalLock(), true);
    const dropEndIdx = dropped.events.findIndex(e => e[0] === 'asc_onEndAction' && e[2] === dropped.c_oAscAsyncAction.RefreshFile);
    const dropDiscIdx = dropped.events.findIndex(e => e[0] === 'asc_onCoAuthoringDisconnect');
    const dropErrIdx = dropped.events.findIndex(e => e[0] === 'asc_onError' && e[2] === dropped.Asc.c_oAscError.Level.Critical);
    assert.ok(dropEndIdx >= 0 && dropDiscIdx > dropEndIdx && dropErrIdx > dropDiscIdx, 'RefreshFile ends before disconnect and critical error in ClosedCoAuth');

    // 6. sessionIdle in WaitAuth (non-ClosedAll) during asc_refreshFile hands off from RefreshFile to Disconnect + waitNotIdle and clears on first auth.
    const waitAuthIdle = createBaseRefreshHarness();
    let waitAuthOpenEndCalls = 0;
    waitAuthIdle.api._openDocumentEndCallback = () => { waitAuthOpenEndCalls++; };
    waitAuthIdle.api.asc_refreshFile({get_Id: () => 'doc-2'});
    const waitAuthCoApi = waitAuthIdle.api.CoAuthoringApi;
    waitAuthIdle.api.CoAuthoringApi.get_state = () => waitAuthIdle.ConnectionState.WaitAuth;
    waitAuthIdle.api.CoAuthoringApi.onDisconnect(null, waitAuthIdle.c_oCloseCode.sessionIdle);
    assert.equal(waitAuthOpenEndCalls, 0);
    assert.equal(waitAuthIdle.api.ServerIdWaitAction, null);
    assert.equal(waitAuthIdle.api.isDisconnectActionActive, true);
    assert.equal(waitAuthIdle.api.IsActionRestrictionCurrent, 1);
    assert.equal(waitAuthIdle.timers.size, 1);
    waitAuthIdle.flushTimers();
    waitAuthIdle.api.lastWork = 200;
    waitAuthIdle.flushTimers();
    assert.equal(waitAuthCoApi.connectCalls, 1);
    waitAuthIdle.api.CoAuthoringApi.onFirstLoadChangesEnd(456);
    assert.equal(waitAuthOpenEndCalls, 1);
    assert.equal(waitAuthIdle.api.isDisconnectActionActive, false);
    assert.equal(waitAuthIdle.api.IsActionRestrictionCurrent, 0);
    assert.equal(waitAuthIdle.api.restrictions, waitAuthIdle.c_oAscRestrictionType.None);
    assert.equal(waitAuthIdle.getGlobalLock(), false);
});

test('repeat refreshFile before document load completes resets stale asset readiness flags across editors', () => {
    // 1. PDF: repeat refreshFile on a new key while isDocumentLoadComplete === false skips asc_CloseFile but still resets readiness flags.
    const pdfHarness = createBaseRefreshHarness();
    let pdfCloseCalls = 0;
    let pdfReadyCalls = 0;
    const pdfRenderer = {isDocumentContentReady: true};
    pdfHarness.api.isDocumentLoadComplete = false;
    pdfHarness.api.isApplyChangesOnOpenEnabled = false;
    pdfHarness.api.ServerIdWaitComplete = false;
    pdfHarness.api.ServerImagesWaitComplete = true;
    pdfHarness.api.FontLoadWaitComplete = true;
    pdfHarness.api.DocumentRenderer = pdfRenderer;
    pdfHarness.api.WordControl = {
        m_oLogicDocument: {},
        m_oDrawingDocument: {m_oDocumentRenderer: pdfRenderer}
    };
    pdfHarness.api.isPdfViewer = true;
    pdfHarness.api.isLiveViewer = () => false;
    pdfHarness.api.externalChartCollector = {init: noop};
    pdfHarness.api.initBroadcastChannelListeners = noop;
    pdfHarness.api.asc_CloseFile = () => { pdfCloseCalls++; };
    pdfHarness.api.sendMathToMenu = noop;
    pdfHarness.api.sendStandartTextures = noop;
    pdfHarness.api._applyPreOpenLocks = noop;
    pdfHarness.api.onDocumentContentReady = () => {
        pdfHarness.api.isDocumentLoadComplete = true;
        pdfReadyCalls++;
    };
    pdfHarness.api._openDocumentEndCallback = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', '_openDocumentEndCallback', {
        Asc: {editor: {asc_registerCallback: noop}},
        AscCommon: {
            PasteElementsId: {},
            CollaborativeEditing: {
                m_aChanges: [],
                CheckLoadChanges: (_api, done) => done()
            }
        }
    });

    pdfHarness.api.asc_refreshFile({get_Id: () => 'doc-2'});
    assert.equal(pdfCloseCalls, 0);
    assert.equal(pdfHarness.api.isApplyChangesOnOpenEnabled, true);
    assert.equal(pdfHarness.api.ServerImagesWaitComplete, false);
    assert.equal(pdfHarness.api.FontLoadWaitComplete, false);
    assert.equal(pdfRenderer.isDocumentContentReady, false);

    // Auth for doc-2 must not fire readiness before doc-2 fonts and renderer finish loading.
    pdfHarness.api.CoAuthoringApi.onFirstLoadChangesEnd(100);
    assert.equal(pdfReadyCalls, 0);
    pdfHarness.api.ServerImagesWaitComplete = true;
    pdfHarness.api._openDocumentEndCallback();
    assert.equal(pdfReadyCalls, 0);
    pdfRenderer.isDocumentContentReady = true;
    pdfHarness.api._openDocumentEndCallback();
    assert.equal(pdfReadyCalls, 1);

    // 2. Visio: stale ServerImagesWaitComplete is cleared on repeat refreshFile before doc-1 completed.
    const visioHarness = createBaseRefreshHarness();
    let visioReadyCalls = 0;
    visioHarness.api.isDocumentLoadComplete = false;
    visioHarness.api.ServerIdWaitComplete = false;
    visioHarness.api.ServerImagesWaitComplete = true;
    visioHarness.api.WordControl = {
        m_oLogicDocument: {toCShapes: noop, Recalculate: noop, DrawingDocument: {OnEndRecalculate: noop}},
        GoToPage: noop,
        InitControl: noop
    };
    visioHarness.api.Document = {getCurrentPage: () => 0};
    visioHarness.api.sendMathToMenu = noop;
    visioHarness.api.sendStandartTextures = noop;
    visioHarness.api._applyPreOpenLocks = noop;
    visioHarness.api.onDocumentContentReady = () => {
        visioHarness.api.isDocumentLoadComplete = true;
        visioReadyCalls++;
    };
    visioHarness.api._openDocumentEndCallback = loadMethod('visio/api.js', 'VisioEditorApi.prototype.', '_openDocumentEndCallback', {
        AscCommon
    });
    visioHarness.api.asc_refreshFile({get_Id: () => 'doc-2'});
    assert.equal(visioHarness.api.ServerImagesWaitComplete, false);
    visioHarness.api.CoAuthoringApi.onFirstLoadChangesEnd(100);
    assert.equal(visioReadyCalls, 0);
    visioHarness.api.ServerImagesWaitComplete = true;
    visioHarness.api._openDocumentEndCallback();
    assert.equal(visioReadyCalls, 1);

    // 3. Cell: stale FontLoadWaitComplete is cleared on repeat refreshFile before doc-1 completed.
    const cellHarness = createBaseRefreshHarness();
    let cellProceeded = 0;
    const cellSentinel = Symbol('cellPastReadinessGuard');
    cellHarness.api.isDocumentLoadComplete = false;
    cellHarness.api.ServerIdWaitComplete = false;
    cellHarness.api.FontLoadWaitComplete = true;
    cellHarness.api.wbModel = {
        getWorksheetCount() {
            cellProceeded++;
            throw cellSentinel;
        }
    };
    cellHarness.api._openDocumentEndCallback = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_openDocumentEndCallback', {
        AscCommon: {}
    });
    cellHarness.api.asc_refreshFile({get_Id: () => 'doc-2'});
    assert.equal(cellHarness.api.FontLoadWaitComplete, false);
    cellHarness.api.CoAuthoringApi.onFirstLoadChangesEnd(100);
    assert.equal(cellProceeded, 0);
    cellHarness.api.FontLoadWaitComplete = true;
    assert.throws(() => cellHarness.api._openDocumentEndCallback(), err => err === cellSentinel);
    assert.equal(cellProceeded, 1);
});

function createOpenCommandHarness(extraGlobals = {}) {
    const openRequests = [];
    const loadedResults = [];
    const events = [];
    const coApis = [];
    const c_oAscAsyncActionType = {Empty: 0, BlockInteraction: 1, Information: 2};
    const c_oAscAsyncAction = {Disconnect: 10, RefreshFile: 11, Open: 12};
    const c_oAscRestrictionType = {None: 0, View: 1};
    const c_oCloseCode = {quiet: 1, sessionIdle: 2, updateVersion: 3, noCache: 4, restore: 5, drop: 6};
    const c_oAscError = {
        Level: {Critical: -1, NoCritical: 0},
        ID: {No: 0, Unknown: 1, UpdateVersion: 2, DownloadError: 3, ConvertationOpenError: 4, Drop: 99}
    };
    const Asc = {
        c_oAscAsyncActionType,
        c_oAscAsyncAction,
        c_oAscRestrictionType,
        c_oAscError,
        c_nCommonRequestTime: 1000
    };
    const baseCommon = {
        ...AscCommon,
        c_oCloseCode,
        c_oSerFormat: {Signature: 'DOCY'},
        ConnectionState: {WaitAuth: 1, ClosedCoAuth: 3, ClosedAll: 4},
        EncryptionWorker: {isNeedCrypt: () => false},
        g_oDocumentUrls: {init: noop},
        asc_CAscEditorPermissions: function() { this.setLicenseType = noop; },
        sendClientLog: noop,
        getClientInfoString: () => 'info',
        getDisconnectErrorCode: () => c_oAscError.ID.Drop,
        getEnableDownloadByCloseCode: () => false,
        mapAscServerErrorToAscError: code => code,
        openFileCommand(docId, url, _changesUrl, _changesToken, _sig, callback) {
            openRequests.push({docId, url, callback});
        },
        CollaborativeEditing: {
            Clear: noop,
            Set_GlobalLock: noop
        },
        CDocsCoApi: function() {
            this.disconnectCalls = [];
            this.rpcCallbacks = [];
            this.disconnect = code => { this.disconnectCalls.push(code); };
            this.get_state = () => 0;
            this.get_isAuth = () => true;
            this.auth = noop;
            this.callPRC = (_req, _timeout, cb) => {
                this.rpcCallbacks.push(cb);
                return true;
            };
            coApis.push(this);
        }
    };
    const globals = {
        Asc,
        AscCommon: baseCommon,
        c_oAscAsyncActionType,
        c_oAscAsyncAction,
        c_oAscRestrictionType,
        c_oAscError,
        c_oCloseCode,
        performance: {now: () => 0},
        setTimeout: fn => { fn(); return 1; },
        clearTimeout: noop,
        ...extraGlobals
    };
    const loadBase = name => loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    const api = {
        documentIsWopi: false,
        documentId: 'doc-1',
        documentFormat: 'docx',
        documentUrlChanges: null,
        documentTokenChanges: null,
        DocInfo: {get_Id: () => 'doc-1', extendWithWopiParams: res => ({get_Id: () => res.id, get_Title: () => res.title})},
        isDocumentLoadComplete: false,
        documentLoadGeneration: 0,
        isOnLoadLicense: false,
        isApplyChangesOnOpenEnabled: true,
        isOpenedFrameEditor: false,
        isUseNativeViewer: false,
        isCoAuthoringEnable: false,
        ServerIdWaitComplete: false,
        ServerImagesWaitComplete: false,
        FontLoadWaitComplete: false,
        DocumentRenderer: null,
        ServerIdWaitAction: null,
        IsLongActionCurrent: 0,
        LongActionCallbacks: [],
        LongActionCallbacksParams: [],
        IsActionRestrictionCurrent: 0,
        IsActionRestrictionPrev: null,
        isDisconnectActionActive: false,
        isOpenActionActive: false,
        isRefreshInProgress: false,
        waitNotIdleTimeout: null,
        restrictions: c_oAscRestrictionType.None,
        asc_isSupportFeature: () => false,
        getViewMode: () => false,
        getRestrictionSettings: () => null,
        asc_setRestriction(value) { this.restrictions = value; },
        asc_setDocInfo(info) {
            this.DocInfo = info;
            this.documentId = info.get_Id();
        },
        asc_CloseFile() { this.isDocumentLoadComplete = false; },
        asc_LoadDocument() { events.push(['asc_LoadDocument']); },
        asc_coAuthoringDisconnect() { events.push(['asc_coAuthoringDisconnect']); },
        isDocumentModified: () => false,
        asc_checkNeedCallback: () => true,
        setViewModeDisconnect: noop,
        setOpenedAt: noop,
        _onEndPermissions() { events.push(['_onEndPermissions']); },
        _openDocumentEndCallback() { events.push(['_openDocumentEndCallback']); },
        onEndLoadFile(result) { loadedResults.push(result); },
        sendEvent(name, ...args) { events.push([name, ...args]); },
        CoAuthoringApi: null
    };
    for (const name of [
        ...loadCancellationMethods,
        'sync_StartAction', 'sync_EndAction', 'isActionWithRestriction', 'isLongActionBase',
        'incrementCounterLongAction', 'decrementCounterLongAction',
        'incrementCounterActionRestriction', 'decrementCounterActionRestriction',
        'clearWaitNotIdle', 'endDisconnectAction', 'endOpenAction', '_onOpenCommand',
        'canRefreshFile', 'onRefreshFile', 'asc_refreshFile',
        'reopenFileWithReconnection', 'asyncServerIdEndLoaded'
    ]) {
        api[name] = loadBase(name);
    }
    const handlerGlobals = {...globals, t: api};
    const handlers = {};
    for (const name of [
        'onDisconnect', 'onDocumentOpen', 'onFirstLoadChangesEnd', 'onServerVersion',
        'onFirstConnect', 'onLicense', 'onLicenseChanged', 'onExpiredToken'
    ]) {
        handlers[name] = loadMethod('common/apiBase.js', 'this.CoAuthoringApi.', name, handlerGlobals);
    }
    api._coAuthoringInit = function() {
        for (const [name, fn] of Object.entries(handlers)) {
            this.CoAuthoringApi[name] = fn;
        }
        this.CoAuthoringApi.onMeta = noop;
    };
    api.CoAuthoringApi = new baseCommon.CDocsCoApi();
    api._coAuthoringInit();
    return {api, openRequests, loadedResults, events, coApis, c_oCloseCode, c_oAscError};
}

test('stale openFileCommand response from an aborted refreshFile cannot load the old file or fail the new one', () => {
    const {api, openRequests, loadedResults, events, c_oCloseCode, c_oAscError} = createOpenCommandHarness();
    api.isDocumentLoadComplete = true;

    // First refresh starts downloading doc-1.
    api.asc_refreshFile({get_Id: () => 'doc-1'});
    api.CoAuthoringApi.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 1, data: {'Editor.bin': 'https://files/doc-1.bin'}}});
    assert.equal(openRequests.length, 1);
    assert.equal(openRequests[0].docId, 'doc-1');

    // Engine requests a second refresh while doc-1 HTTP request is still in flight.
    api.CoAuthoringApi.onDisconnect(null, c_oCloseCode.updateVersion);
    assert.ok(events.some(([name]) => name === 'asc_onRequestRefreshFile'));

    // Even before the host responds with doc-2, a late doc-1 HTTP response must be ignored.
    const doc1Request = openRequests[0];
    doc1Request.callback(c_oAscError.ID.No, {version: 'doc-1-early'});
    assert.deepEqual(loadedResults, []);

    // Host sends doc-2 and doc-2 starts downloading.
    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    api.CoAuthoringApi.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 2, data: {'Editor.bin': 'https://files/doc-2.bin'}}});
    assert.equal(openRequests.length, 2);
    assert.equal(openRequests[1].docId, 'doc-2');

    // Late doc-1 success or download error must neither open doc-1 nor emit critical asc_onError.
    events.length = 0;
    doc1Request.callback(c_oAscError.ID.No, {version: 'doc-1-late'});
    doc1Request.callback(c_oAscError.ID.DownloadError, null);
    assert.deepEqual(loadedResults, []);
    assert.deepEqual(events, []);

    // Completing doc-2 loads doc-2 cleanly.
    openRequests[1].callback(c_oAscError.ID.No, {version: 'doc-2'});
    assert.deepEqual(loadedResults, [{version: 'doc-2'}]);
});

test('stale CoAuthoringApi events and deferred update/WOPI callbacks are ignored after reopen', () => {
    const {api, openRequests, events, coApis, c_oCloseCode} = createOpenCommandHarness();
    const oldCoApi = api.CoAuthoringApi;

    // Trigger asc_onDocumentUpdateVersion callback path (canRefreshFile = false) before reopening.
    api.asc_checkNeedCallback = () => false;
    oldCoApi.onDocumentOpen({data: {type: 'open', status: 'updateversion', openedAt: 1, data: {'Editor.bin': 'https://files/doc-1.bin'}}});
    const updateVersionEvent = events.find(([name]) => name === 'asc_onDocumentUpdateVersion');
    assert.ok(updateVersionEvent);
    const staleContinueOpen = updateVersionEvent[1];

    // Reopen with doc-2 replaces CoAuthoringApi and increments documentLoadGeneration.
    api.asc_checkNeedCallback = () => true;
    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    assert.notEqual(api.CoAuthoringApi, oldCoApi);
    assert.equal(coApis.length, 2);
    events.length = 0;

    // Stale callback from old updateversion dialog must not trigger openFileCommand for doc-1.
    staleContinueOpen();
    assert.equal(openRequests.length, 0);

    // Delayed events from oldCoApi must all be ignored.
    oldCoApi.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 1, data: {'Editor.bin': 'https://files/doc-1.bin'}}});
    oldCoApi.onFirstLoadChangesEnd(123);
    oldCoApi.onDisconnect(null, c_oCloseCode.updateVersion);
    oldCoApi.onServerVersion('9.4.0', 1);
    oldCoApi.onFirstConnect();
    oldCoApi.onLicense({type: 1});
    oldCoApi.onLicenseChanged({licenseType: 1});
    oldCoApi.onExpiredToken({code: 1});
    assert.equal(openRequests.length, 0);
    assert.equal(api.ServerIdWaitComplete, false);
    assert.deepEqual(events, []);

    // Stale WOPI refresh callback must be ignored after a newer reopen.
    api.documentIsWopi = true;
    api.onRefreshFile();
    const staleWopiCallback = api.CoAuthoringApi.rpcCallbacks[0];
    api.documentIsWopi = false;
    api.asc_refreshFile({get_Id: () => 'doc-3', asc_getOptions: () => ({})});
    events.length = 0;
    staleWopiCallback(false, {id: 'doc-wopi-stale', title: 'Stale'});
    assert.equal(api.documentId, 'doc-3');
    assert.deepEqual(events, []);
});

test('stale Apply_Changes callbacks in Word and PDF do not fire onDocumentContentReady after reopen', () => {
    // 1. Word
    const wordApplyCallbacks = [];
    let wordReadyCalls = 0;
    let wordSilentStarts = 0;
    let wordSilentEnds = 0;
    const wordEnd = loadMethod('word/api.js', 'asc_docs_api.prototype.', '_openDocumentEndCallback', {
        Asc: {editor: {asc_registerCallback: noop}},
        c_oAscAdvancedOptionsAction: {None: 0},
        AscCommon: {
            sendClientLog: noop,
            getClientInfoString: () => 'info',
            CollaborativeEditing: {
                m_aChanges: [],
                Have_OtherChanges: () => false,
                Apply_Changes: cb => { wordApplyCallbacks.push(cb); },
                Release_Locks: noop
            }
        },
        performance: {now: () => 0}
    });
    const wordApi = {
        documentLoadGeneration: 1,
        isDocumentLoadComplete: false,
        ServerImagesWaitComplete: true,
        ServerIdWaitComplete: true,
        DocumentType: 2,
        LoadedObject: 1,
        isApplyChangesOnOpenEnabled: true,
        isSaveFonts_Images: false,
        isOnlyReaderMode: false,
        isViewMode: false,
        WordControl: {
            m_oLogicDocument: {
                Start_SilentMode: () => { wordSilentStarts++; },
                End_SilentMode: () => { wordSilentEnds++; },
                MoveCursorToStartOfDocument: noop,
                RecalculateFromStart: noop,
                Document_UpdateInterfaceState: noop,
                Document_UpdateSelectionState: noop,
                theme: null
            },
            m_oDrawingDocument: {TargetStart: noop, CheckGuiControlColors: noop},
            InitControl: noop
        },
        _applyPreOpenLocks: noop,
        onNeedUpdateExternalReferenceOnOpen: noop,
        GenerateStyles: noop,
        sendColorThemes: noop,
        sendEvent: noop,
        sendStandartTextures: noop,
        sendMathToMenu: noop,
        initBroadcastChannelListeners: noop,
        onDocumentContentReady() { wordReadyCalls++; }
    };
    const endOpenSilentMode = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'endOpenSilentMode', {});
    wordApi.endOpenSilentMode = endOpenSilentMode;
    wordEnd.call(wordApi);
    assert.equal(wordApplyCallbacks.length, 1);
    assert.equal(wordSilentStarts, 1);
    assert.equal(wordSilentEnds, 0);
    assert.equal(wordApi.openSilentModeDocument, wordApi.WordControl.m_oLogicDocument);
    // Reopen aborts the load via endOpenSilentMode; stale Apply_Changes must not double-release SilentMode.
    wordApi.documentLoadGeneration = 2;
    wordApi.endOpenSilentMode();
    assert.equal(wordSilentEnds, 1);
    wordApplyCallbacks[0]();
    assert.equal(wordReadyCalls, 0);
    assert.equal(wordSilentEnds, 1, 'stale Apply_Changes must not release SilentMode twice when endOpenSilentMode exists');

    // Fallback path without endOpenSilentMode still releases SilentMode on stale Apply_Changes.
    delete wordApi.endOpenSilentMode;
    wordApi.LoadedObject = 1;
    wordApi.isApplyChangesOnOpenEnabled = true;
    wordEnd.call(wordApi);
    assert.equal(wordApplyCallbacks.length, 2);
    assert.equal(wordSilentStarts, 2);
    wordApi.documentLoadGeneration = 3;
    wordApplyCallbacks[1]();
    assert.equal(wordReadyCalls, 0);
    assert.equal(wordSilentEnds, 2, 'stale Apply_Changes fallback releases SilentMode when endOpenSilentMode is absent');

    // Current generation with endOpenSilentMode still fires readiness cleanly.
    wordApi.endOpenSilentMode = endOpenSilentMode;
    wordApi.LoadedObject = 1;
    wordApi.isApplyChangesOnOpenEnabled = true;
    wordEnd.call(wordApi);
    assert.equal(wordApplyCallbacks.length, 3);
    wordApplyCallbacks[2]();
    assert.equal(wordReadyCalls, 1);
    assert.equal(wordSilentEnds, 3);

    // 2. PDF
    const pdfApplyCallbacks = [];
    let pdfReadyCalls = 0;
    const pdfEnd = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', '_openDocumentEndCallback', {
        Asc: {editor: {asc_registerCallback: noop}},
        AscCommon: {
            PasteElementsId: {},
            sendClientLog: noop,
            getClientInfoString: () => 'info',
            CollaborativeEditing: {
                m_aChanges: [],
                Have_OtherChanges: () => false,
                Apply_Changes: cb => { pdfApplyCallbacks.push(cb); },
                Release_Locks: noop
            }
        },
        performance: {now: () => 0}
    });
    const pdfApi = {
        documentLoadGeneration: 1,
        isDocumentLoadComplete: false,
        ServerImagesWaitComplete: true,
        ServerIdWaitComplete: true,
        isApplyChangesOnOpenEnabled: true,
        isPdfViewer: false,
        isLiveViewer: () => true,
        WordControl: {
            m_oLogicDocument: {},
            m_oDrawingDocument: {m_oDocumentRenderer: {isDocumentContentReady: true}}
        },
        sendMathToMenu: noop,
        sendStandartTextures: noop,
        _applyPreOpenLocks: noop,
        initBroadcastChannelListeners: noop,
        onDocumentContentReady() { pdfReadyCalls++; }
    };
    pdfEnd.call(pdfApi);
    assert.equal(pdfApplyCallbacks.length, 1);
    pdfApi.documentLoadGeneration = 2;
    pdfApplyCallbacks[0]();
    assert.equal(pdfReadyCalls, 0);

    pdfApi.isApplyChangesOnOpenEnabled = true;
    pdfEnd.call(pdfApi);
    assert.equal(pdfApplyCallbacks.length, 2);
    pdfApplyCallbacks[1]();
    assert.equal(pdfReadyCalls, 1);

    // 3. Cell _onApplyChanges
    const cellCallbacks = [];
    let cellDoneCalls = 0;
    const cellApply = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_onApplyChanges', {
        AscCommon: {
            sendClientLog: noop,
            getClientInfoString: () => 'info'
        },
        performance: {now: () => 0}
    });
    const cellApi = {
        documentLoadGeneration: 1,
        isDocumentLoadComplete: false,
        inkDrawer: {startSilentMode: noop, endSilentMode: noop},
        wbModel: {
            DeserializeHistory(_changes, cb) { cellCallbacks.push(cb); }
        }
    };
    cellApply.call(cellApi, [], () => { cellDoneCalls++; });
    assert.equal(cellCallbacks.length, 1);
    cellApi.documentLoadGeneration = 2;
    cellCallbacks[0]();
    assert.equal(cellDoneCalls, 0);

    cellApply.call(cellApi, [], () => { cellDoneCalls++; });
    assert.equal(cellCallbacks.length, 2);
    cellCallbacks[1]();
    assert.equal(cellDoneCalls, 1);
});

test('aborted document loads keep BlockInteraction Open action balanced and clear pending load state', () => {
    const c_oAscAsyncActionType = {Empty: 0, BlockInteraction: 1, Information: 2};
    const c_oAscAsyncAction = {Disconnect: 10, RefreshFile: 11, Open: 12};
    const c_oAscRestrictionType = {None: 0, View: 1};
    const c_oCloseCode = {quiet: 1, sessionIdle: 2, updateVersion: 3, noCache: 4, restore: 5, drop: 6, sessionAbsolute: 7};
    const c_oAscError = {Level: {Critical: -1, NoCritical: 0}, ID: {No: 0, Unknown: 1, UpdateVersion: 2, Warning: 3, Drop: 99}};
    const Asc = {c_oAscAsyncActionType, c_oAscAsyncAction, c_oAscRestrictionType, c_oAscError, c_nMaxConversionTime: 1000};
    const openCryptCallbacks = [];
    const windowObj = {
        IS_NATIVE_EDITOR: true,
        AscDesktopEditor: {
            IsLocalFile: () => false,
            OpenFileCrypt: (_t, _u, cb) => { openCryptCallbacks.push(cb); }
        }
    };
    const baseCommon = {
        ...AscCommon,
        c_oCloseCode,
        c_oAscAdvancedOptionsAction: {Open: 1},
        c_oAscForceSaveTypes: {Button: 1},
        ConnectionState: {ClosedAll: 4},
        CollaborativeEditing: {Clear: noop, Set_GlobalLock: noop, Is_Fast: () => false},
        sendClientLog: noop,
        getClientInfoString: () => 'info',
        getMemoryInfo: () => '',
        mapAscServerErrorToAscError: c => c,
        CDocsCoApi: function() {
            this.disconnectCalls = [];
            this.disconnect = code => { this.disconnectCalls.push(code); };
            this.auth = noop;
            this.get_state = () => 0;
            this.sendClientLog = noop;
        }
    };
    const globals = {
        Asc,
        AscCommon: baseCommon,
        c_oAscAsyncActionType,
        c_oAscAsyncAction,
        c_oAscRestrictionType,
        c_oAscError,
        c_oCloseCode,
        window: windowObj,
        performance: {now: () => 0},
        setInterval: noop,
        setTimeout: noop,
        clearInterval: noop,
        clearTimeout: noop
    };
    const loadBase = name => loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    const openStack = [];
    const cryptCalls = [];
    const events = [];
    const api = {
        documentIsWopi: false,
        DocInfo: {get_Id: () => 'doc-0', get_Encrypted: () => true, get_Title: () => 't', get_Url: () => 'u', asc_getOptions: () => ({}), asc_putTitle: noop},
        isDocumentLoadComplete: false,
        documentLoadGeneration: 0,
        isOpenActionActive: false,
        isDisconnectActionActive: false,
        isOnLoadLicense: true,
        isApplyChangesOnOpenEnabled: true,
        isApplyChangesOnOpen: true,
        isSaveFonts_Images: true,
        asyncMethodCallback: () => {},
        openResult: {stale: true},
        openFileCryptBinary: {stale: true},
        ServerIdWaitComplete: false,
        ServerImagesWaitComplete: false,
        FontLoadWaitComplete: false,
        ServerIdWaitAction: null,
        IsLongActionCurrent: 0,
        LongActionCallbacks: [],
        LongActionCallbacksParams: [],
        IsActionRestrictionCurrent: 0,
        IsActionRestrictionPrev: null,
        waitNotIdleTimeout: null,
        restrictions: c_oAscRestrictionType.None,
        disableRemoveFonts: true,
        isViewMode: false,
        getViewMode: () => false,
        canUnlockDocument: false,
        _getOpenCmd: () => ({}),
        getRestrictionSettings: () => null,
        asc_setRestriction(value) { this.restrictions = value; },
        asc_setDocInfo(info) { this.DocInfo = info; },
        asc_CloseFile() { this.isDocumentLoadComplete = false; },
        isDocumentModified: () => false,
        asc_checkNeedCallback: () => true,
        setOpenedAt: noop,
        _openDocumentEndCallback: noop,
        pluginMethod_SetProperties: noop,
        openFileCryptCallback(stream) { cryptCalls.push(stream); },
        startCollaborationEditing() { events.push(['startCollab']); },
        endCollaborationEditing() { events.push(['endCollab']); },
        sendEvent(name, ...args) {
            events.push([name, ...args]);
            if (name === 'asc_onStartAction' && args[1] === c_oAscAsyncAction.Open) {
                openStack.push(args[0]);
            } else if (name === 'asc_onEndAction' && args[1] === c_oAscAsyncAction.Open) {
                const idx = openStack.lastIndexOf(args[0]);
                if (idx !== -1) openStack.splice(idx, 1);
            }
        },
        CoAuthoringApi: null
    };
    for (const name of [
        ...loadCancellationMethods,
        'sync_StartAction', 'sync_EndAction', 'isActionWithRestriction', 'isLongActionBase',
        'incrementCounterLongAction', 'decrementCounterLongAction',
        'incrementCounterActionRestriction', 'decrementCounterActionRestriction',
        'clearWaitNotIdle', 'endDisconnectAction', 'endOpenAction',
        'asc_LoadDocument', 'onDocumentContentReady', 'canRefreshFile', 'onRefreshFile',
        'asc_refreshFile', 'reopenFileWithReconnection', 'asyncServerIdEndLoaded'
    ]) {
        api[name] = loadBase(name);
    }
    const handlerGlobals = {...globals, t: api};
    const handlers = {};
    for (const name of ['onDisconnect', 'onSession', 'onWarning', 'onMeta', 'onForceSave', 'onStartCoAuthoring', 'onEndCoAuthoring']) {
        handlers[name] = loadMethod('common/apiBase.js', 'this.CoAuthoringApi.', name, handlerGlobals);
    }
    api._coAuthoringInit = function() {
        for (const [name, fn] of Object.entries(handlers)) {
            this.CoAuthoringApi[name] = fn;
        }
    };
    api.CoAuthoringApi = new baseCommon.CDocsCoApi();
    api._coAuthoringInit();

    // Scenario A: initial load doc-0 is interrupted by refreshFile(doc-1), then doc-1 is interrupted by onDisconnect(updateVersion) -> refreshFile(doc-2).
    api.asc_LoadDocument(false);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(openStack.length, 1);
    assert.equal(openCryptCallbacks.length, 1);

    const oldCoApi = api.CoAuthoringApi;
    api.asc_refreshFile({get_Id: () => 'doc-1', get_Encrypted: () => false, asc_getOptions: () => ({}), asc_putTitle: noop});
    assert.equal(api.openResult, null);
    assert.equal(api.openFileCryptBinary, null);
    assert.equal(api.isSaveFonts_Images, false);
    assert.equal(api.isApplyChangesOnOpen, false);
    assert.equal(api.asyncMethodCallback, undefined);
    assert.equal(openStack.length, 0);

    // Stale OpenFileCrypt callback from doc-0 and stale CoAuthoringApi session/meta/warning events must be ignored.
    openCryptCallbacks[0]('stale-stream');
    assert.deepEqual(cryptCalls, []);
    const newCoApi = api.CoAuthoringApi;
    events.length = 0;
    oldCoApi.onSession({code: c_oCloseCode.sessionAbsolute, reason: 'expired', interval: 0});
    oldCoApi.onWarning(10);
    oldCoApi.onMeta({title: 'Stale title'});
    oldCoApi.onForceSave({type: 1, start: true});
    oldCoApi.onStartCoAuthoring(true, false);
    oldCoApi.onEndCoAuthoring(false);
    assert.deepEqual(newCoApi.disconnectCalls, []);
    assert.deepEqual(events, []);

    // doc-1 starts loading, then disconnects with updateVersion before finishing.
    api.asc_LoadDocument(false);
    assert.equal(openStack.length, 1);
    api.CoAuthoringApi.onDisconnect(null, c_oCloseCode.updateVersion);
    assert.equal(openStack.length, 0, 'onRefreshFile must end the interrupted Open action');

    // doc-2 starts loading and completes cleanly.
    api.asc_refreshFile({get_Id: () => 'doc-2', get_Encrypted: () => false, asc_getOptions: () => ({}), asc_putTitle: noop});
    api.asc_LoadDocument(false);
    assert.equal(openStack.length, 1);
    api.asyncServerIdEndLoaded(100);
    api.onDocumentContentReady();
    assert.equal(api.IsLongActionCurrent, 0, 'no leaked BlockInteraction Open long action');
    assert.equal(openStack.length, 0, 'no leaked Open entry in loadMask stack');
    assert.equal(api.isOpenActionActive, false);

    // Scenario B: infinity-loop refresh while Open is active must also clear Open action.
    api.isDocumentLoadComplete = false;
    api.asc_LoadDocument(false);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(openStack.length, 1);
    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(openStack.length, 0);
    assert.equal(api.isOpenActionActive, false);
});

function attachAssetLoaders(harness) {
    const {api, Asc, timers} = harness;
    const pendingFonts = new Set();
    const images = [];
    const window = {};
    const AscFonts = {
        g_font_infos: [],
        ImageLoadStatus: {Loading: 0, Complete: 1},
        CImage: function(src) { this.src = src; },
        g_fontApplication: {
            GetFontInfo: name => ({
                CheckFontLoadStyles: () => pendingFonts.has(name),
                CheckFontLoadStylesNoLoad: () => pendingFonts.has(name)
            })
        }
    };
    vm.runInNewContext(source('common/GlobalLoaders.js'), {
        window, Asc, AscFonts, performance: {now: () => 0},
        AscCommon: {getFullImageSrc2: name => name, backoffOnErrorImg: noop},
        Image: function() { images.push(this); },
        setTimeout: callback => { const id = Symbol(); timers.set(id, callback); return id; },
        clearTimeout: id => timers.delete(id)
    });
    const {g_font_loader: fonts, g_image_loader: loader} = window.AscCommon;
    Object.assign(api, {
        FontLoader: fonts,
        ImageLoader: loader,
        OpenDocumentProgress: {CurrentFont: 0, CurrentImage: 0},
        sync_InitEditorFonts: noop,
        IsNeedDefaultFonts: () => false,
        SendOpenProgress: noop,
        asyncFontsDocumentStartLoaded: noop,
        asyncImagesDocumentStartLoaded: noop,
        asyncFontEndLoaded: noop
    });
    fonts.put_Api(api);
    loader.Api = api;
    return {fonts, loader, pendingFonts, images, FontLoader: window.AscCommon.CGlobalFontLoader};
}

test('reopen invalidates a pending document font poll and starts the next file immediately', () => {
    const harness = createBaseRefreshHarness();
    const {api, timers} = harness;
    const {fonts, pendingFonts} = attachAssetLoaders(harness);
    let completed = 0;
    api.isDocumentLoadComplete = false;
    api.asyncFontsDocumentEndLoaded = () => completed++;
    pendingFonts.add('old');
    fonts.LoadDocumentFonts([{name: 'old'}]);
    const stalePoll = [...timers.values()][0];

    api.reopenFileWithReconnection({get_Id: () => 'next'});
    assert.equal(fonts.isWorking(), false);
    assert.equal(fonts.fonts_loading.length, 0);

    pendingFonts.add('next');
    fonts.LoadDocumentFonts2([{name: 'next'}]);
    stalePoll();
    assert.equal(fonts.isWorking(), true);
    assert.equal(completed, 0);
    pendingFonts.delete('next');
    harness.flushTimers();
    assert.equal(completed, 1);
});

for (const asyncImages of [false, true]) {
    test(`reopen ignores stale ${asyncImages ? 'mixed' : 'blocking'} document image notifications`, () => {
        const harness = createBaseRefreshHarness();
        const {api} = harness;
        const {loader, images} = attachAssetLoaders(harness);
        let completed = 0;
        api.isDocumentLoadComplete = false;
        api.asyncImagesDocumentEndLoaded = () => completed++;
        loader.bIsLoadDocumentFirst = true;
        loader.bIsAsyncLoadDocumentImages = asyncImages;
        loader.LoadDocumentImages(['old', 'old-background'], false, {old: true});
        const previous = images.slice();

        api.reopenFileWithReconnection({get_Id: () => 'next'});
        for (const image of previous) image.onload();
        assert.equal(api.OpenDocumentProgress.CurrentImage, 0);
        assert.equal(completed, 0);
        assert.equal(images.length, previous.length, 'the cancelled batch cannot start background images');

        loader.bIsLoadDocumentFirst = true;
        loader.LoadDocumentImages(['next'], false, {next: true});
        images.at(-1).onload();
        assert.equal(api.OpenDocumentProgress.CurrentImage, 1);
        assert.equal(completed, 1);
    });
}

test('reopen closes a blocked ApplyChanges and clears partial document state', () => {
    const {api, events, c_oAscAsyncActionType, c_oAscAsyncAction, common} = createBaseRefreshHarness();
    const order = [];
    const oldDocument = {
        RecalcInfo: {Paused: true},
        End_SilentMode: () => order.push('silent'),
        StopRecalculate: () => order.push('recalc'),
        StopSpellCheck: () => order.push('spell')
    };
    api.WordControl = {m_oLogicDocument: oldDocument};
    api.isDocumentLoadComplete = false;
    api.openSilentModeDocument = oldDocument;
    api.isSaveFonts_Images = true;
    api.isApplyChangesActionActive = true;
    Object.assign(common.CollaborativeEditing, {
        m_oLogicDocument: oldDocument,
        m_fEndLoadCallBack: () => assert.fail('obsolete callback'),
        m_aEndLoadCallbacks: [() => assert.fail('obsolete callback')],
        m_bApplyChangesInProgress: true,
        m_bRecalculatePausedByApplyChanges: true,
        m_bGlobalLock: 1,
        m_bGlobalLockSelection: 1,
        Set_GlobalLock(value) { this.m_bGlobalLock += value ? 1 : -1; },
        Set_GlobalLockSelection(value) { this.m_bGlobalLockSelection += value ? 1 : -1; },
        Clear() {
            assert.equal(this.m_bGlobalLock, 0);
            assert.equal(this.m_bGlobalLockSelection, 0);
            order.push('clear');
        }
    });
    const collab = common.CollaborativeEditing;
    collab.abortPendingApplyChanges = loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', 'abortPendingApplyChanges');
    api.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.ApplyChanges);

    api.reopenFileWithReconnection({get_Id: () => 'next'});
    assert.deepEqual(order, ['silent', 'recalc', 'spell', 'clear']);
    assert.equal(oldDocument.RecalcInfo.Paused, false);
    assert.equal(collab.m_fEndLoadCallBack, null);
    assert.equal(collab.m_aEndLoadCallbacks.length, 0);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(events.filter(event => event[0] === 'asc_onEndAction' && event[2] === c_oAscAsyncAction.ApplyChanges).length, 1);
});

test('requesting a refresh keeps the live document and ignores superseded WOPI replies', () => {
    const {api, Asc, events} = createBaseRefreshHarness();
    const requests = [];
    const refreshed = [];
    api.documentIsWopi = true;
    api.documentLoadGeneration = 3;
    api.FontLoader = {cancelDocumentLoad: () => assert.fail('live font load was cancelled')};
    api.isPasteFonts_Images = true;
    api.incrementCounterLongAction();
    api.DocInfo.extendWithWopiParams = response => ({get_Title: () => response});
    api.CoAuthoringApi.callPRC = (_data, _timeout, callback) => { requests.push(callback); return true; };
    api.CoAuthoringApi.onMeta = noop;
    api.asc_refreshFile = info => refreshed.push(info.get_Title());

    api.onRefreshFile();
    api.onRefreshFile();
    requests[0](false, 'stale');
    assert.deepEqual(refreshed, []);
    requests[1](false, 'current');
    assert.deepEqual(refreshed, ['current']);
    assert.equal(api.documentLoadGeneration, 3);
    assert.equal(api.isPasteFonts_Images, true);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(events.some(event => event[0] === 'asc_onStartAction' && event[1] === Asc.c_oAscAsyncActionType.BlockInteraction), false);
});

test('a WOPI refresh timeout restores the ready document restrictions', () => {
    const {api, Asc, getGlobalLock, events} = createBaseRefreshHarness();
    api.documentIsWopi = true;
    const requests = [];
    api.CoAuthoringApi.callPRC = (_data, _timeout, callback) => { requests.push(callback); return true; };
    api.onRefreshFile();
    api.onRefreshFile();

    requests[0](true, undefined);
    assert.equal(api.IsActionRestrictionCurrent, 1, 'a stale timeout cannot interrupt the newer request');
    requests[1](true, undefined);

    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, Asc.c_oAscRestrictionType.None);
    assert.equal(getGlobalLock(), false);
    assert.equal(api.isDocumentLoadComplete, true);
    assert.deepEqual(events.filter(event => event[0] === 'asc_onError'), [
        ['asc_onError', Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical]
    ]);
});

test('PDF cancels an old readiness poll and drops pending file bytes without resetting the engine', () => {
    const timers = [];
    const globals = {
        setTimeout: fn => { timers.push(fn); return timers.length; },
        clearTimeout: noop,
        ModuleState: {Loaded: 7},
        window: {AscViewer: {InitializeFonts: noop}}
    };
    const viewerMethod = name => loadMethod('pdf/src/viewer.js', 'this.', name, globals);
    const pdfDoc = {defaultFontsLoaded: 0, fontLoader: {isWorking: () => false}};
    const events = [];
    const viewer = {
        Api: {},
        checkReady: viewerMethod('checkReady'),
        getPDFDoc: () => pdfDoc,
        openSessionId: 0,
        _fileData: 'obsolete bytes',
        isDocumentContentReady: true,
        isDocumentReady: true,
        isStarted: true,
        pageDetector: {},
        initPaintDone: true,
        IsOpenFormsInProgress: false,
        id_main: {style: {}},
        file: {isValid: () => false, pages: []},
        sendEvent: name => events.push(name),
        open: () => assert.fail('obsolete bytes were reopened')
    };
    viewer.cancelPendingLoad = viewerMethod('cancelPendingLoad');
    viewer.checkReady();
    viewerMethod('close').call(viewer);
    timers[0]();
    assert.deepEqual(events, []);
    assert.equal(viewer._fileData, null);
    assert.equal(viewer.isDocumentReady, false);
    assert.equal(viewer.isDocumentContentReady, false);
    viewerMethod('onLoadModule').call(viewer);
    assert.equal(viewer.moduleState, 7);
    assert.equal(viewer.initPaintDone, true);

    viewer.sendEvent = name => {
        events.push(name);
        if (name === 'onFileOpened') viewer.cancelPendingLoad();
    };
    viewer.checkReady();
    timers[1]();
    assert.deepEqual(events, ['onFileOpened'], 'a synchronous reopen suppresses the old page and structure events');
});

test('a synchronous font completion preserves the next font batch and its action type', () => {
    const harness = createBaseRefreshHarness();
    const {api, Asc, events} = harness;
    const {fonts, pendingFonts} = attachAssetLoaders(harness);
    const {BlockInteraction, Information} = Asc.c_oAscAsyncActionType;
    const {LoadDocumentFonts, LoadFont} = Asc.c_oAscAsyncAction;
    const completed = [];
    api.asyncFontsDocumentStartLoaded = type => api.startFontLoadAction(
        type ?? BlockInteraction, type === Information ? LoadFont : LoadDocumentFonts
    );
    api.asyncFontsDocumentEndLoaded = type => {
        api.endFontLoadAction();
        completed.push(type);
        if (type === undefined) {
            pendingFonts.add('changes');
            fonts.LoadDocumentFonts2([{name: 'changes'}], Information);
        }
    };
    fonts.LoadDocumentFonts([]);
    assert.equal(fonts.isWorking(), true);
    pendingFonts.delete('changes');
    harness.flushTimers();
    assert.deepEqual(completed, [undefined, Information]);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.deepEqual(events.filter(e => e[0] === 'asc_onEndAction'), [
        ['asc_onEndAction', BlockInteraction, LoadDocumentFonts],
        ['asc_onEndAction', Information, LoadFont]
    ]);
});

test('standalone font previews release their own actions and theme loaders still work after refresh', () => {
    const harness = createBaseRefreshHarness();
    const {api, Asc, events} = harness;
    const {fonts, pendingFonts, FontLoader} = attachAssetLoaders(harness);
    const preview = new FontLoader();
    preview.put_Api(api);
    const {BlockInteraction, Information} = Asc.c_oAscAsyncActionType;
    const {LoadDocumentFonts, LoadFont} = Asc.c_oAscAsyncAction;
    const finished = [];
    api.isDocumentLoadComplete = false;
    pendingFonts.add('preview');
    preview.LoadDocumentFonts2([{name: 'preview'}], Information, () => assert.fail('stale preview'));
    api.reopenFileWithReconnection({get_Id: () => 'next'});
    api.startFontLoadAction(BlockInteraction, LoadFont);
    harness.flushTimers();
    assert.deepEqual(events.filter(e => e[0] === 'asc_onEndAction'), [
        ['asc_onEndAction', Information, LoadDocumentFonts]
    ]);
    assert.equal(api.IsLongActionCurrent, 1);

    preview.LoadDocumentFonts2([], Information, () => finished.push('preview'));
    fonts.ThemeLoader = {
        asyncFontsStartLoaded: () => finished.push('theme start'),
        asyncFontsEndLoaded: () => finished.push('theme end')
    };
    fonts.LoadDocumentFonts2([]);
    api.endFontLoadAction();
    assert.deepEqual(finished, ['preview', 'theme start', 'theme end']);
    assert.equal(api.IsLongActionCurrent, 0);
});

test('closing Word invalidates pathurls and balances custom images, paste and version history actions', () => {
    const harness = createBaseRefreshHarness();
    const {api, Asc, common, events} = harness;
    const {loader, images} = attachAssetLoaders(harness);
    const {BlockInteraction, Information} = Asc.c_oAscAsyncActionType;
    const {Open, LoadImage, ApplyChanges} = Asc.c_oAscAsyncAction;
    const globals = {
        Asc, AscCommon: common, g_oIdCounter: common.g_oIdCounter, g_oTableId: common.g_oTableId,
        c_oAscAsyncActionType: Asc.c_oAscAsyncActionType, c_oAscAsyncAction: Asc.c_oAscAsyncAction
    };
    const wordMethod = name => loadMethod('word/api.js', 'asc_docs_api.prototype.', name, globals);
    const collabMethod = name => loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', name, globals);
    const callbacks = [];
    common.g_oDocumentUrls = {addUrls: () => assert.fail('obsolete URLs were applied')};
    const collab = common.CollaborativeEditing;
    Object.assign(collab, {
        GetEditorApi: () => api, m_aNewImages: ['old'], CheckWaitingImages: noop,
        abortPendingApplyChanges: collabMethod('abortPendingApplyChanges'),
        SendImagesCallback: () => assert.fail('obsolete assets loaded')
    });
    api.CoAuthoringApi.callPRC = (_data, _timeout, cb) => { callbacks.push(cb); return true; };
    api.WordControl = {
        m_oLogicDocument: {StopRecalculate: noop, StopSpellCheck: noop},
        m_oDrawingDocument: {CloseFile: noop}
    };
    api.turnOffSpecialModes = noop;
    api.asyncImagesDocumentEndLoaded = () => assert.fail('obsolete custom images completed');
    api._openVersionHistoryEndCallback = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_openVersionHistoryEndCallback', globals);
    api.startApplyChangesAction();
    api.isApplyChangesOnVersionHistory = true;
    api.sync_StartAction(BlockInteraction, Open);
    api.isPasteFonts_Images = true;
    api.pasteCallback = () => assert.fail('obsolete paste completed');
    api.incrementCounterLongAction();
    collabMethod('SendImagesUrlsFromChanges').call(collab, ['old']);

    wordMethod('SyncLoadImages').call(api, ['custom']);
    assert.equal(loader.map_image_index.custom.Image, images[0]);
    wordMethod('asc_CloseFile').call(api);
    callbacks[0](false, {status: 'ok', data: ['old-url']});
    images[0].onload();
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(api.isApplyChangesOnVersionHistory, false);
    assert.equal(api.isLoadImagesCustom, false);
    assert.equal(api.pasteCallback, null);
    assert.deepEqual(events.filter(e => e[0] === 'asc_onEndAction'), [
        ['asc_onEndAction', Information, LoadImage],
        ['asc_onEndAction', BlockInteraction, ApplyChanges],
        ['asc_onEndAction', BlockInteraction, Open]
    ]);
});

for (const [file, prototype] of [
    ['word/Editor/CollaborativeEditing.js', 'CWordCollaborativeEditing.prototype.'],
    ['pdf/src/CollaborativeEditing.js', 'CPDFCollaborativeEditing.prototype.']
]) {
    test(`${file.split('/')[0]} does not resume recalculation after a synchronous reopen in ApplyChanges`, () => {
        const harness = createBaseRefreshHarness();
        const {api, common} = harness;
        const oldDoc = {RecalcInfo: {Paused: true}, ResumeRecalculate: () => assert.fail('obsolete recalculation')};
        const collab = common.CollaborativeEditing;
        api.isDocumentLoadComplete = false;
        Object.assign(collab, {
            m_oLogicDocument: oldDoc, m_bRecalculatePausedByApplyChanges: true,
            m_fEndLoadCallBack: () => api.reopenFileWithReconnection({get_Id: () => 'next'}),
            GetEditorApi: () => api, Set_GlobalLock: noop, Set_GlobalLockSelection: noop,
            m_aEndLoadCallbacks: [() => assert.fail('obsolete PDF callback')],
            abortPendingApplyChanges: loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', 'abortPendingApplyChanges')
        });
        api.startApplyChangesAction();
        loadMethod(file, prototype, 'OnEnd_Load_Objects').call(collab);
        assert.equal(oldDoc.RecalcInfo.Paused, false);
        assert.equal(api.IsLongActionCurrent, 0);
        assert.equal(collab.m_aEndLoadCallbacks.length, 0);
    });
}

test('Slide stops the old opening sequence when applying changes synchronously reopens the file', () => {
    const {api, common} = createBaseRefreshHarness();
    Object.assign(common.CollaborativeEditing, {
        m_aChanges: [],
        Apply_Changes: () => api.reopenFileWithReconnection({get_Id: () => 'next'}),
        Release_Locks: () => assert.fail('old Slide open continued')
    });
    Object.assign(api, {
        isDocumentLoadComplete: false, ServerIdWaitComplete: true, ServerImagesWaitComplete: true,
        LoadedObject: 1, DocumentType: 2, WordControl: {m_oLogicDocument: {}},
        _applyPreOpenLocks: noop, onDocumentContentReady: () => assert.fail('old Slide became ready')
    });
    const openEnd = loadMethod('slide/api.js', 'asc_docs_api.prototype.', '_openDocumentEndCallback', {AscCommon: common, performance});
    openEnd.call(api);
    assert.equal(api.bNoSendComments, false);
    assert.equal(api.isApplyChangesOnOpen, false);
});

test('same-key refresh clears queued Cell changes when it gives up on an incomplete file', () => {
    const {api, common} = createBaseRefreshHarness();
    const collab = common.CollaborativeEditing;
    collab.Clear = loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', 'Clear');
    collab.CoHistory = {clear: noop};
    collab.m_arrChanges = ['old cell changes'];
    collab.m_aChanges = ['old shapes'];
    api.isDocumentLoadComplete = false;

    api.asc_refreshFile(api.DocInfo);
    assert.equal(collab.m_arrChanges.length, 0);
    assert.equal(collab.m_aChanges.length, 0);
});

for (const closeVisio of [false, true]) {
    test(`${closeVisio ? 'closing Visio' : 'clearing an incomplete document'} restarts object IDs before registering the new table`, () => {
        const {api, common} = createBaseRefreshHarness();
        common.History = {Clear: noop, Add: noop};
        common.CChangesTableIdAdd = function() {};
        const globals = {AscCommon: common};
        const tableMethod = name => loadMethod('common/TableId.js', 'CTableId.prototype.', name, globals);
        const table = {IsOn: tableMethod('IsOn'), Add: tableMethod('Add'), Clear: tableMethod('Clear'), Get_ById: tableMethod('Get_ById')};
        const counter = {
            m_bLoad: true,
            m_nIdCounterLoad: 4,
            Get_NewId: loadMethod('common/editorscommon.js', 'CIdCounter.prototype.', 'Get_NewId', globals),
            Clear: loadMethod('common/editorscommon.js', 'CIdCounter.prototype.', 'Clear', globals)
        };
        Object.assign(common, {g_oIdCounter: counter, g_oTableId: table});
        api.isDocumentLoadComplete = false;
        api.turnOffSpecialModes = noop;

        if (closeVisio) {
            loadMethod('visio/api.js', 'VisioEditorApi.prototype.', 'asc_CloseFile', globals).call(api);
        } else {
            api.reopenFileWithReconnection({get_Id: () => 'next'});
        }
        assert.equal(table.Id, '1');
        const firstObject = {};
        table.Add(firstObject, counter.Get_NewId());
        assert.equal(firstObject.Id, '2');
        assert.equal(table.Get_ById('1'), table);
    });
}

test('PDF ignores a CMap response for the replaced file before the current file is ready', () => {
    const requests = [];
    const globals = {
        clearTimeout: noop,
        XMLHttpRequest: function() {
            requests.push(this);
            this.status = 200;
            this.response = new ArrayBuffer(0);
            this.open = this.send = this.setRequestHeader = noop;
        }
    };
    const completed = [];
    const file = name => ({
        name, isNeedCMap: () => true,
        setCMap() { completed.push(`CMap for ${name}`); }
    });
    const viewer = {
        Api: {},
        file: file('old'),
        openSessionId: 0,
        isCMapLoading: false,
        getPDFDoc: () => null,
        checkLoadCMap: loadMethod('pdf/src/viewer.js', 'this.', 'checkLoadCMap', globals),
        cancelPendingLoad: loadMethod('pdf/src/viewer.js', 'this.', 'cancelPendingLoad', globals),
        onDocumentReady() { completed.push(`ready for ${this.file.name}`); }
    };
    viewer.checkLoadCMap();
    viewer.cancelPendingLoad();
    viewer.file = file('next');
    viewer.checkLoadCMap();

    requests[0].onload();
    assert.deepEqual(completed, []);
    assert.equal(viewer.isCMapLoading, true);
    requests[1].onload();
    assert.deepEqual(completed, ['CMap for next', 'ready for next']);
    assert.equal(viewer.isCMapLoading, false);
});

test('Cell can refresh before the first workbook has initialized history', () => {
    const {api, common} = createBaseRefreshHarness();
    const globals = {window: {AscCommon: {g_specialPasteHelper: {SpecialPasteButton_Hide: noop}}}};
    const historyMethod = name => loadMethod('cell/model/History.js', 'CHistory.prototype.', name, globals);
    common.History = {
        Points: [], Is_Clear: historyMethod('Is_Clear'), Clear: historyMethod('Clear'),
        _sendCanUndoRedo: noop
    };
    api.editorId = common.c_oEditorId.Spreadsheet;
    api.wbModel = null;
    api.isDocumentLoadComplete = false;

    api.asc_refreshFile({get_Id: () => 'next'});
    assert.equal(api.DocInfo.get_Id(), 'next');

    const events = [];
    api.wbModel = common.History.workbook = {handlers: {trigger: name => events.push(name)}};
    common.History.Points.push('old change');
    api.asc_refreshFile({get_Id: () => 'final'});
    assert.equal(common.History.Points.length, 0);
    assert.deepEqual(events, ['toggleAutoCorrectOptions']);
    assert.equal(api.DocInfo.get_Id(), 'final');
});

test('new collaboration changes wait until the active ApplyChanges batch releases its locks', () => {
    const {api, Asc, common} = createBaseRefreshHarness();
    const globals = {Asc, AscCommon: common, AscFonts: {}, window: {}};
    const collabMethod = name => loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', name, globals);
    const collab = common.CollaborativeEditing;
    const changes = [];
    const doc = {
        RecalcInfo: {Paused: false},
        PauseRecalculate() { this.RecalcInfo.Paused = true; },
        EndPreview_MailMergeResult: noop
    };
    api.WordControl = {m_oLogicDocument: doc};
    Object.assign(collab, {
        m_oLogicDocument: doc,
        m_aChanges: ['first'],
        m_bGlobalLock: 0,
        m_bGlobalLockSelection: 0,
        GetEditorApi: () => api,
        Set_GlobalLock: collabMethod('Set_GlobalLock'),
        Set_GlobalLockSelection: collabMethod('Set_GlobalLockSelection'),
        Apply_Changes: collabMethod('Apply_Changes'),
        OnStart_Load_Objects: collabMethod('OnStart_Load_Objects'),
        abortPendingApplyChanges: collabMethod('abortPendingApplyChanges'),
        CoHistory: {InitTextRecover: noop},
        Apply_OtherChanges() { changes.push(...this.m_aChanges.splice(0)); },
        private_SaveDocumentState: noop,
        private_RestoreDocumentState: noop,
        Clear_NewImages: noop,
        Lock_NeedLock: noop,
        CollectImagesFromChanges: () => ['still-pending'],
        SendImagesUrlsFromChanges: noop
    });

    collab.Apply_Changes();
    collab.m_aChanges.push('second');
    collab.Apply_Changes();
    assert.deepEqual(changes, ['first']);
    assert.deepEqual(collab.m_aChanges, ['second']);
    assert.equal(collab.m_bGlobalLock, 1);
    assert.equal(collab.m_bGlobalLockSelection, 1);

    api.reopenFileWithReconnection({get_Id: () => 'next'});
    assert.equal(collab.m_bGlobalLock, 0);
    assert.equal(collab.m_bGlobalLockSelection, 0);
    assert.equal(api.IsLongActionCurrent, 0);
});

test('reopening drops deferred locks from the old connection and queues locks for the new file', () => {
    const {api} = createBaseRefreshHarness();
    const baseMethod = name => loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name);
    const lock = baseMethod('_coAuthoringCheckEndOpenDocument');
    const apply = baseMethod('_applyPreOpenLocks');
    api.isPreOpenLocks = true;
    api.arrPreOpenLocksObjects = [];
    assert.equal(lock.call(api, () => assert.fail('old lock was applied to the new document')), true);

    const ids = [];
    api.reopenFileWithReconnection({get_Id: () => 'next'});
    assert.equal(lock.call(api, id => ids.push(id), 'next'), true);
    apply.call(api);
    assert.deepEqual(ids, ['next']);

    api.reopenFileWithReconnection({get_Id: () => 'final'});
    assert.equal(lock.call(api, id => ids.push(id), 'final'), true);
    apply.call(api);
    assert.deepEqual(ids, ['next', 'final']);
});

test('closing a document releases its pending refresh restriction', () => {
    const {api, Asc, common, getGlobalLock} = createBaseRefreshHarness();
    api.turnOffSpecialModes = noop;
    api.onRefreshFile();
    assert.equal(getGlobalLock(), true);

    loadMethod('visio/api.js', 'VisioEditorApi.prototype.', 'asc_CloseFile', {AscCommon: common}).call(api);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, Asc.c_oAscRestrictionType.None);
    assert.equal(getGlobalLock(), false);
});

test('an interrupted Cell font load releases formula localization and history without applying old changes', () => {
    const harness = createBaseRefreshHarness();
    const {api, common} = harness;
    const {fonts, pendingFonts} = attachAssetLoaders(harness);
    const locale = {Parse: false, DigitSep: false};
    const history = {TurnOn: () => { history.disabled = false; }, disabled: true};
    const globals = {
        AscCommon: {...common, baseEditorsApi: {prototype: {_abortPendingDocumentLoad: api._abortPendingDocumentLoad}}},
        AscCommonExcel: {oFormulaLocaleInfo: locale}, History: history
    };
    const file = 'cell/api.js';
    api._abortPendingDocumentLoad = loadMethod(file, 'spreadsheet_api.prototype.', '_abortPendingDocumentLoad', globals);
    api.wbModel = {bCollaborativeChanges: true};
    api.asyncMethodCallback = () => assert.fail('old history was applied');
    api.isDocumentLoadComplete = false;
    pendingFonts.add('workbook changes');
    fonts.LoadDocumentFonts2([{name: 'workbook changes'}]);

    api.reopenFileWithReconnection({get_Id: () => 'next'});
    assert.deepEqual(locale, {Parse: true, DigitSep: true});
    assert.equal(history.disabled, false);
    assert.equal(api.wbModel.bCollaborativeChanges, false);
    assert.equal(api.asyncMethodCallback, undefined);
});

test('closing a password-protected PDF prevents its delayed prompt from appearing over the next file', () => {
    const timers = [];
    const globals = {
        setTimeout: fn => timers.push(fn),
        clearTimeout: noop,
        window: {
            AscViewer: {
                createFile: passwordRequired => ({
                    isNeedPassword: () => passwordRequired,
                    isValid: () => false
                })
            }
        }
    };
    const viewerMethod = name => loadMethod('pdf/src/viewer.js', 'this.', name, globals);
    const events = [];
    const viewer = {
        openSessionId: 0,
        getPDFDoc: () => null,
        checkModule: () => true,
        open: viewerMethod('open'),
        close: viewerMethod('close'),
        cancelPendingLoad: viewerMethod('cancelPendingLoad'),
        afterOpen: () => events.push('open'),
        sendEvent: name => events.push(name)
    };

    viewer.open(true);
    viewer.open(false);
    timers[0]();
    assert.deepEqual(events, ['open']);
});

test('the host refresh keeps its restriction through close until the new connection authorizes', () => {
    const {api, Asc, common} = createBaseRefreshHarness();
    api.turnOffSpecialModes = noop;
    api.asc_CloseFile = loadMethod('visio/api.js', 'VisioEditorApi.prototype.', 'asc_CloseFile', {AscCommon: common});
    api.onRefreshFile();
    const waitAction = api.ServerIdWaitAction;

    api.asc_refreshFile({get_Id: () => 'next'});
    assert.equal(api.ServerIdWaitAction, waitAction);
    assert.equal(api.IsActionRestrictionCurrent, 1);
    assert.equal(api.restrictions, Asc.c_oAscRestrictionType.View);

    api.asyncServerIdEndLoaded(100);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsActionRestrictionCurrent, 0);
    assert.equal(api.restrictions, Asc.c_oAscRestrictionType.None);
});

test('terminal disconnect during live refreshFile, re-entrant RefreshFile end, and Open ClosedAll preserve view locks and event order', () => {
    // 1. Live document (isDocumentLoadComplete === true) under onRefreshFile receives fatal disconnect.
    const live = createBaseRefreshHarness();
    live.api.isViewMode = false;
    live.api.asc_setRestriction = function(value) {
        this.restrictions = value;
        this.sendEvent('asc_onChangeRestrictions', value);
    };
    live.api.asc_setViewMode = function(isViewMode) {
        this.isViewMode = Boolean(isViewMode);
        if (isViewMode) {
            live.common.CollaborativeEditing.Set_GlobalLock(true);
        }
    };
    live.api.setViewModeDisconnect = function(enableDownload) {
        this.sendEvent('asc_onCoAuthoringDisconnect', enableDownload);
        this.asc_setViewMode(true);
    };
    live.api.onRefreshFile();
    live.events.length = 0;
    live.api.CoAuthoringApi.onDisconnect('drop', live.c_oCloseCode.drop);
    assert.equal(live.api.ServerIdWaitAction, null);
    assert.equal(live.api.IsActionRestrictionCurrent, 0);
    assert.equal(live.api.isViewMode, true);
    assert.equal(live.getGlobalLock(), true);
    const liveEnd = live.events.findIndex(e => e[0] === 'asc_onEndAction' && e[2] === live.c_oAscAsyncAction.RefreshFile);
    const liveRestr = live.events.findIndex(e => e[0] === 'asc_onChangeRestrictions' && e[1] === live.c_oAscRestrictionType.None);
    const liveDisc = live.events.findIndex(e => e[0] === 'asc_onCoAuthoringDisconnect');
    const liveErr = live.events.findIndex(e => e[0] === 'asc_onError' && e[2] === live.Asc.c_oAscError.Level.NoCritical);
    assert.ok(liveEnd >= 0 && liveRestr >= 0 && liveRestr < liveDisc && liveEnd < liveDisc && liveDisc < liveErr);

    // 2. Re-entrant asc_onEndAction(RefreshFile) calling endServerIdWaitAction does not double-end,
    //    and calling onRefreshFile keeps the newly created ServerIdWaitAction intact (fatal and sessionIdle).
    for (const mode of ['fatal-close', 'fatal-restart', 'idle-restart']) {
        const reentrant = createBaseRefreshHarness();
        let refreshEndCount = 0;
        let triggered = false;
        const origSendEvent = reentrant.api.sendEvent.bind(reentrant.api);
        reentrant.api.sendEvent = function(name, ...args) {
            origSendEvent(name, ...args);
            if (!triggered && name === 'asc_onEndAction' && args[1] === reentrant.c_oAscAsyncAction.RefreshFile) {
                refreshEndCount++;
                triggered = true;
                if (mode === 'fatal-close') {
                    this.endServerIdWaitAction();
                } else {
                    this.onRefreshFile();
                }
            } else if (name === 'asc_onEndAction' && args[1] === reentrant.c_oAscAsyncAction.RefreshFile) {
                refreshEndCount++;
            }
        };
        reentrant.api.asc_refreshFile({get_Id: () => 'doc-2'});
        reentrant.api.CoAuthoringApi.get_state = () => reentrant.ConnectionState.ClosedAll;
        if (mode === 'idle-restart') {
            reentrant.api.CoAuthoringApi.onDisconnect(null, reentrant.c_oCloseCode.sessionIdle);
            assert.equal(refreshEndCount, 1);
            assert.ok(reentrant.api.ServerIdWaitAction && reentrant.api.ServerIdWaitAction[1] === reentrant.c_oAscAsyncAction.RefreshFile);
            assert.equal(reentrant.api.IsActionRestrictionCurrent, 1);
        } else {
            reentrant.api.CoAuthoringApi.onDisconnect('drop', reentrant.c_oCloseCode.drop);
            assert.equal(refreshEndCount, 1);
            if (mode === 'fatal-close') {
                assert.equal(reentrant.api.ServerIdWaitAction, null);
                assert.equal(reentrant.api.IsActionRestrictionCurrent, 0);
            } else {
                assert.ok(reentrant.api.ServerIdWaitAction && reentrant.api.ServerIdWaitAction[1] === reentrant.c_oAscAsyncAction.RefreshFile);
                assert.equal(reentrant.api.IsActionRestrictionCurrent, 1);
            }
        }
    }

    // 3. ServerIdWaitAction === Open on fatal ClosedAll still finishes via asyncServerIdEndLoaded before disconnect.
    const openWait = createBaseRefreshHarness();
    let openServerReadyCount = 0;
    openWait.api.isDocumentLoadComplete = false;
    openWait.api.ServerIdWaitComplete = false;
    openWait.api.ServerIdWaitAction = [openWait.c_oAscAsyncActionType.BlockInteraction, openWait.c_oAscAsyncAction.Open];
    openWait.api.sync_StartAction(...openWait.api.ServerIdWaitAction);
    openWait.api.setViewModeDisconnect = function(enableDownload) {
        this.sendEvent('asc_onCoAuthoringDisconnect', enableDownload);
    };
    const origAsyncServerIdEndLoaded = openWait.api.asyncServerIdEndLoaded.bind(openWait.api);
    openWait.api.asyncServerIdEndLoaded = function(...args) {
        openServerReadyCount++;
        return origAsyncServerIdEndLoaded(...args);
    };
    openWait.api.CoAuthoringApi.get_state = () => openWait.ConnectionState.ClosedAll;
    openWait.events.length = 0;
    openWait.api.CoAuthoringApi.onDisconnect('drop', openWait.c_oCloseCode.drop);
    assert.equal(openServerReadyCount, 1);
    assert.equal(openWait.api.ServerIdWaitAction, null);
    const openEnd = openWait.events.findIndex(e => e[0] === 'asc_onEndAction' && e[2] === openWait.c_oAscAsyncAction.Open);
    const openDisc = openWait.events.findIndex(e => e[0] === 'asc_onCoAuthoringDisconnect');
    assert.ok(openEnd >= 0 && openDisc > openEnd);
});

test('ClosedAll disconnect with allowRefresh or after sessionIdle during refresh does not emit premature readiness', () => {
    // 1. Initial load (ServerIdWaitAction === null) receiving ClosedAll + allowRefresh goes to onRefreshFile without readiness.
    {
        const {api, events, ConnectionState, c_oCloseCode} = createBaseRefreshHarness();
        let readyCalls = 0;
        api.isDocumentLoadComplete = false;
        api.ServerIdWaitComplete = false;
        api._openDocumentEndCallback = () => { readyCalls++; };
        api.CoAuthoringApi.get_state = () => ConnectionState.ClosedAll;
        api.CoAuthoringApi.onDisconnect('update version', c_oCloseCode.updateVersion);
        assert.equal(readyCalls, 0);
        assert.equal(api.ServerIdWaitComplete, false);
        assert.equal(events.filter(e => e[0] === 'asc_onRequestRefreshFile').length, 1);
    }

    // 2. Refresh interrupted by sessionIdle clears ServerIdWaitAction, then ClosedAll + updateVersion or fatal drop must not emit readiness.
    for (const closeCode of ['updateVersion', 'drop']) {
        const {api, events, ConnectionState, c_oCloseCode, Asc} = createBaseRefreshHarness();
        let readyCalls = 0;
        api._openDocumentEndCallback = () => { readyCalls++; };
        api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
        api.CoAuthoringApi.onDisconnect('idle', c_oCloseCode.sessionIdle);
        assert.equal(api.ServerIdWaitAction, null);
        assert.equal(api.isRefreshInProgress, true);
        events.length = 0;

        api.CoAuthoringApi.get_state = () => ConnectionState.ClosedAll;
        api.CoAuthoringApi.onDisconnect(closeCode, c_oCloseCode[closeCode]);
        assert.equal(readyCalls, 0);
        assert.equal(api.ServerIdWaitComplete, false);
        if (closeCode === 'updateVersion') {
            assert.equal(events.filter(e => e[0] === 'asc_onRequestRefreshFile').length, 1);
        } else {
            assert.equal(api.isRefreshInProgress, false);
            assert.deepEqual(events.filter(e => e[0] === 'asc_onError'), [
                ['asc_onError', Asc.c_oAscError.ID.Drop, Asc.c_oAscError.Level.Critical]
            ]);
        }
    }
});

test('PDF ignores stale onFileOpened from a previous document generation after reopen', () => {
    const callbacks = new Map();
    const renderer = {
        isDocumentContentReady: false,
        registerEvent(name, fn) { callbacks.set(name, fn); },
        open: noop,
        close: noop,
        cancelPendingLoad() { this.isDocumentContentReady = false; }
    };
    const pdfCommon = {
        ...AscCommon,
        CViewer: function() { return renderer; },
        InitBrowserInputContext: noop,
        sendClientLog: noop,
        getClientInfoString: () => ''
    };
    const globals = {
        Asc: {editor: {asc_registerCallback: noop}},
        AscCommon: pdfCommon,
        window: {AscViewer: {}},
        performance: {now: () => 0},
        document: {}
    };
    const openDocument = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', 'openDocument', globals);
    const initRenderer = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', 'initDocumentRenderer', globals);
    let readyCalls = 0;
    const api = {
        documentLoadGeneration: 1,
        DocumentRenderer: null,
        WordControl: {m_oLogicDocument: {}, m_oDrawingDocument: {}, initEventsMobile: noop, MobileTouchManager: null, OnResize: noop},
        externalChartCollector: {init: noop},
        FontLoader: {LoadDocumentFonts: noop},
        _openDocumentEndCallback() {
            if (this.DocumentRenderer && this.DocumentRenderer.isDocumentContentReady) readyCalls++;
        },
        initDocumentRenderer: initRenderer
    };

    openDocument.call(api, {data: new Uint8Array([1])});
    assert.equal(api._pdfOpenLoadGeneration, 1);

    // Reopen advances documentLoadGeneration before the replaced file finishes opening.
    api.documentLoadGeneration = 2;
    renderer.cancelPendingLoad();
    callbacks.get('onFileOpened')();
    assert.equal(renderer.isDocumentContentReady, false);
    assert.equal(readyCalls, 0);

    // Opening the current generation lets onFileOpened complete normally.
    openDocument.call(api, {data: new Uint8Array([2])});
    assert.equal(api._pdfOpenLoadGeneration, 2);
    callbacks.get('onFileOpened')();
    assert.equal(renderer.isDocumentContentReady, true);
    assert.equal(readyCalls, 1);
});

test('reopening cancels an active Slide ThemeLoader across XHR, font, and image phases and releases LoadTheme and GlobalLock', () => {
    const ThemeLoaderSource = source('slide/Drawing/ThemeLoader.js');
    for (const phase of ['xhr', 'fonts', 'images']) {
        const harness = createBaseRefreshHarness();
        const {api, events, c_oAscAsyncActionType, c_oAscAsyncAction, getGlobalLockCount} = harness;
        const {fonts, loader: imageLoader, pendingFonts, images} = attachAssetLoaders(harness);
        const sentXhrs = [];
        let endLoadThemeCalls = 0;
        let pluginApplyCbCalls = 0;
        let docFontsEnded = 0;
        let docImagesEnded = 0;
        let abortedXhrs = 0;

        class FakeXHR {
            open() {}
            overrideMimeType() {}
            setRequestHeader() {}
            abort() { abortedXhrs++; }
            send() { sentXhrs.push(this); }
        }
        const slideActionTypes = {BlockInteraction: c_oAscAsyncActionType.BlockInteraction};
        const slideActions = {LoadTheme: 55};
        const themeWindow = {AscCommonSlide: {}};
        const themeCtx = vm.createContext({
            window: themeWindow,
            CAscThemeInfo: function(info) { Object.assign(this, info); },
            XMLHttpRequest: FakeXHR,
            location: {href: 'https://example.test'},
            Uint8Array,
            ArrayBuffer,
            Asc: {c_oAscAsyncActionType: slideActionTypes, c_oAscAsyncAction: slideActions},
            AscCommon: {
                g_defaultThemes: ['Theme1'],
                AscBrowser: {isIE: false},
                History: {MinorChanges: false},
                BinaryPPTYLoader: function() {
                    this.Load = () => {
                        loader.themes_info_editor[0] = {
                            FontMap: [{name: 'theme-font'}],
                            ImageMap: {'theme-img.png': true},
                            Master: {}
                        };
                    };
                }
            },
            AscCommonSlide: {CPresentation: {prototype: {GetWidthEMU: noop, GetHeightEMU: noop, GetWidthMM: noop, GetHeightMM: noop}}},
            editor: {WordControl: {m_oDrawingDocument: {}}}
        });
        vm.runInContext(ThemeLoaderSource, themeCtx);

        const cancelLoadTheme = loadMethod('slide/api.js', 'asc_docs_api.prototype.', 'cancelLoadTheme', {
            AscCommon: harness.common,
            c_oAscAsyncActionType: slideActionTypes,
            c_oAscAsyncAction: slideActions
        });
        api.StartLoadTheme = noop;
        api.EndLoadTheme = () => { endLoadThemeCalls++; };
        api.cancelLoadTheme = cancelLoadTheme;
        api._pluginApplyThemeCallback = () => { pluginApplyCbCalls++; };
        api.asyncFontsDocumentStartLoaded = noop;
        api.asyncFontsDocumentEndLoaded = () => { docFontsEnded++; };
        api.asyncImagesDocumentStartLoaded = noop;
        api.asyncImagesDocumentEndLoaded = () => { docImagesEnded++; };

        const loader = new themeWindow.AscCommonSlide.CThemeLoader();
        loader.Api = api;
        api.ThemeLoader = loader;

        // Simulate ChangeTheme taking GlobalLock before StartLoadTheme.
        harness.common.CollaborativeEditing.Set_GlobalLock(true);
        loader.StartLoadTheme(0);
        assert.equal(getGlobalLockCount(), 1);
        assert.equal(api.IsLongActionCurrent, 1);

        if (phase === 'fonts' || phase === 'images') {
            pendingFonts.add('theme-font');
            sentXhrs[0].readyState = 4;
            sentXhrs[0].status = 200;
            sentXhrs[0].response = new ArrayBuffer(4);
            sentXhrs[0].onload();
            assert.equal(fonts.ThemeLoader, loader);
        }
        if (phase === 'images') {
            pendingFonts.delete('theme-font');
            harness.flushTimers();
            assert.equal(fonts.ThemeLoader, null);
            assert.equal(imageLoader.ThemeLoader, loader);
            assert.equal(images.length, 1);
        }

        let deferredDuringThemeAbort = 0;
        let pluginCallbackSawCleanAbort = false;
        api._pluginApplyThemeCallback = () => {
            pluginApplyCbCalls++;
            pluginCallbackSawCleanAbort = api.IsLongActionCurrent === 0 && fonts.ThemeLoader == null && imageLoader.ThemeLoader == null;
        };
        api.LongActionCallbacks = [];
        api.LongActionCallbacks.push(() => {
            deferredDuringThemeAbort++;
            assert.equal(fonts.ThemeLoader, null, `${phase}: FontLoader.ThemeLoader must be cleared before LongActionCallbacks`);
            assert.equal(imageLoader.ThemeLoader, null, `${phase}: ImageLoader.ThemeLoader must be cleared before LongActionCallbacks`);
            assert.equal(fonts.isWorking(), false, `${phase}: FontLoader must be cancelled before LongActionCallbacks`);
        });

        // Reopen cancels the active theme load while holding LongActionCallbacks until cleanup completes.
        api._abortPendingDocumentLoad(true);
        assert.equal(getGlobalLockCount(), 0, `${phase}: GlobalLock released`);
        assert.equal(api.IsLongActionCurrent, 0, `${phase}: LoadTheme action released`);
        assert.equal(pluginApplyCbCalls, 1, `${phase}: plugin callback settled`);
        assert.equal(pluginCallbackSawCleanAbort, true, `${phase}: plugin theme callback runs after cleanup`);
        assert.equal(deferredDuringThemeAbort, 1, `${phase}: deferred callback runs after cleanup`);
        assert.equal(fonts.ThemeLoader == null, true, `${phase}: FontLoader.ThemeLoader cleared`);
        assert.equal(imageLoader.ThemeLoader == null, true, `${phase}: ImageLoader.ThemeLoader cleared`);
        assert.equal(loader.isLoading, false);
        assert.equal(loader.CurrentLoadThemeIndex, -1);
        assert.equal(loader.themes_info_editor[0], null, `${phase}: partial theme cache cleared on cancel`);
        if (phase === 'xhr') {
            assert.equal(abortedXhrs, 1);
            assert.equal(sentXhrs[0].onload, null);
        }
        if (phase === 'images') {
            // Stale theme image completion must not call EndLoadTheme, even after starting a second theme load on the reused ThemeLoader.
            const staleThemeImage = images[0];
            harness.common.CollaborativeEditing.Set_GlobalLock(true);
            loader.StartLoadTheme(0);
            pendingFonts.add('theme-font');
            sentXhrs[1].readyState = 4;
            sentXhrs[1].status = 200;
            sentXhrs[1].response = new ArrayBuffer(4);
            sentXhrs[1].onload();
            pendingFonts.delete('theme-font');
            harness.flushTimers();
            assert.equal(imageLoader.ThemeLoader, loader);

            staleThemeImage.onload();
            assert.equal(endLoadThemeCalls, 0, 'stale Theme1 image must not complete Theme2');
            assert.equal(loader.isLoading, true);
            assert.equal(loader.CurrentLoadThemeIndex, 0);

            api._abortPendingDocumentLoad(true);
            assert.equal(getGlobalLockCount(), 0);
            assert.equal(api.IsLongActionCurrent, 0);
            assert.equal(loader.themes_info_editor[0], null);
        }
        assert.equal(events.filter(e => e[0] === 'asc_onEndAction' && e[2] === slideActions.LoadTheme).length, phase === 'images' ? 2 : 1);

        // Next document opens and finishes fonts/images cleanly without being blocked by the cancelled theme.
        fonts.IsLoadDocumentFonts2 = true;
        fonts.LoadDocumentFonts([]);
        imageLoader.LoadDocumentImages({});
        assert.equal(docFontsEnded, 1, `${phase}: next document fonts complete`);
        assert.equal(docImagesEnded, 1, `${phase}: next document images complete`);
        assert.equal(endLoadThemeCalls, 0);
    }
});

test('asc_showRevision returns open status and ClosedAll during revision reopen finishes Open action', () => {
    const harness = createBaseRefreshHarness();
    const {api, events, ConnectionState, c_oCloseCode, c_oAscAsyncActionType, c_oAscAsyncAction} = harness;
    const ascShowRevision = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'asc_showRevision', {
        Asc: {c_oAscAsyncActionType, c_oAscAsyncAction},
        AscCommon: harness.common,
        c_oAscAsyncActionType,
        c_oAscAsyncAction,
        window: {
            Asc: {
                asc_CVersionHistory: function(obj) {
                    this.docId = obj.docId;
                    this.url = obj.url;
                    this.urlChanges = obj.urlChanges;
                    this.token = obj.token;
                    this.currentChangeId = obj.currentChangeId || 0;
                    this.update = next => {
                        this.docId = next.docId;
                        this.url = next.url;
                        return true;
                    };
                }
            }
        }
    });
    api.VersionHistory = null;
    api.DocInfo = {
        get_Id: () => 'doc-1',
        clone: () => ({
            put_Id: noop, put_Url: noop, put_Mode: noop, put_CoEditingMode: noop, put_Token: noop,
            get_Id: () => 'rev-2'
        })
    };

    // Blocked while document load is incomplete or docId is empty.
    api.isDocumentLoadComplete = false;
    assert.equal(ascShowRevision.call(api, {docId: 'rev-2', url: '/r2'}), false);
    api.isDocumentLoadComplete = true;
    assert.equal(ascShowRevision.call(api, {docId: '', url: '/r2'}), false);

    // Valid revision starts reopen with ServerIdWaitAction = Open and must NOT set isRefreshInProgress.
    assert.equal(ascShowRevision.call(api, {docId: 'rev-2', url: '/r2'}), true);
    assert.equal(api.isRefreshInProgress, false);
    assert.deepEqual([...api.ServerIdWaitAction], [c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open]);
    assert.equal(api.IsLongActionCurrent, 1);

    // Fatal ClosedAll during revision reopen finishes the Open action instead of leaking it.
    events.length = 0;
    api.CoAuthoringApi.get_state = () => ConnectionState.ClosedAll;
    api.CoAuthoringApi.onDisconnect('drop', c_oCloseCode.drop);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(events.filter(e => e[0] === 'asc_onEndAction' && e[2] === c_oAscAsyncAction.Open).length, 1);
});

test('onRefreshFile guards ClosedAll after sessionIdle and clears isRefreshInProgress on WOPI failure', () => {
    const {api, events, ConnectionState, c_oCloseCode} = createBaseRefreshHarness();
    let readyCalls = 0;
    api.isDocumentLoadComplete = false;
    api.ServerIdWaitComplete = false;
    api.ServerImagesWaitComplete = true;
    api._openDocumentEndCallback = () => { readyCalls++; };

    api.onRefreshFile();
    assert.equal(api.isRefreshInProgress, true);
    api.CoAuthoringApi.onDisconnect('idle', c_oCloseCode.sessionIdle);
    assert.equal(api.ServerIdWaitAction, null);
    assert.equal(api.isRefreshInProgress, true);

    events.length = 0;
    api.CoAuthoringApi.get_state = () => ConnectionState.ClosedAll;
    api.CoAuthoringApi.onDisconnect('drop', c_oCloseCode.drop);
    assert.equal(readyCalls, 0);
    assert.equal(api.ServerIdWaitComplete, false);
    assert.equal(api.isRefreshInProgress, false);

    // WOPI refresh error clears isRefreshInProgress.
    const wopi = createBaseRefreshHarness();
    wopi.api.documentIsWopi = true;
    wopi.api.CoAuthoringApi.callPRC = (_req, _t, cb) => { cb(true, undefined); return true; };
    wopi.api.onRefreshFile();
    assert.equal(wopi.api.isRefreshInProgress, false);
    assert.equal(wopi.api.ServerIdWaitAction, null);
});

test('aborting document load cancels pending plugin pastes, pre_Paste onCancel, and holds LongActionCallbacks until cleanup completes', () => {
    const harness = createBaseRefreshHarness();
    const {api} = harness;
    const pluginReturns = [];
    const removedElems = [];
    const createdElems = new Map();
    let groupStarts = 0;
    let groupEnds = 0;
    const gClipboard = {bSaveFormat: true};
    const fakeBody = {
        appendChild(el) { el.parentNode = this; createdElems.set(el.id, el); },
        removeChild(el) { el.parentNode = null; createdElems.delete(el.id); removedElems.push(el.id); },
        contains(el) { return createdElems.get(el.id) === el; }
    };
    const pluginGlobals = {
        window: {
            g_asc_plugins: {
                setPluginMethodReturnAsync: noop,
                onPluginMethodReturn: ok => pluginReturns.push(ok)
            },
            navigator: {userAgent: ''}
        },
        document: {
            body: fakeBody,
            getElementById: id => createdElems.get(id) || null,
            createElement: () => ({style: {}, innerHTML: '', id: ''})
        },
        AscCommon: {
            g_clipboardBase: gClipboard,
            c_oAscClipboardDataFormat: {HtmlElement: 1, Text: 2},
            c_oEditorId: {Word: 1}
        }
    };
    api.pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteHtml"]', '', pluginGlobals);
    api._pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype.', '_pluginMethod_PasteHtml', pluginGlobals);
    api.pluginMethod_PasteText = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteText"]', '', pluginGlobals);
    api._addPendingPluginPasteCancel = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_addPendingPluginPasteCancel', {});
    api._removePendingPluginPasteCancel = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_removePendingPluginPasteCancel', {});
    api._cancelPendingPluginPastes = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_cancelPendingPluginPastes', {});
    api.executeGroupActionsStart = () => { groupStarts++; };
    api.executeGroupActionsEnd = () => { groupEnds++; };
    api.executeGroupActions = fn => fn();
    api.canEdit = () => true;
    api.get_TextProps = () => null;
    api.getEditorId = () => 1;
    api.WordControl = {m_oLogicDocument: {GetParagraphs: () => []}};

    let deferredRuns = 0;
    let deferredSawCleanState = false;
    api.LongActionCallbacks = [];
    const origDec = api.decrementCounterLongAction.bind(api);
    api.decrementCounterLongAction = function() {
        origDec();
        if (this.IsLongActionCurrent === 0 && this.LongActionCallbacks.length) {
            const cbs = this.LongActionCallbacks.splice(0);
            for (const cb of cbs) cb();
        }
    };
    api.checkLongActionCallback = function(cb) {
        if (this.IsLongActionCurrent > 0) this.LongActionCallbacks.push(cb);
        else cb();
    };
    let pasteCancelRuns = 0;
    api.asc_PasteData = (_fmt, _el, _a, _b, _c, cb) => {
        api.incrementCounterLongAction();
        api.isPasteFonts_Images = true;
        api.pasteCallback = cb;
        api.pasteCancelCallback = () => { pasteCancelRuns++; };
    };

    // Start PasteHtml and queue a LongActionCallback that must wait until cleanup finishes.
    api.pluginMethod_PasteHtml('<p>hello</p>');
    assert.equal(api.IsLongActionCurrent, 2);
    assert.equal(gClipboard.bSaveFormat, false);
    assert.ok(createdElems.has('pmpastehtml'));
    api.checkLongActionCallback(() => {
        deferredRuns++;
        deferredSawCleanState = !createdElems.has('pmpastehtml') && api.isPasteFonts_Images === false && api.pasteCancelCallback === null;
    });

    api._abortPendingDocumentLoad(true);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(pasteCancelRuns, 1);
    assert.equal(gClipboard.bSaveFormat, true);
    assert.equal(groupStarts, 1);
    assert.equal(groupEnds, 1);
    assert.equal(createdElems.has('pmpastehtml'), false);
    assert.deepEqual(pluginReturns, [false]);
    assert.equal(deferredRuns, 1);
    assert.equal(deferredSawCleanState, true, 'LongActionCallbacks must run after _abortPendingDocumentLoad finishes cleanup');

    // Subsequent PasteHtml is not blocked by a leftover #pmpastehtml element.
    api.pluginMethod_PasteHtml('<p>second</p>');
    assert.ok(createdElems.has('pmpastehtml'));
    api._abortPendingDocumentLoad(true);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(createdElems.has('pmpastehtml'), false);

    // PasteText also releases its long action and returns false on abort.
    api.asc_PasteData = noop;
    api.pluginMethod_PasteText('plain');
    assert.equal(api.IsLongActionCurrent, 1);
    api._abortPendingDocumentLoad(true);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(pluginReturns.at(-1), false);
});

test('aborting document load cancels ReplaceTextSmart font wait and releases plugin long actions', () => {
    const harness = createBaseRefreshHarness();
    const {api} = harness;
    const {fonts, pendingFonts} = attachAssetLoaders(harness);
    const pluginReturns = [];
    let replaceCalls = 0;
    let deferredRuns = 0;
    let deferredSawClean = false;
    const callbackObj = {_this: null, _callback: null};
    const fakeAscFonts = {
        FontPickerByCharacter: {
            CallbackObj: callbackObj,
            checkText(_text, ctx, cb) {
                callbackObj._this = ctx;
                callbackObj._callback = cb;
                api.asyncMethodCallback = function() {
                    callbackObj._callback.call(callbackObj._this);
                    callbackObj._this = null;
                    callbackObj._callback = null;
                };
                fonts.LoadDocumentFonts2([{name: 'cjk-font'}]);
                return true;
            }
        }
    };
    api.pluginMethod_ReplaceTextSmart = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_ReplaceTextSmart"]', '', {
        window: {
            g_asc_plugins: {
                setPluginMethodReturnAsync: noop,
                onPluginMethodReturn: ok => pluginReturns.push(ok)
            }
        },
        AscFonts: fakeAscFonts,
        AscCommon: {c_oEditorId: {Word: 1, Presentation: 2, Spreadsheet: 3}}
    });
    api._addPendingPluginPasteCancel = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_addPendingPluginPasteCancel', {});
    api._removePendingPluginPasteCancel = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_removePendingPluginPasteCancel', {});
    api._cancelPendingPluginPastes = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', '_cancelPendingPluginPastes', {});
    api.editorId = 1;
    api.getJsApi = () => ({ReplaceTextSmart: () => { replaceCalls++; }});
    api.asc_canPaste = noop;
    api.asc_Recalculate = noop;
    api.WordControl = {m_oLogicDocument: {FinalizeAction: noop}};

    pendingFonts.add('cjk-font');
    api.pluginMethod_ReplaceTextSmart(['hello']);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(typeof callbackObj._callback, 'function');
    const capturedCallback = callbackObj._callback;

    api.LongActionCallbacks = [() => {
        deferredRuns++;
        deferredSawClean = api.IsLongActionCurrent === 0 && api.asyncMethodCallback === undefined && callbackObj._callback === null;
    }];

    api._abortPendingDocumentLoad(true);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(api.asyncMethodCallback, undefined);
    assert.equal(callbackObj._callback, null);
    assert.equal(replaceCalls, 0);
    assert.deepEqual(pluginReturns, [false]);
    assert.equal(deferredRuns, 1);
    assert.equal(deferredSawClean, true);

    // Late invocation of the cancelled ReplaceTextSmart callback is a no-op.
    capturedCallback.call(api);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(replaceCalls, 0);
    assert.deepEqual(pluginReturns, [false]);
});

test('Cell stops opening tail on synchronous reopen and cleans partial WorkbookView on aborted refresh', () => {
    const harness = createBaseRefreshHarness();
    const {api} = harness;
    let sendChangesCalls = 0;
    let applyPreOpenLocksCalls = 0;
    let goToCalls = 0;
    let resizeCalls = 0;
    let removedHandlers = 0;
    let destroyedTouch = 0;
    let closedDrawing = 0;
    let clearedStyles = 0;
    const fakeExcel = {
        WorkbookView: function() {
            this.removeHandlersList = () => { removedHandlers++; };
            this.MobileTouchManager = {Destroy: () => { destroyedTouch++; }};
        },
        g_StyleCache: {Clear: () => { clearedStyles++; }},
        oFormulaLocaleInfo: {Parse: true, DigitSep: true}
    };
    const cellGlobals = {
        Asc: {editor: {asc_unregisterCallback: noop, asc_registerCallback: noop}},
        AscCommon: {
            ...harness.common,
            EncryptionWorker: null,
            baseEditorsApi: {prototype: {_abortPendingDocumentLoad: api._abortPendingDocumentLoad}}
        },
        AscCommonExcel: fakeExcel,
        History: {TurnOn: noop},
        c_oAscAdvancedOptionsAction: {None: 0},
        c_oAscError: harness.c_oAscError
    };
    api._applyFirstLoadChanges = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_applyFirstLoadChanges', cellGlobals);
    api._openDocumentEndCallback = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_openDocumentEndCallback', cellGlobals);
    api._abortPendingDocumentLoad = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_abortPendingDocumentLoad', cellGlobals);

    api.isDocumentLoadComplete = false;
    api.ServerIdWaitComplete = true;
    api.FontLoadWaitComplete = true;
    api.IsInitControl = true;
    api.wbModel = {
        openErrors: [],
        theme: {},
        getWorksheetCount: () => 1,
        DrawingDocument: {CloseFile: () => { closedDrawing++; }}
    };
    api.collaborativeEditing = {
        applyChanges: () => true,
        sendChanges: () => { sendChangesCalls++; }
    };
    api.registerCustomFunctionsLibrary = noop;
    api.isUseOldMobileVersion = () => false;
    api.asc_CheckGuiControlColors = noop;
    api.sendColorThemes = noop;
    api.asc_ApplyColorScheme = noop;
    api.updateDarkMode = noop;
    api.sendStandartTextures = noop;
    api.sendMathToMenu = noop;
    api._applyPreOpenLocks = () => { applyPreOpenLocksCalls++; api.isPreOpenLocks = false; };
    api.goTo = () => { goToCalls++; };
    api.asc_Resize = () => { resizeCalls++; };
    api.onNeedUpdateExternalReferenceOnOpen = noop;
    api.onUpdateDocumentModified = noop;
    api.initBroadcastChannelListeners = noop;
    api.frameManager = {isFrameEditor: () => false};

    // Synchronous reopen during onDocumentContentReady must halt the old Cell opening tail and clean the pending WorkbookView once.
    api.onDocumentContentReady = () => {
        api.reopenFileWithReconnection({get_Id: () => 'cell-k2'});
    };
    api._openDocumentEndCallback();
    assert.equal(sendChangesCalls, 0);
    assert.equal(applyPreOpenLocksCalls, 0);
    assert.equal(goToCalls, 0);
    assert.equal(resizeCalls, 0);
    assert.equal(api.isPreOpenLocks, true);
    assert.equal(api._hasPendingWorkbookView, false);
    assert.ok(api.wb, 'wb stays non-null for zoom/history queries during reopen');
    assert.equal(removedHandlers, 1);
    assert.equal(destroyedTouch, 1);
    assert.equal(closedDrawing, 1);
    assert.equal(clearedStyles, 1);

    // A subsequent reopen after cleanup (or after asc_CloseFile) does not destroy wb twice.
    api._abortPendingDocumentLoad(true);
    assert.equal(removedHandlers, 1);
    assert.equal(destroyedTouch, 1);
    assert.equal(closedDrawing, 1);
    assert.equal(clearedStyles, 1);
});

test('PDF cancelPendingLoad closes pending encrypted file and asc_setAdvancedOptions ignores stale password', () => {
    let closedFiles = 0;
    let openCalls = 0;
    let overlayCleared = 0;
    let formsCleared = 0;
    let paintCalled = 0;
    const viewer = {
        openSessionId: 0,
        checkReadyTimerId: null,
        isDocumentContentReady: true,
        isDocumentReady: true,
        _fileData: new Uint8Array([1, 2]),
        isCMapLoading: true,
        isFullText: true,
        file: {isValid: () => false, close: () => { closedFiles++; }},
        SearchResults: {matches: []},
        structure: {items: []},
        canvas: {width: 400, height: 600},
        canvasForms: {width: 400, height: 600, getContext: () => ({clearRect: () => { formsCleared++; }})},
        overlay: {Clear: () => { overlayCleared++; }},
        getPDFDoc: () => ({BlurActiveObject: noop}),
        onUpdateOverlay: noop,
        _paint: () => { paintCalled++; }
    };
    const viewerGlobals = {
        clearTimeout: noop,
        CDocumentPagesInfo: function() {}
    };
    viewer._cancelPendingPaintCallbacks = loadMethod('pdf/src/viewer.js', 'this.', '_cancelPendingPaintCallbacks', viewerGlobals);
    viewer.cancelPendingLoad = loadMethod('pdf/src/viewer.js', 'this.', 'cancelPendingLoad', viewerGlobals);
    viewer.close = loadMethod('pdf/src/viewer.js', 'this.', 'close', viewerGlobals);

    // Aborting an incomplete encrypted PDF (not yet valid) closes and clears viewer.file.
    viewer.cancelPendingLoad();
    assert.equal(closedFiles, 1);
    assert.equal(viewer.file, null);
    assert.equal(viewer.SearchResults, null);

    // Cancelling with an already-open valid file runs full close() cleanup and clears viewer.file without re-arming _paint().
    viewer.file = {isValid: () => true, close: () => { closedFiles++; }};
    viewer.structure = {items: [1]};
    viewer.isFullText = true;
    viewer.cancelPendingLoad();
    assert.equal(closedFiles, 2);
    assert.equal(viewer.file, null);
    assert.equal(viewer.structure, null);
    assert.equal(viewer.isFullText, false);
    assert.equal(paintCalled, 0);
    assert.equal(overlayCleared, 1);
    assert.equal(formsCleared, 1);

    // Direct close() still runs full cleanup via cancelPendingLoad(true).
    viewer.file = {isValid: () => true, close: () => { closedFiles++; }};
    viewer.structure = {items: [2]};
    viewer.close();
    assert.equal(closedFiles, 3);
    assert.equal(viewer.structure, null);
    assert.equal(paintCalled, 0);
    assert.equal(overlayCleared, 2);
    assert.equal(formsCleared, 2);

    // Stale password submission after refresh is ignored by PDFEditorApi.asc_setAdvancedOptions.
    const setAdvancedOptions = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', 'asc_setAdvancedOptions', {
        AscCommon: {
            c_oAscAdvancedOptionsAction: {Open: 1},
            EncryptionWorker: {asc_setAdvancedOptions: () => false}
        }
    });
    const pdfApi = {
        advancedOptionsAction: 1,
        documentLoadGeneration: 2,
        _pdfOpenLoadGeneration: 1,
        DocumentRenderer: {file: null, open: () => { openCalls++; }}
    };
    setAdvancedOptions.call(pdfApi, 0, {asc_getPassword: () => 'secret'});
    assert.equal(openCalls, 0);

    // Even if a file existed, mismatched generation still blocks stale password submission.
    pdfApi.DocumentRenderer.file = {isValid: () => true};
    setAdvancedOptions.call(pdfApi, 0, {asc_getPassword: () => 'secret'});
    assert.equal(openCalls, 0);

    // Matching generation and active file allows password submission for the current document.
    pdfApi._pdfOpenLoadGeneration = 2;
    setAdvancedOptions.call(pdfApi, 0, {asc_getPassword: () => 'secret'});
    assert.equal(openCalls, 1);
});

test('infinity-loop refresh, incomplete onRefreshFile, and online Visio OOXML do not emit stale readiness', () => {
    // 1. Infinity-loop branch disconnects the old CoAuthoringApi, clears readiness flags, and ignores subsequent load events.
    {
        const {api, openRequests, events, coApis} = createOpenCommandHarness();
        const co1 = coApis[0];
        api.isDocumentLoadComplete = false;
        api.ServerImagesWaitComplete = true;
        api.FontLoadWaitComplete = true;
        api.asc_refreshFile({get_Id: () => 'doc-1'});
        assert.equal(co1.disconnectCalls.length, 1);
        assert.equal(api.ServerImagesWaitComplete, false);
        assert.equal(api.FontLoadWaitComplete, false);
        events.length = 0;

        co1.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 1, data: {'Editor.bin': 'https://files/doc-1.bin'}}});
        co1.onFirstLoadChangesEnd(1);
        assert.equal(openRequests.length, 0);
        assert.deepEqual(events, []);
    }

    // 2. onRefreshFile during incomplete load disconnects non-WOPI socket and ignores stale onFirstLoadChangesEnd / onDocumentOpen.
    {
        const {api, openRequests, events, coApis} = createOpenCommandHarness();
        const co1 = coApis[0];
        api.isDocumentLoadComplete = false;
        api.ServerImagesWaitComplete = true;
        api.onRefreshFile();
        assert.equal(co1.disconnectCalls.length, 1);
        assert.equal(api.ServerImagesWaitComplete, false);
        assert.equal(api.isRefreshInProgress, true);
        events.length = 0;

        co1.onFirstLoadChangesEnd(1);
        co1.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 1, data: {'Editor.bin': 'https://files/doc-1.bin'}}});
        assert.equal(openRequests.length, 0);
        assert.equal(api.isRefreshInProgress, true);
        assert.deepEqual(events, []);

        // Subsequent refreshFile with a new key re-enables load events on the new connection.
        api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
        api.CoAuthoringApi.onDocumentOpen({data: {type: 'open', status: 'ok', openedAt: 2, data: {'Editor.bin': 'https://files/doc-2.bin'}}});
        assert.equal(openRequests.length, 1);
        assert.equal(openRequests[0].docId, 'doc-2');
    }

    // 3. Online Visio OOXML does not set ServerIdWaitComplete before asyncServerIdEndLoaded.
    {
        const openZip = loadMethod('visio/api.js', 'VisioEditorApi.prototype.', 'OpenDocumentFromZipNoInit', {
            AscCommon: {
                XmlParserContext: function() { this.loadDataLinks = () => ({}); },
                ZLib: function() { this.open = () => true; this.close = noop; },
                pptx_content_loader: {Reader: {}, ImageMapChecker: {}}
            }
        });
        const onlineVisio = {
            InitEditor: noop,
            Document: {fromZip: noop},
            ServerIdWaitComplete: false,
            CoAuthoringApi: {get_onlineWork: () => true}
        };
        assert.equal(openZip.call(onlineVisio, new Uint8Array([1])), true);
        assert.equal(onlineVisio.ServerIdWaitComplete, false);

        const offlineVisio = {
            InitEditor: noop,
            Document: {fromZip: noop},
            ServerIdWaitComplete: false,
            CoAuthoringApi: {get_onlineWork: () => false}
        };
        assert.equal(openZip.call(offlineVisio, new Uint8Array([1])), true);
        assert.equal(offlineVisio.ServerIdWaitComplete, true);
    }
});


function loadRealPluginsManager(api, overrides = {}) {
    const sentMessages = [];
    const sandbox = {
        window: {},
        document: {getElementById: () => null},
        console,
        Asc: {
            checkReturnCommand: () => true,
            ...(overrides.Asc || {})
        },
        AscCommon: {
            isValidJs: () => true,
            safePluginEval: overrides.safePluginEval || (() => undefined),
            History: {Is_LastPointEmpty: () => true},
            ...(overrides.AscCommon || {})
        }
    };
    sandbox.window = sandbox;
    const text = source('common/plugins.js');
    const cPluginDataStart = text.indexOf('function CPluginData()');
    const exportIdx = text.indexOf('// export', cPluginDataStart);
    assert.ok(cPluginDataStart >= 0 && exportIdx > cPluginDataStart);
    const snippet = text.slice(cPluginDataStart, exportIdx) + '\n;({CPluginsManager, CPluginData, CommandTaskType});';
    const {CPluginsManager} = vm.runInNewContext(snippet, sandbox);
    const mgr = new CPluginsManager(api);
    mgr.internalGuid = 'internal-guid';
    mgr.internalCallbacks = [];
    mgr.internalCommandAsync = false;
    mgr.sendMessageToFrame = (frameId, pluginData) => {
        sentMessages.push({
            frameId,
            guid: pluginData.privateData.guid,
            type: pluginData.privateData.type,
            methodReturnData: pluginData.privateData.methodReturnData,
            commandReturnData: pluginData.privateData.commandReturnData
        });
    };
    return {mgr, sentMessages, sandbox};
}

test('PDF thumbnails clear cleanly on refreshFile close and Chrome rAF loop keeps running through reopen', () => {
    const makeCanvas = () => ({
        width: 200,
        height: 400,
        style: {},
        getBoundingClientRect: () => ({left: 0, top: 0, width: 200, height: 400}),
        getContext: () => ({
            fillStyle: '',
            strokeStyle: '',
            lineWidth: 1,
            font: '',
            textAlign: '',
            fillRect: noop,
            strokeRect: noop,
            fillText: noop,
            drawImage: noop,
            save: noop,
            restore: noop,
            translate: noop,
            rotate: noop,
            getImageData: () => ({data: new Uint8Array(100 * 100 * 4)})
        })
    });

    const elements = new Map();
    const panelEl = {id: 'thumbnails-list', offsetWidth: 220, offsetHeight: 500, style: {}};
    const scrollEl = {id: 'id_vertical_scroll_th', style: {}};
    const viewerCanvas = makeCanvas();
    const overlayCanvas = makeCanvas();
    const dragCanvas = makeCanvas();
    elements.set('thumbnails-list', panelEl);
    elements.set('id_vertical_scroll_th', scrollEl);
    elements.set('id_viewer_th', viewerCanvas);
    elements.set('id_overlay_th', overlayCanvas);
    elements.set('id_drag_canvas', dragCanvas);

    const mockDoc = {
        getElementById: id => elements.get(id) || null,
        createElement: () => makeCanvas()
    };

    const pdfDoc = {
        fontLoader: {isWorking: () => false},
        SearchEngine: {Clear: () => { searchCleared++; }, Count: 0},
        pagesTransform: [{a: 1}],
        BlurActiveObject: () => { blurred++; },
        GetThumbnails: () => thumbnails,
        UpdateInterface: noop,
        UpdateCopyCutState: noop,
        RecalculateAll: noop
    };
    let searchCleared = 0;
    let blurred = 0;
    let endedWaiting = 0;
    let getPageCalls = 0;
    let contextMenus = 0;
    const rafCallbacks = [];
    const mockWindow = {
        NATIVE_EDITOR_ENJINE: true,
        requestAnimationFrame: cb => { rafCallbacks.push(cb); return rafCallbacks.length; }
    };
    const localAsc = {
        c_oAscAsyncActionType: {BlockInteraction: 1},
        c_oAscAsyncAction: {Waiting: 99},
        c_oAscPdfContextMenuTypes: {Thumbnails: 1},
        editor: null
    };
    const scheduledTimers = new Map();
    let nextTimerId = 1;
    const mockSetTimeout = fn => {
        const id = nextTimerId++;
        scheduledTimers.set(id, fn);
        return id;
    };
    const mockClearTimeout = id => {
        scheduledTimers.delete(id);
    };
    const realScrollRepos = loadMethod('common/scroll.js', 'ScrollObject.prototype.', 'Repos', {
        AscBrowser: {retinaPixelRatio: 1},
        GetClientWidth: el => (el && el.clientWidth) || 0,
        GetClientHeight: el => (el && el.clientHeight) || 0,
        _HEXTORGB_: () => ({R: 0, G: 0, B: 0})
    });
    const localAscCommon = {
        AscBrowser: {
            isChrome: true,
            isSafariMacOs: false,
            retinaPixelRatio: 1,
            zoom: 1,
            convertToRetinaValue: v => v
        },
        GlobalSkin: {BackgroundColor: '#fff', PageOutline: '#ccc'},
        CollaborativeEditing: {waitingImagesForLoad: false},
        global_mouseEvent: {Button: 2, X: 10, Y: 10, IsLocked: false, LockMouse: noop},
        check_MouseUpEvent: noop,
        addMouseEvent: noop,
        calculateCanvasSize: noop,
        ScrollSettings: function() {},
        ScrollObject: function(id, settings) {
            this.canvas = {style: {display: ''}, parentNode: {clientWidth: 16, clientHeight: settings.screenH || 500}};
            this.settings = {...settings};
            this.scroller = {w: 8, h: 40, x: 0, y: 0};
            this.scrollVCurrentY = 0;
            this.scrollHCurrentX = 0;
            this.canvasH = settings.screenH || 500;
            this.canvasW = 16;
            this.maxScrollY = Math.max(0, (settings.contentH || 0) - (settings.screenH || 0));
            this.maxScrollY2 = this.maxScrollY;
            this.maxScrollX = 0;
            this.maxScrollX2 = 0;
            this.scrollToYCalls = [];
            this._setDimension = (h, w) => { this.canvasH = h; this.canvasW = w; };
            this._setScrollerHW = noop;
            this._roundForScale = v => v;
            this.RecalcScroller = noop;
            this._initPiperImg = noop;
            this._draw = noop;
            this._drawArrows = noop;
            this.Repos = realScrollRepos;
            this.bind = noop;
            this.getMaxScrolledY = () => this.maxScrollY;
            this.scrollToY = y => {
                this.scrollVCurrentY = y;
                this.scrollToYCalls.push(y);
            };
        }
    };

    const thumbSource = source('pdf/src/thumbnails.js');
    const thumbSandbox = {
        window: mockWindow,
        document: mockDoc,
        Asc: localAsc,
        AscCommon: localAscCommon,
        GlobalSkin: localAscCommon.GlobalSkin,
        global_mouseEvent: localAscCommon.global_mouseEvent,
        AscPDF: {}
    };
    mockWindow.Asc = localAsc;
    mockWindow.AscCommon = localAscCommon;
    mockWindow.AscPDF = thumbSandbox.AscPDF;
    vm.runInNewContext(thumbSource, thumbSandbox);
    const CThumbnailsDoc = thumbSandbox.AscCommon.ThumbnailsControl;
    assert.ok(CThumbnailsDoc, 'thumbnails constructor should be exported');

    const makeFile = pageCount => {
        let valid = true;
        const pages = Array.from({length: pageCount}, (_, idx) => ({
            W: 600,
            H: 800,
            Dpi: 96,
            isRecognized: false,
            originIndex: idx,
            text: null
        }));
        return {
            type: 1,
            pages,
            isValid: () => valid && pages.length > 0,
            close: () => { valid = false; pages.length = 0; },
            getPage: () => { getPageCalls++; return makeCanvas(); },
            copy: () => true,
            selectAll: noop,
            removeSelection: noop
        };
    };

    let fontLoadInProgress = false;
    let fontWorking = false;
    let overlayClears = 0;
    let formsClears = 0;
    const pdfDocFull = {
        ...pdfDoc,
        GetDrawingDocument: () => ({m_oDocumentRenderer: null, GetDrawingObjects: () => []}),
        GetSelection: () => null,
        fontLoader: {
            isWorking: () => fontWorking,
            isFontLoadInProgress: () => fontLoadInProgress
        }
    };
    const viewerGlobals = {
        window: mockWindow,
        document: mockDoc,
        Asc: localAsc,
        AscCommon: localAscCommon,
        AscPDF: {PRINT_CONTENT_TYPES: {formsOnly: 1}},
        Date: {now: () => 10000},
        clearTimeout: mockClearTimeout,
        clearInterval: noop,
        setTimeout: mockSetTimeout,
        CDocumentPagesInfo: function() { this.pages = []; this.countTextPages = 0; }
    };
    const mainCanvas = makeCanvas();
    const formsCanvas = {
        width: 200,
        height: 400,
        getContext: () => ({clearRect: () => { formsClears++; }})
    };
    const viewer = {
        Api: {
            isDarkMode: false,
            getPageBackgroundColor: () => ({R: 255, G: 255, B: 255}),
            sync_EndAction: (type, id) => { if (type === 1 && id === 99) endedWaiting++; },
            sync_ContextMenuCallback: () => { contextMenus++; },
            WordControl: {NoneRepaintPages: false}
        },
        openSessionId: 0,
        checkReadyTimerId: null,
        _timerId: null,
        isDocumentContentReady: true,
        isDocumentReady: true,
        _fileData: null,
        isCMapLoading: false,
        IsOpenFormsInProgress: true,
        UseRequestAnimationFrame: true,
        RequestAnimationFrame: mockWindow.requestAnimationFrame,
        RequestAnimationOldTime: -1,
        isRepaint: false,
        startVisiblePage: 0,
        endVisiblePage: 0,
        documentWidth: 600,
        documentHeight: 800,
        scrollX: 10,
        scrollY: 20,
        scrollMaxX: 100,
        scrollMaxY: 200,
        pageDetector: {pages: []},
        initPaintDone: true,
        isMouseDown: true,
        timerScrollSelect: -1,
        scheduledRepaintTimer: null,
        onRepaintCallbacks: [noop],
        onAfterPaintCallback: [noop],
        IsSearch: true,
        CurrentSearchNavi: [{X: 1, Y: 2}],
        isFullTextMessage: true,
        SearchResults: {IsSearch: true, Show: true},
        structure: [{page: 0, y: 0}],
        pagesInfo: {pages: [{IsLocked: () => false, isPainted: true, links: null}], countTextPages: 0},
        drawingPages: [{W: 600, H: 800, Y: 0}],
        drawingDocument: {m_arrPages: [1], m_lPagesCount: 1, m_lDrawingFirst: 0, m_lDrawingEnd: 0},
        file: makeFile(8),
        canvas: mainCanvas,
        canvasForms: formsCanvas,
        overlay: {Clear: () => { overlayClears++; }},
        x: 0,
        y: 0,
        getPDFDoc: () => pdfDocFull,
        isLandscapePage: () => false,
        getPageRotate: () => 0,
        _checkFontsOnPages: () => true,
        navigateToPage: noop,
        getPageByCoords2: () => null,
        showTextMessage: () => { throw new Error('must not start Waiting mask on closed file'); }
    };
    localAsc.editor = {
        getDocumentRenderer: () => viewer,
        getPDFDoc: () => pdfDocFull,
        canEdit: () => true,
        checkLastWork: noop,
        NeedShowTextSelectPanel: () => true
    };
    for (const name of ['_cancelPendingPaintCallbacks', 'cancelPendingLoad', 'close', 'scheduleRepaint', 'paint', '_paint', 'afterPaintCallbacks', 'onUpdateOverlay', 'navigate', 'Copy', 'selectAll', 'removeSelection', 'findText', 'checkPagesLinks', 'checkPagesText', 'getTextCommandsSize']) {
        viewer[name] = loadMethod('pdf/src/viewer.js', 'this.', name, viewerGlobals);
    }
    viewer.GetPageForThumbnails = loadMethod('pdf/src/viewer.js', 'CHtmlPage.prototype.', 'GetPageForThumbnails', viewerGlobals);
    viewer.GetPrintPage = loadMethod('pdf/src/viewer.js', 'CHtmlPage.prototype.', 'GetPrintPage', viewerGlobals);
    viewer._runTimerTick = loadMethod('pdf/src/viewer.js', 'this.', '_runTimerTick', {...viewerGlobals, oThis: viewer});
    viewer.timer = loadMethod('pdf/src/viewer.js', 'this.', 'timer', {...viewerGlobals, oThis: viewer});
    viewer.timerSync = loadMethod('pdf/src/viewer.js', 'this.', 'timerSync', {...viewerGlobals, oThis: viewer});
    viewer.timerAnimation = loadMethod('pdf/src/viewer.js', 'this.', 'timerAnimation', {...viewerGlobals, oThis: viewer});

    const thumbnails = new CThumbnailsDoc('thumbnails-list');
    thumbnails.viewer = viewer;
    viewer.thumbnails = thumbnails;
    thumbnails.init();
    assert.equal(thumbnails.pages.length, 8);
    assert.ok(thumbnails.blocks.length > 0);
    assert.equal(thumbnails.startBlock, 0);
    assert.ok(thumbnails.m_oScrollVerApi && thumbnails.m_oScrollVerApi.maxScrollY > 0);

    // Simulate user scrolling thumbnails down before refreshFile.
    thumbnails.scrollY = 300;
    thumbnails.m_oScrollVerApi.scrollVCurrentY = 300;

    // Hide panel (as LeftMenu.SetDisabled does on refresh-file) and close via cancelPendingLoad()
    // while fontLoader is busy: _paint() must not re-schedule repaint and overlay.Clear() must still run.
    fontLoadInProgress = true;
    fontWorking = true;
    panelEl.offsetWidth = 0;
    viewer.cancelPendingLoad();

    assert.equal(viewer.file, null);
    assert.equal(viewer.pageDetector, null);
    assert.equal(viewer.initPaintDone, false);
    assert.equal(viewer.IsOpenFormsInProgress, false);
    assert.equal(viewer.isRepaint, false);
    assert.equal(viewer.scheduledRepaintTimer, null);
    assert.equal(scheduledTimers.size, 0, 'close() must not re-schedule repaint timers on an invalid file');
    assert.equal(overlayClears, 1, 'close() clears overlay even when fontLoader.isWorking() is true');
    assert.equal(formsClears, 1, 'close() clears canvasForms');
    assert.equal(endedWaiting, 1);
    assert.equal(searchCleared, 1);
    assert.equal(blurred, 1);
    assert.equal(pdfDocFull.pagesTransform.length, 0);
    assert.equal(viewer.drawingDocument.m_lPagesCount, 0);
    assert.equal(thumbnails.pages.length, 0);
    assert.equal(thumbnails.blocks.length, 0);
    assert.equal(thumbnails.startBlock, -1);
    assert.equal(thumbnails.endBlock, -1);
    assert.equal(thumbnails.m_oScrollVerApi.scrollVCurrentY, 0);
    assert.equal(thumbnails.m_oScrollVerApi.canvasH, 0);

    // Calling real _paint() while file is closed and fontLoader is busy does not schedule repaint.
    assert.equal(viewer._paint(), false);
    assert.equal(scheduledTimers.size, 0);
    fontLoadInProgress = false;
    fontWorking = false;

    // User re-opens thumbnails panel during refresh gap while file === null.
    panelEl.offsetWidth = 220;
    thumbnails.resize();
    thumbnails.setZoom(0.5);
    thumbnails.setIsDrawCurrentRect(true);
    thumbnails.onMouseUp({preventDefault: noop});
    assert.equal(contextMenus, 0);

    // Run Chrome rAF ticks while file is closed: no throw, no GetPageForThumbnails, rAF keeps rescheduling.
    viewer.timerAnimation();
    assert.equal(rafCallbacks.length, 1);
    viewer.RequestAnimationOldTime = -1;
    rafCallbacks.shift()();
    assert.equal(rafCallbacks.length, 1);
    assert.equal(scheduledTimers.size, 0);
    assert.equal(getPageCalls, 0);

    // Non-Chrome (Firefox/Safari) timer survives an error inside _runTimerTick and timerSync does not multiply timers.
    viewer.UseRequestAnimationFrame = false;
    const origRunTick = viewer._runTimerTick;
    viewer._runTimerTick = () => { throw new Error('tick-boom'); };
    assert.throws(() => viewer.timer(), /tick-boom/);
    assert.equal(scheduledTimers.size, 1, 'non-Chrome timer reschedules in finally after a throw');
    viewer._runTimerTick = origRunTick;
    viewer.timerSync();
    viewer.timerSync();
    assert.equal(scheduledTimers.size, 1, 'timerSync does not spawn extra 40ms timers');
    mockClearTimeout(viewer._timerId);
    viewer._timerId = null;
    viewer.UseRequestAnimationFrame = true;

    // Reopening a PDF with the same page count resets scrollbar position to 0 through real ScrollObject.Repos.
    viewer.file = makeFile(8);
    thumbnails.init();
    assert.equal(thumbnails.scrollY, 0);
    assert.equal(thumbnails.m_oScrollVerApi.scrollVCurrentY, 0);
    assert.deepEqual(thumbnails.m_oScrollVerApi.scrollToYCalls.at(-1), 0);

    // Reopening a 1-page PDF (maxScrollY === 0) hides the vertical scrollbar cleanly.
    viewer.file = makeFile(1);
    thumbnails.init();
    assert.equal(thumbnails.scrollY, 0);
    assert.equal(thumbnails.m_oScrollVerApi.maxScrollY, 0);
    assert.equal(thumbnails.m_oScrollVerApi.canvas.style.display, 'none');

    // Hidden panel on init defers resize until panel opens.
    panelEl.offsetWidth = 0;
    viewer.file = makeFile(8);
    thumbnails.init();
    assert.equal(thumbnails.needResize, true);
    panelEl.offsetWidth = 220;
    thumbnails.checkTasks(false);
    assert.equal(thumbnails.needResize, false);
    assert.equal(thumbnails.m_oScrollVerApi.scrollVCurrentY, 0);
    viewer.file = null;
    getPageCalls = 0;

    assert.equal(viewer.GetPageForThumbnails(0, 100, 100), null);
    assert.equal(viewer.GetPrintPage(0, 100, 100, 0), null);
    assert.equal(viewer.Copy({Text: ''}), false);
    assert.equal(viewer.findText({text: 'a'}, true), false);
    assert.equal(viewer.getTextCommandsSize(), 0);
    viewer.selectAll();
    viewer.navigate(0);

    // Closed-file PDF helpers & asc_CloseFile with active ink drawer do not throw.
    const updatePagesTransform = loadMethod('pdf/src/document.js', 'CPDFDoc.prototype.', 'UpdatePagesTransform', {AscCommon: localAscCommon});
    updatePagesTransform.call({Viewer: viewer, pagesTransform: [1]});
    const textTrackOnChange = loadMethod('pdf/src/textSelectTrackHandler.js', 'CTextSelectTrackHandler.prototype.', 'OnChangePosition', {Asc: localAsc});
    let hiddenTrack = 0;
    textTrackOnChange.call({OnHide: () => { hiddenTrack++; }, GetBounds: () => null, OnShow: noop}, false);
    assert.equal(hiddenTrack, 1);

    let overlayUpdated = 0;
    viewer.onUpdateOverlay = () => { overlayUpdated++; };
    viewer.DrawingObjects = {onInkDrawerChangeState: noop};
    viewer.setCursorType = noop;
    const pdfApi = {
        DocumentRenderer: viewer,
        WordControl: {OnUpdateOverlay: noop},
        getDocumentRenderer: () => viewer,
        getPDFDoc: () => pdfDoc,
        isInkDrawerOn: () => false
    };
    for (const name of ['onInkDrawerChangeState', 'asc_searchEnabled', 'asc_findText', 'asc_endFindText', 'asc_isSelectSearchingResults', 'asc_EditSelectAll']) {
        pdfApi[name] = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', name, {Asc: localAsc, AscCommon: localAscCommon});
    }
    pdfApi.asc_getPageSize = loadMethod('pdf/api.js', 'PDFEditorApi.prototype["asc_getPageSize"]', '', {Asc: localAsc, AscCommon: localAscCommon});
    pdfApi.onInkDrawerChangeState();
    assert.equal(pdfApi.asc_getPageSize(0), null);
    pdfApi.asc_searchEnabled(false);
    assert.equal(pdfApi.asc_findText({text: 'x'}, true), 0);
    pdfApi.asc_endFindText();
    assert.equal(pdfApi.asc_isSelectSearchingResults(), false);
    pdfApi.asc_EditSelectAll();
    assert.ok(overlayUpdated >= 2);

    // K2 arrives while thumbnails panel is closed: init() sets needResize, then opening panel renders K2 on the live rAF loop.
    panelEl.offsetWidth = 0;
    viewer.file = makeFile(1);
    viewer.pagesInfo = {pages: [{IsLocked: () => false, isPainted: false, links: null}], countTextPages: 0};
    viewer.drawingPages = [{W: 600, H: 800, Y: 0}];
    thumbnails.init();
    assert.equal(thumbnails.pages.length, 1);
    assert.equal(thumbnails.isNeedResize(), true);

    panelEl.offsetWidth = 220;
    viewer.RequestAnimationOldTime = -1;
    rafCallbacks.shift()();
    assert.equal(rafCallbacks.length, 1);
    assert.equal(getPageCalls, 1);
    assert.ok(thumbnails.pages[0].image);

    // timerAnimation keeps Chrome rAF alive in finally even if a tick throws.
    viewer.timer = () => { throw new Error('boom'); };
    viewer.RequestAnimationOldTime = -1;
    assert.throws(() => rafCallbacks.shift()(), /boom/);
    assert.equal(rafCallbacks.length, 1);
});

test('Word refreshFile cancels active and queued plugin pastes without running queued commands under GlobalLock', () => {
    let globalLockCount = 0;
    const pasteDataCalls = [];
    const evalCalls = [];
    const events = [];
    const createdNodes = [];
    const baseCommon = {
        c_oEditorId: {Word: 0, Spreadsheet: 1, Presentation: 2, Pdf: 3, Visio: 4},
        c_oAscClipboardDataFormat: {HtmlElement: 1, Text: 2},
        changestype_Paragraph_Content: 1,
        changestype_Image_Properties: 2,
        c_oCloseCode: {quiet: 0, updateVersion: 1, noCache: 2, restore: 3, sessionIdle: 4},
        ConnectionState: {ClosedAll: 4},
        getDisconnectErrorCode: () => 1,
        getEnableDownloadByCloseCode: () => false,
        g_clipboardBase: {bSaveFormat: true},
        g_specialPasteHelper: {Paste_Process_Start: noop},
        checkOnlyOneImage: () => false,
        History: {Clear: noop, Is_LastPointEmpty: () => true},
        g_oIdCounter: {Clear: noop},
        g_oTableId: {Clear: noop},
        pptx_content_loader: {ImageMapChecker: {}},
        CollaborativeEditing: {
            Clear: noop,
            Set_GlobalLock(value) { globalLockCount = value ? globalLockCount + 1 : Math.max(0, globalLockCount - 1); },
            Get_GlobalLock() { return globalLockCount > 0; }
        },
        Editor_Paste_Exec(_api, format, data1, _d2, _td, _u, callback) {
            pasteDataCalls.push({format, data1, callback, locked: globalLockCount > 0});
        },
        CDocsCoApi: function() {
            this.connect = noop;
            this.disconnect = noop;
            this.get_state = () => 0;
        },
        sendClientLog: noop,
        getClientInfoString: () => '',
        getMemoryInfo: () => ''
    };
    const mockDocument = {
        body: {
            appendChild(el) { el.parentNode = this; createdNodes.push(el); },
            removeChild(el) { el.parentNode = null; },
            contains(el) { return el.parentNode === this; }
        },
        createElement() { return {style: {}, parentNode: null}; },
        getElementById(id) { return createdNodes.find(n => n.id === id && n.parentNode) || null; }
    };
    const mockWindow = {
        IS_NATIVE_EDITOR: true,
        AscCommon: baseCommon
    };
    const Asc = {
        c_oAscAsyncActionType: {BlockInteraction: 1, Information: 2, Empty: 0},
        c_oAscAsyncAction: {Open: 1, Disconnect: 2, RefreshFile: 3},
        c_oAscRestrictionType: {None: 0, View: 1},
        c_oAscError: {ID: {Unknown: 1, Drop: 2, UpdateVersion: 3}, Level: {NoCritical: 0, Critical: 1}},
        c_nCommonRequestTime: 1000
    };
    const globals = {
        window: mockWindow,
        document: mockDocument,
        performance: {now: () => 0},
        Asc,
        AscCommon: baseCommon,
        AscDFH: {historydescription_Document_PasteHotKey: 1},
        AscFonts: {
            FontPickerByCharacter: {
                CallbackObj: {},
                checkText(_text, ctx, cb) {
                    this.CallbackObj._this = ctx;
                    this.CallbackObj._callback = cb;
                }
            }
        },
        g_oIdCounter: baseCommon.g_oIdCounter,
        g_oTableId: baseCommon.g_oTableId,
        c_oAscAsyncActionType: Asc.c_oAscAsyncActionType,
        c_oAscAsyncAction: Asc.c_oAscAsyncAction,
        c_oAscRestrictionType: Asc.c_oAscRestrictionType,
        c_oAscError: Asc.c_oAscError,
        c_oCloseCode: baseCommon.c_oCloseCode,
        setTimeout: noop,
        clearTimeout: noop
    };

    const api = {
        editorId: 0,
        isMobileVersion: false,
        isEmbedVersion: false,
        documentIsWopi: false,
        DocInfo: {get_Id: () => 'doc-1', asc_getOptions: () => ({})},
        isDocumentLoadComplete: true,
        isOnLoadLicense: true,
        isApplyChangesOnOpenEnabled: true,
        ServerIdWaitComplete: true,
        ServerImagesWaitComplete: true,
        FontLoadWaitComplete: true,
        documentLoadGeneration: 0,
        ServerIdWaitAction: null,
        IsLongActionCurrent: 0,
        LongActionCallbacks: [],
        LongActionCallbacksParams: [],
        IsActionRestrictionCurrent: 0,
        IsActionRestrictionPrev: null,
        isDisconnectActionActive: false,
        isOpenActionActive: false,
        isRefreshInProgress: false,
        isAbortingDocumentLoad: false,
        isPluginQueuePausedForRefresh: false,
        _activeAbortNotifications: null,
        restrictions: 0,
        WordControl: {
            m_oLogicDocument: {
                IsSelectionLocked: () => false,
                IsFormFieldEditing: () => false,
                StartAction: noop,
                StopRecalculate: noop,
                StopSpellCheck: noop,
                FinalizeAction: noop,
                IsDocumentEditor: () => false
            },
            m_oDrawingDocument: {CloseFile: noop}
        },
        canEdit: () => true,
        isLogPluginCommands: () => false,
        isLongAction() { return this.IsLongActionCurrent > 0; },
        canRunBuilderScript: () => true,
        _beforeEvalCommand: noop,
        onEndBuilderScript: cb => cb(true),
        executeGroupActions: fn => fn(),
        executeGroupActionsStart: noop,
        executeGroupActionsEnd: noop,
        checkLongActionCallback: () => true,
        getEditorId: () => 0,
        get_TextProps: () => null,
        getJsApi: () => ({ReplaceTextSmart: noop}),
        asc_canPaste: () => true,
        asc_Recalculate: noop,
        asc_setRestriction(value) { this.restrictions = value; },
        getRestrictionSettings: () => null,
        asc_setDocInfo(info) { this.DocInfo = info; },
        isDocumentModified: () => false,
        asc_checkNeedCallback: () => true,
        setViewModeDisconnect: noop,
        setOpenedAt: noop,
        turnOffSpecialModes: noop,
        _openDocumentEndCallback: noop,
        _coAuthoringInit: noop,
        pluginMethod_SetProperties: noop,
        pluginMethod_GetVersion: () => '9.4.0',
        sendEvent(name, ...args) { events.push([name, ...args]); },
        CoAuthoringApi: new baseCommon.CDocsCoApi()
    };
    for (const name of [
        ...loadCancellationMethods,
        '_addPendingPluginPasteCancel', '_removePendingPluginPasteCancel', '_cancelPendingPluginPastes',
        'sync_StartAction', 'sync_EndAction', 'isActionWithRestriction', 'isLongActionBase',
        'incrementCounterLongAction', 'decrementCounterLongAction',
        'incrementCounterActionRestriction', 'decrementCounterActionRestriction',
        'clearWaitNotIdle', 'endDisconnectAction', 'endOpenAction',
        'canRefreshFile', 'onRefreshFile', 'asc_refreshFile',
        'reopenFileWithReconnection', 'asyncServerIdEndLoaded', 'onDocumentContentReady'
    ]) {
        api[name] = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    }
    api.asc_PasteData = loadMethod('word/api.js', 'asc_docs_api.prototype.', 'asc_PasteData', globals);
    api.asc_CloseFile = loadMethod('word/api.js', 'asc_docs_api.prototype.', 'asc_CloseFile', globals);
    api.pluginMethod_PasteText = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteText"]', '', globals);
    api.pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteHtml"]', '', globals);
    api._pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype.', '_pluginMethod_PasteHtml', globals);
    api.pluginMethod_ReplaceTextSmart = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_ReplaceTextSmart"]', '', globals);

    const {mgr, sentMessages} = loadRealPluginsManager(api, {
        safePluginEval: code => { evalCalls.push(code); return 'eval-ok'; }
    });
    mockWindow.g_asc_plugins = mgr;
    mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

    // 1. Exact regression: 1 active + 2 queued PasteText during refreshFile.
    mgr.callMethod('p1', 'PasteText', ['first']);
    mgr.callMethod('p1', 'PasteText', ['second']);
    mgr.callMethod('p1', 'PasteText', ['third']);
    assert.equal(pasteDataCalls.length, 1);
    assert.equal(api.IsLongActionCurrent, 1);
    assert.equal(mgr.queueCommands.length, 3);

    api.onRefreshFile();
    assert.equal(globalLockCount > 0, true);
    assert.equal(api.isPluginQueueSuspended(), true);

    api.asc_refreshFile({get_Id: () => 'doc-2', asc_getOptions: () => ({})});
    assert.equal(pasteDataCalls.length, 1, 'queued pastes must not execute while locked by RefreshFile');
    assert.equal(api.IsLongActionCurrent, 0, 'no leaked long action after refresh close/reopen');
    assert.equal(mgr.queueCommands.length, 0, 'all stale K1 plugin commands settled');
    assert.deepEqual(sentMessages.map(m => m.methodReturnData), [false, false, false]);

    // 2. Command arriving during K2 load (even after asyncServerIdEndLoaded before onDocumentContentReady) pauses until ready.
    api.asyncServerIdEndLoaded(2);
    assert.equal(api.isPluginQueuePausedForRefresh, true);
    mgr.callMethod('p1', 'PasteText', ['after-ready']);
    assert.equal(pasteDataCalls.length, 1);
    assert.equal(mgr.queueCommands.length, 1);

    api.onDocumentContentReady();
    assert.equal(api.isPluginQueuePausedForRefresh, false);
    assert.equal(pasteDataCalls.length, 2);
    assert.equal(pasteDataCalls[1].data1, 'after-ready');
    assert.equal(pasteDataCalls[1].locked, false);
    pasteDataCalls[1].callback(true);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(mgr.queueCommands.length, 0);
    assert.equal(sentMessages.at(-1).methodReturnData, true);

    // 3. Mixed queue [PasteText (active), PasteHtml, ReplaceTextSmart, CallCommand, internal method] during refresh.
    sentMessages.length = 0;
    const internalResults = [];
    mgr.callMethod('p1', 'PasteText', ['m-active']);
    mgr.callMethod('p1', 'PasteHtml', ['<b>hi</b>']);
    mgr.callMethod('p1', 'ReplaceTextSmart', [['a']]);
    mgr.callCommand('p1', 'Api.GetDocument()', false, false, false, false);
    mgr.internalCallbacks.push(res => internalResults.push(res));
    mgr.callMethod(mgr.internalGuid, 'GetVersion', []);

    api.asc_refreshFile({get_Id: () => 'doc-3', asc_getOptions: () => ({})});
    assert.equal(evalCalls.length, 0);
    assert.equal(createdNodes.length, 0);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(mgr.queueCommands.length, 0);
    assert.equal(mgr.internalCallbacks.length, 0);
    assert.deepEqual(internalResults, [false]);
    assert.deepEqual(sentMessages.map(m => [m.type, m.methodReturnData, m.commandReturnData]), [
        ['onMethodReturn', false, undefined],
        ['onMethodReturn', false, undefined],
        ['onMethodReturn', false, undefined],
        ['onCommandCallback', undefined, undefined]
    ]);
    api.asyncServerIdEndLoaded(3);
    api.onDocumentContentReady();

    // 4. Synchronous pluginMethod_OpenFile calling asc_CloseFile() preserves queued follow-up method.
    sentMessages.length = 0;
    api.pluginMethod_OpenFile = function() {
        this.asc_CloseFile();
        return 'opened';
    };
    mgr.callMethod('p1', 'OpenFile', []);
    assert.equal(sentMessages.at(-1).methodReturnData, 'opened');
    // Queue another method while OpenFile is running via nested callMethod.
    sentMessages.length = 0;
    api.pluginMethod_OpenFile = function() {
        mgr.callMethod('p1', 'GetVersion', []);
        this.asc_CloseFile();
        return 'opened-2';
    };
    mgr.callMethod('p1', 'OpenFile', []);
    assert.deepEqual(sentMessages.map(m => m.methodReturnData), ['opened-2', '9.4.0']);

    // 5. Synchronous completion across 2 plugins does not double-shift P2's active task.
    sentMessages.length = 0;
    mgr.runnedPluginsMap['p2'] = {frameId: 'f2', isConnector: false, methodReturnAsync: false};
    api.isDocumentLoadComplete = true;
    api.isRefreshInProgress = false;
    mgr.callMethod('p1', 'PasteText', ['p1-active']);
    mgr.callMethod('p2', 'PasteText', ['p2-queued']);
    // Lock editor so p2-queued rejects synchronously when p1 finishes.
    baseCommon.CollaborativeEditing.Set_GlobalLock(true);
    pasteDataCalls.at(-1).callback(true);
    baseCommon.CollaborativeEditing.Set_GlobalLock(false);
    assert.equal(api.IsLongActionCurrent, 0);
    assert.equal(mgr.queueCommands.length, 0);
    assert.deepEqual(sentMessages.map(m => [m.guid, m.methodReturnData]), [
        ['p1', true],
        ['p2', false]
    ]);
});

test('direct locked PasteText/PasteHtml/ReplaceTextSmart reject cleanly without leaking CounterLongAction across all 5 editors', () => {
    const editors = [
        {name: 'word', file: 'word/api.js', prefix: 'asc_docs_api.prototype.', extra: {WordControl: {m_oLogicDocument: {}}, canEdit: () => true}},
        {name: 'cell', file: 'cell/api.js', prefix: 'spreadsheet_api.prototype.', extra: {canEdit: () => true, wb: {getWorksheet: () => ({}), pasteData: () => { throw new Error('must not paste under lock'); }}, collaborativeEditing: {getGlobalLock: () => true}}},
        {name: 'slide', file: 'slide/api.js', prefix: 'asc_docs_api.prototype.', extra: {canEdit: () => true, isSlideShow: () => false}},
        {name: 'pdf', file: 'pdf/api.js', prefix: 'PDFEditorApi.prototype.', extra: {canEdit: () => true, DocumentRenderer: {file: {isValid: () => true}, canInteract: () => true, getPDFDoc: () => ({StartAction: noop, FinalizeAction: noop})}}},
        {name: 'visio', file: 'visio/api.js', prefix: 'VisioEditorApi.prototype.', extra: {canEdit: () => true}}
    ];
    for (const mode of ['refresh-in-progress', 'global-lock']) {
        for (const spec of editors) {
            const mockDocument = {
                body: {appendChild: noop, removeChild: noop, contains: () => false},
                createElement: () => ({style: {}}),
                getElementById: () => null
            };
            const localCommon = {
                c_oEditorId: {Word: 0, Spreadsheet: 1, Presentation: 2},
                c_oAscClipboardDataFormat: {HtmlElement: 1, Text: 2},
                g_clipboardBase: {bSaveFormat: true},
                CollaborativeEditing: {Get_GlobalLock: () => mode === 'global-lock'}
            };
            const mockWindow = {AscCommon: localCommon};
            const globals = {
                window: mockWindow,
                document: mockDocument,
                AscCommon: localCommon,
                AscFonts: {
                    FontPickerByCharacter: {
                        CallbackObj: {},
                        checkText(_text, ctx, cb) { cb.call(ctx); }
                    }
                }
            };
            const api = {
                editorId: 0,
                IsLongActionCurrent: 0,
                LongActionCallbacks: [],
                LongActionCallbacksParams: [],
                isRefreshInProgress: mode === 'refresh-in-progress',
                canEdit: () => true,
                isLogPluginCommands: () => false,
                executeGroupActions: fn => fn(),
                executeGroupActionsStart: noop,
                executeGroupActionsEnd: noop,
                get_TextProps: () => null,
                getJsApi: () => ({ReplaceTextSmart: () => { throw new Error('must not run ReplaceTextSmart under lock'); }}),
                asc_canPaste: () => true,
                ...spec.extra
            };
            api.isLongActionBase = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'isLongActionBase', globals);
            api.incrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'incrementCounterLongAction', globals);
            api.decrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'decrementCounterLongAction', globals);
            api.asc_PasteData = loadMethod(spec.file, spec.prefix, 'asc_PasteData', globals);
            api.pluginMethod_PasteText = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteText"]', '', globals);
            api.pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteHtml"]', '', globals);
            api._pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype.', '_pluginMethod_PasteHtml', globals);
            api.pluginMethod_ReplaceTextSmart = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_ReplaceTextSmart"]', '', globals);

            const {mgr, sentMessages} = loadRealPluginsManager(api);
            mockWindow.g_asc_plugins = mgr;
            mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

            mgr.callMethod('p1', 'PasteText', ['hello']);
            mgr.callMethod('p1', 'PasteHtml', ['<p>hello</p>']);
            mgr.callMethod('p1', 'ReplaceTextSmart', [['hello']]);

            assert.equal(api.IsLongActionCurrent, 0, `${spec.name}/${mode}: IsLongActionCurrent must stay 0`);
            assert.equal(mgr.queueCommands.length, 0, `${spec.name}/${mode}: queueCommands must drain cleanly`);
            assert.deepEqual(sentMessages.map(m => m.methodReturnData), [false, false, false], `${spec.name}/${mode}: all 3 methods reject cleanly`);
        }
    }

    // PDF asc_PasteData rejects immediately on an invalid/closed file, and defers onCancel into _activeAbortNotifications on close().
    {
        const localCommon = {
            c_oAscClipboardDataFormat: {HtmlElement: 1, Text: 2},
            CollaborativeEditing: {Get_GlobalLock: () => false}
        };
        let paintCalls = 0;
        const closedViewer = {
            file: {isValid: () => false},
            canInteract: () => false,
            paint: () => { paintCalls++; }
        };
        const pdfApi = {
            DocumentRenderer: closedViewer,
            documentLoadGeneration: 0
        };
        pdfApi.asc_PasteData = loadMethod('pdf/api.js', 'PDFEditorApi.prototype.', 'asc_PasteData', {AscCommon: localCommon});
        let rejected = 0;
        pdfApi.asc_PasteData(2, 'hi', undefined, undefined, undefined, noop, undefined, () => { rejected++; });
        assert.equal(rejected, 1);
        assert.equal(paintCalls, 0);

        // Deferred asc_PasteData receives rejectCallback on close() while plain onAfterPaintCallback (e.g. OpenChartEditor) is dropped silently.
        const viewerGlobals = {clearTimeout: noop, clearInterval: noop, setTimeout: () => 1, CDocumentPagesInfo: function() {}};
        const openViewer = {
            Api: null,
            openSessionId: 0,
            checkReadyTimerId: null,
            scheduledRepaintTimer: null,
            onRepaintCallbacks: [],
            onAfterPaintCallback: [],
            file: {isValid: () => true, close: noop},
            canInteract: () => false,
            getPDFDoc: () => ({BlurActiveObject: noop}),
            onUpdateOverlay: noop
        };
        for (const name of ['_cancelPendingPaintCallbacks', 'cancelPendingLoad', 'close', 'scheduleRepaint', 'paint', 'afterPaintCallbacks']) {
            openViewer[name] = loadMethod('pdf/src/viewer.js', 'this.', name, viewerGlobals);
        }
        let plainCalled = 0;
        let pasteRejected = 0;
        openViewer.paint(null, () => { plainCalled++; });
        pdfApi.DocumentRenderer = openViewer;
        pdfApi.asc_PasteData(2, 'hi', undefined, undefined, undefined, noop, undefined, () => { pasteRejected++; });
        openViewer.close();
        assert.equal(plainCalled, 0);
        assert.equal(pasteRejected, 1);

        // Inside _abortPendingDocumentLoad, close() queues onCancel into Api._activeAbortNotifications instead of firing mid-cleanup.
        const deferred = [];
        openViewer.Api = {_activeAbortNotifications: deferred};
        openViewer.file = {isValid: () => true, close: noop};
        let abortRejected = 0;
        pdfApi.asc_PasteData(2, 'hi', undefined, undefined, undefined, noop, undefined, () => { abortRejected++; });
        openViewer.cancelPendingLoad();
        assert.equal(abortRejected, 0, 'onCancel is deferred into _activeAbortNotifications');
        assert.equal(deferred.length, 1);
        deferred[0]();
        assert.equal(abortRejected, 1);
    }
});

test('Cell applyChanges, partial model cleanup, action handoff, and Slide/Visio/FontLoader re-entrancy stay guarded', () => {
    // 1. Cell collaborativeEditing.applyChanges stops if asc_onBeforeApplyChanges synchronously refreshes.
    {
        let triggerCalls = 0;
        const sentEvents = [];
        const mockApi = {
            documentLoadGeneration: 0,
            sendEvent(name) {
                sentEvents.push(name);
                if (name === 'asc_onBeforeApplyChanges') {
                    this.documentLoadGeneration++;
                }
            }
        };
        const applyChanges = loadMethod('cell/model/CollaborativeEditing.js', 'CCollaborativeEditing.prototype.', 'applyChanges', {
            Asc: {editor: mockApi}
        });
        const collab = {
            m_arrChanges: [{id: 1}],
            handlers: {trigger: () => { triggerCalls++; }}
        };
        assert.equal(applyChanges.call(collab, null), false);
        assert.equal(triggerCalls, 0);
        assert.deepEqual(sentEvents, ['asc_onBeforeApplyChanges']);
        assert.equal(collab.m_arrChanges.length, 1);
    }

    // 2. Cell partial model cleanup runs on abort during font load (before WorkbookView exists).
    {
        let closedDrawing = 0;
        let clearedStyles = 0;
        const cellGlobals = {
            AscCommon: {
                baseEditorsApi: {prototype: {_abortPendingDocumentLoad: noop}}
            },
            AscCommonExcel: {
                oFormulaLocaleInfo: {Parse: true, DigitSep: true},
                g_StyleCache: {Clear: () => { clearedStyles++; }}
            },
            History: {TurnOn: noop}
        };
        const abortCell = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_abortPendingDocumentLoad', cellGlobals);
        const cellApi = {
            isDocumentLoadComplete: false,
            _hasPendingWorkbookView: false,
            _hasPendingWorkbookModel: true,
            wb: null,
            wbModel: {DrawingDocument: {CloseFile: () => { closedDrawing++; }}}
        };
        abortCell.call(cellApi, true);
        assert.equal(closedDrawing, 1);
        assert.equal(clearedStyles, 1);
        assert.equal(cellApi._hasPendingWorkbookModel, false);
        abortCell.call(cellApi, true);
        assert.equal(closedDrawing, 1);
        assert.equal(clearedStyles, 1);
    }

    // 3. startFontLoadAction / startImageLoadAction do not transiently flush LongActionCallbacks.
    {
        let flushed = 0;
        const globals = {
            Asc: {c_oAscAsyncActionType: {BlockInteraction: 1}},
            AscCommon: {CollaborativeEditing: {Set_GlobalLock: noop}},
            c_oAscAsyncActionType: {BlockInteraction: 1, Empty: 0}
        };
        const api = {
            IsLongActionCurrent: 0,
            LongActionCallbacks: [() => { flushed++; }],
            LongActionCallbacksParams: [undefined],
            activeFontLoadAction: null,
            activeImageLoadAction: null,
            sendEvent: noop
        };
        for (const name of ['sync_StartAction', 'sync_EndAction', 'isLongActionBase', 'incrementCounterLongAction', 'decrementCounterLongAction', 'startFontLoadAction', 'endFontLoadAction', 'startImageLoadAction', 'endImageLoadAction']) {
            api[name] = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
        }
        api.startFontLoadAction(1, 10);
        api.startFontLoadAction(1, 11);
        assert.equal(flushed, 0);
        assert.equal(api.IsLongActionCurrent, 1);
        api.startImageLoadAction(1, 20);
        api.startImageLoadAction(1, 21);
        assert.equal(flushed, 0);
        assert.equal(api.IsLongActionCurrent, 2);
        api.endFontLoadAction();
        api.endImageLoadAction();
        assert.equal(api.IsLongActionCurrent, 0);
        assert.equal(flushed, 1);
    }

    // 4. Visio _openDocumentEndCallback stops opening tail if onDocumentContentReady synchronously reopens.
    {
        let initControlCalls = 0;
        const openEnd = loadMethod('visio/api.js', 'VisioEditorApi.prototype.', '_openDocumentEndCallback', {AscCommon});
        const visioApi = {
            isDocumentLoadComplete: false,
            ServerImagesWaitComplete: true,
            ServerIdWaitComplete: true,
            documentLoadGeneration: 0,
            WordControl: {
                m_oLogicDocument: {
                    toCShapes: noop,
                    Recalculate: noop,
                    DrawingDocument: {OnEndRecalculate: noop}
                },
                GoToPage: noop,
                InitControl: () => { initControlCalls++; }
            },
            Document: {getCurrentPage: () => 0},
            onDocumentContentReady() { this.documentLoadGeneration++; }
        };
        openEnd.call(visioApi);
        assert.equal(initControlCalls, 0);
    }

    // 5. GlobalLoaders.LoadDocumentFonts stops if asyncFontsDocumentStartLoaded synchronously cancels the session.
    {
        let checkCalled = 0;
        let loadCalled = 0;
        const loader = {
            check_loaded_timer_id: -1,
            endLoadingCallback: null,
            IsLoadDocumentFonts2: false,
            documentLoadSessionId: 0,
            fontInfos: [],
            startDocumentLoad() { this.documentLoadSessionId++; },
            cancelDocumentLoad() { this.documentLoadSessionId++; },
            AddLoadFonts: noop,
            CheckFontsNeedLoadingLoad: () => { checkCalled++; },
            _LoadFonts: () => { loadCalled++; },
            Api: {
                sync_InitEditorFonts: noop,
                IsNeedDefaultFonts: () => false,
                asyncFontsDocumentStartLoaded: () => { loader.cancelDocumentLoad(); }
            }
        };
        loader.LoadDocumentFonts = loadMethod('common/GlobalLoaders.js', 'this.', 'LoadDocumentFonts', {AscFonts: {CFont: function() {}}});
        loader.LoadDocumentFonts({});
        assert.equal(checkCalled, 0);
        assert.equal(loadCalled, 0);
    }
});

test('plugin queue exception recovery, pause safety, ReplaceTextSmart throw cleanup, and PDF getViewingRect stay safe', () => {
    // 1. Throwing plugin method releases isCommandRunning; out-of-plugin onPluginMethodReturn while paused does not pop queued task;
    //    cancelling queued task for the same plugin does not clobber active task's methodReturnAsync.
    {
        const api = {
            documentLoadGeneration: 0,
            isPluginQueuePausedForRefresh: false,
            isPluginQueueSuspended() { return this.isPluginQueuePausedForRefresh; },
            isLogPluginCommands: () => false,
            executeGroupActions: fn => fn(),
            pluginMethod_Boom() { throw new Error('plugin-boom'); },
            pluginMethod_Ok: () => 'ok',
            pluginMethod_AsyncHold() {
                windowObj.g_asc_plugins.setPluginMethodReturnAsync();
            }
        };
        const {mgr, sentMessages} = loadRealPluginsManager(api);
        const windowObj = {g_asc_plugins: mgr};
        mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

        assert.throws(() => mgr.callMethod('p1', 'Boom', []), /plugin-boom/);
        assert.equal(mgr.isCommandRunning, false);
        assert.equal(mgr.queueCommands.length, 0);

        mgr.callMethod('p1', 'Ok', []);
        assert.equal(sentMessages.at(-1).methodReturnData, 'ok');

        // Pause queue and verify stray onPluginMethodReturn does not pop unstarted queued command.
        api.isPluginQueuePausedForRefresh = true;
        mgr.callMethod('p1', 'Ok', []);
        assert.equal(mgr.isCommandRunning, false);
        assert.equal(mgr.queueCommands.length, 1);
        mgr.onPluginMethodReturn('stray');
        assert.equal(mgr.queueCommands.length, 1);

        // Unpause and verify queued command executes normally.
        api.isPluginQueuePausedForRefresh = false;
        mgr.resumeQueue();
        assert.equal(mgr.queueCommands.length, 0);
        assert.equal(sentMessages.at(-1).methodReturnData, 'ok');

        // Active async task at index 0 keeps methodReturnAsync=true when queued task at index 1 is cancelled.
        mgr.callMethod('p1', 'AsyncHold', []);
        mgr.callMethod('p1', 'Ok', []);
        assert.equal(mgr.runnedPluginsMap['p1'].methodReturnAsync, true);
        mgr.cancelQueuedCommands(false, 1);
        assert.equal(mgr.queueCommands.length, 1);
        assert.equal(mgr.runnedPluginsMap['p1'].methodReturnAsync, true);
        mgr.onPluginMethodReturn('done');
        assert.equal(mgr.queueCommands.length, 0);
        assert.equal(mgr.runnedPluginsMap['p1'].methodReturnAsync, false);
    }

    // 2. Sync, async, LongActionCallback, and generation-change failures inside ReplaceTextSmart keep plugin queue unblocked.
    {
        for (const mode of ['sync-throw', 'async-throw', 'async-longaction-throw', 'generation-change']) {
            let pendingCheck = null;
            let recalcCalls = 0;
            const mockWindow = {};
            const globals = {
                window: mockWindow,
                AscCommon: {c_oEditorId: {Word: 0, Presentation: 1, Spreadsheet: 2}, CollaborativeEditing: {Get_GlobalLock: () => false}},
                AscFonts: {
                    FontPickerByCharacter: {
                        CallbackObj: {},
                        checkText(_text, ctx, cb) {
                            if (mode === 'sync-throw') {
                                cb.call(ctx);
                            } else {
                                pendingCheck = () => cb.call(ctx);
                            }
                        }
                    }
                }
            };
            const api = {
                editorId: 0,
                documentLoadGeneration: 0,
                IsLongActionCurrent: 0,
                LongActionCallbacks: [],
                LongActionCallbacksParams: [],
                isRefreshInProgress: false,
                isLogPluginCommands: () => false,
                executeGroupActions: fn => fn(),
                getJsApi: () => ({
                    ReplaceTextSmart() {
                        if (mode === 'sync-throw' || mode === 'async-throw') {
                            throw new Error('replace-boom');
                        }
                        if (mode === 'generation-change') {
                            api.documentLoadGeneration++;
                        }
                        if (mode === 'async-longaction-throw') {
                            api.LongActionCallbacks.push(() => { throw new Error('long-action-boom'); });
                            api.LongActionCallbacksParams.push(null);
                        }
                    }
                }),
                WordControl: {m_oLogicDocument: {FinalizeAction: noop}},
                asc_canPaste: () => true,
                asc_Recalculate: () => { recalcCalls++; },
                pluginMethod_Next: () => 'next-ok'
            };
            api.isLongActionBase = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'isLongActionBase', globals);
            api.incrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'incrementCounterLongAction', globals);
            api.decrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'decrementCounterLongAction', globals);
            api.pluginMethod_ReplaceTextSmart = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_ReplaceTextSmart"]', '', globals);

            const {mgr, sentMessages} = loadRealPluginsManager(api);
            mockWindow.g_asc_plugins = mgr;
            mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

            mgr.callMethod('p1', 'ReplaceTextSmart', [['hello']]);
            mgr.callMethod('p1', 'Next', []);
            if (pendingCheck) {
                if (mode === 'async-longaction-throw') {
                    assert.throws(() => pendingCheck(), /long-action-boom/);
                } else {
                    pendingCheck();
                }
            }
            assert.equal(api.IsLongActionCurrent, 0, `${mode}: IsLongActionCurrent must return to 0`);
            assert.equal(mgr.isCommandRunning, false, `${mode}: plugin queue must not deadlock`);
            assert.equal(mgr.runnedPluginsMap['p1'].methodReturnAsync, false, `${mode}: methodReturnAsync cleared`);
            assert.equal(mgr.queueCommands.length, 0, `${mode}: subsequent command ran`);
            assert.deepEqual(
                sentMessages.map(m => m.methodReturnData),
                [mode === 'async-longaction-throw', 'next-ok'],
                `${mode}: ReplaceTextSmart reports result and advances queue`
            );
            if (mode === 'generation-change') {
                assert.equal(recalcCalls, 0);
            }
        }

        // Async ReplaceTextSmart cancelReplace still notifies plugin if releaseLongAction throws in LongActionCallbacks.
        {
            let pendingCheck = null;
            const mockWindow = {};
            const globals = {
                window: mockWindow,
                AscCommon: {c_oEditorId: {Word: 0, Presentation: 1, Spreadsheet: 2}, CollaborativeEditing: {Get_GlobalLock: () => false}},
                AscFonts: {
                    FontPickerByCharacter: {
                        CallbackObj: {},
                        checkText(_text, ctx, cb) { pendingCheck = () => cb.call(ctx); }
                    }
                }
            };
            const api = {
                editorId: 0,
                documentLoadGeneration: 0,
                IsLongActionCurrent: 0,
                LongActionCallbacks: [],
                LongActionCallbacksParams: [],
                isRefreshInProgress: false,
                isLogPluginCommands: () => false,
                executeGroupActions: fn => fn(),
                getJsApi: () => ({ReplaceTextSmart: noop}),
                asc_canPaste: () => true,
                pluginMethod_Next: () => 'next-ok'
            };
            api.isLongActionBase = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'isLongActionBase', globals);
            api.incrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'incrementCounterLongAction', globals);
            api.decrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'decrementCounterLongAction', globals);
            api.pluginMethod_ReplaceTextSmart = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_ReplaceTextSmart"]', '', globals);
            const {mgr, sentMessages} = loadRealPluginsManager(api);
            mockWindow.g_asc_plugins = mgr;
            mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

            mgr.callMethod('p1', 'ReplaceTextSmart', [['hello']]);
            mgr.callMethod('p1', 'Next', []);
            api.documentLoadGeneration++;
            api.LongActionCallbacks.push(() => { throw new Error('cancel-longaction-boom'); });
            api.LongActionCallbacksParams.push(null);
            assert.throws(() => pendingCheck(), /cancel-longaction-boom/);
            assert.equal(api.IsLongActionCurrent, 0);
            assert.equal(mgr.isCommandRunning, false);
            assert.equal(mgr.queueCommands.length, 0);
            assert.deepEqual(sentMessages.map(m => m.methodReturnData), [false, 'next-ok']);
        }

        // Async PasteHtml completion still cleans up DOM, bSaveFormat, group actions, and plugin queue if releaseLongAction throws.
        {
            const bodyChildren = [];
            const mockDocument = {
                body: {
                    appendChild: el => { el.parentNode = mockDocument.body; bodyChildren.push(el); },
                    removeChild: el => {
                        const idx = bodyChildren.indexOf(el);
                        if (idx !== -1) bodyChildren.splice(idx, 1);
                        el.parentNode = null;
                    },
                    contains: el => bodyChildren.includes(el)
                },
                createElement: () => ({style: {}}),
                getElementById: () => null
            };
            const localCommon = {
                c_oEditorId: {Word: 0, Spreadsheet: 1, Presentation: 2},
                c_oAscClipboardDataFormat: {HtmlElement: 1, Text: 2},
                g_clipboardBase: {bSaveFormat: true},
                CollaborativeEditing: {Get_GlobalLock: () => false}
            };
            const mockWindow = {AscCommon: localCommon};
            const globals = {window: mockWindow, document: mockDocument, AscCommon: localCommon};
            let pasteComplete = null;
            let groupEndCalls = 0;
            const api = {
                documentLoadGeneration: 0,
                IsLongActionCurrent: 0,
                LongActionCallbacks: [],
                LongActionCallbacksParams: [],
                isRefreshInProgress: false,
                canEdit: () => true,
                getEditorId: () => 2,
                isLogPluginCommands: () => false,
                executeGroupActions: fn => fn(),
                executeGroupActionsStart: noop,
                executeGroupActionsEnd: () => { groupEndCalls++; },
                get_TextProps: () => null,
                asc_PasteData(_fmt, _el, _d2, _txt, _cur, cb) { pasteComplete = cb; },
                pluginMethod_Next: () => 'next-ok'
            };
            api.isLongActionBase = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'isLongActionBase', globals);
            api.incrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'incrementCounterLongAction', globals);
            api.decrementCounterLongAction = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'decrementCounterLongAction', globals);
            api.checkLongActionCallback = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', 'checkLongActionCallback', globals);
            api.pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype["pluginMethod_PasteHtml"]', '', globals);
            api._pluginMethod_PasteHtml = loadMethod('common/apiBase_plugins.js', 'Api.prototype.', '_pluginMethod_PasteHtml', globals);

            const {mgr, sentMessages} = loadRealPluginsManager(api);
            mockWindow.g_asc_plugins = mgr;
            mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

            mgr.callMethod('p1', 'PasteHtml', ['<p>hi</p>']);
            mgr.callMethod('p1', 'Next', []);
            api.LongActionCallbacks.push(() => { throw new Error('paste-longaction-boom'); });
            api.LongActionCallbacksParams.push(null);
            assert.throws(() => pasteComplete(true), /paste-longaction-boom/);
            assert.equal(api.IsLongActionCurrent, 0);
            assert.equal(bodyChildren.length, 0);
            assert.equal(localCommon.g_clipboardBase.bSaveFormat, true);
            assert.equal(groupEndCalls, 1);
            assert.equal(mgr.isCommandRunning, false);
            assert.equal(mgr.queueCommands.length, 0);
            assert.deepEqual(sentMessages.map(m => m.methodReturnData), [true, 'next-ok']);
        }
    }

    // 3. PDF getViewingRect returns an empty rect when pageDetector is null after close().
    {
        const getViewingRect = loadMethod('pdf/src/viewer.js', 'CHtmlPage.prototype.', 'getViewingRect', {});
        const rect = getViewingRect.call({pageDetector: null}, 0);
        assert.equal(rect.num, 0);
        assert.equal(rect.x, 0);
        assert.equal(rect.r, 0);
    }

    // 4. LoadDocumentFonts stops if sync_InitEditorFonts synchronously cancels the load session.
    {
        let startLoadedCalls = 0;
        const loader = {
            check_loaded_timer_id: -1,
            endLoadingCallback: null,
            IsLoadDocumentFonts2: false,
            documentLoadSessionId: 0,
            fontInfos: [],
            startDocumentLoad() { this.documentLoadSessionId++; },
            cancelDocumentLoad() { this.documentLoadSessionId++; },
            AddLoadFonts: noop,
            CheckFontsNeedLoadingLoad: noop,
            _LoadFonts: noop,
            Api: {
                sync_InitEditorFonts: () => { loader.cancelDocumentLoad(); },
                IsNeedDefaultFonts: () => false,
                asyncFontsDocumentStartLoaded: () => { startLoadedCalls++; }
            }
        };
        loader.LoadDocumentFonts = loadMethod('common/GlobalLoaders.js', 'this.', 'LoadDocumentFonts', {AscFonts: {CFont: function() {}}});
        loader.LoadDocumentFonts({});
        assert.equal(startLoadedCalls, 0);
    }
});

test('abort scope recovers from LongActionCallbacks throw, flushes post-abort tasks, and preserves external suspendQueue', () => {
    const mockWindow = {};
    const globals = {window: mockWindow};
    const api = {
        documentLoadGeneration: 0,
        IsLongActionCurrent: 0,
        LongActionCallbacks: [],
        LongActionCallbacksParams: [],
        isAbortingDocumentLoad: false,
        isPluginQueuePausedForRefresh: false,
        isLogPluginCommands: () => false,
        executeGroupActions: fn => fn(),
        pluginMethod_Ping: () => 'pong'
    };
    for (const name of ['isLongActionBase', 'incrementCounterLongAction', 'decrementCounterLongAction', 'isPluginQueueSuspended', '_pausePluginQueueForRefresh', '_finishPluginQueuePause', '_runAbortScope']) {
        api[name] = loadMethod('common/apiBase.js', 'baseEditorsApi.prototype.', name, globals);
    }
    const {mgr, sentMessages} = loadRealPluginsManager(api);
    mockWindow.g_asc_plugins = mgr;
    mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

    // 1. Throw inside LongActionCallbacks still clears isAbortingDocumentLoad, runs notifications, and flushes new-generation commands.
    let notifRan = 0;
    api.LongActionCallbacks.push(() => { throw new Error('long-cb-boom'); });
    api.LongActionCallbacksParams.push(null);
    assert.throws(() => {
        api._runAbortScope(function() {
            this.documentLoadGeneration++;
            this._activeAbortNotifications.push(() => {
                notifRan++;
                mgr.callMethod('p1', 'Ping', []);
            });
        });
    }, /long-cb-boom/);
    assert.equal(api.isAbortingDocumentLoad, false);
    assert.equal(api.isPluginQueueSuspended(), false);
    assert.equal(notifRan, 1);
    assert.equal(mgr.queueCommands.length, 0);
    assert.deepEqual(sentMessages.map(m => m.methodReturnData), ['pong']);

    // 2. When fn throws, _runAbortScope does not mask the original error and does not flush queued commands.
    sentMessages.length = 0;
    assert.throws(() => {
        api._runAbortScope(function() {
            this.documentLoadGeneration++;
            mgr.callMethod('p1', 'Ping', []);
            throw new Error('fn-boom');
        });
    }, /fn-boom/);
    assert.equal(api.isAbortingDocumentLoad, false);
    assert.equal(mgr.queueCommands.length, 1);
    assert.deepEqual(sentMessages, []);
    mgr.flushQueue();
    assert.equal(mgr.queueCommands.length, 0);
    assert.deepEqual(sentMessages.map(m => m.methodReturnData), ['pong']);

    // 3. Finishing refresh pause uses flushQueue and does not decrement an external suspendQueue() hold.
    sentMessages.length = 0;
    mgr.suspendQueue();
    api._pausePluginQueueForRefresh();
    mgr.callMethod('p1', 'Ping', []);
    api._finishPluginQueuePause(true);
    assert.equal(mgr.queueSuspendCount, 1);
    assert.equal(mgr.queueCommands.length, 1);
    assert.deepEqual(sentMessages, []);
    mgr.resumeQueue();
    assert.equal(mgr.queueSuspendCount, 0);
    assert.equal(mgr.queueCommands.length, 0);
    assert.deepEqual(sentMessages.map(m => m.methodReturnData), ['pong']);
});

test('re-entrant callbacks clear one-shot slots before invoking across Cell, Word, Slide, PDF, ThemeLoader, and onDocumentOpen', () => {
    // 1. Cell asyncFontsDocumentEndLoaded -> onDocumentContentReady -> flushQueue -> new asyncMethodCallback.
    {
        let secondCbCalls = 0;
        const cellApi = {
            asyncMethodCallback: null,
            FontLoadWaitComplete: false,
            endFontLoadAction: noop,
            _openDocumentEndCallback: () => { throw new Error('must not enter openDocumentEndCallback for second callback'); }
        };
        cellApi.asyncFontsDocumentEndLoaded = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', 'asyncFontsDocumentEndLoaded', {});
        cellApi.asyncMethodCallback = function() {
            this.asyncMethodCallback = () => { secondCbCalls++; };
        };
        cellApi.asyncFontsDocumentEndLoaded();
        assert.ok(cellApi.asyncMethodCallback, 're-entrant asyncMethodCallback is preserved');
        cellApi.asyncFontsDocumentEndLoaded();
        assert.equal(secondCbCalls, 1);
        assert.equal(cellApi.asyncMethodCallback, undefined);
    }

    // 2. Word, Slide, and PDF asyncImagesDocumentEndLoaded + Word/Slide fast pre_Paste preserve re-entrant pasteCallback.
    for (const [file, prefix] of [
        ['word/api.js', 'asc_docs_api.prototype.'],
        ['slide/api.js', 'asc_docs_api.prototype.'],
        ['pdf/api.js', 'PDFEditorApi.prototype.']
    ]) {
        let secondPaste = 0;
        let decCalls = 0;
        const api = {
            isPasteFonts_Images: true,
            isSaveFonts_Images: false,
            isLoadImagesCustom: false,
            EndActionLoadImages: 0,
            pasteImageMap: {a: 1},
            pasteCancelCallback: noop,
            pasteCallback: null,
            ImageLoader: {bIsLoadDocumentFirst: false},
            endImageLoadAction: noop,
            decrementCounterLongAction: () => { decCalls++; },
            WordControl: {m_oDrawingDocument: {m_oDocumentRenderer: null}}
        };
        api.asyncImagesDocumentEndLoaded = loadMethod(file, prefix, 'asyncImagesDocumentEndLoaded', {
            c_oAscAsyncActionType: {Information: 1, BlockInteraction: 2},
            c_oAscAsyncAction: {LoadImage: 1, LoadDocumentImages: 2}
        });
        api.pasteCallback = function() {
            this.isPasteFonts_Images = true;
            this.pasteCallback = () => { secondPaste++; };
        };
        api.asyncImagesDocumentEndLoaded();
        assert.equal(typeof api.pasteCallback, 'function', `${file}: re-entrant pasteCallback preserved`);
        api.asyncImagesDocumentEndLoaded();
        assert.equal(secondPaste, 1, `${file}: second pasteCallback ran`);
        assert.equal(decCalls, 2, `${file}: CounterLongAction decremented for both pastes`);
    }

    // 3. Slide ThemeLoader.asyncImagesEndLoaded & EndLoadTheme preserve re-entrant theme load state.
    {
        let secondThemeCb = 0;
        let nextThemeStarted = false;
        const slideApi = {
            bSelectedSlidesTheme: false,
            _pluginApplyThemeCallback: null,
            WordControl: {
                m_oLogicDocument: {GetSelectedSlides: () => [], changeTheme: noop, FinalizeAction: noop},
                ThemeGenerateThumbnails: noop,
                CheckLayouts: noop
            },
            sync_EndAction: noop
        };
        slideApi.EndLoadTheme = loadMethod('slide/api.js', 'asc_docs_api.prototype.', 'EndLoadTheme', {
            AscCommon: {CollaborativeEditing: {Set_GlobalLock: noop}},
            c_oAscAsyncActionType: {BlockInteraction: 1},
            c_oAscAsyncAction: {LoadTheme: 2}
        });
        const themeLoader = {
            CurrentLoadThemeIndex: 0,
            isLoading: true,
            themes_info_editor: [{Master: 1}, {Master: 2}],
            Api: {
                ImageLoader: {ThemeLoader: null},
                EndLoadTheme(info) {
                    slideApi.EndLoadTheme(info);
                }
            }
        };
        themeLoader.asyncImagesEndLoaded = loadMethod('slide/Drawing/ThemeLoader.js', 'this.', 'asyncImagesEndLoaded', {});
        slideApi._pluginApplyThemeCallback = function() {
            if (!nextThemeStarted) {
                nextThemeStarted = true;
                themeLoader.CurrentLoadThemeIndex = 1;
                this._pluginApplyThemeCallback = () => { secondThemeCb++; };
            }
        };
        themeLoader.asyncImagesEndLoaded();
        assert.equal(themeLoader.CurrentLoadThemeIndex, 1);
        assert.equal(typeof slideApi._pluginApplyThemeCallback, 'function');
        themeLoader.asyncImagesEndLoaded();
        assert.equal(secondThemeCb, 1);
        assert.equal(themeLoader.CurrentLoadThemeIndex, -1);
    }

    // 4. apiBase.js CoAuthoringApi.onDocumentOpen preserves re-entrant fCurCallback (e.g. chained GetFileToDownload).
    {
        const {api} = createOpenCommandHarness();
        let secondDownload = null;
        api.fCurCallback = () => {
            api.fCurCallback = data => { secondDownload = data; };
        };
        api.CoAuthoringApi.onDocumentOpen({data: {type: 'download', url: 'first'}});
        assert.equal(typeof api.fCurCallback, 'function');
        api.CoAuthoringApi.onDocumentOpen({data: {type: 'download', url: 'second'}});
        assert.deepEqual(secondDownload, {type: 'download', url: 'second'});
        assert.equal(api.fCurCallback, null);
    }
});

test('onExpiredToken, direct updateversion refresh, empty Cell workbook, and Apply_Changes start action stay balanced', () => {
    // 1. CoAuthoringApi.onExpiredToken clears Open, Disconnect, RefreshFile, wopiRefreshRequestId, and plugin pause.
    {
        const mockWindow = {};
        const {api, events} = createOpenCommandHarness({window: mockWindow});
        const {mgr, sentMessages} = loadRealPluginsManager(api);
        mockWindow.g_asc_plugins = mgr;
        api.isLogPluginCommands = () => false;
        api.executeGroupActions = fn => fn();
        api.pluginMethod_Ping = () => 'ok';
        mgr.runnedPluginsMap['p1'] = {frameId: 'f1', isConnector: false, methodReturnAsync: false};

        api.isDocumentLoadComplete = true;
        api.isOpenActionActive = false;
        api.isDisconnectActionActive = true;
        api.isRefreshInProgress = true;
        api.wopiRefreshRequestId = 3;
        api.ServerIdWaitAction = [2, 11, 1];
        api._pausePluginQueueForRefresh();
        mgr.callMethod('p1', 'Ping', []);
        assert.equal(mgr.queueCommands.length, 1);

        api.CoAuthoringApi.onExpiredToken({code: 1});
        assert.equal(api.isRefreshInProgress, false);
        assert.equal(api.ServerIdWaitAction, null);
        assert.equal(api.isDisconnectActionActive, false);
        assert.equal(api.wopiRefreshRequestId, 4);
        assert.equal(api.isPluginQueuePausedForRefresh, false);
        assert.equal(events.filter(e => e[0] === 'asc_onEndAction' && e[2] === 12).length, 0, 'does not double-end closed Open');
        assert.equal(events.filter(e => e[0] === 'asc_onEndAction' && e[2] === 11).length, 1);
        assert.equal(mgr.queueCommands.length, 0, 'resumes plugin queue on live document');
        assert.deepEqual(sentMessages.map(m => m.methodReturnData), ['ok']);
        assert.equal(events.at(-1)[0], 'asc_onExpiredToken');

        // Incomplete load cancels queued plugin commands.
        sentMessages.length = 0;
        api.isDocumentLoadComplete = false;
        api.isOpenActionActive = true;
        api.isRefreshInProgress = true;
        api.ServerIdWaitAction = [2, 11, 1];
        api._pausePluginQueueForRefresh();
        mgr.callMethod('p1', 'Ping', []);
        api.CoAuthoringApi.onExpiredToken({code: 1});
        assert.equal(api.isOpenActionActive, false);
        assert.equal(api.isRefreshInProgress, false);
        assert.equal(api.ServerIdWaitAction, null);
        assert.equal(mgr.queueCommands.length, 0);
        assert.deepEqual(sentMessages.map(m => m.methodReturnData), [false]);
    }

    // 2. Direct asc_refreshFile without canRefreshFile() completes ServerIdWaitComplete after updateversion confirmation
    // even when DocsCoApi.disconnect() without opt_code only enters ClosedCoAuth and does not fire onDisconnect.
    {
        let openedUrl = null;
        let openCallbackCalls = 0;
        const {api, events} = createOpenCommandHarness();
        api.canRefreshFile = () => false;
        api.getViewMode = () => false;
        api.isCoAuthoringEnable = true;
        api.asc_coAuthoringDisconnect = function() {
            this.isCoAuthoringEnable = false;
            this.CoAuthoringApi.get_state = () => 3; // ClosedCoAuth
        };
        api._onOpenCommand = url => { openedUrl = url; };
        api._openDocumentEndCallback = () => { openCallbackCalls++; };

        api.isDocumentLoadComplete = false;
        api.ServerIdWaitComplete = false;
        api.isRefreshInProgress = true;
        api.ServerIdWaitAction = [2, 11, 1];

        api.CoAuthoringApi.onDocumentOpen({data: {type: 'open', status: 'updateversion', openedAt: 1, data: {'Editor.bin': '/doc.bin'}}});
        const updateEvt = events.find(e => e[0] === 'asc_onDocumentUpdateVersion');
        assert.ok(updateEvt && typeof updateEvt[1] === 'function');
        updateEvt[1]();
        assert.equal(api.isRefreshInProgress, false);
        assert.equal(api.ServerIdWaitAction, null);
        assert.equal(api.ServerIdWaitComplete, true);
        assert.equal(openedUrl, '/doc.bin');
        assert.equal(openCallbackCalls, 1);
    }

    // 3. Empty Cell workbook ends Open action and cancels paused plugin commands before reporting ConvertationOpenError.
    {
        let openEnded = 0;
        let disconnectEnded = 0;
        let pauseFinished = [];
        const sentEvents = [];
        const cellApi = {
            isDocumentLoadComplete: false,
            ServerIdWaitComplete: true,
            FontLoadWaitComplete: true,
            wbModel: {getWorksheetCount: () => 0},
            endOpenAction: () => { openEnded++; },
            endDisconnectAction: () => { disconnectEnded++; },
            _finishPluginQueuePause: canResume => { pauseFinished.push(canResume); },
            sendEvent: (name, ...args) => { sentEvents.push([name, ...args]); }
        };
        cellApi._openDocumentEndCallback = loadMethod('cell/api.js', 'spreadsheet_api.prototype.', '_openDocumentEndCallback', {
            AscCommon: {EncryptionWorker: null},
            c_oAscError: {ID: {ConvertationOpenError: -4}, Level: {Critical: 1}}
        });
        cellApi._openDocumentEndCallback();
        assert.equal(openEnded, 1);
        assert.equal(disconnectEnded, 1);
        assert.deepEqual(pauseFinished, [false]);
        assert.deepEqual(sentEvents, [['asc_onError', -4, 1]]);
    }

    // 4. CollaborativeEditingBase.Apply_Changes stops if startApplyChangesAction synchronously bumps documentLoadGeneration.
    {
        let savedState = 0;
        const AscFonts = {IsCheckSymbols: false};
        const api = {
            documentLoadGeneration: 0,
            WordControl: {m_oLogicDocument: {PauseRecalculate: noop, EndPreview_MailMergeResult: noop}},
            sendEvent: noop,
            startApplyChangesAction() {
                this.documentLoadGeneration++;
            }
        };
        const collab = {
            m_nUndoBeforeApply: 0,
            m_aChanges: [{id: 1}],
            m_bRecalculatePausedByApplyChanges: false,
            GetEditorApi: () => api,
            private_SaveDocumentState: () => { savedState++; },
            Clear_NewImages: noop,
            Apply_OtherChanges: noop,
            Lock_NeedLock: noop,
            private_RestoreDocumentState: noop,
            OnStart_Load_Objects: noop
        };
        const applyChanges = loadMethod('common/CollaborativeEditingBase.js', 'CCollaborativeEditingBase.prototype.', 'Apply_Changes', {
            window: {},
            AscFonts,
            AscCommon: {CollaborativeEditing: {CoHistory: {InitTextRecover: noop}}},
            Asc: {c_oAscAsyncActionType: {BlockInteraction: 1}, c_oAscAsyncAction: {ApplyChanges: 2}}
        });
        applyChanges.call(collab, noop);
        assert.equal(savedState, 0);
        assert.equal(AscFonts.IsCheckSymbols, false);
    }
});
