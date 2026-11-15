'use strict';
const assert = require('node:assert/strict'), {test} = require('node:test'), vm = require('node:vm');
const {source, method} = require('./source-harness.cjs');
function world(desktop = false) {
    const calls = [], listeners = new Set(), popups = [], timers = new Map();
    let timerId=0;
    const origin = 'https://editor.test';
    const opener = {closed: false, postMessage: (...args) => calls.push(['opener', ...args])};
    const window = {location: {origin, href: origin + '/index.html'}, innerWidth: 1200, innerHeight: 800, screenLeft: 0, screenTop: 0,
        opener, postMessage: (...args) => calls.push(['self', ...args]),
        addEventListener(type, fn) {assert.equal(type, 'message'); listeners.add(fn);},
        removeEventListener(type, fn) {listeners.delete(fn);}, close() {calls.push(['selfClose']);},
        open(url) {calls.push(['open', url, listeners.size]); if (window.blockPopup) return null;
            const popup = {closed: false, postMessage: (...args) => calls.push(['popup', ...args]),
                addEventListener() {}, removeEventListener() {}, close() {this.closed = true; calls.push(['close']); if (this.onbeforeunload) this.onbeforeunload();}};
            popups.push(popup); return popup;}}
    if (desktop) window.AscDesktopEditor = {startReporter() {}, endReporter() {}, sendToReporter: v => calls.push(['desktopTo',v]), sendFromReporter: v => calls.push(['desktopFrom',v])};
    const common = {translateManager: {mapTranslate: {}}, GlobalSkin: {}, AscBrowser: {}, EncryptionWorker: {isPasswordCryptoPresent: false}};
    const api = {locale: 'ru', isReporterMode: false, sendEvent: (...a) => calls.push(['event', ...a]), getAnnotations: () => null};
    const manager = {Mode: false, waitReporterObject: ['preview',0,true],
        EndWaitReporter(cancel) {calls.push(['wait', cancel]); this.waitReporterObject = null;},
        OnNextSlide() {calls.push(['next']);}, OnPrevSlide() {calls.push(['prev']);}, Pause() {calls.push(['pause']);},
        PointerMove(...args) {calls.push(['pointer', ...args]);},
        End() {if(this.waitReporterObject) this.EndWaitReporter(true); api.DemonstrationReporterEnd(); calls.push(['end']);}};
    api.WordControl = {DemonstrationManager: manager}; window.editor = api;
    const ctx = {window, document: {documentElement: {}}, screen: {}, setInterval: fn => {timers.set(++timerId,fn);return timerId;}, clearInterval: id => timers.delete(id), AscCommon: common, Asc: {editor: api}};
    for (const name of ['DemonstrationReporterStart','DemonstrationReporterEnd','DemonstrationReporterMessages','DemonstrationToReporterMessages','sendFromReporter','sendToReporter','EndDemonstration','closeDemonstration','endDemoMode'])
        api[name] = vm.runInNewContext('(' + method(source('slide/api.js'), name) + ')', ctx);
    function start() {api.DemonstrationReporterStart({key:'test', token:'secret-token', translations: {endSlideshow: 'Завершить показ & <закрыть>'}}); return api.reporterWindow;}
    function message(data, from = api.reporterWindow, fromOrigin = origin) {
        const e = {source: from, origin: fromOrigin, data: typeof data === 'string' ? data : JSON.stringify(data)};
        for (const fn of [...listeners]) fn(e);
    }
    return {api, window, opener, calls, listeners, popups, timers, manager, common, start, message, origin};
}

test('browser sends raw messages once only to its trusted main window', () => {
    const w = world(); w.api.isReporterMode = true; w.api.sendFromReporter('next');
    assert.deepEqual(w.calls, [['opener', 'next', w.origin]]);
    w.calls.length = 0; w.api.disableReporterEvents = true;
    Object.defineProperty(w.window, 'opener', {get() {throw new Error('must not access opener');}});
    w.api.sendFromReporter('next'); assert.equal(w.calls.length, 0);
});

test('ordinary editor never sends presenter messages to its portal opener', () => {
    const w = world(); w.api.sendFromReporter('next'); assert.equal(w.calls.length, 0);
});

test('desktop transport remains raw and never accesses opener', () => {
    const w = world(true); Object.defineProperty(w.window, 'opener', {get() {throw new Error('desktop opener');}});
    w.api.sendFromReporter('raw'); w.api.sendToReporter('raw');
    assert.deepEqual(w.calls, [['desktopFrom','raw'], ['desktopTo','raw']]);
    w.api.disableReporterEvents = true; w.api.sendFromReporter('ignored'); assert.equal(w.calls.length, 2);
});

test('main subscribes before opening and sends metadata once from the editor', () => {
    const w = world(), popup = w.start();
    assert.equal(w.calls[0][2], 1);
    const url = new URL(w.calls[0][1], w.origin);
    assert.equal(url.searchParams.get('reporterEndLabel'), 'Завершить показ & <закрыть>'); assert.ok(!w.calls[0][1].includes('secret-token'));
    let reads = 0; w.common.EncryptionWorker.isPasswordCryptoPresent = true;
    for (const key of ['currentPassword','currentDocumentHash','currentDocumentInfo']) Object.defineProperty(w.api, key, {get() {reads++; return key;}});
    w.message('i:am:ready', {}, w.origin); w.message('i:am:ready', popup, 'https://portal.test'); assert.equal(reads, 0);
    w.message('i:am:ready'); w.message('i:am:ready');
    const sent = w.calls.filter(c => c[0] === 'popup'); assert.equal(sent.length, 1); assert.equal(sent[0][2], w.origin);
    assert.equal(JSON.parse(sent[0][1]).data.cryptoCurrentPassword, 'currentPassword'); assert.equal(reads, 3);
});

test('show starts once after metadata and commands execute once', () => {
    const w = world(); w.start(); w.message({reporter_command:'start_show'}); assert.equal(w.calls.filter(c => c[0] === 'wait').length, 0);
    w.message('i:am:ready'); w.message({reporter_command:'start_show'}); w.message({reporter_command:'start_show'});
    assert.deepEqual(w.calls.filter(c => c[0] === 'wait'), [['wait', undefined]]);
    for (const command of ['next','prev','pause','pointer_move']) w.message({reporter_command:command,x:1,y:2,w:3,h:4});
    assert.equal(w.calls.filter(c => c[0] === 'next').length, 1); assert.equal(w.calls.filter(c => c[0] === 'prev').length, 1);
    assert.equal(w.calls.filter(c => c[0] === 'pause').length, 1); assert.equal(w.calls.filter(c => c[0] === 'pointer').length, 1);
});

test('three shows clean listeners and ignore commands from older windows', () => {
    const w = world(); let old;
    for (let i = 0; i < 3; i++) {
        const popup = w.start(); assert.equal(w.listeners.size, 1);
        if (old) {w.message({reporter_command:'end'}, old); w.message({reporter_command:'start_show'}, old); assert.equal(w.api.reporterWindow, popup);}
        w.message('i:am:ready'); w.message({reporter_command:'start_show'});
        w.api.DemonstrationReporterEnd(); w.api.DemonstrationReporterEnd(); assert.equal(w.listeners.size, 0);
        assert.equal(w.api.reporterWindow, null); assert.equal(w.api.reporterStartObject, null); old = popup;
    }
});

test('early reporter end sends one cancel even without a demonstration manager', () => {
    for (const missing of [false,true]) {
        const w = world(); w.api.isReporterMode = true; if (missing) delete w.api.WordControl;
        w.api.EndDemonstration(); w.api.EndDemonstration();
        assert.equal(w.calls.filter(c => c[0] === 'opener').length, 1);
        assert.equal(JSON.parse(w.calls.find(c => c[0] === 'opener')[1]).reporter_command, 'end');
        assert.ok(w.calls.some(c => c[0] === 'selfClose')); assert.ok(!w.calls.some(c => c[0] === 'event'));
    }
});

test('main cancels waiting before metadata without starting a show', () => {
    const w = world(); w.start(); w.message({reporter_command:'end'});
    assert.deepEqual(w.calls.filter(c => c[0] === 'wait'), [['wait', true]]); assert.equal(w.listeners.size, 0);
});

test('window close cancels waiting and its stale handler cannot end a new show', () => {
    const w = world(), popup = w.start(), close = popup.onbeforeunload;
    close(); assert.equal(w.api.reporterWindow, null); assert.equal(w.listeners.size, 0);
    const fresh = w.start(); close(); assert.equal(w.api.reporterWindow, fresh);
});

test('blocked popup removes state and subscriptions', () => {
    const w = world(); w.window.blockPopup = true; w.start();
    assert.equal(w.listeners.size, 0); assert.equal(w.api.reporterWindow, null); assert.equal(w.api.reporterStartObject, null);
});

test('reporter accepts main commands only from its opener at the editor origin', () => {
    const w = world(); w.api.isReporterMode = true;
    for (const [source, origin] of [[{}, w.origin], [w.opener, 'https://portal.test'], [w.opener, w.origin]])
        w.api.DemonstrationToReporterMessages({source, origin, data: JSON.stringify({main_command:true,next:true})});
    assert.equal(w.calls.filter(c => c[0] === 'next').length, 1);
});

test('initial popup navigation does not cancel a pending show', () => {
    const w = world(), popup = w.start();
    if (popup.onunload) popup.onunload();
    assert.equal(w.api.reporterWindow,popup);
    assert.deepEqual(w.calls.filter(c => c[0] === 'wait'),[]);
});


test('closing a popup before its scripts load cancels waiting and clears the close monitor', () => {
    const w=world(), popup=w.start();
    assert.equal(w.timers.size,1);
    popup.closed=true;
    for(const check of [...w.timers.values()]) check();
    assert.deepEqual(w.calls.filter(c=>c[0]==='wait'),[['wait',true]]);
    assert.equal(w.timers.size,0); assert.equal(w.listeners.size,0);
    w.start(); w.api.DemonstrationReporterEnd(); assert.equal(w.timers.size,0);
});

test('close notifications request one annotation decision while completion is pending', () => {
    const w=world(), popup=w.start();
    w.api.getAnnotations=()=>({isEmpty:()=>false});
    popup.onbeforeunload(); popup.closed=true;
    for(const check of [...w.timers.values()]) check();
    w.message({reporter_command:'end'}); w.message({reporter_command:'end'});
    assert.equal(w.calls.filter(c=>c[0]==='event'&&c[1]==='asc_onEndDemoWithAnnotations').length,1);
    assert.equal(w.timers.size,0);
});

test('an end command clears the close monitor before an annotation decision completes', () => {
    const w=world(); w.start(); w.api.getAnnotations=()=>({isEmpty:()=>false});
    w.message({reporter_command:'end'});
    assert.equal(w.timers.size,0);
    w.message({reporter_command:'end'});
    assert.equal(w.calls.filter(c=>c[0]==='event'&&c[1]==='asc_onEndDemoWithAnnotations').length,1);
});
