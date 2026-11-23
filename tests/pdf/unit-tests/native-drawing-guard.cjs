/* Run: node tests/pdf/unit-tests/native-drawing-guard.cjs
 * Executes the real guardNativeDrawing from pdf/src/file.js against a drawing engine double.
 * Does not load drawingfile.wasm or the viewer.
 */
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '../../..');
const fileSource = fs.readFileSync(process.env.PDF_FILE_SOURCE || path.join(root, 'pdf/src/file.js'), 'utf8');

function extractFunction(source, name) {
	const start = source.indexOf('function ' + name + '(');
	assert.ok(start >= 0, 'real function exists: ' + name);
	const open = source.indexOf('{', start);
	let depth = 0;
	for (let i = open; i < source.length; i++) {
		if (source[i] === '{') depth++;
		else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
	}
	throw new Error('unbalanced function: ' + name);
}

const guardNativeDrawing = vm.runInThisContext('(' + extractFunction(fileSource, 'guardNativeDrawing') + ')');

const METHODS = ['getPagePixmap', 'getAnnotationsAP', 'getInteractiveFormsAP', 'getButtonIcons', 'getGlyphs'];

function createEngine(behavior) {
	const engine = { calls: {}, unlockCalls: 0 };
	for (const name of METHODS) {
		engine.calls[name] = [];
		engine[name] = function () {
			engine.calls[name].push({ self: this, args: Array.from(arguments) });
			return behavior(name, Array.from(arguments));
		};
	}
	engine.unlockPageNumForFontsLoader = function () {
		engine.unlockCalls++;
	};
	return engine;
}

function createReporter() {
	const reports = [];
	const report = function (failure) {
		reports.push(failure);
	};
	report.reports = reports;
	return report;
}

const engineError = new WebAssembly.RuntimeError('memory access out of bounds');
const throwAlways = function () {
	throw engineError;
};

function createPageImageApi(engine, pagesCount = 1, options = {}) {
	engine.pages = Array.from({ length: pagesCount }, (_, originIndex) => ({ W: 24, H: 32, originIndex }));
	engine.free = () => {};
	engine.loadFromData = () => 0;
	engine.getPages = () => engine.pages;
	engine.getType = () => 0;
	const window = {
		AscViewer: { CDrawingFile: function () { return engine; } },
		AscPDF: { APPEARANCE_TYPES: {} },
		PDFEditorApi: function () {}
	};
	const canvas = { toDataURL: () => 'data:image/png;base64,background' };
	canvas.getContext = () => ({ canvas, fillRect() {} });
	const context = vm.createContext({
		window,
		AscPDF: window.AscPDF,
		AscCommon: {
			AscBrowser: {}, sendClientLog() {},
			translateManager: { getValue: (text) => text }
		},
		AscCommonWord: { g_dKoef_pt_to_mm: 1, g_dKoef_mm_to_pix: 1 },
		editor: { getDocumentRenderer: () => api.DocumentRenderer },
		Asc: { editor: { IsEditFieldsMode: () => false }, c_oAscError: { Level: { NoCritical: 0 } } },
		navigator: { userAgent: "" },
		RangeError,
		WebAssembly,
		performance,
		document: { createElement: () => canvas }
	});
	vm.runInContext(fileSource, context);
	vm.runInContext(fs.readFileSync(path.join(root, 'pdf/api_plugins.js'), 'utf8'), context);
	const api = new window.PDFEditorApi();
	Object.assign(api, options);
	window.Asc = { editor: api };
	const file = window.AscViewer.createFile();
	const rendered = { toDataURL: () => 'data:image/png;base64,content' };
	file._pixelsToCanvas = () => ({ getContext: () => ({ canvas: rendered }) });
	api.DocumentRenderer = { file, _drawMarkupAnnotsOnCtx() {} };
	api.getPageBackgroundColor = () => ({ R: 255, G: 255, B: 255 });
	const graphics = {
		GetDrawingPageW: () => 24,
		GetDrawingPageH: () => 32,
		SetGlobalAlpha() {},
		DrawImageXY() {}
	};
	function loadMethods(filename, type, target, names) {
		const source = fs.readFileSync(path.join(root, filename), 'utf8');
		for (const name of names) {
			const prefix = type + '.prototype.' + name + ' = function(';
			const [, body] = source.split(prefix);
			target[name] = vm.runInContext('(' + extractFunction('function method(' + body, 'method') + ')', context);
		}
	}
	return { api, file, graphics, loadMethods };
}

const tests = {
	'page image plugin distinguishes failed rendering from empty and recognized pages'() {
		const engine = createEngine((name, args) => {
			if (args[0] === 0) throw engineError;
			return 42;
		});
		const { api, file } = createPageImageApi(engine, 2);

		assert.equal(api.pluginMethod_GetPageImage(0), null);
		assert.equal(api.pluginMethod_GetPageImage(0), null);
		assert.equal(engine.calls.getPagePixmap.length, 1);
		assert.equal(api.pluginMethod_GetPageImage(1), 'data:image/png;base64,content');
		file.pages[0].isRecognized = true;
		assert.equal(api.pluginMethod_GetPageImage(0), 'data:image/png;base64,background');
		file.pages.unshift({ W: 24, H: 32 });
		assert.equal(api.pluginMethod_GetPageImage(0), 'data:image/png;base64,background');
	},

	'uses the native capped size to detect a failed page export'() {
		const engine = createEngine((name, args) => {
			if (args[2] === 16) throw engineError;
			return null;
		});
		const { api, file } = createPageImageApi(engine);
		file.maxCanvasSize = 16;
		assert.equal(api.pluginMethod_GetPageImage(0), null);
		assert.deepEqual(engine.calls.getPagePixmap[0].args.slice(1, 3), [12, 16]);
		assert.equal(api.pluginMethod_GetPageImage(0, { maxSize: 8 }), 'data:image/png;base64,background');
	},

	'warns separately when retries stop after the PDF becomes ready'() {
		const events = [];
		const engine = createEngine(throwAlways);
		const { api, file } = createPageImageApi(engine, 1, { sendEvent: (...args) => events.push(args) });
		assert.equal(api.pluginMethod_GetPageImage(0), null);
		file.showDrawingWarning(api);
		assert.deepEqual(events, []);
		api.isDocumentLoadComplete = true;
		file.showDrawingWarning(api);
		assert.deepEqual(events, [['asc_onError', 'Some PDF content could not be displayed.', 0]]);
		api.pluginMethod_GetPageImage(0);
		assert.equal(events.length, 1);
		api.pluginMethod_GetPageImage(0, { maxSize: 8 });
		api.pluginMethod_GetPageImage(0, { maxSize: 9 });
		assert.deepEqual(events, [
			['asc_onError', 'Some PDF content could not be displayed.', 0],
			['asc_onError', 'The editor stopped trying to display some PDF content after repeated errors.', 0]
		]);
	},

	'page image export fails only for a requested annotation with an unread appearance'() {
		const engine = createEngine((name, args) => {
			if (name === 'getAnnotationsAP' && args[4] === 7) throw engineError;
			return name === 'getAnnotationsAP' ? [] : null;
		});
		const { api, graphics, loadMethods } = createPageImageApi(engine);
		const base = 'pdf/src/annotations/base.js';
		let hidden = false;
		const stamp = {
			GetApIdx: () => 7, GetOriginPage: () => 0, IsStamp: () => true, IsHidden: () => hidden,
			GetOriginViewScale: () => 1, GetRotate: () => 0
		};
		const annotation = { GetApIdx: () => 8, GetOriginPage: () => 0, IsStamp: () => false, IsHidden: () => false };
		loadMethods(base, 'CAnnotationBase', annotation, ['DrawFromStream', 'GetOriginView', 'GetOriginViewInfo']);
		loadMethods(base, 'CAnnotationBase', stamp, ['GetOriginView']);
		loadMethods('pdf/src/annotations/stamp.js', 'CAnnotationStamp', stamp, ['DrawFromStream', 'GetOriginViewInfo']);
		let annotations = [annotation, stamp];
		api.DocumentRenderer._drawAnnotsOnCtx = () => annotations.forEach((annot) => annot.DrawFromStream(graphics));

		stamp.DrawFromStream(graphics);
		assert.equal(api.pluginMethod_GetPageImage(0), 'data:image/png;base64,background');
		assert.equal(api.pluginMethod_GetPageImage(0, { annotations: true }), null);
		hidden = true;
		assert.equal(api.pluginMethod_GetPageImage(0, { annotations: true }), 'data:image/png;base64,background');
		hidden = false;
		annotations = [annotation];
		assert.equal(api.pluginMethod_GetPageImage(0, { annotations: true }), 'data:image/png;base64,background');
		assert.equal(engine.calls.getAnnotationsAP.length, 2);
	},

	'page image export uses a cached field appearance after a failed native read'() {
		const engine = createEngine((name) => {
			if (name === 'getInteractiveFormsAP') throw engineError;
			return null;
		});
		const { api, graphics, loadMethods } = createPageImageApi(engine);
		let hidden = false;
		const field = {
			_apIdx: 7, _originView: {}, GetOriginPage: () => 0, IsHidden: () => hidden,
			IsNeedDrawFromStream: () => true, GetRect: () => [0, 0, 1, 1], DrawLocks() {}, DrawEdit() {}
		};
		loadMethods('pdf/src/forms/base/base.js', 'CBaseField', field,
			['DrawOnPage', 'DrawFromStream', 'GetOriginView', 'GetOriginViewInfo']);
		api.DocumentRenderer._drawFieldsOnCtx = () => field.DrawOnPage(graphics, {}, 0);

		assert.equal(api.pluginMethod_GetPageImage(0, { fields: true }), null);
		hidden = true;
		assert.equal(api.pluginMethod_GetPageImage(0, { fields: true }), 'data:image/png;base64,background');
		hidden = false;
		const cachedAppearance = { width: 24, height: 32 };
		field._originView.normal = cachedAppearance;
		let drawnAppearance;
		graphics.DrawImageXY = (appearance) => { drawnAppearance = appearance; };
		assert.equal(api.pluginMethod_GetPageImage(0, { fields: true }), 'data:image/png;base64,background');
		assert.equal(drawnAppearance, cachedAppearance);
		assert.equal(engine.calls.getInteractiveFormsAP.length, 1);
	},

	'retries an appearance at the same size after a RangeError'() {
		const engine = createEngine((name) => {
			if (name !== 'getInteractiveFormsAP') return null;
			if (engine.calls.getInteractiveFormsAP.length === 1) throw new RangeError('Maximum call stack size exceeded');
			return [];
		});
		const { api, graphics, loadMethods } = createPageImageApi(engine);
		const field = {
			_apIdx: 7, _originView: {}, GetOriginPage: () => 0, IsHidden: () => false,
			IsNeedDrawFromStream: () => true, GetRect: () => [0, 0, 1, 1], DrawLocks() {}, DrawEdit() {}
		};
		loadMethods('pdf/src/forms/base/base.js', 'CBaseField', field,
			['DrawOnPage', 'DrawFromStream', 'GetOriginView', 'GetOriginViewInfo']);
		api.DocumentRenderer._drawFieldsOnCtx = () => field.DrawOnPage(graphics, {}, 0);

		assert.equal(api.pluginMethod_GetPageImage(0, { fields: true }), null);
		assert.equal(api.pluginMethod_GetPageImage(0, { fields: true }), 'data:image/png;base64,background');
		assert.equal(engine.calls.getInteractiveFormsAP.length, 2);
	},

	'passes results and arguments through'() {
		const results = { getPagePixmap: 1234, getAnnotationsAP: [{ i: 1 }], getInteractiveFormsAP: [{ i: 2 }], getButtonIcons: { MK: [], View: [] }, getGlyphs: [{ line: 1 }] };
		const engine = createEngine((name) => results[name]);
		const report = createReporter();
		guardNativeDrawing(engine, report);
		for (const name of METHODS) {
			assert.equal(engine[name](2, 800, 1100, 0xFFFFFF), results[name], name);
			assert.deepEqual(engine.calls[name][0].args, [2, 800, 1100, 0xFFFFFF], name);
			assert.equal(engine.calls[name][0].self, engine, name);
		}
		assert.equal(report.reports.length, 0);
		assert.equal(engine.unlockCalls, 0);
	},

	'returns empty values on failure'() {
		const engine = createEngine(throwAlways);
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.equal(engine.getPagePixmap(2, 1200, 1700, 0xFFFFFF), null);
		assert.deepEqual(engine.getAnnotationsAP(2, 800, 1100), []);
		assert.deepEqual(engine.getInteractiveFormsAP(2, 800, 1100), []);
		assert.equal(engine.getButtonIcons(2, 800, 1100, undefined, true), null);
		assert.deepEqual(engine.getGlyphs(2), []);
		assert.equal(engine.unlockCalls, 5);
		assert.equal(report.reports.length, 5);
		assert.deepEqual(report.reports[0], {
			method: 'getPagePixmap',
			page: 2,
			width: 1200,
			height: 1700,
			objectId: undefined,
			appearance: undefined,
			error: 'RuntimeError: memory access out of bounds',
			engineCrash: true,
			stack: engineError.stack
		});
		assert.deepEqual(report.reports.map((r) => r.method), METHODS);
	},

	'retries a failed page at a different size within the page limit'() {
		const engine = createEngine((name, args) => {
			if (args[1] >= 1200) throw engineError;
			return 42;
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.equal(engine.getPagePixmap(2, 1200, 1700), null);
		assert.equal(engine.getPagePixmap(2, 1200, 1700), null);
		assert.equal(engine.getPagePixmap(2, 600, 850), 42);
		assert.equal(engine.getPagePixmap(2, 1300, 1800), null);
		assert.equal(engine.getPagePixmap(2, 1400, 1900), null);
		assert.equal(engine.getPagePixmap(3, 600, 850), 42);
		assert.equal(engine.calls.getPagePixmap.length, 4);
		assert.equal(report.reports.length, 2);
	},

	'keeps other pages and methods working'() {
		const engine = createEngine((name, args) => {
			if (name === 'getPagePixmap' && args[0] === 2) throw engineError;
			return name === 'getPagePixmap' ? 99 : [];
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.equal(engine.getPagePixmap(2, 1200, 1700), null);
		assert.equal(engine.getPagePixmap(3, 1200, 1700), 99);
		engine.getAnnotationsAP(2, 800, 1100);
		assert.equal(engine.calls.getPagePixmap.length, 2);
		assert.equal(engine.calls.getAnnotationsAP.length, 1);
		assert.equal(report.reports.length, 1);
	},

	'wraps only once'() {
		const engine = createEngine(() => 7);
		const report = createReporter();
		assert.equal(guardNativeDrawing(engine, report), engine);
		assert.equal(guardNativeDrawing(engine, report), engine);
		engine.getPagePixmap(0, 10, 10);
		assert.equal(engine.calls.getPagePixmap.length, 1);
		assert.equal(engine['isDrawingGuarded'], true);
	},

	'survives a failing reporter and unlock'() {
		const engine = createEngine(throwAlways);
		engine.unlockPageNumForFontsLoader = function () {
			throw new TypeError("Cannot set properties of null (setting 'fontPageUpdateType')");
		};
		guardNativeDrawing(engine, function () {
			throw new Error('report failed');
		});
		assert.equal(engine.getPagePixmap(1, 100, 100), null);
		assert.deepEqual(engine.getAnnotationsAP(1, 100, 100), []);
	},

	'returns empty after a non-Error throw and retries the same call'() {
		const engine = createEngine(() => {
			if (engine.calls.getPagePixmap.length === 1) throw 42;
			return 7;
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.equal(engine.getPagePixmap(0, 100, 200), null);
		assert.equal(engine.getPagePixmap(0, 100, 200), 7);
		assert.equal(engine.calls.getPagePixmap.length, 2);
		assert.equal(report.reports[0].error, '42');
		assert.equal(report.reports[0].stack, '');
	},

	'retries a page image at the same size after a RangeError'() {
		const engine = createEngine(() => {
			if (engine.calls.getPagePixmap.length === 1) throw new RangeError('Maximum call stack size exceeded');
			return 42;
		});
		const { api } = createPageImageApi(engine);
		assert.equal(api.pluginMethod_GetPageImage(0), null);
		assert.equal(api.pluginMethod_GetPageImage(0), 'data:image/png;base64,content');
		assert.equal(engine.calls.getPagePixmap.length, 2);
	},

	'source wiring'() {
		const created = fileSource.indexOf('file.nativeFile = new window["AscViewer"]["CDrawingFile"]();');
		assert.ok(created >= 0, 'engine creation line exists');
		const rest = fileSource.slice(created).split('\n').slice(1).map((l) => l.trim()).filter(Boolean);
		assert.equal(rest[0], 'guardNativeDrawing(file.nativeFile, reportNativeDrawingFailure, function(limitReached) {');
		assert.ok(fileSource.includes('nativeFile["isDrawingGuarded"]'), 'guard mark uses a quoted property');
		assert.ok(!fileSource.includes('nativeFile.isDrawingGuarded'), 'no unquoted guard mark');
		const guardStart = fileSource.indexOf('function guardNativeDrawing(');
		assert.ok(fileSource.indexOf('Error["stackTraceLimit"] = 50;') > guardStart, 'stack trace limit is raised when a PDF engine is guarded');
	},

	'extends the stack trace limit only while drawing'() {
		const saved = Error.stackTraceLimit;
		try {
			Error.stackTraceLimit = 10;
			const engine = createEngine(() => Error.stackTraceLimit);
			guardNativeDrawing(engine, createReporter());
			assert.equal(Error.stackTraceLimit, 10);
			assert.equal(engine.getPagePixmap(0, 100, 100), 50);
			assert.equal(Error.stackTraceLimit, 10);
			Error.stackTraceLimit = 100;
			assert.equal(engine.getPagePixmap(0, 100, 100), 100);
			assert.equal(Error.stackTraceLimit, 100);
		} finally {
			Error.stackTraceLimit = saved;
		}
	},

	'marks engine crashes'() {
		const engine = createEngine((name) => {
			if (name === 'getPagePixmap') throw new WebAssembly.RuntimeError('memory access out of bounds');
			throw new TypeError('bad page');
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		engine.getPagePixmap(0, 10, 10);
		assert.deepEqual(engine.getGlyphs(0), []);
		assert.equal(report.reports[0].engineCrash, true);
		assert.equal(report.reports[1].engineCrash, false);
	},

	'tracks single annotations separately'() {
		const engine = createEngine((name, args) => {
			if (args[4] === 7) throw engineError;
			return [];
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.deepEqual(engine.getAnnotationsAP(2, 800, 1100, undefined, 7), []);
		engine.getAnnotationsAP(2, 800, 1100, undefined, 7);
		engine.getAnnotationsAP(2, 800, 1100);
		engine.getAnnotationsAP(2, 800, 1100, undefined, 8);
		assert.equal(engine.calls.getAnnotationsAP.length, 3);
		assert.equal(report.reports.length, 1);
	},

	'failed icon reads never block a healthy page bitmap'() {
		const engine = createEngine((name, args) => {
			if (name === 'getButtonIcons' && args[0] < 5) throw engineError;
			return name === 'getPagePixmap' ? 42 : { View: [] };
		});
		const { api } = createPageImageApi(engine, 6);
		for (let i = 0; i < 5; i++)
			engine.getButtonIcons(i, 10, 10);
		assert.equal(api.pluginMethod_GetPageImage(5), 'data:image/png;base64,content');
		assert.deepEqual(engine.getButtonIcons(5, 10, 10), { View: [] });
		assert.equal(engine.calls.getPagePixmap.length, 1);
		assert.equal(engine.calls.getButtonIcons.length, 6);
	},

	'limits diagnostic volume while continuing to draw other pages'() {
		const engine = createEngine((name, args) => {
			if (args[0] < 12) throw engineError;
			return 42;
		});
		const report = createReporter();
		guardNativeDrawing(engine, report);
		for (let i = 0; i < 12; i++)
			assert.equal(engine.getPagePixmap(i, 10, 10), null);
		assert.equal(report.reports.length, 10);
		assert.equal(engine.getPagePixmap(12, 10, 10), 42);
		assert.equal(engine.calls.getPagePixmap.length, 13);
	},

	'formats the log record'() {
		const sent = [];
		const api = { getEditorErrorInfo: () => 'editor: pdf' };
		const context = {
			window: { Asc: { editor: api } },
			navigator: { userAgent: 'UA/1.0' },
			AscCommon: { getMemoryInfo: () => 'used 1', sendClientLog: (level, msg, target) => sent.push({ level, msg, target }) }
		};
		const reportFailure = vm.runInNewContext('(' + extractFunction(fileSource, 'reportNativeDrawingFailure') + ')', context);
		reportFailure({ method: 'getPagePixmap', page: 2, width: 1200, height: 1700, error: 'RuntimeError: memory access out of bounds', engineCrash: true, stack: 'at blend' });
		assert.equal(sent.length, 1);
		assert.equal(sent[0].level, 'error');
		assert.equal(sent[0].target, api);
		assert.equal(sent[0].msg,
			'pdfDrawingError: method=getPagePixmap page=2 size=1200x1700 engineCrash=true error=RuntimeError: memory access out of bounds' +
			'\n userAgent: UA/1.0\n editor: pdf\n performance.memory: used 1\n StackTrace: at blend');
	},

	'tracks the failed page by its origin index'() {
		const engine = createEngine((name, args) => {
			if (engine.pages[args[0]].originIndex === 5) throw engineError;
			return 1;
		});
		engine.pages = [0, 1, 2, 3, 4, 5, 6].map((i) => ({ originIndex: i }));
		const report = createReporter();
		guardNativeDrawing(engine, report);
		assert.equal(engine.getPagePixmap(5, 100, 100), null);
		engine.pages.splice(2, 1);
		assert.equal(engine.getPagePixmap(4, 100, 100), null);
		assert.equal(engine.calls.getPagePixmap.length, 1, 'shifted failed page is not drawn again');
		assert.equal(engine.getPagePixmap(5, 100, 100), 1, 'page shifted into the failed index is drawn');
		assert.equal(report.reports.length, 1);
		assert.equal(report.reports[0].page, 5);
	},

	'expires failures for undone imports whose origin indices are reused'() {
		const historySource = fs.readFileSync(path.join(root, 'pdf/src/history/documentChanges.js'), 'utf8');
		const [, undoSource] = historySource.match(/CChangesPDFDocumentEndMergePages\.prototype\.Undo = (function\(\) \{[\s\S]*?\n\});/);
		const undo = vm.runInThisContext('(' + undoSource + ')');
		const originalPages = [{ originIndex: 0 }, { originIndex: 1 }];
		const oldImport = { originIndex: 2, broken: true };
		const newImport = { originIndex: 2 };
		const engine = createEngine((name, args) => {
			const page = name === 'getPagePixmap'
				? engine.pages[args[0]]
				: engine.pages.find((page) => page.originIndex === args[0]);
			if (page.originIndex === 1 || page.broken) throw engineError;
			return name === 'getPagePixmap' ? 99 : [{ i: 7 }];
		});
		engine.pages = [oldImport, ...originalPages];
		let nativePages = [...originalPages, oldImport];
		engine.UndoMergePages = function () {
			nativePages.pop();
		};
		engine.getPagesInfo = () => nativePages;
		const file = { nativeFile: engine, originalPagesCount: 3 };
		const doc = { Viewer: { file }, mergedPagesData: [{}] };
		guardNativeDrawing(engine, createReporter());

		assert.equal(engine.getPagePixmap(0, 100, 100), null);
		assert.equal(engine.getPagePixmap(2, 100, 100), null);
		assert.deepEqual(engine.getAnnotationsAP(2, 100, 100, undefined, 7), []);
		undo.call({ Class: doc });
		assert.equal(file.originalPagesCount, 2);

		engine.pages.splice(0, 1, newImport);
		nativePages.push(newImport);
		assert.equal(engine.getPagePixmap(0, 100, 100), 99);
		assert.deepEqual(engine.getAnnotationsAP(2, 100, 100, undefined, 7), [{ i: 7 }]);
		assert.equal(engine.getPagePixmap(2, 100, 100), null);
		assert.equal(engine.calls.getPagePixmap.length, 3);
	},

	'looks up failures by the current page index and annotation arguments'() {
		const engine = createEngine(throwAlways);
		engine.pages = [{ originIndex: 2 }, { originIndex: 1 }];
		guardNativeDrawing(engine, createReporter());

		engine.getPagePixmap(0, 100, 100);
		engine.getAnnotationsAP(2, 100, 100, undefined, 7);
		engine.pages.reverse();
		assert.equal(engine.hasDrawingFailure('getPagePixmap', 0, 100, 100), false);
		assert.equal(engine.hasDrawingFailure('getPagePixmap', 1, 100, 100), true);
		assert.equal(engine.hasDrawingFailure('getAnnotationsAP', 2, 100, 100), false);
		assert.equal(engine.hasDrawingFailure('getAnnotationsAP', 2, 100, 100, 7), true);
	}
};

let failed = 0;
for (const [name, test] of Object.entries(tests)) {
	try {
		test();
		console.log('PASS ' + name);
	} catch (e) {
		failed++;
		console.log('FAIL ' + name + '\n  ' + (e && e.message));
	}
}
if (failed) {
	console.log(failed + ' failed');
	process.exit(1);
}
console.log('all passed');
