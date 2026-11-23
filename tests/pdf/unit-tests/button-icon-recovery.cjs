const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../../..');
const appearance = {normal: 0, mouseDown: 1, rollover: 2};
const position = {textOnly: 0, iconOnly: 1};
const highlight = {invert: 1, push: 2};
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAJ+VGNkAAAAASUVORK5CYII=';

function createDocument(getButtonIcons, pages, imageLoader, recordHistory = false) {
	const calls = [];
	const errors = [];
	const changes = [];
	let noHistory = 0;
	const historyRecorder = {
		UndoRedoInProgress: false,
		IsOn() { return noHistory === 0; },
		StartNoHistoryMode() { noHistory++; },
		EndNoHistoryMode() { noHistory--; },
		Add(change) {
			if (recordHistory && this.IsOn() && !this.UndoRedoInProgress)
				changes.push(change);
		}
	};
	const view = {
		file: {
			pages,
			nativeFile: {
				getButtonIcons(page, width, height, background, base64, widget, state) {
					calls.push([page, width, height, background, base64, widget, state]);
					return getButtonIcons(page, widget, state);
				}
			}
		},
		IsOpenFormsInProgress: true
	};
	function BaseField(name, type, rect, doc) { this.doc = doc; this.name = name; }
	BaseField.prototype.SetWasChanged = function() {};
	BaseField.prototype.AddToRedraw = function() {};
	BaseField.prototype.GetParentPage = function() { return this.parentPage; };
	BaseField.prototype.GetParent = function() { return null; };
	BaseField.prototype.GetPartialName = function() { return this.name; };
	BaseField.prototype.IsNeedDrawFromStream = function() { return false; };
	BaseField.prototype.SetNeedRecalc = function() {};
	BaseField.prototype.SetHasOriginView = function() {};
	BaseField.prototype.GetType = function() { return 27; };
	BaseField.prototype.SetOriginPage = function(page) { this.originPage = page; };
	BaseField.prototype.GetOriginPage = function() { return this.originPage; };
	BaseField.prototype.SetApIdx = function(id) { this.apIdx = id; };
	BaseField.prototype.GetApIdx = function() { return this.apIdx; };
	BaseField.prototype.WriteToBinaryBase = function() {};
	BaseField.prototype.WriteToBinaryBase2 = function(memory) {
		memory.posForWidgetFlags = memory.GetCurPosition();
		memory.Skip(4);
		memory.posForFieldDataFlags = memory.GetCurPosition();
		memory.Skip(4);
		memory.buttonFlagsPos = memory.GetCurPosition();
		memory.fieldDataFlags = 0;
		memory.widgetFlags = 0;
	};
	const AscPDF = {
		Api: {Types: {position, scaleWhen: {always: 0}, scaleHow: {proportional: 0}, highlight}},
		FIELD_TYPES: {button: 27},
		APPEARANCE_TYPES: appearance,
		ALIGN_TYPE: {center: 1},
		BUTTON_HIGHLIGHT_TYPES: highlight,
		CBaseField: BaseField,
		CTextBoxContent: function() {
			const run = {ClearContent() {}};
			const paragraph = {
				Content: [run], GetElementsCount: () => 1, GetElement: () => run,
				RemoveFromContent() {}, CorrectContent() {}
			};
			this.Content = [paragraph];
			this.GetElement = () => paragraph;
			this.GetAllDrawingObjects = () => [];
			this.SetAlign = function() {};
		}
	};
	const editor = {
		getDocumentRenderer: () => view,
		ImageLoader: imageLoader,
		sendEvent(...args) { errors.push(args); }
	};
	const context = vm.createContext({
		window: {AscPDF, NATIVE_EDITOR_ENJINE: !imageLoader},
		AscPDF,
		editor,
		Asc: {editor, c_oAscError: {ID: {CopyDisabled: -62}, Level: {NoCritical: 0}}},
		AscCommon: {
			AscBrowser: {retinaPixelRatio: 1},
			CWordCollaborativeEditing: function() {},
			CCollaborativeEditingBase: function() {},
			contentchanges_Add: 1,
			History: historyRecorder,
			getFullImageSrc2: (src) => src,
			CommandType: {ctAnnotField: 164}
		},
		AscFormat: {
			CBaseNoIdObject: function() {},
			CCopyObjectProperties: function() {},
			InitClass(derived, base) { Object.setPrototypeOf(derived.prototype, base.prototype); },
			ExecuteNoHistory(fn, self, args) { return fn.apply(self, args); },
			getTargetTextObject: () => null,
			fResetConnectorsIds() {}
		},
		AscDFH: {
			CChangesBaseContentChange: function(object, pos, items, add) {
				this.Class = object;
				this.Pos = pos;
				this.Items = items;
				this.IsAdd = () => add;
			},
			CChangesBaseStringProperty: function(object, oldValue, newValue) {
				this.Class = object;
				this.Old = oldValue;
				this.New = newValue;
			}
		},
		PDFEditorApi: function() {},
		isRealObject: (value) => !!value && typeof value === 'object'
	});
	context.AscDFH.CChangesBaseLongProperty = context.AscDFH.CChangesBaseStringProperty;
	context.AscDFH.CChangesPDFArrayOfDoubleProperty = context.AscDFH.CChangesBaseStringProperty;
	context.AscCommon.CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges = function() {};
	for (const file of ['pdf/src/document.js', 'pdf/src/forms/pushbutton.js', 'pdf/src/CollaborativeEditing.js']) {
		vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, {filename: file});
	}
	const history = fs.readFileSync(path.join(root, 'pdf/src/history/formsChanges.js'), 'utf8');
	const historyStart = history.indexOf('function CChangesPDFPushbuttonImage(');
	const historyEnd = history.indexOf('function CChangesPDFPushbuttonScaleWhenType(', historyStart);
	vm.runInContext(history.slice(historyStart, historyEnd), context, {filename: 'pdf/src/history/formsChanges.js'});
	const documentHistory = fs.readFileSync(path.join(root, 'pdf/src/history/documentChanges.js'), 'utf8');
	const fieldsStart = documentHistory.indexOf('function CChangesPDFDocumentFieldsContent(');
	const fieldsEnd = documentHistory.indexOf('/**', fieldsStart);
	vm.runInContext(documentHistory.slice(fieldsStart, fieldsEnd), context, {filename: 'pdf/src/history/documentChanges.js'});
	const api = fs.readFileSync(path.join(root, 'pdf/api.js'), 'utf8');
	const start = api.indexOf('PDFEditorApi.prototype.asc_CheckCopy =');
	const end = api.indexOf('PDFEditorApi.prototype.asc_SelectionCut =', start);
	vm.runInContext(api.slice(start, end), context, {filename: 'pdf/api.js'});
	editor.asc_CheckCopy = context.PDFEditorApi.prototype.asc_CheckCopy;
	const layoutStart = api.indexOf('PDFEditorApi.prototype.SetButtonFieldLayout =');
	const layoutEnd = api.indexOf('PDFEditorApi.prototype.SetButtonFieldScaleWhen =', layoutStart);
	vm.runInContext(api.slice(layoutStart, layoutEnd), context, {filename: 'pdf/api.js'});
	editor.SetButtonFieldLayout = context.PDFEditorApi.prototype.SetButtonFieldLayout;
	const clearStart = api.indexOf('PDFEditorApi.prototype.ClearButtonFieldImage =');
	const clearEnd = api.indexOf('///////// For drawings', clearStart);
	vm.runInContext(api.slice(clearStart, clearEnd), context, {filename: 'pdf/api.js'});
	editor.ClearButtonFieldImage = context.PDFEditorApi.prototype.ClearButtonFieldImage;
	editor.DocumentRenderer = view;
	const doc = Object.create(AscPDF.CPDFDoc.prototype);
	doc.Viewer = view;
	doc.Api = editor;
	doc.CollaborativeEditing = new AscPDF.CPDFCollaborativeEditing(doc);
	doc.private_UpdateTargetForCollaboration = () => {};
	editor.getPDFDoc = () => doc;
	function addButton(page, id) {
		const button = new AscPDF.CPushButtonField('button', [0, 0, 10, 10], doc);
		button.GetOriginPage = () => page;
		button.GetApIdx = () => id;
		button.GetType = () => 27;
		return button;
	}
	const page = {
		fields: [],
		GetIndex: () => 0,
		GetOriginIndex: () => pages[0].originIndex,
		GetDocument: () => doc,
		fieldsContentChanges: {Check: (type, pos) => pos}
	};
	doc.GetPageInfo = () => page;
	doc.DoAction = (fn) => fn();
	doc.CreateField = (name, type, rect) => {
		const button = new AscPDF.CPushButtonField(name, rect, doc);
		button.DoInitialRecalc = function() { return true; };
		button._UpdateImage = function() {};
		return button;
	};
	page.AddField = insertButton;
	function readButton(properties = {}) {
		doc.private_AddFormsByInfo({Fields: [{
			name: 'button', type: 27, rect: {x1: 0, y1: 0, x2: 10, y2: 10},
			page: pages[0].originIndex, AP: {i: 11, have: true}, ...properties
		}]});
		return page.fields.at(-1);
	}
	function insertButton(button) {
		const change = new context.CChangesPDFDocumentFieldsContent(page, page.fields.length, [button], true);
		changes.push(change);
		change.Redo();
		return change;
	}
	function replay(direction, point = changes) {
		historyRecorder.UndoRedoInProgress = true;
		try {
			const undo = direction === 'Undo';
			for (const change of undo ? [...point].reverse() : point) {
				if (change[direction])
					change[direction]();
				else
					change.private_SetValue(undo ? change.Old : change.New);
			}
		} finally {
			historyRecorder.UndoRedoInProgress = false;
		}
	}
	function applyImageFromPeer(button, raster, type) {
		const sent = new context.CChangesPDFPushbuttonImage(button, undefined, raster, type);
		const encoded = [];
		sent.WriteToBinary({WriteLong: (value) => encoded.push(value), WriteString2: (value) => encoded.push(value)});
		const change = new context.CChangesPDFPushbuttonImage(button);
		change.ReadFromBinary({GetLong: () => encoded.shift(), GetString2: () => encoded.shift()});
		historyRecorder.StartNoHistoryMode();
		try {
			change.private_SetValue(change.New);
		} finally {
			historyRecorder.EndNoHistoryMode();
		}
	}
	return {doc, view, calls, errors, changes, addButton, editor, insertButton, replay, readButton, page, applyImageFromPeer};
}

function icon(widget, state, base64) {
	return {MK: [{i: widget, [state]: 7}], View: [{j: 7, w: 2, h: 2, retValue: base64}]};
}

function recoveryFixture() {
	const fixture = createDocument((page) => (
		page === 5 ? icon(11, 'I', 'AAAA') : null
	), [
		{originIndex: 4, W: 600, H: 800, Dpi: 96},
		{originIndex: 5, W: 600, H: 800, Dpi: 96}
	]);
	fixture.healthy = fixture.addButton(5, 11);
	fixture.otherDamaged = fixture.addButton(4, 13);
	fixture.damaged = fixture.addButton(4, 12);
	fixture.doc.private_FillButtonsIconsOnOpen([fixture.healthy, fixture.damaged, fixture.otherDamaged]);
	return fixture;
}

test('a failed page icon read is attempted once and other pages still load their icons', () => {
	const {calls, healthy, damaged, otherDamaged, view} = recoveryFixture();
	assert.deepEqual(calls, [
		[4, 600, 800, undefined, true, undefined, undefined],
		[5, 600, 800, undefined, true, undefined, undefined]
	]);
	assert.equal(healthy.GetImageRasterId(), 'data:image/png;base64,AAAA');
	assert.deepEqual(serializeButton(damaged), {flags: 0, images: []});
	assert.deepEqual(serializeButton(otherDamaged), {flags: 0, images: []});
	assert.equal(view.IsOpenFormsInProgress, false);
});

test('an empty page icon result completes loading without retrying widgets', () => {
	const {doc, view, calls, addButton} = createDocument(() => ({}), [{originIndex: 0, W: 600, H: 800, Dpi: 96}]);
	const button = addButton(0, 11);
	doc.private_FillButtonsIconsOnOpen([button]);
	assert.equal(calls.length, 1);
	assert.equal(button.GetImageRasterId(), '');
	assert.equal(view.IsOpenFormsInProgress, false);
});

function serializeButton(button, mode = {}) {
	const data = Buffer.alloc(256);
	let pos = 0;
	const memory = {
		...mode,
		images: [],
		GetCurPosition: () => pos,
		Skip(length) { pos += length; },
		Seek(target) { pos = target; },
		WriteByte(value) { data.writeUInt8(value, pos); pos++; },
		WriteLong(value) { data.writeInt32LE(value, pos); pos += 4; },
		WriteDouble(value) { data.writeDoubleLE(value, pos); pos += 8; }
	};
	button.WriteToBinary(memory);
	return {flags: data.readInt32LE(memory.buttonFlagsPos) & 0xe0, images: memory.images};
}

test('incremental saves write only explicitly replaced appearances after a failed read', () => {
	const {damaged} = recoveryFixture();
	assert.deepEqual(serializeButton(damaged), {flags: 0, images: []});
	damaged.SetImageRasterId('data:image/png;base64,CCCC', appearance.normal);
	assert.deepEqual(serializeButton(damaged), {flags: 0x20, images: ['data:image/png;base64,CCCC']});
});

test('split serialization embeds recovered PNG data without decoding it again', () => {
	const {doc, addButton} = createDocument(() => icon(11, 'I', png), [{originIndex: 0, W: 600, H: 800, Dpi: 96}]);
	const button = addButton(0, 11);
	doc.private_FillButtonsIconsOnOpen([button]);
	assert.deepEqual(serializeButton(button, {isForSplit: true}), {flags: 0xe0, images: ['data:image/png;base64,' + png, '']});
});

test('a failed browser image decode leaves the original appearance untouched on save', () => {
	const imageLoader = {
		map_image_index: {},
		LoadImagesWithCallback([src], callback) {
			this.map_image_index[src] = {Image: null};
			callback();
		}
	};
	const {doc, view, addButton} = createDocument(() => icon(11, 'I', 'YWJj'), [{originIndex: 0, W: 600, H: 800, Dpi: 96}], imageLoader);
	const button = addButton(0, 11);
	doc.private_FillButtonsIconsOnOpen([button]);
	assert.deepEqual(serializeButton(button), {flags: 0xc0, images: ['']});
	assert.equal(view.IsOpenFormsInProgress, false);
});

test('a remote editor loads and preserves original icons from the imported page', () => {
	const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
	const peer = createDocument(() => ({
		MK: [{i: 20, I: 7}, {i: 21, I: 8}, {i: 22, I: 7}],
		View: [{j: 7, w: 2, h: 2, retValue: png}]
	}), pages, undefined, true);
	const healthy = peer.addButton(5, 20);
	const damaged = peer.addButton(5, 21);
	const cleared = peer.addButton(5, 22);
	for (const button of [healthy, damaged, cleared]) {
		const insertion = peer.insertButton(button);
		insertion.Undo();
		insertion.Load();
	}
	cleared.SetImage = function() {};
	peer.applyImageFromPeer(cleared, '', appearance.normal);
	peer.doc.CollaborativeEditing.OnEnd_ReadForeignChanges();
	assert.equal(peer.calls.length, 1);
	assert.equal(healthy.GetImageRasterId(), 'data:image/png;base64,' + png);
	assert.deepEqual(serializeButton(damaged), {flags: 0xc0, images: ['']});
	assert.equal(cleared.GetImageRasterId(), '');
	peer.applyImageFromPeer(cleared, undefined, appearance.normal);
	assert.equal(cleared.GetImageRasterId(), 'data:image/png;base64,' + png);
	assert.equal(peer.changes.length, 3);
});

function delayedImageLoader() {
	return {
		map_image_index: {},
		LoadImagesWithCallback([src], callback) {
			this.finish = () => {
				this.map_image_index[src] = {Image: {}};
				callback();
			};
		}
	};
}

function selectButton(doc, view, button) {
	view.file.Selection = {};
	const shape = {
		selected: true, IsEditFieldShape: () => true, IsAnnot: () => false, GetEditField: () => button,
		GetDocContent: () => null, IsInTextBox: () => false
	};
	doc.GetController = () => ({selection: {}, selectedObjects: [shape]});
	doc.activeDrawing = shape;
}

test('opening an invert-highlight button loads original rollover and mouse-down icons', () => {
	const imageLoader = delayedImageLoader();
	const pages = [{originIndex: 0, W: 600, H: 800, Dpi: 96}];
	const {readButton} = createDocument(() => ({
		MK: [{i: 11, RI: 7, IX: 7}], View: [{j: 7, retValue: png}]
	}), pages, imageLoader);
	const button = readButton({position: position.iconOnly, highlight: highlight.invert});
	assert.equal(typeof imageLoader.finish, 'function');
	imageLoader.finish();
	assert.equal(button.GetImageRasterId(appearance.rollover), 'data:image/png;base64,' + png);
	assert.equal(button.GetImageRasterId(appearance.mouseDown), 'data:image/png;base64,' + png);
});

for (const redoWhileLoading of [false, true]) {
	test(`undo and redo preserve the copied icon-only button when redo happens ${redoWhileLoading ? 'before' : 'after'} decoding`, () => {
		const imageLoader = delayedImageLoader();
		const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
		const {doc, view, page, readButton, replay} =
			createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
		const button = readButton({position: position.iconOnly, highlight: highlight.invert});
		assert.equal(button.GetLayout(), position.iconOnly);
		assert.equal(button.HasUnreadOriginalIcons(), true);
		replay('Undo');
		assert.equal(page.fields.length, 0);
		assert.equal(button.GetLayout(), position.textOnly);

		selectButton(doc, view, button);
		if (redoWhileLoading) {
			replay('Redo');
			assert.equal(doc.CanCopyCut().copy, false);
		}
		imageLoader.finish();
		if (!redoWhileLoading) {
			button.CheckImageOnce();
			replay('Redo');
		}

		assert.equal(page.fields[0], button);
		assert.equal(button.GetLayout(), position.iconOnly);
		button.CheckImageOnce();
		const displayedImages = [];
		button.getAllRasterImages(displayedImages);
		assert.deepEqual(displayedImages, ['data:image/png;base64,' + png]);
		assert.equal(doc.CanCopyCut().copy, true);
		assert.equal(doc.CanCopyCut().cut, true);
		assert.deepEqual(serializeButton(button, {isCopyPaste: true}), {
			flags: 0xe0, images: ['data:image/png;base64,' + png, '']
		});
	});
}

test('clearing an imported icon-only button through the editor wins over pending decode', () => {
	const imageLoader = delayedImageLoader();
	const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
	const {doc, view, editor, readButton} = createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
	const button = readButton({position: position.iconOnly, highlight: highlight.invert});
	selectButton(doc, view, button);
	doc.activeForm = button;
	assert.equal(editor.ClearButtonFieldImage(appearance.normal), true);
	imageLoader.finish();
	assert.equal(button.GetImageRasterId(appearance.normal), '');
	assert.equal(button.HasUnreadOriginalIcons(), false);
	assert.deepEqual(serializeButton(button, {isCopyPaste: true}), {flags: 0xe0, images: ['']});
});

test('an explicit change to text-only clears the icon and its own undo restores it', () => {
	const imageLoader = delayedImageLoader();
	const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
	const {doc, view, editor, changes, readButton, replay} =
		createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
	const button = readButton({position: position.iconOnly, highlight: highlight.invert});
	imageLoader.finish();
	selectButton(doc, view, button);
	doc.activeForm = button;

	const firstChange = changes.length;
	assert.equal(editor.SetButtonFieldLayout(position.textOnly), true);
	assert.equal(button.GetImageRasterId(), '');
	const layoutChanges = changes.slice(firstChange);
	replay('Undo', layoutChanges);
	assert.equal(button.GetLayout(), position.iconOnly);
	assert.equal(button.GetImageRasterId(), 'data:image/png;base64,' + png);
	replay('Redo', layoutChanges);
	assert.equal(button.GetLayout(), position.textOnly);
	assert.equal(button.GetImageRasterId(), '');
});

for (const redoWhileLoading of [false, true]) {
	test(`an inserted button can be copied and cut when redo happens ${redoWhileLoading ? 'before' : 'after'} decoding its icon`, () => {
		const imageLoader = delayedImageLoader();
		const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
		const {doc, view, changes, addButton, insertButton, replay} =
			createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
		const button = addButton(5, 11);
		button.SetImage = function() {};
		const insertion = insertButton(button);
		doc.private_FillButtonsIconsOnOpen([button]);
		assert.equal(button.HasUnreadOriginalIcons(), true);
		replay('Undo');
		assert.equal(insertion.Class.fields.length, 0);

		if (redoWhileLoading)
			replay('Redo');
		imageLoader.finish();
		if (!redoWhileLoading)
			replay('Redo');

		assert.equal(insertion.Class.fields[0], button);
		assert.equal(button.GetImageRasterId(), 'data:image/png;base64,' + png);
		assert.deepEqual(changes, [insertion]);
		selectButton(doc, view, button);
		const {copy, cut} = doc.CanCopyCut();
		assert.equal(copy, true);
		assert.equal(cut, true);
	});
}

test('delayed decode updates the undone appearance while preserving other edits', () => {
	const imageLoader = delayedImageLoader();
	const pages = [{originIndex: 0, W: 600, H: 800, Dpi: 96}];
	const {doc, view, changes, addButton} = createDocument(() => ({
		MK: [{i: 11, I: 7, RI: 7, IX: 7}],
		View: [{j: 7, w: 2, h: 2, retValue: png}]
	}), pages, imageLoader, true);
	const button = addButton(0, 11);
	button.SetImage = function() {};
	doc.private_FillButtonsIconsOnOpen([button]);
	button.SetImageRasterId('replacement', appearance.normal);
	button.SetImageRasterId('', appearance.rollover);
	button.SetImageRasterId('replacement', appearance.mouseDown);
	const replacement = changes.at(-1);
	replacement.private_SetValue(replacement.Old);
	imageLoader.finish();
	assert.equal(button.GetImageRasterId(appearance.normal), 'replacement');
	assert.equal(button.GetImageRasterId(appearance.rollover), '');
	assert.deepEqual(serializeButton(button), {
		flags: 0xe0, images: ['replacement', '', 'data:image/png;base64,' + png]
	});
	assert.equal(view.IsOpenFormsInProgress, false);
});

for (const undoWhileLoading of [false, true]) {
	test(`undo restores the original icon ${undoWhileLoading ? 'before' : 'after'} decoding finishes`, () => {
		const imageLoader = delayedImageLoader();
		const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
		const {doc, view, changes, readButton, replay} =
			createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
		const button = readButton({position: position.iconOnly, highlight: highlight.invert});
		selectButton(doc, view, button);

		const firstChange = changes.length;
		button.AddImage2('replacement', appearance.normal);
		const replacement = changes.slice(firstChange);
		assert.equal(replacement.length, 1);

		if (undoWhileLoading) {
			replay('Undo', replacement);
			assert.equal(doc.CanCopyCut().copy, false);
			assert.deepEqual(serializeButton(button), {flags: 0xc0, images: ['']});
			replay('Redo', replacement);
			assert.equal(button.GetImageRasterId(), 'replacement');
			replay('Undo', replacement);
			imageLoader.finish();
		} else {
			imageLoader.finish();
			assert.equal(button.GetImageRasterId(), 'replacement');
			replay('Undo', replacement);
		}

		button.CheckImageOnce();
		const original = 'data:image/png;base64,' + png;
		assert.equal(button.GetImageRasterId(), original);
		const displayedImages = [];
		button.getAllRasterImages(displayedImages);
		assert.deepEqual(displayedImages, [original]);
		assert.equal(doc.CanCopyCut().copy, true);
		assert.deepEqual(serializeButton(button, {isCopyPaste: true}), {
			flags: 0xe0, images: [original, '']
		});
		assert.equal(changes.length, firstChange + 1);
	});
}

test('a peer undo restores an already decoded original after a clear', () => {
	const imageLoader = delayedImageLoader();
	const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
	const {doc, view, changes, readButton, applyImageFromPeer} =
		createDocument(() => icon(11, 'I', png), pages, imageLoader, true);
	const button = readButton({position: position.iconOnly, highlight: highlight.invert});
	imageLoader.finish();
	selectButton(doc, view, button);

	const changeCount = changes.length;
	applyImageFromPeer(button, '', appearance.normal);
	assert.equal(button.GetImageRasterId(), '');
	applyImageFromPeer(button, undefined, appearance.normal);
	const original = 'data:image/png;base64,' + png;
	assert.equal(button.GetImageRasterId(), original);
	const displayedImages = [];
	button.getAllRasterImages(displayedImages);
	assert.deepEqual(displayedImages, [original]);
	assert.equal(doc.CanCopyCut().copy, true);
	assert.equal(changes.length, changeCount);
});

test('a peer undo keeps a button copyable when the original appearance has no icon', () => {
	const pages = [{originIndex: 5, W: 600, H: 800, Dpi: 96}];
	const {doc, view, readButton, applyImageFromPeer} =
		createDocument(() => ({MK: [], View: []}), pages);
	const button = readButton({position: position.iconOnly, highlight: highlight.invert});
	selectButton(doc, view, button);

	applyImageFromPeer(button, '', appearance.normal);
	applyImageFromPeer(button, undefined, appearance.normal);
	assert.equal(button.GetImageRasterId(), '');
	assert.equal(doc.CanCopyCut().copy, true);
});

test('a standalone widget with an unread icon cannot be copied or cut', () => {
	const {doc, view, errors, damaged, editor} = recoveryFixture();
	view.file.Selection = {Glyph1: 0, Glyph2: 0};
	const shape = {selected: true, IsEditFieldShape: () => true, IsAnnot: () => false, GetEditField: () => damaged};
	doc.GetController = () => ({selection: {}, selectedObjects: [shape], GetSelectedText: () => ''});
	doc.GetActiveObject = () => damaged;
	assert.equal(editor.asc_CheckCopy(null, 0), false);
	assert.deepEqual(errors, [['asc_onError', -62, 0]]);
	assert.equal(doc.CanCopyCut().copy, false);
	assert.equal(doc.CanCopyCut().cut, false);
});
