/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
	};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getCellCache = function (col, row) {
		let _cell = null;
		this.model.getRange3(row, col, row, col)._foreachNoEmpty(function(cell, row, col) {
			if (cell && !cell.isEmptyTextString()) {
				_cell = {cellType: cell.getType()}
			}
		}, null, true);
		return _cell;
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (range, subType, callback) {
		callback(true);
		return true;
	};
	AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
	};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function(){};

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	}

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	wb.handlers.add("asc_onConfirmAction", function (test1, callback) {
		callback(true);
	});
	api.wb.cellCommentator = new AscCommonExcel.CCellCommentator({
		model: api.wbModel.aWorksheets[0],
		collaborativeEditing: null,
		draw: function() {
		},
		handlers: {
			trigger: function() {
				return false;
			}
		}
	});

	AscCommonExcel.CCellCommentator.prototype.isLockedComment = function (oComment, callbackFunc) {
		callbackFunc(true);
	};
	AscCommonExcel.CCellCommentator.prototype.drawCommentCells = function () {
	};
	AscCommonExcel.CCellCommentator.prototype.ascCvtRatio = function () {
	};

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};
	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	};

	function checkUndoRedo(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	}

	let getRgbColor = function(clr){
		var color = (typeof(clr) == 'object') ? clr.color : clr;

		color=color.replace(/#/,'');
		if(color.length==3) color=color.replace(/(.)/g,'$1$1');
		color=parseInt(color,16);
		var c = new Asc.asc_CColor();
		c.put_type( (typeof(clr) == 'object' && clr.effectId !== undefined)? Asc.c_oAscColor.COLOR_TYPE_SCHEME : Asc.c_oAscColor.COLOR_TYPE_SRGB);
		c.put_r(color>>16);
		c.put_g((color&0xff00)>>8);
		c.put_b(color&0xff);
		c.put_a(0xff);
		if (clr.effectId !== undefined)
			c.put_value(clr.effectId);
		return c;
	};

	function compareAscColorAndRgbColor (ascColor, rgbColor) {
		if ((ascColor && !rgbColor) || (!ascColor && rgbColor)) {
			return false;
		}
		return ascColor.get_r() === rgbColor.getR() && ascColor.get_g() === rgbColor.getG() && ascColor.get_b() === rgbColor.getB();
	};

	QUnit.test('Conditional formatting: test apply to', function (assert) {

		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		tableOptions.range = "A1:B3";
		api.asc_addAutoFilter("TableStyleMedium2", tableOptions);

		let cf = new AscCommonExcel.CConditionalFormattingRule();
		cf.asc_setType(Asc.c_oAscCFType.cellIs);
		cf.asc_setLocation("A5");

		api.asc_setCF([cf]);

		wsView.setSelection(new Asc.Range(0, 4, 0, 4));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		let cfLocation;
		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		let ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$5", "compare location conditional formatting in cell");


		cf = new AscCommonExcel.CConditionalFormattingRule();
		cf.asc_setType(Asc.c_oAscCFType.cellIs);
		cf.asc_setLocation("=Table1[Column1]");

		api.asc_setCF([cf]);

		wsView.setSelection(new Asc.Range(0, 1, 0, 1));
		modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);

		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$2:$A$4", "compare location conditional formatting in table");


		clearData(0, 6, 0, 6);
	});


	QUnit.test("Test: \"simple tests\"", function (assert) {
		let testData = [
			["100"]
		];

		let range = ws.getRange4(0, 0);
		range.fillData(testData);


		let newRule = new AscCommonExcel.CConditionalFormattingRule();
		newRule.asc_setType(Asc.c_oAscCFType.cellIs);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("9C0006");
		let fillColor = getRgbColor("FFC7CE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		newRule.asc_setDxf(xfs);
		newRule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		newRule.asc_setValue1("99");
		newRule.asc_setLocation("$A$1");

		api.asc_setCF([newRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, AscCommon.gc_nMaxCol0, AscCommon.gc_nMaxRow0)]));

		wsView.setSelection(new Asc.Range(0, 0, 0, 0));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);

		let cfLocation;
		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		let ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$1", "compare location conditional formatting");


		let compiledStyle = ws.getCompiledStyle(0, 0);
		let rgbColor = compiledStyle.fill.bg();
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor), true, "compare fill color _1");


		//Clearing data of sheet
		clearData(0, 0, 1, 8);
	});

	QUnit.test('Conditional formatting: contains text tests', function (assert) {
		// Preparing test data
		let testData = [
			["This is sample text"],
			["Text with sample word"],
			["No matching here"],
			["SAMPLE in uppercase"]
		];

		let range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Test 1: Check simple text condition
		let textRule = new AscCommonExcel.CConditionalFormattingRule();
		textRule.asc_setType(Asc.c_oAscCFType.containsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("2F75B5");
		let fillColor = getRgbColor("BDD7EE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		textRule.asc_setDxf(xfs);
		textRule.asc_setContainsText("sample");
		textRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([textRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check that the rule was applied
		wsView.setSelection(new Asc.Range(0, 0, 0, 0));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		assert.ok(modelCf && modelCf[0] && modelCf[0][0], "CF rule exists");

		// Check that the style was applied to cells with 'sample'
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1.fill.bg();
		let rgbColor2 = compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 matches 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 matches 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't match 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 matches 'sample' condition (case insensitive)");

		// Test 2: Check cell reference in condition
		clearData(0, 0, 0, 4);

		// New test data
		testData = [
			["apple"], // A1: search word
			["This is an apple text"], // A2: contains the word
			["Orange juice"], // A3: doesn't contain the word
			["Pineapple is good"] // A4: contains the word as part of another word
		];

		range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Create rule with cell reference
		let cellRefRule = new AscCommonExcel.CConditionalFormattingRule();
		cellRefRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("006400");
		fillColor = getRgbColor("C6EFCE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		cellRefRule.asc_setDxf(xfs);
		cellRefRule.asc_setContainsText("=A1"); // Cell reference, not text
		cellRefRule.asc_setLocation("$A$2:$A$4");

		api.asc_setCF([cellRefRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 3)]));

		// Check that the rule was applied
		wsView.setSelection(new Asc.Range(0, 1, 0, 1));
		modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		assert.ok(modelCf && modelCf[0] && modelCf[0][0], "CF rule with cell reference exists");

		// Check that the style was applied to cells with 'apple'
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 matches value from A1");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), true, "Cell A3 doesn't match value from A1");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 matches value from A1 (as part of word)");

		// Test 3: Check begins with / ends with
		clearData(0, 0, 0, 5);

		// New test data for begins/ends with
		testData = [
			["Spreadsheet testing"],
			["Testing Spreadsheet"],
			["Microsoft Spreadsheet"],
			["spreadsheet lowercase"]
		];

		range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Rule "Begins with Spreadsheet"
		let beginsWithRule = new AscCommonExcel.CConditionalFormattingRule();
		beginsWithRule.asc_setType(Asc.c_oAscCFType.beginsWith);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		fillColor = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		beginsWithRule.asc_setDxf(xfs);
		beginsWithRule.asc_setContainsText("Spreadsheet");
		beginsWithRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([beginsWithRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check cell styles
		compiledStyle1 = ws.getCompiledStyle(0, 0);
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 begins with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), false, "Cell A2 doesn't begin with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't begin with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 begins with 'spreadsheet' (case insensitive)");

		// Clean up data after tests
		clearData(0, 0, 0, 5);
	});

	QUnit.test('Conditional formatting: not contains text tests', function (assert) {
		// Preparing test data
		let testData = [
			["Report 2023"],
			["Sales data"],
			["Error in data"],
			["Missing values"]
		];

		let range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Rule "Does not contain Error"
		let notContainsRule = new AscCommonExcel.CConditionalFormattingRule();
		notContainsRule.asc_setType(Asc.c_oAscCFType.notContainsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("3F3F76");
		let fillColor = getRgbColor("B4C6E7");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		notContainsRule.asc_setDxf(xfs);
		notContainsRule.asc_setContainsText("Error");
		notContainsRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([notContainsRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check cell styles
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		let rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 doesn't contain 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 doesn't contain 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 contains 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 doesn't contain 'Error'");

		// Clean up data after tests
		clearData(0, 0, 5, 5);
	});

	QUnit.test('Conditional formatting: asc_setContainsText with formula and text', function (assert) {
		// Preparing test data
		let testData = [
			["Product"], // A1: header
			["CPU Intel"], // A2
			["Graphics card AMD"], // A3
			["Intel Motherboard"], // A4
			["AMD Processor"] // A5
		];

		let range = ws.getRange4(0, 0, 4, 0);
		range.fillData(testData);

		// Additional value for formula reference
		let referenceData = [
			["Intel"] // B1: value for reference
		];
		ws.getRange4(0, 1).fillData(referenceData);

		// Test 1: Using plain text
		let textRule = new AscCommonExcel.CConditionalFormattingRule();
		textRule.asc_setType(Asc.c_oAscCFType.containsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("2F75B5");
		let fillColor = getRgbColor("BDD7EE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		textRule.asc_setDxf(xfs);
		textRule.asc_setContainsText("Intel"); // Plain text
		textRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([textRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for text search
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);
		let compiledStyle5 = ws.getCompiledStyle(4, 0);

		let rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		let rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 contains 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't contain 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 contains 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), false, "Cell A5 doesn't contain 'Intel' text");


		// Test 2: Using cell reference formula
		let formulaRule = new AscCommonExcel.CConditionalFormattingRule();
		formulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("006400");
		fillColor = getRgbColor("C6EFCE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		formulaRule.asc_setDxf(xfs);
		formulaRule.asc_setContainsText("=B1"); // Formula reference
		formulaRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([formulaRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for formula reference
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);
		compiledStyle5 = ws.getCompiledStyle(4, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		//TODO!
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 contains value from B1 (Intel)");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't contain value from B1");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 contains value from B1 (Intel)");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), false, "Cell A5 doesn't contain value from B1");

		// Test 3: Using more complex formula
		// Change reference data
		referenceData = [
			["AMD"] // B1: changed value
		];
		ws.getRange4(0, 1).fillData(referenceData);

		let complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		fillColor = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=UPPER(B1)"); // Formula with function
		complexFormulaRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([complexFormulaRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for complex formula
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);
		compiledStyle5 = ws.getCompiledStyle(4, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), false, "Cell A2 doesn't contain 'AMD'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), true, "Cell A3 contains 'AMD'");
		//TODO!
		//assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), false, "Cell A4 doesn't contain 'AMD'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), true, "Cell A5 contains 'AMD'");

		clearData(0, 0, 5, 5);

		// Test 4
		testData = [
			["test"],//A1
			["test"],//A2
			["test"],//A3
			["test"],//A4
			["test"]//A5
		];

		range = ws.getRange4(0, 0, 4, 0);
		range.fillData(testData);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor1 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor1);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=test");
		complexFormulaRule.asc_setLocation("$A$1");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor2 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor2);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=\"test\"");
		complexFormulaRule.asc_setLocation("$A$2");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor3 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor3);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("test");
		complexFormulaRule.asc_setLocation("$A$3");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor4 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor4);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("\"test\"");
		complexFormulaRule.asc_setLocation("$A$4");

		api.asc_setCF([complexFormulaRule]);

		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 4)]));

		// Check that styles are applied correctly for complex formula
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor1, rgbColor1), false, "formula + text without quotes");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor2, rgbColor2), true, "formula + text with quotes");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor3, rgbColor3), true, "text without quote");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor4, rgbColor4), false, "text with quotes");

		// Clean up data after tests
		clearData(0, 0, 5, 5);
	});

	QUnit.test("Deleting and restoring a formula rule updates its dependencies", function (assert) {
		const history = AscCommon.History;
		const cell = ws.getRange4(1, 1);
		const source = ws.getRange4(0, 0);
		const fillColor = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fillColor);

		cell.setValue("1");
		source.setValue("70");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2:$B$6");
		rule.asc_setValue1("=$A$1>50");
		rule.asc_setDxf(dxf);
		api.asc_setCF([rule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(rule.ranges));

		const isHighlighted = () => compareAscColorAndRgbColor(fillColor, ws.getCompiledStyle(1, 1)?.fill?.bg());
		assert.true(isHighlighted(), "The formula applies the fill");
		const formula = rule.aRuleElements[0].getFormula(ws);
		assert.true(formula.parent?.ws === ws, "The formula invalidates the owning worksheet");
		assert.true(formula.isInDependencies, "The formula is registered");

		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		assert.false(formula.isInDependencies, "Deleting the rule unregisters the formula");
		assert.false(isHighlighted(), "Deleting the rule clears the fill");

		history.Undo();
		assert.true(isHighlighted(), "Undo restores the fill");
		assert.true(rule.aRuleElements[0].getFormula(ws) === formula, "Undo keeps the parsed formula");
		assert.true(formula.isInDependencies, "Undo restores the dependency");

		const copy = new AscCommonExcel.Worksheet(wb, 1);
		copy.copyFrom(ws, "Copy");
		assert.true(copy.getCFRuleById(rule.id) !== rule, "The copy owns a separate rule");
		copy.deleteCFRule(rule.id);
		assert.true(formula.isInDependencies, "Deleting the copy does not unregister the original formula");

		source.setValue("30");
		assert.false(isHighlighted(), "Changing A1 recalculates the restored rule in B2");

		ws.deleteCFRule(rule.id);
		clearData(0, 0, 0, 6);
	});

	QUnit.test("Deleting a resized formula rule removes its original subscription", function (assert) {
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=A1>50");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(getRgbColor("FFC7CE"));
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);
		ws.getCompiledStyle(1, 1);

		const formula = rule.aRuleElements[0].getFormula(ws);
		const graph = wb.dependencyFormulas;
		const hasSubscription = () => Object.values(graph.sheetListeners[ws.getId()]?.cellMap || {})
			.some(entry => entry.listeners[formula.getListenerId()] === formula);
		assert.true(hasSubscription(), "A1 has the original single-cell listener");

		rule.setLocation([new Asc.Range(1, 1, 1, 5)], ws, false);
		ws.deleteCFRule(rule.id);
		assert.false(hasSubscription(), "Deleting the resized rule removes the original listener");
		clearData(0, 0, 0, 6);
	});

	QUnit.test("Copying a rule keeps the reference to a renamed worksheet", function (assert) {
		const source = wb.createWorksheet(wb.getWorksheetCount(), "Reference");
		source.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=Reference!$A$1>50");
		const dxf = new Asc.asc_CellXfs();
		const fillColor = getRgbColor("FFC7CE");
		dxf.asc_setFillColor(fillColor);
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);
		assert.true(compareAscColorAndRgbColor(fillColor, ws.getCompiledStyle(1, 1)?.fill?.bg()));

		source.setName("Renamed");
		wb.copyWorksheet(ws.getIndex(), null, "Copy");
		const copy = wb.getWorksheet(wb.getWorksheetCount() - 1);
		const copiedRule = copy.getCFRuleById(rule.id);
		assert.strictEqual(copiedRule.aRuleElements[0].Text, "Renamed!$A$1>50");
		copy.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(copiedRule.ranges));
		assert.true(compareAscColorAndRgbColor(fillColor, copy.getCompiledStyle(1, 1)?.fill?.bg()));
		const formula = copiedRule.aRuleElements[0].getFormula(copy);
		assert.true(formula.isInDependencies, "The copied formula is registered");

		const history = AscCommon.History;
		history.Create_NewPoint();
		wb.removeWorksheet(copy.getIndex());
		assert.false(formula.isInDependencies, "Removing the copy unregisters its formula");
		history.Undo();
		assert.true(formula.isInDependencies, "Undo restores the copied formula subscription");
		history.Redo();
		assert.false(formula.isInDependencies, "Redo unregisters the copied formula");

		ws.deleteCFRule(rule.id);
		wb.removeWorksheet(source.getIndex());
		clearData(0, 0, 0, 6);
	});

	QUnit.test("Redoing row insertion updates a formula rule reference", function (assert) {
		ws.getRange4(1, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2:$B$6");
		rule.asc_setValue1("=$A$2>50");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(getRgbColor("FFC7CE"));
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);
		ws.getCompiledStyle(1, 1);
		const formula = rule.aRuleElements[0].getFormula(ws);

		const history = AscCommon.History;
		history.Create_NewPoint();
		history.StartTransaction();
		ws.insertRowsBefore(1, 1);
		history.EndTransaction();
		assert.strictEqual(formula.getFormula(), "$A$3>50", "Insertion shifts the formula");
		history.Undo();
		assert.strictEqual(formula.getFormula(), "$A$2>50", "Undo shifts the formula back");
		history.Redo();
		assert.strictEqual(ws.getRange4(2, 0).getValue(), "70", "Redo inserts the row again");
		assert.strictEqual(formula.getFormula(), "$A$3>50", "Redo shifts the formula again");

		ws.deleteCFRule(rule.id);
		clearData(0, 0, 0, 6);
	});

	QUnit.test("Replacing a formula rule updates its dependencies in history", function (assert) {
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(0, 2).setValue("30");
		ws.getRange4(1, 1).setValue("1");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=$A$1>50");
		const fillColor = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fillColor);
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);

		const isHighlighted = () => compareAscColorAndRgbColor(fillColor, ws.getCompiledStyle(1, 1)?.fill?.bg());
		assert.true(isHighlighted());
		const original = rule.aRuleElements[0].getFormula(ws);
		const edited = rule.clone();
		edited.asc_setValue1("=$C$1>50");

		const history = AscCommon.History;
		history.Create_NewPoint();
		ws.changeCFRule(rule, edited, true);
		assert.false(isHighlighted(), "The edited formula changes the fill");
		const replacement = rule.aRuleElements[0].getFormula(ws);
		assert.false(original.isInDependencies, "The old formula is unregistered");
		assert.true(replacement.isInDependencies, "The new formula is registered");

		history.Undo();
		assert.true(isHighlighted(), "Undo restores the original fill");
		assert.true(original.isInDependencies, "Undo restores the original formula subscription");
		assert.false(replacement.isInDependencies, "Undo unregisters the replacement");
		history.Redo();
		assert.false(isHighlighted(), "Redo restores the edited fill");
		assert.false(original.isInDependencies, "Redo unregisters the original formula");

		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 6);
	});

	QUnit.test("Collaborative undo turns a rule deletion into a serializable addition", function (assert) {
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2:$B$6");
		rule.asc_setValue1("=$A$1>50");
		ws.addCFRule(rule);

		const history = AscCommon.History;
		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		const deletion = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete);
		const reverse = history.Item_ToSerializable(deletion).CreateReverseChange();
		assert.ok(reverse, "A synchronized deletion has a reverse operation");
		if (reverse) {
			assert.strictEqual(reverse.nActionType, AscCH.historyitem_Worksheet_CFRuleAdd);
			assert.strictEqual(reverse.oData.id, rule.id);
			const restoredRule = reverse.oData.getProperty(reverse.oData.Properties.to);
			assert.notStrictEqual(restoredRule, rule, "Restoration has its own rule snapshot");
			assert.strictEqual(restoredRule.aRuleElements[0].Text, "$A$1>50");
			const worksheetHistory = AscCommonExcel.g_oUndoRedoWorksheet;
			assert.true(worksheetHistory.CommuteRelated(reverse, {
				oClass: worksheetHistory,
				nActionType: AscCH.historyitem_Worksheet_RemoveCell,
			}), "A cell value change does not invalidate restoration");
			assert.false(worksheetHistory.CommuteRelated(reverse, {
				oClass: worksheetHistory,
				nActionType: AscCH.historyitem_Worksheet_AddRows,
			}), "Restoration does not apply stale references after a structural edit");
			reverse.Redo();
			assert.strictEqual(ws.getCFRuleById(rule.id), restoredRule);
			ws.deleteCFRule(rule.id);
		}
	});

	QUnit.test("Collaborative undo does not restore a rule changed by another user after deletion", function (assert) {
		const history = AscCommon.History;
		const worksheetHistory = AscCommonExcel.g_oUndoRedoWorksheet;
		const ruleHistory = AscCommonExcel.g_oUndoRedoCF;
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2:$B$6");
		rule.asc_setValue1("=$A$1>50");
		rule.priority = 2;
		ws.addCFRule(rule);

		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		const deletion = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete);
		const sheetId = ws.getId();
		const edited = rule.clone();
		edited.asc_setValue1("=$C$1>50");
		const range = new AscCommonExcel.UndoRedoData_BBox(new Asc.Range(3, 1, 3, 5));
		const scenarios = [
			[ruleHistory, AscCH.historyitem_CFRule_SetRuleElements, sheetId, rule.id, edited.aRuleElements,
				false, "A concurrent formula edit blocks restoration"],
			[ruleHistory, AscCH.historyitem_CFRule_SetRanges, sheetId, rule.id, [range],
				false, "A concurrent range edit blocks restoration"],
			[worksheetHistory, AscCH.historyitem_Worksheet_CFRuleDelete, sheetId, rule.id, null,
				false, "A concurrent deletion blocks restoration"],
			[ruleHistory, AscCH.historyitem_CFRule_SetRuleElements, sheetId, edited.id, edited.aRuleElements,
				true, "A change to another rule does not block restoration"],
			[ruleHistory, AscCH.historyitem_CFRule_SetPriority, sheetId, edited.id, rule.priority,
				false, "A later rule that takes the deleted rule's priority blocks restoration"],
			[ruleHistory, AscCH.historyitem_CFRule_SetRanges, sheetId + "-copy", rule.id, [range],
				true, "A change to the same rule ID on a different sheet does not block restoration"],
			[worksheetHistory, AscCH.historyitem_Worksheet_CFRuleAdd, sheetId + "-copy", edited.id, edited,
				true, "An addition on a different sheet does not block restoration"],
		];
		for (const [changeClass, type, editSheetId, editRuleId, value, canRestore, message] of scenarios) {
			const collaborativeHistory = new AscCommon.CCollaborativeHistory(api.collaborativeEditing);
			collaborativeHistory.AddOwnChanges([history.Item_ToSerializable(deletion)], null);
			const data = new AscCommonExcel.UndoRedoData_CF(editRuleId, null, value);
			const change = new AscCommonExcel.UndoRedoItemSerializable(changeClass, type, editSheetId, null, data);
			change.Redo();
			collaborativeHistory.AddChange(change);

			const reverse = collaborativeHistory.GetReverseOwnChanges();
			assert.strictEqual(reverse.length, canRestore ? 1 : 0, message);
		}
	});

	QUnit.test("Collaborative undo sends the current formula of a formatting threshold", function (assert) {
		ws.getRange4(0, 0).setValue("70");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.dataBar);
		rule.asc_setLocation("$B$2:$B$6");
		const bar = new AscCommonExcel.CDataBar();
		const threshold = new AscCommonExcel.CConditionalFormatValueObject();
		threshold.asc_setType(AscCommonExcel.ECfvoType.Formula);
		threshold.asc_setVal("=$A$1");
		bar.aCFVOs.push(threshold);
		rule.asc_setColorScaleOrDataBarOrIconSetRule(bar);
		ws.addCFRule(rule);
		assert.strictEqual(rule.getMin([], ws), 70);
		ws.insertRowsBefore(0, 1);
		assert.strictEqual(threshold.formula.getFormula(ws).getFormulaRaw(), "$A$2");

		const history = AscCommon.History;
		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		const deletion = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete);
		const reverse = history.Item_ToSerializable(deletion).CreateReverseChange();
		const writer = new AscCommon.CMemory();
		reverse.Serialize(writer, api.collaborativeEditing);
		const received = new AscCommonExcel.UndoRedoItemSerializable();
		received.Deserialize(new AscCommon.FT_Stream2(writer.GetData(), writer.GetCurPosition()));
		assert.strictEqual(received.oData.to.aRuleElements[0].aCFVOs[0].Val, "$A$2");
		assert.strictEqual(received.oData.to.getMin([], ws), 70);
		received.oData.to.removeFormulaDependencies();
		clearData(0, 0, 0, 7);
	});

	QUnit.test("Editing a shifted data bar threshold uses its current formula", function (assert) {
		ws.getRange4(0, 0).setValue("70");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.dataBar);
		rule.asc_setLocation("$B$2:$B$6");
		const bar = new AscCommonExcel.CDataBar();
		const threshold = new AscCommonExcel.CConditionalFormatValueObject();
		threshold.asc_setType(AscCommonExcel.ECfvoType.Formula);
		threshold.asc_setVal("=$A$1");
		bar.aCFVOs.push(threshold);
		rule.asc_setColorScaleOrDataBarOrIconSetRule(bar);
		ws.addCFRule(rule);

		ws.insertRowsBefore(0, 1);
		const shiftedFormula = threshold.formula;
		assert.strictEqual(threshold.getValueText(), "$A$2");

		threshold.asc_setVal("=$A$2");
		assert.strictEqual(threshold.formula, shiftedFormula, "Setting the current value preserves the parsed formula");
		ws.insertRowsBefore(0, 1);
		assert.strictEqual(threshold.getValueText(), "$A$3", "The threshold follows the next insertion");

		ws.getRange4(0, 0).setValue("30");
		threshold.asc_setVal("=$A$1");
		assert.strictEqual(rule.getMin([], ws), 30);
		ws.insertRowsBefore(0, 1);
		ws.getRange4(0, 0).setValue("50");
		threshold.asc_setVal("=$A$1");
		assert.strictEqual(threshold.getValueText(), "$A$1", "The original reference can be entered again");
		assert.strictEqual(rule.getMin([], ws), 50);

		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 9);
	});

	for (const type of ["Number", "Percent"]) {
		QUnit.test(type + " thresholds follow row shifts before calculation and after undo", function (assert) {
			const rule = new AscCommonExcel.CConditionalFormattingRule();
			rule.asc_setType(Asc.c_oAscCFType.dataBar);
			rule.asc_setLocation("$B$3:$B$6");
			const bar = new AscCommonExcel.CDataBar();
			const threshold = new AscCommonExcel.CConditionalFormatValueObject();
			threshold.asc_setType(AscCommonExcel.ECfvoType[type]);
			threshold.asc_setVal("=$A$2");
			bar.aCFVOs.push(threshold);
			rule.asc_setColorScaleOrDataBarOrIconSetRule(bar);
			ws.addCFRule(rule);

			ws.insertRowsBefore(0, 1);
			assert.strictEqual(threshold.getValueText(), "$A$3", "An uncalculated threshold follows the inserted row");
			threshold.asc_setVal("=$A$3");
			rule.getMin([0, 100], ws);

			ws.removeRows(0, 0);
			assert.strictEqual(threshold.getValueText(), "$A$2", "Deleting the row shifts the reference");
			AscCommon.History.Undo();
			const restored = rule.aRuleElements[0].aCFVOs[0];
			assert.strictEqual(restored.getValueText(), "$A$3", "Undo restores the reference");
			ws.insertRowsBefore(0, 1);
			assert.strictEqual(restored.getValueText(), "$A$4", "The restored reference follows the next insertion");

			ws.deleteCFRule(rule.id);
		});
	}

	QUnit.test("Copying a sheet registers formatting formulas before the copy is shown", function (assert) {
		const reference = wb.createWorksheet(wb.getWorksheetCount(), "Reference");
		reference.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("90");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.cellIs);
		rule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=Reference!$A$1");
		const fill = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fill);
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The source rule applies");

		const index = wb.getWorksheetCount();
		const event = new AscCommonExcel.UndoRedoData_SheetAdd(index, "Copy", ws.getId(), "copied-sheet", []);
		AscCommonExcel.g_oUndoRedoWorkbook.Redo(AscCH.historyitem_Workbook_SheetAdd, event, null);
		const copy = wb.getWorksheet(index);
		reference.setName("Renamed");
		const copiedRule = copy.getCFRuleById(rule.id);
		assert.strictEqual(copiedRule.aRuleElements[0].getFormula(copy).getFormulaRaw(), "Renamed!$A$1");
		copy.setDirtyConditionalFormatting();
		assert.true(compareAscColorAndRgbColor(fill, copy.getCompiledStyle(1, 1)?.fill?.bg()), "The copied rule applies");

		wb.removeWorksheet(copy.getIndex());
		ws.deleteCFRule(rule.id);
		wb.removeWorksheet(reference.getIndex());
		clearData(0, 1, 0, 1);
	});

	QUnit.test("A formatting formula can resolve a worksheet added later", function (assert) {
		ws.getRange4(1, 1).setValue("90");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.cellIs);
		rule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=Later!$A$1");
		const fill = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fill);
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);

		const index = wb.getWorksheetCount();
		wb.copyWorksheet(ws.getIndex(), index, "Imported", undefined, true, undefined, ws);
		const copy = wb.getWorksheet(index);
		const reference = wb.createWorksheet(wb.getWorksheetCount(), "Later");
		reference.getRange4(0, 0).setValue("70");
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The reference resolves");
		copy.setDirtyConditionalFormatting();
		assert.true(compareAscColorAndRgbColor(fill, copy.getCompiledStyle(1, 1)?.fill?.bg()), "The copied reference resolves");

		wb.removeWorksheet(copy.getIndex());
		ws.deleteCFRule(rule.id);
		wb.removeWorksheet(reference.getIndex());
		clearData(1, 1, 1, 1);
	});

	QUnit.test("Collaborative undo preserves the priority of a rule added after deletion", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		function addRule(id, color) {
			const rule = new AscCommonExcel.CConditionalFormattingRule();
			rule.id = id;
			rule.asc_setType(Asc.c_oAscCFType.expression);
			rule.asc_setLocation("$B$2");
			rule.asc_setValue1("=$A$1>50");
			const dxf = new Asc.asc_CellXfs();
			dxf.asc_setFillColor(color);
			rule.asc_setDxf(dxf);
			api.asc_setCF([rule]);
			return rule;
		}

		const deleted = addRule("owner-a_1", getRgbColor("FF0000"));
		history.Create_NewPoint();
		ws.deleteCFRule(deleted.id, true);
		const deletion = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete);
		const collaborativeHistory = new AscCommon.CCollaborativeHistory(api.collaborativeEditing);
		collaborativeHistory.AddOwnChanges([history.Item_ToSerializable(deletion)], null);
		const green = getRgbColor("00FF00");
		const added = addRule("owner-b_1", green);
		const addition = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleAdd);
		const wire = wb._SerializeHistoryItem(new AscCommon.CMemory(), history.Item_ToSerializable(addition));
		collaborativeHistory.AddChange(wb._DeserializeUndoRedoElems([wire])[0]);
		assert.true(compareAscColorAndRgbColor(green, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The new rule applies");

		for (const change of collaborativeHistory.GetReverseOwnChanges()) {
			change.Redo();
		}
		assert.notOk(ws.getCFRuleById(deleted.id), "Undo skips restoration when the rule order has changed");
		assert.strictEqual(added.priority, 1, "The new rule keeps its priority");
		assert.true(compareAscColorAndRgbColor(green, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The new rule still applies");

		ws.deleteCFRule(deleted.id);
		ws.deleteCFRule(added.id);
		clearData(0, 1, 0, 1);
	});

	QUnit.test("Restored formatting formulas follow row insertion in the same history batch", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=$A$1>50");
		const color = getRgbColor("FF0000");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(color);
		rule.asc_setDxf(dxf);
		ws.addCFRule(rule);
		assert.true(compareAscColorAndRgbColor(color, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The rule applies in B2");

		const thresholdRule = new AscCommonExcel.CConditionalFormattingRule();
		thresholdRule.asc_setType(Asc.c_oAscCFType.dataBar);
		thresholdRule.asc_setLocation("$C$2");
		const bar = new AscCommonExcel.CDataBar();
		const threshold = new AscCommonExcel.CConditionalFormatValueObject();
		threshold.asc_setType(AscCommonExcel.ECfvoType.Formula);
		threshold.asc_setVal("=$A$1");
		bar.aCFVOs.push(threshold);
		thresholdRule.asc_setColorScaleOrDataBarOrIconSetRule(bar);
		ws.addCFRule(thresholdRule);
		assert.strictEqual(thresholdRule.getMin([], ws), 70);

		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		ws.deleteCFRule(thresholdRule.id, true);
		const restore = history.Points[history.Index].Items
			.filter(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete)
			.map(item => history.Item_ToSerializable(item).CreateReverseChange());
		const addRows = new AscCommonExcel.UndoRedoItemSerializable(AscCommonExcel.g_oUndoRedoWorksheet,
			AscCH.historyitem_Worksheet_AddRows, ws.getId(), new Asc.Range(0, 0, 0, 0),
			new AscCommonExcel.UndoRedoData_FromToRowCol(true, 0, 0));
		const wire = [...restore, addRows].map(item => wb._SerializeHistoryItem(new AscCommon.CMemory(), item));
		history.Clear();
		wb.snapshot = wb._getSnapshot();
		wb.DeserializeHistory(wire);

		const restored = ws.getCFRuleById(rule.id);
		assert.strictEqual(restored.ranges[0].getName(), "B3", "The range shifts to B3");
		assert.strictEqual(restored.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$2>50", "The formula follows A2");
		assert.true(compareAscColorAndRgbColor(color, ws.getCompiledStyle(2, 1)?.fill?.bg()), "The rule applies in B3");
		assert.strictEqual(ws.getCFRuleById(thresholdRule.id).getMin([], ws), 70, "The data bar threshold follows A2");

		const reloaded = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, false);
		const file = new AscCommonExcel.BinaryFileWriter(wb).Write();
		new AscCommonExcel.BinaryFileReader().Read(file, reloaded);
		const reloadedRules = Object.values(reloaded.getWorksheet(0).getConditionalFormattingRules());
		const expression = reloadedRules.find(item => item.type === Asc.c_oAscCFType.expression);
		const dataBar = reloadedRules.find(item => item.type === Asc.c_oAscCFType.dataBar);
		assert.strictEqual(expression.aRuleElements[0].Text, "$A$2>50", "The saved formula follows A2");
		assert.strictEqual(dataBar.aRuleElements[0].aCFVOs[0].Val, "$A$2", "The saved threshold follows A2");

		const json = new window.AscJsonConverter.WriterToJSON();
		const thresholdValue = ws.getCFRuleById(thresholdRule.id).aRuleElements[0].aCFVOs[0];
		assert.strictEqual(json.SerFormulaCF(restored.aRuleElements[0]).formula, "$A$2>50", "The exported formula follows A2");
		assert.strictEqual(json.SerCondFmtValObj(thresholdValue).val, "$A$2", "The exported threshold follows A2");

		ws.deleteCFRule(rule.id);
		ws.deleteCFRule(thresholdRule.id);
		clearData(0, 2, 0, 2);
	});

	QUnit.test("Public API modifies a formula after inserting a row", function (assert) {
		const Api = AscBuilder.Cell.Api;
		const sheet = Api.GetActiveSheet();
		sheet.GetRange("A1").SetValue(70);
		sheet.GetRange("B2").SetValue(90);
		const condition = sheet.GetRange("B2").GetFormatConditions().Add("xlCellValue", "xlGreater", "=$A$1");
		condition.SetFillColor(Api.CreateColorFromRGB(255,199,206));
		sheet.GetRows(1).Insert("down");
		sheet.GetRange("A1").SetValue(100);
		condition.Modify("xlCellValue", "xlGreater", "$A$1");
		const rule = condition.rule;
		const fill = getRgbColor("FFC7CE");
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$1", "The rule uses the new reference");
		assert.false(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(2, 1)?.fill?.bg()), "B3 is not formatted after threshold becomes 100");
		ws.deleteCFRule(rule.id);
		clearData(0, 1, 0, 3);
	});

	QUnit.test("Editing only a rule's color keeps its reference through a second row insert", function (assert) {
		const Api = AscBuilder.Cell.Api;
		const sheet = Api.GetActiveSheet();
		sheet.GetRange("A1").SetValue(70);
		sheet.GetRange("B2").SetValue(90);
		const condition = sheet.GetRange("B2").GetFormatConditions().Add("xlCellValue", "xlGreater", "=$A$1");
		condition.SetFillColor(Api.CreateColorFromRGB(255,199,206));
		ws.setDirtyConditionalFormatting();
		ws.getCompiledStyle(1, 1);
		sheet.GetRows(1).Insert("down");
		condition.SetFillColor(Api.CreateColorFromRGB(0,255,0));
		sheet.GetRows(1).Insert("down");
		sheet.GetRange("A2").SetValue(100);
		const rule = condition.rule;
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$3", "The formula still follows the shifted reference");
		assert.true(compareAscColorAndRgbColor(getRgbColor("00FF00"), ws.getCompiledStyle(3, 1)?.fill?.bg()), "B4 uses the edited fill");
		ws.deleteCFRule(rule.id);
		clearData(0, 1, 0, 4);
	});

	QUnit.test("Edited rule follows row insertion after fast undo in the same history batch", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("1");
		ws.getRange4(0, 2).setValue("90");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.expression);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=$A$1>50");
		const dxf = new Asc.asc_CellXfs();
		const color = getRgbColor("FF0000");
		dxf.asc_setFillColor(color);
		rule.asc_setDxf(dxf);
		api.asc_setCF([rule]);
		assert.true(compareAscColorAndRgbColor(color, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The rule applies in B2");
		history.Create_NewPoint();
		ws.deleteCFRule(rule.id, true);
		const deletion = history.Points[history.Index].Items.find(item => item.Type === AscCH.historyitem_Worksheet_CFRuleDelete);
		const reverse = history.Item_ToSerializable(deletion).CreateReverseChange();
		assert.ok(reverse, "Fast undo creates a restoration");
		if (!reverse) {
			return;
		}
		const addRows = new AscCommonExcel.UndoRedoItemSerializable(AscCommonExcel.g_oUndoRedoWorksheet,
			AscCH.historyitem_Worksheet_AddRows, ws.getId(), new Asc.Range(0, 0, 0, 0),
			new AscCommonExcel.UndoRedoData_FromToRowCol(true, 0, 0));
		history.Clear();
		wb.snapshot = wb._getSnapshot();
		const replacement = rule.clone();
		replacement.asc_setValue1("=$C$1>50");
		const edit = new AscCommonExcel.UndoRedoItemSerializable(AscCommonExcel.g_oUndoRedoCF,
			AscCH.historyitem_CFRule_SetRuleElements, ws.getId(), rule.getUnionRange(),
			new AscCommonExcel.UndoRedoData_CF(rule.id, null, replacement.aRuleElements));
		const wire = [reverse, edit, addRows].map(item => wb._SerializeHistoryItem(new AscCommon.CMemory(), item));
		wb.DeserializeHistory(wire);
		const restored = ws.getCFRuleById(rule.id);
		assert.strictEqual(restored.ranges[0].getName(), "B3", "Applies-to range shifts to B3");
		const f = restored.aRuleElements[0].getFormula(ws).getFormulaRaw();
		assert.strictEqual(f, "$C$2>50", "The edited formula shifts with the source cell");
		assert.true(compareAscColorAndRgbColor(color, ws.getCompiledStyle(2, 1)?.fill?.bg()), "The rule still applies to B3");
		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 3);
	});


	QUnit.test("Undo restores a removed rule with its original reference", function (assert) {
		const history = AscCommon.History;
		const fill = getRgbColor("FFC7CE");
		const scenarios = [
			["A3", 2, 0, 1, 0, () => ws.removeRows(1, 1, false)],
			["C2", 1, 2, 1, 0, () => ws.removeCols(1, 1)],
			["B3", 2, 1, 0, 1, () => ws.getRange4(1, 1).deleteCellsShiftUp()],
			["C2", 1, 2, 1, 0, () => ws.getRange4(1, 1).deleteCellsShiftLeft()],
		];
		for (const [reference, row, col, otherRow, otherCol, remove] of scenarios) {
			ws.getRange4(row, col).setValue("70");
			ws.getRange4(otherRow, otherCol).setValue("100");
			ws.getRange4(1, 1).setValue("90");
			const rule = new AscCommonExcel.CConditionalFormattingRule();
			rule.asc_setType(Asc.c_oAscCFType.cellIs);
			rule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
			rule.asc_setLocation("$B$2");
			const address = "$" + reference[0] + "$" + reference[1];
			rule.asc_setValue1("=" + address);
			const dxf = new Asc.asc_CellXfs();
			dxf.asc_setFillColor(fill);
			rule.asc_setDxf(dxf);
			ws.addCFRule(rule);
			assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()), "The rule applies");

			history.Create_NewPoint();
			history.StartTransaction();
			remove();
			history.EndTransaction();
			assert.notOk(ws.getCFRuleById(rule.id), "Removing the range deletes the rule");
			history.Undo();
			const restored = ws.getCFRuleById(rule.id);
			assert.strictEqual(restored?.aRuleElements[0].getFormula(ws).getFormulaRaw(), address, "Undo restores the reference");
			assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()), "Undo restores the fill");
			history.Redo();
			assert.notOk(ws.getCFRuleById(rule.id), "Redo removes the rule again");

			clearData(0, 2, 0, 3);
		}
	});

	QUnit.test("Undo type changes and immediately insert row", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(1, 1).setValue("90");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.cellIs);
		rule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=$A$1");
		const fill = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fill);
		rule.asc_setDxf(dxf);
		api.asc_setCF([rule]);
		ws.setDirtyConditionalFormatting();
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()));
		const edit = rule.clone();
		edit.id = rule.id;
		edit.asc_setType(Asc.c_oAscCFType.containsErrors);
		api.asc_setCF([edit]);
		assert.strictEqual(rule.type, Asc.c_oAscCFType.containsErrors);
		history.Undo();
		assert.strictEqual(rule.type, Asc.c_oAscCFType.cellIs);
		ws.insertRowsBefore(0, 1);
		ws.getRange4(0, 0).setValue("100");
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$2");
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(2, 1)?.fill?.bg()));
		ws.deleteCFRule(rule.id);
		clearData(0, 1, 0, 3);
	});

	QUnit.test("Undo replacing a text condition formula with a literal then shift before paint", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("pass");
		ws.getRange4(1, 1).setValue("pass");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.containsText);
		rule.asc_setLocation("$B$2");
		rule.asc_setContainsText("=$A$1");
		const fill = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fill);
		rule.asc_setDxf(dxf);
		api.asc_setCF([rule]);
		ws.setDirtyConditionalFormatting();
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(1, 1)?.fill?.bg()), "original formula condition applies");
		const edit = rule.clone();
		edit.id = rule.id;
		edit.asc_setContainsText("different");
		api.asc_setCF([edit]);
		assert.strictEqual(rule.text, "different");
		history.Undo();
		assert.strictEqual(rule.text, null);
		ws.insertRowsBefore(0, 1);
		ws.getRange4(0, 0).setValue("different");
		assert.true(rule.aRuleElements[0].getFormula(ws).getFormulaRaw().includes("$A$2"), "restored formula follows the moved threshold");
		assert.true(compareAscColorAndRgbColor(fill, ws.getCompiledStyle(2, 1)?.fill?.bg()), "restored formula condition still applies");
		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 3);
	});

	QUnit.test("Undo and redo keep subscriptions when switching between a formula and a data bar", function (assert) {
		const history = AscCommon.History;
		ws.getRange4(0, 0).setValue("70");
		ws.getRange4(0,2).setValue("40");
		ws.getRange4(1, 1).setValue("90");
		const rule = new AscCommonExcel.CConditionalFormattingRule();
		rule.asc_setType(Asc.c_oAscCFType.cellIs);
		rule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		rule.asc_setLocation("$B$2");
		rule.asc_setValue1("=$A$1");
		const fill = getRgbColor("FFC7CE");
		const dxf = new Asc.asc_CellXfs();
		dxf.asc_setFillColor(fill);
		rule.asc_setDxf(dxf);
		api.asc_setCF([rule]);
		const edit = rule.clone();
		edit.id = rule.id;
		edit.asc_setType(Asc.c_oAscCFType.dataBar);
		const bar = new AscCommonExcel.CDataBar();
		const threshold = new AscCommonExcel.CConditionalFormatValueObject();
		threshold.asc_setType(AscCommonExcel.ECfvoType.Formula);
		threshold.asc_setVal("=$C$1");
		bar.aCFVOs.push(threshold);
		edit.asc_setColorScaleOrDataBarOrIconSetRule(bar);
		api.asc_setCF([edit]);
		assert.strictEqual(rule.type, Asc.c_oAscCFType.dataBar);
		history.Undo();
		assert.strictEqual(rule.type, Asc.c_oAscCFType.cellIs);
		assert.true(rule.getFormulaParsers().some(f => f.isInDependencies && f.getFormulaRaw() === "$A$1"));
		history.Redo();
		assert.strictEqual(rule.type, Asc.c_oAscCFType.dataBar);
		assert.true(rule.getFormulaParsers().some(f => f.isInDependencies && f.getFormulaRaw() === "$C$1"));
		ws.insertRowsBefore(0, 1);
		ws.getRange4(0,2).setValue("80");
		assert.strictEqual(rule.getMin([],ws),40);
		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 3);
	});

	QUnit.test("Undo partial range deletion restores original threshold reference", function (assert) {
		const Api = AscBuilder.Cell.Api;
		const sheet = Api.GetActiveSheet();
		sheet.GetRange("A2").SetValue(70);
		sheet.GetRange("B2").SetValue(90);
		sheet.GetRange("B3").SetValue(95);
		const condition = sheet.GetRange("B2:B3").GetFormatConditions().Add("xlCellValue", "xlGreater", "=$A$2");
		condition.SetFillColor(Api.CreateColorFromRGB(255,199,206));
		const rule = condition.rule;
		sheet.GetRows(2).Delete("up");
		assert.ok(ws.getCFRuleById(rule.id), "Partially covered rule remains after deletion");
		const afterDeletion = rule.aRuleElements[0].getFormula(ws).getFormulaRaw();
		AscCommon.History.Undo();
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$2");
		ws.setDirtyConditionalFormatting();
		assert.true(compareAscColorAndRgbColor(getRgbColor("FFC7CE"), ws.getCompiledStyle(1, 1)?.fill?.bg()));
		AscCommon.History.Redo();
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), afterDeletion);
		AscCommon.History.Undo();
		assert.strictEqual(rule.aRuleElements[0].getFormula(ws).getFormulaRaw(), "$A$2");
		ws.deleteCFRule(rule.id);
		clearData(0, 2, 0, 3);
	});

	QUnit.module("Conditional formatting");
});
