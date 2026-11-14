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
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
		return 1;
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function () {
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	api.wbModel.handlers.add("getSelectionState", function () {
		return null;
	});

	var ws = api.wbModel.aWorksheets[0];
	var wsView = api.wb.getWorksheet();
	var styles = Asc.c_oAscBorderStyles;

	function prop(style, r, g, b) {
		var res = new AscCommonExcel.BorderProp();
		res.setStyle(style);
		res.c = AscCommonExcel.createRgbColor(r || 0, g || 0, b || 0);
		return res;
	}

	function setBorder(ref, sides) {
		var border = new AscCommonExcel.Border();
		for (var name in sides) {
			if (sides.hasOwnProperty(name)) {
				border[name] = sides[name];
			}
		}
		ws.getRange2(ref).setBorderSrc(border);
	}

	function select(ref, activeRef) {
		var bbox = ws.getRange2(ref).getBBox0();
		var active = ws.getRange2(activeRef || ref).getBBox0();
		ws.selectionRange.ranges = [bbox.clone()];
		ws.selectionRange.setActiveCell(active.r1, active.c1);
		return wsView._getSelectionInfoCell();
	}

	// The default sizes WorksheetView._init derives from the fonts (stubbed here): hidden rows and columns are the
	// ones laid out with a zero size, as in the editor.
	wsView.defaultColWidthPx = 64;
	AscCommonExcel.oDefaultMetrics.RowHeight = 15;

	function layout() {
		wsView._calcWidthColumns(AscCommonExcel.recalcType.full);
		wsView._calcHeightRows(AscCommonExcel.recalcType.full);
	}

	// Style of a side reported by asc_getBorders(), None when the side is absent.
	function styleOf(info, getter) {
		var side = info.asc_getBorders()[getter]();
		return side ? side.asc_getStyle() : styles.None;
	}

	function colorOf(info, getter) {
		var color = info.asc_getBorders()[getter]().asc_getColor();
		return [color.asc_getR(), color.asc_getG(), color.asc_getB()];
	}

	function clean() {
		ws.getRange2("A1:J20").cleanAll();
		ws.getRange2("A1:J20").unmerge();
		ws.setColHidden(false, 0, 9);
		ws.setRowHidden(false, 0, 19);
		layout();
	}

	QUnit.module("Selection info borders");

	QUnit.test("Own sides and an empty cell", function (assert) {
		clean();
		setBorder("C3", {l: prop(styles.Thin), t: prop(styles.Medium), r: prop(styles.Dashed), b: prop(styles.Double)});
		var info = select("C3");
		assert.strictEqual(styleOf(info, "asc_getLeft"), styles.Thin, "left");
		assert.strictEqual(styleOf(info, "asc_getTop"), styles.Medium, "top");
		assert.strictEqual(styleOf(info, "asc_getRight"), styles.Dashed, "right");
		assert.strictEqual(styleOf(info, "asc_getBottom"), styles.Double, "bottom");
		assert.strictEqual(info.asc_getBorders().asc_getDiagDown(), null, "no diagonal down");

		info = select("F6");
		["asc_getLeft", "asc_getTop", "asc_getRight", "asc_getBottom"].forEach(function (getter) {
			assert.strictEqual(info.asc_getBorders()[getter](), null, "empty cell: " + getter);
		});
	});

	QUnit.test("A side drawn by the neighbouring cell", function (assert) {
		clean();
		setBorder("D3", {l: prop(styles.Thick, 255, 0, 0)});
		setBorder("C4", {t: prop(styles.Thin, 0, 0, 255)});
		var info = select("C3");
		assert.strictEqual(styleOf(info, "asc_getRight"), styles.Thick, "right comes from the left side of D3");
		assert.deepEqual(colorOf(info, "asc_getRight"), [255, 0, 0], "with its color");
		assert.strictEqual(styleOf(info, "asc_getBottom"), styles.Thin, "bottom comes from the top side of C4");
	});

	QUnit.test("Both cells set the shared side: the one that is drawn wins", function (assert) {
		clean();
		setBorder("C3", {r: prop(styles.Thin), b: prop(styles.Thick, 0, 0, 0)});
		setBorder("D3", {l: prop(styles.Thick, 0, 128, 0)});
		setBorder("C4", {t: prop(styles.Thin, 255, 0, 0)});
		var info = select("C3");
		assert.strictEqual(styleOf(info, "asc_getRight"), styles.Thick, "the wider neighbour side, not the own thin one");
		assert.deepEqual(colorOf(info, "asc_getRight"), [0, 128, 0], "with the neighbour color");
		assert.strictEqual(styleOf(info, "asc_getBottom"), styles.Thick, "the wider own side");
	});

	QUnit.test("Hidden rows and columns between the cell and its neighbour", function (assert) {
		clean();
		setBorder("C3", {});
		setBorder("D3", {l: prop(styles.Dotted)});
		setBorder("E3", {l: prop(styles.Medium)});
		setBorder("C5", {t: prop(styles.Dashed)});
		setBorder("A3", {r: prop(styles.Thin)});
		ws.setColHidden(true, 3, 3);
		ws.setColHidden(true, 1, 1);
		ws.setRowHidden(true, 3, 3);
		layout();
		assert.strictEqual(wsView._getColumnWidth(3), 0, "column D is hidden in the view");
		var info = select("C3");
		assert.strictEqual(styleOf(info, "asc_getRight"), styles.Medium, "right comes from E3 behind the hidden D");
		assert.strictEqual(styleOf(info, "asc_getBottom"), styles.Dashed, "bottom comes from C5 behind the hidden row 4");
		assert.strictEqual(styleOf(info, "asc_getLeft"), styles.Thin, "left comes from A3 behind the hidden B");
	});

	QUnit.test("Merged area", function (assert) {
		clean();
		ws.getRange2("C3:D4").merge(Asc.c_oAscMergeOptions.Merge);
		setBorder("C3:D4", {l: prop(styles.Thin), t: prop(styles.Thin), r: prop(styles.Thick), b: prop(styles.Thick)});
		var info = select("C3:D4", "C3");
		assert.strictEqual(styleOf(info, "asc_getRight"), styles.Thick, "right of the merged area");
		assert.strictEqual(styleOf(info, "asc_getBottom"), styles.Thick, "bottom of the merged area");
		assert.strictEqual(info.asc_getOneCell(), true, "exactly the merged area is one cell");
		assert.strictEqual(select("C3:E5", "C3").asc_getOneCell(), false, "a range around the merged area is not");
	});

	QUnit.test("Diagonals", function (assert) {
		clean();
		var border = {d: prop(styles.Thin, 128, 0, 128), dd: true, du: false};
		setBorder("C3", border);
		var info = select("C3");
		assert.strictEqual(styleOf(info, "asc_getDiagDown"), styles.Thin, "diagonal down");
		assert.strictEqual(info.asc_getBorders().asc_getDiagUp(), null, "no diagonal up");
	});

	QUnit.test("One cell or a range", function (assert) {
		clean();
		assert.strictEqual(select("C3").asc_getOneCell(), true, "one cell");
		assert.strictEqual(select("C3:E5", "D4").asc_getOneCell(), false, "a range");
		var info = select("C3:E5", "D4");
		assert.ok(info.asc_getBorders(), "a range still reports the borders of its active cell");
		ws.selectionRange.ranges = [ws.getRange2("C3").getBBox0().clone(), ws.getRange2("E5").getBBox0().clone()];
		ws.selectionRange.setActiveCell(4, 4);
		assert.strictEqual(wsView._getSelectionInfoCell().asc_getOneCell(), false, "several ranges");
	});
});
