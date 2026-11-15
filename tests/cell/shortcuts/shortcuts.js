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

"use strict";

QUnit.config.autostart = false;
(function (window)
{
	const {
		tableHotkeyTypes,
		cellEditorHotkeyTypes,
		graphicHotkeyTypes,
		tableEvents,
		graphicEvents,
		ExecuteCellEditorHotkey,
		ExecuteTableHotkey,
		ExecuteGraphicHotkey,
		ExecuteShortcut,
		InitEditor
	} = AscTestShortcut;

	let editor, wb, wbView, ws, wsView, cellEditor;
	InitEditor(function ()
	{
		editor = window["Asc"]["editor"];
		wb = editor.wbModel;
		wbView = editor.wb;
		ws = wb.aWorksheets[0];
		wsView = wbView.getWorksheet();
		cellEditor = wbView.cellEditor;
		QUnit.start();
	});

	function UpdateView()
	{
		wsView._cleanCache(new Asc.Range(0, 0, wsView.cols.length - 1, wsView.rows.length - 1));
		wsView.changeWorksheet("update", {reinitRanges: true});
	}

	function GetCellCacheText(column, row)
	{
		return wsView._getCellTextCache(column, row).state.chars.map((e) => String.fromCharCode(e)).join('');
	}

	function GetParagraphText(paragraph)
	{
		return paragraph.GetText({ParaSeparator : ""});
	}

	function ClearShapeAndAddParagraph(sText)
	{
		const textShape = AddShape(0, 0, 100, 100);
		const txBody = AscFormat.CreateTextBodyFromString(sText, wb.DrawingDocument, textShape);
		textShape.setTxBody(txBody);
		textShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
		const paragraph = txBody.content.Content[0];
		paragraph.SetThisElementCurrent();
		paragraph.MoveCursorToStartPos();
		textShape.recalculate();
		return {shape: textShape, paragraph: paragraph};
	}

	function AddShape(x, y, height, width)
	{
		const shapeTrack = new AscFormat.NewShapeTrack('rect', x, y, wb.theme, null, null, null, 0);
		shapeTrack.track({}, x + width, y + height);
		const shape = shapeTrack.getShape(false, wb.DrawingDocument, null);
		shape.setBDeleted(false);
		shape.setWorksheet(ws);
		shape.setParent(GetDrawingObjects().drawingObjects);
		shape.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
		shape.checkDrawingBaseCoords();

		shape.select(GetDrawingObjects(), 0);
		return shape;
	}

	function EnterText(text)
	{
		const codePoints = text.split('').map((e) => e.charCodeAt());
		wbView.EnterText(codePoints);
	}

	function AddChart()
	{
		ws.getRange4(0, 0).fillData([['1']]);
		const props = editor.asc_getChartSettings(Asc.c_oAscChartTypeSettings.lineNormal);
		props.left = 0;
		props.top = 0;
		props.width = 100;
		props.height = 100;
		return GetDrawingObjects().addChartDrawingObject(props);
	}

	function CreateGroup(arrObjects)
	{
		const controller = GetDrawingObjects();
		controller.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			arrObjects[i].select(controller, 0);
		}

		return controller.createGroup();
	}

	function GetDirectGraphicTextPr()
	{
		return GetDrawingObjects().getParagraphTextPr();
	}

	function GetDirectGraphicParaPr()
	{
		return GetDrawingObjects().getParagraphParaPr();
	}

	function Round(number, amount)
	{
		const power = Math.pow(10, amount);
		return Math.round(number * power) / power;
	}

	function SelectDrawings(arrDrawings)
	{
		const drawingController = GetDrawingObjects();
		drawingController.resetSelection()
		for (let i = 0; i < arrDrawings.length; i += 1)
		{
			arrDrawings[i].select(drawingController, 0);
		}
	}

	function GetDrawingObjects()
	{
		return editor.getGraphicController();
	}

	function FillActiveCell(text)
	{
		const activeCell = ws.selectionRange.activeCell;
		ws.getRange4(activeCell.row, activeCell.col).fillData([[text]]);

	}

	function SetVisibleRange(r1, c1, r2, c2)
	{
		wsView.visibleRange = new Asc.Range(r1, c1, r2, c2);
		wsView.nRowsCount = 49;
		wsView.nColsCount = 49;
	}

	function FillExampleData()
	{
		ws.getRange3(0, 0, 2, 2).fillData([
			['test11', 'test12', 'test13'],
			['1', '2', '3'],
			['1', '2', '3']
		]);

		Select(0, 0, 0, 0, 2, 2);
	}

	function GetRange(c1, r1, c2, r2)
	{
		return new Asc.Range(c1, r1, c2, r2);
	}

	function Select(activeR, activeC, r1, c1, r2, c2)
	{
		CloseCellEditor(true);
		ws.selectionRange.ranges = [GetRange(c1, r1, c2, r2)];
		ws.selectionRange.activeCell = new AscCommon.CellBase(activeR, activeC);
		wbView._updateSelectionInfo();
	}

	function ResetData(r1, c1, r2, c2)
	{
		ws.autoFilters.deleteAutoFilter(GetRange(0, 0, 0, 0));
		while (ws.TableParts.length)
		{
			ws.deleteTablePart(0);
		}
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	}

	function GetCellText(row, column)
	{
		return ws.getRange4(row, column).getValueWithFormat();
	}

	function CloseCellEditor(save)
	{
		wbView.closeCellEditor(!save);
	}

	function CreateTable()
	{
		FillExampleData();
		ws.autoFilters.addAutoFilter('TableStyleMedium2', GetRange(0, 0, 2, 2));
	}

	function CreateSlicer()
	{
		CreateTable();
		editor.asc_insertSlicer(['test11']);
		return GetDrawingObjects().selectedObjects[0];
	}

	function RemoveGraphicObjects()
	{
		GetDrawingObjects().resetSelection();
		GetDrawingObjects().selectAll();
		GetDrawingObjects().remove();
	}

	function GetCellEditorText()
	{
		return AscCommonExcel.getFragmentsText(cellEditor.options.fragments);
	}

	function OpenCellEditor()
	{
		const enterOptions = new AscCommonExcel.CEditorEnterOptions();
		enterOptions.newText = '';
		enterOptions.quickInput = true;
		enterOptions.focus = true;
		wbView._onEditCell(enterOptions);
	}

	function Undo()
	{
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo);
	}

	function Redo()
	{
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo);
	}

	function GetSelectedCellEditorText()
	{
		const fragments = cellEditor.copySelection() || [];
		return AscCommonExcel.getFragmentsText(fragments)
	}

	function CreateWorksheets(names)
	{
		editor.asc_insertWorksheet(names);
	}

	function RemoveWorksheets(indexes)
	{
		editor.asc_deleteWorksheet(indexes);
	}

	QUnit.module('Test cell editor shortcut', {
		afterEach: function ()
		{
			CloseCellEditor();
		}
	});
	QUnit.test('Check actions with text movements', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'World Hello World');

		cellEditor._moveCursor(-2);

		function CheckCursorPosition(expected, description)
		{
			assert.strictEqual(cellEditor.cursorPos, expected, description);
		}

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine);
		CheckCursorPosition(18, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		CheckCursorPosition(18, 'Check move to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftWord);
		CheckCursorPosition(12, 'Check move to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftWord, 1);
		CheckCursorPosition(6, 'Check move to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord, 1);
		CheckCursorPosition(12, 'Check move to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord);
		CheckCursorPosition(18, 'Check move to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord);
		CheckCursorPosition(24, 'Check move to right word');


		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToDownLine);
		CheckCursorPosition(36, 'Check move down');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToUpLine);
		CheckCursorPosition(18, 'Check move up');*/

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndDocument);
		CheckCursorPosition(161, 'Check move to end document');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartDocument);
		CheckCursorPosition(0, 'Check move to start document');

		function CheckSelectedText(expectedText, description)
		{
			assert.strictEqual(GetSelectedCellEditorText(), expectedText, description);
		}

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine);
		CheckSelectedText('Hello World Hello ', 'Select to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftChar);
		CheckSelectedText('Hello World Hello ', 'Select to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord, 1);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord, 1);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello World ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine);
		CheckSelectedText('Hello World Hello ', 'Select to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToDownLine);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToUpLine);
		CheckSelectedText('Hello World Hello ', 'Select up');*/

		cellEditor._moveCursor(-4);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftChar);
		CheckSelectedText('d', 'Select to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord);
		CheckSelectedText('World', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to start content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndDocument);
		CheckSelectedText('', 'Select to end content');

		cellEditor._moveCursor(-2);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to end content');
		cellEditor._moveCursor(-4);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(144, 'Check move to start line');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(144, 'Check move to start line');*/
	});

	QUnit.test('Check remove parts of text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('Hello Hello Hello Hello Hello Hello Hello');
		cellEditor._moveCursor(-4);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeCharBack);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeWordBack);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeWordBack, 1);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

		cellEditor._moveCursor(-2);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontChar);
		assert.strictEqual(GetCellEditorText(), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontWord);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello ', 'Check removing front word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontWord, 1);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello ', 'Check removing front word');
	});

	QUnit.test('Check save and moving from cell', (assert) =>
	{

		function CheckSaveAndMoving(row, column, text)
		{
			assert.false(wbView.getCellEditMode(), 'Check closing cell editor');
			assert.deepEqual([ws.selectionRange.activeCell.row, ws.selectionRange.activeCell.col], [row, column], 'Check active cell');
			assert.strictEqual(GetCellText(5, 5), text, 'Check initial text');
		}

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello1');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveDown);
		CheckSaveAndMoving(6, 5, 'Hello1');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello2');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveUp);
		CheckSaveAndMoving(4, 5, 'Hello2');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello3');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveRight);
		CheckSaveAndMoving(5, 6, 'Hello3');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello4');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveLeft);
		CheckSaveAndMoving(5, 4, 'Hello4');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello5');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.closeWithoutSave);
		CheckSaveAndMoving(5, 5, 'Hello4');
	});

	QUnit.test('Check change text formatting', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello Hello');
		cellEditor.selectAll();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getStrikeout()), 'Check cell editor strikeout format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getStrikeout()), 'Check cell editor strikeout format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getBold()), 'Check cell editor bold format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getBold()), 'Check cell editor bold format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getItalic()), 'Check cell editor italic format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getItalic()), 'Check cell editor italic format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getVerticalAlign() === AscCommon.vertalign_SuperScript), 'Check cell editor superscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getVerticalAlign() === AscCommon.vertalign_Baseline), 'Check cell editor baseline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getVerticalAlign() === AscCommon.vertalign_SuperScript), 'Check cell editor superscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getVerticalAlign() === AscCommon.vertalign_SubScript), 'Check cell editor subscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getVerticalAlign() === AscCommon.vertalign_Baseline), 'Check cell editor baseline format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getUnderline() === Asc.EUnderline.underlineSingle), 'Check cell editor underline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getUnderline() === Asc.EUnderline.underlineNone), 'Check cell editor underline format');

		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check init font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 12), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 14), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 16), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 18), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 20), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 22), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 24), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 26), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 28), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 36), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 48), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 72), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 72), "Check increase font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 48), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 36), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 28), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 26), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 24), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 22), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 20), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 18), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 16), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 14), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 12), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 10), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 9), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 8), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 8), "Check decrease font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check initial value");
	});

	QUnit.test('Check enter text in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), '.', 'Check add separator');
		EnterText('');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 2, 'Check add new line');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 3, 'Check add new line');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 4, 'Check add new line');
	});

	QUnit.test('Check disabling shortcuts', (assert) =>
	{
		OpenCellEditor();
		assert.strictEqual(ExecuteCellEditorHotkey(cellEditorHotkeyTypes.disableScrollLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		assert.strictEqual(ExecuteCellEditorHotkey(cellEditorHotkeyTypes.disableNumLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		AscCommon.AscBrowser.isOpera = true;
		assert.strictEqual(ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		AscCommon.AscBrowser.isOpera = false;
		assert.strictEqual(ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint) & keydownresult_PreventDefault, keydownresult_PreventDefault);
	});

	QUnit.test('Check select all', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		assert.strictEqual(GetSelectedCellEditorText(), 'Hello', 'Check select all');
	});

	QUnit.test('Check undo/redo in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('H');
		EnterText('e');
		EnterText('l');
		EnterText('l');
		EnterText('o');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the whole typed word');

		Redo();
		assert.strictEqual(GetCellEditorText(), 'Hello', 'Check redo of the whole typed word');

		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo after redo');
	});

	QUnit.test('Check undo/redo in cell editor groups typing by words', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello world');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'Hello ', 'Check undo of the last word, the space stays with the previous word');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the first word with the space');

		Redo();
		assert.strictEqual(GetCellEditorText(), 'Hello ', 'Check redo of the first word');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'Hello world', 'Check redo of the second word');
	});

	QUnit.test('Check undo/redo in cell editor joins several spaces to the previous word', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab   cd');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab   ', 'Check undo of the last word');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the first word with spaces');
	});

	QUnit.test('Check a pause breaks the typing group', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		const now = Date.now;
		let time = 10000;
		Date.now = () => time;
		try {
			EnterText('ab');
			time += 1001;
			EnterText('cd');
			Undo();
			assert.strictEqual(GetCellEditorText(), 'ab');
			Redo();
			assert.strictEqual(GetCellEditorText(), 'abcd');
		} finally { Date.now = now; }
	});

	QUnit.test('Check autocomplete continues the typed word', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		cellEditor.options.autoComplete = ['Hello'];
		cellEditor.options.autoCompleteLC = ['hello'];
		cellEditor.objAutoComplete.clear();
		EnterText('Hel');
		assert.strictEqual(GetCellEditorText(), 'Hello');
		Undo();
		assert.strictEqual(GetCellEditorText(), '');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'Hello');
	});

	QUnit.test('Check formula bar input groups words', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		const checkText = AscFonts.FontPickerByCharacter.checkText;
		AscFonts.FontPickerByCharacter.checkText = (text, context, callback) => callback.call(context);
		try {
			for (const ch of 'Hello world') {
				cellEditor.input.value = GetCellEditorText() + ch;
				cellEditor.input.selectionStart = cellEditor.input.selectionEnd = cellEditor.input.value.length;
				cellEditor._onInputTextArea({inputType: 'insertText'});
			}
			Undo();
			assert.strictEqual(GetCellEditorText(), 'Hello ');
			Redo();
			assert.strictEqual(GetCellEditorText(), 'Hello world');
		} finally { AscFonts.FontPickerByCharacter.checkText = checkText; }
	});

	QUnit.test('Check composition is undone and redone as one edit', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		cellEditor.Begin_CompositeInput();
		cellEditor.Replace_CompositeText([0x4e2d]);
		cellEditor.Replace_CompositeText([0x4e2d, 0x6587]);
		cellEditor.End_CompositeInput();
		assert.strictEqual(GetCellEditorText(), 'ab中文');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'ab中文');
		Undo();
		Undo();
		assert.strictEqual(GetCellEditorText(), '');
	});

	QUnit.test('Check cancelling composition preserves redo', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('abc');
		Undo();
		cellEditor.Begin_CompositeInput();
		cellEditor.End_CompositeInput();
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abc');
		Undo();
		cellEditor.Begin_CompositeInput();
		cellEditor.Replace_CompositeText([0x4e2d]);
		cellEditor.Replace_CompositeText([]);
		cellEditor.End_CompositeInput();
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abc');
	});

	QUnit.test('Check composition in an empty editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		cellEditor.Begin_CompositeInput();
		cellEditor.Replace_CompositeText([0x4e2d]);
		cellEditor.End_CompositeInput();
		Undo();
		assert.strictEqual(GetCellEditorText(), '');
		Redo();
		assert.strictEqual(GetCellEditorText(), '中');
	});

	QUnit.test('Check composition restores a replaced selection', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('abc');
		cellEditor._moveCursor(-11, 1);
		cellEditor._selectChars(-11, 3);
		cellEditor.Begin_CompositeInput();
		cellEditor.Replace_CompositeText([0x4e2d]);
		cellEditor.End_CompositeInput();
		Undo();
		assert.strictEqual(GetCellEditorText(), 'abc');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'a中');
	});

	QUnit.test('Check autocomplete does not swallow a rejected suggestion', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		cellEditor.options.autoComplete = ['Hello'];
		cellEditor.options.autoCompleteLC = ['hello'];
		cellEditor.objAutoComplete.clear();
		EnterText('Hex');
		assert.strictEqual(GetCellEditorText(), 'Hex');
		Undo();
		assert.strictEqual(GetCellEditorText(), '');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'Hex');
	});

	QUnit.test('Check cursor movement breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('Hello');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		EnterText('abc');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'Hello', 'Check undo of the text typed after the cursor returned to the same position');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the text typed before the cursor movement');
	});

	QUnit.test('Check typing inside the text is grouped separately in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('ad');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		EnterText('bc');
		assert.strictEqual(GetCellEditorText(), 'abcd', 'Check text');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ad', 'Check undo of the inserted group');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abcd', 'Check redo of the inserted group');
	});

	QUnit.test('Check deletion breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeCharBack);
		EnterText('p');
		assert.strictEqual(GetCellEditorText(), 'Hellp', 'Check text');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'Hell', 'Check undo of the text typed after deletion');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'Hello', 'Check undo of the deletion');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the text typed before deletion');
	});

	QUnit.test('Check typing after undoing deletion starts a new group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeCharBack);
		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo restores the deleted character');
		EnterText('c');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo removes only the new typing');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abc', 'Check redo restores the new typing');
		Undo();
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check the original word remains a separate undo step');
	});

	QUnit.test('Check typing after undoing rich text paste starts a new group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		const fragment = cellEditor.options.fragments[0].clone();
		fragment.setCharCodes([88, 89]);
		fragment.format.setBold(true);
		cellEditor.paste([fragment]);
		assert.strictEqual(GetCellEditorText(), 'abXY', 'Check rich text is pasted');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo removes the pasted text');
		EnterText('c');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo removes only the new typing');
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abc', 'Check redo restores the new typing');
	});

	QUnit.test('Check Delete at the end breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontChar);
		EnterText('c');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo removes only the text typed after Delete');
	});

	QUnit.test('Check forward deletion breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('abcd');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		EnterText('x');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontChar);
		EnterText('y');
		assert.strictEqual(GetCellEditorText(), 'abxyd', 'Check text');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'abxd', 'Check undo of the text typed after forward deletion');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'abxcd', 'Check undo of the forward deletion');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'abcd', 'Check undo of the text typed before forward deletion');
	});

	QUnit.test('Check format change breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		cellEditor.setTextStyle('b', true);
		EnterText('cd');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo of the bold text');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the regular text');

		Redo();
		Redo();
		assert.strictEqual(GetCellEditorText(), 'abcd', 'Check redo of both groups');
		const fragments = cellEditor.options.fragments;
		assert.strictEqual(fragments.length, 2, 'Check redo keeps two fragments');
		assert.notOk(fragments[0].format.getBold(), 'Check redo keeps the first group regular');
		assert.ok(fragments[1].format.getBold(), 'Check redo keeps the second group bold');
	});

	QUnit.test('Check paste breaks the typing group in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ab');
		cellEditor.pasteText('XY');
		EnterText('cd');
		assert.strictEqual(GetCellEditorText(), 'abXYcd', 'Check text');

		Undo();
		assert.strictEqual(GetCellEditorText(), 'abXY', 'Check undo of the text typed after paste');
		Undo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check undo of the paste');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the text typed before paste');
	});

	QUnit.test('Check undo/redo in cell editor after group undo starts a new group', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello');
		Undo();
		EnterText('ab');
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check text');

		Redo();
		assert.strictEqual(GetCellEditorText(), 'ab', 'Check typing clears the redo list');
		Undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check undo of the new group');
	});
	QUnit.test('Check toolbar undo/redo in cell editor with empty workbook history', (assert) =>
	{
		AscCommon.History.Clear();
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello');
		assert.false(AscCommon.History.Can_Undo(), 'Check workbook history is empty');

		wbView.undo();
		assert.strictEqual(GetCellEditorText(), '', 'Check toolbar undo of the whole typed word');

		wbView.redo();
		assert.strictEqual(GetCellEditorText(), 'Hello', 'Check toolbar redo');
	});
	QUnit.test('Check switch reference of formula', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('=F4');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=$F$4', 'Check switch to absolute reference');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=F$4', 'Check switch to absolute row');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=$F4', 'Check switch to absolute column');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=F4', 'Check switch to relative reference');
	});

	QUnit.test('Test catch events', (assert) =>
	{
		function ExecuteTestWithCatchEvent(sSendEvent, fCustomCheck, customExpectedValue, oEvent)
		{
			let check = false;
			const Check = function (...args)
			{
				if (fCustomCheck)
				{
					check = fCustomCheck(...args);
				}
				else
				{
					check = true;
				}
			}
			editor.asc_registerCallback(sSendEvent, Check);
			ExecuteShortcut(oEvent);
			assert.strictEqual(check, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
			editor.asc_unregisterCallback(sSendEvent, Check);
		}

		ExecuteTestWithCatchEvent('asc_onPrint', () => true, true, Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint);
		ExecuteTestWithCatchEvent('asc_onContextMenu', () => true, true, tableEvents[tableHotkeyTypes.contextMenu][0]);


		ClearShapeAndAddParagraph('');
		const hyperlink = new AscCommonExcel.Hyperlink();
		hyperlink.Hyperlink = "https://www.onlyoffice.com/";
		const ascHyperlink = new Asc.asc_CHyperlink(hyperlink);
		ascHyperlink.text = "Hello";
		editor.asc_insertHyperlink(ascHyperlink);
		GetDrawingObjects().cursorMoveLeft();
		GetDrawingObjects().cursorMoveLeft();

		ExecuteTestWithCatchEvent('asc_onHyperlinkClick', (sValue) => sValue, 'https://www.onlyoffice.com/', graphicEvents[graphicHotkeyTypes.visitHyperink][0]);
		const selectedInfo = new CSelectedElementsInfo();
		GetDrawingObjects().getTargetDocContent().GetSelectedElementsInfo(selectedInfo);
		assert.true(selectedInfo.m_oHyperlink.Visited, 'Check hyperlink visited');
		RemoveGraphicObjects();

	});
	QUnit.module('Test graphic shortcuts', {
		afterEach: function ()
		{
			RemoveGraphicObjects();
		}
	});
	QUnit.test('Check remove parts of text', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('Hello Hello Hello Hello Hello Hello Hello');
		GetDrawingObjects().cursorMoveToEndPos();

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackChar);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackWord);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackWord, 1);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

		GetDrawingObjects().cursorMoveToStartPos();
		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontChar);
		assert.strictEqual(GetParagraphText(paragraph), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontWord);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello ', 'Check removing front word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontWord, 1);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello ', 'Check removing front word');
	});

	QUnit.test('Check remove graphic objects', (assert) =>
	{

		function CheckRemoveObject(spTree, shape)
		{
			assert.true(spTree.indexOf(shape) === -1, 'Check remove shape from spTree');
		}

		function CheckRemoveDrawingObject(shape)
		{
			assert.false(ws.Drawings.some((e) => e.graphicObject === shape), 'Check remove shape from spTree');
		}

		const CheckActionsWithEvent = (eventIndex) =>
		{
			const shape = AddShape(0, 0, 100, 100);
			SelectDrawings([shape]);
			assert.true(ws.Drawings.some((e) => e.graphicObject === shape), 'Check available shape in spTree');
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeShape, eventIndex);
			CheckRemoveDrawingObject(shape);

			const chart = AddChart();
			SelectDrawings([chart]);
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeChart, eventIndex);
			CheckRemoveDrawingObject(chart);


			const shape1 = AddShape(0, 0, 100, 100);
			const shape2 = AddShape(0, 0, 100, 100);
			const shape3 = AddShape(0, 0, 100, 100);

			const group = CreateGroup([shape1, shape2, shape3]);
			group.selectObject(shape1, 0);
			GetDrawingObjects().selection.groupSelection = group;

			ExecuteGraphicHotkey(graphicHotkeyTypes.removeShapeInGroup, eventIndex);
			CheckRemoveObject(group.spTree, shape1);

			SelectDrawings([group]);
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeGroup, eventIndex);
			CheckRemoveDrawingObject(group);
		};

		CheckActionsWithEvent(0);
		CheckActionsWithEvent(1);
	});

	QUnit.test('Check main actions with shapes', (assert) =>
	{
		const drawing1 = AddShape(100, 100, 100, 200);

		function CheckShapePosition(X, Y)
		{
			assert.deepEqual([Round(drawing1.x, 12), Round(drawing1.y, 12), Round(drawing1.extX, 12), Round(drawing1.extY, 12)], [Round(X, 12), Round(Y, 12), 200, 100], 'Check shape position after movement ' + X + ' ' + Y);
		}

		SelectDrawings([drawing1]);


		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectLeft);
		CheckShapePosition(100 - 5 * AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectLeft);
		CheckShapePosition(100 - 6 * AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectRight);
		CheckShapePosition(100 - AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectRight);
		CheckShapePosition(100, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectDown);
		CheckShapePosition(100, 100 + 5 * AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectDown);
		CheckShapePosition(100, 100 + 6 * AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectUp);
		CheckShapePosition(100, 100 + AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectUp);
		CheckShapePosition(100, 100);

		function CheckSelectedObjects(arrOfDrawings)
		{
			const length = Math.max(arrOfDrawings.length, GetDrawingObjects().selectedObjects.length);
			for (let i = 0; i < length; i++)
			{
				assert.true(GetDrawingObjects().selectedObjects[i] === arrOfDrawings[i], 'Check selection movement between objects');
			}
		}

		GetDrawingObjects().resetSelection();
		const drawing2 = AddShape(0, 0, 10, 10);
		const drawing3 = AddShape(0, 0, 10, 10);
		SelectDrawings([drawing3]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing1]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing3]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing1]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing3]);

		drawing2.setWordShape(false);
		SelectDrawings([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.createTxBody);
		assert.true(!!drawing2.txBody, 'Check create textBody after enter');

		SelectDrawings([drawing2]);
		ExecuteGraphicHotkey(graphicHotkeyTypes.moveToStartInEmptyContent);
		const paragraph = drawing2.getDocContent().Content[0];
		assert.true(drawing2.parent.controller.selection.textSelection === drawing2, 'Check text selection');
		assert.true(paragraph.IsCursorAtBegin(), 'Check movement to start position in empty content');

		EnterText('Hello');
		SelectDrawings([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectAllAfterEnter);
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Hello\r\n', 'Check select non empty content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetTextSelection);
		assert.true(GetDrawingObjects().selectedObjects[0] === drawing2, 'Check reset text selection');
		assert.true(!GetDrawingObjects().selection.textSelection, 'Check reset text selection');


		let group = CreateGroup([drawing1, drawing2]);
		group.selectObject(drawing1, 0);
		GetDrawingObjects().selection.groupSelection = group;

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetStepSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects[0].getObjectType(), AscDFH.historyitem_type_GroupShape, 'Check reset step drawing selection');
		assert.strictEqual(group.selectedObjects.length, 0, 'Check reset step selection');

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetStepSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects.length, 0, 'Check reset all drawing selection');
	});

	QUnit.test('Check actions with text movements', (assert) =>
	{
		const {shape} = ClearShapeAndAddParagraph(
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'World Hello World');

		function CheckCursorPosition(expected, description)
		{
			const position = shape.txBody.content.GetContentPosition();
			assert.strictEqual(position[position.length - 1].Position, expected, description);
		}

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine);
		CheckCursorPosition(18, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftChar);
		CheckCursorPosition(18, 'Check move to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftWord);
		CheckCursorPosition(12, 'Check move to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftWord, 1);
		CheckCursorPosition(6, 'Check move to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord, 1);
		CheckCursorPosition(12, 'Check move to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord);
		CheckCursorPosition(18, 'Check move to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord);
		CheckCursorPosition(24, 'Check move to right word');


		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorDown);
		CheckCursorPosition(36, 'Check move down');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorUp);
		CheckCursorPosition(18, 'Check move up');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndDocument);
		CheckCursorPosition(161, 'Check move to end document');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartDocument);
		CheckCursorPosition(0, 'Check move to start document');

		function CheckSelectedText(expectedText, description)
		{
			const selectedText = GetDrawingObjects().GetSelectedText();
			assert.strictEqual(selectedText, expectedText, description);
		}

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine);
		CheckSelectedText('Hello World Hello ', 'Select to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftChar);
		CheckSelectedText('Hello World Hello ', 'Select to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord, 1);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord, 1);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello World ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine);
		CheckSelectedText('Hello World Hello ', 'Select to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorDown);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorUp);
		CheckSelectedText('Hello World Hello ', 'Select up');

		GetDrawingObjects().cursorMoveToEndPos();
		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftChar);
		CheckSelectedText('d', 'Select to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord);
		CheckSelectedText('World', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to start content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndDocument);
		CheckSelectedText('\r\n', 'Select to end content');

		GetDrawingObjects().cursorMoveToStartPos();

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World\r\n', 'Select to end content');
	});

	QUnit.test('Check paragraph property change', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello world');

		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check align left");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CenterPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Center, "Check turn on center para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CenterPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on center para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.JustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.JustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on justify para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.JustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.LeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.LeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on left para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.LeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RightPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Right, "Check turn on right para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RightPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on right para");
	});

	QUnit.test('Check text property change', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello world');
		GetDrawingObjects().selectAll();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.strictEqual(GetDirectGraphicTextPr().GetBold(), true, 'Check turn on bold');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.strictEqual(GetDirectGraphicTextPr().GetBold(), false, 'Check turn off bold');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.strictEqual(GetDirectGraphicTextPr().GetItalic(), true, 'Check turn on italic');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.strictEqual(GetDirectGraphicTextPr().GetItalic(), false, 'Check turn off italic');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.strictEqual(GetDirectGraphicTextPr().GetUnderline(), true, 'Check turn on underline');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.strictEqual(GetDirectGraphicTextPr().GetUnderline(), false, 'Check turn off underline');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_SuperScript, 'Check turn on superscript');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off superscript');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_SubScript, 'Check turn on subscript');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off subscript');

		// defaultSize = 10
		// 10 -> 11 -> 12 -> 14 -> 16 -> 14 -> 12 -> 11 -> 10
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 11, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 12, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 14, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 16, 'Check increase font size');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 14, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 12, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 11, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 10, 'Check decrease font size');
	});

	QUnit.test('Check add various characters', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EnDash);
		assert.strictEqual(GetParagraphText(paragraph), String.fromCharCode(0x2013), 'Check add en dash');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EnDash);
		assert.strictEqual(GetParagraphText(paragraph), String.fromCharCode(0x2013, 0x2013), 'Check add en dash');
	});

	QUnit.test('Check select all', (assert) =>
	{
		const {paragraph, shape} = ClearShapeAndAddParagraph('Hello');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);

		assert.true(shape.txBody.content.IsSelectionUse(), 'Check content selection');
		assert.true(paragraph.IsSelectedAll(), 'Check paragraph selection');
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Hello\r\n', 'Check selected text');

		const drawing2 = AddShape();

		SelectDrawings([shape]);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		assert.strictEqual(GetDrawingObjects().selectedObjects.length, 2, 'Check amount of selected shapes');
		assert.true(GetDrawingObjects().selectedObjects[0] === drawing2, 'Check selected shape');
		assert.true(GetDrawingObjects().selectedObjects[1] === shape, 'Check selected shape');
	});

	QUnit.test('Check select all in chart title', (assert) =>
	{
		const chart = AddChart();
		SelectDrawings([chart]);
		const titles = chart.getAllTitles();
		const controller = GetDrawingObjects();
		controller.selection.chartSelection = chart;
		chart.selectTitle(titles[0], 0);
		ExecuteGraphicHotkey(graphicHotkeyTypes.selectAllTitleAfterEnter);
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Diagram Title\r\n', 'Check select all content in chart title');
	});

	QUnit.test('Check add tab symbol', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteGraphicHotkey(graphicHotkeyTypes.addTab);
		assert.true(paragraph.Content[0].Content[0].IsTab(), 'Check add tab symbol');
	});

	QUnit.test('Check add new paragraph', (assert) =>
	{
		const {shape} = ClearShapeAndAddParagraph('');
		ExecuteGraphicHotkey(graphicHotkeyTypes.addParagraph);
		assert.strictEqual(shape.txBody.content.Content.length, 2, 'Check add new paragraph');
	});

	QUnit.test('Check add break line', (assert) =>
	{
		function CheckAddBreakLine(shape)
		{
			assert.strictEqual(shape.txBody.content.Content.length, 1, 'Check amount paragraphs');
			assert.strictEqual(shape.txBody.content.Content[0].GetLinesCount(), 2, 'Check amount lines');
		}

		{

			const {shape} = ClearShapeAndAddParagraph('Hello Hello');
			GetDrawingObjects().cursorMoveRight(false, true);
			ExecuteGraphicHotkey(graphicHotkeyTypes.addBreakLine);
			CheckAddBreakLine(shape);

		}

		{
			const {shape} = ClearShapeAndAddParagraph('');
			editor.asc_AddMath2(c_oAscMathType.FractionVertical);
			GetDrawingObjects().recalculate();
			const paragraph = shape.txBody.content.Content[0];
			GetDrawingObjects().cursorMoveLeft();
			GetDrawingObjects().cursorMoveLeft();
			EnterText('8888888');
			GetDrawingObjects().cursorMoveLeft();
			GetDrawingObjects().cursorMoveLeft();
			ExecuteGraphicHotkey(graphicHotkeyTypes.addLineInMath);
			const paraMath = paragraph.GetAllParaMaths()[0];
			const fraction = paraMath.Root.GetFirstElement();
			const numerator = fraction.getNumerator();
			const eqArray = numerator.GetFirstElement();
			assert.strictEqual(eqArray.getRowsCount(), 2, 'Check add new line in math');
		}
	});

	QUnit.test('Check slicer shortcuts', (assert) =>
	{
		const ExecuteHotkeyWithEvent = (index) =>
		{
			const slicer = CreateSlicer();
			ExecuteGraphicHotkey(graphicHotkeyTypes.invertMultiselectSlicer, index);
			assert.true(slicer.isMultiSelect(), 'Check multiselect slicer');

			slicer.buttonsContainer.buttons[0].setUnselectTmpState();
			ExecuteGraphicHotkey(graphicHotkeyTypes.cleanSlicer, index);
			assert.true(slicer.buttonsContainer.buttons[0].isSelected(), 'Check clean slicer');
			RemoveGraphicObjects();
			ResetData(0, 0, 49, 49);
		};
		ExecuteHotkeyWithEvent(0);
		ExecuteHotkeyWithEvent(1);
	});


	QUnit.module('Test table hotkeys', {
		before   : function f()
		{
			SetVisibleRange(0, 0, 49, 49);
		},
		afterEach: function ()
		{
			CloseCellEditor(true);
			ResetData(0, 0, 49, 49);
		}
	});
	QUnit.test('Check drop down menu', (assert) =>
	{
		let check = false;
		const ShowAutoComplete = function () {check = true;};
		wbView.handlers.add("asc_onEntriesListMenu", ShowAutoComplete);
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showAutoComplete);
		assert.true(check);
		wbView.handlers.remove("asc_onEntriesListMenu", ShowAutoComplete);
		check = false;

		const props = ws.getDataValidationProps();
		props.setType(Asc.EDataValidationType.List);
		props.setFormula1(new Asc.CDataFormula('"1"'));
		ws.setDataValidationProps(props);
		const ShowDataValidations = function () {check = true;};
		wbView.handlers.add("asc_onValidationListMenu", ShowDataValidations);
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showDataValidation);
		assert.true(check);
		check = false;
		ExecuteTableHotkey(tableHotkeyTypes.showDataValidation, 1);
		assert.true(check);
		ws.deleteDataValidationById(props.Get_Id());
		wbView.handlers.remove("asc_onValidationListMenu", ShowDataValidations);
		check = false;

		const SetAutoFiltersDialog = function () {check = true;};
		wbView.handlers.add("setAutoFiltersDialog", SetAutoFiltersDialog);
		CreateTable();
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showFilterOptions);
		assert.true(check);
		wbView.handlers.remove("setAutoFiltersDialog", SetAutoFiltersDialog);
	});
	QUnit.test('Check cell movement', (assert) =>
	{
		function CheckActiveCell(expectedRow, expectedColumn, description)
		{
			assert.deepEqual([ws.selectionRange.activeCell.row, ws.selectionRange.activeCell.col], [expectedRow, expectedColumn], description);
		}

		function CheckActiveRange(r1, c1, r2, c2, description)
		{
			const lastRange = ws.selectionRange.getLast();
			assert.deepEqual([lastRange.r1, lastRange.c1, lastRange.r2, lastRange.c2], [r1, c1, r2, c2], description);
		}


		Select(0, 15, 0, 15, 0, 15);
		FillActiveCell('Hello');
		Select(25, 0, 25, 0, 25, 0);
		FillActiveCell('Hello');
		Select(27, 0, 27, 0, 27, 0);
		FillActiveCell('Hello');

		Select(5, 5, 5, 5, 5, 5);
		FillActiveCell('Hello');
		Select(4, 8, 4, 8, 4, 8);
		FillActiveCell('Hello');

		Select(35, 0, 35, 0, 35, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(27, 0, 'Check move to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(25, 0, 'Check move to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(0, 0, 'Check move to top cell');


		Select(0, 39, 0, 39, 0, 39);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeCell);
		CheckActiveCell(0, 15, 'Check move to left edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeCell);
		CheckActiveCell(0, 0, 'Check move to left edge cell');

		Select(0, 2, 0, 2, 0, 2);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftCell);
		CheckActiveCell(0, 1, 'Check move to left cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell);
		CheckActiveCell(0, 15, 'Check move to right edge cell');

		Select(4, 0, 4, 0, 4, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell, 1);
		CheckActiveCell(4, 8, 'Check move to right edge cell');

		Select(5, 0, 5, 0, 5, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell);
		CheckActiveCell(5, 5, 'Check move to right edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToRightCell);
		CheckActiveCell(5, 6, 'Check move to right cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightCell);
		CheckActiveCell(5, 7, 'Check move to right cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToDownCell);
		CheckActiveCell(6, 7, 'Check move to down cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToDownCell);
		CheckActiveCell(7, 7, 'Check move to down cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToUpCell);
		CheckActiveCell(6, 7, 'Check move to up cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToUpCell);
		CheckActiveCell(5, 7, 'Check move to up cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToBottomCell);
		CheckActiveCell(25, 0, 'Check move to bottom cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToBottomCell);
		CheckActiveCell(27, 0, 'Check move to bottom cell');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.moveToFirstColumn);
		CheckActiveCell(5, 0, 'Check move to first column');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeTop);
		CheckActiveCell(0, 0, 'Check move to left-top edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToRightBottomEdge);
		CheckActiveCell(27, 15, 'Check move to right-bottom edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToLowerCell);
		CheckActiveCell(77, 15, 'Check move to lower cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToUpperCell);
		CheckActiveCell(27, 15, 'Check move to upper cell');

		Select(35, 0, 35, 0, 35, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(27, 0, 35, 0, 'Check select to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(25, 0, 35, 0, 'Check select to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(0, 0, 35, 0, 'Check select to top cell');

		Select(0, 39, 0, 39, 0, 39);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeCell);
		CheckActiveRange(0, 15, 0, 39, 'Check select to left edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeCell);
		CheckActiveRange(0, 0, 0, 39, 'Check select to left edge cell');

		Select(0, 2, 0, 2, 0, 2);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftCell);
		CheckActiveRange(0, 1, 0, 2, 'Check select to left cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell);
		CheckActiveRange(0, 0, 0, 15, 'Check select to right edge cell');

		Select(4, 0, 4, 0, 4, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell, 1);
		CheckActiveRange(4, 0, 4, 8, 'Check select to right edge cell');

		Select(5, 0, 5, 0, 5, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell);
		CheckActiveRange(5, 0, 5, 5, 'Check select to right edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToRightCell);
		CheckActiveRange(5, 0, 5, 6, 'Check select to right cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightCell);
		CheckActiveRange(5, 0, 5, 7, 'Check select to right cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToDownCell);
		CheckActiveRange(5, 0, 6, 7, 'Check select to down cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToDownCell);
		CheckActiveRange(5, 0, 7, 7, 'Check select to down cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToUpCell);
		CheckActiveRange(5, 0, 6, 7, 'Check select to up cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToUpCell);
		CheckActiveRange(5, 0, 5, 7, 'Check select to up cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToBottomCell);
		CheckActiveRange(0, 0, 25, 0, 'Check select to bottom cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToBottomCell);
		CheckActiveRange(0, 0, 27, 0, 'Check select to bottom cell');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.selectToFirstColumn);
		CheckActiveRange(5, 0, 5, 25, 'Check select to first column');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeTop);
		CheckActiveRange(0, 0, 5, 25, 'Check select to left-top edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToRightBottomEdge);
		CheckActiveRange(5, 15, 27, 25, 'Check select to right-bottom edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToLowerCell);
		CheckActiveRange(5, 15, 77, 25, 'Check select to lower cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToUpperCell);
		CheckActiveRange(5, 15, 27, 25, 'Check select to upper cell');

		Select(0, 0, 0, 0, 3, 3);

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 1, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 2, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 3, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(1, 0, 'Check move active cell to right');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 3, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 2, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 1, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 0, 'Check move active cell to left');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(1, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(2, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(3, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(0, 1, 'Check move active cell to down');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(3, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(2, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(1, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(0, 0, 'Check move active cell to up');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectRow);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, 0, 'Check select row');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectColumn);
		CheckActiveRange(0, 0, 0, AscCommon.gc_nMaxCol0, 'Check select column');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectSheet);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0, 'Check select sheet');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0, 'Check select sheet');
	});

	QUnit.test('Check reset actions', (assert) =>
	{
		editor.asc_startAddShape('rect');
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(GetDrawingObjects().checkEndAddShape(), false, 'Check resetting add shape');
		Select(0, 0, 0, 0, 0, 0);
		editor.asc_SelectionCut();
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(wsView.getCutRange(), null, 'Check resetting cut range');

		Select(0, 0, 0, 0, 0, 0);
		editor.asc_Copy();
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(wsView.getCutRange(), null, 'Check resetting copy range');
	});


	QUnit.test('Check actions with removing', (assert) =>
	{
		Select(0, 15, 0, 15, 0, 15);
		FillActiveCell('Hello');
		Select(25, 0, 25, 0, 25, 0);
		FillActiveCell('Hello');
		Select(27, 0, 27, 0, 27, 0);
		FillActiveCell('Hello');

		Select(0, 15, 0, 0, 30, 30);
		ExecuteTableHotkey(tableHotkeyTypes.removeActiveCell);
		CloseCellEditor(true);

		assert.strictEqual(GetCellText(0, 15), '', 'Check remove active cell');
		assert.strictEqual(GetCellText(25, 0), 'Hello', 'Check remove active cell');
		assert.strictEqual(GetCellText(27, 0), 'Hello', 'Check remove active cell');

		ExecuteTableHotkey(tableHotkeyTypes.emptyRange);
		assert.strictEqual(GetCellText(0, 15), '', 'Check remove range');
		assert.strictEqual(GetCellText(25, 0), '', 'Check remove range');
		assert.strictEqual(GetCellText(27, 0), '', 'Check remove range');
	});

	QUnit.test('Check actions with filling cell', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), '.', 'Check add separator');
	});

	QUnit.test('Check actions with formatting cell', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('0.1');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat);
		assert.strictEqual(GetCellText(0, 0), '1.00E-01', 'set exponential format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat);
		assert.strictEqual(GetCellText(0, 0), '10.00%', 'set percent format');

		FillActiveCell('49990');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat);
		assert.strictEqual(GetCellText(0, 0), '$49,990.00', 'set currency format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat);
		assert.strictEqual(GetCellText(0, 0), '12:00:00 AM', 'set time format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellDateFormat);
		assert.strictEqual(GetCellText(0, 0), '11/11/2036', 'set date format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat);
		assert.strictEqual(GetCellText(0, 0), '49990.00', 'set number format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat);
		assert.strictEqual(GetCellText(0, 0), '49990', 'set general format');

		function GetCellFormatting(row, column)
		{
			return ws.getRange4(row, column).getXfs();
		}

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.true(GetCellFormatting(0, 0).asc_getFontStrikeout(), 'Check turn on strikeout format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.false(GetCellFormatting(0, 0).asc_getFontStrikeout(), 'Check turn of strikeout format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.true(GetCellFormatting(0, 0).asc_getFontBold(), 'Check turn on bold format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.false(GetCellFormatting(0, 0).asc_getFontBold(), 'Check turn off bold format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.true(GetCellFormatting(0, 0).asc_getFontItalic(), 'Check turn on italic format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.false(GetCellFormatting(0, 0).asc_getFontItalic(), 'Check turn off italic format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(GetCellFormatting(0, 0).asc_getFontUnderline(), 'Check turn on underline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.false(GetCellFormatting(0, 0).asc_getFontUnderline(), 'Check turn off underline format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontVerticalAlign(), AscCommon.vertalign_SuperScript, 'Check cell editor superscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontVerticalAlign(), AscCommon.vertalign_Baseline, 'Check cell editor baseline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontVerticalAlign(), AscCommon.vertalign_SuperScript, 'Check cell editor superscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontVerticalAlign(), AscCommon.vertalign_SubScript, 'Check cell editor subscript format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontVerticalAlign(), AscCommon.vertalign_Baseline, 'Check cell editor baseline format');

		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check init font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 12, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 14, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 16, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 18, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 20, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 22, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 24, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 26, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 28, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 36, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 48, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 72, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 72, "Check increase font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 48, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 36, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 28, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 26, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 24, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 22, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 20, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 18, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 16, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 14, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 12, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 10, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 9, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 8, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 8, "Check decrease font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check initial value");
	});

	QUnit.test('Check undo/redo', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('0.1');
		CloseCellEditor(true);

		assert.strictEqual(GetCellText(0, 0), '0.1', 'Check cell value');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo);
		assert.strictEqual(GetCellText(0, 0), '', 'Check undo');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo);
		assert.strictEqual(GetCellText(0, 0), '0.1', 'Check redo');
	});

	QUnit.test('Check focus on cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();

		ExecuteTableHotkey(tableHotkeyTypes.openCellEditor);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();

		ExecuteTableHotkey(tableHotkeyTypes.openCellEditor, 1);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();
	});
	QUnit.test('Check show formulas shortcut', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell("=sum(1+2)");
		UpdateView();
		assert.strictEqual(GetCellCacheText(0, 0), "3", "Check show value");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.ShowFormulas);
		assert.strictEqual(GetCellCacheText(0, 0), "=SUM(1+2)", "Check show formula");
	});

	QUnit.test('Check prevent default in Opera', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		assert.strictEqual(ExecuteTableHotkey(tableHotkeyTypes.disableNumLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		assert.strictEqual(ExecuteTableHotkey(tableHotkeyTypes.disableScrollLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
	});

	QUnit.test('Check add sum formula', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('1');
		Select(0, 1, 0, 1, 0, 1);
		EnterText('2');

		Select(7, 7, 7, 7, 7, 7);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction);
		EnterText('A1,B1');
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(7, 7), "3", 'Check sum result');
	});

	QUnit.test('Check move on worksheets', (assert) =>
	{
		CreateWorksheets(['name1', 'name2']);
		const ws1 = wb.getWorksheetByName(['name1'], 0);
		const ws2 = wb.getWorksheetByName(['name2'], 0);
		assert.strictEqual(wb.getActive(), 0, 'Check select first sheet');
		assert.true(wb.getWorksheet(0) === ws1, 'Check select first sheet');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet);
		assert.strictEqual(wb.getActive(), 1, 'Check select second sheet');
		assert.true(wb.getWorksheet(1) === ws2, 'Check select second sheet');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet);
		assert.strictEqual(wb.getActive(), 2, 'Check select third sheet');
		assert.true(wb.getWorksheet(2) === ws, 'Check select third sheet');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet);
		assert.strictEqual(wb.getActive(), 1, 'Check select second sheet');
		assert.true(wb.getWorksheet(1) === ws2, 'Check select second sheet');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet);
		assert.strictEqual(wb.getActive(), 0, 'Check select first sheet');
		assert.true(wb.getWorksheet(0) === ws1, 'Check select first sheet');

		RemoveWorksheets([0, 1]);
	});

	QUnit.test('Check refresh connections', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ad');
		Select(1, 0, 1, 0, 1, 0);
		EnterText('1');
		Select(2, 0, 2, 0, 2, 0);
		EnterText('2');
		CloseCellEditor(true);
		editor.asc_insertPivotExistingWorksheet("Sheet1!$A$1:$A$3", "Sheet1!$C$1");
		const pivotTable1 = ws.getPivotTable(2, 1);
		pivotTable1.asc_addField(editor, 0);

		Select(0, 1, 0, 1, 0, 1);
		EnterText('ap');
		Select(1, 1, 1, 1, 1, 1);
		EnterText('2');
		Select(2, 1, 2, 1, 2, 1);
		EnterText('3');
		CloseCellEditor(true);
		editor.asc_insertPivotExistingWorksheet("Sheet1!$B$1:$B$3", "Sheet1!$D$1");
		const pivotTable2 = ws.getPivotTable(3, 1);
		pivotTable2.asc_addField(editor, 0);

		Select(2, 0, 2, 0, 2, 0);
		EnterText('4');
		Select(2, 1, 2, 1, 2, 1);
		EnterText('4');

		Select(1, 2, 1, 2, 1, 2);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots);
		assert.strictEqual(GetCellText(1, 2), '5', 'Check refresh selected connections');

		Select(1, 3, 1, 3, 1, 3);
		assert.strictEqual(GetCellText(1, 3), '5', 'Check refresh selected connections');

		Select(2, 0, 2, 0, 2, 0);
		EnterText('5');

		Select(1, 2, 1, 2, 1, 2);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots);
		assert.strictEqual(GetCellText(1, 2), '6', 'Check refresh all connections');

		Select(1, 3, 1, 3, 1, 3);
		assert.strictEqual(GetCellText(1, 3), '6', 'Check refresh all connections');

	});
	QUnit.test('Change format table info', (assert) =>
	{
		CreateTable();
		Select(1, 0, 1, 0, 1, 0);
		ExecuteTableHotkey(tableHotkeyTypes.changeFormatTableInfo);
		assert.strictEqual(GetCellText(3, 2), '6', 'Check change format table info');
		ExecuteTableHotkey(tableHotkeyTypes.changeFormatTableInfo);
	});

	QUnit.test('Check recalculating data', (assert) =>
	{
		CreateWorksheets(['name1']);
		const wsName = wb.getWorksheetByName(['name1'], 0);

		ws.getRange4(0, 2).fillData([['=sum(a1+b1)']]);
		ws.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(1);
		});
		ws.getRange4(0, 1)._foreach(function (cell)
		{
			cell.setValueNumberInternal(2);
		});

		wsName.getRange4(0, 2).fillData([['=sum(a1+b1)']]);
		wsName.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(1);
		});
		wsName.getRange4(0, 1)._foreach(function (cell)
		{
			cell.setValueNumberInternal(2);
		});

		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating values');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating values');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet);
		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating value');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '3', 'Check calculating value');

		wsName.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(0);
		});

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateAll);
		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '3', 'Check calculating value');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '2', 'Check calculating value');
		RemoveWorksheets([0]);
	});

	QUnit.module('Test BiDi cell editor', {
		afterEach: function ()
		{
			CloseCellEditor();
		}
	});

	QUnit.test('Arrow keys swap in RTL text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('\u05E9\u05DC\u05D5\u05DD');

		assert.strictEqual(cellEditor.cursorPos, 4, 'Cursor at end after entering Hebrew text');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		assert.strictEqual(cellEditor.cursorPos, 3, 'Right arrow in RTL moves to prev char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		assert.strictEqual(cellEditor.cursorPos, 2, 'Right arrow in RTL moves to prev char again');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		assert.strictEqual(cellEditor.cursorPos, 3, 'Left arrow in RTL moves to next char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		assert.strictEqual(cellEditor.cursorPos, 4, 'Left arrow in RTL moves to next char again');
	});

	QUnit.test('Cursor visual position in RTL text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('\u05E9\u05DC\u05D5\u05DD');

		let tr = cellEditor.textRender;
		let offsetBegin = tr.calcCharOffset(0);
		let offsetEnd = tr.calcCharOffset(4);

		assert.ok(offsetBegin !== null, 'Offset at pos=0 is not null');
		assert.ok(offsetEnd !== null, 'Offset at pos=4 is not null');
		assert.ok(offsetBegin.left > offsetEnd.left, 'In RTL: cursor at pos=0 is visually to the right of cursor at pos=4');

		let offsetMid = tr.calcCharOffset(2);
		assert.ok(offsetMid.left > offsetEnd.left && offsetMid.left < offsetBegin.left,
			'Cursor at pos=2 is between pos=0 and pos=4');
	});

	QUnit.test('Hit-testing in RTL text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('\u05E9\u05DC\u05D5\u05DD');

		let tr = cellEditor.textRender;
		let zoom = cellEditor.getZoom();
		let offsetBegin = tr.calcCharOffset(0);
		let offsetEnd = tr.calcCharOffset(4);

		let posAtRight = tr.getCharPosByXY(offsetBegin.left, 1, 0, zoom);
		assert.strictEqual(posAtRight, 0, 'Click at right edge of RTL text gives pos=0');

		let posAtLeft = tr.getCharPosByXY(offsetEnd.left, 1, 0, zoom);
		assert.strictEqual(posAtLeft, 4, 'Click at left edge of RTL text gives pos=4');
	});

	QUnit.test('Selection rects in RTL text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('\u05E9\u05DC\u05D5\u05DD');

		let rects = cellEditor._collectSelectionRects(0, 1, 3);
		assert.strictEqual(rects.length, 1, 'Pure RTL selection produces one contiguous rect');
		assert.ok(rects[0].w > 0, 'Selection rect has positive width');
	});

	QUnit.test('Mixed LTR/RTL cursor positions', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('AB\u05E9\u05DCCD');

		let tr = cellEditor.textRender;
		let offset0 = tr.calcCharOffset(0);
		let offset2 = tr.calcCharOffset(2);
		let offset4 = tr.calcCharOffset(4);
		let offset6 = tr.calcCharOffset(6);

		assert.ok(offset0.left < offset2.left, 'LTR part: pos=0 is left of pos=2');
		assert.ok(offset2.left >= offset4.left, 'RTL part: pos=2 boundary is at or right of pos=4 boundary');
		assert.ok(offset4.left < offset6.left, 'LTR part after RTL: pos=4 is left of pos=6');
	});

	QUnit.test('Cursor at BiDi boundary with trailing edge hint', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('AB\u05E9\u05DCCD');

		let tr = cellEditor.textRender;

		tr.cursorAtTrailingEdge = true;
		let offsetTrailing = tr.calcCharOffset(2);

		tr.cursorAtTrailingEdge = false;
		let offsetLeading = tr.calcCharOffset(2);

		tr.cursorAtTrailingEdge = undefined;

		assert.ok(offsetTrailing !== null && offsetLeading !== null,
			'Both cursor positions at BiDi boundary are valid');
	});

	QUnit.test('Home/End keys work in RTL text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('\u05E9\u05DC\u05D5\u05DD');

		cellEditor._moveCursor(-11, 2);
		assert.strictEqual(cellEditor.cursorPos, 2, 'Cursor moved to pos=2');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		assert.strictEqual(cellEditor.cursorPos, 0, 'Home moves to logical start (pos=0)');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine);
		assert.strictEqual(cellEditor.cursorPos, 4, 'End moves to logical end (pos=4)');
	});

	QUnit.module('Cell editor autosave', {
		beforeEach: function ()
		{
			wbView.collaborativeEditing.m_nUseType = 1;
		},
		afterEach: function ()
		{
			CloseCellEditor();
			editor._cellEditorCheckpoint = false;
		}
	});

	[false, true].forEach(function (fast)
	{
		QUnit.test('Send a pending checkpoint after the save delay, fast collaboration: ' + fast, (assert) =>
		{
			const DateOriginal = window.Date;
			const askSaveOriginal = editor.CoAuthoringApi.askSaveChanges;
			const lastSaveTimeOriginal = editor.lastSaveTime;
			const collaboration = wbView.collaborativeEditing;
			const fastOriginal = collaboration.getFast();
			const useTypeOriginal = collaboration.m_nUseType;
			const openedFrameOriginal = editor.isOpenedFrameEditor;
			editor.isOpenedFrameEditor = false;
			const askLock = function (locks, callback) { callback({lock: true}); };
			collaboration.handlers.add('askLock', askLock);
			collaboration.m_nUseType = 1;
			let now = DateOriginal.now();
			let saveRequests = 0;
			let saveCallback;
			window.Date = class extends DateOriginal {
				constructor(...args) { super(...(args.length ? args : [now])); }
				static now() { return now; }
			};
			editor.CoAuthoringApi.askSaveChanges = function (callback) { saveRequests++; saveCallback = callback; };
			editor.lastSaveTime = null;
			if (fast) {
				collaboration.startCollaborationEditing();
				collaboration.setFast(true);
			}
			try {
				Select(0, 0, 0, 0, 0, 0);
				FillActiveCell('before');
				OpenCellEditor();
				EnterText('after');
				cellEditor.autosaveValue();
				assert.strictEqual(GetCellText(0, 0), 'after', 'Checkpoint reached the workbook');
				assert.true(!!editor._cellEditorCheckpoint, 'Pending marker created');
				assert.true(editor._saveCheck(), 'Save preconditions met');
				editor._autoSaveInner();
				assert.strictEqual(saveRequests, 0, 'The normal save delay still applies');
				now += 12000;
				editor._autoSaveInner();
				assert.strictEqual(saveRequests, 1, 'An unchanged pending checkpoint reaches the save lifecycle');
				assert.true(cellEditor.isOpened, 'Sending keeps the editor open');
				if (saveCallback) {
					saveCallback({saveLock: true});
					editor._autoSaveInner();
					assert.strictEqual(saveRequests, 2, 'A busy save lock does not discard the pending checkpoint');
				}
			} finally {
				window.Date = DateOriginal;
				editor.CoAuthoringApi.askSaveChanges = askSaveOriginal;
				editor.lastSaveTime = lastSaveTimeOriginal;
				editor.canSave = true;
				editor._cellEditorSaveInProgress = false;
				if (fast) {
					collaboration.endCollaborationEditing();
					collaboration.setFast(fastOriginal);
				}
				collaboration.m_nUseType = useTypeOriginal;
				collaboration.handlers.remove('askLock', askLock);
				collaboration.m_arrNeedUnlock2.length = 0;
				editor.isOpenedFrameEditor = openedFrameOriginal;
			}
		});
	});

	QUnit.test('Document autosave does not commit before the input debounce', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('before');
		OpenCellEditor();
		EnterText('unfinished');
		editor._autoSaveInner();
		cellEditor.close(false);
		assert.strictEqual(GetCellText(0, 0), 'before', 'Esc before the pause discards the input');
	});

	QUnit.test('Resume checkpointing after typing during a regular document save', (assert) =>
	{
		const done = assert.async();
		const collaboration = wbView.collaborativeEditing;
		const sendChangesOriginal = collaboration.sendChanges;
		const lastSaveTimeOriginal = editor.lastSaveTime;
		const onUnSaveLockOriginal = editor.CoAuthoringApi.onUnSaveLock;
		collaboration.sendChanges = function () {};
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('before');
		editor._cellEditorSaveInProgress = false;
		editor.canSave = false;
		editor._onSaveCallbackInner();
		OpenCellEditor();
		EnterText('after');
		assert.strictEqual(GetCellText(0, 0), 'before', 'Typing does not mutate history during the save');
		editor.CoAuthoringApi.onUnSaveLock();
		window.setTimeout(function () {
			try {
				assert.strictEqual(GetCellText(0, 0), 'after', 'Save completion resumes the paused checkpoint');
				assert.true(cellEditor.isOpened, 'The edit session stays open');
			} finally {
				collaboration.sendChanges = sendChangesOriginal;
				editor.lastSaveTime = lastSaveTimeOriginal;
				editor.CoAuthoringApi.onUnSaveLock = onUnSaveLockOriginal;
				editor.canSave = true;
				done();
			}
		}, 5100);
	});

	QUnit.test('Validate a checkpoint before changing the workbook', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('5');
		const props = ws.getDataValidationProps();
		props.setType(Asc.EDataValidationType.Whole);
		props.setShowErrorMessage(true);
		props.setFormula1(new Asc.CDataFormula('1'));
		props.setFormula2(new Asc.CDataFormula('10'));
		ws.setDataValidationProps(props);
		const errors = [];
		const onError = function (id) { errors.push(id); };
		wb.handlers.add('asc_onError', onError);
		try {
			OpenCellEditor();
			EnterText('99');
			cellEditor.autosaveValue();
			assert.strictEqual(GetCellText(0, 0), '5', 'Invalid input does not replace the cell');
			assert.false(cellEditor._autosaveInProgress, 'Rejected checkpoint finishes its callback');
			assert.deepEqual(errors, [], 'Background validation does not interrupt typing');
			cellEditor.selectAll();
			EnterText('7');
			cellEditor.autosaveValue();
			assert.strictEqual(GetCellText(0, 0), '7', 'Valid input can still be checkpointed');
			cellEditor.selectAll();
			EnterText('99');
			cellEditor.autosaveValue();
			cellEditor.close(false);
			assert.strictEqual(GetCellText(0, 0), '7', 'Esc preserves the last valid checkpoint');
		} finally {
			wb.handlers.remove('asc_onError', onError);
			ws.clearDataValidation([new Asc.Range(0, 0, 0, 0)], true);
		}
	});

	QUnit.test('Checkpoint cannot overwrite part of a classic formula array', (assert) =>
	{
		const range = ws.getRange2('A1:A2');
		ws.getRange2('B1').setValue('1');
		ws.getRange2('B2').setValue('2');
		range.setValue('=B1:B2', null, null, range.bbox);
		Select(1, 0, 1, 0, 1, 0);
		OpenCellEditor();
		EnterText('replacement');
		cellEditor.autosaveValue();
		assert.false(cellEditor._autosaveInProgress, 'Rejected checkpoint finishes its callback');
		cellEditor.close(false);
		[0, 1].forEach(function (row) {
			ws._getCell(row, 0, function (cell) {
				assert.true(!!cell.formulaParsed, 'Array element retains its formula');
				assert.true(!!cell.formulaParsed && cell.formulaParsed.ref.isEqual(range.bbox), 'Array range stays intact');
			});
		});
		range.cleanAll();
	});

	[false, true].forEach(function (grant)
	{
		QUnit.test('Wait for the editing lock before checkpointing, granted: ' + grant, (assert) =>
		{
			const collaboration = wbView.collaborativeEditing;
			let pendingLock;
			const askLock = function (locks, callback) { pendingLock = callback; };
			collaboration.handlers.add('askLock', askLock);
			collaboration.startCollaborationEditing();
			try {
				Select(0, 0, 0, 0, 0, 0);
				FillActiveCell('before');
				OpenCellEditor();
				EnterText('pending');
				assert.strictEqual(typeof pendingLock, 'function', 'Editing lock is awaiting a server response');
				assert.false(cellEditor.autosaveValue(), 'Pending lock forbids a checkpoint');
				assert.strictEqual(GetCellText(0, 0), 'before', 'No value is written before permission');
				pendingLock(grant ? {lock: true} : {error: 'locked'});
				if (grant) {
					assert.true(!!cellEditor._autosaveTimer, 'Lock grant schedules the pending input');
					cellEditor.autosaveValue();
					assert.strictEqual(GetCellText(0, 0), 'pending', 'Granted edit can be checkpointed');
				} else {
					assert.false(cellEditor.isOpened, 'Denied edit is closed');
					assert.strictEqual(GetCellText(0, 0), 'before', 'Denied edit leaves the cell unchanged');
				}
			} finally {
				collaboration.handlers.remove('askLock', askLock);
				collaboration.Set_GlobalLock(false);
				collaboration.m_arrNeedUnlock2.length = 0;
				collaboration.endCollaborationEditing();
			}
		});
	});

	QUnit.test('Save a text snapshot without closing the editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('saved');
		cellEditor._moveCursor(-4);
		cellEditor._selectChars(1, 3);
		const cursorPos = cellEditor.cursorPos;
		const selectionBegin = cellEditor.selectionBegin;
		const selectionEnd = cellEditor.selectionEnd;
		const undoCount = cellEditor.undoList.length;
		const fragments = cellEditor.options.fragments;

		assert.true(cellEditor.autosaveValue(), 'Changed text starts a checkpoint');
		assert.strictEqual(GetCellText(0, 0), 'saved', 'The workbook contains current text');
		assert.true(cellEditor.isOpened, 'The editor remains open');
		assert.strictEqual(GetCellEditorText(), 'saved', 'The editor still has its original text');
		assert.strictEqual(cellEditor.cursorPos, cursorPos, 'Cursor position is unchanged');
		assert.strictEqual(cellEditor.selectionBegin, selectionBegin, 'Selection start is unchanged');
		assert.strictEqual(cellEditor.selectionEnd, selectionEnd, 'Selection end is unchanged');
		assert.strictEqual(cellEditor.undoList.length, undoCount, 'Editor undo list is unchanged');
		assert.strictEqual(cellEditor.redoList.length, 0, 'Editor redo list is unchanged');
		assert.strictEqual(cellEditor.options.fragments, fragments, 'Editor fragments are not replaced');

		assert.false(cellEditor.autosaveValue(), 'An unchanged text needs no new checkpoint');
		assert.strictEqual(GetCellEditorText(), 'saved', 'Deduplication does not change editor text');
	});

	QUnit.test('Keep the cell lock while sending an intermediate checkpoint', (assert) =>
	{
		const collaboration = wbView.collaborativeEditing;
		const lockInfo = collaboration.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Range, null, ws.Id,
			new AscCommonExcel.asc_CCollaborativeRange(0, 0, 0, 0));
		const lock = new AscCommonExcel.CLock(lockInfo);
		lock.setType(AscCommon.c_oAscLockTypes.kLockTypeMine);
		const releasedLocks = [];
		collaboration.startCollaborationEditing();
		collaboration.addUnlock2(lock);
		const releaseLocksHandler = function (guid)
		{
			releasedLocks.push(guid);
		};
		collaboration.handlers.add('releaseLocks', releaseLocksHandler);
		AscCommon.History.Clear();

		collaboration.sendChanges(undefined, undefined, undefined, true);
		assert.strictEqual(collaboration.m_arrNeedUnlock2.length, 1, 'The checkpoint retains the own lock');
		assert.strictEqual(lock.getType(), AscCommon.c_oAscLockTypes.kLockTypeMine, 'The cell remains locked by us');
		assert.deepEqual(releasedLocks, [], 'The checkpoint does not release the cell');
		const saveChangesOriginal = editor.CoAuthoringApi.saveChanges;
		const isDocumentLoadCompleteOriginal = editor.isDocumentLoadComplete;
		editor.isDocumentLoadComplete = true;
		let releaseLocks;
		editor.CoAuthoringApi.saveChanges = function (arrayChanges, deleteIndex, additionalInfo, canUnlock, canRelease)
		{
			releaseLocks = canRelease;
		};
		editor._onSaveChanges(undefined, undefined, undefined, ['change'], true);
		assert.strictEqual(releaseLocks, false, 'The transport also retains collaborative locks');
		editor.CoAuthoringApi.saveChanges = saveChangesOriginal;
		editor.isDocumentLoadComplete = isDocumentLoadCompleteOriginal;

		collaboration.sendChanges();
		assert.strictEqual(collaboration.m_arrNeedUnlock2.length, 0, 'Normal sending releases the lock');
		assert.deepEqual(releasedLocks, [lockInfo.guid], 'The lock is released on the final send');
		collaboration.handlers.remove('releaseLocks', releaseLocksHandler);
		collaboration.endCollaborationEditing();
	});

	QUnit.test('Keep transport locks during an intermediate document save', (assert) =>
	{
		const coAuthoringApi = new AscCommon.CDocsCoApi();
		coAuthoringApi._onlineWork = true;
		coAuthoringApi._CoAuthoringApi._locks = {cell: {state: 2}};
		let message;
		coAuthoringApi._CoAuthoringApi._send = function (data)
		{
			message = data;
		};
		coAuthoringApi.saveChanges([], null, null, false, false);
		const lock = coAuthoringApi._CoAuthoringApi._locks.cell;
		assert.strictEqual(!!lock, true, 'The transport retains its own cell lock');
		assert.strictEqual(message.releaseLocks, false, 'The outgoing message forbids lock release');
		clearTimeout(coAuthoringApi._CoAuthoringApi.saveCallbackErrorTimeOutId);
		coAuthoringApi.saveChanges([], null, null, false, true);
		assert.strictEqual(coAuthoringApi._CoAuthoringApi._locks.cell, undefined, 'Normal save still releases the transport lock');
		assert.strictEqual(message.releaseLocks, true, 'Normal save retains the original protocol');
		clearTimeout(coAuthoringApi._CoAuthoringApi.saveCallbackErrorTimeOutId);
	});

	QUnit.test('Debounce checkpoints after real text changes', (assert) =>
	{
		const timers = [];
		const setTimeoutOriginal = window.setTimeout;
		const clearTimeoutOriginal = window.clearTimeout;
		window.setTimeout = function (callback, delay)
		{
			const timer = {callback: callback, delay: delay, cleared: false};
			timers.push(timer);
			return timer;
		};
		window.clearTimeout = function (timer)
		{
			timer.cleared = true;
		};
		try
		{
			Select(0, 0, 0, 0, 0, 0);
			OpenCellEditor();
			EnterText('h');
			assert.true(!!cellEditor._autosaveTimer, 'Text input schedules a checkpoint');
			assert.strictEqual(cellEditor._autosaveTimer.delay, 5000, 'Checkpoint waits five seconds after input');
			const firstTimer = cellEditor._autosaveTimer;
			EnterText('i');
			assert.true(firstTimer.cleared, 'New input resets the previous timer');
			assert.notStrictEqual(cellEditor._autosaveTimer, firstTimer, 'Only the latest timer is active');
			cellEditor._autosaveTimer.callback();
			assert.strictEqual(GetCellText(0, 0), 'hi', 'The timer saves the latest text');
			assert.true(cellEditor.isOpened, 'The timer does not close the editor');
		}
		finally
		{
			window.setTimeout = setTimeoutOriginal;
			window.clearTimeout = clearTimeoutOriginal;
			CloseCellEditor();
		}
	});

	QUnit.test('Esc keeps the last checkpoint and discards later text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('saved');
		cellEditor.autosaveValue();
		EnterText(' later');
		cellEditor.close(false);
		assert.strictEqual(GetCellText(0, 0), 'saved', 'Esc discards only the unsaved suffix');
	});

	QUnit.test('Keep typing while a checkpoint is pending', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('first');
		const saveValueCallback = cellEditor.options.saveValueCallback;
		let pendingCallback;
		cellEditor.options.saveValueCallback = function (fragments, flags, callback)
		{
			pendingCallback = callback;
		};
		assert.true(cellEditor.autosaveValue(), 'The first checkpoint is pending');
		EnterText(' second');
		assert.false(cellEditor.autosaveValue(), 'A second checkpoint cannot overlap the first');
		pendingCallback(true);
		assert.strictEqual(GetCellEditorText(), 'first second', 'Completion does not overwrite new input');
		assert.true(!!cellEditor._autosaveTimer, 'The new suffix is scheduled after completion');
		cellEditor.options.saveValueCallback = saveValueCallback;
	});

	QUnit.test('Checkpoint deleting the cell text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('before');
		OpenCellEditor();
		cellEditor.empty();
		assert.true(cellEditor.autosaveValue(), 'Deleting the text starts a checkpoint');
		assert.strictEqual(GetCellText(0, 0), '', 'The cell is cleared in the workbook');
	});

	QUnit.test('Do not checkpoint formulas', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('original');
		OpenCellEditor();
		cellEditor.empty();
		EnterText('=1+');
		assert.false(cellEditor.autosaveValue(), 'An incomplete formula is not sent');
		assert.strictEqual(GetCellText(0, 0), 'original', 'Original workbook value is unchanged');
	});
	['+1+2', '-A1'].forEach(function (text)
	{
		QUnit.test('Do not checkpoint formula-like text: ' + text, (assert) =>
		{
			Select(0, 0, 0, 0, 0, 0);
			FillActiveCell('original');
			OpenCellEditor();
			cellEditor.empty();
			EnterText(text);
			assert.false(cellEditor.autosaveValue(), 'Formula-like text is not sent');
			assert.strictEqual(GetCellText(0, 0), 'original', 'Original value is preserved');
		});
	});

	QUnit.test('Skip incomplete input modes and cancel timer on close', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('partial');
		const isMobileVersion = editor.isMobileVersion;
		const isNativeEditor = window['IS_NATIVE_EDITOR'];
		try
		{
			cellEditor.Begin_CompositeInput();
			assert.false(cellEditor.autosaveValue(), 'IME composition is not committed');
			cellEditor.End_CompositeInput();

			cellEditor._clearAutosaveTimer();
			editor.isMobileVersion = true;
			cellEditor._scheduleAutosave();
			assert.false(!!cellEditor._autosaveTimer, 'Mobile editor does not schedule an intermediate commit');
			editor.isMobileVersion = false;

			window['IS_NATIVE_EDITOR'] = true;
			assert.false(cellEditor.autosaveValue(), 'Native editor does not commit unsupported partial input');
		}
		finally
		{
			editor.isMobileVersion = isMobileVersion;
			window['IS_NATIVE_EDITOR'] = isNativeEditor;
		}
	});

	QUnit.test('Cancel a pending checkpoint on close', (assert) =>
	{
		const setTimeoutOriginal = window.setTimeout;
		const clearTimeoutOriginal = window.clearTimeout;
		window.setTimeout = function (callback, delay)
		{
			return {callback: callback, delay: delay, cleared: false};
		};
		window.clearTimeout = function (timer)
		{
			timer.cleared = true;
		};
		try
		{
			Select(0, 0, 0, 0, 0, 0);
			OpenCellEditor();
			EnterText('partial');
			const timer = cellEditor._autosaveTimer;
			cellEditor.close(false);
			assert.true(timer.cleared, 'Closing cancels the pending callback');
			assert.strictEqual(cellEditor._autosaveTimer, null, 'Closing releases the timer reference');
		}
		finally
		{
			window.setTimeout = setTimeoutOriginal;
			window.clearTimeout = clearTimeoutOriginal;
		}
	});

	QUnit.test('Autosave sends a checkpoint without releasing the cell', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('shared');
		const collaboration = wbView.collaborativeEditing;
		const sendChangesOriginal = collaboration.sendChanges;
		const askSaveOriginal = editor.CoAuthoringApi.askSaveChanges;
		const lastSaveTimeOriginal = editor.lastSaveTime;
		const waitOriginal = editor.intervalWaitAutoSave;
		const saveCheckOriginal = editor._saveCheck;
		const calls = [];
		collaboration.sendChanges = function ()
		{
			calls.push(Array.prototype.slice.call(arguments));
		};
		editor.CoAuthoringApi.askSaveChanges = function (callback)
		{
			calls.askSave = callback;
		};
		editor.intervalWaitAutoSave = 0;
		editor.lastSaveTime = new Date(0);
		editor.isDocumentCanSave = true;
		editor._saveCheck = function ()
		{
			return true;
		};
		try
		{
			cellEditor.autosaveValue();
			editor._autoSaveInner();
			assert.strictEqual(typeof calls.askSave, 'function', 'The normal document save lifecycle is requested');
			assert.true(cellEditor.isOpened, 'The autosave preparation keeps the editor open');
			assert.strictEqual(calls.length, 0, 'Changes are sent only after the save lock is granted');
			assert.true(editor.canSave === false, 'Save progress remains tracked normally');

			calls.askSave({saveLock: false});
			assert.strictEqual(calls.length, 1, 'Granted save sends the checkpoint');
			assert.strictEqual(calls[0][3], true, 'The save sends without releasing the cell lock');
			assert.true(cellEditor.isOpened, 'The editor remains open after the save request');
			editor.CoAuthoringApi.onUnSaveLock();
			assert.true(editor.canSave, 'Normal save completion unlocks subsequent saves');
		}
		finally
		{
			collaboration.sendChanges = sendChangesOriginal;
			editor.CoAuthoringApi.askSaveChanges = askSaveOriginal;
			editor.CoAuthoringApi.onUnSaveLock = null;
			editor.lastSaveTime = lastSaveTimeOriginal;
			editor.intervalWaitAutoSave = waitOriginal;
			editor._saveCheck = saveCheckOriginal;
			editor.canSave = true;
		}
	});

	[false, true].forEach(function (typeDuringSave)
	{
		QUnit.test('Saved cell status with an open editor, typing during save: ' + typeDuringSave, (assert) =>
		{
			const openedFrameOriginal = editor.isOpenedFrameEditor;
			editor.isOpenedFrameEditor = false;
			const history = AscCommon.History;
			const collaboration = wbView.collaborativeEditing;
			const sendChangesOriginal = collaboration.sendChanges;
			const triggerOriginal = editor.handlers.trigger;
			const onUnSaveLockOriginal = editor.CoAuthoringApi.onUnSaveLock;
			let modified;
			editor.handlers.trigger = function (name, value)
			{
				if (name === 'asc_onDocumentModifiedChanged') {
					modified = value;
				}
				return triggerOriginal.apply(this, arguments);
			};
			collaboration.sendChanges = function () { history.Reset_SavedIndex(true); };
			try
			{
				Select(0, 0, 0, 0, 0, 0);
				FillActiveCell('before');
				history.Reset_SavedIndex(true);
				OpenCellEditor();
				EnterText('checkpoint');
				assert.true(editor.asc_isDocumentModified(), 'Uncommitted input requires an unload warning');
				cellEditor.autosaveValue();
				assert.true(editor.asc_isDocumentModified(), 'A local checkpoint is not yet saved on the server');
				editor.canSave = false;
				editor._cellEditorSaveInProgress = true;
				editor._onSaveCallbackInner();
				assert.true(editor.asc_isDocumentModified(), 'Keep the warning until the save completes');
				if (typeDuringSave) {
					EnterText(' suffix');
				}
				editor.CoAuthoringApi.onUnSaveLock();
				assert.strictEqual(editor.asc_isDocumentModified(), typeDuringSave, 'Only newer input keeps the unload warning');
				assert.strictEqual(modified, typeDuringSave, 'The status notification agrees with the unload guard');
				assert.strictEqual(editor.asc_isDocumentCanSave(), typeDuringSave, 'The save button reflects outstanding input');
				assert.true(cellEditor.isOpened, 'Reporting a saved state does not close the cell');
				if (!typeDuringSave) {
					EnterText(' new');
					assert.true(editor.asc_isDocumentModified(), 'Typing again restores the unload warning');
					assert.true(modified, 'Typing again immediately updates the status');
				}
				cellEditor.close(false);
				assert.false(editor.asc_isDocumentModified(), 'Esc removes only the unsaved suffix');
				assert.false(modified, 'Esc restores the saved status');
			}
			finally
			{
				collaboration.sendChanges = sendChangesOriginal;
				editor.handlers.trigger = triggerOriginal;
				editor.CoAuthoringApi.onUnSaveLock = onUnSaveLockOriginal;
				editor.canSave = true;
				editor._cellEditorSaveInProgress = false;
				editor.isOpenedFrameEditor = openedFrameOriginal;
			}
		});
	});

	QUnit.test('Formatting and formula input still require an unload warning after a checkpoint', (assert) =>
	{
		const openedFrameOriginal = editor.isOpenedFrameEditor;
		editor.isOpenedFrameEditor = false;
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('saved');
		cellEditor.autosaveValue();
		AscCommon.History.Reset_SavedIndex(true);
		assert.false(editor.asc_isDocumentModified(), 'An unchanged saved checkpoint is clean');
		cellEditor.selectAll();
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.strictEqual(GetCellEditorText(), 'saved', 'Formatting does not change the text');
		assert.true(editor.asc_isDocumentModified(), 'Unsaved formatting must not be mistaken for saved text');
		cellEditor.selectAll();
		EnterText('=SUM(');
		assert.false(cellEditor.autosaveValue(), 'Incomplete formulas cannot be checkpointed');
		assert.true(editor.asc_isDocumentModified(), 'An incomplete formula still needs the warning');
		editor.isOpenedFrameEditor = openedFrameOriginal;
	});

	QUnit.test('Keep the ordinary save path after skipped intermediate input', (assert) =>
	{
		const askSaveOriginal = editor.CoAuthoringApi.askSaveChanges;
		const saveCheckOriginal = editor._saveCheck;
		const lastSaveTimeOriginal = editor.lastSaveTime;
		const waitOriginal = editor.intervalWaitAutoSave;
		const prepareSaveOriginal = editor._prepareSave;
		let intermediatePreparation;
		editor.CoAuthoringApi.askSaveChanges = function () {};
		editor._saveCheck = function ()
		{
			return true;
		};
		editor.lastSaveTime = new Date(0);
		editor.intervalWaitAutoSave = 0;
		editor._prepareSave = function (isIdle)
		{
			prepareSaveOriginal.call(this, isIdle);
			if (!this.IsUserSave) {
				intermediatePreparation = true;
			}
		};
		try
		{
			Select(0, 0, 0, 0, 0, 0);
			FillActiveCell('original');
			OpenCellEditor();
			cellEditor.empty();
			EnterText('=1+2');
			editor._autoSaveInner();
			assert.false(!!intermediatePreparation, 'No checkpoint is prepared for a formula');
			assert.strictEqual(GetCellText(0, 0), 'original', 'Autosave does not commit formula text');

			editor.asc_Save(false);
			assert.false(cellEditor.isOpened, 'Manual save still finishes the formula edit normally');
			assert.false(!!editor._cellEditorSaveInProgress, 'Manual save is not marked as intermediate');
		}
		finally
		{
			editor.CoAuthoringApi.askSaveChanges = askSaveOriginal;
			editor.CoAuthoringApi.onUnSaveLock = null;
			editor.canSave = true;
			editor.lastSaveTime = lastSaveTimeOriginal;
			editor.intervalWaitAutoSave = waitOriginal;
			editor._saveCheck = saveCheckOriginal;
			editor._prepareSave = prepareSaveOriginal;
		}
	});
})(window);
