/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	const editor = AscTest.Editor;
	const drawingDocument = AscTest.DrawingDocument;
	const thumbnails = editor.WordControl.Thumbnails;
	let restorations;

	function replace(object, key, value)
	{
		const hadOwn = Object.prototype.hasOwnProperty.call(object, key);
		const previous = object[key];
		object[key] = value;
		restorations.push(function ()
		{
			if (hadOwn)
				object[key] = previous;
			else
				delete object[key];
		});
	}

	function createFixture()
	{
		const presentation = new AscCommonSlide.CPresentation(drawingDocument, true);
		presentation.Api = editor;
		replace(editor.WordControl, 'm_oLogicDocument', presentation);
		replace(drawingDocument, 'm_oLogicDocument', presentation);
		replace(thumbnails, 'm_oWordControl', drawingDocument);
		replace(thumbnails, 'm_arrPages', []);
		replace(editor, 'SelectedObjectsStack', []);
		replace(editor, 'sendEvent', function () {});
		replace(drawingDocument, 'DrawImageTextureFillSlide', function () {});
		replace(presentation, 'Recalculate', function () {});
		presentation.createNecessaryObjectsIfNoPresent();
		presentation.CurPage = -1;
		presentation.addNextSlideAction(0);
		presentation.CurPage = 0;
		thumbnails.CalculatePlaces();

		const realUpdate = presentation.Document_UpdateInterfaceState;
		const realSync = AscCommon.SlideEditorApi.prototype.sync_slidePropCallback;
		const updates = [];
		const properties = [];
		replace(presentation, 'Document_UpdateInterfaceState', function ()
		{
			const result = realUpdate.apply(this, arguments);
			updates.push(thumbnails.GetSelectedArray());
			return result;
		});
		replace(editor, 'sync_slidePropCallback', function (slide)
		{
			realSync.call(this, slide);
			const selected = this.SelectedObjectsStack[this.SelectedObjectsStack.length - 1];
			properties.push({isHidden: selected.Value.isHidden});
		});
		return {presentation: presentation, updates: updates, properties: properties};
	}

	QUnit.module('Hidden slide properties after hideSlides', {
		beforeEach: function ()
		{
			restorations = [];
			AscCommon.History.Clear();
		},
		afterEach: function ()
		{
			restorations.reverse().forEach(function (restore) { restore(); });
			AscCommon.History.Clear();
		}
	});

	QUnit.test('hide_slides_sends_properties_once', function (assert)
	{
		const fixture = createFixture();
		fixture.presentation.hideSlides(true, [0]);
		assert.deepEqual(fixture.properties.map(function (item) { return item.isHidden; }), [true], 'hide reports the hidden state once');
		fixture.presentation.hideSlides(false, [0]);
		assert.deepEqual(fixture.properties.map(function (item) { return item.isHidden; }), [true, false], 'show reports the visible state once');
		assert.deepEqual(fixture.updates, [[0], [0]], 'each command preserves the selected thumbnail');
	});

	QUnit.test('hide_slides_locked_sends_nothing', function (assert)
	{
		const fixture = createFixture();
		replace(fixture.presentation, 'Document_Is_SelectionLocked', function () { return true; });
		fixture.presentation.hideSlides(true, [0]);
		assert.true(fixture.presentation.Slides[0].isVisible(), 'locked slide stays visible');
		assert.deepEqual(fixture.updates, [], 'locked command sends no properties');
		assert.deepEqual(fixture.properties, [], 'locked command sends no slide state');
	});

	QUnit.test('undo_redo_hide_keep_existing_notifications', function (assert)
	{
		const fixture = createFixture();
		fixture.presentation.hideSlides(true, [0]);
		fixture.updates.length = 0;
		fixture.properties.length = 0;
		fixture.presentation.Document_Undo();
		assert.true(fixture.presentation.Slides[0].isVisible(), 'undo restores visibility');
		fixture.presentation.Document_Redo();
		assert.false(fixture.presentation.Slides[0].isVisible(), 'redo restores the hidden state');
		assert.deepEqual(fixture.properties.map(function (item) { return item.isHidden; }), [false, true], 'undo and redo keep their direct properties');
		assert.strictEqual(fixture.updates.length, 2, 'undo and redo send one direct update each');
	});
});
