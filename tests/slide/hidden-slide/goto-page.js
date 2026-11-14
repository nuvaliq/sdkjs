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
	const goToPage = AscCommonSlide.CEditorPage.prototype.GoToPage;
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

	function createPresentation()
	{
		const presentation = new AscCommonSlide.CPresentation(drawingDocument, true);
		presentation.Api = editor;
		replace(editor.WordControl, 'm_oLogicDocument', presentation);
		replace(drawingDocument, 'm_oLogicDocument', presentation);
		replace(thumbnails, 'm_oWordControl', drawingDocument);
		replace(thumbnails, 'm_arrPages', []);
		replace(thumbnails, 'SelectPageEnabled', true);
		presentation.createNecessaryObjectsIfNoPresent();
		presentation.CurPage = -1;
		presentation.addNextSlideAction(0);
		presentation.CurPage = 0;
		presentation.addNextSlideAction(0);
		presentation.CurPage = 0;
		thumbnails.CalculatePlaces();
		replace(drawingDocument, 'SlideCurrent', 0);
		return presentation;
	}

	function createEditorPage(presentation)
	{
		replace(editor, 'isReporterMode', false);
		replace(editor, 'isViewMode', true);
		replace(editor, 'SelectedObjectsStack', []);
		replace(editor, 'isEyedropperStarted', function () { return false; });
		replace(editor, 'sendEvent', function () {});
		replace(editor, 'sync_currentPageCallback', function () {});
		replace(drawingDocument, 'TransitionSlide', {
			IsPlaying: function () { return false; },
			End: function () {}
		});
		replace(drawingDocument, 'DrawImageTextureFillSlide', function () {});
		return {
			m_oApi: editor,
			m_oLogicDocument: presentation,
			m_oDrawingDocument: drawingDocument,
			DemonstrationManager: editor.WordControl.DemonstrationManager,
			m_oScrollVerApi: {scrollToY: function () {}},
			m_oScrollHorApi: {scrollToX: function () {}},
			Thumbnails: thumbnails,
			SlideDrawer: {CheckSlide: function () {}},
			IsNotesSupported: function () { return false; },
			m_oNotesApi: null,
			m_oAnimPaneApi: null,
			CheckLayouts: function () {},
			CalculateDocumentSize: function () {},
			CreateBackgroundHorRuler: function () {},
			CreateBackgroundVerRuler: function () {},
			OnCalculatePagesPlace: function () {},
			OnScroll: function () {},
			IsGoToPageMAXPosition: false,
			m_dScrollY_Central: 0,
			m_dScrollY_max: 0,
			m_bIsHorScrollVisible: false
		};
	}

	function captureProperties(presentation)
	{
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
			properties.push({
				isHidden: selected.Value.isHidden,
				selected: thumbnails.GetSelectedArray()
			});
		});
		return {updates: updates, properties: properties};
	}

	function createFixture()
	{
		const presentation = createPresentation();
		return {
			presentation: presentation,
			page: createEditorPage(presentation),
			capture: captureProperties(presentation)
		};
	}

	QUnit.module('GoToPage hidden slide properties', {
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

	QUnit.test('goto_page_reports_final_properties_after_select', function (assert)
	{
		const fixture = createFixture();
		fixture.presentation.Slides[1].setShow(false);
		goToPage.call(fixture.page, 1);
		assert.strictEqual(fixture.presentation.CurPage, 1, 'the requested slide is current');
		assert.true(fixture.capture.updates.length >= 2, 'early notification is preserved');
		assert.deepEqual(fixture.capture.properties[0].selected, [0], 'early properties use the previous selection');
		assert.false(fixture.capture.properties[0].isHidden, 'early properties describe the visible slide');
		const last = fixture.capture.properties[fixture.capture.properties.length - 1];
		assert.deepEqual(last.selected, [1], 'final properties follow thumbnail selection');
		assert.true(last.isHidden, 'final properties describe the hidden slide');
	});

	QUnit.test('goto_page_zoom_keeps_existing_notifications', function (assert)
	{
		const fixture = createFixture();
		goToPage.call(fixture.page, 0, true);
		let nested = false;
		fixture.page.CalculateDocumentSize = function ()
		{
			if (!nested)
			{
				nested = true;
				goToPage.call(fixture.page, 0, true);
			}
		};
		goToPage.call(fixture.page, 0, true);
		assert.strictEqual(fixture.capture.updates.length, 0, 'zoom adds no interface properties');
		assert.deepEqual(thumbnails.GetSelectedArray(), [0], 'zoom preserves the selection');
	});

	QUnit.test('goto_page_keeps_preset_selection', function (assert)
	{
		const fixture = createFixture();
		thumbnails.m_arrPages[1].IsSelected = true;
		thumbnails.SelectPageEnabled = false;
		goToPage.call(fixture.page, 1);
		assert.deepEqual(thumbnails.GetSelectedArray(), [0, 1], 'caller selection is preserved');
	});

	QUnit.test('goto_page_show_exits_unchanged', function (assert)
	{
		const fixture = createFixture();
		const demonstrationCalls = [];
		replace(fixture.page.DemonstrationManager, 'Mode', true);
		replace(editor, 'DemonstrationGoToSlide', function (pageNum) { demonstrationCalls.push(pageNum); });
		goToPage.call(fixture.page, 1, true);
		goToPage.call(fixture.page, 1, false, true);
		goToPage.call(fixture.page, 1);
		assert.deepEqual(demonstrationCalls, [1], 'only the ordinary transition enters the show');
		assert.deepEqual(thumbnails.GetSelectedArray(), [0], 'show exits before thumbnail selection');
		assert.strictEqual(fixture.capture.updates.length, 0, 'show exits send no editor properties');
	});

	QUnit.test('goto_page_empty_selects_nothing', function (assert)
	{
		const fixture = createFixture();
		goToPage.call(fixture.page, -1);
		assert.strictEqual(fixture.presentation.CurPage, -1, 'empty presentation keeps page -1');
		assert.deepEqual(thumbnails.GetSelectedArray(), [0], 'page -1 does not select a thumbnail');
		assert.strictEqual(fixture.capture.updates.length, 2, 'existing page notifications are preserved');
	});
});
