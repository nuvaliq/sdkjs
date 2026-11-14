/* Run: node tests/word/unit-tests/selection-restore.cjs
 * Executes real selection, removal and position methods with table/row/cell doubles.
 * Does not verify layout, merged cells or the complete editor history.
 */
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '../../..');
const tableFile = process.env.TABLE_SELECTION_SOURCE || path.join(root, 'word/Editor/Table.js');
const rowFile = process.env.TABLE_ROW_CHANGES_SOURCE || path.join(root, 'word/Editor/Table/TableRowChanges.js');
const positionsFile = process.env.DOCUMENT_POSITIONS_SOURCE || path.join(root, 'common/CollaborativeEditingBase.js');
const tableSource = fs.readFileSync(tableFile, 'utf8');

function extract(source, name, next) {
	const start = source.indexOf(name + ' = function');
	assert.ok(start >= 0, 'real method exists: ' + name);
	const end = source.indexOf(next, start + name.length);
	assert.ok(end > start, 'next method exists: ' + next);
	return source.slice(start + name.length + 3, end).trim().replace(/;$/, '');
}

const selectionMethod = extract(tableSource, 'CTable.prototype.SetContentSelection',
	'CTable.prototype.SetContentPosition');
const removeMethod = extract(fs.readFileSync(rowFile, 'utf8'), 'CChangesTableRowRemoveCell.prototype.Load',
	'CChangesTableRowRemoveCell.prototype.IsRelated');
const positionsMethod = extract(fs.readFileSync(positionsFile, 'utf8'),
	'CDocumentPositionsManager.prototype.Update_DocumentPositionsOnRemove',
	'CDocumentPositionsManager.prototype.OnStart_SplitRun');
const constants = {};
for (const name of ['Cell', 'Text', 'Common', 'Rows']) {
	const key = 'table_Selection_' + name;
	const match = tableSource.match(new RegExp('var ' + key + '\\s*=\\s*(0x[0-9a-f]+|[0-9]+);', 'i'));
	assert.ok(match, 'real selection constant exists: ' + key);
	constants[key] = Number(match[1]);
}
const CELL = constants.table_Selection_Cell;
const TEXT = constants.table_Selection_Text;
const COMMON = constants.table_Selection_Common;
const ROWS = constants.table_Selection_Rows;
const plain = value => JSON.parse(JSON.stringify(value));

function fixture(counts) {
	const calls = {arrays: [], data: [], content: [], removals: []};
	const context = vm.createContext({...constants});
	const compile = source => vm.runInContext('(' + source + ')', context);
	const table = {
		Content: [],
		Selection: {Use: false, StartPos: {Pos: {Row: 9, Cell: 9}}, EndPos: {Pos: {Row: 8, Cell: 8}},
			CurRow: 7, Type: TEXT, Type2: ROWS, Data2: 'old'},
		Get_Row(index) { return this.Content[index]; },
		private_SetSelectionData(value) { calls.data.push(value); },
		private_UpdateSelectedCellsArray(value) { calls.arrays.push(value); }
	};
	for (const count of counts) {
		const row = {
			Content: [],
			Get_CellsCount() { return this.Content.length; },
			Get_Cell(index) { return this.Content[index]; },
			m_oContentChanges: {Check(type, pos) { assert.equal(type, 2); return pos; }},
			Internal_ReIndexing() {}, private_CheckCurCell() {}, private_UpdateTableGrid() {}
		};
		for (let index = 0; index < count; index++) {
			const cell = {index, SetIndex(value) { this.index = value; }, Content: {
				SetContentSelection(start, end, depth, startFlag, endFlag) {
					calls.content.push({cell, start, end, depth, startFlag, endFlag});
				}
			}};
			row.Content.push(cell);
		}
		table.Content.push(row);
	}
	const manager = {m_aDocumentPositions: []};
	const update = compile(positionsMethod);
	context.AscCommon = {contentchanges_Remove: 2, CollaborativeEditing: {
		Update_DocumentPositionsOnRemove(row, pos, count) {
			calls.removals.push({row, pos, count});
			update.call(manager, row, pos, count);
		}
	}};
	return {
		table, calls,
		position(row, cell, type = CELL, type2 = COMMON) {
			const value = [{Class: table, Position: row, Type: type, Type2: type2},
				{Class: table.Content[row], Position: cell}];
			manager.m_aDocumentPositions.push(value);
			return value;
		},
		restore: compile(selectionMethod),
		remove: compile(removeMethod)
	};
}

function verify(f, start, end, expected, startFlag = 0, endFlag = 0) {
	f.restore.call(f.table, start, end, 0, startFlag, endFlag);
	const selection = f.table.Selection;
	assert.deepEqual(plain(selection.StartPos.Pos), {Row: expected.start[0], Cell: expected.start[1]});
	assert.deepEqual(plain(selection.EndPos.Pos), {Row: expected.end[0], Cell: expected.end[1]});
	assert.equal(selection.Use, true);
	assert.equal(selection.CurRow, expected.end[0]);
	assert.equal(selection.Type, expected.type === undefined ? CELL : expected.type);
	assert.equal(selection.Type2, expected.type2 === undefined ? COMMON : expected.type2);
	assert.equal(selection.Data2, null);
	assert.deepEqual(f.calls.data, [null]);
	assert.deepEqual(f.calls.arrays, expected.arrays || [false, false]);
	for (const pos of [selection.StartPos.Pos, selection.EndPos.Pos]) {
		assert.ok(pos.Row >= 0 && pos.Row < f.table.Content.length, 'restored row is in bounds');
		assert.ok(pos.Cell >= 0 && pos.Cell < f.table.Content[pos.Row].Content.length,
			'restored cell is in bounds');
	}
	if (expected.delegate || expected.type === TEXT) {
		assert.equal(f.calls.content.length, 1);
		const call = f.calls.content[0];
		assert.equal(call.cell, f.table.Content[expected.start[0]].Content[expected.start[1]]);
		assert.equal(f.table.CurCell, call.cell);
		assert.equal(call.start, start);
		assert.equal(call.end, end);
		assert.equal(call.depth, 2);
		assert.equal(call.startFlag, startFlag);
		assert.equal(call.endFlag, endFlag);
	} else {
		assert.equal(f.calls.content.length, 0);
	}
}

let passed = 0, failed = 0;
function check(name, fn) {
	try { fn(); console.log('PASS', name); passed++; }
	catch (error) { console.error('FAIL', name, error.stack); failed++; }
}

const deletedCases = [
	{name: 'deleted_last_cell_restores_end', counts: [1], start: [0, 0], end: [0, 1], target: [0, 0]},
	{name: 'both_ends_deleted', counts: [1], start: [0, 1], end: [0, 1], target: [0, 0],
		startDeleted: true, startTarget: [0, 0]},
	{name: 'deleted_middle_cell', counts: [2], start: [0, 0], end: [0, 1], target: [0, 1]},
	{name: 'multirow_end_at_first_cell', counts: [3, 2], start: [0, 0], end: [1, 0], target: [1, 0]},
	{name: 'multirow_end_at_deleted_last_cell', counts: [2, 2], start: [1, 0], end: [1, 2], target: [1, 1]},
	{name: 'ragged_rows_use_endrow', counts: [1, 4], start: [0, 0], end: [0, 1], target: [0, 0]},
	{name: 'third_row_end_uses_its_own_length', counts: [1, 3, 2], start: [0, 0], end: [2, 0], target: [2, 0]},
	{name: 'reversed_selection', counts: [1, 3], start: [1, 2], end: [0, 1], target: [0, 0]}
];
for (const spec of deletedCases) {
	check(spec.name, () => {
		const f = fixture(spec.counts);
		const start = f.position(...spec.start), end = f.position(...spec.end);
		end[1].Deleted = true;
		if (spec.startDeleted) start[1].Deleted = true;
		verify(f, start, end, {start: spec.startTarget || spec.start, end: spec.target});
	});
}

check('unchanged_non_deleted_end', () => {
	for (const type of [TEXT, CELL]) for (const type2 of [COMMON, ROWS]) {
		const f = fixture([2, 3]);
		const start = f.position(1, 1, type, type2), end = f.position(1, 1, type, type2);
		verify(f, start, end, {start: [1, 1], end: [1, 1], type, type2, delegate: true,
			arrays: type === CELL ? [type2 === ROWS] : []});
		assert.equal(f.calls.content.length, 1, 'same-cell text selection is delegated before type restoration');
	}
});

check('positions_without_selection_type', () => {
	const f = fixture([2]);
	const start = f.position(0, 0), end = f.position(0, 1);
	for (const position of [start, end]) {
		delete position[0].Type;
		delete position[0].Type2;
	}
	verify(f, start, end, {start: [0, 0], end: [0, 1], arrays: [false]});
});

for (const side of ['start', 'end', 'both']) {
	check('deleted_text_selection_' + side, () => {
		const f = fixture([1]);
		const start = f.position(0, side === 'end' ? 0 : 1, TEXT);
		const end = f.position(0, side === 'start' ? 0 : 1, TEXT);
		if (side !== 'end') start[1].Deleted = true;
		if (side !== 'start') end[1].Deleted = true;
		verify(f, start, end, {start: [0, 0], end: [0, 0], arrays: [false]});
	});
}

for (const side of ['start', 'end', 'both']) {
	check('deleted_text_selection_row_' + side, () => {
		const f = fixture([1]);
		const start = f.position(side === 'end' ? 0 : 1, 0, TEXT);
		const end = f.position(side === 'start' ? 0 : 1, 0, TEXT);
		if (side !== 'end') start[0].Deleted = true;
		if (side !== 'start') end[0].Deleted = true;
		verify(f, start, end, {start: [0, 0], end: [0, 0], arrays: [false]});
	});
}

check('deleted_row_selection_preserves_type', () => {
	const f = fixture([2, 2]);
	const start = f.position(0, 1, CELL, ROWS), end = f.position(1, 2, CELL, ROWS);
	end[1].Deleted = true;
	verify(f, start, end, {start: [0, 1], end: [1, 1], type2: ROWS, arrays: [false, true]});
});

check('collaborative_load_deleted_text_selection', () => {
	const f = fixture([2]);
	const start = f.position(0, 1, TEXT), end = f.position(0, 1, TEXT);
	const row = f.table.Content[0], removed = row.Content[1];
	f.remove.call({Class: row, Items: [removed], PosArray: [1]});
	assert.equal(start[1].Deleted, true);
	assert.equal(end[1].Deleted, true);
	assert.equal(row.Content.includes(removed), false);
	verify(f, start, end, {start: [0, 0], end: [0, 0], arrays: [false]});
});

check('start_deleted_control', () => {
	const f = fixture([1, 3]);
	const start = f.position(0, 1), end = f.position(1, 2);
	start[1].Deleted = true;
	verify(f, start, end, {start: [0, 0], end: [1, 2]});
});

check('end_row_deleted', () => {
	for (const row of [1, 2]) {
		const f = fixture([2, 4]);
		const start = f.position(0, 0), end = f.position(row, 9);
		end[0].Deleted = true;
		verify(f, start, end, {start: [0, 0], end: row === 1 ? [1, 0] : [1, 3]});
	}
});

check('boundary_flags', () => {
	const f = fixture([2, 3]);
	verify(f, null, null, {start: [0, 0], end: [1, 2], arrays: [true]}, 1, -1);
});

check('class_guard', () => {
	for (const side of [0, 1]) {
		const f = fixture([2]);
		const start = f.position(0, 0), end = f.position(0, 1);
		const selection = f.table.Selection;
		const before = plain(selection);
		(side === 0 ? start : end)[0].Class = {};
		f.restore.call(f.table, start, end, 0, 0, 0);
		assert.equal(f.table.Selection, selection);
		assert.deepEqual(plain(selection), before);
		assert.deepEqual(f.calls, {arrays: [], data: [], content: [], removals: []});
	}
});

for (const spec of [
	{name: 'collaborative_load_last_cell', counts: [2], start: [0, 0], end: [0, 1],
		remove: [0, 1], target: [0, 0]},
	{name: 'collaborative_load_ragged_rows', counts: [2, 4], start: [1, 3], end: [0, 1],
		remove: [0, 1], target: [0, 0], startAfter: [1, 3]}
]) {
	check(spec.name, () => {
		const f = fixture(spec.counts);
		const start = f.position(...spec.start), end = f.position(...spec.end);
		const row = f.table.Content[spec.remove[0]], index = spec.remove[1];
		const removed = row.Content[index];
		const survivor = row.Content[index - 1];
		assert.equal(start[1].Deleted, undefined);
		assert.equal(end[1].Deleted, undefined);
		f.remove.call({Class: row, Items: [removed], PosArray: [index]});
		assert.equal(row.Content.length, spec.counts[spec.remove[0]] - 1);
		assert.equal(row.Content.includes(removed), false);
		assert.equal(row.Content[index - 1], survivor);
		assert.equal(removed.index, -1);
		assert.deepEqual(f.calls.removals, [{row, pos: index, count: 1}]);
		assert.deepEqual(end.map(item => [item.Position, item.Deleted === true]),
			[[spec.end[0], false], [index, true]]);
		assert.deepEqual(start.map(item => [item.Position, item.Deleted === true]),
			[[spec.start[0], false], [spec.start[1], false]]);
		verify(f, start, end, {start: spec.startAfter || spec.start, end: spec.target});
	});
}

console.log(`${passed} passed, ${failed} failed`);
process.exitCode = failed ? 1 : 0;
