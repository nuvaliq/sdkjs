/* Run: node tests/word/unit-tests/text-to-table-autofit.cjs
 * Executes the real converter with table/environment doubles; does not verify layout geometry.
 */
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const filename = process.env.TEXT_TO_TABLE_SOURCE || path.resolve(__dirname, '../../../word/Editor/Document.js');
const modes = {Fixed: 1, Content: 2, Window: 3, Unknown: 99};
const fits = [-1, -10, -25.4, -25.4 / 72, 20, '-10'];
function converter(file) {
    const source = fs.readFileSync(file, 'utf8');
    const start = source.indexOf('CDocument.prototype.private_ConvertTextToTable = function(');
    assert.ok(start >= 0, 'real converter exists');
    const end = source.indexOf('\n};', start);
    assert.ok(end > start, 'converter terminator exists');
    return source.slice(start, end + 3);
}
function run(source, mode, fit, haveTable, inheritedLayout = 0) {
    const applied = [];
    let table, observedHaveTable = false, reset = false;
    function CDocument() {}
    const document = new CDocument();
    Object.assign(document, {
        GetCurrentSectPr: () => ({GetColumnCount: () => 1}),
        Get_PageFields: () => ({X: 10, XLimit: 175}),
        GetCompatibilityMode: () => 15
    });
    function CTable(drawing, owner, inline, rows, cols, grid) {
        table = this;
        this.inheritedLayout = inheritedLayout;
        this.grid = grid;
        this.GetRow = () => ({GetCell: () => ({GetContent: () => ({
            Internal_Content_RemoveAll() {}, Internal_Content_Add() {}
        })})});
        this.SelectAll = () => {};
        this.SetTableProps = props => {
            applied.push({...props});
            if ('TableLayout' in props) this.directLayout = props.TableLayout === 0 ? 1 : 0;
        };
    }
    function CTableProp() {}
    for (const key of ['CellsWidth', 'CellSelect', 'TableLayout']) {
        CTableProp.prototype['put_' + key] = function (value) { this[key] = value; };
    }
    function Engine() {}
    Engine.prototype = {SetConvertMode() {}, FinalizeConvert() {}, GetContentControl: () => null,
        GetRows: () => [[[{}], [{}]]]};
    const context = {CDocument, CTable, CTextToTableEngine: Engine,
        Asc: {CTableProp, c_oAscTextToTableAutoFitType: modes, c_oAscTableLayout: {AutoFit: 0, Fixed: 1}},
        AscCommon: {document_compatibility_mode_Word14: 14},
        AscCommonWord: {CSelectedElement: function (element) { this.Element = element; }}};
    vm.runInNewContext(source, context, {filename});
    const selected = {Elements: [{Element: {IsBlockLevelSdt: () => false,
        CalculateTextToTable() {}, GetParent: () => document}}],
        HaveTable() { assert.equal(reset, false); observedHaveTable = true; return haveTable; },
        Reset() { assert.ok(observedHaveTable); reset = true; this.Elements = []; },
        Add(element) { this.Elements.push(element); }};
    const props = {get_ColsCount: () => 2, get_RowsCount: () => 1,
        get_SeparatorType: () => 1, get_Separator: () => 0,
        get_AutoFitType: () => mode, get_Fit: () => fit};
    assert.equal(document.private_ConvertTextToTable(props, selected), true);
    assert.equal(selected.Elements.length, 1);
    assert.equal(selected.Elements[0].Element, table);
    return {applied, directLayout: table.directLayout,
        effectiveLayout: table.directLayout === undefined ? table.inheritedLayout : table.directLayout,
        grid: Array.from(table.grid)};
}
const source = converter(filename);
const baseline = process.env.TEXT_TO_TABLE_BASE_SOURCE && converter(process.env.TEXT_TO_TABLE_BASE_SOURCE);
let passed = 0, failed = 0;
function check(name, fn) {
    try { fn(); console.log('PASS', name); passed++; }
    catch (error) { console.error('FAIL', name, error.message); failed++; }
}
for (const fit of fits) {
    check('content_uses_explicit_autofit_and_auto_width_' + String(fit), () => {
        const result = run(source, modes.Content, fit, false);
        assert.deepEqual(result.applied, [{CellsWidth: null, CellSelect: true, TableLayout: 0}]);
        assert.equal(result.directLayout, 1);
        assert.equal(result.effectiveLayout, 1, 'explicit AutoFit overrides inherited Fixed');
    });
}
for (const [name, mode] of Object.entries(modes)) {
    for (const fit of fits) for (const haveTable of [false, true]) {
        if (mode === modes.Content && !haveTable) continue;
        check(`${name}_${String(fit)}_haveTable_${haveTable}_matches_base`, () => {
            const actual = run(source, mode, fit, haveTable);
            const applied = mode === modes.Window ? [] : [{CellsWidth: fit > 0 ? fit : null, CellSelect: true,
                ...(fit === -10 && !haveTable ? {TableLayout: 1} : {})}];
            assert.deepEqual(actual.applied, applied);
            if (baseline) assert.deepEqual(actual, run(baseline, mode, fit, haveTable));
        });
    }
}
console.log(`${passed} passed, ${failed} failed`);
process.exitCode = failed ? 1 : 0;
