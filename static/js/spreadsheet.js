/** Workbook parsing and results export, independent of UI state. */
export function parseWorkbook(workbook, xlsx = XLSX) {
    const sheets = {};
    for (const name of workbook.SheetNames) {
        const ws = workbook.Sheets[name];
        const matrix = xlsx.utils.sheet_to_json(ws, { header: 1, defval: null });
        if (!matrix.length) continue;
        const width = matrix.reduce((max, row) => Math.max(max, row.length), 0);
        const header = matrix[0];
        const columns = Array.from({ length: width }, (_, i) => String(header[i] ?? '').trim());
        if (columns.some(c => !c)) throw new Error(`Sheet "${name}" has blank headers. Give each column a name.`);
        if (new Set(columns).size !== columns.length) throw new Error(`Sheet "${name}" has duplicate headers. Give each column a unique name.`);
        const data = matrix.slice(1).filter(row => row.some(v => v !== null && v !== '')).map(row =>
            Object.fromEntries(columns.map((column, i) => [column, row[i] ?? null])));
        sheets[name] = { columns, data };
    }
    if (!Object.keys(sheets).length) throw new Error('No populated sheets found in this file.');
    return sheets;
}

export function exportResults(result, filename, prefix = 'analyzed', xlsx = XLSX) {
    if (!result) return;
    const ws = xlsx.utils.json_to_sheet(result.data, { header: result.columns });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, (result.sheetName || 'Sheet1').slice(0, 31));
    const base = filename.replace(/\.(xlsx|xls|csv)$/i, '');
    xlsx.writeFile(wb, `${prefix}_${base}.xlsx`);
}

/** Escape a literal separator for use inside a RegExp character sequence. */
function escapeRegExp(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Split one cell on a separator, longest-first so "||" wins over "|".
 *
 * Separators are matched literally, never as regex, so a user typing "|" gets
 * a pipe rather than an alternation. Blank fragments are dropped: "a || || b"
 * is two categories, not four.
 */
export function splitCell(value, separators, { trim = true } = {}) {
    if (value === null || value === undefined) return [];
    const text = String(value);
    const active = separators.filter(Boolean);
    if (!active.length) return text.trim() === '' ? [] : [trim ? text.trim() : text];
    // Longest first: "||" must be consumed before "|" can match half of it.
    const ordered = [...new Set(active)].sort((a, b) => b.length - a.length);
    const pattern = new RegExp(ordered.map(escapeRegExp).join('|'), 'g');
    return text.split(pattern)
        .map(part => (trim ? part.trim() : part))
        .filter(part => part !== '');
}

/**
 * Explode a multi-value column into one row per value, copying every other
 * column down. Rows whose target cell holds a single value pass through
 * unchanged, so this is safe to run over a whole mixed sheet.
 *
 * @param {{columns: string[], data: object[]}} result - sheet to expand.
 * @param {string} column - the column holding separated values.
 * @param {string[]} separators - literal separators, e.g. ['||', '|', ';'].
 * @param {object} [options]
 * @param {string} [options.outputColumn] - rename the exploded column.
 * @param {boolean} [options.keepEmpty] - emit a row with an empty value when a
 *   cell has no values at all, instead of dropping the source row.
 * @param {boolean} [options.dedupe] - drop repeated values within one row.
 * @returns {{columns: string[], data: object[], sheetName?: string}}
 */
export function explodeColumn(result, column, separators, options = {}) {
    const { outputColumn = column, keepEmpty = true, dedupe = true } = options;
    if (!result || !Array.isArray(result.data)) throw new Error('No data to clean.');
    if (!result.columns.includes(column)) throw new Error(`Column "${column}" is not in this sheet.`);
    if (outputColumn !== column && result.columns.includes(outputColumn)) {
        throw new Error(`Column "${outputColumn}" already exists. Choose another name.`);
    }

    const columns = result.columns.map(name => (name === column ? outputColumn : name));
    // Build each output row from `columns` so a renamed column leaves no
    // stray key behind and the field order matches the header exactly.
    const buildRow = (row, value) => Object.fromEntries(
        columns.map(name => [name, name === outputColumn ? value : row[name]]));

    const data = [];
    for (const row of result.data) {
        let values = splitCell(row[column], separators);
        if (dedupe) values = [...new Set(values)];
        if (!values.length) {
            // Nothing to explode; preserve the row so counts still reconcile.
            if (keepEmpty) data.push(buildRow(row, ''));
            continue;
        }
        for (const value of values) data.push(buildRow(row, value));
    }
    return { ...result, columns, data, partial: result.partial };
}
