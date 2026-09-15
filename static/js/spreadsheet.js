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
