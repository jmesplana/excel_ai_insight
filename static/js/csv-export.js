export function csvLine(values) {
    return values.map(value => {
        const text = value == null ? '' : typeof value === 'object' ? JSON.stringify(value) : String(value);
        return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
    }).join(',') + '\r\n';
}

/** CSV-only blobs avoid a workbook, per-cell objects and a giant joined string. */
export function resultsCsvBlob(rows, columns) {
    const chunks = [new Blob(['\ufeff', csvLine(columns)])];
    for (let start = 0; start < rows.length; start += 250) {
        const lines = [];
        for (let i = start; i < Math.min(start + 250, rows.length); i++) {
            lines.push(csvLine(columns.map(column => rows[i][column])));
        }
        chunks.push(new Blob(lines));
    }
    return new Blob(chunks, {type:'text/csv;charset=utf-8'});
}
