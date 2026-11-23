let sheet, csvChunks, csvLine, rowCount = 0;
function sendWorkbook(wb) {
    const bytes = XLSX.write(wb, {bookType:'xlsx', type:'array', compression:true});
    postMessage({file:new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})});
}
onmessage = async ({data}) => {
    try {
        if (data.audit) {
            importScripts('../vendor/xlsx.full.min.js');
            const {buildWorkbook} = await import('./pipeline-workbook.js?v=3');
            sendWorkbook(buildWorkbook(data.audit));
            return;
        }
        if (data.headers && data.format === 'csv') {
            ({csvLine} = await import('./csv-export.js'));
            csvChunks = [new Blob(['\ufeff', csvLine(data.headers)])];
        } else if (csvChunks && data.rows) {
            csvChunks.push(new Blob(data.rows.map(csvLine)));
        } else if (csvChunks) {
            postMessage({file:new Blob(csvChunks, {type:'text/csv;charset=utf-8'})});
            return;
        } else if (data.headers) {
            importScripts('../vendor/xlsx.full.min.js');
            sheet = XLSX.utils.aoa_to_sheet([data.headers], {dense:true});
        } else if (data.rows) {
            rowCount += data.rows.length;
            if (rowCount > 1048575) throw new Error('This selection exceeds the worksheet row limit.');
            for (const row of data.rows) {
                if (row.some(v => typeof v === 'string' && v.length > 32767)) throw new Error('A selected value exceeds the Excel cell length limit. The audit workbook preserves long values.');
            }
            XLSX.utils.sheet_add_aoa(sheet, data.rows, {origin:-1});
        } else {
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, sheet, 'Results');
            sendWorkbook(wb);
            return;
        }
        postMessage({ready:true});
    } catch (error) { postMessage({error:error.message}); }
};
