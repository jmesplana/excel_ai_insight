/** Exact aggregation over every processed row; never the chat sample. */
export function buildJevReport(result) {
    result = result.jev.sourceResult || result;
    const config = result.jev.config;
    const definitions = [...config.questions, ...(config.derived || [])];
    const coverage = {total: result.jev.totalRows, selected: result.jev.selectedRows,
        processed: result.data.length, successful: 0, review: 0, failed: 0, empty: 0, blocked: 0};
    const reviewRows = [];
    const columns = definitions.map(def => ({name: def.outputColumnName, type: def.questionType || def.type || 'lookup',
        instructions: def.instructions, options: def.options, reviewPolicy: def.review,
        statuses: {ok: 0, review: 0, error: 0, empty: 0, blocked: 0}, counts: new Map(),
        numeric: {count: 0, sum: 0, min: null, max: null}, expectedCounts: new Map()}));
    const groups = new Map();
    const crossTabs = (config.report?.crossTabs || []).map(fields => ({fields, counts: new Map(), excluded: 0}));
    const models = new Set();
    let inputTokens = 0;
    const examples = [];
    const exampleKeys = new Set();
    const evidenceColumns = config.report?.evidenceColumns || config.sourceColumns;
    const maxExamples = config.report?.maxExamples ?? 12;
    for (let index = 0; index < result.data.length; index++) {
        const row = result.data[index];
        const audit = result.jev.audit[index];
        const decisions = audit?.decisions || {};
        const statuses = Object.values(decisions).map(d => d.status);
        const status = !statuses.length || statuses.includes('error') ? 'failed' : statuses.every(s => s === 'empty') ? 'empty'
            : statuses.includes('blocked') ? 'blocked' : statuses.includes('review') ? 'review' : 'successful';
        coverage[status]++;
        if (['failed', 'blocked', 'review'].includes(status)) reviewRows.push({rowId: (audit?.rowIndex ?? index) + 1,
            issues: Object.entries(decisions).filter(([, d]) => ['error', 'blocked', 'review'].includes(d.status))
                .map(([name, d]) => ({question: name, status: d.status, reason: d.reason, value: d.value}))});
        for (const call of audit?.calls || []) {
            if (call.model) models.add(call.model);
            if (Number.isFinite(call.usage?.input_tokens)) inputTokens += call.usage.input_tokens;
        }
        for (const column of columns) {
            const decision = decisions[column.name] || {status: 'error'};
            column.statuses[decision.status]++;
            if (!['ok', 'review'].includes(decision.status)) continue;
            const value = decision.value;
            if (value !== null && value !== undefined) column.counts.set(String(value), (column.counts.get(String(value)) || 0) + 1);
            const number = decision.detail ?? (column.type === 'composite' ? value : null);
            if (typeof number === 'number' && Number.isFinite(number)) {
                column.numeric.count++; column.numeric.sum += number;
                column.numeric.min = column.numeric.min === null ? number : Math.min(column.numeric.min, number);
                column.numeric.max = column.numeric.max === null ? number : Math.max(column.numeric.max, number);
            }
            for (const [key, probability] of Object.entries(decision.raw?.probabilities || {})) {
                const label = decision.labels?.[key] ?? key;
                column.expectedCounts.set(label, (column.expectedCounts.get(label) || 0) + probability);
            }
            for (const field of config.report?.groupBy || []) {
                const groupDecision = Object.hasOwn(decisions, field) ? decisions[field] : undefined;
                if (groupDecision && !['ok', 'review'].includes(groupDecision.status)) continue;
                const group = groupDecision ? groupDecision.value : row[field];
                if (group === null || group === undefined || group === '' || value === null || value === undefined) continue;
                const key = JSON.stringify([field, String(group), column.name, String(value)]);
                groups.set(key, (groups.get(key) || 0) + 1);
            }
            const exampleKey = JSON.stringify([column.name, value, decision.status]);
            if (examples.length < maxExamples && !exampleKeys.has(exampleKey)) {
                exampleKeys.add(exampleKey);
                examples.push({rowId: (audit?.rowIndex ?? index) + 1, question: column.name, value, status: decision.status,
                    evidence: Object.fromEntries(evidenceColumns.map(field => [field, String(row[field] ?? '').slice(0, 1000)]))});
            }
        }
        for (const tab of crossTabs) {
            const values = tab.fields.map(field => {
                const d = Object.hasOwn(decisions, field) ? decisions[field] : undefined;
                return d ? (['ok', 'review'].includes(d.status) ? d.value : null) : row[field];
            });
            if (values.some(v => v === null || v === undefined || v === '')) { tab.excluded++; continue; }
            const key = JSON.stringify(values.map(String));
            tab.counts.set(key, (tab.counts.get(key) || 0) + 1);
        }
    }
    coverage.unprocessed = coverage.total - coverage.processed;
    const distributions = columns.map(column => {
        const denominator = [...column.counts.values()].reduce((a, b) => a + b, 0);
        return {...column, denominator,
            counts: [...column.counts].sort((a, b) => b[1] - a[1]).map(([value, count]) => ({value, count, percent: denominator ? count / denominator * 100 : 0})),
            expectedCounts: Object.fromEntries(column.expectedCounts),
            numeric: {...column.numeric, mean: column.numeric.count ? column.numeric.sum / column.numeric.count : null}};
    });
    return {title: config.report?.title || 'Dataset summary', sheet: result.sheetName,
        runId: result.jev.runId, coverage, models: [...models], inputTokens,
        scope: coverage.unprocessed ? 'Partial dataset' : 'All rows in the selected sheet',
        definitions: config.questions, distributions,
        groups: [...groups].map(([key, count]) => ({fields: JSON.parse(key), count})),
        crossTabs: crossTabs.map(tab => ({...tab, counts: [...tab.counts].map(([key, count]) => ({values: JSON.parse(key), count}))})),
        reviewRows, examples,
        notes: ['Statistics describe original classified records, before any subsequent Clean Data splits.',
            'Percentages use non-empty classified values per question, including uncertain assigned values. Errors and blocked/empty answers are excluded.',
            'Score means use the original ordered rubric, not a physical measurement. Noul means are mean yes-probabilities.',
            'Examples illustrate selected outcomes; they are not a representative sample. Evidence text is limited to 1,000 characters per field.',
            'Row IDs refer to the 1-based imported data records, not physical Excel row numbers.',
            'Token totals cover successful API responses; failed calls may incur additional usage.']};
}

export function narrativePacket(report, config) {
    const maxCategories = config.report?.maxCategories ?? 30;
    const maxGroups = config.report?.maxGroups ?? 100;
    return {...report, reviewRows: report.reviewRows.slice(0, 20),
        reviewRowsOmitted: Math.max(0, report.reviewRows.length - 20),
        definitions: undefined,
        distributions: report.distributions.map(d => ({...d, expectedCounts: undefined,
            counts: d.counts.slice(0, maxCategories), categoriesOmitted: Math.max(0, d.counts.length - maxCategories)})),
        groups: report.groups.slice(0, maxGroups), groupsOmitted: Math.max(0, report.groups.length - maxGroups),
        crossTabs: report.crossTabs.map(t => ({...t, counts: t.counts.slice(0, maxGroups), groupsOmitted: Math.max(0, t.counts.length - maxGroups)}))};
}

export function downloadJSON(value, filename) {
    const url = URL.createObjectURL(new Blob([JSON.stringify(value, null, 2)], {type: 'application/json'}));
    const a = document.createElement('a'); a.href = url; a.download = filename; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}
