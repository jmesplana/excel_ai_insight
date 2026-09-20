import {buildJevReport, narrativePacket, downloadJSON} from './jev-report.js';
import {escapeHtml as esc, renderMarkdown} from './rendering.js';

export function renderJevReport(result, {streamChat, getLLMConfig}) {
    const box = document.getElementById('jev-dataset-report');
    if (!box) return;
    box.classList.toggle('hidden', !result?.jev);
    if (!result?.jev) { box.innerHTML = ''; return; }
    const report = buildJevReport(result);
    const source = result.jev.sourceResult || result;
    const coverage = report.coverage;
    result.jev.report = report;
    const table = (headers, rows) => `<div class="table-responsive"><table class="table table-sm"><thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${row.map(v => `<td>${esc(v ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    box.innerHTML = `<div class="card"><div class="card-header"><strong>${esc(report.title)}</strong></div><div class="card-body">
        <p id="jev-coverage"><strong>${esc(report.scope)}</strong>: ${coverage.processed} of ${coverage.total} records processed in ${esc(report.sheet)}.
        ${coverage.successful} successful; ${coverage.review} need review; ${coverage.failed} failed; ${coverage.blocked} blocked; ${coverage.empty} empty; ${coverage.unprocessed} unprocessed.</p>
        <p class="small text-muted">All processed original records contribute to these statistics, before any later Clean Data splits. Percentages exclude missing values and failed/blocked answers; uncertain assigned values are included and counted separately. Row IDs identify imported records.</p>
        ${report.distributions.map(d => `<details class="mb-2" open><summary>${esc(d.name)} — ${d.denominator} assigned values, ${d.statuses.review} uncertain</summary>
            ${table(['Value', 'Count', '% of assigned'], d.counts.slice(0, 50).map(c => [c.value, c.count, c.percent.toFixed(1)]))}
            ${d.counts.length > 50 ? '<p>First 50 categories shown. Download the summary for all categories.</p>' : ''}
            ${d.numeric.count ? `<p>Raw ${esc(d.type)}: mean ${d.numeric.mean.toFixed(3)}, min ${d.numeric.min}, max ${d.numeric.max} (${d.numeric.count} answers).</p>` : ''}</details>`).join('')}
        ${report.groups.length ? `<details><summary>Configured group breakdowns (${report.groups.length})</summary>${table(['Group column', 'Group', 'Question', 'Value', 'Count'], report.groups.slice(0, 100).map(g => [...g.fields, g.count]))}<p>First 100 entries shown; all entries are in the download.</p></details>` : ''}
        ${report.crossTabs.map(t => `<details><summary>${t.fields.map(esc).join(' × ')} (${t.excluded} excluded)</summary>${table([...t.fields, 'Count'], t.counts.slice(0, 100).map(c => [...c.values, c.count]))}<p>First 100 entries shown; all entries are in the download.</p></details>`).join('')}
        <details class="mt-3"><summary>Review queue (${report.reviewRows.length} records)</summary>
        <div style="max-height:320px;overflow:auto">${report.reviewRows.slice(0, 200).map(r => `<p><button class="btn btn-sm btn-outline-secondary jev-show-source" data-row="${r.rowId}">Record ${r.rowId}</button> ${r.issues.map(i => `${esc(i.question)}: ${esc(i.status)} — ${esc(i.reason)}`).join('; ')}</p>`).join('')}</div>
        <p class="small">First 200 review records shown; download includes the complete queue.</p><pre id="jev-source-record" style="white-space:pre-wrap"></pre></details>
        <div class="d-flex flex-wrap gap-2 mt-3"><button id="jev-continue-results" class="btn btn-outline-secondary">Continue unfinished run</button><button id="jev-retry-results" class="btn btn-outline-secondary">Retry failed decisions</button><button id="jev-summary-download" class="btn btn-outline-primary">Download summary JSON</button>
        <button id="jev-narrative-generate" class="btn btn-primary">Generate written report</button>
        <button id="jev-narrative-download" class="btn btn-outline-secondary">Download report</button></div>
        <p class="small text-muted mt-2">The written report uses your configured OpenAI/Azure provider with aggregate statistics and selected examples. It needs those credentials separately from Jev. Statistics require no additional AI call.</p>
        <div id="jev-narrative" class="markdown-content"></div></div></div>`;
    box.querySelectorAll('.jev-show-source').forEach(button => button.onclick = () => {
        const index = source.jev.audit.findIndex(a => a.rowIndex + 1 === Number(button.dataset.row));
        box.querySelector('#jev-source-record').textContent = JSON.stringify({rowId: Number(button.dataset.row), record: source.data[index], decisions: source.jev.audit[index]?.decisions}, null, 2);
    });
    box.querySelector('#jev-continue-results').onclick = () => document.dispatchEvent(new Event('jev-continue'));
    box.querySelector('#jev-retry-results').onclick = () => document.dispatchEvent(new Event('jev-retry'));
    box.querySelector('#jev-summary-download').onclick = () => downloadJSON(report, 'dataset-summary.json');
    box.querySelector('#jev-narrative').innerHTML = renderMarkdown(result.jev.narrative || '');
    box.querySelector('#jev-narrative-download').onclick = () => {
        downloadJSON({report: result.jev.narrative || '', evidence: report}, 'dataset-findings.json');
    };
    box.querySelector('#jev-narrative-generate').onclick = async event => {
        event.target.disabled = true;
        const target = box.querySelector('#jev-narrative');
        target.textContent = 'Generating report…';
        try {
            const text = await streamChat({report: narrativePacket(report, result.jev.config),
                question: result.jev.config.report?.instructions || 'Write a concise findings report with evidence, limitations and suggested follow-up.', ...getLLMConfig()},
                full => { target.innerHTML = renderMarkdown(full); });
            result.jev.narrative = text;
        } catch (error) { target.textContent = error.message; }
        finally { event.target.disabled = false; }
    };
}
