import {buildJevState} from './jev-config.js';

/** Reuse within one frozen run only, keyed by every selected source value. */
export class JevRowReuse {
    constructor(run, enabled = run.config.execution?.reuseIdenticalRows === true) {
        this.run = run;
        this.enabled = enabled;
        this.completed = new Map();
        this.outputs = [...run.config.questions, ...(run.config.derived || [])].map(q => q.outputColumnName);
        if (enabled) for (let i = 0; i < run.cursor; i++) this.remember(i);
    }

    state(i) {
        return buildJevState(this.run.rows[i], this.run.config.sourceColumns, this.run.config.execution?.structuredState);
    }

    key(i) { return JSON.stringify(this.state(i)); }

    remember(i) {
        if (!this.enabled) return;
        const decisions = this.run.audit[i]?.decisions;
        if (this.outputs.every(name => ['ok', 'review', 'empty'].includes(decisions?.[name]?.status)
            && (decisions[name].value != null || decisions[name].status === 'empty'))) {
            this.completed.set(this.key(i), i);
        }
    }

    plan(indices, retryFailed = false) {
        const seen = new Map(), rows = [], reused = new Map();
        for (const i of indices) {
            const state = this.state(i), key = JSON.stringify(state);
            if (this.enabled && (seen.has(key) || this.completed.has(key))) reused.set(i, seen.get(key) ?? this.completed.get(key));
            else {
                rows.push({rowIndex: i, state, previous: retryFailed ? this.run.audit[i]?.decisions : undefined});
                seen.set(key, i);
            }
        }
        return {rows, reused};
    }

    copy(from, to) {
        const values = Object.fromEntries(this.run.outputColumns.map(name => [name, this.run.results[from][name]]));
        this.run.results[to] = {...this.run.rows[to], ...values};
        this.run.audit[to] = {rowIndex: to, decisions: this.run.audit[from].decisions, calls: [], reusedFrom: from};
    }
}
