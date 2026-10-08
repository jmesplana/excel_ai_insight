/** Measure the current active session; previously completed rows are not new work. */
export class ProgressTimer {
    constructor(now = () => performance.now()) { this.now = now; }

    start({completed = 0, total = null} = {}) {
        this.startedAt = this.now();
        this.stoppedAt = null;
        this.initialCompleted = completed;
        this.completed = completed;
        this.total = total;
    }

    update(completed, total) {
        this.completed = completed;
        this.total = total;
    }

    stop() { this.stoppedAt = this.now(); }

    snapshot() {
        const elapsedMs = Math.max(0, (this.stoppedAt ?? this.now()) - this.startedAt);
        const processed = this.completed - this.initialCompleted;
        const remaining = this.total === null ? null : Math.max(0, this.total - this.completed);
        // Include slow requests/checkpointing in the observed rate. An overdue
        // batch must not count down to a false "finished" estimate.
        const remainingMs = remaining === 0 ? 0
            : remaining !== null && processed > 0 && elapsedMs >= 1000
                ? remaining * elapsedMs / processed : null;
        return {elapsedMs, remainingMs, total: this.total};
    }
}

export function formatDuration(milliseconds) {
    const seconds = Math.max(0, Math.floor(milliseconds / 1000));
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor(seconds % 3600 / 60);
    return [hours ? `${hours}h` : '', hours || minutes ? `${minutes}m` : '', `${seconds % 60}s`]
        .filter(Boolean).join(' ');
}

/** One interval for the shared overlay, cleared on completion, pause or error. */
export class ProgressTimerView {
    constructor(root = document) {
        this.root = root;
        this.timer = new ProgressTimer();
        this.interval = null;
    }

    start(options) {
        clearInterval(this.interval);
        this.timer.start(options);
        this.render();
        this.interval = setInterval(() => this.render(), 1000);
    }

    update(completed, total) {
        this.timer.update(completed, total);
        this.render();
    }

    stop() {
        clearInterval(this.interval);
        this.interval = null;
        this.timer.stop();
    }

    render() {
        const {elapsedMs, remainingMs, total} = this.timer.snapshot();
        this.root.getElementById('progress-elapsed').textContent = formatDuration(elapsedMs);
        this.root.getElementById('progress-remaining-row').classList.toggle('hidden', total === null);
        this.root.getElementById('progress-remaining').textContent = remainingMs === null ? 'Estimating…'
            : remainingMs === 0 ? 'Finishing…'
                : `About ${formatDuration(Math.max(1000, Math.ceil(remainingMs / 1000) * 1000))}`;
    }
}
