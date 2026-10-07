/** Await durable storage before sending another batch. Unsaved results remain downloadable. */
export class BatchRun {
    constructor({ rows, batchSize, processBatch, onProgress = () => {}, onCheckpoint = async () => {} }) {
        this.rows = rows;
        this.batchSize = batchSize;
        this.processBatch = processBatch;
        this.onProgress = onProgress;
        this.onCheckpoint = onCheckpoint;
        this.pendingCheckpoint = [];
        this.results = [];
        this.errors = 0;
        this.cursor = 0;
        this.running = false;
        this.stopped = false;
    }
    stop() { this.stopped = true; }
    get complete() { return this.cursor === this.rows.length; }
    async run() {
        if (this.running) return;
        this.running = true;
        this.stopped = false;
        try {
            await this.flushCheckpoint();
            while (!this.complete && !this.stopped) {
                const end = Math.min(this.cursor + this.batchSize, this.rows.length);
                const batch = this.rows.slice(this.cursor, end);
                const response = await this.processBatch(batch, this.cursor);
                if (!Array.isArray(response.rows) || response.rows.length !== batch.length) {
                    throw new Error('Incomplete batch response. Resume to retry this batch.');
                }
                this.pendingCheckpoint = Array.from({length: end - this.cursor}, (_, i) => this.cursor + i);
                this.results.push(...response.rows);
                this.errors += response.errors || 0;
                this.cursor = end;
                await this.flushCheckpoint();
            }
        } finally { this.running = false; }
        return this.results;
    }
    async flushCheckpoint() {
        if (!this.pendingCheckpoint.length) return;
        await this.onCheckpoint(this, this.pendingCheckpoint);
        this.pendingCheckpoint = [];
        await this.onProgress(this);
    }
}
