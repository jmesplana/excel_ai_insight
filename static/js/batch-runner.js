/** In-memory checkpoint. Only acknowledged batches advance the cursor. */
export class BatchRun {
    constructor({ rows, batchSize, processBatch, onProgress = () => {} }) {
        this.rows = rows;
        this.batchSize = batchSize;
        this.processBatch = processBatch;
        this.onProgress = onProgress;
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
            while (!this.complete && !this.stopped) {
                const end = Math.min(this.cursor + this.batchSize, this.rows.length);
                const batch = this.rows.slice(this.cursor, end);
                const response = await this.processBatch(batch, this.cursor);
                if (!Array.isArray(response.rows) || response.rows.length !== batch.length) {
                    throw new Error('Incomplete batch response. Resume to retry this batch.');
                }
                this.results.push(...response.rows);
                this.errors += response.errors || 0;
                this.cursor = end;
                this.onProgress(this);
            }
        } finally { this.running = false; }
        return this.results;
    }
}
