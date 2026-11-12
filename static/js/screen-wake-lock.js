/** Keep the screen awake only while work is active and the page is visible. */
export class ScreenWakeLock {
    constructor({navigator = globalThis.navigator, document = globalThis.document, onStatus = () => {}} = {}) {
        this.navigator = navigator;
        this.document = document;
        this.onStatus = onStatus;
        this.active = false;
        this.session = 0;
        this.lock = null;
        this.pending = null;
        document.addEventListener('visibilitychange', () => {
            if (this.active && document.visibilityState === 'visible') void this.acquire();
        });
    }

    start() {
        if (this.active) return;
        this.active = true;
        this.session++;
        void this.acquire();
    }

    stop() {
        this.active = false;
        this.session++;
        this.pending = null;
        const lock = this.lock;
        this.lock = null;
        if (lock) void this.release(lock);
    }

    async release(lock) {
        try { await lock.release(); } catch { /* Already released by the browser. */ }
    }

    async acquire() {
        if (!this.active || this.pending !== null || (this.lock && !this.lock.released)) return;
        if (!this.navigator?.wakeLock?.request) {
            this.onStatus('Automatic screen wake protection is unavailable. Keep your device awake in system settings.');
            return;
        }
        if (this.document.visibilityState !== 'visible') {
            this.onStatus('Keep this tab visible to enable screen wake protection.');
            return;
        }
        const session = this.session;
        this.pending = session;
        this.onStatus('Requesting screen wake protection…');
        try {
            const lock = await this.navigator.wakeLock.request('screen');
            if (!this.active || session !== this.session) {
                await this.release(lock);
                return;
            }
            this.lock = lock;
            const released = () => {
                if (this.lock !== lock) return;
                this.lock = null;
                if (this.active) this.onStatus('Screen wake protection was interrupted. Keep your device awake in system settings.');
            };
            lock.addEventListener('release', released, {once: true});
            if (lock.released) released();
            else this.onStatus('Keeping your screen awake while processing.');
        } catch {
            if (this.active && session === this.session) {
                this.onStatus('Screen wake protection could not start. Keep your device awake in system settings.');
            }
        } finally {
            if (this.pending === session) this.pending = null;
        }
    }
}
