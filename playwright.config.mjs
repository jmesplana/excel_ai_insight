import { defineConfig } from '@playwright/test';
export default defineConfig({
    testDir: './tests/browser',
    timeout: 45000,
    use: { baseURL: 'http://127.0.0.1:8091', channel: 'chrome', headless: true },
    webServer: {
        command: 'venv/bin/python app.py',
        env: { PORT: '8091', HOST: '127.0.0.1', FLASK_DEBUG: '0' },
        url: 'http://127.0.0.1:8091',
        reuseExistingServer: false
    }
});
