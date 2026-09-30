/** Laya requests stay on this computer; Jev requests use the hosted backend. */
export function classifierURL(path, config) {
    if (config.classificationProvider !== 'laya') return path;
    const address = config.layaEndpoint?.trim() || 'http://127.0.0.1:8001';
    let url;
    try { url = new URL(address); } catch { throw new Error('Enter a local connector address, such as http://127.0.0.1:8001.'); }
    if (!['http:', 'https:'].includes(url.protocol) || !['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)
        || url.username || url.password || url.search || url.hash || url.pathname !== '/') {
        throw new Error('The Laya connector address must be a localhost URL without a path.');
    }
    return url.origin + path;
}

export async function classifierFetch(path, payload, config = payload) {
    const local = config.classificationProvider === 'laya';
    const url = classifierURL(path, config);
    try {
        return await fetch(url, {method:'POST', headers:{'Content-Type':'application/json'},
            body:JSON.stringify(payload), ...(local ? {mode:'cors', credentials:'omit', redirect:'error',
                targetAddressSpace:'loopback'} : {})});
    } catch (error) {
        if (!local) throw error;
        throw new Error('Could not reach local Laya. Run the Laya connector, allow this website with --origin, and grant local-network access if your browser asks. Use the connector address (normally port 8001), not the raw model address (port 8000).');
    }
}
