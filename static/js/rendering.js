import DOMPurify from '../vendor/purify.es.mjs';

/** Escape untrusted values in HTML text and quoted attributes. */
export function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
}

export function renderMarkdown(text) {
    return DOMPurify.sanitize(marked.parse(String(text ?? '')), {
        USE_PROFILES: { html: true }, FORBID_TAGS: ['img', 'style', 'form', 'input'],
        FORBID_ATTR: ['style']
    });
}
