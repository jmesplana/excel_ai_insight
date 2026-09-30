# Bundled browser assets

These files are served locally so a fresh browser session can load the app offline.
Versions are pinned here; retain license notices when updating.

| Asset | Version / source | License |
| --- | --- | --- |
| Bootstrap CSS and JS | 5.3.2, cdn.jsdelivr.net/npm/bootstrap | MIT, licenses/bootstrap.txt |
| Bootstrap Icons CSS and fonts | 1.11.0, cdn.jsdelivr.net/npm/bootstrap-icons | MIT, licenses/bootstrap-icons.txt |
| SheetJS xlsx.full.min.js | 0.16.9, cdnjs.cloudflare.com/ajax/libs/xlsx | Apache 2.0, licenses/xlsx.txt |
| Chart.js UMD | 4.5.1, cdn.jsdelivr.net/npm/chart.js | MIT, licenses/chart.txt |
| Marked | 15.0.12, cdn.jsdelivr.net/npm/marked | MIT, licenses/marked.txt |
| Inter fonts (text-0 through text-3) | Google Fonts inter/v20, weights 300/400/500/600 | OFL, licenses/inter.txt |
| Space Grotesk fonts (text-4 through text-6) | Google Fonts spacegrotesk/v22, weights 400/500/700 | OFL, licenses/space-grotesk.txt |
| DOMPurify purify.es.mjs | Existing bundled dependency; version and license in file header | Apache 2.0 or MPL 2.0 |

`fonts.css` contains local URLs for the Google Fonts stylesheet used by this app.
`SHA256SUMS` records the downloaded files (including the pre-existing DOMPurify).
