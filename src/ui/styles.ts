const STYLE_ID = 'anyview-styles';

const LIGHT = `--av-bg:#f3f4f6;--av-surface:#ffffff;--av-fg:#1f2328;--av-muted:#656d76;--av-border:#d0d7de;--av-accent:#0969da;--av-hover:rgba(125,135,150,.18);--av-code-bg:#f6f8fa;--av-tok-comment:#6e7781;--av-tok-string:#0a3069;--av-tok-number:#0550ae;--av-tok-keyword:#cf222e;--av-tok-function:#8250df;--av-tok-tag:#116329;--av-tok-attr:#953800;--av-grid-head:#f6f8fa;--av-shadow:0 1px 4px rgba(0,0,0,.18);`;
const DARK = `--av-bg:#0d1117;--av-surface:#161b22;--av-fg:#e6edf3;--av-muted:#8b949e;--av-border:#30363d;--av-accent:#58a6ff;--av-hover:rgba(140,150,165,.22);--av-code-bg:#1c2128;--av-tok-comment:#8b949e;--av-tok-string:#a5d6ff;--av-tok-number:#79c0ff;--av-tok-keyword:#ff7b72;--av-tok-function:#d2a8ff;--av-tok-tag:#7ee787;--av-tok-attr:#ffa657;--av-grid-head:#1c2128;--av-shadow:0 1px 4px rgba(0,0,0,.6);`;

/* Every rule is scoped under .av-root so the viewer never restyles the host page. */
export const CSS = `
.av-root{${LIGHT}position:relative;display:flex;flex-direction:column;width:100%;height:100%;min-height:120px;box-sizing:border-box;background:var(--av-bg);color:var(--av-fg);font:14px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;overflow:hidden;text-align:left}
.av-root[data-theme=dark]{${DARK}color-scheme:dark}
@media (prefers-color-scheme:dark){.av-root[data-theme=auto]{${DARK}color-scheme:dark}}
.av-root *,.av-root *::before,.av-root *::after{box-sizing:border-box}
.av-root:fullscreen{width:100vw;height:100vh}
.av-toolbar{display:flex;align-items:center;gap:8px;flex:0 0 auto;min-height:40px;padding:4px 8px;background:var(--av-surface);border-bottom:1px solid var(--av-border)}
.av-toolbar-group{display:flex;align-items:center;gap:2px}
.av-toolbar-title{flex:1 1 0;min-width:0;display:flex;align-items:baseline;gap:8px;overflow:hidden}
.av-file-name{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.av-file-kind{color:var(--av-muted);font-size:12px;white-space:nowrap}
.av-toolbar-info{min-width:56px;text-align:center;font-size:13px;color:var(--av-muted);font-variant-numeric:tabular-nums}
.av-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;min-width:30px;height:30px;padding:0 6px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;cursor:pointer}
.av-btn:hover:not(:disabled){background:var(--av-hover)}
.av-btn:focus-visible{outline:2px solid var(--av-accent);outline-offset:1px}
.av-btn:disabled{opacity:.4;cursor:default}
.av-btn-text{padding:0 12px;border:1px solid var(--av-border);background:var(--av-surface)}
.av-body{position:relative;flex:1 1 auto;min-height:0;display:flex;flex-direction:column}
.av-content{flex:1 1 auto;min-height:0;display:flex;flex-direction:column;overflow:hidden}
.av-root [hidden]{display:none!important}
.av-status{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;padding:24px;text-align:center;background:var(--av-bg);overflow:auto}
.av-status-title{font-size:16px;font-weight:600;margin:0}
.av-status-text{color:var(--av-muted);margin:0;max-width:520px;overflow-wrap:anywhere}
.av-status-icon{color:var(--av-muted)}
.av-spinner{width:28px;height:28px;border:3px solid var(--av-border);border-top-color:var(--av-accent);border-radius:50%;animation:av-spin .8s linear infinite}
@keyframes av-spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion:reduce){.av-spinner{animation-duration:2.4s}}
.av-child{position:absolute;inset:0;z-index:5;display:flex;flex-direction:column;background:var(--av-bg)}
.av-child-bar{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--av-surface);border-bottom:1px solid var(--av-border)}
.av-child-body{flex:1 1 auto;min-height:0}
.av-scroll{flex:1 1 auto;min-height:0;overflow:auto}
.av-notice{flex:0 0 auto;padding:6px 12px;font-size:12px;color:var(--av-muted);background:var(--av-surface);border-bottom:1px solid var(--av-border)}

.av-image{flex:1 1 auto;min-height:0;overflow:auto;display:flex;padding:16px;background-color:var(--av-bg);background-image:linear-gradient(45deg,var(--av-hover) 25%,transparent 25%,transparent 75%,var(--av-hover) 75%),linear-gradient(45deg,var(--av-hover) 25%,transparent 25%,transparent 75%,var(--av-hover) 75%);background-size:20px 20px;background-position:0 0,10px 10px}
.av-image img{margin:auto;display:block;flex:0 0 auto;max-width:none;image-rendering:auto}
.av-image img.av-fit{max-width:100%;max-height:100%;object-fit:contain}
.av-media{flex:1 1 auto;min-height:0;display:flex;align-items:center;justify-content:center;padding:16px;background:#000}
.av-media video{max-width:100%;max-height:100%;outline:0}
.av-media-audio{background:var(--av-bg)}
.av-media audio{width:min(560px,100%)}
.av-frame{flex:1 1 auto;min-height:0;width:100%;border:0;background:#fff}

.av-code{flex:1 1 auto;min-height:0;overflow:auto;background:var(--av-surface)}
.av-code-table{display:flex;min-width:max-content;min-height:100%;font:13px/1.55 ui-monospace,SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace}
.av-code-gutter{flex:0 0 auto;margin:0;padding:12px 10px;text-align:right;color:var(--av-muted);background:var(--av-code-bg);border-right:1px solid var(--av-border);user-select:none;position:sticky;left:0;white-space:pre}
.av-code-lines{flex:1 1 auto;margin:0;padding:12px 16px;white-space:pre;tab-size:4}
.av-code-gutter,.av-code-lines{font:inherit}
.av-tok-comment{color:var(--av-tok-comment);font-style:italic}
.av-tok-string{color:var(--av-tok-string)}
.av-tok-number,.av-tok-literal{color:var(--av-tok-number)}
.av-tok-keyword{color:var(--av-tok-keyword)}
.av-tok-function{color:var(--av-tok-function)}
.av-tok-tag{color:var(--av-tok-tag)}
.av-tok-attr{color:var(--av-tok-attr)}

.av-prose{max-width:860px;margin:0 auto;padding:32px 40px;background:var(--av-surface);min-height:100%;overflow-wrap:break-word}
.av-prose h1,.av-prose h2,.av-prose h3,.av-prose h4,.av-prose h5,.av-prose h6{margin:1.4em 0 .6em;line-height:1.25;font-weight:600}
.av-prose h1{font-size:2em;padding-bottom:.3em;border-bottom:1px solid var(--av-border)}
.av-prose h2{font-size:1.5em;padding-bottom:.3em;border-bottom:1px solid var(--av-border)}
.av-prose h3{font-size:1.25em}.av-prose h4{font-size:1em}.av-prose h5{font-size:.9em}.av-prose h6{font-size:.85em;color:var(--av-muted)}
.av-prose>:first-child{margin-top:0}
.av-prose p,.av-prose ul,.av-prose ol,.av-prose blockquote,.av-prose pre,.av-prose table{margin:0 0 1em}
.av-prose ul,.av-prose ol{padding-left:2em}
.av-prose li.av-task{list-style:none;margin-left:-1.4em}
.av-prose li.av-task input{margin-right:.5em}
.av-prose a{color:var(--av-accent)}
.av-prose img{max-width:100%;height:auto}
.av-prose blockquote{padding:0 1em;color:var(--av-muted);border-left:4px solid var(--av-border)}
.av-prose code{font:85%/1.45 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;background:var(--av-code-bg);padding:.2em .4em;border-radius:4px}
.av-prose pre{padding:12px 16px;overflow:auto;background:var(--av-code-bg);border-radius:6px}
.av-prose pre code{padding:0;background:none;font-size:13px;white-space:pre}
.av-prose hr{height:1px;border:0;background:var(--av-border);margin:1.5em 0}
.av-prose table{border-collapse:collapse;display:block;overflow:auto;max-width:100%}
.av-prose th,.av-prose td{border:1px solid var(--av-border);padding:6px 12px}
.av-prose th{background:var(--av-code-bg);font-weight:600}

.av-grid{flex:1 1 auto;min-height:0;overflow:auto;background:var(--av-surface);position:relative;font-size:13px}
.av-grid-canvas{position:relative}
.av-grid-colhead{position:sticky;top:0;z-index:4;background:var(--av-grid-head)}
.av-grid-rowhead{position:sticky;left:0;z-index:3;background:var(--av-grid-head)}
.av-grid-corner{position:sticky;left:0;z-index:5;background:var(--av-grid-head);border-right:1px solid var(--av-border);border-bottom:1px solid var(--av-border)}
.av-grid-cells{position:absolute}
.av-grid-overlays{position:absolute;z-index:2;pointer-events:none}
.av-grid-overlay{position:absolute;overflow:hidden}
.av-grid-overlay>*{display:block;width:100%;height:100%}
.av-grid-h,.av-grid-c{position:absolute;overflow:hidden;white-space:nowrap;border-right:1px solid var(--av-border);border-bottom:1px solid var(--av-border);padding:0 6px}
.av-grid-h{padding:0 2px}
.av-grid-h{color:var(--av-muted);text-align:center;text-overflow:ellipsis;background:var(--av-grid-head);user-select:none}
.av-grid-c{display:flex;align-items:center;background:var(--av-surface)}
.av-grid-t{position:relative;flex:1 1 auto;min-width:0;max-height:100%;overflow:hidden;text-overflow:ellipsis}
.av-grid-wrap{white-space:pre-wrap;overflow-wrap:anywhere}
.av-grid-fill{display:flex;justify-content:space-between;white-space:pre}
.av-grid-spill{overflow:visible}
.av-grid-spill>.av-grid-t{z-index:1}
.av-grid-bar{position:absolute;top:2px;bottom:2px;opacity:.75;border-radius:1px}
.av-grid-icon{position:relative;flex:0 0 auto;margin-right:3px}
.av-grid-filter{position:relative;flex:0 0 auto;display:flex;align-items:center;justify-content:center;width:1.3em;height:1.3em;margin-left:2px;border:1px solid #c6c6c6;border-radius:3px;background:#fff;color:#444}
.av-grid-sheet{background:#fff}
.av-grid-sheet .av-grid-cells{--av-surface:#fff;--av-border:#e1e1e1;color:#000}
.av-grid-sheet .av-grid-c{padding:0 3px 1px;line-height:1.2}
.av-grid-sheet .av-grid-t{white-space:pre;text-overflow:clip}
.av-grid-sheet .av-grid-t.av-grid-wrap{white-space:pre-wrap}
.av-tabs{flex:0 0 auto;display:flex;gap:2px;padding:4px 8px;overflow-x:auto;background:var(--av-surface);border-top:1px solid var(--av-border)}
.av-tab{flex:0 0 auto;height:28px;padding:0 12px;border:0;border-radius:6px;background:transparent;color:var(--av-muted);font:inherit;font-size:13px;cursor:pointer;white-space:nowrap}
.av-tab:hover{background:var(--av-hover)}
.av-tab[aria-selected=true]{background:var(--av-hover);color:var(--av-fg);font-weight:600}
.av-tab:focus-visible{outline:2px solid var(--av-accent);outline-offset:-2px}

.av-pages{flex:1 1 auto;min-height:0;overflow:auto;padding:16px;display:flex;flex-direction:column;align-items:center;gap:16px}
.av-page{flex:0 0 auto;position:relative;background:#fff;color:#000;box-shadow:var(--av-shadow)}
.av-paper{flex:0 0 auto;background:#fff;color:#000;box-shadow:var(--av-shadow);margin:0 auto;max-width:100%;overflow-wrap:break-word;font-family:Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:11pt;line-height:1.3;transform-origin:top center}
.av-paper p{margin:0;min-height:1em}
.av-paper table{border-collapse:collapse;max-width:100%}
.av-paper td,.av-paper th{vertical-align:top;padding:2px 6px}
.av-paper img{max-width:100%;height:auto}
.av-paper a{color:#0563c1}
.av-paper h1,.av-paper h2,.av-paper h3,.av-paper h4,.av-paper h5,.av-paper h6{margin:0;font-size:inherit;font-weight:inherit}
.av-paper ul,.av-paper ol{margin:0;padding-left:2em}
.av-doc{flex:0 0 auto;display:flex;flex-direction:column;align-items:center;gap:16px;color:#000;line-height:1.2}
.av-doc-page{position:relative;z-index:0;display:flex;flex-direction:column;overflow:hidden;max-width:none;margin:0;line-height:inherit;font-size:inherit;font-family:inherit}
.av-doc-body{flex:1 1 auto;min-height:0;display:flow-root;position:relative;z-index:1}
.av-doc-header,.av-doc-footer{position:absolute}
.av-doc-footer{display:flex;flex-direction:column;justify-content:flex-end;pointer-events:none}
.av-doc-footer>*{pointer-events:auto}
.av-doc-footnotes{position:relative;z-index:1}
.av-doc-footnotes{flex:0 0 auto}
.av-doc p{min-height:0}
.av-doc a{color:inherit;text-decoration:inherit}
.av-doc a[href]{cursor:pointer}
.av-doc-tab{display:inline-flex;min-width:0;text-indent:0;vertical-align:baseline}
.av-doc-tab>span{flex:1 1 0;min-width:0;overflow:hidden;white-space:pre;tab-size:0;direction:rtl}
.av-doc-tab[data-av-leader-style=dot]>span::before{content:"................................................................................................................................................................................................................................................"}
.av-doc-tab[data-av-leader-style=hyphen]>span::before{content:"----------------------------------------------------------------------------------------------------------------------------------------------------------------"}
.av-doc-tab[data-av-leader-style=line]>span::before{content:"________________________________________________________________________________________________________________________________________________________________"}
.av-doc-table{border-collapse:collapse;max-width:none!important}
.av-doc-table td{overflow-wrap:break-word}
.av-doc-gridlines .av-doc-table td[data-av-g]{position:relative}
.av-doc-gridlines .av-doc-table td[data-av-g]::after{content:"";position:absolute;inset:0;border:0 dashed #9aa0a6;pointer-events:none}
.av-doc-gridlines .av-doc-table td[data-av-g*=t]::after{border-top-width:1px}
.av-doc-gridlines .av-doc-table td[data-av-g*=l]::after{border-left-width:1px}
.av-doc-gridlines .av-doc-table td[data-av-g*=b]::after{border-bottom-width:1px}
.av-doc-gridlines .av-doc-table td[data-av-g*=r]::after{border-right-width:1px}
.av-slide{position:absolute;top:0;left:0;transform-origin:0 0;overflow:hidden;background:#fff;color:#000;font-family:Calibri,Carlito,Arial,sans-serif}
.av-slide-shape{position:absolute;display:flex;flex-direction:column;overflow:visible}
.av-slide-shape p{margin:0;white-space:pre-wrap;overflow-wrap:break-word}
.av-slide-shape img{width:100%;height:100%;display:block}
.av-slide table{border-collapse:collapse;width:100%;height:100%}
.av-slide td{border:1px solid #999;padding:4px 8px;vertical-align:top}

.av-list{width:100%;border-collapse:collapse;background:var(--av-surface)}
.av-list th,.av-list td{padding:6px 12px;border-bottom:1px solid var(--av-border);text-align:left;white-space:nowrap}
.av-list th{position:sticky;top:0;background:var(--av-grid-head);font-weight:600}
.av-list td:first-child{white-space:normal;overflow-wrap:anywhere;width:100%}
.av-link{border:0;padding:0;background:none;color:var(--av-accent);font:inherit;cursor:pointer;text-align:left}
.av-link:hover{text-decoration:underline}
.av-mail{flex:1 1 auto;min-height:0;display:flex;flex-direction:column;background:var(--av-surface)}
.av-mail-head{flex:0 0 auto;padding:12px 16px;border-bottom:1px solid var(--av-border);display:grid;grid-template-columns:auto 1fr;gap:2px 12px;max-height:40%;overflow:auto}
.av-mail-head dt{color:var(--av-muted)}
.av-mail-head dd{margin:0;overflow-wrap:anywhere}
.av-mail-attachments{flex:0 0 auto;padding:8px 16px;border-bottom:1px solid var(--av-border);display:flex;flex-wrap:wrap;gap:4px 16px}
.av-mail-text{flex:1 1 auto;min-height:0;overflow:auto;margin:0;padding:16px;white-space:pre-wrap;overflow-wrap:anywhere;font:inherit}
.av-nb-cell{margin:0 0 16px}
.av-nb-code,.av-nb-output{margin:0;padding:10px 14px;overflow:auto;border-radius:6px;font:13px/1.5 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;white-space:pre}
.av-nb-code{background:var(--av-code-bg);border:1px solid var(--av-border)}
.av-nb-output{white-space:pre-wrap;overflow-wrap:anywhere}
.av-nb-error{color:#cf222e}
.av-nb-cell img{max-width:100%;background:#fff}
.av-font{padding:24px 32px;background:var(--av-surface);min-height:100%;overflow-wrap:anywhere}
.av-font p{margin:0 0 16px;line-height:1.3}
.av-font small{display:block;color:var(--av-muted);font:12px/1.4 system-ui,sans-serif;margin-bottom:2px}
`;

/** Injects the stylesheet once per document. */
export function ensureStyles(doc: Document, nonce?: string): void {
  if (doc.getElementById(STYLE_ID)) return;
  const style = doc.createElement('style');
  style.id = STYLE_ID;
  if (nonce) style.setAttribute('nonce', nonce);
  style.textContent = CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}
