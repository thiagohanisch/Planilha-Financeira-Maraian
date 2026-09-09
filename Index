<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>V9 · FinTrack V8 — Persistência Segura</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.5/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.5/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.5/firebase-database-compat.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f1117;--bg2:#171b26;--bg3:#1e2333;--bg4:#252b3d;
  --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
  --text:#f0f2f8;--text2:#8b93a8;--text3:#5a6280;
  --green:#22d3a0;--green-bg:rgba(34,211,160,0.1);
  --red:#f56565;--red-bg:rgba(245,101,101,0.1);
  --blue:#60a5fa;--amber:#fbbf24;--purple:#a78bfa;
  --radius:12px;--radius-sm:8px;--sidebar:220px;
}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;line-height:1.6;-webkit-tap-highlight-color:transparent}
.app{display:flex;min-height:100vh}
.sidebar{width:var(--sidebar);background:var(--bg2);border-right:1px solid var(--border);padding:24px 0;display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;transition:transform .25s}
.logo{padding:0 20px 24px;font-size:16px;font-weight:600;letter-spacing:-.3px;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border);margin-bottom:16px;flex-shrink:0}
.logo-icon{width:30px;height:30px;background:linear-gradient(135deg,var(--green),#0ea5e9);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.sidebar-nav{display:flex;flex-direction:column;flex:1;overflow-y:auto;overflow-x:hidden}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 20px;cursor:pointer;transition:background .15s;color:var(--text2);font-size:13.5px;font-weight:400;border-left:2px solid transparent;user-select:none;flex-shrink:0}
.nav-item:hover{background:var(--bg3);color:var(--text)}
.nav-item.active{background:var(--bg3);color:var(--green);border-left-color:var(--green)}
.nav-spacer{flex:1}
.nav-icon{font-size:16px;width:20px;text-align:center;flex-shrink:0}
.main{margin-left:var(--sidebar);flex:1;padding:28px 32px;max-width:1280px}
.page-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:24px;flex-wrap:wrap;gap:12px}
.page-title{font-size:21px;font-weight:600;letter-spacing:-.4px}
.page-sub{font-size:13px;color:var(--text2);margin-top:2px}
.hdr-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.cards-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px}
.card-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px}
.card-value{font-size:22px;font-weight:600;font-family:'DM Mono',monospace;letter-spacing:-.5px}
.card-sub{font-size:12px;color:var(--text2);margin-top:3px}
.content-grid{display:grid;grid-template-columns:1fr 360px;gap:18px;align-items:start}
.chart-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:22px}
.chart-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:8px}
.chart-title{font-size:14px;font-weight:500}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--text2)}
.legend-item{display:flex;align-items:center;gap:5px}
.legend-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0}
.transactions-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.tx-header{display:flex;align-items:center;justify-content:space-between;padding:16px 18px 14px;border-bottom:1px solid var(--border);flex-wrap:wrap;gap:8px}
.tx-title{font-size:14px;font-weight:500}
.tx-list{overflow-y:auto;max-height:420px}
.tx-list::-webkit-scrollbar{width:4px}
.tx-list::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}
.tx-item{display:flex;align-items:center;gap:11px;padding:11px 18px;border-bottom:1px solid var(--border);transition:background .1s}
.tx-item:hover{background:var(--bg3)}
.tx-item:last-child{border-bottom:none}
.tx-icon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.tx-info{flex:1;min-width:0}
.tx-name{font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tx-cat{font-size:11px;color:var(--text2)}
.tx-date{font-size:11px;color:var(--text3);white-space:nowrap;margin-right:4px}
.tx-amount{font-family:'DM Mono',monospace;font-size:13px;font-weight:500;white-space:nowrap}
.tx-actions{display:flex;gap:2px;opacity:0;transition:opacity .15s}
.tx-item:hover .tx-actions{opacity:1}
.tx-btn{background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;padding:3px 6px;border-radius:5px;transition:all .15s}
.tx-btn:hover{background:var(--bg4)}
.btn{padding:7px 14px;border-radius:var(--radius-sm);border:1px solid var(--border2);background:var(--bg3);color:var(--text);font-size:12px;font-weight:500;cursor:pointer;font-family:'DM Sans',sans-serif;transition:all .15s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:hover{background:var(--bg4)}
.btn-primary{background:var(--green);color:#0f1117;border-color:var(--green)}
.btn-primary:hover{background:#1ab88c}
.btn-danger{background:var(--red-bg);color:var(--red);border-color:rgba(245,101,101,.3)}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:200;align-items:center;justify-content:center;padding:16px}
.modal-bg.open{display:flex}
.modal{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--radius);padding:26px;width:460px;max-width:100%;max-height:90vh;overflow-y:auto}
.modal-title{font-size:15px;font-weight:600;margin-bottom:18px}
.form-row{margin-bottom:13px}
.form-label{font-size:11.5px;color:var(--text2);margin-bottom:5px;display:block}
.form-input,.form-select{width:100%;padding:9px 11px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;outline:none;transition:border .15s}
.form-input:focus,.form-select:focus{border-color:var(--green)}
.form-select option{background:var(--bg3)}
.form-row-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.radio-group{display:flex;gap:8px}
.radio-btn{flex:1;padding:8px 10px;border-radius:var(--radius-sm);border:1px solid var(--border2);cursor:pointer;text-align:center;font-size:13px;font-weight:500;transition:all .15s;color:var(--text2);background:var(--bg3)}
.radio-btn.sel-income{background:var(--green-bg);border-color:var(--green);color:var(--green)}
.radio-btn.sel-expense{background:var(--red-bg);border-color:var(--red);color:var(--red)}
.modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:18px}
.period-select{padding:7px 11px;background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:13px;font-family:'DM Sans',sans-serif;cursor:pointer;outline:none}
/* Filter bar with selectable chips */
.filter-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;align-items:center}
.filter-chip{padding:5px 12px;border-radius:20px;border:1px solid var(--border);font-size:12px;cursor:pointer;background:var(--bg2);color:var(--text2);transition:all .15s;user-select:none;display:flex;align-items:center;gap:5px}
.filter-chip:hover{border-color:var(--border2);color:var(--text)}
.filter-chip.active-income{background:var(--green-bg);color:var(--green);border-color:var(--green)}
.filter-chip.active-expense{background:var(--red-bg);color:var(--red);border-color:var(--red)}
.filter-chip.active{background:var(--bg4);color:var(--text);border-color:var(--border2)}
.filter-chip .chip-x{font-size:10px;opacity:.6;margin-left:2px}
.filter-chip .chip-x:hover{opacity:1}
.filter-section{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;margin-bottom:16px}
.filter-section-title{font-size:12px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
.chip-group{display:flex;gap:6px;flex-wrap:wrap}
/* Cat bars */
.cat-list{padding:6px 0}
.cat-row{display:flex;align-items:center;gap:11px;padding:9px 20px}
.cat-bar-wrap{flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden}
.cat-bar{height:100%;border-radius:3px;transition:width .4s ease}
.cat-name-col{font-size:13px;width:130px;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cat-amt{font-family:'DM Mono',monospace;font-size:12px;color:var(--text);min-width:90px;text-align:right}
.cat-pct{font-size:11px;color:var(--text3);min-width:34px;text-align:right}
.cat-manager{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.cat-section-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px}
.cat-section-title{font-size:14px;font-weight:500;margin-bottom:14px}
.cat-item-row{display:flex;align-items:center;gap:9px;padding:7px 0;border-bottom:1px solid var(--border)}
.cat-item-row:last-child{border-bottom:none}
.cat-color-dot{width:11px;height:11px;border-radius:50%;flex-shrink:0}
.cat-item-name{flex:1;font-size:13px}
.cat-item-badge{font-size:10px;padding:2px 7px;border-radius:10px;background:var(--bg3);color:var(--text3)}
.cat-item-edit-btn{background:none;border:none;color:var(--text3);cursor:pointer;font-size:13px;opacity:.5;padding:2px 6px}
.cat-item-edit-btn:hover{color:var(--blue);opacity:1}
.cat-edit-inline{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);padding:10px 12px;margin-top:6px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.cat-total-badge{font-size:11px;font-family:'DM Mono',monospace;padding:2px 8px;border-radius:20px;background:var(--bg3);color:var(--text2)}
.cat-item-del{background:none;border:none;color:var(--text3);cursor:pointer;font-size:13px;opacity:.5;padding:2px 6px}
.cat-item-del:hover{color:var(--red);opacity:1}
.new-cat-form{display:flex;gap:7px;margin-top:14px;flex-wrap:wrap}
.color-picker{width:38px;height:34px;border:1px solid var(--border2);border-radius:var(--radius-sm);cursor:pointer;background:none;padding:2px}
.drop-zone{border:2px dashed var(--border2);border-radius:var(--radius);padding:44px 20px;text-align:center;cursor:pointer;transition:all .2s;background:var(--bg2)}
.drop-zone:hover,.drop-zone.dragover{border-color:var(--green);background:var(--green-bg)}
.drop-icon{font-size:36px;margin-bottom:10px}
.drop-title{font-size:15px;font-weight:500;margin-bottom:5px}
.drop-sub{font-size:13px;color:var(--text2)}
.import-preview{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-top:18px}
.import-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0}
.istat{background:var(--bg3);border-radius:var(--radius-sm);padding:12px;text-align:center}
.istat-val{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.istat-label{font-size:11px;color:var(--text2);margin-top:3px}
.import-table-wrap{max-height:260px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius-sm);margin-top:14px}
.import-table{width:100%;border-collapse:collapse;font-size:12px}
.import-table th{padding:8px 12px;background:var(--bg3);color:var(--text2);font-weight:500;text-align:left;position:sticky;top:0}
.import-table td{padding:8px 12px;border-bottom:1px solid var(--border)}
.import-table tr:last-child td{border-bottom:none}
.import-warn{background:rgba(251,191,36,.1);border:1px solid rgba(251,191,36,.3);border-radius:var(--radius-sm);padding:11px 14px;font-size:13px;color:var(--amber);margin-top:10px}
.cloud-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:24px;margin-bottom:20px}
.cloud-step{display:flex;gap:14px;margin-bottom:18px;align-items:flex-start}
.cloud-step-num{width:26px;height:26px;border-radius:50%;background:var(--green-bg);border:1px solid var(--green);color:var(--green);font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.cloud-step-body{flex:1}
.cloud-step-title{font-size:13.5px;font-weight:500;margin-bottom:4px}
.cloud-step-desc{font-size:12.5px;color:var(--text2);line-height:1.5}
.cloud-step-desc a{color:var(--blue);text-decoration:none}
.sync-status{display:inline-flex;align-items:center;gap:6px;font-size:12px;padding:4px 10px;border-radius:20px}
.sync-dot{width:7px;height:7px;border-radius:50%}
.sync-ok{background:var(--green-bg);color:var(--green)}
.sync-dot.ok{background:var(--green)}
.sync-err{background:var(--red-bg);color:var(--red)}
.sync-dot.err{background:var(--red)}
.sync-local{background:var(--bg3);color:var(--text2)}
.sync-dot.local{background:var(--text3)}
.section{display:none}
.section.active{display:block}
.empty{text-align:center;padding:36px 20px;color:var(--text3);font-size:13px}
.empty-icon{font-size:28px;margin-bottom:8px}
/* Insights */
.insights-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}
.insight-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px}
.insight-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.insight-value{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.insight-sub{font-size:11px;color:var(--text2);margin-top:4px}
/* Patrimônio */
.pat-input-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}
.pat-field{display:flex;flex-direction:column;gap:5px}
.pat-field label{font-size:11.5px;color:var(--text2)}
.pat-input{width:100%;padding:9px 11px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;outline:none;transition:border .15s}
.pat-input:focus{border-color:var(--green)}
.pat-input-hint{font-size:10.5px;color:var(--text3);margin-top:2px}
.pat-api-badge{display:inline-flex;align-items:center;gap:5px;font-size:11px;padding:3px 9px;border-radius:20px;margin-left:8px}
.pat-api-ok{background:var(--green-bg);color:var(--green)}
.pat-api-err{background:var(--red-bg);color:var(--red)}
.pat-api-load{background:var(--bg3);color:var(--text3)}
.pat-milestone{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:8px;font-size:12px;background:var(--bg3);border:1px solid var(--border)}
.pat-milestone.reached{background:var(--green-bg);border-color:var(--green);color:var(--green)}
.pat-table-wrap{max-height:420px;overflow-y:auto}
.pat-table{width:100%;border-collapse:collapse;font-size:12px}
.pat-table th{padding:9px 14px;background:var(--bg3);color:var(--text2);font-weight:500;text-align:right;position:sticky;top:0}
.pat-table th:first-child{text-align:left}
.pat-table td{padding:9px 14px;border-bottom:1px solid var(--border);text-align:right;font-family:'DM Mono',monospace}
.pat-table td:first-child{text-align:left;font-family:'DM Sans',sans-serif;font-weight:500;color:var(--text)}
.pat-table tr:last-child td{border-bottom:none}
.pat-table tr.highlight-year td{background:rgba(34,211,160,0.06)}
.owner-split{display:flex;height:7px;background:var(--bg3);border-radius:4px;overflow:hidden;min-width:90px}
.owner-mariana{background:var(--blue)}
.owner-luiza{background:#f59e0b}
.owner-legend{display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:var(--text2);margin-top:8px}
.owner-legend span{display:flex;align-items:center;gap:5px}
.owner-dot{width:8px;height:8px;border-radius:2px}
.recurring-tag{font-size:10px;padding:1px 6px;border-radius:10px;background:rgba(167,139,250,.13);color:#f59e0b;margin-left:4px}
.history-list{max-height:55vh;overflow:auto;border:1px solid var(--border);border-radius:var(--radius-sm)}
.history-row{display:grid;grid-template-columns:90px 1fr 110px;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid var(--border)}
.history-row:last-child{border-bottom:none}

.pat-excluded-cats{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
.cmp-badge{display:inline-flex;align-items:center;gap:4px;font-size:11px;padding:2px 8px;border-radius:20px;font-weight:500}
.cmp-up{background:rgba(34,211,160,0.15);color:var(--green)}
.cmp-down{background:rgba(245,101,101,0.15);color:var(--red)}
.cmp-flat{background:var(--bg3);color:var(--text3)}
.cmp-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}
.cmp-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px}
.cmp-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.cmp-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:5px}
.cmp-year-label{font-size:11px;color:var(--text2)}
.cmp-value{font-family:'DM Mono',monospace;font-size:13px;font-weight:500}
.cmp-divider{height:1px;background:var(--border);margin:8px 0}
.rank-row{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid var(--border)}
.rank-row:last-child{border-bottom:none}
.rank-num{width:22px;height:22px;border-radius:50%;background:var(--bg3);color:var(--text2);font-size:11px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.rank-bar-wrap{flex:1;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden}
.rank-bar{height:100%;border-radius:2px}
.mobile-bar{display:none;position:fixed;bottom:0;left:0;right:0;background:var(--bg2);border-top:1px solid var(--border);z-index:100;padding:6px 0 calc(6px + env(safe-area-inset-bottom))}
/* Lock screen */
.lock-screen{position:fixed;inset:0;background:var(--bg);z-index:9999;display:flex;align-items:center;justify-content:center}
.lock-card{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--radius);padding:40px 36px;width:360px;max-width:95vw;text-align:center}
.lock-logo{width:56px;height:56px;background:linear-gradient(135deg,var(--green),#0ea5e9);border-radius:16px;display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 20px}
.lock-title{font-size:20px;font-weight:600;margin-bottom:6px}
.lock-sub{font-size:13px;color:var(--text2);margin-bottom:28px}
.lock-input{width:100%;padding:12px 16px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:15px;outline:none;text-align:center;letter-spacing:4px;transition:border .2s}
.lock-input:focus{border-color:var(--green)}
.lock-input.err{border-color:var(--red);animation:shake .3s}
@keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-6px)}75%{transform:translateX(6px)}}
.lock-btn{width:100%;margin-top:14px;padding:11px;border-radius:var(--radius-sm);background:var(--green);color:#0f1117;font-size:14px;font-weight:600;border:none;cursor:pointer;font-family:'DM Sans',sans-serif;transition:background .15s}
.lock-btn:hover{background:#1ab88c}
.lock-err{color:var(--red);font-size:12px;margin-top:10px;min-height:18px}
.mobile-bar-inner{display:flex;justify-content:space-around}
.mob-btn{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 12px;cursor:pointer;color:var(--text3);font-size:10px;border-radius:8px;transition:color .15s;border:none;background:none}
.mob-btn.active{color:var(--green)}
.mob-btn .mob-icon{font-size:18px}
@media(max-width:768px){
  :root{--sidebar:0px}
  .sidebar{transform:translateX(-220px);width:220px}
  .sidebar.open{transform:translateX(0)}
  .main{margin-left:0;padding:16px 16px 80px}
  .cards-grid{grid-template-columns:repeat(2,1fr);gap:10px}
  .content-grid{grid-template-columns:1fr}
  .cat-manager{grid-template-columns:1fr}
  .import-stats{grid-template-columns:repeat(2,1fr)}
  .form-row-2{grid-template-columns:1fr}
  .mobile-bar{display:block}
  .insights-grid{grid-template-columns:repeat(2,1fr)}
}
@media(min-width:769px) and (max-width:1024px){
  :root{--sidebar:60px}
  .sidebar .nav-item span,.sidebar .logo span{display:none}
  .cards-grid{grid-template-columns:repeat(2,1fr)}
}
</style>
</head>
<body>
<!-- Login seguro com Firebase Authentication -->
<div id="lock-screen" class="lock-screen">
  <div class="lock-card">
    <div class="lock-logo">💰</div>
    <div class="lock-title">FinTrack</div>
    <div class="lock-sub" id="lock-sub">Entre com sua conta autorizada · V4 Mariana</div>
    <form id="firebase-login-form" onsubmit="loginWithFirebase(event)">
      <input class="form-input" id="login-email" type="email" autocomplete="username" placeholder="E-mail" required style="margin-bottom:10px;padding:12px 14px">
      <input class="form-input" id="login-password" type="password" autocomplete="current-password" placeholder="Senha" required style="padding:12px 14px">
      <button class="lock-btn" id="login-btn" type="button" onclick="loginWithFirebase(event)">Entrar</button>
    </form>
    <button type="button" onclick="resetFirebasePassword()" style="margin-top:12px;background:none;border:none;color:var(--blue);font-size:12px;cursor:pointer">Esqueci minha senha</button>
    <div class="lock-err" id="lock-err"></div>
    <div id="firebase-config-help" style="display:none;margin-top:14px;font-size:11px;color:var(--amber);text-align:left;line-height:1.5"></div>
  </div>
</div>

<div class="app" id="app-root" style="display:none">

<aside class="sidebar" id="sidebar">
  <div class="logo"><div class="logo-icon">💰</div><span>FinTrack</span></div>

  <div style="padding:0 14px 14px;border-bottom:1px solid var(--border);margin-bottom:10px">
    <div style="font-size:10.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Logado como</div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
      <span id="profile-badge" style="font-size:13px;font-weight:500">—</span>
      <a href="#" onclick="logoutProfile();return false" style="font-size:11px;color:var(--blue);text-decoration:none">Trocar</a>
    </div>
    <div style="font-size:10.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Visualizando</div>
    <div style="display:flex;gap:5px" id="view-switcher">
      <div class="view-chip" id="view-mariana" onclick="setViewOwner('mariana')" style="flex:1;text-align:center;padding:7px 6px;border-radius:7px;font-size:11px;cursor:default;border:1px solid var(--border2);background:var(--green);color:#0f1117;font-weight:600">Mariana</div>
    </div>
  </div>

  <nav class="sidebar-nav">
    <div class="nav-item active" data-section="dashboard" onclick="showSection('dashboard')"><span class="nav-icon">📊</span><span>Dashboard</span></div>
    <div class="nav-item" data-section="transacoes" onclick="showSection('transacoes')"><span class="nav-icon">↕️</span><span>Transações</span></div>
    <div class="nav-item" data-section="evolucao" onclick="showSection('evolucao')"><span class="nav-icon">📈</span><span>Evolução</span></div>
    <div class="nav-item" data-section="comparativo" onclick="showSection('comparativo')"><span class="nav-icon">🔄</span><span>Comparativo</span></div>
    <div class="nav-item" data-section="insights" onclick="showSection('insights')"><span class="nav-icon">🔍</span><span>Insights</span></div>
    <div class="nav-item" data-section="patrimonio" onclick="showSection('patrimonio')"><span class="nav-icon">🏦</span><span>Patrimônio</span></div>
    <div class="nav-item" data-section="categorias" onclick="showSection('categorias')"><span class="nav-icon">🏷️</span><span>Categorias</span></div>
    <div class="nav-item" data-section="anual" onclick="showSection('anual')"><span class="nav-icon">📅</span><span>Visão Anual</span></div>
    <div class="nav-item" data-section="exportar" onclick="showSection('exportar')"><span class="nav-icon">📤</span><span>Exportar Excel</span></div>
    <div class="nav-item" data-section="gcat" onclick="showSection('gcat')"><span class="nav-icon">⚙️</span><span>Ger. Categorias</span></div>
    <div class="nav-item" data-section="cloud" onclick="showSection('cloud')"><span class="nav-icon">☁️</span><span>Nuvem</span></div>
    <div class="nav-spacer"></div>
    <div class="nav-item" data-section="reset" onclick="showSection('reset')" style="color:var(--red);border-top:1px solid var(--border)"><span class="nav-icon">🗑️</span><span>Zerar dados</span></div>
  </nav>
</aside>

<main class="main">

<!-- DASHBOARD -->
<section id="sec-dashboard" class="section active">
  <div class="page-header">
    <div><div class="page-title">Dashboard</div><div class="page-sub" id="dash-period-label">—</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="dash-month" onchange="updateDashboard()"></select>
      <select class="period-select" id="dash-year" onchange="updateDashboard()"></select>
      <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
    </div>
  </div>
  <div class="cards-grid">
    <div class="card"><div class="card-label">Receitas</div><div class="card-value" id="c-income" style="color:var(--green)">R$ 0</div><div class="card-sub" id="c-income-count">—</div></div>
    <div class="card"><div class="card-label">Despesas</div><div class="card-value" id="c-expense" style="color:var(--red)">R$ 0</div><div class="card-sub" id="c-expense-count">—</div></div>
    <div class="card"><div class="card-label">Saldo do mês</div><div class="card-value" id="c-balance">R$ 0</div><div class="card-sub" id="c-balance-pct">—</div></div>
    <div class="card"><div class="card-label">Maior gasto</div><div class="card-value" id="c-topcat" style="font-size:14px;padding-top:4px;color:var(--amber)">—</div><div class="card-sub" id="c-topcat-amt">—</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header"><div class="chart-title">⚖️ Balanço do casal — quem pagou os gastos compartilhados</div></div>
    <div id="balance-card-body" style="font-size:13px;color:var(--text2)">—</div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header">
        <div class="chart-title">Receitas vs Despesas — últimos 6 meses</div>
        <div class="legend"><span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span><span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span></div>
      </div>
      <div style="position:relative;height:220px"><canvas id="chart-monthly"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Gastos por categoria</div></div>
      <div style="position:relative;height:170px"><canvas id="chart-pie"></canvas></div>
      <div id="pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
  <div class="transactions-card" style="margin-top:18px">
    <div class="tx-header"><div class="tx-title">Últimas transações</div><button class="btn" onclick="showSection('transacoes')">Ver todas →</button></div>
    <div class="tx-list" id="recent-list"></div>
  </div>
</section>

<!-- TRANSAÇÕES -->
<section id="sec-transacoes" class="section">
  <div class="page-header">
    <div><div class="page-title">Transações</div><div class="page-sub" id="tx-count-label">—</div></div>
    <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
  </div>

  <!-- Filtros visuais -->
  <div class="filter-section">
    <div class="filter-section-title">Período</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px">
      <select class="period-select" id="tx-month" onchange="renderTxList()"></select>
      <select class="period-select" id="tx-year" onchange="renderTxList()"></select>
    </div>
    <div class="filter-section-title">Tipo</div>
    <div class="chip-group" style="margin-bottom:14px">
      <div class="filter-chip active" id="chip-type-all" onclick="setTxTypeFilter('all')">Todos</div>
      <div class="filter-chip" id="chip-type-income" onclick="setTxTypeFilter('income')">⬆ Receitas</div>
      <div class="filter-chip" id="chip-type-expense" onclick="setTxTypeFilter('expense')">⬇ Despesas</div>
    </div>
    <div class="filter-section-title">Categoria <span style="color:var(--text3);font-weight:400;font-size:11px">(clique para filtrar, clique novamente para remover)</span></div>
    <div class="chip-group" id="cat-filter-chips"></div>
  </div>

  <div class="transactions-card">
    <div class="tx-header">
      <div class="tx-title" id="tx-list-title">Todas as transações</div>
      <button class="btn btn-danger" id="btn-delete-filtered" onclick="deleteFiltered()" style="display:none;font-size:12px">🗑 Excluir selecionadas</button>
    </div>
    <div class="tx-list" id="tx-list" style="max-height:none"></div>
  </div>
</section>

<!-- EVOLUÇÃO MENSAL -->
<section id="sec-evolucao" class="section">
  <div class="page-header">
    <div><div class="page-title">Evolução Mensal</div><div class="page-sub">Acompanhe receitas e despesas mês a mês</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="ev-year" onchange="renderEvolucao()"></select>
    </div>
  </div>
  <div class="cards-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="card"><div class="card-label">Melhor mês (receita)</div><div class="card-value" id="ev-best-income" style="color:var(--green);font-size:16px">—</div><div class="card-sub" id="ev-best-income-val">—</div></div>
    <div class="card"><div class="card-label">Pior mês (despesa)</div><div class="card-value" id="ev-worst-expense" style="color:var(--red);font-size:16px">—</div><div class="card-sub" id="ev-worst-expense-val">—</div></div>
    <div class="card"><div class="card-label">Saldo acumulado</div><div class="card-value" id="ev-accumulated">R$ 0</div><div class="card-sub" id="ev-accumulated-sub">—</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Receitas, Despesas e Saldo mensais</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Saldo</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--amber);border-radius:50%"></span>Saldo acum.</span>
      </div>
    </div>
    <div style="position:relative;height:300px"><canvas id="chart-evolucao"></canvas></div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Saldo acumulado no ano</div></div>
      <div style="position:relative;height:220px"><canvas id="chart-acumulado"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Médias mensais</div></div>
      <div id="ev-medias" style="padding:4px 0"></div>
    </div>
  </div>
</section>

<!-- COMPARATIVO ANO ANTERIOR -->
<section id="sec-comparativo" class="section">
  <div class="page-header">
    <div><div class="page-title">Comparativo</div><div class="page-sub" id="cmp-subtitle">Período atual vs ano anterior</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="cmp-mode" onchange="renderComparativo()">
        <option value="ytd">Acumulado no ano (Jan–mês atual)</option>
        <option value="month">Mês específico</option>
        <option value="full">Ano completo</option>
      </select>
      <select class="period-select" id="cmp-month" onchange="renderComparativo()" style="display:none"></select>
      <select class="period-select" id="cmp-year" onchange="renderComparativo()"></select>
    </div>
  </div>

  <!-- Cards resumo -->
  <div class="cmp-grid">
    <div class="cmp-card">
      <div class="cmp-label">Receitas</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label">2025</span><span class="cmp-value" id="cmp-inc-cur" style="color:var(--green)">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label">2024</span><span class="cmp-value" id="cmp-inc-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-inc-badge">—</span>
      </div>
    </div>
    <div class="cmp-card">
      <div class="cmp-label">Despesas</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label2">2025</span><span class="cmp-value" id="cmp-exp-cur" style="color:var(--red)">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label2">2024</span><span class="cmp-value" id="cmp-exp-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-exp-badge">—</span>
      </div>
    </div>
    <div class="cmp-card">
      <div class="cmp-label">Saldo</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label3">2025</span><span class="cmp-value" id="cmp-bal-cur">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label3">2024</span><span class="cmp-value" id="cmp-bal-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-bal-badge">—</span>
      </div>
    </div>
  </div>

  <!-- Gráfico comparativo mês a mês -->
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title" id="cmp-chart-title">Despesas mês a mês — ano atual vs anterior</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span><span id="cmp-legend-cur">2025</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:rgba(245,101,101,0.3)"></span><span id="cmp-legend-prev">2024</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span><span id="cmp-legend-inc-cur">Receita 2025</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:rgba(34,211,160,0.3)"></span><span id="cmp-legend-inc-prev">Receita 2024</span></span>
      </div>
    </div>
    <div style="position:relative;height:300px"><canvas id="chart-comparativo"></canvas></div>
  </div>

  <!-- Comparativo por categoria -->
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Despesas por categoria — comparativo</div></div>
      <div id="cmp-cat-list" class="cat-list"></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Maiores variações</div></div>
      <div id="cmp-variations"></div>
    </div>
  </div>
</section>

<!-- INSIGHTS -->
<section id="sec-insights" class="section">
  <div class="page-header">
    <div><div class="page-title">Insights</div><div class="page-sub">Análise detalhada do seu histórico financeiro</div></div>
    <select class="period-select" id="ins-year" onchange="renderInsights()"></select>
  </div>
  <div class="insights-grid">
    <div class="insight-card"><div class="insight-label">Maior receita no ano</div><div class="insight-value" id="ins-top-income" style="color:var(--green)">—</div><div class="insight-sub" id="ins-top-income-sub">—</div></div>
    <div class="insight-card"><div class="insight-label">Maior despesa no ano</div><div class="insight-value" id="ins-top-expense" style="color:var(--red)">—</div><div class="insight-sub" id="ins-top-expense-sub">—</div></div>
    <div class="insight-card"><div class="insight-label">Taxa média de poupança</div><div class="insight-value" id="ins-save-rate">—%</div><div class="insight-sub" id="ins-save-sub">—</div></div>
  </div>
  <div class="content-grid" style="margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Ranking de meses — Despesas</div></div>
      <div id="ins-rank-expense"></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Ranking de meses — Receitas</div></div>
      <div id="ins-rank-income"></div>
    </div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">% de gasto por categoria</div></div>
      <div style="position:relative;height:260px"><canvas id="chart-ins-cat"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Detalhamento por categoria</div></div>
      <div id="ins-cat-detail"></div>
    </div>
  </div>
</section>

<!-- CATEGORIAS -->
<section id="sec-categorias" class="section">
  <div class="page-header">
    <div><div class="page-title">Categorias</div><div class="page-sub">Análise por categoria</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="cat-month" onchange="renderCategorias()">
        <option value="-1">Todos os meses</option>
      </select>
      <select class="period-select" id="cat-year" onchange="renderCategorias()"></select>
    </div>
  </div>
  <div class="content-grid">
    <div>
      <div class="chart-card" style="margin-bottom:18px"><div class="chart-header"><div class="chart-title">Despesas por categoria</div></div><div class="cat-list" id="cat-expense-bars"></div></div>
      <div class="chart-card"><div class="chart-header"><div class="chart-title">Receitas por categoria</div></div><div class="cat-list" id="cat-income-bars"></div></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Distribuição de despesas</div></div>
      <div style="position:relative;height:260px"><canvas id="chart-cat-pie"></canvas></div>
      <div id="cat-pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
</section>

<!-- VISÃO ANUAL -->
<section id="sec-anual" class="section">
  <div class="page-header">
    <div><div class="page-title">Visão Anual</div><div class="page-sub">Resumo e análise por categoria</div></div>
    <select class="period-select" id="anual-year" onchange="renderAnual()"></select>
  </div>
  <div class="cards-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="card"><div class="card-label">Total Receitas</div><div class="card-value" id="anual-income" style="color:var(--green)">R$ 0</div></div>
    <div class="card"><div class="card-label">Total Despesas</div><div class="card-value" id="anual-expense" style="color:var(--red)">R$ 0</div></div>
    <div class="card"><div class="card-label">Saldo Anual</div><div class="card-value" id="anual-balance">R$ 0</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div><div class="chart-title">Mensal — Receitas vs Despesas</div><div style="font-size:10.5px;color:var(--text3);margin-top:2px">Receitas: tons vivos · Despesas: tons escuros e borda reforçada</div></div>
      <div class="legend"><span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span><span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span><span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Saldo</span></div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-anual"></canvas></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div><div class="chart-title">Evolução anual por categoria</div><div style="font-size:10.5px;color:var(--text3);margin-top:2px">Clique em uma barra mensal para ver os lançamentos</div></div>
      <div class="hdr-actions" style="gap:8px">
        <select class="period-select" id="cat-anual-type" onchange="populateCatAnualSelect();renderCatAnual()" style="font-size:12px"><option value="expense">Despesas</option><option value="income">Receitas</option></select>
        <select class="period-select" id="cat-anual-sel" onchange="renderCatAnual()" style="font-size:12px"></select>
      </div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-cat-anual"></canvas></div>
    <div id="cat-anual-summary" style="display:flex;gap:12px;flex-wrap:wrap;margin-top:14px;font-size:12px"></div>
  </div>
  <div class="transactions-card">
    <div class="tx-header"><div class="tx-title">Resumo mensal</div></div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="padding:11px 18px;text-align:left;color:var(--text2);font-weight:500">Mês</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Receitas</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Despesas</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Saldo</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Qtd.</th>
        </tr></thead>
        <tbody id="anual-tbody"></tbody>
      </table>
    </div>
  </div>
</section>

<!-- EXPORTAR -->
<section id="sec-exportar" class="section">
  <div class="page-header"><div><div class="page-title">Exportar Excel</div><div class="page-sub">Exportação completa de toda a base Mariana</div></div></div>
  <div class="chart-card" style="margin-bottom:18px">
    <div style="font-size:14px;font-weight:600;margin-bottom:8px">Planilha completa</div>
    <div style="font-size:13px;color:var(--text2);line-height:1.6;margin-bottom:18px">Exporta todas as transações registradas, sem filtros de mês, ano, tipo ou categoria. O arquivo também inclui resumo anual e categorias.</div>
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <button class="btn btn-primary" onclick="exportExcel()" style="font-size:13px;padding:9px 20px">📥 Exportar planilha completa</button>
      <span id="exp-preview-count" style="font-size:12px;color:var(--text2)">—</span>
    </div>
  </div>
  <div class="chart-card"><div style="font-size:14px;font-weight:500;margin-bottom:14px">Prévia de toda a base</div><div class="import-table-wrap" style="max-height:380px"><table class="import-table"><thead><tr><th>Data</th><th>Tipo</th><th>Usuário</th><th>Categoria</th><th>Valor</th><th>Descrição</th></tr></thead><tbody id="exp-preview-tbody"></tbody></table></div></div>
</section>

<!-- PATRIMÔNIO -->
<section id="sec-patrimonio" class="section">
  <div class="page-header">
    <div>
      <div class="page-title">Evolução Patrimonial</div>
      <div class="page-sub">Projeção baseada no seu histórico + SELIC + IPCA projetados pelo BCB</div>
    </div>
    <button class="btn btn-primary" onclick="calcularPatrimonio()">▶ Calcular</button>
  </div>

  <!-- Parâmetros -->
  <div class="chart-card" style="margin-bottom:18px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px">
      <div style="font-size:14px;font-weight:500">Parâmetros da projeção</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <span id="pat-selic-badge" class="pat-api-badge pat-api-load">⏳ Buscando SELIC...</span>
        <span id="pat-ipca-badge" class="pat-api-badge pat-api-load">⏳ Buscando IPCA...</span>
        <button class="btn" id="pat-refresh-rates" type="button" onclick="atualizarTaxasBCB(true)" style="font-size:11px;padding:4px 10px">↻ Atualizar taxas</button>
      </div>
    </div>

    <div class="pat-input-grid">
      <div class="pat-field">
        <label>💰 Saldo atual aplicado (R$)</label>
        <input class="pat-input" id="pat-saldo" type="number" step="0.01" placeholder="Ex: 50000">
        <span class="pat-input-hint">Valor total que você tem aplicado hoje</span>
      </div>
      <div class="pat-field">
        <label>📅 Horizonte de projeção (anos)</label>
        <input class="pat-input" id="pat-anos" type="number" min="1" max="10" value="10" placeholder="Ex: 10" onchange="normalizarHorizontePatrimonio();renderSelicFocusTable()">
        <span class="pat-input-hint">Horizonte permitido: de 1 a 10 anos</span>
      </div>
      <div class="pat-field">
        <label>📈 SELIC de contingência (%)</label>
        <input class="pat-input" id="pat-selic" type="number" step="0.01" placeholder="Buscando...">
        <span class="pat-input-hint" id="pat-selic-hint">Usada apenas quando o Focus não trouxer projeção para determinado ano.</span>
      </div>
      <div class="pat-field">
        <label>📊 IPCA anual projetado (%)</label>
        <input class="pat-input" id="pat-ipca" type="number" step="0.01" placeholder="Buscando...">
        <span class="pat-input-hint" id="pat-ipca-hint">Buscando Boletim Focus...</span>
      </div>
      <div class="pat-field">
        <label>⬆ Crescimento anual de receita (%)</label>
        <input class="pat-input" id="pat-cresc-rec" type="number" step="0.1" value="3" placeholder="Ex: 3">
        <span class="pat-input-hint">Ex: aumento de salário esperado</span>
      </div>
      <div class="pat-field">
        <label>⬇ Crescimento anual de despesa (%)</label>
        <input class="pat-input" id="pat-cresc-desp" type="number" step="0.1" value="4" placeholder="Ex: 4">
        <span class="pat-input-hint">Ex: inflação sobre seus gastos</span>
      </div>
      <div class="pat-field">
        <label>➕ Aporte extra mensal (R$)</label>
        <input class="pat-input" id="pat-aporte" type="number" step="0.01" value="0" placeholder="0">
        <span class="pat-input-hint">Depósito adicional todo mês</span>
      </div>
    </div>

    <div id="pat-selic-projection-card" style="background:var(--bg3);border-radius:var(--radius-sm);padding:14px 16px;margin-top:4px;margin-bottom:14px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;margin-bottom:10px">
        <div>
          <div style="font-size:12px;color:var(--text2);font-weight:500">Projeção anual da SELIC — Boletim Focus/BCB</div>
          <div id="pat-selic-projection-meta" style="font-size:10.5px;color:var(--text3);margin-top:2px">Buscando projeções para os próximos anos...</div>
        </div>
      </div>
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead>
            <tr style="border-bottom:1px solid var(--border)">
              <th style="padding:8px 10px;text-align:left;color:var(--text2);font-weight:500">Ano</th>
              <th style="padding:8px 10px;text-align:right;color:var(--text2);font-weight:500">Mediana Focus</th>
              <th style="padding:8px 10px;text-align:left;color:var(--text2);font-weight:500">Aplicação no cálculo</th>
            </tr>
          </thead>
          <tbody id="pat-selic-projection-tbody">
            <tr><td colspan="3" style="padding:12px 10px;color:var(--text3)">Consultando Banco Central...</td></tr>
          </tbody>
        </table>
      </div>
      <div style="font-size:10.5px;color:var(--text3);margin-top:8px">
        Para anos sem projeção publicada, o sistema mantém a última mediana anual disponível até completar o horizonte máximo de 10 anos. Se o Focus não responder, usa a taxa de contingência.
      </div>
    </div>

    <!-- Receita e despesa base calculadas -->
    <div style="background:var(--bg3);border-radius:var(--radius-sm);padding:14px 16px;margin-top:4px">
      <div style="font-size:12px;color:var(--text2);margin-bottom:10px;font-weight:500">Base calculada do histórico (últimos 12 meses)</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px">Receita operacional média/mês</div>
          <div style="font-size:16px;font-weight:600;font-family:DM Mono;color:var(--green)" id="pat-rec-base">—</div>
          <div style="font-size:10.5px;color:var(--text3);margin-top:3px">Categorias incluídas: <span id="pat-rec-cats" style="color:var(--text2)">—</span></div>
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px">Despesa média/mês</div>
          <div style="font-size:16px;font-weight:600;font-family:DM Mono;color:var(--red)" id="pat-desp-base">—</div>
          <div style="font-size:10.5px;color:var(--text3);margin-top:3px">Todas as categorias de despesa</div>
        </div>
      </div>
      <div style="margin-top:10px;padding-top:10px;border-top:1px solid var(--border)">
        <div style="font-size:10.5px;color:var(--text3);margin-bottom:4px">Categorias excluídas da projeção de receita:</div>
        <div class="pat-excluded-cats" id="pat-excluded-cats"></div>
      </div>
    </div>
  </div>

  <!-- Cards resultado -->
  <div class="cards-grid" id="pat-cards" style="display:none">
    <div class="card"><div class="card-label">Patrimônio em 1 ano</div><div class="card-value" id="pat-1y" style="color:var(--green)">—</div><div class="card-sub" id="pat-1y-real">—</div></div>
    <div class="card"><div class="card-label" id="pat-mid-label">Patrimônio em 5 anos</div><div class="card-value" id="pat-mid" style="color:var(--green)">—</div><div class="card-sub" id="pat-mid-real">—</div></div>
    <div class="card"><div class="card-label" id="pat-long-label">Patrimônio em 10 anos</div><div class="card-value" id="pat-long" style="color:var(--green)">—</div><div class="card-sub" id="pat-long-real">—</div></div>
    <div class="card"><div class="card-label">Rendimento total acumulado</div><div class="card-value" id="pat-rend-total" style="color:var(--amber)">—</div><div class="card-sub" id="pat-rend-pct">—</div></div>
  </div>

  <!-- Marcos -->
  <div id="pat-milestones" style="display:none;margin-bottom:18px">
    <div style="font-size:13px;font-weight:500;margin-bottom:10px">🏆 Marcos projetados — R$ 1M, R$ 2M e R$ 5M</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap" id="pat-milestone-list"></div>
  </div>

  <!-- Gráfico -->
  <div class="chart-card" style="margin-bottom:18px" id="pat-chart-card" style="display:none">
    <div class="chart-header">
      <div class="chart-title">Evolução patrimonial projetada</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Patrimônio nominal</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Valor real (descontado IPCA)</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--amber)"></span>Rendimento acumulado</span>
      </div>
    </div>
    <div style="position:relative;height:320px"><canvas id="chart-patrimonio"></canvas></div>
  </div>

  <!-- Tabela anual -->
  <div class="transactions-card" id="pat-table-card" style="display:none">
    <div class="tx-header">
      <div class="tx-title">Projeção ano a ano</div>
      <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer">
        <input type="checkbox" id="pat-show-monthly" onchange="toggleMonthlyDetail()" style="accent-color:var(--green)">
        Ver detalhamento mensal
      </label>
    </div>
    <div class="pat-table-wrap">
      <table class="pat-table">
        <thead>
          <tr>
            <th style="text-align:left">Período</th>
            <th>Receita</th>
            <th>Despesa</th>
            <th>Fluxo</th>
            <th>SELIC usada</th>
            <th>Rendimento</th>
            <th>Patrimônio Nominal</th>
            <th>Valor Real</th>
          </tr>
        </thead>
        <tbody id="pat-tbody"></tbody>
      </table>
    </div>
  </div>
</section>

<!-- GER. CATEGORIAS -->
<section id="sec-gcat" class="section">
  <div class="page-header"><div><div class="page-title">Gerenciar Categorias</div><div class="page-sub">Crie e edite categorias personalizadas</div></div></div>
  <div class="cat-manager">
    <div class="cat-section-card"><div class="cat-section-title">⬇ Despesas</div><div id="gcat-expense-list"></div><div class="new-cat-form"><input class="form-input" id="new-exp-name" placeholder="Nome" style="flex:1;min-width:100px"><input type="color" class="color-picker" id="new-exp-color" value="#f56565"><input class="form-input" id="new-exp-icon" placeholder="🏷️" style="width:46px;text-align:center"><button class="btn btn-primary" onclick="addCat('expense')">+</button></div></div>
    <div class="cat-section-card"><div class="cat-section-title">⬆ Receitas</div><div id="gcat-income-list"></div><div class="new-cat-form"><input class="form-input" id="new-inc-name" placeholder="Nome" style="flex:1;min-width:100px"><input type="color" class="color-picker" id="new-inc-color" value="#22d3a0"><input class="form-input" id="new-inc-icon" placeholder="💰" style="width:46px;text-align:center"><button class="btn btn-primary" onclick="addCat('income')">+</button></div></div>
  </div>
</section>

<!-- NUVEM -->
<section id="sec-cloud" class="section">
  <div class="page-header">
    <div><div class="page-title">Sincronização na Nuvem</div><div class="page-sub">Firebase protegido por autenticação</div></div>
    <div id="sync-status-badge" class="sync-status sync-local"><span class="sync-dot local"></span>Aguardando autenticação</div>
  </div>
  <div class="cloud-card">
    <div style="font-size:15px;font-weight:500;margin-bottom:10px">Conta conectada</div>
    <div style="display:grid;gap:8px;font-size:13px;color:var(--text2)">
      <div><strong style="color:var(--text)">Mariana:</strong> <span id="cloud-user-email">—</span></div>
      <div><strong style="color:var(--text)">Perfil:</strong> <span id="cloud-user-profile">—</span></div>
      <div><strong style="color:var(--text)">Banco:</strong> <span id="cloud-url-display">—</span></div>
    </div>
  </div>
  <div class="chart-card" id="cloud-actions">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:16px">
      <div><div style="font-size:14px;font-weight:500">Sincronização segura</div><div style="font-size:12px;color:var(--text2);margin-top:2px">Os dados são baixados antes de liberar o sistema e salvos automaticamente após alterações.</div></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn" onclick="syncFromCloud(true)">⬇ Baixar novamente</button>
        <button class="btn btn-primary" onclick="syncToCloud(true)">⬆ Enviar tudo</button>
        <button class="btn btn-danger" onclick="logoutProfile()">Sair da conta</button>
      </div>
    </div>
    <div id="cloud-log" style="font-size:12px;color:var(--text2);background:var(--bg3);border-radius:8px;padding:12px;min-height:48px;line-height:1.7"></div>
  </div>
</section>

<!-- ZERAR DADOS -->
<section id="sec-reset" class="section">
  <div class="page-header"><div><div class="page-title">Zerar Dados</div><div class="page-sub">Apague definitivamente os dados locais e da nuvem</div></div></div>
  <div style="display:grid;gap:16px;max-width:600px">
    <div class="chart-card" style="border-color:rgba(245,101,101,0.2)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">📋</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px">Apagar todas as transações</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove todos os lançamentos. As categorias são mantidas.</div>
          <div style="background:var(--bg3);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text2);margin-bottom:14px" id="reset-tx-count">—</div>
          <button class="btn btn-danger" onclick="resetTransactions()">🗑 Apagar transações</button>
        </div>
      </div>
    </div>
    <div class="chart-card" style="border-color:rgba(251,191,36,0.2)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">🏷️</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px">Remover categorias personalizadas</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove apenas as categorias personalizadas. As padrão são mantidas.</div>
          <div style="background:var(--bg3);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text2);margin-bottom:14px" id="reset-cats-count">—</div>
          <button class="btn" onclick="resetCustomCats()" style="background:rgba(251,191,36,0.1);border-color:rgba(251,191,36,0.3);color:var(--amber)">🏷 Remover criadas</button>
        </div>
      </div>
    </div>
    <div class="chart-card" style="border:2px solid rgba(245,101,101,0.4);background:rgba(245,101,101,0.04)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">⚠️</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px;color:var(--red)">Apagar tudo</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove <strong style="color:var(--text)">tudo</strong>. Categorias padrão são restauradas. <strong style="color:var(--red)">Irreversível.</strong></div>
          <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer;margin-bottom:14px"><input type="checkbox" id="confirm-reset-all" style="accent-color:var(--red)"> Entendo que esta ação não pode ser desfeita</label>
          <button class="btn btn-danger" onclick="resetAll()">⚠️ Apagar tudo</button>
        </div>
      </div>
    </div>
  </div>
</section>

</main>
</div><!-- end app-root -->

<!-- Mobile bar -->
<div class="mobile-bar">
  <div class="mobile-bar-inner">
    <button class="mob-btn active" id="mob-dashboard" onclick="showSection('dashboard')"><span class="mob-icon">📊</span>Início</button>
    <button class="mob-btn" id="mob-transacoes" onclick="showSection('transacoes')"><span class="mob-icon">↕️</span>Trans.</button>
    <button class="mob-btn" id="mob-add" onclick="openModal()"><span class="mob-icon">➕</span>Novo</button>
    <button class="mob-btn" id="mob-evolucao" onclick="showSection('evolucao')"><span class="mob-icon">📈</span>Evolução</button>
    <button class="mob-btn" id="mob-menu" onclick="toggleMobileMenu()"><span class="mob-icon">☰</span>Menu</button>
  </div>
</div>

<!-- Modal -->
<div class="modal-bg" id="modal-bg" onclick="closeModalBg(event)">
  <div class="modal">
    <div class="modal-title" id="modal-title">Nova transação</div>
    <div class="form-row">
      <label class="form-label">Tipo</label>
      <div class="radio-group">
        <div class="radio-btn sel-income" id="rb-income" onclick="setType('income')">⬆ Receita</div>
        <div class="radio-btn" id="rb-expense" onclick="setType('expense')">⬇ Despesa</div>
      </div>
    </div>
    <div class="form-row"><label class="form-label">Descrição</label><input class="form-input" id="f-desc" placeholder="Ex: Salário, Aluguel, Mercado..."></div>
    <div class="form-row-2">
      <div class="form-row" style="margin-bottom:0"><label class="form-label">Valor (R$)</label><input class="form-input" id="f-amount" type="number" step="0.01" placeholder="0,00"><div id="f-amount-hint" style="font-size:11px;color:var(--text3);margin-top:4px;display:none">Valor negativo: usado para abater de "a receber" (débito sobre receita)</div></div>
      <div class="form-row" style="margin-bottom:0"><label class="form-label">Data</label><input class="form-input" id="f-date" type="date"></div>
    </div>
    <div class="form-row" style="margin-top:13px">
      <label class="form-label">Categoria</label>
      <input class="form-input" id="f-cat-search" placeholder="Buscar categoria..." oninput="filterCatOptions()" onclick="showCatDropdown()" autocomplete="off" style="margin-bottom:4px">
      <div id="f-cat-dropdown" style="display:none;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);max-height:200px;overflow-y:auto;position:relative;z-index:10"></div>
      <input type="hidden" id="f-cat">
    </div>
    <!-- Compartilhado -->
    <div class="form-row" style="display:none;margin-top:4px" id="f-shared-row">
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;color:var(--text2)">
        <input type="checkbox" id="f-shared" onchange="toggleShared()" style="accent-color:var(--green)">
        🤝 Compartilhado com <span id="f-shared-partner-name"></span>
      </label>
    </div>
    <div id="f-shared-split-row" style="display:none;background:var(--bg3);border-radius:var(--radius-sm);padding:12px 14px;margin-bottom:13px">
      <div style="font-size:11.5px;color:var(--text3);margin-bottom:8px">Divisão do valor total (some 100%)</div>
      <div class="form-row-2" style="margin-bottom:0">
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label" id="f-shared-mine-label">Minha parte (%)</label>
          <input class="form-input" id="f-shared-mine-pct" type="number" min="0" max="100" value="50" oninput="syncSharedPct('mine')">
        </div>
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label" id="f-shared-partner-label">Parte do parceiro (%)</label>
          <input class="form-input" id="f-shared-partner-pct" type="number" min="0" max="100" value="50" oninput="syncSharedPct('partner')">
        </div>
      </div>
      <div style="font-size:11px;color:var(--text3);margin-top:8px" id="f-shared-preview">—</div>
    </div>
    <!-- Parcelamento -->
    <div class="form-row" style="margin-top:4px">
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;color:var(--text2)">
        <input type="checkbox" id="f-parceled" onchange="toggleParcelamento()" style="accent-color:var(--green)">
        Pagamento parcelado
      </label>
    </div>
    <div id="f-parcelas-row" style="display:none" class="form-row-2">
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Nº de parcelas</label>
        <input class="form-input" id="f-parcelas" type="number" min="2" max="120" value="2" placeholder="Ex: 12">
      </div>
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Valor de cada parcela</label>
        <input class="form-input" id="f-parcela-valor" type="number" min="0.01" step="0.01" placeholder="Calculado automaticamente" readonly style="opacity:.7">
      </div>
    </div>
    <!-- Recorrência mensal -->
    <div class="form-row" style="margin-top:4px" id="f-recurring-check-row">
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;color:var(--text2)">
        <input type="checkbox" id="f-recurring" onchange="toggleRecorrencia()" style="accent-color:#f59e0b">
        🔁 Gasto recorrente mensal
      </label>
    </div>
    <div id="f-recurring-row" style="display:none;background:var(--bg3);border-radius:var(--radius-sm);padding:12px 14px;margin-bottom:13px">
      <div class="form-row-2">
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label">Quantidade de meses</label>
          <input class="form-input" id="f-recurring-months" type="number" min="2" max="120" value="12">
        </div>
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label">Valor mensal</label>
          <input class="form-input" id="f-recurring-value" type="text" readonly style="opacity:.75">
        </div>
      </div>
      <div style="font-size:11px;color:var(--text3);margin-top:8px">
        O valor informado será lançado integralmente em cada mês. Não é parcelamento.
      </div>
    </div>
    <input type="hidden" id="f-editing-id">
    <div class="modal-actions"><button class="btn" onclick="closeModal()">Cancelar</button><button class="btn btn-primary" onclick="saveTransaction()">Salvar</button></div>
  </div>
</div>

<div class="modal-bg" id="history-modal-bg" onclick="if(event.target.id==='history-modal-bg')closeCategoryHistory()">
  <div class="modal" style="width:680px">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:16px">
      <div>
        <div class="modal-title" id="history-modal-title" style="margin-bottom:3px">Histórico da categoria</div>
        <div style="font-size:12px;color:var(--text2)" id="history-modal-sub">—</div>
      </div>
      <button class="btn" onclick="closeCategoryHistory()">✕ Fechar</button>
    </div>
    <div id="history-modal-summary" style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px"></div>
    <div id="history-modal-list" class="history-list"></div>
  </div>
</div>

<div class="modal-bg" id="chart-summary-modal-bg" onclick="if(event.target.id==='chart-summary-modal-bg')closeChartSummary()">
  <div class="modal" style="width:620px">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:16px">
      <div>
        <div class="modal-title" id="chart-summary-title" style="margin-bottom:3px">Resumo do gráfico</div>
        <div style="font-size:12px;color:var(--text2)" id="chart-summary-sub">—</div>
      </div>
      <button class="btn" onclick="closeChartSummary()">✕ Fechar</button>
    </div>
    <div id="chart-summary-body"></div>
  </div>
</div>

<script>

function chartTitleFromCanvas(chart){
  const card=chart?.canvas?.closest('.chart-card');
  return card?.querySelector('.chart-title')?.textContent?.trim()||'Resumo do gráfico';
}
function monthIndexFromLabel(label){
  const short=MONTHS_SHORT.findIndex(m=>m===String(label));
  if(short>=0)return short;
  return MONTHS.findIndex(m=>m===String(label));
}
function transactionsForChartPoint(chart,elements){
  const canvasId=chart.canvas?.id||'';
  const point=elements[0];
  const index=point.index;
  const label=chart.data.labels?.[index];
  let txs=getVisibleTx();
  let title=chartTitleFromCanvas(chart);
  let subtitle=String(label??'');
  let type=null,category=null,month=null,year=null;

  if(['chart-monthly','chart-evolucao','chart-acumulado'].includes(canvasId)){
    month=monthIndexFromLabel(label);
    const selectedYear=parseInt(document.getElementById(canvasId==='chart-monthly'?'dash-year':'ev-year')?.value)||new Date().getFullYear();
    year=selectedYear;
    const dsLabel=String(chart.data.datasets?.[point.datasetIndex]?.label||'').toLowerCase();
    if(dsLabel.includes('receita'))type='income';
    if(dsLabel.includes('despesa'))type='expense';
  }else if(canvasId==='chart-anual'){
    month=index;
    year=parseInt(document.getElementById('anual-year')?.value)||new Date().getFullYear();
    const dsLabel=String(chart.data.datasets?.[point.datasetIndex]?.label||'').toLowerCase();
    if(dsLabel.includes('receita'))type='income';
    if(dsLabel.includes('despesa'))type='expense';
    if(dsLabel.includes('mariana'))txs=txs.filter(t=>t.owner==='mariana');
    if(dsLabel.includes('mariana'))txs=txs.filter(t=>t.owner==='mariana');
  }else if(canvasId==='chart-pie'){
    month=parseInt(document.getElementById('dash-month')?.value);
    year=parseInt(document.getElementById('dash-year')?.value);
    type='expense';
    const catName=String(label);
    category=allCats().find(c=>c.name===catName)?.id;
  }else if(canvasId==='chart-cat-pie'){
    month=parseInt(document.getElementById('cat-month')?.value);
    year=parseInt(document.getElementById('cat-year')?.value);
    type='expense';
    category=allCats().find(c=>c.name===String(label))?.id;
  }else if(canvasId==='chart-ins-cat'){
    year=parseInt(document.getElementById('ins-year')?.value);
    type='expense';
    category=allCats().find(c=>c.name===String(label))?.id;
  }else if(canvasId==='chart-cat-anual'){
    month=index;
    year=parseInt(document.getElementById('anual-year')?.value);
    type=document.getElementById('cat-anual-type')?.value||'expense';
    category=document.getElementById('cat-anual-sel')?.value;
    const dsLabel=String(chart.data.datasets?.[point.datasetIndex]?.label||'').toLowerCase();
    if(dsLabel.includes('mariana'))txs=txs.filter(t=>t.owner==='mariana');
    if(dsLabel.includes('mariana'))txs=txs.filter(t=>t.owner==='mariana');
  }else if(canvasId==='chart-comparativo'){
    month=index;
    year=parseInt(document.getElementById('cmp-year')?.value)||new Date().getFullYear();
    const dsLabel=String(chart.data.datasets?.[point.datasetIndex]?.label||'').toLowerCase();
    if(dsLabel.includes(String(year-1)))year-=1;
    if(dsLabel.includes('receita'))type='income';
    if(dsLabel.includes('despesa'))type='expense';
  }

  if(Number.isFinite(year))txs=txs.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);
  if(Number.isFinite(month)&&month>=0)txs=txs.filter(t=>new Date(t.date+'T12:00').getMonth()===month);
  if(type)txs=txs.filter(t=>t.type===type);
  if(category)txs=txs.filter(t=>t.category===category);

  return{txs:txs.sort((a,b)=>new Date(b.date)-new Date(a.date)),title,subtitle};
}
function openChartSummary(chart,elements){
  if(!chart||!elements?.length)return;
  const result=transactionsForChartPoint(chart,elements);
  const txs=result.txs;
  const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);

  document.getElementById('chart-summary-title').textContent=result.title;
  document.getElementById('chart-summary-sub').textContent=result.subtitle;
  document.getElementById('chart-summary-body').innerHTML=`
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">
      <span style="background:var(--green-bg);color:var(--green);padding:6px 10px;border-radius:8px">Receitas: <strong>${fmt(income)}</strong></span>
      <span style="background:var(--red-bg);color:var(--red);padding:6px 10px;border-radius:8px">Despesas: <strong>${fmt(expense)}</strong></span>
      <span style="background:var(--bg3);padding:6px 10px;border-radius:8px">${txs.length} lançamento(s)</span>
    </div>
    <div class="history-list">
      ${txs.length?txs.map(t=>{
        const d=new Date(t.date+'T12:00');
        const cat=getCat(t.category);
        return`<div class="history-row">
          <div style="font-size:11px;color:var(--text2)">${d.toLocaleDateString('pt-BR')}</div>
          <div>
            <div style="font-size:13px;font-weight:500">${t.description}</div>
            <div style="font-size:10.5px;color:var(--text3)">${cat.icon} ${cat.name} · ${profileName(t.owner)}${t.shared?' · compartilhado':''}${t.recurring?' · recorrente':''}</div>
          </div>
          <div style="text-align:right;font-family:'DM Mono',monospace;color:${t.type==='expense'?'var(--red)':'var(--green)'}">${t.type==='expense'?'−':'+'}${fmt(t.amount)}</div>
        </div>`;
      }).join(''):'<div class="empty">Nenhum lançamento relacionado a este item.</div>'}
    </div>`;
  document.getElementById('chart-summary-modal-bg').classList.add('open');
}
function closeChartSummary(){document.getElementById('chart-summary-modal-bg').classList.remove('open')}

const chartClickSummaryPlugin={
  id:'chartClickSummaryPlugin',
  afterEvent(chart,args){
    if(args.event.type!=='click'||chart.canvas?.id==='chart-cat-anual')return;
    const points=chart.getElementsAtEventForMode(args.event.native,'nearest',{intersect:true},true);
    if(points.length)openChartSummary(chart,points);
  }
};
Chart.register(chartClickSummaryPlugin);

// FinTrack MARIANA V4 — Firebase Mariana auditado e isolado
// ===================== FIREBASE AUTHENTICATION — MARIANA ONLY =====================
// PREENCHA os campos abaixo com os dados do aplicativo Web do Firebase.
// Firebase Console > Configurações do projeto > Seus apps > App Web > Configuração do SDK.
// Realtime Database EXCLUSIVO: projeto Firebase da Mariana
const EXPECTED_FIREBASE_PROJECT_ID='planilha-financeira-mariana';
const EXPECTED_DATABASE_URL='https://planilha-financeira-mariana-default-rtdb.asia-southeast1.firebasedatabase.app/';
const MARIANA_AUTH_EMAIL='mcarneiroramos@gmail.com';
const MARIANA_AUTH_UID='bFp5D7cOTWW5q8uYBFsEQoaqeas1';
const EXPECTED_AUTH_DOMAIN='planilha-financeira-mariana.firebaseapp.com';
const EXPECTED_STORAGE_BUCKET='planilha-financeira-mariana.firebasestorage.app';
const EXPECTED_MESSAGING_SENDER_ID='546792812228';
const EXPECTED_APP_ID='1:546792812228:web:e1eaf13d4925dc56262391';
const FIREBASE_APP_NAME='fintrack-mariana-v4';

const FIREBASE_CONFIG={
  // IMPORTANTE: use SOMENTE os 3 valores do App Web criado dentro do projeto
  // planilha-financeira-mariana. Valores do projeto antigo são bloqueados abaixo.
  apiKey:'AIzaSyBf-bBwIzDd488gQGAoI0M0nhJDcpS0sI4',
  authDomain:'planilha-financeira-mariana.firebaseapp.com',
  databaseURL:EXPECTED_DATABASE_URL,
  projectId:EXPECTED_FIREBASE_PROJECT_ID,
  storageBucket:'planilha-financeira-mariana.firebasestorage.app',
  messagingSenderId:'546792812228',
  appId:'1:546792812228:web:e1eaf13d4925dc56262391',
  measurementId:'G-2BQ15FJJKR'
};

let firebaseApp=null,auth=null,realtimeDb=null;
let _authReady=false,_cloudLoaded=false,_appOpening=false,_syncTimer=null,_syncInProgress=false,_syncPending=false,_suppressAutoSync=false,_cloudWriteQueue=Promise.resolve();
let _cloudTxMap={},_localDirtyIds=new Set(),_localDeletedIds=new Set(),_transactionsListener=null;
let _lastRemoteTxFingerprint='',_remoteApplyQueue=Promise.resolve();
let _categoriesDirty=false,_importsDirty=false,_lastCloudError=null,_retryAttempt=0,_retryTimer=null,_storageWriteQueue=Promise.resolve();

function normalizeDatabaseUrl(v){return String(v||'').trim().replace(/\/+$/,'')+'/'}
function firebaseConfigured(){
  const vals=[FIREBASE_CONFIG.apiKey,FIREBASE_CONFIG.authDomain,FIREBASE_CONFIG.databaseURL,FIREBASE_CONFIG.projectId,FIREBASE_CONFIG.storageBucket,FIREBASE_CONFIG.messagingSenderId,FIREBASE_CONFIG.appId,MARIANA_AUTH_EMAIL,MARIANA_AUTH_UID];
  return vals.every(v=>v && !String(v).startsWith('COLE_'));
}
function assertMarianaFirebaseConfig(){
  if(FIREBASE_CONFIG.projectId!==EXPECTED_FIREBASE_PROJECT_ID)throw new Error('Projeto Firebase incorreto. Esperado: '+EXPECTED_FIREBASE_PROJECT_ID);
  if(normalizeDatabaseUrl(FIREBASE_CONFIG.databaseURL)!==EXPECTED_DATABASE_URL)throw new Error('Realtime Database incorreto. Esperado: '+EXPECTED_DATABASE_URL);
  if(FIREBASE_CONFIG.authDomain!==EXPECTED_AUTH_DOMAIN)throw new Error('Auth Domain não pertence ao projeto Mariana.');
  if(FIREBASE_CONFIG.storageBucket!==EXPECTED_STORAGE_BUCKET)throw new Error('Storage Bucket não pertence ao projeto Mariana.');
  if(String(FIREBASE_CONFIG.messagingSenderId)!==EXPECTED_MESSAGING_SENDER_ID)throw new Error('Messaging Sender ID não pertence ao projeto Mariana.');
  if(FIREBASE_CONFIG.appId!==EXPECTED_APP_ID)throw new Error('App ID não pertence ao projeto Mariana.');
}
function normalizeEmail(v){return String(v||'').trim().toLowerCase()}
function profileFromUser(user){
  const email=normalizeEmail(user?.email),uid=String(user?.uid||'');
  return email===MARIANA_AUTH_EMAIL && uid===MARIANA_AUTH_UID ? 'mariana' : null;
}
function assertAuthorizedMariana(user){
  if(!user)throw new Error('Mariana não autenticado.');
  if(normalizeEmail(user.email)!==MARIANA_AUTH_EMAIL)throw new Error('Conta não autorizada para esta base.');
  if(String(user.uid||'')!==MARIANA_AUTH_UID)throw new Error('UID do Firebase não corresponde ao usuário Mariana autorizado.');
  return true;
}

// Base exclusiva do projeto Mariana. O acesso é bloqueado por e-mail + UID autorizado.
window.soloUserRoot=function soloUserRoot(){
  if(!realtimeDb)throw new Error('Firebase Mariana ainda não foi inicializado.');
  const user=auth?.currentUser;assertAuthorizedMariana(user);
  return realtimeDb.ref('fintrack');
};
var soloUserRoot=window.soloUserRoot;

function setLoginError(msg){const el=document.getElementById('lock-err');if(el)el.textContent=msg||''}
function showLogin(){
  document.getElementById('app-root').style.display='none';
  document.getElementById('lock-screen').style.display='flex';
  const email=document.getElementById('login-email');if(email&&!email.value)email.value=MARIANA_AUTH_EMAIL;
  const btn=document.getElementById('login-btn');if(btn){btn.disabled=false;btn.textContent='Entrar'}
}
function showApp(){document.getElementById('lock-screen').style.display='none';document.getElementById('app-root').style.display='flex'}
async function loginWithFirebase(e){
  if(e&&typeof e.preventDefault==='function')e.preventDefault();
  if(!auth){try{initFirebase()}catch(err){setLoginError('Falha ao iniciar o Firebase Mariana: '+(err?.message||err));return;} if(!auth){setLoginError('Firebase Mariana não inicializou. Recarregue o arquivo V4 Mariana.');return;}}
  setLoginError('');
  if(!firebaseConfigured()){setLoginError('Preencha apiKey, messagingSenderId e appId do App Web do projeto Firebase da Mariana.');return;}
  const email=normalizeEmail(document.getElementById('login-email').value);
  if(email!==MARIANA_AUTH_EMAIL){setLoginError('Use a conta Mariana autorizada: '+MARIANA_AUTH_EMAIL);return;}
  const password=document.getElementById('login-password').value;
  const btn=document.getElementById('login-btn');btn.disabled=true;btn.textContent='Entrando...';
  try{
    const cred=await auth.signInWithEmailAndPassword(email,password);
    assertAuthorizedMariana(cred.user);
  }catch(err){try{await auth.signOut()}catch{}setLoginError(firebaseAuthMessage(err));btn.disabled=false;btn.textContent='Entrar'}
}
function firebaseAuthMessage(err){
  const msg=String(err?.message||'');
  if(msg.includes('UID do Firebase'))return'Esta conta existe, mas o UID não corresponde ao Mariana autorizado neste HTML.';
  const code=err?.code||'';
  if(code.includes('invalid-credential')||code.includes('wrong-password')||code.includes('user-not-found'))return'E-mail ou senha incorretos no Firebase Mariana.';
  if(code.includes('too-many-requests'))return'Muitas tentativas. Aguarde alguns minutos.';
  if(code.includes('network-request-failed'))return'Sem conexão com a internet.';
  if(code.includes('unauthorized-domain'))return'Este domínio ainda não foi autorizado no Firebase Mariana.';
  return'Não foi possível entrar: '+(err?.message||code);
}
async function resetFirebasePassword(){
  if(!auth||!firebaseConfigured()){setLoginError('Configure o App Web do projeto Firebase da Mariana.');return}
  try{await auth.sendPasswordResetEmail(MARIANA_AUTH_EMAIL);setLoginError('E-mail de redefinição enviado para a conta Mariana.')}catch(err){setLoginError(firebaseAuthMessage(err))}
}
async function logoutProfile(){
  try{if(realtimeDb&&_transactionsListener)soloUserRoot().child('transactions').off('value',_transactionsListener);_transactionsListener=null;if(auth)await auth.signOut();}catch{}
  _currentProfile=null;_viewOwner='mariana';_cloudLoaded=false;_lastRemoteTxFingerprint='';showLogin();setLoginError('');
}
async function handleAuthenticatedUser(user){
  if(_appOpening)return;_appOpening=true;
  try{
    assertAuthorizedMariana(user);
    _currentProfile='mariana';_viewOwner='mariana';
    if(window._fintrackLocalReady)await window._fintrackLocalReady;
    const sub=document.getElementById('lock-sub');if(sub)sub.textContent='Sincronizando a base exclusiva de Mariana...';
    const btn=document.getElementById('login-btn');if(btn){btn.disabled=true;btn.textContent='Carregando...'}
    await syncFromCloud(false);
    _cloudLoaded=true;updateProfileUI();populateAllSelects();refreshAll();renderCloudSection();showApp();
  }catch(err){try{await auth?.signOut()}catch{}showLogin();setLoginError('Acesso bloqueado: '+err.message)}
  finally{_appOpening=false;const sub=document.getElementById('lock-sub');if(sub)sub.textContent='Entre com a conta Mariana autorizada'}
}
function initFirebase(){
  assertMarianaFirebaseConfig();
  if(!firebaseConfigured()){
    throw new Error('Configuração interna do Firebase Mariana inválida. Abra especificamente o arquivo V4 Mariana.');
  }
  const h=document.getElementById('firebase-config-help');if(h){h.style.display='none';h.textContent='';}
  // Nunca reutiliza o app Firebase default ou um app legado já carregado na página.
  try{firebaseApp=firebase.app(FIREBASE_APP_NAME)}catch{firebaseApp=firebase.initializeApp(FIREBASE_CONFIG,FIREBASE_APP_NAME)}
  const opts=firebaseApp.options||{};
  if(opts.projectId!==EXPECTED_FIREBASE_PROJECT_ID||normalizeDatabaseUrl(opts.databaseURL)!==EXPECTED_DATABASE_URL||opts.authDomain!==EXPECTED_AUTH_DOMAIN||opts.appId!==EXPECTED_APP_ID){
    throw new Error('O app Firebase carregado não pertence integralmente ao projeto Mariana. Inicialização interrompida.');
  }
  auth=firebaseApp.auth();
  realtimeDb=firebaseApp.database();
  auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL).catch(()=>{});
  auth.onAuthStateChanged(async user=>{_authReady=true;if(user)await handleAuthenticatedUser(user);else showLogin();});
}

// ===================== STORAGE =====================
// V4 auditada: dados financeiros não são persistidos no navegador.
// O único estado persistente local é a sessão do Firebase Authentication.
function parseCloudValue(value,fallback){
  if(value==null)return fallback;
  if(typeof value==='string'){try{return JSON.parse(value)}catch{return fallback}}
  if(typeof value==='object')return value;
  return fallback;
}

const DEFAULT_CATS={
  expense:[
    {id:'moradia',name:'Moradia',icon:'🏠',color:'#60a5fa',builtin:true},
    {id:'alimentacao',name:'Alimentação',icon:'🍽️',color:'#fb923c',builtin:true},
    {id:'transporte',name:'Transporte',icon:'🚗',color:'#a78bfa',builtin:true},
    {id:'saude',name:'Saúde',icon:'❤️',color:'#f87171',builtin:true},
    {id:'educacao',name:'Educação',icon:'📚',color:'#34d399',builtin:true},
    {id:'lazer',name:'Lazer',icon:'🎮',color:'#fbbf24',builtin:true},
    {id:'vestuario',name:'Vestuário',icon:'👕',color:'#c084fc',builtin:true},
    {id:'servicos',name:'Serviços',icon:'📱',color:'#38bdf8',builtin:true},
    {id:'outros_d',name:'Outros',icon:'📦',color:'#94a3b8',builtin:true},
  ],
  income:[
    {id:'salario',name:'Salário',icon:'💼',color:'#22d3a0',builtin:true},
    {id:'freelance',name:'Freelance',icon:'💻',color:'#60a5fa',builtin:true},
    {id:'investimento',name:'Investimento',icon:'📈',color:'#fbbf24',builtin:true},
    {id:'bonus',name:'Bônus',icon:'🎁',color:'#f472b6',builtin:true},
    {id:'outros_r',name:'Outros',icon:'💰',color:'#94a3b8',builtin:true},
  ]
};

// ===================== PERFIL SOLO =====================
const PROFILES={mariana:{id:'mariana',name:'Mariana',icon:'👤'}};
let _currentProfile=null;
let _viewOwner='mariana';
function otherProfile(){return'mariana'}
function profileName(){return'Mariana'}
function updateProfileUI(){
  const badge=document.getElementById('profile-badge');if(badge)badge.textContent='👤 Mariana';
  const el=document.getElementById('view-mariana');if(el){el.style.background='var(--green)';el.style.color='#0f1117';el.style.borderColor='var(--green)';el.style.fontWeight='600'}
}
function viewSuffix(){return ' · Mariana'}
function setViewOwner(){_viewOwner='mariana';updateProfileUI();refreshAll()}
function getVisibleTx(){return _db.transactions||[]}
function visibleDB(){return{transactions:getVisibleTx()}}
function ownerTotals(txs,type='expense',category=null){const filtered=txs.filter(t=>t.type===type&&(!category||t.category===category));const total=filtered.reduce((s,t)=>s+t.amount,0);return{mariana:total,luiza:0}}
function ownerSplitHtml(totals,showValues=true){return showValues?`<div style="margin-top:6px;font-size:10px;color:var(--text3)"><span style="color:var(--blue)">Mariana ${fmt(totals.mariana||0)}</span></div>`:''}

let _db={transactions:[]},_cats=JSON.parse(JSON.stringify(DEFAULT_CATS)),_imports=[];

// V9: a nuvem é a única fonte persistente de dados. Não restauramos transações antigas do
// localStorage/window.storage, evitando que uma base apagada volte a aparecer ao reabrir a página.
async function clearLegacyLocalPersistence(){
  // Apenas remove chaves legadas; nunca grava dados financeiros localmente.
  const legacy=['fintrack_v2','fintrack_cats_v2','fintrack_imports','fintrack_pending_sync_v2'];
  for(const k of legacy){try{localStorage.removeItem(k)}catch{}}
}
async function loadAll(){
  _db={transactions:[]};_cats=JSON.parse(JSON.stringify(DEFAULT_CATS));_imports=[];
  _localDirtyIds.clear();_localDeletedIds.clear();_categoriesDirty=false;_importsDirty=false;
  await clearLegacyLocalPersistence();
}
function loadCats(){return JSON.parse(JSON.stringify(_cats))}
function loadDB(){return JSON.parse(JSON.stringify(_db))}
function loadImports(){return[]}
async function saveCats(c){
  const next=JSON.parse(JSON.stringify(c));
  if(auth?.currentUser&&realtimeDb){await soloUserRoot().child('categories').set(JSON.stringify(next));}
  else throw new Error('Firebase não está autenticado. As categorias não foram salvas.');
  _cats=next;renderCloudSection();
}
async function persistPendingSyncState(){return true}
async function saveDB(db){_db=JSON.parse(JSON.stringify(db||{transactions:[]}));return true}
async function saveImports(){return true}
async function saveDBAndSync(db){
  if(!auth?.currentUser||!realtimeDb)throw new Error('Firebase Mariana não está autenticado. A alteração não foi salva.');
  assertAuthorizedMariana(auth.currentUser);
  const next=JSON.parse(JSON.stringify(db||{transactions:[]}));
  const map={};(next.transactions||[]).forEach(t=>{if(t?.id!=null)map[firebaseSafeKey(t.id)]={...t,id:String(t.id)}});
  // Uma única gravação confirma a nova base de transações. O listener apenas lê o eco e nunca regrava.
  await soloUserRoot().child('transactions').set(Object.keys(map).length?map:null);
  _db=next;_cloudTxMap=map;_lastRemoteTxFingerprint=stableTxFingerprint(map);_cloudLoaded=true;_lastCloudError=null;
  renderCloudSection();
  return true;
}
function scheduleAutoSync(){/* V9: sem fila persistente/auto-reenvio */}
async function autoSyncToCloud(){return true}
function scheduleCloudRecovery(){/* V9: reconexão é tratada pela próxima leitura/gravação */}

// ===================== HELPERS BASE (V13) =====================
// Funções fundamentais restauradas. A remoção acidental deste bloco nas versões V9-V12
// causava o erro `populateMonthSelect is not defined` logo após o login.
function allCats(){const c=loadCats();return[...c.expense,...c.income]}
function getCat(id){return allCats().find(c=>c.id===id)||{name:id||'—',icon:'📦',color:'#888'}}
function slugify(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').substring(0,30)}
function randomColor(){const p=['#f87171','#fb923c','#fbbf24','#34d399','#60a5fa','#a78bfa','#f472b6','#38bdf8','#22d3a0','#c084fc'];return p[Math.floor(Math.random()*p.length)]}

const MONTHS=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const MONTHS_SHORT=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function fmt(v){return 'R$ '+Math.abs(Number(v)||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}

function getYears(){
  const db=loadDB(),now=new Date().getFullYear(),ys=new Set([now]);
  db.transactions.forEach(t=>{const d=new Date(t.date+'T12:00');if(!Number.isNaN(d.getTime()))ys.add(d.getFullYear())});
  return[...ys].sort((a,b)=>b-a);
}
function populateMonthSelect(id,cur){
  const sel=document.getElementById(id);if(!sel)return;
  const previous=sel.value;
  const includeAll=id.startsWith('tx-')||id==='exp-month';
  sel.innerHTML=(includeAll?'<option value="-1">Todos os meses</option>':'')+MONTHS.map((m,i)=>`<option value="${i}">${m}</option>`).join('');
  if(previous!==''&&[...sel.options].some(o=>o.value===String(previous)))sel.value=String(previous);
  else sel.value=String(cur);
}
function populateYearSelect(id,cur){
  const sel=document.getElementById(id);if(!sel)return;
  const previous=sel.value;
  const years=getYears();
  const previousYear=parseInt(previous);
  const preferred=Number.isFinite(previousYear)&&years.includes(previousYear)?previousYear:(years.includes(cur)?cur:years[0]);
  sel.innerHTML=years.map(y=>`<option value="${y}">${y}</option>`).join('');
  if(preferred!==undefined)sel.value=String(preferred);
}
function filterTx(txs,month,year){
  return (Array.isArray(txs)?txs:[]).filter(t=>{const d=new Date(t.date+'T12:00');return !Number.isNaN(d.getTime())&&d.getFullYear()===year&&(month===-1||d.getMonth()===month)});
}

// ===================== CHARTS =====================
let chartMonthly=null,chartPie=null,chartCatPie=null,chartAnual=null,chartCatAnual=null,chartEvolucao=null,chartAcumulado=null,chartInsCat=null;
function dc(c){if(c){try{c.destroy()}catch{}}return null}

// ===================== SECTIONS =====================
function showSection(name){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const sec=document.getElementById('sec-'+name);
  if(sec)sec.classList.add('active');
  const ni=document.querySelector(`.nav-item[data-section="${name}"]`);
  if(ni)ni.classList.add('active');
  document.querySelectorAll('.mob-btn').forEach(b=>b.classList.remove('active'));
  const mm={dashboard:'mob-dashboard',transacoes:'mob-transacoes',evolucao:'mob-evolucao'};
  if(mm[name])document.getElementById(mm[name])?.classList.add('active');
  document.getElementById('sidebar').classList.remove('open');
  if(name==='dashboard')updateDashboard();
  if(name==='transacoes')renderTxList();
  if(name==='evolucao')renderEvolucao();
  if(name==='comparativo')renderComparativo();
  if(name==='insights')renderInsights();
  if(name==='patrimonio')renderPatrimonio();
  if(name==='categorias')renderCategorias();
  if(name==='anual')renderAnual();
  if(name==='exportar')renderExportSection();
  if(name==='gcat')renderCatManager();
  if(name==='cloud')renderCloudSection();
  if(name==='reset')renderResetSection();
}
function toggleMobileMenu(){document.getElementById('sidebar').classList.toggle('open')}

// ===================== MODAL =====================
let currentType='expense';
function openModal(id){
  document.getElementById('modal-bg').classList.add('open');
  document.getElementById('modal-title').textContent=id?'Editar transação':'Nova transação';
  document.getElementById('f-editing-id').value=id||'';
  // reset parcelamento
  document.getElementById('f-parceled').checked=false;
  document.getElementById('f-parcelas-row').style.display='none';
  document.getElementById('f-parcelas').value=2;
  document.getElementById('f-parcela-valor').value='';
  document.getElementById('f-recurring').checked=false;
  document.getElementById('f-recurring-row').style.display='none';
  document.getElementById('f-recurring-months').value=12;
  document.getElementById('f-recurring-value').value='';
  document.getElementById('f-recurring-check-row').style.display=id?'none':'block';
  // hide parcelamento when editing
  document.getElementById('f-parceled').closest('.form-row').style.display=id?'none':'block';
  document.getElementById('f-parcelas-row').style.display='none';
  // reset compartilhado
  document.getElementById('f-shared').checked=false;
  document.getElementById('f-shared-split-row').style.display='none';
  document.getElementById('f-shared-mine-pct').value=50;
  document.getElementById('f-shared-partner-pct').value=50;
  document.getElementById('f-shared-partner-name').textContent=profileName(otherProfile(_currentProfile));
  document.getElementById('f-shared-mine-label').textContent='Minha parte ('+profileName(_currentProfile)+') %';
  document.getElementById('f-shared-partner-label').textContent='Parte de '+profileName(otherProfile(_currentProfile))+' (%)';
  if(id){
    const t=loadDB().transactions.find(x=>x.id===id);if(!t)return;
    setType(t.type);
    document.getElementById('f-desc').value=t.description;
    document.getElementById('f-amount').value=t.amount;
    document.getElementById('f-date').value=t.date;
    const c=getCat(t.category);
    document.getElementById('f-cat').value=t.category;
    document.getElementById('f-cat-search').value=c.icon+' '+c.name;
    // não permite alterar o compartilhamento na edição — só na criação
    document.getElementById('f-shared-row').style.display='none';
    if(t.shared){
      document.getElementById('f-shared-split-row').style.display='block';
      document.getElementById('f-shared-split-row').innerHTML=`<div style="font-size:12px;color:var(--text2)">🤝 Esta transação faz parte de um gasto compartilhado (total ${fmt(t.sharedTotal||0)}, pago por ${profileName(t.paidBy)}). Editar aqui só ajusta esta linha — a parte do(a) ${profileName(otherProfile(t.owner))} não muda automaticamente. Se precisar desfazer a divisão, exclua e lance novamente.</div>`;
    }
  }else{
    document.getElementById('f-date').value=new Date().toISOString().split('T')[0];
    document.getElementById('f-desc').value='';
    document.getElementById('f-amount').value='';
    document.getElementById('f-shared-row').style.display='none';
    setType('expense');
  }
  // auto-calc parcela when amount changes
  document.getElementById('f-amount').oninput=()=>{if(document.getElementById('f-parceled').checked)calcParcelaValor();updateRecurringPreview();updateSharedPreview()};
  document.getElementById('f-parcelas').oninput=calcParcelaValor;
  updateSharedPreview();
}
function closeModal(){document.getElementById('modal-bg').classList.remove('open')}
function closeModalBg(e){if(e.target.id==='modal-bg')closeModal()}

function setType(type){
  currentType=type;
  document.getElementById('rb-income').className='radio-btn'+(type==='income'?' sel-income':'');
  document.getElementById('rb-expense').className='radio-btn'+(type==='expense'?' sel-expense':'');
  document.getElementById('f-cat').value='';
  document.getElementById('f-cat-search').value='';
  renderCatDropdown('');
  // mostra dica de valor negativo apenas para receita
  document.getElementById('f-amount-hint').style.display=type==='income'?'block':'none';
  // parcelamento só faz sentido para despesas
  const pRow=document.getElementById('f-parceled').closest('.form-row');
  const rRow=document.getElementById('f-recurring-check-row');
  if(type==='income'){
    document.getElementById('f-parceled').checked=false;
    document.getElementById('f-parcelas-row').style.display='none';
    document.getElementById('f-recurring').checked=false;
    document.getElementById('f-recurring-row').style.display='none';
    pRow.style.display='none';
    if(rRow)rRow.style.display='none';
  } else if(!document.getElementById('f-editing-id').value){
    pRow.style.display='block';
    if(rRow)rRow.style.display='block';
  }
}
function getAllCatsForType(type){const c=loadCats();return[...c[type].map(x=>({...x,_primary:true})),...c[type==='expense'?'income':'expense'].map(x=>({...x,_primary:false}))]}
function renderCatDropdown(filter){
  const dd=document.getElementById('f-cat-dropdown');if(!dd)return;
  const all=getAllCatsForType(currentType);
  const q=filter.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const filtered=all.filter(c=>c.name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').includes(q));
  if(!filtered.length){dd.innerHTML='<div style="padding:10px 14px;font-size:12px;color:var(--text3)">Nenhuma encontrada</div>';return}
  const pri=filtered.filter(c=>c._primary),sec=filtered.filter(c=>!c._primary);
  let html=pri.map(c=>catOpt(c)).join('');
  if(sec.length){if(pri.length)html+=`<div style="padding:4px 14px;font-size:10px;color:var(--text3);background:var(--bg4);border-top:1px solid var(--border)">Outras</div>`;html+=sec.map(c=>catOpt(c)).join('')}
  dd.innerHTML=html;
}
function catOpt(c){
  const sel=document.getElementById('f-cat')?.value===c.id;
  return`<div onclick="selectCat('${c.id}','${c.name.replace(/'/g,"\\'")}','${c.icon}')" style="display:flex;align-items:center;gap:9px;padding:9px 14px;cursor:pointer;font-size:13px;background:${sel?'var(--bg4)':'transparent'}" onmouseover="this.style.background='var(--bg4)'" onmouseout="this.style.background='${sel?'var(--bg4)':'transparent'}'"><span style="width:8px;height:8px;border-radius:50%;background:${c.color};flex-shrink:0"></span><span>${c.icon}</span><span>${c.name}</span>${sel?'<span style="margin-left:auto;color:var(--green);font-size:11px">✓</span>':''}</div>`;
}
function selectCat(id,name,icon){document.getElementById('f-cat').value=id;document.getElementById('f-cat-search').value=icon+' '+name;document.getElementById('f-cat-dropdown').style.display='none'}
function showCatDropdown(){document.getElementById('f-cat-dropdown').style.display='block';renderCatDropdown(document.getElementById('f-cat-search').value);setTimeout(()=>document.addEventListener('click',function h(e){if(!document.getElementById('f-cat-dropdown')?.contains(e.target)&&e.target!==document.getElementById('f-cat-search')){document.getElementById('f-cat-dropdown').style.display='none';document.removeEventListener('click',h)}},{once:true}),10)}
function filterCatOptions(){document.getElementById('f-cat-dropdown').style.display='block';renderCatDropdown(document.getElementById('f-cat-search').value)}

function toggleParcelamento(){
  const checked=document.getElementById('f-parceled').checked;
  document.getElementById('f-parcelas-row').style.display=checked?'grid':'none';
  if(checked){
    document.getElementById('f-recurring').checked=false;
    document.getElementById('f-recurring-row').style.display='none';
    calcParcelaValor();
  }
}
function calcParcelaValor(){
  const total=parseFloat(document.getElementById('f-amount').value)||0;
  const n=parseInt(document.getElementById('f-parcelas').value)||2;
  document.getElementById('f-parcela-valor').value=n>0?(total/n).toFixed(2):'';
}

function toggleRecorrencia(){
  const checked=document.getElementById('f-recurring').checked;
  document.getElementById('f-recurring-row').style.display=checked?'block':'none';
  if(checked){
    document.getElementById('f-parceled').checked=false;
    document.getElementById('f-parcelas-row').style.display='none';
  }
  updateRecurringPreview();
}
function updateRecurringPreview(){
  const value=parseFloat(document.getElementById('f-amount')?.value)||0;
  const el=document.getElementById('f-recurring-value');
  if(el)el.value=fmt(value)+' por mês';
}

function toggleShared(){
  const checked=document.getElementById('f-shared').checked;
  document.getElementById('f-shared-split-row').style.display=checked?'block':'none';
  updateSharedPreview();
}
function syncSharedPct(which){
  const mineEl=document.getElementById('f-shared-mine-pct'),partnerEl=document.getElementById('f-shared-partner-pct');
  let mine=parseFloat(mineEl.value)||0,partner=parseFloat(partnerEl.value)||0;
  if(which==='mine')partner=100-mine; else mine=100-partner;
  mine=Math.max(0,Math.min(100,mine));partner=Math.max(0,Math.min(100,partner));
  mineEl.value=mine;partnerEl.value=partner;
  updateSharedPreview();
}
function updateSharedPreview(){
  const prev=document.getElementById('f-shared-preview');if(!prev)return;
  if(!document.getElementById('f-shared').checked){prev.textContent='—';return}
  const total=parseFloat(document.getElementById('f-amount').value)||0;
  const mine=parseFloat(document.getElementById('f-shared-mine-pct').value)||0;
  const partner=parseFloat(document.getElementById('f-shared-partner-pct').value)||0;
  prev.textContent=`Total ${fmt(total)} → ${profileName(_currentProfile)}: ${fmt(total*mine/100)} (${mine}%) · ${profileName(otherProfile(_currentProfile))}: ${fmt(total*partner/100)} (${partner}%)`;
}

async function saveTransaction(){
  const desc=document.getElementById('f-desc').value.trim();
  const amount=parseFloat(document.getElementById('f-amount').value);
  const date=document.getElementById('f-date').value;
  const cat=document.getElementById('f-cat').value;
  if(!desc||!date||isNaN(amount)||amount===0){alert('Preencha todos os campos.');return}
  if(currentType==='expense'&&amount<0){alert('Despesas devem ter valor positivo.');return}
  if(!cat){alert('Selecione uma categoria.');return}
  const db=loadDB();
  const editingId=document.getElementById('f-editing-id').value;
  const isParceled=document.getElementById('f-parceled').checked&&!editingId;
  const isRecurring=document.getElementById('f-recurring').checked&&!editingId;
  const isShared=false; // V9 SOLO Mariana: divisão entre perfis desativada
  if(isRecurring){
    const months=parseInt(document.getElementById('f-recurring-months').value)||12;
    if(months<2||months>120){alert('A recorrência deve ter entre 2 e 120 meses.');return}
    const baseDate=new Date(date+'T12:00');
    const recurringGroup='rec_'+Date.now().toString(36);
    const sharedChecked=document.getElementById('f-shared').checked;
    const minePct=parseFloat(document.getElementById('f-shared-mine-pct').value)||0;
    const partnerPct=parseFloat(document.getElementById('f-shared-partner-pct').value)||0;
    if(sharedChecked&&Math.round(minePct+partnerPct)!==100){alert('A divisão precisa somar 100%.');return}
    const partner=otherProfile(_currentProfile);

    for(let i=0;i<months;i++){
      const d=new Date(baseDate);
      const day=baseDate.getDate();
      d.setDate(1);
      d.setMonth(d.getMonth()+i);
      const lastDay=new Date(d.getFullYear(),d.getMonth()+1,0).getDate();
      d.setDate(Math.min(day,lastDay));
      const ds=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      const monthGroup=recurringGroup+'_'+i;

      if(sharedChecked){
        db.transactions.unshift({
          id:Date.now().toString()+Math.random().toString(36).slice(2),type:currentType,description:desc,
          amount:parseFloat((amount*minePct/100).toFixed(2)),date:ds,category:cat,
          owner:_currentProfile,shared:true,groupId:monthGroup,sharedTotal:amount,splitPct:minePct,paidBy:_currentProfile,
          recurring:true,recurringGroup,recurringNum:i+1,recurringTotal:months
        });
        db.transactions.unshift({
          id:Date.now().toString()+Math.random().toString(36).slice(2),type:currentType,description:desc,
          amount:parseFloat((amount*partnerPct/100).toFixed(2)),date:ds,category:cat,
          owner:partner,shared:true,groupId:monthGroup,sharedTotal:amount,splitPct:partnerPct,paidBy:_currentProfile,
          recurring:true,recurringGroup,recurringNum:i+1,recurringTotal:months
        });
      }else{
        db.transactions.unshift({
          id:Date.now().toString()+Math.random().toString(36).slice(2),type:currentType,description:desc,
          amount,date:ds,category:cat,owner:_currentProfile,
          recurring:true,recurringGroup,recurringNum:i+1,recurringTotal:months
        });
      }
    }
  } else if(isShared){
    const minePct=parseFloat(document.getElementById('f-shared-mine-pct').value)||0;
    const partnerPct=parseFloat(document.getElementById('f-shared-partner-pct').value)||0;
    if(Math.round(minePct+partnerPct)!==100){alert('A divisão precisa somar 100%.');return}
    const partner=otherProfile(_currentProfile);
    const groupId='sh_'+Date.now().toString(36);
    db.transactions.unshift({
      id:Date.now().toString(),type:currentType,description:desc,
      amount:parseFloat((amount*minePct/100).toFixed(2)),date,category:cat,
      owner:_currentProfile,shared:true,groupId,sharedTotal:amount,splitPct:minePct,paidBy:_currentProfile
    });
    db.transactions.unshift({
      id:Date.now().toString()+'_p',type:currentType,description:desc,
      amount:parseFloat((amount*partnerPct/100).toFixed(2)),date,category:cat,
      owner:partner,shared:true,groupId,sharedTotal:amount,splitPct:partnerPct,paidBy:_currentProfile
    });
  } else if(isParceled){
    const n=parseInt(document.getElementById('f-parcelas').value)||2;
    if(n<2||n>120){alert('Número de parcelas deve ser entre 2 e 120.');return}
    const parcelaValor=amount/n;
    const baseDate=new Date(date+'T12:00');
    const groupId='parc_'+Date.now().toString(36);
    for(let i=0;i<n;i++){
      const d=new Date(baseDate);
      d.setMonth(d.getMonth()+i);
      const ds=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      db.transactions.unshift({
        id:Date.now().toString()+Math.random().toString(36).slice(2),
        type:currentType,
        description:`${desc} (${i+1}/${n})`,
        amount:parseFloat(parcelaValor.toFixed(2)),
        date:ds,
        category:cat,
        parcelaGroup:groupId,
        parcelaNum:i+1,
        parcelaTotal:n,
        owner:_currentProfile
      });
    }
  } else {
    if(editingId){
      const idx=db.transactions.findIndex(t=>t.id===editingId);
      if(idx>=0)db.transactions[idx]={...db.transactions[idx],type:currentType,description:desc,amount,date,category:cat};
    } else {
      db.transactions.unshift({id:Date.now().toString(),type:currentType,description:desc,amount,date,category:cat,owner:_currentProfile});
    }
  }
  db.transactions.sort((a,b)=>new Date(b.date)-new Date(a.date));
  try{await saveDBAndSync(db)}catch(err){alert('Não foi possível salvar no Firebase. A alteração não foi confirmada.\n'+err.message);return}
  closeModal();populateAllSelects();refreshAll();
}
async function delTx(id){
  const db=loadDB();const t=db.transactions.find(x=>x.id===id);if(!t)return;
  if(!confirm(t.parcelaGroup?'Excluir somente esta parcela?':'Excluir esta transação?'))return;
  db.transactions=db.transactions.filter(x=>x.id!==id);
  try{await saveDBAndSync(db);populateAllSelects();refreshAll()}catch(err){alert('Não foi possível excluir no Firebase.\n'+err.message)}
}
async function delParcelGroup(id){
  const db=loadDB();const t=db.transactions.find(x=>x.id===id);if(!t?.parcelaGroup)return;
  const group=db.transactions.filter(x=>x.parcelaGroup===t.parcelaGroup);
  if(!confirm(`Excluir TODAS as ${group.length} parcelas deste gasto?\n\nEsta ação remove todos os lançamentos vinculados e não pode ser desfeita.`))return;
  db.transactions=db.transactions.filter(x=>x.parcelaGroup!==t.parcelaGroup);
  try{await saveDBAndSync(db);populateAllSelects();refreshAll();alert(`✓ ${group.length} parcelas excluídas do Firebase.`)}catch(err){alert('Não foi possível excluir as parcelas no Firebase.\n'+err.message)}
}

// ===================== DASHBOARD =====================
function updateDashboard(){
  const now=new Date();
  const month=parseInt(document.getElementById('dash-month')?.value??now.getMonth());
  const year=parseInt(document.getElementById('dash-year')?.value??now.getFullYear());
  document.getElementById('dash-period-label').textContent=MONTHS[month]+' de '+year+viewSuffix();
  const db=visibleDB(),txs=filterTx(db.transactions,month,year);
  const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const balance=income-expense;
  document.getElementById('c-income').textContent=fmt(income);
  document.getElementById('c-income-count').textContent=txs.filter(t=>t.type==='income').length+' entrada(s)';
  document.getElementById('c-expense').textContent=fmt(expense);
  document.getElementById('c-expense-count').textContent=txs.filter(t=>t.type==='expense').length+' saída(s)';
  document.getElementById('c-balance').textContent=fmt(balance);
  document.getElementById('c-balance').style.color=balance>=0?'var(--green)':'var(--red)';
  document.getElementById('c-balance-pct').textContent=income>0?'Poupança: '+Math.round((balance/income)*100)+'%':'—';
  const catMap={};txs.filter(t=>t.type==='expense').forEach(t=>{catMap[t.category]=(catMap[t.category]||0)+t.amount});
  const top=Object.entries(catMap).sort((a,b)=>b[1]-a[1])[0];
  if(top){const c=getCat(top[0]);document.getElementById('c-topcat').textContent=c.icon+' '+c.name;document.getElementById('c-topcat-amt').textContent=fmt(top[1])}
  else{document.getElementById('c-topcat').textContent='—';document.getElementById('c-topcat-amt').textContent='Sem despesas'}
  const last6=[];
  for(let i=5;i>=0;i--){let m=month-i,y=year;if(m<0){m+=12;y--}const t=filterTx(db.transactions,m,y);last6.push({label:MONTHS_SHORT[m],income:t.filter(x=>x.type==='income').reduce((s,x)=>s+x.amount,0),expense:t.filter(x=>x.type==='expense').reduce((s,x)=>s+x.amount,0)})}
  chartMonthly=dc(chartMonthly);
  chartMonthly=new Chart(document.getElementById('chart-monthly'),{type:'bar',data:{labels:last6.map(x=>x.label),datasets:[{label:'Receitas',data:last6.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},{label:'Despesas',data:last6.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  const expCats=allCats().filter(c=>loadCats().expense.find(e=>e.id===c.id)).map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartPie=dc(chartPie);
  const pieLeg=document.getElementById('pie-legend');
  if(!expCats.length){pieLeg.innerHTML='<div style="color:var(--text3);font-size:12px">Sem despesas</div>';return}
  chartPie=new Chart(document.getElementById('chart-pie'),{type:'doughnut',data:{labels:expCats.map(x=>x.cat.name),datasets:[{data:expCats.map(x=>x.total),backgroundColor:expCats.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  pieLeg.innerHTML=expCats.map(x=>`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-family:DM Mono;font-size:12px">${fmt(x.total)}</span></span>`).join('');
  pieLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
  renderTxItems(document.getElementById('recent-list'),db.transactions.slice(0,8));
  renderBalanceCard(month,year);
}

// ===================== BALANÇO DO CASAL =====================
// Considera SEMPRE os dados completos (não é afetado pelo filtro Meu/Dela/Casal),
// pois o balanço é, por natureza, uma visão conjunta.
function renderBalanceCard(month,year){
  const el=document.getElementById('balance-card-body');
  if(!el)return;
  const monthTx=filterTx(_db.transactions,month,year).filter(t=>t.shared&&t.groupId);
  const seen=new Set();
  let owedToMariana=0, owedTo=0, totalShared=0;
  monthTx.forEach(t=>{
    if(seen.has(t.groupId))return;
    seen.add(t.groupId);
    const group=_db.transactions.filter(x=>x.groupId===t.groupId);
    const payer=t.paidBy;
    if(!payer||group.length<2)return;
    totalShared+=t.sharedTotal||0;
    const otherLine=group.find(x=>x.owner!==payer);
    if(!otherLine)return;
    if(payer==='mariana')owedToMariana+=otherLine.amount;
    else if(payer==='mariana')owedTo+=otherLine.amount;
  });
  if(!seen.size){
    el.innerHTML='<div class="empty" style="padding:14px 0"><div class="empty-icon">🤝</div>Nenhum gasto compartilhado neste mês ainda.</div>';
    return;
  }
  const net=owedToMariana-owedTo; // >0:  deve a Mariana; <0: Mariana deve a 
  let resultLine;
  if(Math.abs(net)<0.01){
    resultLine=`<span style="color:var(--green);font-weight:600">✅ Contas equilibradas — ninguém deve nada este mês.</span>`;
  } else if(net>0){
    resultLine=`<span style="color:var(--amber);font-weight:600"> deve ${fmt(net)} a Mariana</span>`;
  } else {
    resultLine=`<span style="color:var(--amber);font-weight:600">Mariana deve ${fmt(Math.abs(net))} a </span>`;
  }
  el.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:14px">
      <div><div style="font-size:11px;color:var(--text3);margin-bottom:3px">Total pago por Mariana (compartilhado)</div><div style="font-family:DM Mono;font-size:15px;font-weight:600">${fmt((_db.transactions.filter(t=>t.shared&&t.paidBy==='mariana'&&filterTx([t],month,year).length).reduce((s,t2)=>s+ (t2.sharedTotal&&t2.owner===t2.paidBy?t2.sharedTotal:0),0)))}</div></div>
      <div><div style="font-size:11px;color:var(--text3);margin-bottom:3px">Total pago por  (compartilhado)</div><div style="font-family:DM Mono;font-size:15px;font-weight:600">${fmt((_db.transactions.filter(t=>t.shared&&t.paidBy==='mariana'&&filterTx([t],month,year).length).reduce((s,t2)=>s+ (t2.sharedTotal&&t2.owner===t2.paidBy?t2.sharedTotal:0),0)))}</div></div>
    </div>
    <div style="padding:12px 14px;background:var(--bg3);border-radius:var(--radius-sm);font-size:14px">${resultLine}</div>
  `;
}

// ===================== TX LIST com filtros =====================
let txTypeFilter='all', txCatFilters=new Set();

function setTxTypeFilter(type){
  txTypeFilter=type;
  ['all','income','expense'].forEach(t=>{
    const chip=document.getElementById('chip-type-'+t);
    if(!chip)return;
    chip.className='filter-chip'+(type===t?(t==='all'?' active':t==='income'?' active-income':' active-expense'):'');
  });
  renderTxList();
}

function toggleCatFilter(catId){
  if(txCatFilters.has(catId))txCatFilters.delete(catId);
  else txCatFilters.add(catId);
  renderTxList();
}

function buildCatFilterChips(){
  const usedCats=[...new Set((_viewOwner==='mariana'?_db.transactions:_db.transactions.filter(t=>t.owner===_viewOwner)).map(t=>t.category))];
  const container=document.getElementById('cat-filter-chips');
  if(!container)return;
  container.innerHTML=usedCats.map(cid=>{
    const c=getCat(cid),active=txCatFilters.has(cid);
    return`<div class="filter-chip ${active?'active':''}" onclick="toggleCatFilter('${cid}')" id="catfc-${cid}"><span style="width:7px;height:7px;border-radius:50%;background:${c.color};display:inline-block"></span>${c.icon} ${c.name}${active?'<span class="chip-x">✕</span>':''}</div>`;
  }).join('');
}

function renderTxList(){
  buildCatFilterChips();
  const month=parseInt(document.getElementById('tx-month')?.value??-1);
  const year=parseInt(document.getElementById('tx-year')?.value??new Date().getFullYear());
  const listOwner=_viewOwner;
  const db={transactions:_viewOwner==='mariana'?_db.transactions:_db.transactions.filter(t=>t.owner===_viewOwner)};
  let txs=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(txTypeFilter!=='all')txs=txs.filter(t=>t.type===txTypeFilter);
  if(txCatFilters.size>0)txs=txs.filter(t=>txCatFilters.has(t.category));
  const lbl=document.getElementById('tx-count-label');
  if(lbl)lbl.textContent=txs.length+' transação(ões) · '+(_viewOwner==='mariana'?'Mariana':profileName(listOwner));
  const title=document.getElementById('tx-list-title');
  if(title)title.textContent=txs.length+' transação(ões) encontrada(s)';
  // mostrar botão excluir filtradas só quando há filtro ativo
  const hasFilter=txTypeFilter!=='all'||txCatFilters.size>0||month!==-1;
  const delBtn=document.getElementById('btn-delete-filtered');
  if(delBtn)delBtn.style.display=(hasFilter&&txs.length>0)?'inline-flex':'none';
  renderTxItems(document.getElementById('tx-list'),txs);
}

async function deleteFiltered(){
  const month=parseInt(document.getElementById('tx-month')?.value??-1);
  const year=parseInt(document.getElementById('tx-year')?.value??new Date().getFullYear());
  const listOwner=_viewOwner==='mariana'?_currentProfile:_viewOwner;
  const db={transactions:_db.transactions.filter(t=>t.owner===listOwner)};
  let toDelete=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(txTypeFilter!=='all')toDelete=toDelete.filter(t=>t.type===txTypeFilter);
  if(txCatFilters.size>0)toDelete=toDelete.filter(t=>txCatFilters.has(t.category));
  if(!toDelete.length){alert('Nenhuma transação selecionada.');return}
  if(!confirm(`Excluir ${toDelete.length} transação(ões) filtrada(s)? Esta ação não pode ser desfeita.`))return;
  const deleteIds=new Set(toDelete.map(t=>t.id));
  const fullDb=loadDB();fullDb.transactions=fullDb.transactions.filter(t=>!deleteIds.has(t.id));
  try{await saveDBAndSync(fullDb)}catch(err){alert('Não foi possível excluir no Firebase.\n'+err.message);return}
  refreshAll();
}

function renderTxItems(container,txs){
  if(!container)return;
  if(!txs.length){container.innerHTML='<div class="empty"><div class="empty-icon">🔍</div>Nenhuma transação encontrada</div>';return}
  container.innerHTML=txs.map(t=>{
    const c=getCat(t.category),d=new Date(t.date+'T12:00');
    const isNegIncome=t.type==='income'&&t.amount<0;
    const amtColor=t.type==='expense'?'var(--red)':(isNegIncome?'var(--amber)':'var(--green)');
    const sign=t.amount<0?'-':(t.type==='income'?'+':'-');
    const ownerTag=t.owner?`<span style="font-size:10px;padding:1px 6px;border-radius:10px;background:var(--bg4);color:var(--text3);margin-left:6px">${profileName(t.owner)}</span>`:'';
    const sharedTag=t.shared?`<span style="font-size:10px;padding:1px 6px;border-radius:10px;background:var(--green-bg);color:var(--green);margin-left:4px" title="Total ${fmt(t.sharedTotal||0)} · pago por ${profileName(t.paidBy)} · ${t.splitPct}%">🤝 ${t.splitPct}%</span>`:'';
    const recurringTag=t.recurring?`<span class="recurring-tag" title="Recorrência ${t.recurringNum||''}/${t.recurringTotal||''}">🔁 ${t.recurringNum||''}/${t.recurringTotal||''}</span>`:'';
    return`<div class="tx-item"><div class="tx-icon" style="background:${c.color}22">${c.icon}</div><div class="tx-info"><div class="tx-name">${t.description}${ownerTag}${sharedTag}${recurringTag}</div><div class="tx-cat">${c.name}${isNegIncome?' · abatimento':''}</div></div><div class="tx-date">${d.toLocaleDateString('pt-BR')}</div><div class="tx-amount" style="color:${amtColor}">${sign} ${fmt(Math.abs(t.amount))}</div><div class="tx-actions"><button class="tx-btn" onclick="openModal('${t.id}')" title="Editar">✏️</button>${t.parcelaGroup?`<button class="tx-btn" onclick="delParcelGroup('${t.id}')" title="Excluir todas as parcelas vinculadas">🧾✕</button>`:''}<button class="tx-btn" onclick="delTx('${t.id}')" title="Excluir somente este registro">🗑</button></div></div>`;
  }).join('');
}

// ===================== EVOLUÇÃO =====================
function renderEvolucao(){
  const year=parseInt(document.getElementById('ev-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const months=MONTHS.map((name,m)=>{
    const txs=filterTx(db.transactions,m,year);
    const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
    const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
    return{name,short:MONTHS_SHORT[m],income,expense,balance:income-expense,count:txs.length};
  });
  // acumulado
  let acc=0;const accumulated=months.map(m=>{acc+=m.balance;return acc});
  // cards
  const bestIncome=months.reduce((a,b)=>b.income>a.income?b:a,months[0]);
  const worstExpense=months.reduce((a,b)=>b.expense>a.expense?b:a,months[0]);
  document.getElementById('ev-best-income').textContent=bestIncome.income>0?bestIncome.name:'—';
  document.getElementById('ev-best-income-val').textContent=bestIncome.income>0?fmt(bestIncome.income):'Sem receitas';
  document.getElementById('ev-worst-expense').textContent=worstExpense.expense>0?worstExpense.name:'—';
  document.getElementById('ev-worst-expense-val').textContent=worstExpense.expense>0?fmt(worstExpense.expense):'Sem despesas';
  document.getElementById('ev-accumulated').textContent=fmt(acc);
  document.getElementById('ev-accumulated').style.color=acc>=0?'var(--green)':'var(--red)';
  document.getElementById('ev-accumulated-sub').textContent='Saldo acumulado em '+year;
  // médias
  const activeMths=months.filter(m=>m.count>0);
  const avgInc=activeMths.length?activeMths.reduce((s,m)=>s+m.income,0)/activeMths.length:0;
  const avgExp=activeMths.length?activeMths.reduce((s,m)=>s+m.expense,0)/activeMths.length:0;
  const avgBal=activeMths.length?activeMths.reduce((s,m)=>s+m.balance,0)/activeMths.length:0;
  document.getElementById('ev-medias').innerHTML=[
    {label:'Receita média mensal',value:fmt(avgInc),color:'var(--green)'},
    {label:'Despesa média mensal',value:fmt(avgExp),color:'var(--red)'},
    {label:'Saldo médio mensal',value:fmt(avgBal),color:avgBal>=0?'var(--green)':'var(--red)'},
    {label:'Meses com movimentação',value:activeMths.length+' de 12',color:'var(--text)'},
  ].map(r=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border)"><span style="font-size:13px;color:var(--text2)">${r.label}</span><span style="font-family:DM Mono;font-size:13px;color:${r.color}">${r.value}</span></div>`).join('');
  // gráfico principal
  chartEvolucao=dc(chartEvolucao);
  chartEvolucao=new Chart(document.getElementById('chart-evolucao'),{
    type:'bar',
    data:{labels:MONTHS_SHORT,datasets:[
      {label:'Receitas',data:months.map(m=>m.income),backgroundColor:'rgba(34,211,160,0.65)',borderRadius:4},
      {label:'Despesas',data:months.map(m=>m.expense),backgroundColor:'rgba(245,101,101,0.65)',borderRadius:4},
      {type:'line',label:'Saldo',data:months.map(m=>m.balance),borderColor:'rgba(96,165,250,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3},
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });
  // gráfico acumulado
  chartAcumulado=dc(chartAcumulado);
  chartAcumulado=new Chart(document.getElementById('chart-acumulado'),{
    type:'line',
    data:{labels:MONTHS_SHORT,datasets:[{label:'Saldo acumulado',data:accumulated,borderColor:'rgba(251,191,36,0.9)',backgroundColor:'rgba(251,191,36,0.08)',borderWidth:2,pointRadius:4,tension:0.3,fill:true}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });
}

// ===================== INSIGHTS =====================
function renderInsights(){
  const year=parseInt(document.getElementById('ins-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const allTxs=db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);
  // year totals
  const months=MONTHS.map((name,m)=>{
    const txs=filterTx(db.transactions,m,year);
    const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
    const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
    return{name,income,expense,balance:income-expense,count:txs.length};
  });
  const totInc=months.reduce((s,m)=>s+m.income,0),totExp=months.reduce((s,m)=>s+m.expense,0);
  // top individual tx — maior receita positiva
  const topInc=allTxs.filter(t=>t.type==='income').sort((a,b)=>b.amount-a.amount)[0];
  const topExp=allTxs.filter(t=>t.type==='expense').sort((a,b)=>b.amount-a.amount)[0];
  document.getElementById('ins-top-income').textContent=topInc?fmt(topInc.amount):'—';
  document.getElementById('ins-top-income-sub').textContent=topInc?topInc.description+' · '+new Date(topInc.date+'T12:00').toLocaleDateString('pt-BR'):'Nenhuma receita';
  document.getElementById('ins-top-expense').textContent=topExp?fmt(topExp.amount):'—';
  document.getElementById('ins-top-expense-sub').textContent=topExp?topExp.description+' · '+new Date(topExp.date+'T12:00').toLocaleDateString('pt-BR'):'Nenhuma despesa';
  // meses com movimentação de receita (inclusive negativos)
  const actMths=months.filter(m=>m.count>0&&allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getMonth()===months.indexOf(m)).length>0);
  const avgSave=actMths.length?actMths.reduce((s,m)=>s+(m.income!==0?(m.balance/Math.abs(m.income))*100:0),0)/actMths.length:0;
  document.getElementById('ins-save-rate').textContent=avgSave.toFixed(1)+'%';
  document.getElementById('ins-save-rate').style.color=avgSave>=0?'var(--green)':'var(--red)';
  document.getElementById('ins-save-sub').textContent='Média dos meses com receita';
  // ranking meses despesa
  const rankExp=months.filter(m=>m.expense>0).sort((a,b)=>b.expense-a.expense);
  const maxExp=rankExp[0]?.expense||1;
  document.getElementById('ins-rank-expense').innerHTML=rankExp.slice(0,6).map((m,i)=>{
    const monthIdx=MONTHS.indexOf(m.name);
    const monthTx=allTxs.filter(t=>t.type==='expense'&&new Date(t.date+'T12:00').getMonth()===monthIdx);
    const split=_viewOwner==='mariana'?ownerSplitHtml(ownerTotals(monthTx,'expense')):'';
    return`<div class="rank-row" style="display:block">
      <div style="display:flex;align-items:center;gap:10px">
        <div class="rank-num">${i+1}</div>
        <div style="font-size:13px;color:var(--text2);min-width:70px">${m.name.substring(0,3)}</div>
        <div class="rank-bar-wrap"><div class="rank-bar" style="width:${(m.expense/maxExp*100).toFixed(0)}%;background:var(--red)"></div></div>
        <div style="font-family:DM Mono;font-size:12px;color:var(--red);min-width:90px;text-align:right">${fmt(m.expense)}</div>
      </div>${split}
    </div>`;
  }).join('')||'<div class="empty">Sem dados</div>';

  const rankInc=months.filter(m=>{
    const mIdx=MONTHS.indexOf(m.name);
    return allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getMonth()===mIdx).length>0;
  }).sort((a,b)=>b.income-a.income);
  const maxIncAbs=Math.max(...rankInc.map(m=>Math.abs(m.income)),1);
  document.getElementById('ins-rank-income').innerHTML=rankInc.slice(0,6).map((m,i)=>{
    const color=m.income<0?'var(--amber)':'var(--green)';
    const monthIdx=MONTHS.indexOf(m.name);
    const monthTx=allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getMonth()===monthIdx);
    const split=_viewOwner==='mariana'?ownerSplitHtml(ownerTotals(monthTx,'income')):'';
    return`<div class="rank-row" style="display:block">
      <div style="display:flex;align-items:center;gap:10px">
        <div class="rank-num">${i+1}</div>
        <div style="font-size:13px;color:var(--text2);min-width:70px">${m.name.substring(0,3)}</div>
        <div class="rank-bar-wrap"><div class="rank-bar" style="width:${(Math.abs(m.income)/maxIncAbs*100).toFixed(0)}%;background:${color}"></div></div>
        <div style="font-family:DM Mono;font-size:12px;color:${color};min-width:90px;text-align:right">${m.income<0?'−':''}${fmt(m.income)}</div>
      </div>${split}
    </div>`;
  }).join('')||'<div class="empty">Sem dados</div>';
  // % por categoria despesa
  const cats=loadCats();
  const catData=cats.expense.map(c=>({cat:c,total:allTxs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  // categorias de receita com transações (inclusive negativas)
  const incCatData=cats.income.map(c=>{
    const catTxs=allTxs.filter(t=>t.type==='income'&&t.category===c.id);
    return{cat:c,total:catTxs.reduce((s,t)=>s+t.amount,0),count:catTxs.length};
  }).filter(x=>x.count>0).sort((a,b)=>b.total-a.total);
  chartInsCat=dc(chartInsCat);
  if(catData.length){
    chartInsCat=new Chart(document.getElementById('chart-ins-cat'),{type:'doughnut',data:{labels:catData.map(x=>x.cat.name),datasets:[{data:catData.map(x=>x.total),backgroundColor:catData.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)+' ('+(totExp>0?(ctx.raw/totExp*100).toFixed(1):0)+'%)'}}}}});
  }
  const expenseHtml=catData.map(x=>{
    const pct=totExp>0?(x.total/totExp*100).toFixed(1):0;
    const split=_viewOwner==='mariana'?ownerSplitHtml(ownerTotals(allTxs,'expense',x.cat.id)):'';
    return`<div class="rank-row" style="display:block">
      <div style="display:flex;align-items:center;gap:10px">
        <div style="width:10px;height:10px;border-radius:50%;background:${x.cat.color};flex-shrink:0"></div>
        <div style="font-size:13px;flex:1">${x.cat.icon} ${x.cat.name}</div>
        <div style="font-size:11px;color:var(--text3);margin-right:8px">${pct}%</div>
        <div style="font-family:DM Mono;font-size:12px;color:var(--red)">${fmt(x.total)}</div>
      </div>${split}
    </div>`;
  }).join('');

  const incomeHtml=incCatData.map(x=>{
    const incomeTotal=incCatData.reduce((s,i)=>s+Math.max(0,i.total),0);
    const pct=incomeTotal>0?(Math.max(0,x.total)/incomeTotal*100).toFixed(1):0;
    const split=_viewOwner==='mariana'?ownerSplitHtml(ownerTotals(allTxs,'income',x.cat.id)):'';
    const color=x.total<0?'var(--amber)':'var(--green)';
    return`<div class="rank-row" style="display:block">
      <div style="display:flex;align-items:center;gap:10px">
        <div style="width:10px;height:10px;border-radius:50%;background:${x.cat.color};flex-shrink:0"></div>
        <div style="font-size:13px;flex:1">${x.cat.icon} ${x.cat.name}</div>
        <div style="font-size:11px;color:var(--text3);margin-right:8px">${pct}%</div>
        <div style="font-family:DM Mono;font-size:12px;color:${color}">${x.total<0?'−':''}${fmt(x.total)}</div>
      </div>${split}
    </div>`;
  }).join('');

  document.getElementById('ins-cat-detail').innerHTML=`
    ${_viewOwner==='mariana'?'<div class="owner-legend" style="margin-bottom:10px"><span><i class="owner-dot owner-mariana"></i>Mariana</span><span><i class="owner-dot owner-luiza"></i></span></div>':''}
    <div style="padding:6px 0;font-size:11px;color:var(--red);text-transform:uppercase;letter-spacing:.5px">Despesas por categoria</div>
    ${expenseHtml||'<div class="empty">Sem despesas no período</div>'}
    <div style="padding:14px 0 6px;font-size:11px;color:var(--green);text-transform:uppercase;letter-spacing:.5px">Receitas por categoria</div>
    ${incomeHtml||'<div class="empty">Sem receitas no período</div>'}`;

}

// ===================== CATEGORIAS =====================
function renderCategorias(){
  const month=parseInt(document.getElementById('cat-month')?.value??new Date().getMonth());
  const year=parseInt(document.getElementById('cat-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  // month=-1 → todos os meses do ano selecionado
  const txs=month===-1
    ?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year)
    :filterTx(db.transactions,month,year);
  const cats=loadCats();
  function makeBars(type,cid){
    const grossTotal=txs.filter(t=>t.type===type).reduce((s,t)=>s+Math.abs(t.amount),0);
    const data=cats[type].map(c=>{
      const catTxs=txs.filter(t=>t.type===type&&t.category===c.id);
      const total=catTxs.reduce((s,t)=>s+t.amount,0);
      const count=catTxs.length;
      return{cat:c,total,count};
    // Para receita: mostra mesmo se total negativo (tem transações); para despesa: só se >0
    }).filter(x=>type==='income'?x.count>0:x.total>0).sort((a,b)=>b.total-a.total);
    document.getElementById(cid).innerHTML=!data.length?'<div class="empty" style="padding:16px">Sem dados</div>':data.map(x=>{
      const pct=grossTotal>0?Math.round((Math.abs(x.total)/grossTotal)*100):0;
      const amtColor=type==='income'&&x.total<0?'var(--amber)':'inherit';
      const sign=type==='income'&&x.total<0?'−':'';
      const split=_viewOwner==='mariana'?ownerSplitHtml(ownerTotals(txs,type,x.cat.id),true):'';
      return`<div class="cat-row" style="align-items:flex-start"><div class="cat-name-col">${x.cat.icon} ${x.cat.name}</div><div style="flex:1"><div class="cat-bar-wrap"><div class="cat-bar" style="width:${pct}%;background:${x.cat.color}"></div></div>${split}</div><div class="cat-amt" style="color:${amtColor}">${sign}${fmt(x.total)}</div><div class="cat-pct">${pct}%</div></div>`;
    }).join('');
  }
  makeBars('expense','cat-expense-bars');makeBars('income','cat-income-bars');
  if(_viewOwner==='mariana'){
    ['cat-expense-bars','cat-income-bars'].forEach(id=>{
      const el=document.getElementById(id);
      if(el)el.insertAdjacentHTML('afterbegin','<div class="owner-legend" style="padding:0 20px 8px"><span><i class="owner-dot owner-mariana"></i>Mariana</span><span><i class="owner-dot owner-luiza"></i></span></div>');
    });
  }
  const expTotal=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const expData=cats.expense.map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartCatPie=dc(chartCatPie);
  const catLeg=document.getElementById('cat-pie-legend');
  if(!expData.length){catLeg.innerHTML='';return}
  chartCatPie=new Chart(document.getElementById('chart-cat-pie'),{type:'doughnut',data:{labels:expData.map(x=>x.cat.name),datasets:[{data:expData.map(x=>x.total),backgroundColor:expData.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  catLeg.innerHTML=expData.map(x=>{const pct=expTotal>0?Math.round((x.total/expTotal)*100):0;return`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-size:11px;color:var(--text3)">${pct}%</span></span>`}).join('');
  catLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
}

// ===================== ANUAL =====================
function renderAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const allTxs=db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);
  const months=MONTHS.map((name,m)=>{const txs=filterTx(db.transactions,m,year);const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);return{name,income,expense,balance:income-expense,count:txs.length}});
  const totI=months.reduce((s,m)=>s+m.income,0),totE=months.reduce((s,m)=>s+m.expense,0),totB=totI-totE;
  document.getElementById('anual-income').textContent=fmt(totI);document.getElementById('anual-expense').textContent=fmt(totE);document.getElementById('anual-balance').textContent=fmt(totB);document.getElementById('anual-balance').style.color=totB>=0?'var(--green)':'var(--red)';
  chartAnual=dc(chartAnual);
  const annualDatasets=_viewOwner==='mariana'?[
    {label:'Receitas Mariana',data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type==='income'),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:'rgba(37,99,235,.92)',borderColor:'#60a5fa',borderWidth:1,stack:'receitas',borderRadius:3},
    {label:'Receitas ',data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type==='income'),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:'rgba(245,158,11,.92)',borderColor:'#fbbf24',borderWidth:1,stack:'receitas',borderRadius:3},
    {label:'Despesas Mariana',data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type==='expense'),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:'rgba(30,64,175,.78)',borderColor:'#1e3a8a',borderWidth:2,stack:'despesas',borderRadius:3},
    {label:'Despesas ',data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type==='expense'),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:'rgba(180,83,9,.78)',borderColor:'#92400e',borderWidth:2,stack:'despesas',borderRadius:3},
    {type:'line',label:'Saldo do casal',data:months.map(x=>x.balance),borderColor:'rgba(34,211,160,.95)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:.3}
  ]:[
    {label:'Receitas',data:months.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},
    {label:'Despesas',data:months.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4},
    {type:'line',label:'Saldo',data:months.map(x=>x.balance),borderColor:'rgba(96,165,250,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3}
  ];
  chartAnual=new Chart(document.getElementById('chart-anual'),{type:'bar',data:{labels:MONTHS_SHORT,datasets:annualDatasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:_viewOwner==='mariana',labels:{color:'#8b93a8',boxWidth:10}}},scales:{x:{stacked:_viewOwner==='mariana',grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{stacked:_viewOwner==='mariana',grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  document.getElementById('anual-tbody').innerHTML=months.map(m=>{
    const bc=m.balance>=0?'var(--green)':'var(--red)';
    const incColor=m.income<0?'var(--amber)':'var(--green)';
    const hasIncome=allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getFullYear()===year&&new Date(t.date+'T12:00').getMonth()===months.indexOf(m)).length>0;
    return`<tr style="border-bottom:1px solid var(--border);${!m.count?'opacity:.35':''}"><td style="padding:11px 18px;font-weight:500">${m.name}</td><td style="padding:11px 18px;text-align:right;color:${incColor};font-family:DM Mono;font-size:12px">${hasIncome?fmt(m.income):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--red);font-family:DM Mono;font-size:12px">${m.expense>0?fmt(m.expense):'—'}</td><td style="padding:11px 18px;text-align:right;color:${bc};font-family:DM Mono;font-size:12px">${m.count?fmt(m.balance):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--text2)">${m.count||'—'}</td></tr>`;
  }).join('');
  populateCatAnualSelect();renderCatAnual();
}
function populateCatAnualSelect(){
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const sel=document.getElementById('cat-anual-sel');if(!sel)return;
  sel.innerHTML=loadCats()[type].map(c=>`<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
}
function renderCatAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const catId=document.getElementById('cat-anual-sel')?.value;if(!catId)return;
  const db=visibleDB(),cat=getCat(catId);
  const monthData=MONTHS.map((_,m)=>{const txs=filterTx(db.transactions,m,year).filter(t=>t.type===type&&t.category===catId);return{short:MONTHS_SHORT[m],name:MONTHS[m],total:txs.reduce((s,t)=>s+t.amount,0),count:txs.length}});
  const annual=monthData.reduce((s,m)=>s+m.total,0);
  const actv=monthData.filter(m=>m.total>0);
  const avg=actv.length?actv.reduce((s,m)=>s+m.total,0)/actv.length:0;
  chartCatAnual=dc(chartCatAnual);
  const catAnnualDatasets=_viewOwner==='mariana'?[
    {label:`${type==='income'?'Receita':'Despesa'} — Mariana`,data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type===type&&t.category===catId),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:type==='income'?'rgba(37,99,235,.92)':'rgba(30,64,175,.78)',borderColor:type==='income'?'#60a5fa':'#1e3a8a',borderWidth:type==='income'?1:2,stack:'categoria',borderRadius:4},
    {label:`${type==='income'?'Receita':'Despesa'} — `,data:MONTHS_SHORT.map((_,m)=>filterTx(_db.transactions.filter(t=>t.owner==='mariana'&&t.type===type&&t.category===catId),m,year).reduce((s,t)=>s+t.amount,0)),backgroundColor:type==='income'?'rgba(245,158,11,.92)':'rgba(180,83,9,.78)',borderColor:type==='income'?'#fbbf24':'#92400e',borderWidth:type==='income'?1:2,stack:'categoria',borderRadius:4},
    {type:'line',label:'Média do casal',data:monthData.map(()=>avg),borderColor:'rgba(255,255,255,.35)',borderDash:[4,4],borderWidth:1.5,pointRadius:0,backgroundColor:'transparent'}
  ]:[
    {label:cat.name,data:monthData.map(x=>x.total),backgroundColor:monthData.map(x=>x.total>0?cat.color+'bb':'rgba(255,255,255,0.05)'),borderColor:cat.color,borderWidth:1,borderRadius:5},
    {type:'line',label:'Média',data:monthData.map(()=>avg),borderColor:'rgba(255,255,255,0.25)',borderDash:[4,4],borderWidth:1.5,pointRadius:0,backgroundColor:'transparent'}
  ];
  chartCatAnual=new Chart(document.getElementById('chart-cat-anual'),{type:'bar',data:{labels:monthData.map(x=>x.short),datasets:catAnnualDatasets},options:{responsive:true,maintainAspectRatio:false,onClick:(evt,elements)=>{if(elements?.length&&elements[0].datasetIndex<(_viewOwner==='mariana'?2:1))openCategoryHistory(year,elements[0].index,type,catId)},plugins:{legend:{display:_viewOwner==='mariana',labels:{color:'#8b93a8',boxWidth:10}},tooltip:{callbacks:{label:ctx=>` ${fmt(ctx.raw)}`,afterLabel:ctx=>ctx.datasetIndex<(_viewOwner==='mariana'?2:1)?' Clique para abrir o histórico':''}}},scales:{x:{stacked:_viewOwner==='mariana',grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{stacked:_viewOwner==='mariana',grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(1)+'k'}}}}});
  const best=monthData.reduce((a,b)=>b.total>a.total?b:a,monthData[0]);
  document.getElementById('cat-anual-summary').innerHTML=[`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📦 Total: <strong style="color:${cat.color}">${fmt(annual)}</strong></span>`,`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📊 Média: <strong>${fmt(avg)}</strong></span>`,best.total>0?`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🏆 Maior: <strong>${best.name} (${fmt(best.total)})</strong></span>`:'',`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🔢 Lançamentos: <strong>${monthData.reduce((s,m)=>s+m.count,0)}</strong></span>`].join('');
}


function openCategoryHistory(year,month,type,catId){
  const db=visibleDB();
  const cat=getCat(catId);
  const txs=filterTx(db.transactions,month,year)
    .filter(t=>t.type===type&&t.category===catId)
    .sort((a,b)=>new Date(b.date)-new Date(a.date));

  document.getElementById('history-modal-title').textContent=`${cat.icon} ${cat.name}`;
  document.getElementById('history-modal-sub').textContent=`${MONTHS[month]} de ${year}${viewSuffix()}`;
  const total=txs.reduce((s,t)=>s+t.amount,0);
  const owners=ownerTotals(txs,type,catId);
  document.getElementById('history-modal-summary').innerHTML=`
    <span style="background:var(--bg3);padding:6px 10px;border-radius:8px">Total: <strong>${fmt(total)}</strong></span>
    <span style="background:var(--bg3);padding:6px 10px;border-radius:8px">${txs.length} lançamento(s)</span>
    ${_viewOwner==='mariana'?`<span style="background:var(--bg3);padding:6px 10px;border-radius:8px;color:var(--blue)">Mariana: ${fmt(owners.mariana)}</span><span style="background:var(--bg3);padding:6px 10px;border-radius:8px;color:#f59e0b">: ${fmt(owners.luiza)}</span>`:''}`;

  document.getElementById('history-modal-list').innerHTML=txs.length?txs.map(t=>{
    const d=new Date(t.date+'T12:00');
    return`<div class="history-row">
      <div style="font-size:11px;color:var(--text2)">${d.toLocaleDateString('pt-BR')}</div>
      <div>
        <div style="font-size:13px;font-weight:500">${t.description}</div>
        <div style="font-size:10.5px;color:var(--text3)">${profileName(t.owner)}${t.shared?' · compartilhado':''}${t.recurring?' · recorrente':''}</div>
      </div>
      <div style="text-align:right;font-family:'DM Mono',monospace;color:${type==='expense'?'var(--red)':'var(--green)'}">${fmt(t.amount)}</div>
    </div>`;
  }).join(''):'<div class="empty">Nenhum lançamento neste mês.</div>';

  document.getElementById('history-modal-bg').classList.add('open');
}
function closeCategoryHistory(){document.getElementById('history-modal-bg').classList.remove('open')}

// ===================== EXPORT =====================
function renderExportSection(){
  updateExportPreview();
}
function getExportTxs(){
  return [...(_db.transactions||[])].sort((a,b)=>new Date(a.date)-new Date(b.date));
}
function updateExportPreview(){
  const txs=getExportTxs();
  const cnt=document.getElementById('exp-preview-count');if(cnt)cnt.textContent=txs.length+' transação(ões) em toda a base';
  const tb=document.getElementById('exp-preview-tbody');if(!tb)return;
  if(!txs.length){tb.innerHTML='<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text3)">Base vazia</td></tr>';return}
  tb.innerHTML=txs.slice(0,50).map(t=>{const c=getCat(t.category),d=new Date(t.date+'T12:00'),color=t.type==='income'?'var(--green)':'var(--red)',sign=t.type==='income'?'+':'-';return`<tr><td>${d.toLocaleDateString('pt-BR')}</td><td style="color:${color}">${t.type==='income'?'Receita':'Despesa'}</td><td>Mariana</td><td>${c.icon} ${c.name}</td><td style="color:${color};font-family:DM Mono">${sign} ${fmt(Math.abs(t.amount))}</td><td>${t.description}</td></tr>`}).join('')+(txs.length>50?`<tr><td colspan="6" style="text-align:center;padding:8px;color:var(--text3);font-size:12px">... e mais ${txs.length-50} transações</td></tr>`:'');
}
function exportExcel(){
  const txs=getExportTxs();
  if(!txs.length){alert('A base está vazia. Não há dados para exportar.');return}
  const rows=[['Data','Tipo','Usuário','Categoria','Valor','Descrição','Parcelado','Parcela','Grupo de parcelas','Recorrente']];
  txs.forEach(t=>{const c=getCat(t.category);rows.push([t.date,t.type==='income'?'Receita':'Despesa','Mariana',c.name,t.type==='expense'?-Math.abs(t.amount):t.amount,t.description,t.parcelaGroup?'Sim':'Não',t.parcelaGroup?`${t.parcelaNum}/${t.parcelaTotal}`:'',t.parcelaGroup||'',t.recurring?'Sim':'Não'])});
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:12},{wch:11},{wch:12},{wch:22},{wch:14},{wch:36},{wch:10},{wch:10},{wch:24},{wch:10}];
  for(let r=1;r<rows.length;r++){const vc=ws[XLSX.utils.encode_cell({r,c:4})];if(vc){vc.t='n';vc.z='#,##0.00'}}
  XLSX.utils.book_append_sheet(wb,ws,'Transações');

  const years=[...new Set(txs.map(t=>new Date(t.date+'T12:00').getFullYear()))].sort((a,b)=>a-b);
  const summary=[['Ano','Receitas','Despesas','Saldo','Quantidade']];
  years.forEach(year=>{const y=txs.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);const inc=y.filter(t=>t.type==='income').reduce((a,t)=>a+t.amount,0);const exp=y.filter(t=>t.type==='expense').reduce((a,t)=>a+Math.abs(t.amount),0);summary.push([year,inc,-exp,inc-exp,y.length])});
  const ws2=XLSX.utils.aoa_to_sheet(summary);ws2['!cols']=[{wch:10},{wch:16},{wch:16},{wch:16},{wch:12}];XLSX.utils.book_append_sheet(wb,ws2,'Resumo Anual');

  const catRows=[['Tipo','Categoria','Ícone','Padrão']];
  (_cats.expense||[]).forEach(c=>catRows.push(['Despesa',c.name,c.icon,c.builtin?'Sim':'Não']));
  (_cats.income||[]).forEach(c=>catRows.push(['Receita',c.name,c.icon,c.builtin?'Sim':'Não']));
  const ws3=XLSX.utils.aoa_to_sheet(catRows);ws3['!cols']=[{wch:12},{wch:24},{wch:8},{wch:10}];XLSX.utils.book_append_sheet(wb,ws3,'Categorias');

  const stamp=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`FinTrack_Mariana_Completo_${stamp}.xlsx`);
}

// ===================== CAT MANAGER =====================
function renderCatManager(){
  const cats=loadCats();
  const db=loadDB();
  // compute total per category
  const totals={};
  db.transactions.forEach(t=>{totals[t.category]=(totals[t.category]||0)+t.amount});

  ['expense','income'].forEach(type=>{
    const el=document.getElementById('gcat-'+type+'-list');
    if(!el)return;
    el.innerHTML=cats[type].map(c=>{
      const total=totals[c.id]||0;
      const inUse=db.transactions.some(t=>t.category===c.id);
      const totalColor=total<0?'var(--amber)':total>0?'var(--green)':'var(--text3)';
      const totalBadge=inUse?`<span class="cat-total-badge" style="color:${totalColor}">${total<0?'−':''}${fmt(total)}</span>`:'';
      const editBtn=`<button class="cat-item-edit-btn" onclick="toggleCatEdit('${type}','${c.id}')" title="Editar">✏️</button>`;
      const delBtn=`<button class="cat-item-del" onclick="deleteCat('${type}','${c.id}')" title="Excluir">✕</button>`;
      return`<div>
        <div class="cat-item-row" id="crow-${c.id}">
          <div class="cat-color-dot" style="background:${c.color}"></div>
          <div style="font-size:16px">${c.icon}</div>
          <div class="cat-item-name">${c.name}</div>
          ${totalBadge}
          ${c.builtin?'<span class="cat-item-badge">padrão</span>':''}
          ${editBtn}
          ${delBtn}
        </div>
        <div id="cedit-${c.id}" class="cat-edit-inline" style="display:none">
          <input id="cename-${c.id}" class="form-input" value="${c.name}" placeholder="Nome" style="flex:1;min-width:80px;height:32px;padding:4px 8px;font-size:12px">
          <input type="color" id="cecolor-${c.id}" value="${c.color}" class="color-picker" style="width:32px;height:32px">
          <input id="ceicon-${c.id}" class="form-input" value="${c.icon}" placeholder="🏷️" style="width:40px;text-align:center;height:32px;padding:4px;font-size:14px">
          <button class="btn btn-primary" onclick="saveCatEdit('${type}','${c.id}')" style="padding:4px 10px;font-size:12px">✓</button>
          <button class="btn" onclick="toggleCatEdit('${type}','${c.id}')" style="padding:4px 10px;font-size:12px">✕</button>
        </div>
      </div>`;
    }).join('')||'<div style="color:var(--text3);font-size:13px;padding:8px 0">Nenhuma</div>';
  });
}

function toggleCatEdit(type,id){
  const el=document.getElementById('cedit-'+id);
  if(!el)return;
  el.style.display=el.style.display==='none'?'flex':'none';
}

async function saveCatEdit(type,id){
  const name=document.getElementById('cename-'+id)?.value.trim();
  const color=document.getElementById('cecolor-'+id)?.value;
  const icon=document.getElementById('ceicon-'+id)?.value.trim()||'📦';
  if(!name){alert('Nome não pode ser vazio.');return}
  const cats=loadCats();
  const idx=cats[type].findIndex(c=>c.id===id);
  if(idx>=0){cats[type][idx]={...cats[type][idx],name,color,icon};}
  await saveCats(cats);
  renderCatManager();
}
async function addCat(type){
  const nE=document.getElementById(type==='expense'?'new-exp-name':'new-inc-name');
  const cE=document.getElementById(type==='expense'?'new-exp-color':'new-inc-color');
  const iE=document.getElementById(type==='expense'?'new-exp-icon':'new-inc-icon');
  const name=nE.value.trim();if(!name){alert('Digite um nome.');return}
  const cats=loadCats();
  cats[type].push({id:slugify(name)+'_'+Date.now().toString(36),name,icon:iE.value.trim()||(type==='expense'?'📦':'💰'),color:cE.value,builtin:false});
  await saveCats(cats);nE.value='';iE.value='';renderCatManager();
}
async function deleteCat(type,id){
  const cats=loadCats(),cat=cats[type].find(c=>c.id===id);
  const inUse=loadDB().transactions.some(t=>t.category===id);
  if(inUse){
    if(!confirm(`A categoria "${cat?.name}" tem transações registradas.\nAs transações existentes manterão a referência à categoria, mas ela não aparecerá mais nos filtros.\n\nExcluir mesmo assim?`))return;
  } else if(cat?.builtin){
    if(!confirm(`"${cat?.name}" é uma categoria padrão. Deseja excluí-la mesmo assim?`))return;
  }
  cats[type]=cats[type].filter(c=>c.id!==id);
  await saveCats(cats);renderCatManager();
}

// ===================== CLOUD =====================
function renderCloudSection(){
  const user=auth?.currentUser;
  const badge=document.getElementById('sync-status-badge');
  const pending=0;
  if(user&&_lastCloudError){badge.className='sync-status sync-err';badge.innerHTML='<span class="sync-dot err"></span>Firebase pendente · '+pending+' alteração(ões)'}
  else if(user&&_cloudLoaded&&pending===0){badge.className='sync-status sync-ok';badge.innerHTML='<span class="sync-dot ok"></span>Firebase protegido e sincronizado'}
  else if(user&&_cloudLoaded){badge.className='sync-status sync-local';badge.innerHTML='<span class="sync-dot local"></span>'+pending+' alteração(ões) aguardando envio'}
  else{badge.className='sync-status sync-local';badge.innerHTML='<span class="sync-dot local"></span>Aguardando sincronização'}
  const emailEl=document.getElementById('cloud-user-email');if(emailEl)emailEl.textContent=user?.email||'—';
  const profileEl=document.getElementById('cloud-user-profile');if(profileEl)profileEl.textContent=_currentProfile?profileName(_currentProfile):'—';
  const urlEl=document.getElementById('cloud-url-display');if(urlEl)urlEl.textContent=EXPECTED_DATABASE_URL+'fintrack';
}

function firebaseSafeKey(value){
  // Realtime Database não aceita . # $ [ ] / nas chaves.
  // Mantemos o ID original dentro do objeto e usamos uma chave segura no Firebase.
  return String(value ?? '').replace(/[.#$\[\]\/]/g,'_');
}


function txArrayToMap(txs){
  const map={};
  (txs||[]).forEach(t=>{if(t?.id!=null)map[String(t.id)]=t});
  return map;
}
// Campos abaixo existem apenas para controle da nuvem e não fazem parte do dado local.
// Removê-los evita que um eco do Firebase pareça uma edição feita pelo usuário.
function cloudTxToLocal(tx){
  if(!tx||typeof tx!=='object')return tx;
  const {firebaseKey,updatedAt,updatedBy,...local}=tx;
  return local;
}
function txMapToArray(map){
  return Object.values(map||{}).filter(Boolean).map(cloudTxToLocal).sort((a,b)=>new Date(b.date)-new Date(a.date));
}
function stableTxFingerprint(map){
  const normalized=Object.values(map||{}).filter(Boolean).map(cloudTxToLocal)
    .sort((a,b)=>String(a.id).localeCompare(String(b.id)));
  return JSON.stringify(normalized);
}
function localDbFingerprint(db){
  return stableTxFingerprint(txArrayToMap(db?.transactions||[]));
}
function cloudPayload(){return{transactions:txArrayToMap(_db.transactions),categories:JSON.stringify(_cats),metadata:{updated:new Date().toISOString(),updatedBy:auth?.currentUser?.uid||null}}}
async function syncToCloud(showFeedback=false){
  if(!auth?.currentUser||!realtimeDb)throw new Error('Usuário não autenticado.');
  assertAuthorizedMariana(auth.currentUser);
  if(showFeedback)logCloud('⏳ Enviando base atual ao Firebase Mariana...');
  const payload=cloudPayload();
  await soloUserRoot().update({transactions:Object.keys(payload.transactions).length?payload.transactions:null,categories:payload.categories,'metadata/updated':firebase.database.ServerValue.TIMESTAMP,'metadata/updatedBy':auth.currentUser.uid});
  _lastRemoteTxFingerprint=stableTxFingerprint(payload.transactions);_cloudTxMap=payload.transactions;_cloudLoaded=true;_lastCloudError=null;
  if(showFeedback)logCloud('✓ Base confirmada no Firebase às '+new Date().toLocaleTimeString('pt-BR'));
  renderCloudSection();return true;
}
async function migrateLegacyCloudIfNeeded(data){return data||null}
async function ensureSoloUserData(){if(!auth?.currentUser||!realtimeDb)throw new Error('Usuário não autenticado.');assertAuthorizedMariana(auth.currentUser);await soloUserRoot().once('value')}
function attachTransactionsListener(){
  if(_transactionsListener)soloUserRoot().child('transactions').off('value',_transactionsListener);
  _transactionsListener=snap=>{
    const remote=snap.val()||{};const fp=stableTxFingerprint(remote);
    if(fp===_lastRemoteTxFingerprint)return;
    _lastRemoteTxFingerprint=fp;_cloudTxMap=remote;
    _db={transactions:txMapToArray(remote)};
    populateAllSelects();refreshAll();renderCloudSection();
    logCloud('↻ Firebase atualizado às '+new Date().toLocaleTimeString('pt-BR'));
  };
  soloUserRoot().child('transactions').on('value',_transactionsListener,err=>{_lastCloudError=err;logCloud('✗ Falha no acompanhamento em tempo real: '+err.message);renderCloudSection()});
}
async function syncFromCloud(showFeedback=false){
  if(!auth?.currentUser||!realtimeDb)throw new Error('Usuário não autenticado.');
  assertAuthorizedMariana(auth.currentUser);
  if(showFeedback)logCloud('⏳ Baixando Firebase Mariana...');
  const snap=await soloUserRoot().once('value');const data=snap.val()||{};
  const remote=data.transactions||{};
  _db={transactions:txMapToArray(remote)};
  _cloudTxMap=remote;_lastRemoteTxFingerprint=stableTxFingerprint(remote);
  _cats=data.categories!=null?parseCloudValue(data.categories,JSON.parse(JSON.stringify(DEFAULT_CATS))):JSON.parse(JSON.stringify(DEFAULT_CATS));
  _imports=[];_cloudLoaded=true;_lastCloudError=null;
  attachTransactionsListener();populateAllSelects();refreshAll();
  if(showFeedback)logCloud('✓ Firebase carregado às '+new Date().toLocaleTimeString('pt-BR'));
  return true;
}

function logCloud(msg){const el=document.getElementById('cloud-log');if(el)el.innerHTML=msg+'<br>'+el.innerHTML}
async function forceFirebaseSync(){
  try{
    if(!auth?.currentUser){throw new Error('Mariana não autenticado.')}
    await syncToCloud(true);
    await syncFromCloud(true);
    alert('✓ Firebase sincronizado e confirmado.');
  }catch(err){
    alert('✗ Falha na sincronização.\n\n'+err.message);
  }
}
function firebaseSyncDiagnostics(){const u=auth?.currentUser;return{projectId:firebaseApp?.options?.projectId||null,databaseURL:firebaseApp?.options?.databaseURL||null,uid:u?.uid||null,email:u?.email||null,authorized:!!u&&u.uid===MARIANA_AUTH_UID&&normalizeEmail(u.email)===MARIANA_AUTH_EMAIL,root:'fintrack',cloudLoaded:_cloudLoaded,lastCloudError:_lastCloudError?.message||null,transactions:_db.transactions.length}}


// ===================== AUTOTESTES V9 (não alteram dados reais) =====================
function fintrackPersistenceSelfTest(){
  const results=[];const ok=(name,pass,detail='')=>results.push({name,pass:!!pass,detail});
  try{
    const txA={id:'1',description:'A',amount:10,date:'2026-01-01',type:'expense',category:'x'};
    const txB={...txA,description:'B'};
    const map=txArrayToMap([txA]);
    ok('array→map preserva ID',map['1']?.description==='A');
    ok('map→array preserva conteúdo',txMapToArray(map)[0]?.amount===10);
    ok('fingerprint detecta edição',stableTxFingerprint(txArrayToMap([txA]))!==stableTxFingerprint(txArrayToMap([txB])));
    const cloud={...txA,firebaseKey:'1',updatedAt:123,updatedBy:'u'};
    ok('metadados da nuvem não contaminam local',!('updatedAt' in cloudTxToLocal(cloud))&&!('firebaseKey' in cloudTxToLocal(cloud)));
    ok('parseCloudValue aceita JSON string',parseCloudValue('{"a":1}',{}).a===1);
    ok('parseCloudValue aceita objeto',parseCloudValue({a:2},{}).a===2);
    ok('parseCloudValue rejeita JSON inválido',parseCloudValue('{x}',{safe:true}).safe===true);
    ok('firebaseSafeKey remove caracteres proibidos',!/[[\].#$\/]/.test(firebaseSafeKey('a/b.c#d$e[f]')));
    const sorted=txMapToArray({a:{...txA,id:'a',date:'2026-01-01'},b:{...txA,id:'b',date:'2026-02-01'}});
    ok('ordenação por data descendente',sorted[0]?.id==='b');
  }catch(err){ok('execução dos autotestes',false,err.message)}
  const passed=results.filter(r=>r.pass).length;
  return {passed,total:results.length,failed:results.length-passed,results};
}
window.fintrackPersistenceSelfTest=fintrackPersistenceSelfTest;

// ===================== RESET =====================
function renderResetSection(){
  const db=loadDB(),cats=loadCats();const cExp=cats.expense.filter(c=>!c.builtin).length,cInc=cats.income.filter(c=>!c.builtin).length;
  const tc=db.transactions.length,ti=db.transactions.filter(t=>t.type==='income').length,te=db.transactions.filter(t=>t.type==='expense').length;
  const txEl=document.getElementById('reset-tx-count');if(txEl)txEl.innerHTML=tc?`<strong style="color:var(--text)">${tc}</strong> transações — <span style="color:var(--green)">${ti} receitas</span> · <span style="color:var(--red)">${te} despesas</span>`:'Nenhuma transação.';
  const catEl=document.getElementById('reset-cats-count');if(catEl)catEl.innerHTML=(cExp+cInc)?`<strong style="color:var(--text)">${cExp+cInc}</strong> personalizadas`:'Nenhuma personalizada.';
  const cb=document.getElementById('confirm-reset-all');if(cb)cb.checked=false;
}
async function resetTransactions(){
  if(!_db.transactions.length){alert('Não há transações.');return}
  if(!confirm(`Apagar definitivamente ${_db.transactions.length} transações do Firebase?`))return;
  try{await soloUserRoot().child('transactions').remove();_db={transactions:[]};_cloudTxMap={};_lastRemoteTxFingerprint=stableTxFingerprint({});refreshAll();renderResetSection();renderCloudSection();alert('✓ Transações apagadas da nuvem.')}catch(err){alert('Falha ao apagar no Firebase. Nenhum dado local será usado para restaurar a base.\n'+err.message)}
}
async function resetCustomCats(){
  const cats=loadCats(),total=cats.expense.filter(c=>!c.builtin).length+cats.income.filter(c=>!c.builtin).length;if(!total){alert('Nenhuma categoria personalizada.');return}
  if(!confirm(`Remover ${total} categoria(s) personalizada(s) do Firebase?`))return;
  cats.expense=cats.expense.filter(c=>c.builtin);cats.income=cats.income.filter(c=>c.builtin);
  try{await saveCats(cats);renderResetSection();renderCatManager();alert('✓ Categorias removidas da nuvem.')}catch(err){alert('Falha ao salvar no Firebase.\n'+err.message)}
}
async function resetImportHistory(){return true}
async function resetAll(){
  const cb=document.getElementById('confirm-reset-all');if(!cb?.checked){alert('Marque a confirmação.');return}
  if(!confirm('Apagar TODA a base do Firebase? Esta ação é irreversível.'))return;
  try{
    // remove todo o nó /fintrack; em seguida grava apenas metadados mínimos. Nenhuma transação antiga é restaurada.
    await soloUserRoot().remove();
    await soloUserRoot().child('metadata').set({updated:firebase.database.ServerValue.TIMESTAMP,updatedBy:auth.currentUser.uid,reset:true});
    _db={transactions:[]};_cats=JSON.parse(JSON.stringify(DEFAULT_CATS));_imports=[];_cloudTxMap={};_lastRemoteTxFingerprint=stableTxFingerprint({});
    await clearLegacyLocalPersistence();populateAllSelects();refreshAll();renderResetSection();renderCloudSection();alert('✓ Base zerada no Firebase e neste navegador.');
  }catch(err){alert('Falha ao zerar o Firebase.\n'+err.message)}
}

// ===================== REFRESH =====================
function refreshAll(){
  const id=document.querySelector('.section.active')?.id;
  if(id==='sec-dashboard')updateDashboard();
  if(id==='sec-transacoes')renderTxList();
  if(id==='sec-categorias')renderCategorias();
  if(id==='sec-anual')renderAnual();
  if(id==='sec-evolucao')renderEvolucao();
  if(id==='sec-comparativo')renderComparativo();
  if(id==='sec-insights')renderInsights();
  if(id==='sec-patrimonio')renderPatrimonio();
}
function populateAllSelects(){
  const now=new Date(),cy=now.getFullYear(),cm=now.getMonth();
  populateMonthSelect('dash-month',cm);populateYearSelect('dash-year',cy);
  populateMonthSelect('tx-month',cm);populateYearSelect('tx-year',cy);
  // cat-month: always keep "Todos os meses" as first option
  const catMonthSel=document.getElementById('cat-month');
  if(catMonthSel){
    catMonthSel.innerHTML='<option value="-1">Todos os meses</option>'+
      MONTHS.map((m,i)=>`<option value="${i}"${i===cm?' selected':''}>${m}</option>`).join('');
  }
  populateYearSelect('cat-year',cy);
  populateYearSelect('anual-year',cy);
  populateYearSelect('ev-year',cy);populateYearSelect('ins-year',cy);
  populateMonthSelect('cmp-month',cm);populateYearSelect('cmp-year',cy);
}

// ===================== COMPARATIVO =====================
let chartComparativo=null;

function renderComparativo(){
  const mode=document.getElementById('cmp-mode')?.value||'ytd';
  const year=parseInt(document.getElementById('cmp-year')?.value??new Date().getFullYear());
  const prevYear=year-1;
  const cmpMonthSel=document.getElementById('cmp-month');
  if(cmpMonthSel) cmpMonthSel.style.display=mode==='month'?'inline-block':'none';
  const selMonth=mode==='month'?parseInt(cmpMonthSel?.value??new Date().getMonth()):null;
  const db=visibleDB();

  // Define which months to include
  const now=new Date();
  const maxMonth=mode==='ytd'?(year===now.getFullYear()?now.getMonth():11):11;

  function getTxsForPeriod(y,months){
    return db.transactions.filter(t=>{
      const d=new Date(t.date+'T12:00');
      return d.getFullYear()===y&&months.includes(d.getMonth());
    });
  }

  let monthRange=[];
  if(mode==='month') monthRange=[selMonth];
  else if(mode==='ytd') monthRange=Array.from({length:maxMonth+1},(_,i)=>i);
  else monthRange=Array.from({length:12},(_,i)=>i);

  const curTxs=getTxsForPeriod(year,monthRange);
  const prevTxs=getTxsForPeriod(prevYear,monthRange);

  const curInc=curTxs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const curExp=curTxs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const curBal=curInc-curExp;
  const prevInc=prevTxs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const prevExp=prevTxs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const prevBal=prevInc-prevExp;

  // subtitle
  const modeLabel=mode==='ytd'?`Jan–${MONTHS_SHORT[maxMonth]}`:(mode==='month'?MONTHS[selMonth]:'Ano completo');
  document.getElementById('cmp-subtitle').textContent=`${modeLabel} ${year} vs ${modeLabel} ${prevYear}${viewSuffix()}`;

  // update labels
  ['','2','3'].forEach(s=>{
    const cl=document.getElementById('cmp-year-cur-label'+s);
    const pl=document.getElementById('cmp-year-prev-label'+s);
    if(cl)cl.textContent=year;if(pl)pl.textContent=prevYear;
  });

  // fill cards
  function setBadge(id,cur,prev,invertGood){
    const el=document.getElementById(id);if(!el)return;
    if(!prev){el.innerHTML='<span class="cmp-badge cmp-flat">—</span>';return}
    const pct=((cur-prev)/Math.abs(prev)*100);
    const up=cur>prev;
    // for expenses: up is bad; for income/balance: up is good
    const good=invertGood?!up:up;
    const cls=Math.abs(pct)<0.5?'cmp-flat':good?'cmp-up':'cmp-down';
    const arrow=Math.abs(pct)<0.5?'→':up?'↑':'↓';
    el.innerHTML=`<span class="cmp-badge ${cls}">${arrow} ${Math.abs(pct).toFixed(1)}%</span>`;
  }

  document.getElementById('cmp-inc-cur').textContent=fmt(curInc);
  document.getElementById('cmp-inc-prev').textContent=fmt(prevInc);
  setBadge('cmp-inc-badge',curInc,prevInc,false);

  document.getElementById('cmp-exp-cur').textContent=fmt(curExp);
  document.getElementById('cmp-exp-prev').textContent=fmt(prevExp);
  setBadge('cmp-exp-badge',curExp,prevExp,true);

  document.getElementById('cmp-bal-cur').textContent=fmt(curBal);
  document.getElementById('cmp-bal-cur').style.color=curBal>=0?'var(--green)':'var(--red)';
  document.getElementById('cmp-bal-prev').textContent=fmt(prevBal);
  setBadge('cmp-bal-badge',curBal,prevBal,false);

  // legend labels
  document.getElementById('cmp-legend-cur').textContent='Despesa '+year;
  document.getElementById('cmp-legend-prev').textContent='Despesa '+prevYear;
  document.getElementById('cmp-legend-inc-cur').textContent='Receita '+year;
  document.getElementById('cmp-legend-inc-prev').textContent='Receita '+prevYear;

  // Chart: mês a mês dentro do período
  const labels=mode==='month'?[MONTHS[selMonth]]:monthRange.map(m=>MONTHS_SHORT[m]);
  const curExpByMonth=monthRange.map(m=>getTxsForPeriod(year,[m]).filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0));
  const prevExpByMonth=monthRange.map(m=>getTxsForPeriod(prevYear,[m]).filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0));
  const curIncByMonth=monthRange.map(m=>getTxsForPeriod(year,[m]).filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0));
  const prevIncByMonth=monthRange.map(m=>getTxsForPeriod(prevYear,[m]).filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0));

  chartComparativo=dc(chartComparativo);
  chartComparativo=new Chart(document.getElementById('chart-comparativo'),{
    type:'bar',
    data:{
      labels,
      datasets:[
        {label:'Despesa '+year,data:curExpByMonth,backgroundColor:'rgba(245,101,101,0.8)',borderRadius:3},
        {label:'Despesa '+prevYear,data:prevExpByMonth,backgroundColor:'rgba(245,101,101,0.25)',borderRadius:3},
        {type:'line',label:'Receita '+year,data:curIncByMonth,borderColor:'rgba(34,211,160,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3},
        {type:'line',label:'Receita '+prevYear,data:prevIncByMonth,borderColor:'rgba(34,211,160,0.3)',backgroundColor:'transparent',borderWidth:1.5,borderDash:[4,3],pointRadius:2,tension:0.3},
      ]
    },
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},
              y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });

  // Category comparison
  const cats=loadCats();
  const allCatIds=[...cats.expense,...cats.income].map(c=>c.id);
  const catRows=cats.expense.map(c=>{
    const curCatTxs=curTxs.filter(t=>t.type==='expense'&&t.category===c.id);
    const prevCatTxs=prevTxs.filter(t=>t.type==='expense'&&t.category===c.id);
    const cur=curCatTxs.reduce((s,t)=>s+t.amount,0);
    const prev=prevCatTxs.reduce((s,t)=>s+t.amount,0);
    return{cat:c,cur,prev,curOwners:ownerTotals(curCatTxs),prevOwners:ownerTotals(prevCatTxs),diff:cur-prev,pct:prev>0?((cur-prev)/prev*100):null};
  }).filter(x=>x.cur>0||x.prev>0).sort((a,b)=>b.cur-a.cur);

  const maxCat=Math.max(...catRows.map(r=>Math.max(r.cur,r.prev)),1);
  document.getElementById('cmp-cat-list').innerHTML=catRows.map(r=>{
    const pctCur=Math.round(r.cur/maxCat*100);
    const pctPrev=Math.round(r.prev/maxCat*100);
    const diffColor=r.diff>0?'var(--red)':r.diff<0?'var(--green)':'var(--text3)';
    const diffSign=r.diff>0?'+':'';
    return`<div style="padding:10px 20px;border-bottom:1px solid var(--border)">
      <div style="display:flex;justify-content:space-between;margin-bottom:5px">
        <span style="font-size:13px">${r.cat.icon} ${r.cat.name}</span>
        <span style="font-size:11px;color:${diffColor};font-family:DM Mono">${diffSign}${fmt(r.diff)}</span>
      </div>
      <div style="display:flex;gap:6px;align-items:center;margin-bottom:3px">
        <span style="font-size:10px;color:var(--text3);width:28px">${year}</span>
        <div style="flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden"><div style="width:${pctCur}%;height:100%;background:${r.cat.color};border-radius:3px"></div></div>
        <span style="font-size:11px;font-family:DM Mono;color:var(--red);min-width:80px;text-align:right">${r.cur>0?fmt(r.cur):'—'}</span>
      </div>
      ${_viewOwner==='mariana'?ownerSplitHtml(r.curOwners):''}
      <div style="display:flex;gap:6px;align-items:center">
        <span style="font-size:10px;color:var(--text3);width:28px">${prevYear}</span>
        <div style="flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden"><div style="width:${pctPrev}%;height:100%;background:${r.cat.color};border-radius:3px;opacity:.35"></div></div>
        <span style="font-size:11px;font-family:DM Mono;color:var(--text2);min-width:80px;text-align:right">${r.prev>0?fmt(r.prev):'—'}</span>
      </div>
      ${_viewOwner==='mariana'?ownerSplitHtml(r.prevOwners):''}
    </div>`;
  }).join('')||'<div class="empty">Sem dados para comparar</div>';

  // Maiores variações
  const variations=catRows.filter(r=>r.prev>0&&r.cur>0&&r.pct!==null).sort((a,b)=>Math.abs(b.pct)-Math.abs(a.pct)).slice(0,8);
  document.getElementById('cmp-variations').innerHTML=variations.map(r=>{
    const up=r.diff>0;
    const cls=up?'cmp-down':'cmp-up'; // expenses: up is bad
    return`<div class="rank-row">
      <div style="font-size:16px">${r.cat.icon}</div>
      <div style="flex:1;min-width:0"><div style="font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.cat.name}</div><div style="font-size:11px;color:var(--text2)">${fmt(r.prev)} → ${fmt(r.cur)}</div></div>
      <span class="cmp-badge ${cls}">${up?'↑':'↓'} ${Math.abs(r.pct).toFixed(1)}%</span>
    </div>`;
  }).join('')||'<div class="empty">Sem variações para exibir</div>';
}

// ===================== PATRIMÔNIO =====================
let chartPatrimonio=null;
let _patData=null; // cache da projeção

// Categorias excluídas da projeção de receita (por nome normalizado)
const PAT_EXCLUDED_PATTERNS=[
  /rendimento/,/investimento/,/freelance/,/bonus/,/bônus/,
  /projeto/,/receber/,/recebido/,/doacao/,/doação/
];

function isExcludedFromProjection(catName){
  const n=catName.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  return PAT_EXCLUDED_PATTERNS.some(p=>p.test(n));
}

// Consulta taxas oficiais no Banco Central do Brasil.
// Mantém múltiplas tentativas para contornar bloqueios de CORS em hospedagens estáticas.
async function fetchJsonComTimeout(url, timeoutMs=9000){
  const controller = new AbortController();
  const timer = setTimeout(()=>controller.abort(), timeoutMs);
  try{
    const response = await fetch(url,{
      method:'GET',
      cache:'no-store',
      headers:{'Accept':'application/json'},
      signal:controller.signal
    });
    if(!response.ok) throw new Error(`HTTP ${response.status}`);
    return await response.json();
  } finally {
    clearTimeout(timer);
  }
}

function parseValorBCB(valor){
  if(valor===null || valor===undefined || valor==='') return null;
  if(typeof valor==='number') return Number.isFinite(valor)?valor:null;

  let s=String(valor).trim().replace(/\s/g,'');
  if(!s) return null;

  // Quando há ponto e vírgula, o último separador é tratado como decimal.
  if(s.includes('.') && s.includes(',')){
    if(s.lastIndexOf(',') > s.lastIndexOf('.')){
      s=s.replace(/\./g,'').replace(',','.');
    }else{
      s=s.replace(/,/g,'');
    }
  }else if(s.includes(',')){
    // Formato brasileiro: 14,25
    s=s.replace(/\./g,'').replace(',','.');
  }
  // Se houver apenas ponto, preserva como separador decimal: 14.25.
  const n=Number(s);
  return Number.isFinite(n)?n:null;
}

async function fetchBCBSerie(serie){
  const base=`https://api.bcb.gov.br/dados/serie/bcdata.sgs.${serie}/dados/ultimos/12?formato=json`;
  const urls=[
    `${base}&_=${Date.now()}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(base)}`,
    `https://api.allorigins.win/get?url=${encodeURIComponent(base)}`
  ];

  for(const url of urls){
    try{
      const payload=await fetchJsonComTimeout(url);
      let data=payload;
      if(payload && typeof payload.contents==='string'){
        data=JSON.parse(payload.contents);
      }
      if(Array.isArray(data) && data.length){
        for(let i=data.length-1;i>=0;i--){
          const valor=parseValorBCB(data[i]?.valor);
          if(valor!=null) return {
            valor,
            data:data[i]?.data||'',
            serie
          };
        }
      }
    }catch(err){
      console.warn(`Falha ao consultar série BCB ${serie}:`,err);
    }
  }
  return null;
}

async function fetchFocusSelic(){
  // 432: Meta Selic definida pelo Copom (% a.a.).
  // 4390 fica como contingência para instalações antigas que já a utilizavam.
  return (await fetchBCBSerie(432)) || (await fetchBCBSerie(4390));
}

async function fetchFocusIPCA(){
  // Expectativa de IPCA em 12 meses; usa série alternativa como contingência.
  return (await fetchBCBSerie(13522)) || (await fetchBCBSerie(4175));
}


let _selicFocusAnual=[];
let _selicFocusData='';

async function fetchFocusSelicAnual(){
  const filtro=encodeURIComponent("Indicador eq 'Selic'");
  const base=`https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$top=500&$filter=${filtro}&$orderby=Data%20desc&$format=json`;
  const urls=[
    `${base}&_=${Date.now()}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(base)}`,
    `https://api.allorigins.win/get?url=${encodeURIComponent(base)}`
  ];

  for(const url of urls){
    try{
      const payload=await fetchJsonComTimeout(url,12000);
      let obj=payload;
      if(payload && typeof payload.contents==='string') obj=JSON.parse(payload.contents);
      const rows=Array.isArray(obj?.value)?obj.value:(Array.isArray(obj)?obj:[]);
      if(!rows.length) continue;

      const valid=rows
        .map(r=>({
          ano:Number(r.DataReferencia),
          mediana:Number(r.Mediana),
          media:Number(r.Media),
          data:String(r.Data||''),
          baseCalculo:r.baseCalculo
        }))
        .filter(r=>Number.isFinite(r.ano)&&Number.isFinite(r.mediana));

      if(!valid.length) continue;

      valid.sort((a,b)=>String(b.data).localeCompare(String(a.data)));
      const latestDate=valid[0].data;
      const latest=valid.filter(r=>r.data===latestDate);
      const byYear=new Map();

      latest.forEach(r=>{
        if(!byYear.has(r.ano)) byYear.set(r.ano,r);
      });

      const result=[...byYear.values()].sort((a,b)=>a.ano-b.ano);
      if(result.length) return {data:latestDate,projecoes:result};
    }catch(err){
      console.warn('Falha ao consultar projeções anuais Focus:',err);
    }
  }
  return null;
}

function normalizarHorizontePatrimonio(){
  const input=document.getElementById('pat-anos');
  if(!input) return 10;
  let anos=parseInt(input.value,10);
  if(!Number.isFinite(anos)) anos=10;
  anos=Math.max(1,Math.min(10,anos));
  input.value=anos;
  return anos;
}

function taxaSelicParaAno(ano,fallback){
  const exact=_selicFocusAnual.find(x=>x.ano===ano);
  if(exact) return {taxa:exact.mediana,origem:'Focus'};
  const anteriores=_selicFocusAnual.filter(x=>x.ano<ano).sort((a,b)=>b.ano-a.ano);
  if(anteriores.length) return {taxa:anteriores[0].mediana,origem:`Última projeção (${anteriores[0].ano})`};
  const futuras=_selicFocusAnual.filter(x=>x.ano>ano).sort((a,b)=>a.ano-b.ano);
  if(futuras.length) return {taxa:futuras[0].mediana,origem:`Primeira projeção (${futuras[0].ano})`};
  return {taxa:fallback,origem:'Contingência manual'};
}

function renderSelicFocusTable(){
  const tbody=document.getElementById('pat-selic-projection-tbody');
  const meta=document.getElementById('pat-selic-projection-meta');
  if(!tbody||!meta)return;

  const anos=normalizarHorizontePatrimonio();
  const anoInicial=new Date().getFullYear();
  const fallback=parseFloat(document.getElementById('pat-selic')?.value)||0;

  if(!_selicFocusAnual.length){
    meta.textContent='O Banco Central não retornou projeções anuais neste momento.';
    tbody.innerHTML=Array.from({length:anos},(_,i)=>{
      const ano=anoInicial+i;
      return `<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:8px 10px;font-weight:500">${ano}</td>
        <td style="padding:8px 10px;text-align:right;font-family:'DM Mono',monospace;color:var(--amber)">${fallback.toFixed(2)}% a.a.</td>
        <td style="padding:8px 10px;color:var(--text2)">Contingência manual</td>
      </tr>`;
    }).join('');
    return;
  }

  meta.textContent=`Focus atualizado em ${_selicFocusData||'data não informada'} • horizonte exibido: ${anos} ano(s)`;
  tbody.innerHTML=Array.from({length:anos},(_,i)=>{
    const ano=anoInicial+i;
    const info=taxaSelicParaAno(ano,fallback);
    const exata=_selicFocusAnual.some(x=>x.ano===ano);
    const cor=exata?'var(--green)':'var(--amber)';
    const descricao=exata
      ?'Projeção anual publicada pelo Focus'
      :info.origem.startsWith('Última projeção')
        ?`${info.origem} — mantida constante até o fim do horizonte`
        :info.origem;

    return `<tr style="border-bottom:1px solid var(--border)">
      <td style="padding:8px 10px;font-weight:500">${ano}</td>
      <td style="padding:8px 10px;text-align:right;font-family:'DM Mono',monospace;color:${cor}">${info.taxa.toFixed(2)}% a.a.</td>
      <td style="padding:8px 10px;color:var(--text2)">${descricao}</td>
    </tr>`;
  }).join('');
}

let _patRatesLoading=false;

async function atualizarTaxasBCB(forcar=false){
  if(_patRatesLoading) return;
  _patRatesLoading=true;

  const selicBadge=document.getElementById('pat-selic-badge');
  const ipcaBadge=document.getElementById('pat-ipca-badge');
  const btn=document.getElementById('pat-refresh-rates');

  if(selicBadge){
    selicBadge.className='pat-api-badge pat-api-load';
    selicBadge.textContent='⏳ Buscando SELIC...';
  }
  if(ipcaBadge){
    ipcaBadge.className='pat-api-badge pat-api-load';
    ipcaBadge.textContent='⏳ Buscando IPCA...';
  }
  if(btn){
    btn.disabled=true;
    btn.textContent='⏳ Atualizando...';
  }

  try{
    const [selicInfo,ipcaInfo,selicFocus]=await Promise.all([
      fetchFocusSelic(),
      fetchFocusIPCA(),
      fetchFocusSelicAnual()
    ]);

    if(selicFocus?.projecoes?.length){
      _selicFocusAnual=selicFocus.projecoes;
      _selicFocusData=selicFocus.data||'';
    }else{
      _selicFocusAnual=[];
      _selicFocusData='';
    }
    renderSelicFocusTable();

    const selicInput=document.getElementById('pat-selic');
    const ipcaInput=document.getElementById('pat-ipca');

    if(selicInfo){
      if(selicInput && (forcar || !selicInput.value)) selicInput.value=selicInfo.valor.toFixed(2);
      const dataTxt=selicInfo.data?` • ${selicInfo.data}`:'';
      document.getElementById('pat-selic-hint').textContent=`BCB — Meta Selic: ${selicInfo.valor.toFixed(2)}% a.a.${dataTxt}`;
      selicBadge.className='pat-api-badge pat-api-ok';
      selicBadge.textContent=_selicFocusAnual.length
        ?`✓ Focus: ${_selicFocusAnual.length} ano(s)`
        :`✓ SELIC ${selicInfo.valor.toFixed(2)}% a.a.`;
    }else{
      selicBadge.className='pat-api-badge pat-api-err';
      selicBadge.textContent='✗ Não foi possível consultar';
      document.getElementById('pat-selic-hint').textContent='A API do BCB não respondeu. Tente “Atualizar taxas” ou preencha manualmente.';
    }

    if(ipcaInfo){
      if(ipcaInput && (forcar || !ipcaInput.value)) ipcaInput.value=ipcaInfo.valor.toFixed(2);
      const dataTxt=ipcaInfo.data?` • ${ipcaInfo.data}`:'';
      document.getElementById('pat-ipca-hint').textContent=`BCB — IPCA esperado em 12 meses: ${ipcaInfo.valor.toFixed(2)}% a.a.${dataTxt}`;
      ipcaBadge.className='pat-api-badge pat-api-ok';
      ipcaBadge.textContent=`✓ IPCA ${ipcaInfo.valor.toFixed(2)}% a.a.`;
    }else{
      ipcaBadge.className='pat-api-badge pat-api-err';
      ipcaBadge.textContent='✗ Não foi possível consultar';
      document.getElementById('pat-ipca-hint').textContent='A API do BCB não respondeu. Tente “Atualizar taxas” ou preencha manualmente.';
    }
  }catch(err){
    console.error('Erro ao atualizar taxas do BCB:',err);
    if(selicBadge){
      selicBadge.className='pat-api-badge pat-api-err';
      selicBadge.textContent='✗ Erro na consulta';
    }
    if(ipcaBadge){
      ipcaBadge.className='pat-api-badge pat-api-err';
      ipcaBadge.textContent='✗ Erro na consulta';
    }
  }finally{
    renderSelicFocusTable();
    _patRatesLoading=false;
    if(btn){
      btn.disabled=false;
      btn.textContent='↻ Atualizar taxas';
    }
  }
}

async function renderPatrimonio(){
  // Calcula base histórica
  const db=visibleDB();
  const cats=loadCats();
  const now=new Date();
  const y=now.getFullYear(), m=now.getMonth();

  // últimos 12 meses
  const last12Months=[];
  for(let i=11;i>=0;i--){
    let mo=m-i, yr=y;
    if(mo<0){mo+=12;yr--;}
    last12Months.push({m:mo,y:yr});
  }

  // categorias de receita excluídas/incluídas
  const allIncCats=cats.income||[];
  const includedCats=allIncCats.filter(c=>!isExcludedFromProjection(c.name));
  const excludedCats=allIncCats.filter(c=>isExcludedFromProjection(c.name));

  // receita operacional média (só categorias incluídas)
  const includedIds=new Set(includedCats.map(c=>c.id));
  let totalRec=0;
  last12Months.forEach(({m:mo,y:yr})=>{
    const txs=filterTx(db.transactions,mo,yr);
    totalRec+=txs.filter(t=>t.type==='income'&&includedIds.has(t.category)).reduce((s,t)=>s+t.amount,0);
  });
  const recBase=totalRec/12;

  // despesa média (todas)
  let totalDesp=0;
  last12Months.forEach(({m:mo,y:yr})=>{
    const txs=filterTx(db.transactions,mo,yr);
    totalDesp+=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  });
  const despBase=totalDesp/12;

  // Atualiza UI base
  document.getElementById('pat-rec-base').textContent=fmt(recBase);
  document.getElementById('pat-desp-base').textContent=fmt(despBase);
  document.getElementById('pat-rec-cats').textContent=includedCats.map(c=>c.name).join(', ')||'Nenhuma';
  document.getElementById('pat-excluded-cats').innerHTML=excludedCats.map(c=>
    `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:11px;color:var(--text3)">${c.icon} ${c.name}</span>`
  ).join('')||'<span style="font-size:11px;color:var(--text3)">Nenhuma excluída</span>';

  // Busca automaticamente as taxas oficiais do Banco Central.
  await atualizarTaxasBCB(false);

  // guarda base para o cálculo
  _patData={recBase,despBase,includedCats,excludedCats};
}

function calcularPatrimonio(){
  if(!_patData){alert('Aguarde o carregamento dos dados.');return}
  const saldoInicial=parseFloat(document.getElementById('pat-saldo').value)||0;
  const anos=normalizarHorizontePatrimonio();
  const selicAnual=parseFloat(document.getElementById('pat-selic').value)||0;
  const ipcaAnual=parseFloat(document.getElementById('pat-ipca').value)||0;
  const crescRec=parseFloat(document.getElementById('pat-cresc-rec').value)||0;
  const crescDesp=parseFloat(document.getElementById('pat-cresc-desp').value)||0;
  const aporte=parseFloat(document.getElementById('pat-aporte').value)||0;

  if(!saldoInicial){alert('Informe o saldo atual aplicado.');return}
  if(!selicAnual){alert('Informe a taxa SELIC.');return}

  const ipcaMensal=(Math.pow(1+ipcaAnual/100,1/12)-1);
  const anoCalendarioInicial=new Date().getFullYear();

  const {recBase,despBase}=_patData;

  // Projeção mês a mês
  let saldo=saldoInicial;
  let rendimentoAcum=0;
  let deflator=1; // para valor real

  const monthly=[];
  const yearly=[];

  for(let ano=1;ano<=anos;ano++){
    const anoCalendario=anoCalendarioInicial+ano-1;
    const selicAnoInfo=taxaSelicParaAno(anoCalendario,selicAnual);
    const selicAno=selicAnoInfo.taxa;
    const selicMensal=(Math.pow(1+selicAno/100,1/12)-1);
    const recAnual=recBase*Math.pow(1+crescRec/100,ano-1);
    const despAnual=despBase*Math.pow(1+crescDesp/100,ano-1);
    const recMes=recAnual;
    const despMes=despAnual;

    let saldoInicioAno=saldo;
    let rendAnual=0;
    let recTotalAno=0;
    let despTotalAno=0;

    for(let mes=1;mes<=12;mes++){
      const rend=saldo*selicMensal;
      const fluxo=recMes-despMes+aporte;
      saldo=saldo+rend+fluxo;
      rendimentoAcum+=rend;
      rendAnual+=rend;
      recTotalAno+=recMes;
      despTotalAno+=despMes;
      deflator*=(1+ipcaMensal);

      monthly.push({
        ano,mes,anoCalendario,selicAno,selicOrigem:selicAnoInfo.origem,
        saldo,
        saldobReal:saldo/deflator,
        rend,
        rec:recMes,
        desp:despMes,
        fluxo,
        rendimentoAcum
      });
    }

    yearly.push({
      ano,
      anoCalendario,
      selicAno,
      selicOrigem:selicAnoInfo.origem,
      saldoInicio:saldoInicioAno,
      saldoFim:saldo,
      saldoReal:saldo/deflator,
      rendAnual,
      recTotal:recTotalAno,
      despTotal:despTotalAno,
      fluxoTotal:recTotalAno-despTotalAno+(aporte*12),
      rendimentoAcum
    });
  }

  // Marcos
  const milestones=[1000000,2000000,5000000];
  const milestoneResults=milestones.map(target=>{
    const hit=monthly.find(m=>m.saldo>=target);
    return{target,hit};
  });

  // Atualiza cards
  const y1=yearly[0];
  const yMid=yearly[Math.min(4,yearly.length-1)];
  const yLong=yearly[Math.min(9,yearly.length-1)];
  const yFinal=yearly[yearly.length-1];

  document.getElementById('pat-cards').style.display='grid';
  document.getElementById('pat-1y').textContent=fmt(y1.saldoFim);
  document.getElementById('pat-1y-real').textContent='Real: '+fmt(y1.saldoReal);

  const midLabel=`Patrimônio em ${Math.min(5,anos)} ano${anos>=5?'s':''}`;
  document.getElementById('pat-mid-label').textContent=midLabel;
  document.getElementById('pat-mid').textContent=fmt(yMid.saldoFim);
  document.getElementById('pat-mid-real').textContent='Real: '+fmt(yMid.saldoReal);

  const longLabel=`Patrimônio em ${Math.min(10,anos)} ano${anos>=10?'s':''}`;
  document.getElementById('pat-long-label').textContent=longLabel;
  document.getElementById('pat-long').textContent=fmt(yLong.saldoFim);
  document.getElementById('pat-long-real').textContent='Real: '+fmt(yLong.saldoReal);

  document.getElementById('pat-rend-total').textContent=fmt(yFinal.rendimentoAcum);
  const pctGain=((yFinal.rendimentoAcum/saldoInicial)*100).toFixed(1);
  document.getElementById('pat-rend-pct').textContent=`+${pctGain}% sobre o capital inicial`;

  // Marcos
  const mileEl=document.getElementById('pat-milestones');
  const mileList=document.getElementById('pat-milestone-list');
  mileEl.style.display='block';
  const patMonthNames=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  mileList.innerHTML=milestoneResults.map(({target,hit})=>{
    const reached=!!hit;
    const label=`R$ ${(target/1000000).toFixed(0)}M`;
    const when=hit
      ?`${patMonthNames[(hit.mes||1)-1]}/${hit.anoCalendario}`
      :'Além do horizonte';
    return`<span class="pat-milestone ${reached?'reached':''}">
      ${reached?'✅':'⏳'} <strong>${label}</strong>
      <span style="font-size:10px;color:${reached?'inherit':'var(--text3)'}">· ${when}</span>
    </span>`;
  }).join('');

  // Marcadores de R$ 1M, R$ 2M e R$ 5M diretamente no gráfico
  const patMilestonePlugin={
    id:'patMilestonePlugin',
    afterDatasetsDraw(chart){
      const atingidos=milestoneResults.filter(m=>m.hit);
      if(!atingidos.length)return;

      const {ctx,chartArea,scales}=chart;
      if(!chartArea||!scales?.x||!scales?.y)return;

      ctx.save();
      ctx.font='10px DM Sans';
      ctx.textBaseline='middle';

      atingidos.forEach(({target,hit},idx)=>{
        const yearIndex=yearly.findIndex(y=>y.anoCalendario===hit.anoCalendario);
        if(yearIndex<0)return;

        const x=scales.x.getPixelForValue(yearIndex);
        const y=scales.y.getPixelForValue(target);
        if(y<chartArea.top||y>chartArea.bottom)return;

        ctx.setLineDash([4,4]);
        ctx.strokeStyle='rgba(255,255,255,.28)';
        ctx.lineWidth=1;
        ctx.beginPath();
        ctx.moveTo(chartArea.left,y);
        ctx.lineTo(chartArea.right,y);
        ctx.stroke();
        ctx.setLineDash([]);

        const label=`R$ ${(target/1000000).toFixed(0)}M · ${patMonthNames[(hit.mes||1)-1]}/${hit.anoCalendario}`;
        const width=ctx.measureText(label).width+14;
        const boxX=Math.min(Math.max(x-width/2,chartArea.left+3),chartArea.right-width-3);
        const boxY=Math.max(chartArea.top+3,y-23-(idx*2));

        ctx.fillStyle='rgba(15,17,23,.94)';
        ctx.fillRect(boxX,boxY,width,18);
        ctx.strokeStyle='rgba(255,255,255,.16)';
        ctx.strokeRect(boxX,boxY,width,18);
        ctx.fillStyle='#f0f2f8';
        ctx.fillText(label,boxX+7,boxY+9);
      });

      ctx.restore();
    }
  };

  // Gráfico
  document.getElementById('pat-chart-card').style.display='block';
  chartPatrimonio=dc(chartPatrimonio);
  chartPatrimonio=new Chart(document.getElementById('chart-patrimonio'),{
    plugins:[patMilestonePlugin],
    type:'line',
    data:{
      labels:yearly.map(y=>String(y.anoCalendario)),
      datasets:[
        {
          label:'Patrimônio Nominal',
          data:yearly.map(y=>y.saldoFim),
          borderColor:'rgba(34,211,160,0.9)',
          backgroundColor:'rgba(34,211,160,0.08)',
          borderWidth:2.5,pointRadius:3,tension:0.3,fill:true
        },
        {
          label:'Valor Real (−IPCA)',
          data:yearly.map(y=>y.saldoReal),
          borderColor:'rgba(96,165,250,0.9)',
          backgroundColor:'rgba(96,165,250,0.05)',
          borderWidth:2,pointRadius:2,tension:0.3,
          borderDash:[5,3],fill:true
        },
        {
          label:'Rendimento Acumulado',
          data:yearly.map(y=>y.rendimentoAcum),
          borderColor:'rgba(251,191,36,0.8)',
          backgroundColor:'transparent',
          borderWidth:1.5,pointRadius:2,tension:0.3,
          borderDash:[3,3]
        }
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>` ${fmt(ctx.raw)}`}}
      },
      scales:{
        x:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#8b93a8',maxTicksLimit:12}},
        y:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#8b93a8',callback:v=>{
          if(v>=1000000)return'R$'+(v/1000000).toFixed(1)+'M';
          if(v>=1000)return'R$'+(v/1000).toFixed(0)+'k';
          return'R$'+v;
        }}}
      }
    }
  });

  // Tabela — modo anual por padrão
  document.getElementById('pat-table-card').style.display='block';
  document.getElementById('pat-show-monthly').checked=false;

  // guarda dados para toggle mensal
  window._patYearly=yearly;
  window._patMonthly=monthly;

  renderPatTable(false);
}

function renderPatTable(showMonthly){
  const tbody=document.getElementById('pat-tbody');
  if(!tbody)return;

  if(!showMonthly){
    tbody.innerHTML=(window._patYearly||[]).map((y,i)=>{
      const isHighlight=(y.ano===1||y.ano===5||y.ano===10);
      const balColor=y.fluxoTotal>=0?'var(--green)':'var(--red)';
      return`<tr class="${isHighlight?'highlight-year':''}">
        <td>${y.anoCalendario} <span style="color:var(--text3);font-size:10px">(Ano ${y.ano})</span></td>
        <td style="color:var(--green)">${fmt(y.recTotal)}</td>
        <td style="color:var(--red)">${fmt(y.despTotal)}</td>
        <td style="color:${balColor}">${y.fluxoTotal>=0?'+':''}${fmt(y.fluxoTotal)}</td>
        <td style="color:var(--blue)">${y.selicAno.toFixed(2)}%<div style="font-size:9px;color:var(--text3)">${y.selicOrigem}</div></td>
        <td style="color:var(--amber)">${fmt(y.rendAnual)}</td>
        <td style="color:var(--green);font-weight:600">${fmt(y.saldoFim)}</td>
        <td style="color:var(--blue)">${fmt(y.saldoReal)}</td>
      </tr>`;
    }).join('');
  } else {
    tbody.innerHTML=(window._patMonthly||[]).map(m=>{
      const balColor=m.fluxo>=0?'var(--green)':'var(--red)';
      return`<tr>
        <td style="color:var(--text2);font-weight:400">${m.anoCalendario} · ${MONTHS_SHORT[m.mes-1]}</td>
        <td style="color:var(--green)">${fmt(m.rec)}</td>
        <td style="color:var(--red)">${fmt(m.desp)}</td>
        <td style="color:${balColor}">${m.fluxo>=0?'+':''}${fmt(m.fluxo)}</td>
        <td style="color:var(--blue)">${m.selicAno.toFixed(2)}%</td>
        <td style="color:var(--amber)">${fmt(m.rend)}</td>
        <td style="color:var(--green);font-weight:600">${fmt(m.saldo)}</td>
        <td style="color:var(--blue)">${fmt(m.saldobReal)}</td>
      </tr>`;
    }).join('');
  }
}

function toggleMonthlyDetail(){
  const show=document.getElementById('pat-show-monthly')?.checked;
  renderPatTable(show);
}

// ===================== INIT =====================
(function bootFinTrack(){
  // Limpa qualquer persistência local legada antes de autenticar. O Firebase Mariana é a fonte única.
  window._fintrackLocalReady=loadAll().then(()=>populateAllSelects()).catch(err=>{
    console.error('Falha ao carregar cache local:',err);
  });
  try{initFirebase()}catch(err){
    showLogin();
    setLoginError('Falha ao iniciar o Firebase: '+(err?.message||err));
  }
  const form=document.getElementById('firebase-login-form');
  if(form){
    form.addEventListener('submit',loginWithFirebase);
  }
  const pass=document.getElementById('login-password');
  if(pass){
    pass.addEventListener('keydown',ev=>{if(ev.key==='Enter'){ev.preventDefault();loginWithFirebase(ev)}});
  }
  window.addEventListener('error',ev=>{
    if(document.getElementById('lock-screen')?.style.display!=='none'){
      setLoginError('Erro no carregamento do sistema: '+(ev.message||'erro desconhecido'));
    }
  });
  window.addEventListener('online',()=>{logCloud('↻ Conexão restabelecida. Recarregando Firebase...');if(auth?.currentUser)syncFromCloud(false).catch(err=>{_lastCloudError=err;renderCloudSection()})});
  window.addEventListener('offline',()=>{_lastCloudError=new Error('Sem conexão com a internet.');renderCloudSection()});
})();
</script>
</body>
</html>
