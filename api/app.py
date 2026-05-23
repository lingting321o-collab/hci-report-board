from flask import Flask, Response, jsonify, request

from api.convert import build_actions, build_docx, parse_report

app = Flask(__name__)

INDEX_HTML = r'''<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>HCI 巡检整改看板生成器</title>
  <style>
    :root{--ink:#1f2937;--muted:#6b7280;--line:#d9e2f3;--blue:#1f4e79;--blue-soft:#eef6fc;--red:#c00000;--red-soft:#fce4d6;--amber:#b45f06;--amber-soft:#fff2cc;--surface:#fff;--page:#f6f8fb;--green:#548235;--green-soft:#e2f0d9}
    *{box-sizing:border-box} body{margin:0;min-height:100vh;background:var(--page);color:var(--ink);font-family:"Microsoft YaHei","PingFang SC",system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
    main{width:min(1180px,calc(100vw - 40px));margin:0 auto;padding:32px 0 42px}.topbar{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:22px}h1{margin:0;color:var(--blue);font-size:28px}.sub{margin:8px 0 0;color:var(--muted);font-size:14px}.badge{border:1px solid var(--line);background:var(--surface);color:var(--blue);border-radius:999px;padding:8px 12px;font-size:13px;white-space:nowrap}
    .layout{display:grid;grid-template-columns:minmax(380px,.92fr) minmax(440px,1.08fr);gap:18px}.panel{border:1px solid var(--line);background:var(--surface);border-radius:8px;padding:20px}.panel-title{margin:0 0 12px;color:var(--blue);font-size:17px}.upload{min-height:242px;display:grid;place-items:center;border:2px dashed #b8c7df;border-radius:8px;background:var(--blue-soft);text-align:center}.upload.dragging{border-color:var(--blue);background:#e4f0fa}.upload input{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}.upload-title{margin:0 0 8px;font-size:20px;font-weight:700}.upload-note{margin:0 0 18px;color:var(--muted);font-size:14px}button,.file-button{display:inline-flex;align-items:center;justify-content:center;min-height:40px;border:0;border-radius:7px;background:var(--blue);color:#fff;padding:0 16px;font-weight:700;cursor:pointer;font-size:14px}button:disabled{cursor:not-allowed;opacity:.5}.file-name{margin-top:14px;color:var(--blue);font-size:13px;word-break:break-all}.actions{display:flex;align-items:center;gap:12px;margin-top:16px}.status{color:var(--muted);font-size:13px}.error{color:var(--red)}
    .progress-box{margin-top:14px;border:1px solid var(--line);border-radius:8px;padding:12px;background:#fff}.progress-meta{display:flex;justify-content:space-between;gap:12px;color:var(--muted);font-size:13px;margin-bottom:8px}.bar{height:10px;background:#edf2f7;border-radius:999px;overflow:hidden}.bar span{display:block;height:100%;width:0;background:var(--blue);transition:width .15s ease}.speed{font-variant-numeric:tabular-nums}.hint{margin:12px 0 0;color:var(--muted);font-size:13px;line-height:1.55}
    .preview-window{min-height:520px;border:1px solid var(--line);border-radius:8px;background:#fbfdff;overflow:hidden}.preview-head{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;background:#eef6fc;border-bottom:1px solid var(--line)}.preview-head strong{color:var(--blue)}.preview-body{padding:16px}.empty{display:grid;place-items:center;min-height:420px;text-align:center;color:var(--muted)}.empty b{display:block;color:var(--ink);margin-bottom:8px}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:14px}.metric{border:1px solid var(--line);border-radius:8px;padding:12px;background:#fff}.metric span{display:block;color:var(--muted);font-size:12px}.metric strong{display:block;font-size:21px;color:var(--blue);margin-top:4px}.metric.critical strong{color:var(--red)}.section-label{font-size:13px;color:var(--muted);margin:14px 0 8px}.action-list{display:grid;gap:8px}.action{border:1px solid var(--line);border-left:5px solid var(--amber);border-radius:8px;background:#fff;padding:10px 12px}.action.p0{border-left-color:var(--red);background:#fffafa}.action-top{display:flex;gap:8px;align-items:center;margin-bottom:5px}.pill{font-size:12px;font-weight:700;border-radius:999px;padding:3px 7px;background:var(--amber-soft);color:var(--amber)}.p0 .pill{background:var(--red-soft);color:var(--red)}.action b{font-size:14px}.action p{margin:0;color:var(--muted);font-size:13px;line-height:1.5}.small{color:var(--muted);font-size:12px}.ok{color:var(--green)}
    @media(max-width:900px){main{width:min(100vw - 24px,1180px);padding-top:22px}.topbar,.layout{display:block}.badge{display:inline-block;margin-top:14px}.panel{margin-bottom:14px;padding:16px}.cards{grid-template-columns:repeat(2,minmax(0,1fr))}}
  </style>
</head>
<body>
<main>
  <div class="topbar"><div><h1>HCI 巡检整改看板生成器</h1><p class="sub">上传深信服原始巡检报告，先预览识别结果，再生成客户展示版 Word 看板。</p></div><div class="badge">输出格式：.docx</div></div>
  <section class="layout">
    <div class="panel">
      <h2 class="panel-title">上传与生成</h2>
      <form id="form">
        <label class="upload" id="drop"><input id="file" type="file" accept=".docx"/><span><p class="upload-title">拖入巡检报告，或选择 Word 文件</p><p class="upload-note">选择后会自动上传并生成预览</p><span class="file-button">选择文件</span><div class="file-name" id="fileName"></div></span></label>
        <div class="progress-box"><div class="progress-meta"><span id="progressLabel">等待文件</span><span><span id="percent">0%</span> · <span class="speed" id="speed">0 KB/s</span></span></div><div class="bar"><span id="bar"></span></div></div>
        <div class="actions"><button id="submit" type="submit" disabled>生成整改看板</button><span class="status" id="status">请先上传并确认预览</span></div>
      </form>
      <p class="hint">说明：预览会读取报告中的得分、异常/告警、备份覆盖率和核心整改项；最终下载的 Word 仍会按整改看板样式生成。</p>
    </div>
    <aside class="panel">
      <h2 class="panel-title">预览窗口</h2>
      <div class="preview-window"><div class="preview-head"><strong>整改看板预览</strong><span class="small" id="previewState">未上传</span></div><div class="preview-body" id="preview"><div class="empty"><div><b>这里会显示识别结果</b><span>选择巡检报告后，系统会先生成预览。</span></div></div></div></div>
    </aside>
  </section>
</main>
<script>
const form=document.querySelector('#form'),fileInput=document.querySelector('#file'),fileName=document.querySelector('#fileName'),submit=document.querySelector('#submit'),statusEl=document.querySelector('#status'),drop=document.querySelector('#drop'),bar=document.querySelector('#bar'),percentEl=document.querySelector('#percent'),speedEl=document.querySelector('#speed'),progressLabel=document.querySelector('#progressLabel'),preview=document.querySelector('#preview'),previewState=document.querySelector('#previewState');
let selectedFile=null, previewReady=false;
function fmtSpeed(bytesPerSecond){if(!isFinite(bytesPerSecond)||bytesPerSecond<=0)return '0 KB/s'; if(bytesPerSecond>1024*1024)return (bytesPerSecond/1024/1024).toFixed(2)+' MB/s'; return Math.max(1,Math.round(bytesPerSecond/1024))+' KB/s'}
function resetProgress(label){bar.style.width='0%';percentEl.textContent='0%';speedEl.textContent='0 KB/s';progressLabel.textContent=label}
function setProgress(evt,start,last){if(!evt.lengthComputable)return last;const now=Date.now(),pct=Math.round((evt.loaded/evt.total)*100);bar.style.width=pct+'%';percentEl.textContent=pct+'%';const elapsed=(now-start)/1000;speedEl.textContent=fmtSpeed(evt.loaded/Math.max(elapsed,.1));return {time:now,loaded:evt.loaded}}
function setFile(file){if(!file)return;if(!file.name.toLowerCase().endsWith('.docx')){statusEl.textContent='请选择 .docx 文件';statusEl.className='status error';submit.disabled=true;return}selectedFile=file;previewReady=false;submit.disabled=true;fileName.textContent=file.name+' · '+(file.size/1024/1024).toFixed(2)+' MB';statusEl.textContent='正在上传并解析预览...';statusEl.className='status';previewState.textContent='解析中';resetProgress('预览上传');loadPreview(file)}
fileInput.addEventListener('change',()=>setFile(fileInput.files[0]));['dragenter','dragover'].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.add('dragging')}));['dragleave','drop'].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.remove('dragging')}));drop.addEventListener('drop',e=>{const file=e.dataTransfer.files[0];if(!file)return;const transfer=new DataTransfer();transfer.items.add(file);fileInput.files=transfer.files;setFile(file)});
function upload(url,file,onDone){const xhr=new XMLHttpRequest();const data=new FormData();data.append('file',file);let start=Date.now(),last={time:start,loaded:0};xhr.upload.onprogress=e=>{last=setProgress(e,start,last)};xhr.onload=()=>{if(xhr.status>=200&&xhr.status<300)onDone(null,xhr);else{let msg='请求失败';try{msg=JSON.parse(xhr.responseText).error||msg}catch(e){}onDone(new Error(msg),xhr)}};xhr.onerror=()=>onDone(new Error('网络连接失败'),xhr);xhr.open('POST',url);xhr.send(data)}
function loadPreview(file){preview.innerHTML='<div class="empty"><div><b>正在解析预览...</b><span>文件上传中，请稍等。</span></div></div>';upload('/api/preview',file,(err,xhr)=>{if(err){preview.innerHTML='<div class="empty"><div><b class="error">预览失败</b><span>'+err.message+'</span></div></div>';statusEl.textContent=err.message;statusEl.className='status error';previewState.textContent='失败';return}const data=JSON.parse(xhr.responseText);renderPreview(data);previewReady=true;submit.disabled=false;statusEl.textContent='预览已生成，可以下载整改看板';statusEl.className='status ok';previewState.textContent='已就绪';progressLabel.textContent='预览完成'})}
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function renderPreview(data){const actions=data.actions||[];preview.innerHTML='<div class="cards"><div class="metric"><span>巡检得分</span><strong>'+esc(data.score)+'</strong></div><div class="metric critical"><span>异常项</span><strong>'+esc(data.critical)+'</strong></div><div class="metric"><span>告警项</span><strong>'+esc(data.warning)+'</strong></div><div class="metric critical"><span>备份覆盖率</span><strong>'+esc(data.backup_rate)+'</strong></div></div><div class="small">设备：HCI-'+esc(data.device_id)+'（'+esc(data.ip)+'） · 巡检日期：'+esc(data.date)+' · 风险项：'+esc(data.risk_count)+' 条</div><div class="section-label">核心整改项</div><div class="action-list">'+actions.map(a=>'<div class="action '+(a.level==='P0'?'p0':'')+'"><div class="action-top"><span class="pill">'+esc(a.level)+'</span><b>'+esc(a.name)+'</b></div><p>'+esc(a.what)+'</p><p><b>动作：</b>'+esc(a.do)+'</p></div>').join('')+'</div>'}
form.addEventListener('submit',e=>{e.preventDefault();if(!selectedFile||!previewReady)return;submit.disabled=true;statusEl.textContent='正在上传并生成 Word...';statusEl.className='status';resetProgress('生成上传');upload('/api/convert',selectedFile,(err,xhr)=>{submit.disabled=false;if(err){statusEl.textContent=err.message;statusEl.className='status error';return}const blob=xhr.response;const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='HCI整改看板.docx';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);statusEl.textContent='已生成并开始下载';statusEl.className='status ok';progressLabel.textContent='生成完成'});const oldOpen=XMLHttpRequest.prototype.open});
const originalUpload=upload;
function upload(url,file,onDone){const xhr=new XMLHttpRequest();const data=new FormData();data.append('file',file);let start=Date.now(),last={time:start,loaded:0};xhr.upload.onprogress=e=>{last=setProgress(e,start,last)};xhr.onload=()=>{if(xhr.status>=200&&xhr.status<300)onDone(null,xhr);else{let msg='请求失败';try{msg=JSON.parse(xhr.responseText).error||msg}catch(e){}onDone(new Error(msg),xhr)}};xhr.onerror=()=>onDone(new Error('网络连接失败'),xhr);xhr.open('POST',url);if(url==='/api/convert')xhr.responseType='blob';xhr.send(data)}
</script>
</body></html>'''


@app.get("/")
def home():
    return Response(INDEX_HTML, content_type="text/html; charset=utf-8")


@app.post("/api/preview")
def preview_report():
    uploaded = request.files.get("file")
    if uploaded is None:
        return jsonify({"error": "没有找到上传文件"}), 400
    data = uploaded.read()
    if not data:
        return jsonify({"error": "上传文件为空"}), 400
    try:
        report = parse_report(data)
        actions = build_actions(report)
        return jsonify({
            "device_id": report["device_id"],
            "ip": report["ip"],
            "date": report["date"],
            "score": report["score"],
            "critical": report["critical"],
            "warning": report["warning"],
            "backup_rate": report["backup_rate"],
            "risk_count": len(report.get("risks", [])),
            "actions": [
                {"level": item[0], "name": item[1], "what": item[2], "do": item[4], "deadline": item[5]}
                for item in actions
            ],
        })
    except Exception as exc:
        return jsonify({"error": f"预览失败：{exc}"}), 400


@app.post("/api/convert")
def convert_report():
    uploaded = request.files.get("file")
    if uploaded is None:
        return jsonify({"error": "没有找到上传文件"}), 400
    data = uploaded.read()
    if not data:
        return jsonify({"error": "上传文件为空"}), 400
    try:
        output = build_docx(parse_report(data))
    except Exception as exc:
        return jsonify({"error": f"生成失败：{exc}"}), 400
    return Response(
        output,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="HCI-remediation-board.docx"'},
    )


@app.route("/api/<path:_path>", methods=["OPTIONS"])
def api_options(_path):
    return Response(status=204, headers={
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
    })
