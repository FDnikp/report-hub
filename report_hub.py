#!/usr/bin/env python3
"""
REPORT HUB v3.0 - Unified Web UI for Conversion Report Generators
Setup: pip install flask pandas openpyxl numpy
Usage: python report_hub.py -> http://localhost:5000
"""
import os, sys, io, traceback, json, urllib.parse
from datetime import datetime
from flask import (Flask, render_template_string, request,
                   send_file, redirect, url_for, flash, jsonify)

app = Flask(__name__)
app.secret_key = "report-hub-v3-2024"

def _import_balancing():
    import balancing_agent_v6_recent as bal
    return bal

def _import_exception():
    import exception_report_builder_New as exc
    return exc

def _import_report_tool():
    import report_tool as rt
    return rt

class OutputCapture:
    def __init__(self):
        self.buffer = io.StringIO()
        self._old = None
    def __enter__(self):
        self._old = sys.stdout
        sys.stdout = self.buffer
        return self
    def __exit__(self, *a):
        sys.stdout = self._old
    @property
    def text(self):
        return self.buffer.getvalue()

def _clean_path(s):
    s = s.strip()
    if len(s) >= 2:
        if (s[0] == '"' and s[-1] == '"') or (s[0] == "'" and s[-1] == "'"):
            s = s[1:-1]
    return s.strip()

def _encode_path(p):
    return urllib.parse.quote(p, safe='')

DEFAULT_DESC_PATH = r"G:\zSecure1\conv\CONVPROJDOC\Conversions and Startups\DMG\Excel - Exception Template and Description File\Exception Description New.xlsx"

SHARED_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*{margin:0;padding:0;box-sizing:border-box;}
html,body{height:100%;font-family:'Inter',sans-serif;background:#f0f4f8;color:#2d3748;}

@keyframes fadeIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:none;}}
@keyframes slideIn{from{transform:translateX(-250px);}to{transform:none;}}
@keyframes float{0%,100%{transform:translateY(0);}50%{transform:translateY(-18px);}}
@keyframes gradientShift{0%{background-position:0% 50%;}50%{background-position:100% 50%;}100%{background-position:0% 50%;}}
@keyframes bounce{to{transform:translateY(-14px);opacity:.3;}}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}

.sidebar{position:fixed;left:0;top:0;bottom:0;width:250px;
background:linear-gradient(180deg,#0c1222,#162032,#1a2744);
color:#fff;display:flex;flex-direction:column;z-index:100;box-shadow:4px 0 30px rgba(0,0,0,.2);animation:slideIn .4s ease-out;}
.sidebar-brand{padding:24px 20px 18px;border-bottom:1px solid rgba(255,255,255,.06);
background:linear-gradient(135deg,rgba(59,130,246,.08),rgba(147,51,234,.08));}
.sidebar-brand h1{font-size:1.25rem;font-weight:800;}
.sidebar-brand h1 .accent{background:linear-gradient(135deg,#60a5fa,#a78bfa);-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.sidebar-brand p{font-size:.68rem;color:#64748b;margin-top:4px;letter-spacing:1.5px;text-transform:uppercase;font-weight:600;}
.sidebar-nav{flex:1;padding:14px 0;}
.sidebar-nav a{display:flex;align-items:center;gap:12px;padding:12px 20px;color:#94a3b8;
text-decoration:none;font-size:.87rem;font-weight:500;transition:all .25s;border-left:3px solid transparent;margin:1px 0;}
.sidebar-nav a:hover{background:rgba(255,255,255,.04);color:#e2e8f0;}
.sidebar-nav a.active{background:linear-gradient(90deg,rgba(59,130,246,.12),transparent);color:#60a5fa;border-left-color:#3b82f6;font-weight:600;}
.sidebar-nav a .icon{font-size:1.05rem;width:24px;text-align:center;}
.sidebar-footer{padding:14px 20px;border-top:1px solid rgba(255,255,255,.06);font-size:.7rem;color:#475569;}

.main{margin-left:250px;min-height:100vh;display:flex;flex-direction:column;}
.topbar{background:rgba(255,255,255,.88);backdrop-filter:blur(12px);padding:12px 32px;
display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid rgba(226,232,240,.6);position:sticky;top:0;z-index:50;}
.breadcrumb{display:flex;align-items:center;gap:8px;font-size:.82rem;color:#64748b;}
.breadcrumb a{color:#3b82f6;text-decoration:none;font-weight:500;}
.breadcrumb a:hover{text-decoration:underline;}
.breadcrumb .sep{color:#cbd5e1;font-size:.7rem;}
.topbar-right{font-size:.78rem;color:#64748b;font-weight:500;display:flex;align-items:center;gap:6px;}
.live-dot{width:8px;height:8px;border-radius:50%;background:#22c55e;animation:pulse 2s infinite;}
.content{flex:1;padding:28px 32px;animation:fadeIn .5s ease-out;position:relative;z-index:1;}

.bg-decoration{position:fixed;top:0;left:250px;right:0;bottom:0;pointer-events:none;z-index:0;overflow:hidden;}
.bg-circle{position:absolute;border-radius:50%;opacity:.035;background:linear-gradient(135deg,#3b82f6,#8b5cf6);}
.bg-circle:nth-child(1){width:500px;height:500px;top:-120px;right:-80px;animation:float 22s ease-in-out infinite;}
.bg-circle:nth-child(2){width:350px;height:350px;bottom:-100px;left:8%;animation:float 17s ease-in-out infinite reverse;}
.bg-circle:nth-child(3){width:250px;height:250px;top:35%;right:15%;animation:float 20s ease-in-out infinite 3s;}

.hero{background:linear-gradient(135deg,#1e3a5f,#2563eb 40%,#7c3aed);background-size:200% 200%;animation:gradientShift 8s ease infinite;
border-radius:18px;padding:44px 36px;color:#fff;margin-bottom:32px;position:relative;overflow:hidden;}
.hero::before{content:'';position:absolute;top:-50%;right:-15%;width:450px;height:450px;border-radius:50%;background:rgba(255,255,255,.04);}
.hero::after{content:'';position:absolute;bottom:-35%;left:-8%;width:350px;height:350px;border-radius:50%;background:rgba(255,255,255,.03);}
.hero h2{font-size:1.7rem;font-weight:800;margin-bottom:8px;position:relative;z-index:1;}
.hero p{font-size:.95rem;opacity:.85;max-width:580px;line-height:1.65;position:relative;z-index:1;}

.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:22px;}
.card{background:rgba(255,255,255,.92);backdrop-filter:blur(8px);border-radius:16px;overflow:hidden;
box-shadow:0 4px 18px rgba(0,0,0,.05);transition:all .3s cubic-bezier(.4,0,.2,1);
text-decoration:none;color:#2d3748;position:relative;border:1px solid rgba(226,232,240,.5);}
.card:hover{transform:translateY(-7px) scale(1.01);box-shadow:0 18px 45px rgba(59,130,246,.12);border-color:rgba(59,130,246,.2);}
.card-top{padding:26px 22px 20px;color:#fff;position:relative;overflow:hidden;}
.card-top::before{content:'';position:absolute;right:-12px;top:-12px;width:90px;height:90px;border-radius:50%;background:rgba(255,255,255,.1);}
.card-top h2{font-size:1.1rem;font-weight:700;position:relative;z-index:1;}
.card-top .tag{font-size:.72rem;opacity:.8;font-weight:500;letter-spacing:.4px;position:relative;z-index:1;margin-top:3px;display:block;}
.card-body{padding:18px 22px 22px;}
.card-body p{font-size:.84rem;color:#64748b;line-height:1.6;margin-bottom:12px;}
.card-body .chip{display:inline-block;background:#f1f5f9;padding:4px 11px;border-radius:18px;font-size:.73rem;color:#475569;font-weight:600;border:1px solid #e2e8f0;}
.card-arrow{position:absolute;bottom:16px;right:18px;font-size:1.2rem;color:#cbd5e1;transition:all .3s;}
.card:hover .card-arrow{color:#3b82f6;transform:translateX(4px);}
.bg-bal{background:linear-gradient(135deg,#1d4ed8,#3b82f6);}
.bg-exc{background:linear-gradient(135deg,#c2410c,#ea580c);}
.bg-rpt{background:linear-gradient(135deg,#15803d,#16a34a);}

.page-header{margin-bottom:22px;}
.page-header h2{font-size:1.3rem;font-weight:700;color:#1e293b;}
.page-header p{font-size:.86rem;color:#64748b;margin-top:3px;}
.form-section{background:rgba(255,255,255,.92);backdrop-filter:blur(6px);border-radius:14px;
box-shadow:0 2px 14px rgba(0,0,0,.04);margin-bottom:18px;overflow:hidden;border:1px solid rgba(226,232,240,.5);transition:box-shadow .3s;}
.form-section:hover{box-shadow:0 4px 22px rgba(0,0,0,.07);}
.section-header{padding:14px 22px;color:#fff;font-weight:600;font-size:.88rem;letter-spacing:.3px;}
.section-body{padding:20px 22px;}

label{display:block;font-weight:600;font-size:.82rem;color:#374151;margin:13px 0 5px;}
label:first-child{margin-top:0;}
.hint{font-size:.75rem;color:#94a3b8;margin-top:3px;}
input[type=text],select{width:100%;padding:9px 13px;border:1.5px solid #d1d5db;border-radius:9px;font-size:.87rem;font-family:inherit;transition:all .25s;background:#fff;}
input[type=text]:focus,select:focus{outline:none;border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.1);}

.browse-row{display:flex;gap:7px;align-items:stretch;}
.browse-row input{flex:1;}
.browse-btn{padding:9px 16px;background:linear-gradient(135deg,#1e293b,#334155);color:#fff;border:none;border-radius:9px;font-size:.82rem;font-weight:600;cursor:pointer;white-space:nowrap;transition:all .25s;}
.browse-btn:hover{background:linear-gradient(135deg,#334155,#475569);transform:translateY(-1px);box-shadow:0 3px 10px rgba(0,0,0,.12);}
.scan-btn{padding:9px 14px;background:linear-gradient(135deg,#0369a1,#0ea5e9);color:#fff;border:none;border-radius:9px;font-size:.82rem;font-weight:600;cursor:pointer;transition:all .25s;}
.scan-btn:hover{transform:translateY(-1px);box-shadow:0 3px 10px rgba(14,165,233,.3);}

.radio-cards{display:flex;gap:12px;margin-top:8px;}
.radio-card{flex:1;border:2px solid #e2e8f0;border-radius:11px;padding:16px;cursor:pointer;transition:all .25s;text-align:center;background:#fff;}
.radio-card:hover{border-color:#93c5fd;background:#eff6ff;transform:translateY(-2px);}
.radio-card.selected{border-color:#3b82f6;background:linear-gradient(135deg,#eff6ff,#dbeafe);box-shadow:0 4px 14px rgba(59,130,246,.12);}
.radio-card input{display:none;}
.radio-card h4{font-size:.86rem;margin-bottom:3px;color:#1e293b;}
.radio-card p{font-size:.74rem;color:#64748b;}

.toggle-wrap{display:flex;align-items:center;gap:12px;margin:13px 0;}
.toggle{position:relative;width:46px;height:25px;background:#cbd5e1;border-radius:13px;cursor:pointer;transition:background .3s;}
.toggle.on{background:#3b82f6;}
.toggle .knob{position:absolute;top:3px;left:3px;width:19px;height:19px;background:#fff;border-radius:50%;transition:left .3s;box-shadow:0 2px 5px rgba(0,0,0,.18);}
.toggle.on .knob{left:24px;}
.toggle-label{font-size:.86rem;font-weight:500;color:#475569;}

.file-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:7px;margin-top:10px;}
.file-item{display:flex;align-items:center;gap:7px;padding:8px 11px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:7px;font-size:.81rem;transition:all .2s;cursor:pointer;}
.file-item:hover{background:#eff6ff;border-color:#93c5fd;transform:translateX(2px);}
.file-item input[type=checkbox]{width:15px;height:15px;accent-color:#3b82f6;cursor:pointer;}
.file-item label{cursor:pointer;font-weight:500;color:#334155;margin:0;flex:1;font-size:.81rem;}
.ws-name-input{width:130px;padding:5px 8px;border:1px solid #d1d5db;border-radius:5px;font-size:.78rem;}
.file-item-select-all{background:#dbeafe !important;border-color:#60a5fa !important;font-weight:700;}

.event-card{background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:11px;padding:16px;margin:9px 0;position:relative;transition:all .25s;animation:fadeIn .3s ease-out;}
.event-card:hover{border-color:#93c5fd;box-shadow:0 2px 10px rgba(0,0,0,.04);}
.event-card h4{font-size:.84rem;color:#64748b;margin-bottom:9px;font-weight:600;}
.remove-event{position:absolute;top:10px;right:12px;background:#fee2e2;border:1px solid #fca5a5;color:#dc2626;cursor:pointer;font-size:.78rem;font-weight:700;border-radius:5px;padding:3px 9px;transition:all .2s;}
.remove-event:hover{background:#dc2626;color:#fff;}
.add-event-btn{display:flex;align-items:center;justify-content:center;gap:8px;padding:13px;background:#f8fafc;border:2px dashed #cbd5e1;border-radius:11px;color:#64748b;font-size:.86rem;font-weight:600;cursor:pointer;transition:all .25s;margin-top:8px;}
.add-event-btn:hover{border-color:#3b82f6;color:#3b82f6;background:#eff6ff;transform:translateY(-2px);}

.btn{display:inline-flex;align-items:center;gap:8px;padding:11px 26px;border:none;border-radius:11px;color:#fff;font-size:.88rem;font-weight:600;cursor:pointer;transition:all .3s;margin-top:18px;text-decoration:none;letter-spacing:.3px;}
.btn:hover{transform:translateY(-2px);box-shadow:0 7px 22px rgba(0,0,0,.18);}
.btn:active{transform:translateY(0);}
.btn-blue{background:linear-gradient(135deg,#1d4ed8,#3b82f6);}
.btn-orange{background:linear-gradient(135deg,#c2410c,#ea580c);}
.btn-green{background:linear-gradient(135deg,#15803d,#16a34a);}
.btn-gray{background:linear-gradient(135deg,#374151,#4b5563);}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none;}

.console-wrap{margin-top:18px;animation:fadeIn .4s ease-out;}
.console-title{font-size:.82rem;font-weight:600;color:#64748b;margin-bottom:6px;}
.console{background:#0c1222;color:#cbd5e1;font-family:'Cascadia Code','Fira Code',Consolas,monospace;font-size:.77rem;padding:16px;border-radius:11px;max-height:320px;overflow-y:auto;white-space:pre-wrap;line-height:1.55;border:1px solid #1e293b;}

.result-box{margin-top:16px;padding:16px 20px;border-radius:11px;display:flex;align-items:flex-start;gap:10px;animation:fadeIn .4s ease-out;}
.result-ok{background:linear-gradient(135deg,#ecfdf5,#d1fae5);border:1px solid #86efac;}
.result-err{background:linear-gradient(135deg,#fef2f2,#fee2e2);border:1px solid #fca5a5;}
.result-box .icon-big{font-size:1.3rem;}
.result-box .result-content{flex:1;}
.result-box .result-content h4{font-size:.9rem;margin-bottom:6px;color:#1e293b;}
.result-box a.dl-link{display:inline-flex;align-items:center;gap:4px;padding:6px 14px;background:linear-gradient(135deg,#1d4ed8,#2563eb);color:#fff;border-radius:7px;text-decoration:none;font-size:.8rem;font-weight:600;margin:3px 4px 3px 0;transition:all .2s;box-shadow:0 2px 7px rgba(29,78,216,.2);}
.result-box a.dl-link:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(29,78,216,.3);}

.spinner{display:none;margin-top:18px;text-align:center;padding:18px;}
.spinner.active{display:block;}
.spinner .dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin:0 3px;animation:bounce .6s infinite alternate;}
.spinner .dot:nth-child(1){background:#3b82f6;animation-delay:0s;}
.spinner .dot:nth-child(2){background:#f97316;animation-delay:.15s;}
.spinner .dot:nth-child(3){background:#22c55e;animation-delay:.3s;}
.spinner p{margin-top:8px;font-size:.84rem;color:#64748b;font-weight:500;}

.detected-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 12px;border-radius:18px;font-size:.77rem;font-weight:600;margin-top:7px;}
.badge-ok{background:#dcfce7;color:#15803d;border:1px solid #86efac;}
.badge-warn{background:#fef9c3;color:#a16207;border:1px solid #fde047;}
.hidden{display:none;}
.main-footer{padding:16px 32px;text-align:center;color:#94a3b8;font-size:.74rem;border-top:1px solid #e2e8f0;margin-top:auto;background:rgba(255,255,255,.5);}
</style>
"""


SHARED_JS = """
<script>
function browsePath(inputId, mode) {
    var ep = mode === 'file' ? '/api/browse-file' : '/api/browse-folder';
    fetch(ep).then(function(r){return r.json();}).then(function(d){
        if (d.path) { var el=document.getElementById(inputId); el.value=d.path; el.dispatchEvent(new Event('input')); el.dispatchEvent(new Event('change')); }
    }).catch(function(e){console.error(e);});
}
function showLoading(fid) {
    var sp=document.getElementById('loading-spinner'); if(sp) sp.classList.add('active');
    var btn=document.querySelector('#'+fid+' button[type=submit]'); if(btn){btn.disabled=true;btn.textContent='Generating... please wait';}
}
function selectRadio(el) {
    document.querySelectorAll('.radio-card').forEach(function(c){c.classList.remove('selected');});
    el.classList.add('selected'); el.querySelector('input').checked=true;
}
function toggleSection(tid, sid) {
    document.getElementById(tid).classList.toggle('on');
    document.getElementById(sid).classList.toggle('hidden');
}
function toggleSelectAll(name, checked) {
    document.querySelectorAll('input[name="'+name+'"]').forEach(function(cb){cb.checked=checked;});
}
function updateClock(){
    var n=new Date();
    var m=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var s=m[n.getMonth()]+' '+String(n.getDate()).padStart(2,'0')+', '+n.getFullYear()+
        ' | '+String(n.getHours()).padStart(2,'0')+':'+String(n.getMinutes()).padStart(2,'0')+':'+String(n.getSeconds()).padStart(2,'0');
    var el=document.getElementById('live-clock'); if(el) el.textContent=s;
}
setInterval(updateClock,1000);
document.addEventListener('DOMContentLoaded',updateClock);
</script>
"""


def _sidebar(active="home"):
    items = [("home","/","&#127968;","Dashboard"),("balancing","/balancing","&#128202;","Balancing Report"),
             ("exception","/exception","&#128203;","Exception Report"),("report-tool","/report-tool","&#128193;","Report Tool")]
    nav = ""
    for key, href, icon, label in items:
        cls = " active" if key == active else ""
        nav += '<a href="'+href+'" class="'+cls+'"><span class="icon">'+icon+'</span>'+label+'</a>'
    yr = str(datetime.now().year)
    return ('<div class="sidebar"><div class="sidebar-brand"><h1>&#128209; <span class="accent">Report Hub</span></h1>'
        '<p>Conversions Team</p></div><div class="sidebar-nav">'+nav+'</div>'
        '<div class="sidebar-footer"><p>&#128295; AI integration coming soon</p>'
        '<p style="margin-top:4px;">&copy; '+yr+' Report Hub</p></div></div>')

def _topbar(crumbs):
    bc = '<a href="/">Home</a>'
    for label, href in crumbs:
        bc += ' <span class="sep">&#8250;</span> '
        bc += '<a href="'+href+'">'+label+'</a>' if href else '<span style="color:#1e293b;font-weight:600;">'+label+'</span>'
    return '<div class="topbar"><div class="breadcrumb">'+bc+'</div><div class="topbar-right"><span class="live-dot"></span><span id="live-clock"></span></div></div>'

def _page(active, crumbs, body, extra_js=""):
    yr = str(datetime.now().year)
    return ('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">'
        '<title>Report Hub</title>'+SHARED_CSS+'</head><body>'
        +_sidebar(active)+'<div class="main">'+_topbar(crumbs)+
        '<div class="bg-decoration"><div class="bg-circle"></div><div class="bg-circle"></div><div class="bg-circle"></div></div>'
        '<div class="content">'+body+'</div>'
        '<div class="main-footer">Report Hub &mdash; Conversions Team &mdash; '+yr+'</div>'
        '</div>'+SHARED_JS+extra_js+'</body></html>')


@app.route("/")
def home():
    body = ('<div class="hero"><h2>&#128640; Welcome to Report Hub</h2>'
        '<p>Your one-stop platform for generating conversion reports. Select a tool below to get started.</p></div>'
        '<div class="cards">'
        '<a href="/balancing" class="card"><div class="card-top bg-bal"><h2>&#128202; Balancing Report</h2>'
        '<span class="tag">Balance Verification &amp; Analysis</span></div>'
        '<div class="card-body"><p>Summary, BALMSTR, FLAPMSTR, Delinquency, Cycle Codes, Cross-check.</p>'
        '<span class="chip">BALDATA &middot; PRICDATA &middot; FLAPDATA</span></div><span class="card-arrow">&rarr;</span></a>'
        '<a href="/exception" class="card"><div class="card-top bg-exc"><h2>&#128203; Exception Report</h2>'
        '<span class="tag">Exception Analysis &amp; Tracking</span></div>'
        '<div class="card-body"><p>Event-to-Event pivots, severity parsing, description lookups.</p>'
        '<span class="chip">EXCPXLS &middot; EXCPSUM &middot; Description</span></div><span class="card-arrow">&rarr;</span></a>'
        '<a href="/report-tool" class="card"><div class="card-top bg-rpt"><h2>&#128193; Report Tool</h2>'
        '<span class="tag">Raw Data to Excel Converter</span></div>'
        '<div class="card-body"><p>Smart column classification, data profiling, auto-detected delimiters.</p>'
        '<span class="chip">CLOG &middot; CROSS &middot; PASS &middot; FAIL &amp; more</span></div><span class="card-arrow">&rarr;</span></a>'
        '</div>')
    return _page("home", [], body)


@app.route("/balancing")
def balancing_form():
    body = ('<div class="page-header"><h2>&#128202; Balancing Report Generator</h2>'
        '<p>Generate a comprehensive balancing workbook.</p></div>'
        '<form method="POST" action="/balancing/run" id="bal-form" onsubmit="showLoading(&apos;bal-form&apos;)">'
        '<div class="form-section"><div class="section-header bg-bal">Data Folder</div><div class="section-body">'
        '<label>Folder Path</label><div class="browse-row">'
        '<input type="text" id="bal-folder" name="folder_path" placeholder="Select or paste folder path..." required>'
        '<button type="button" class="browse-btn" onclick="browsePath(&apos;bal-folder&apos;,&apos;folder&apos;)">&#128194; Browse</button></div>'
        '<p class="hint">Must contain BALDATA.txt. PRICDATA.txt and FLAPDATA.txt are optional.</p>'
        '<label>Output Filename <span style="font-weight:400;color:#94a3b8;">(optional)</span></label>'
        '<input type="text" name="output_name" placeholder="Leave blank for auto-generated name">'
        '</div></div>'
        '<button type="submit" class="btn btn-blue">&#9654; Generate Balancing Report</button></form>'
        '<div class="spinner" id="loading-spinner"><div class="dot"></div><div class="dot"></div><div class="dot"></div><p>Generating report...</p></div>')
    return _page("balancing", [("Balancing Report", None)], body)

@app.route("/balancing/run", methods=["POST"])
def balancing_run():
    folder = _clean_path(request.form.get("folder_path",""))
    output_name = request.form.get("output_name","").strip() or None
    console_log, result_file, error_msg = "", None, None
    if not folder or not os.path.isdir(folder):
        error_msg = "Folder not found: " + folder
    else:
        try:
            bal = _import_balancing()
            with OutputCapture() as cap:
                result_path = bal.generate_balancing_report(folder, output_name)
            console_log = cap.text
            if result_path and os.path.isfile(result_path):
                result_file = result_path
            else:
                error_msg = "Report generation failed. Ensure BALDATA.txt exists."
        except Exception as e:
            console_log = traceback.format_exc()
            error_msg = str(e)
    rhtml = ""
    if console_log:
        rhtml += '<div class="console-wrap"><div class="console-title">Console Output</div><div class="console">'+console_log+'</div></div>'
    if result_file:
        rhtml += '<div class="result-box result-ok"><span class="icon-big">&#9989;</span><div class="result-content"><h4>Report generated!</h4><a href="/download?path='+_encode_path(result_file)+'" class="dl-link">&#128229; Download '+os.path.basename(result_file)+'</a></div></div>'
    if error_msg:
        rhtml += '<div class="result-box result-err"><span class="icon-big">&#10060;</span><div class="result-content"><h4>Error</h4><p>'+error_msg+'</p></div></div>'
    body = '<div class="page-header"><h2>&#128202; Balancing Report</h2><p>Results</p></div><a href="/balancing" class="btn btn-gray" style="margin-top:0;margin-bottom:14px;">&larr; Run Again</a>'+rhtml
    return _page("balancing", [("Balancing Report","/balancing"),("Results",None)], body)


EXCEPTION_JS = """
<script>
function excScanFolder() {
    var folder = document.getElementById('exc-folder').value.trim();
    if (!folder) { alert('Enter or browse a folder path first.'); return; }
    var det = document.getElementById('detected-file');
    var grid = document.getElementById('exc-file-grid');
    det.innerHTML = '<span style="color:#64748b;font-size:.82rem;">Scanning...</span>';
    grid.innerHTML = '<p style="color:#64748b;">Loading files...</p>';
    fetch('/api/scan-folder?path=' + encodeURIComponent(folder))
        .then(function(r){return r.json();})
        .then(function(data){
            if (data.error) { det.innerHTML='<span class="detected-badge badge-warn">'+data.error+'</span>'; return; }
            det.innerHTML = data.detected ?
                '<span class="detected-badge badge-ok">&#9989; Auto-detected: '+data.detected+'</span>' :
                '<span class="detected-badge badge-warn">&#9888; No EXCPXLS/EXCPRPT1 found</span>';
            grid.innerHTML = '';
            if (data.files && data.files.length > 0) {
                var sa = document.createElement('div'); sa.className='file-item file-item-select-all';
                var saCb = document.createElement('input'); saCb.type='checkbox'; saCb.id='exc-select-all';
                saCb.addEventListener('change',function(){toggleSelectAll('extra_files',this.checked);});
                var saLbl = document.createElement('label'); saLbl.htmlFor='exc-select-all';
                saLbl.textContent='Select All ('+data.files.length+' files)'; saLbl.style.fontWeight='700';
                sa.appendChild(saCb); sa.appendChild(saLbl); grid.appendChild(sa);
                data.files.forEach(function(f){
                    var d=document.createElement('div'); d.className='file-item';
                    var cb=document.createElement('input'); cb.type='checkbox'; cb.name='extra_files'; cb.value=f.name; cb.id='ef-'+f.name;
                    var lb=document.createElement('label'); lb.htmlFor='ef-'+f.name; lb.textContent=f.name+' ('+f.size_kb+' KB)';
                    var ws=document.createElement('input'); ws.type='text'; ws.name='ws_'+f.name; ws.placeholder='Sheet name'; ws.className='ws-name-input';
                    d.appendChild(cb); d.appendChild(lb); d.appendChild(ws); grid.appendChild(d);
                });
            } else { grid.innerHTML='<p style="color:#94a3b8;">No files found.</p>'; }
        }).catch(function(e){ det.innerHTML='<span class="detected-badge badge-warn">Error: '+e+'</span>'; });
}
var eventIdx=1;
function addEvent(){
    eventIdx++;
    var c=document.getElementById('events-container');
    var card=document.createElement('div'); card.className='event-card'; card.id='event-'+eventIdx;
    var rb=document.createElement('button'); rb.type='button'; rb.className='remove-event'; rb.textContent='Remove';
    rb.addEventListener('click',function(){card.remove();});
    var h4=document.createElement('h4'); h4.textContent='Event #'+eventIdx;
    var l1=document.createElement('label'); l1.textContent='Event Name';
    var i1=document.createElement('input'); i1.type='text'; i1.name='event_name_'+eventIdx; i1.placeholder='e.g. Mock 2, LIVE';
    var l2=document.createElement('label'); l2.textContent='Event Folder Path';
    var br=document.createElement('div'); br.className='browse-row';
    var i2=document.createElement('input'); i2.type='text'; i2.id='evt-folder-'+eventIdx; i2.name='event_folder_'+eventIdx; i2.placeholder='Path containing EXCPSUM.txt';
    var bb=document.createElement('button'); bb.type='button'; bb.className='browse-btn'; bb.innerHTML='&#128194; Browse';
    var idx=eventIdx; bb.addEventListener('click',function(){browsePath('evt-folder-'+idx,'folder');});
    br.appendChild(i2); br.appendChild(bb);
    card.appendChild(rb); card.appendChild(h4); card.appendChild(l1); card.appendChild(i1); card.appendChild(l2); card.appendChild(br);
    c.appendChild(card);
    document.getElementById('event_count').value=eventIdx;
}
</script>
"""

@app.route("/exception")
def exception_form():
    desc = DEFAULT_DESC_PATH
    body = ('<div class="page-header"><h2>&#128203; Exception Report Generator</h2>'
        '<p>Build exception analysis with Event-to-Event comparison.</p></div>'
        '<form method="POST" action="/exception/run" id="exc-form" onsubmit="showLoading(&apos;exc-form&apos;)">'
        '<div class="form-section"><div class="section-header bg-exc">Step 1 &mdash; Exception Folder</div><div class="section-body">'
        '<label>Exception Folder Path</label><div class="browse-row">'
        '<input type="text" id="exc-folder" name="folder_path" placeholder="Select or paste folder path..." required>'
        '<button type="button" class="browse-btn" onclick="browsePath(&apos;exc-folder&apos;,&apos;folder&apos;)">&#128194; Browse</button>'
        '<button type="button" class="scan-btn" onclick="excScanFolder()">&#128269; Scan</button></div>'
        '<div id="detected-file" style="margin-top:7px;"></div>'
        '<div class="toggle-wrap" style="margin-top:16px;">'
        '<div class="toggle" id="extra-toggle" onclick="toggleSection(&apos;extra-toggle&apos;,&apos;extra-section&apos;)"><div class="knob"></div></div>'
        '<span class="toggle-label">Load additional exception files?</span></div>'
        '<div id="extra-section" class="hidden">'
        '<p class="hint" style="margin-bottom:7px;">Select files to include. Click Scan first to load the list.</p>'
        '<div class="file-grid" id="exc-file-grid"><p style="color:#94a3b8;font-size:.82rem;font-style:italic;">Click Scan after entering a folder path.</p></div>'
        '</div></div></div>'
        '<div class="form-section"><div class="section-header bg-exc">Step 2 &mdash; Summary Events</div><div class="section-body">'
        '<p class="hint" style="margin-bottom:9px;">Add each conversion event with its EXCPSUM folder.</p>'
        '<input type="hidden" id="event_count" name="event_count" value="1">'
        '<div id="events-container"><div class="event-card" id="event-1"><h4>Event #1</h4>'
        '<label>Event Name</label><input type="text" name="event_name_1" placeholder="e.g. Mock 1" required>'
        '<label>Event Folder Path</label><div class="browse-row">'
        '<input type="text" id="evt-folder-1" name="event_folder_1" placeholder="Path containing EXCPSUM.txt" required>'
        '<button type="button" class="browse-btn" onclick="browsePath(&apos;evt-folder-1&apos;,&apos;folder&apos;)">&#128194; Browse</button>'
        '</div></div></div>'
        '<div class="add-event-btn" onclick="addEvent()">+ Add Another Event</div></div></div>'
        '<div class="form-section"><div class="section-header bg-exc">Step 3 &mdash; Description File</div><div class="section-body">'
        '<label>Exception Description File</label><div class="browse-row">'
        '<input type="text" id="desc-file" name="desc_file" value="'+desc+'">'
        '<button type="button" class="browse-btn" onclick="browsePath(&apos;desc-file&apos;,&apos;file&apos;)">&#128194; Browse</button></div>'
        '<p class="hint">Pre-filled with default. Change or clear to skip.</p></div></div>'
        '<button type="submit" class="btn btn-orange">&#9654; Generate Exception Report</button></form>'
        '<div class="spinner" id="loading-spinner"><div class="dot"></div><div class="dot"></div><div class="dot"></div><p>Generating report...</p></div>')
    return _page("exception", [("Exception Report", None)], body, EXCEPTION_JS)

@app.route("/exception/run", methods=["POST"])
def exception_run():
    folder = _clean_path(request.form.get("folder_path",""))
    console_log, result_file, error_msg = "", None, None
    if not folder or not os.path.isdir(folder):
        error_msg = "Folder not found: " + folder
    else:
        try:
            exc = _import_exception()
            builder = exc.ExceptionReportBuilder()
            builder.folder_path = folder
            df = exc.find_exc(folder)
            if df:
                builder.exception_files = [(df, os.path.basename(df), "All Exceptions", "Y")]
            else:
                builder.exception_files = []
            extra = request.form.getlist("extra_files")
            if extra:
                builder.exception_files = []
                for fname in extra:
                    fpath = os.path.join(folder, fname)
                    if os.path.isfile(fpath):
                        ws = request.form.get("ws_"+fname, "").strip()
                        if not ws:
                            ws = os.path.splitext(fname)[0]
                        for ch in [":", "/", "?", "*", "[", "]", "_"]:
                            ws = ws.replace(ch, " ")
                        ws = ws[:31].strip()
                        builder.exception_files.append((fpath, fname, ws, "Y"))
            if not builder.exception_files:
                error_msg = "No exception files found."
            else:
                ec = int(request.form.get("event_count", 1))
                events = []
                for i in range(1, ec + 1):
                    en = request.form.get("event_name_"+str(i), "").strip()
                    ef = _clean_path(request.form.get("event_folder_"+str(i), ""))
                    if en and ef and os.path.isdir(ef):
                        sf = None
                        for n in ["EXCPSUM.txt","EXCPSUM.zip"]:
                            c = os.path.join(ef, n)
                            if os.path.isfile(c): sf=c; break
                        if not sf:
                            for f in os.listdir(ef):
                                if f.lower().startswith("excpsum"): sf=os.path.join(ef,f); break
                        if sf: events.append((en, ef, sf))
                builder.events = events
                desc = _clean_path(request.form.get("desc_file",""))
                builder.desc_file = desc if desc and os.path.isfile(desc) else ""
                with OutputCapture() as cap:
                    builder.create_directory_sheet()
                    builder.load_exceptions()
                    builder.build_exception_summary()
                    builder.build_event_to_event_summary()
                    builder.load_description()
                    builder.save_output()
                console_log = cap.text
                for nm in ["Exception_Report_Output.xlsx","Exception_Report_Output_NEW.xlsx"]:
                    p = os.path.join(folder, nm)
                    if os.path.isfile(p): result_file=p; break
                if not result_file: error_msg="Output not created."
        except Exception as e:
            console_log = traceback.format_exc()
            error_msg = str(e)
    rhtml = ""
    if console_log:
        rhtml += '<div class="console-wrap"><div class="console-title">Console Output</div><div class="console">'+console_log+'</div></div>'
    if result_file:
        rhtml += '<div class="result-box result-ok"><span class="icon-big">&#9989;</span><div class="result-content"><h4>Exception Report generated!</h4><a href="/download?path='+_encode_path(result_file)+'" class="dl-link">&#128229; Download '+os.path.basename(result_file)+'</a></div></div>'
    if error_msg:
        rhtml += '<div class="result-box result-err"><span class="icon-big">&#10060;</span><div class="result-content"><h4>Error</h4><p>'+error_msg+'</p></div></div>'
    body = '<div class="page-header"><h2>&#128203; Exception Report</h2><p>Results</p></div><a href="/exception" class="btn btn-gray" style="margin-top:0;margin-bottom:14px;">&larr; Run Again</a>'+rhtml
    return _page("exception", [("Exception Report","/exception"),("Results",None)], body)


REPORT_TOOL_JS = """
<script>
function rptScanFolder() {
    var folder = document.getElementById('rpt-folder').value.trim();
    if (!folder) { alert('Enter or browse a folder path first.'); return; }
    var grid = document.getElementById('rpt-all-files-grid');
    var info = document.getElementById('rpt-file-info');
    grid.innerHTML = '<p style="color:#64748b;">Scanning...</p>';
    fetch('/api/scan-report-files?path=' + encodeURIComponent(folder))
        .then(function(r){return r.json();})
        .then(function(data){
            if (data.error) { grid.innerHTML='<p style="color:#ef4444;">'+data.error+'</p>'; return; }
            var files = data.all_files || [];
            var dc = files.filter(function(f){return f.is_default;}).length;
            info.textContent = dc+' default + '+(files.length-dc)+' other = '+files.length+' total';
            grid.innerHTML = '';
            var sa=document.createElement('div'); sa.className='file-item file-item-select-all';
            var saCb=document.createElement('input'); saCb.type='checkbox'; saCb.id='rpt-select-all';
            saCb.addEventListener('change',function(){toggleSelectAll('rpt_files',this.checked);});
            var saLbl=document.createElement('label'); saLbl.htmlFor='rpt-select-all';
            saLbl.textContent='Select All ('+files.length+' files)'; saLbl.style.fontWeight='700';
            sa.appendChild(saCb); sa.appendChild(saLbl); grid.appendChild(sa);
            files.forEach(function(f){
                var d=document.createElement('div'); d.className='file-item';
                if(f.is_default){d.style.background='#f0fdf4';d.style.borderColor='#86efac';}
                var cb=document.createElement('input'); cb.type='checkbox'; cb.name='rpt_files'; cb.value=f.name; cb.id='rf-'+f.name;
                if(f.is_default) cb.checked=true;
                var lb=document.createElement('label'); lb.htmlFor='rf-'+f.name;
                lb.innerHTML=f.name+' <span style="color:#94a3b8;font-size:.73rem;">('+f.size_kb+' KB'+(f.is_default?' &middot; <b style=color:#15803d>default</b>':'')+')</span>';
                d.appendChild(cb); d.appendChild(lb); grid.appendChild(d);
            });
        }).catch(function(e){ grid.innerHTML='<p style="color:#ef4444;">Error: '+e+'</p>'; });
}
</script>
"""

@app.route("/report-tool")
def report_tool_form():
    body = ('<div class="page-header"><h2>&#128193; Report Tool</h2>'
        '<p>Convert raw .txt report files to professionally formatted Excel workbooks.</p></div>'
        '<form method="POST" action="/report-tool/run" id="rpt-form" onsubmit="showLoading(&apos;rpt-form&apos;)">'
        '<div class="form-section"><div class="section-header bg-rpt">Step 1 &mdash; Report Folder</div><div class="section-body">'
        '<label>Folder Path</label><div class="browse-row">'
        '<input type="text" id="rpt-folder" name="folder_path" placeholder="Select or paste folder path..." required>'
        '<button type="button" class="browse-btn" onclick="browsePath(&apos;rpt-folder&apos;,&apos;folder&apos;)">&#128194; Browse</button>'
        '<button type="button" class="scan-btn" onclick="rptScanFolder()">&#128269; Scan</button></div>'
        '<div style="margin-top:16px;"><h4 style="font-size:.87rem;color:#475569;">All Files in Folder '
        '<span id="rpt-file-info" style="font-weight:400;color:#94a3b8;font-size:.79rem;"></span></h4>'
        '<p class="hint">Default files are pre-selected (green). Check more to include them.</p>'
        '<div class="file-grid" id="rpt-all-files-grid"><p style="color:#94a3b8;font-size:.82rem;font-style:italic;">Click Scan after selecting a folder.</p></div>'
        '</div></div></div>'
        '<div class="form-section"><div class="section-header bg-rpt">Step 2 &mdash; Output Mode</div><div class="section-body">'
        '<label>Choose output format</label><div class="radio-cards">'
        '<div class="radio-card selected" onclick="selectRadio(this)"><input type="radio" name="output_mode" value="F" checked>'
        '<h4>&#128196; Separate Files</h4><p>Each report as its own .xlsx</p></div>'
        '<div class="radio-card" onclick="selectRadio(this)"><input type="radio" name="output_mode" value="T">'
        '<h4>&#128209; Single Workbook</h4><p>All reports as tabs in one file</p></div></div></div></div>'
        '<button type="submit" class="btn btn-green">&#9654; Generate Reports</button></form>'
        '<div class="spinner" id="loading-spinner"><div class="dot"></div><div class="dot"></div><div class="dot"></div><p>Generating reports...</p></div>')
    return _page("report-tool", [("Report Tool", None)], body, REPORT_TOOL_JS)

@app.route("/report-tool/run", methods=["POST"])
def report_tool_run():
    folder = _clean_path(request.form.get("folder_path",""))
    mode = request.form.get("output_mode","F").upper()
    console_log, gen_files, error_msg = "", [], None
    if not folder or not os.path.isdir(folder):
        error_msg = "Folder not found: " + folder
    else:
        try:
            rt = _import_report_tool()
            selected_files = request.form.getlist("rpt_files")
            file_selection = selected_files if selected_files else None
            with OutputCapture() as cap:
                results = rt.run_report_tool(folder, output_mode=mode, file_selection=file_selection)
            console_log = cap.text
            if mode == "F":
                for r in results:
                    if r.get("status") == "Success":
                        fp = os.path.join(folder, r["output_name"])
                        if os.path.isfile(fp):
                            gen_files.append({"name": r["output_name"], "path": fp})
            else:
                c = os.path.join(folder, "Report_Output.xlsx")
                if os.path.isfile(c):
                    gen_files.append({"name": "Report_Output.xlsx", "path": c})
            s = os.path.join(folder, "Report_Summary.xlsx")
            if os.path.isfile(s):
                gen_files.append({"name": "Report_Summary.xlsx", "path": s})
            if not gen_files:
                error_msg = "No reports generated."
        except Exception as e:
            console_log = traceback.format_exc()
            error_msg = str(e)
    rhtml = ""
    if console_log:
        rhtml += '<div class="console-wrap"><div class="console-title">Console Output</div><div class="console">'+console_log+'</div></div>'
    if gen_files:
        lnk = ""
        for f in gen_files:
            lnk += '<a href="/download?path='+_encode_path(f["path"])+'" class="dl-link">&#128229; '+f["name"]+'</a>'
        rhtml += '<div class="result-box result-ok"><span class="icon-big">&#9989;</span><div class="result-content"><h4>Reports generated!</h4>'+lnk+'</div></div>'
    if error_msg:
        rhtml += '<div class="result-box result-err"><span class="icon-big">&#10060;</span><div class="result-content"><h4>Error</h4><p>'+error_msg+'</p></div></div>'
    body = '<div class="page-header"><h2>&#128193; Report Tool</h2><p>Results</p></div><a href="/report-tool" class="btn btn-gray" style="margin-top:0;margin-bottom:14px;">&larr; Run Again</a>'+rhtml
    return _page("report-tool", [("Report Tool","/report-tool"),("Results",None)], body)


@app.route("/api/browse-folder")
def api_browse_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); root.focus_force()
        path = filedialog.askdirectory(title="Select Folder")
        root.destroy()
        return jsonify({"path": path or ""})
    except Exception as e:
        return jsonify({"path": "", "error": str(e)})

@app.route("/api/browse-file")
def api_browse_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True); root.focus_force()
        path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        root.destroy()
        return jsonify({"path": path or ""})
    except Exception as e:
        return jsonify({"path": "", "error": str(e)})

@app.route("/api/scan-folder")
def api_scan_folder():
    folder = request.args.get("path","").strip()
    if not folder or not os.path.isdir(folder):
        return jsonify({"error": "Folder not found", "files": [], "detected": None})
    raw = sorted([f for f in os.listdir(folder) if f.lower().endswith((".txt",".zip"))])
    files = []
    for f in raw:
        try:
            sz = round(os.path.getsize(os.path.join(folder, f)) / 1024, 2)
        except:
            sz = 0
        files.append({"name": f, "size_kb": sz})
    detected = None
    for n in ["EXCPXLS.txt","EXCPRPT1.txt","EXCPXLS.zip","EXCPRPT1.zip"]:
        for fi in files:
            if fi["name"].lower() == n.lower():
                detected = fi["name"]; break
        if detected: break
    return jsonify({"files": files, "detected": detected})

@app.route("/api/scan-report-files")
def api_scan_report_files():
    folder = request.args.get("path","").strip()
    if not folder or not os.path.isdir(folder):
        return jsonify({"error": "Folder not found"})
    try:
        rt = _import_report_tool()
        all_files = rt.discover_files(folder)
        selected, _ = rt.select_default_files(all_files)
        sel_names = set(f["name"] for f in selected)
        out = []
        for f in all_files:
            is_def = f["name"] in sel_names
            oname = rt.map_filename(f["name"]) + ".xlsx" if is_def else f["name"].replace(".txt","") + ".xlsx"
            out.append({"name":f["name"],"size_kb":f["size_kb"],"is_default":is_def,"output":oname})
        return jsonify({"all_files":out,"total":len(all_files)})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/download")
def download_file():
    fpath = urllib.parse.unquote(request.args.get("path",""))
    if fpath and os.path.isfile(fpath):
        return send_file(fpath, as_attachment=True, download_name=os.path.basename(fpath))
    flash("File not found: " + fpath, "error")
    return redirect(url_for("home"))

if __name__ == "__main__":
    print()
    print("=" * 60)
    print("  REPORT HUB v3.0")
    print("=" * 60)
    print()
    print("  Open:  http://localhost:5000")
    print()
    print("=" * 60)
    print()
    app.run(debug=True, host="0.0.0.0", port=5000)
