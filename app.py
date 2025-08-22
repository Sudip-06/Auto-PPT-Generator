import io, os, re, json, tempfile, logging, secrets, itertools
from typing import List, Dict, Any, Optional
import requests
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse
from starlette.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from markdown import markdown
from bs4 import BeautifulSoup

# ---------- Logging (no secrets) ----------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger("auto-ppt")
logger.propagate = False

app = FastAPI(title="Auto PPT Generator", version="1.3")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=False, allow_methods=["*"], allow_headers=["*"]
)

# ============================
# Helpers: Text -> Slide JSON
# ============================

LLM_SYSTEM_INSTRUCTIONS = (
    "You convert long text or markdown into a clean slide deck structure.\n"
    "Return STRICT JSON with this schema:\n"
    "{ \"slides\": [ {\"title\": \"...\", \"bullets\": [\"...\"], \"notes\": \"(optional)\"} ] }\n"
    "Guidelines:\n"
    "- Prefer 6–10 slides for a typical 700–1200 word input; scale reasonably.\n"
    "- Titles concise; bullets short, no numbering, no markdown formatting.\n"
    "- Preserve input section order if clear. Keep notes brief (optional)."
)

def _safe_json(s: str) -> Dict[str, Any]:
    try:
        return json.loads(s)
    except Exception as e:
        raise ValueError(f"LLM did not return valid JSON: {e}")

def _bulletize(paragraph: str) -> List[str]:
    parts = re.split(r"[\n\r]+|(?<=[.?!])\s+", paragraph)
    bullets, seen = [], set()
    for l in parts:
        l = l.strip(" •-–—\t ")
        if not l:
            continue
        if len(l) > 180:
            for p in re.split(r"[;:—–-]|, ", l):
                p = p.strip()
                if p and len(p) >= 3:
                    k = p.lower()
                    if k not in seen:
                        seen.add(k); bullets.append(p[:180])
        else:
            k = l.lower()
            if k not in seen and len(l) > 2:
                seen.add(k); bullets.append(l)
    return bullets[:10]

def _title_from_text(text: str) -> str:
    words = re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]+", text)
    return (" ".join(words[:6]).title()[:60]) if words else "Slide"

def _default_chunker(raw_text: str) -> Dict[str, Any]:
    """Rule-based fallback if no LLM key/model is provided."""
    text = raw_text.strip()
    html = markdown(text)
    soup = BeautifulSoup(html, "html.parser")
    headings = soup.find_all(re.compile("^h[1-3]$"))
    slides = []

    if headings:
        for h in headings:
            title = h.get_text(" ", strip=True) or "Slide"
            bullets: List[str] = []
            for sib in h.find_all_next():
                if sib.name and re.match(r"h[1-3]", sib.name) and sib != h:
                    break
                if sib.name == "p":
                    txt = sib.get_text(" ", strip=True)
                    if txt: bullets.extend(_bulletize(txt))
                elif sib.name in ("ul", "ol"):
                    for li in sib.find_all("li"):
                        li_txt = li.get_text(" ", strip=True)
                        if li_txt: bullets.append(li_txt)
            slides.append({"title": title, "bullets": bullets[:8]})
    else:
        paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()] or [text]
        chunk_size = max(1, len(paras) // 8)
        for i in range(0, len(paras), chunk_size):
            chunk = " ".join(paras[i:i+chunk_size])
            slides.append({"title": _title_from_text(chunk), "bullets": _bulletize(chunk)[:8]})

    if not slides:
        slides = [{"title": "Overview", "bullets": _bulletize(text)[:8]}]
    return {"slides": slides[:25]}

# =====================
# Provider LLM adapters
# =====================

def _env_key_for(provider: str) -> str:
    if provider == "gemini":
        return os.getenv("GEMINI_API_KEY", "")
    if provider == "openai":
        return os.getenv("OPENAI_API_KEY", "")
    if provider in ("gorq", "groq"):
        return os.getenv("GORQ_API_KEY", "")
    return ""

def call_llm_struct(
    provider: str,
    model: str,
    user_text: str,
    guidance: str = "",
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    provider = (provider or "").lower().strip() or "gemini"  # default: Gemini
    model = (model or "").strip()
    api_key = (api_key or "").strip()

    # Enforce: must have an API key in the request OR in env; otherwise fail early to UI
    effective_key = api_key or _env_key_for(provider)
    if not effective_key:
        raise HTTPException(status_code=400, detail="API key required. Please enter your API key in the UI.")

    if not model:
        # No model -> use provider defaults
        if provider == "gemini":
            model = "gemini-2.5-flash"
        elif provider == "openai":
            model = "gpt-4o"
        else:
            model = "deepseek-r1"

    try:
        if provider == "gemini":
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={effective_key}"
            body = {
                "contents": [
                    {"role": "user", "parts": [{"text": LLM_SYSTEM_INSTRUCTIONS}]},
                    {"role": "user", "parts": [{"text": f"Guidance: {guidance or 'none'}\n\nInput:\n{user_text}"}]}
                ],
                "generationConfig": {"temperature": 0.2}
            }
            r = requests.post(url, json=body, timeout=60)
            r.raise_for_status()
            data = r.json()
            text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
            return _safe_json(text)

        elif provider == "openai":
            base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
            url = f"{base_url}/chat/completions"
            headers = {"Authorization": f"Bearer {effective_key}"}
            payload = {
                "model": model,
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
                "messages": [
                    {"role": "system", "content": LLM_SYSTEM_INSTRUCTIONS},
                    {"role": "user", "content": f"Guidance: {guidance or 'none'}\n\nInput:\n{user_text}"}
                ]
            }
            r = requests.post(url, headers=headers, json=payload, timeout=60)
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return _safe_json(content)

        elif provider in ("gorq", "groq"):  # Groq (OpenAI-compatible)
            base_url = os.getenv("GORQ_BASE_URL", "https://api.groq.com/openai/v1")
            url = f"{base_url}/chat/completions"
            headers = {"Authorization": f"Bearer {effective_key}"}
            payload = {
                "model": model,  # e.g., "deepseek-r1" or "openai120b"
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
                "messages": [
                    {"role": "system", "content": LLM_SYSTEM_INSTRUCTIONS},
                    {"role": "user", "content": f"Guidance: {guidance or 'none'}\n\nInput:\n{user_text}"}
                ]
            }
            r = requests.post(url, headers=headers, json=payload, timeout=60)
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return _safe_json(content)

        else:
            return _default_chunker(user_text)

    except HTTPException:
        raise
    except Exception as e:
        logger.warning(f"LLM error -> fallback: {e}")
        return _default_chunker(user_text)

# ==============================
# PPTX building from a template
# ==============================

def extract_template_images(prs: Presentation):
    imgs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    imgs.append(shape.image.blob)
                except Exception:
                    pass
    return imgs

def choose_layout(prs: Presentation):
    layout = None
    for l in prs.slide_layouts:
        name = (l.name or "").lower()
        if "title and content" in name:
            layout = l; break
    if layout is None:
        for l in prs.slide_layouts:
            try:
                if any(ph.placeholder_format.type for ph in l.placeholders):
                    layout = l; break
            except Exception:
                continue
    return layout or prs.slide_layouts[0]

def add_slide(prs: Presentation, layout, title: str, bullets: List[str], note: Optional[str], img_bytes: Optional[bytes]):
    slide = prs.slides.add_slide(layout)
    # Title
    title_shape = slide.shapes.title if slide.shapes.title else None
    if not title_shape:
        for ph in slide.placeholders:
            try:
                if ph.placeholder_format.type and "TITLE" in str(ph.placeholder_format.type):
                    title_shape = ph; break
            except Exception:
                continue
    if title_shape:
        title_shape.text_frame.clear()
        title_shape.text_frame.text = title[:150]

    # Body
    body = None
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.type and "BODY" in str(ph.placeholder_format.type):
                body = ph; break
        except Exception:
            continue
    if body:
        tf = body.text_frame
        tf.clear(); tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for i, b in enumerate((bullets or [])[:8]):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = b; p.level = 0; p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

    # Image
    if img_bytes:
        pic_placeholder = None
        for ph in slide.placeholders:
            try:
                if "PICTURE" in str(ph.placeholder_format.type):
                    pic_placeholder = ph; break
            except Exception:
                continue
        img_stream = io.BytesIO(img_bytes)
        try:
            if pic_placeholder:
                pic_placeholder.insert_picture(img_stream)
            else:
                slide.shapes.add_picture(img_stream, Inches(8.0), Inches(1.5), width=Inches(2.5))
        except Exception:
            pass

    # Notes
    if note:
        try:
            slide.notes_slide.notes_text_frame.text = note[:1000]
        except Exception:
            pass

def build_ppt_from_template(template_bytes: bytes, struct: Dict[str, Any], reuse_images: bool=True) -> bytes:
    prs = Presentation(io.BytesIO(template_bytes))
    layout = choose_layout(prs)
    images = extract_template_images(prs) if reuse_images else []
    img_cycle = itertools.cycle(images) if images else None

    for s in struct.get("slides", []):
        title = s.get("title") or "Slide"
        bullets = s.get("bullets") or []
        notes = s.get("notes")
        img_bytes = next(img_cycle) if img_cycle else None
        add_slide(prs, layout, title, bullets, notes, img_bytes)

    out = io.BytesIO(); prs.save(out); out.seek(0)
    return out.read()

# ===============
# Polished UI v3 (API-Key Modal)
# ===============

INDEX_HTML = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Auto PPT Generator</title>
<style>
:root{
  --bg:#070c18; --bg2:#0d1426; --card:#0f1a33; --muted:#9fb0d9; --fg:#eaf1ff; --acc:#7cc4ff; --acc2:#78f3d3;
  --border:#1d2b4f; --good:#33d69f; --bad:#ff6b6b; --warn:#ffc857;
}
*{box-sizing:border-box}
html,body{height:100%}
body{
  margin:0; font-family:ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu;
  background:radial-gradient(1200px 600px at 10% -10%, #152045 0%, #070c18 60%), linear-gradient(180deg, #0c1224 0%, #070c18 100%);
  color:var(--fg);
}
.container{max-width:1100px; margin:40px auto; padding:24px}
.header{display:flex; align-items:center; gap:14px; margin-bottom:16px}
.brand{
  display:flex; align-items:center; gap:12px; padding:10px 14px; border:1px solid var(--border);
  border-radius:14px; background:linear-gradient(180deg,#0f1a33,#0b1326);
  box-shadow:0 8px 30px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.05);
}
.logo{width:22px; height:22px; border-radius:6px; background:linear-gradient(135deg,var(--acc),var(--acc2))}
.badge{padding:2px 10px; border-radius:999px; background:#15234a; color:#a6c8ff; font-size:12px; border:1px solid var(--border)}
.grid{display:grid; grid-template-columns:1.4fr .9fr; gap:18px}
.card{background:var(--card); border:1px solid var(--border); border-radius:16px; padding:18px; box-shadow:0 10px 40px rgba(0,0,0,.35)}
.card h3{margin:0 0 8px 0}
label{display:block; font-size:13px; color:#c9d4f7; margin:10px 0 6px}
input[type=text],input[type=password],textarea,select{
  width:100%; padding:12px 12px; border-radius:12px; border:1px solid #223158; background:#0b1427; color:#e6ecff; outline: none;
}
textarea{min-height:200px; resize:vertical}
small.help{display:block; color:var(--muted); margin-top:6px}
.row{display:flex; gap:12px; align-items:center}
.controls{display:grid; grid-template-columns:1fr 1fr; gap:12px}
.pill{padding:6px 10px; border-radius:10px; background:#0d1730; border:1px solid var(--border); color:#9bb5f7; font-size:12px}
.drop{
  display:flex; gap:14px; align-items:center; justify-content:center; padding:16px; border:1px dashed #315099;
  background:linear-gradient(180deg,#0c1733,#0a142b); border-radius:12px; cursor:pointer;
}
.drop.drag{border-color:var(--acc); background:linear-gradient(180deg,#0e2044,#0a142b)}
.kpi{display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:6px}
.kpi .tile{background:#0d1730; border:1px solid var(--border); border-radius:12px; padding:10px 12px}
.kpi .tile b{display:block; font-size:18px}
.btn{
  padding:12px 16px; border-radius:12px; border:1px solid #174f7b; background:linear-gradient(180deg,#7cc4ff,#69b2ff);
  color:#06243c; font-weight:800; cursor:pointer; box-shadow:0 6px 18px rgba(50,110,255,.35);
}
.btn.secondary{background:#142246; color:#d8e6ff; border:1px solid #243a73; box-shadow:none}
.btn:disabled{opacity:.6; cursor:not-allowed}
.status{margin-top:10px; font-size:13px}
.toast{
  position:fixed; right:18px; bottom:18px; padding:12px 14px; background:#0f1a33; border:1px solid var(--border);
  border-radius:12px; color:#cfe1ff; box-shadow:0 10px 30px rgba(0,0,0,.4); display:none;
}
.progress{height:10px; width:100%; background:#101a34; border:1px solid var(--border); border-radius:999px; overflow:hidden}
.progress > div{height:100%; width:0%; background:linear-gradient(90deg, var(--acc), var(--acc2)); transition:width .2s}
.footer{margin-top:14px; color:#7d8eb9; font-size:12px; text-align:center}
@media (max-width: 1020px){ .grid{grid-template-columns:1fr} }

/* API key modal */
.modal-backdrop{
  position:fixed; inset:0; background:rgba(0,0,0,.6); display:none; align-items:center; justify-content:center; z-index:9999;
}
.modal{
  width:min(480px, 92vw); background:#0f1a33; border:1px solid var(--border); border-radius:16px; padding:18px;
  box-shadow:0 20px 60px rgba(0,0,0,.5);
}
.modal h4{ margin:0 0 10px 0 }
.modal .actions{ display:flex; gap:10px; justify-content:flex-end; margin-top:12px }
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="brand">
        <div class="logo"></div>
        <div>
          <div style="font-weight:800; letter-spacing:.2px">Auto PPT Generator</div>
          <div style="font-size:12px; color:#9fb0d9">Gemini by default · Render-ready</div>
        </div>
      </div>
      <span class="badge">v1.3</span>
    </div>

    <div class="grid">
      <!-- Left: Content -->
      <div class="card">
        <h3>1) Content</h3>
        <label for="text">Input Text / Markdown</label>
        <textarea id="text" placeholder="Paste your long text or markdown here…" oninput="countChars()"></textarea>
        <div class="row" style="justify-content:space-between">
          <small class="help">Tip: Use markdown headings (## …) to become slide titles.</small>
          <span class="pill" id="charCount">0 chars</span>
        </div>

        <label for="guidance">One-line Guidance (optional)</label>
        <input id="guidance" type="text" placeholder="e.g., investor pitch, research summary, technical brief"/>
      </div>

      <!-- Right: Settings -->
      <div class="card">
        <h3>2) Settings</h3>
        <div class="controls">
          <div>
            <label for="provider">LLM Provider</label>
            <select id="provider" onchange="hydrateModels()">
              <option value="gemini" selected>Gemini (default)</option>
              <option value="openai">OpenAI</option>
              <option value="gorq">Gorq (Groq-compatible)</option>
            </select>
            <small class="help">API key is required to generate slides.</small>
          </div>
          <div>
            <label for="model">Model</label>
            <select id="model"></select>
          </div>
        </div>

        <label for="api_key">API Key <span style="color:#ffbdbd">(required)</span></label>
        <input id="api_key" type="password" placeholder="Paste your API key"/>

        <div class="kpi">
          <div class="tile"><small>Provider</small><b id="kProv">Gemini</b></div>
          <div class="tile"><small>Model</small><b id="kModel">gemini-2.5-flash</b></div>
          <div class="tile"><small>Notes</small><b id="kNotes">Off</b></div>
        </div>

        <div class="row" style="margin-top:12px">
          <input id="speaker_notes" type="checkbox" onchange="document.getElementById('kNotes').innerText=this.checked?'On':'Off'"/>
          <label for="speaker_notes" style="margin:0">Generate short speaker notes</label>
        </div>
      </div>
    </div>

    <div class="grid" style="margin-top:18px">
      <div class="card">
        <h3>3) Template (.pptx / .potx)</h3>
        <label>Upload or drag a PowerPoint template</label>
        <div id="drop" class="drop" onclick="pickFile()" ondragover="evtDragOver(event)" ondragleave="evtDragLeave()" ondrop="evtDrop(event)">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" style="opacity:.9">
            <path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3M12 3v12m0 0l-4-4m4 4l4-4" stroke="#a6c8ff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
          <span id="fileLabel">Drop template here or click to browse</span>
          <input id="template" type="file" accept=".pptx,.potx" style="display:none" onchange="filePicked(this)"/>
        </div>
        <small class="help">Max 25 MB. Your slide theme is inherited from the template.</small>
      </div>

      <div class="card">
        <h3>4) Generate</h3>
        <div class="row">
          <button class="btn" id="submitBtn" onclick="precheckAndMaybeAskKey()">Generate Deck</button>
          <button class="btn secondary" onclick="loadSample()">Load sample content</button>
        </div>
        <div class="status" id="status"></div>
        <div class="progress" style="margin-top:10px"><div id="bar"></div></div>
        <div class="footer">MIT Licensed • No keys are stored • Template images may be reused; no AI images are generated.</div>
      </div>
    </div>
  </div>

  <!-- Modal to request API key when empty -->
  <div class="modal-backdrop" id="keyBackdrop" role="dialog" aria-modal="true" aria-labelledby="keyTitle">
    <div class="modal">
      <h4 id="keyTitle">Enter your API key</h4>
      <p style="margin:6px 0 10px; color:#cfe1ff">An API key is required for the selected provider to generate slides.</p>
      <label for="api_key_modal">API Key</label>
      <input id="api_key_modal" type="password" placeholder="Paste your API key"/>
      <div class="actions">
        <button class="btn secondary" onclick="closeKeyModal()">Cancel</button>
        <button class="btn" onclick="saveKeyAndGenerate()">Continue</button>
      </div>
    </div>
  </div>

  <div class="toast" id="toast"></div>

<script>
function $(id){return document.getElementById(id);}

const MODELS = {
  gemini: [
    {id: 'gemini-2.5-flash', label: '2.5 flash (default)'},
    {id: 'gemini-2.5-pro', label: '2.5 pro'},
    {id: 'gemini-2.0-pro', label: '2 pro'},
    {id: 'gemini-2.0-flash', label: '2 flash'}
  ],
  openai: [
    {id: 'gpt-4o', label: 'GPT-4o'},
    {id: 'gpt-4o-mini', label: 'GPT-4o-mini'},
    {id: 'gpt-5-mini', label: 'GPT-5 mini'},
    {id: 'gpt-5', label: 'GPT-5'}
  ],
  gorq: [
    {id: 'openai120b', label: 'openai120b'},
    {id: 'deepseek-r1', label: 'deepseekr1'}
  ]
};

function hydrateModels(){
  const provider = $('provider').value;
  const modelSel = $('model');
  modelSel.innerHTML = '';
  (MODELS[provider] || []).forEach((m, idx) => {
    const opt = document.createElement('option');
    opt.value = m.id; opt.textContent = m.label;
    if(provider==='gemini' && idx===0) opt.selected = true;
    modelSel.appendChild(opt);
  });
  updateKPIs();
}

function updateKPIs(){
  $('kProv').innerText = $('provider').options[$('provider').selectedIndex].text;
  const modelSel = $('model');
  $('kModel').innerText = modelSel.value || '(manual)';
}

function countChars(){
  $('charCount').innerText = ($('text').value || '').length + ' chars';
}

function showToast(msg, kind='info'){
  const t = $('toast');
  t.style.display='block';
  t.textContent = msg;
  t.style.borderColor = (kind==='error')?'#ff6b6b':(kind==='warn')?'#ffc857':'#1d2b4f';
  setTimeout(()=>{ t.style.display='none'; }, 3200);
}

function pickFile(){ $('template').click(); }

function filePicked(inp){
  const f = inp.files[0];
  if(!f) return;
  if(!/\.(pptx|potx)$/i.test(f.name)) { showToast('Please choose a .pptx or .potx', 'warn'); inp.value=''; return; }
  if(f.size > 25 * 1024 * 1024) { showToast('File too large (max 25MB)', 'warn'); inp.value=''; return; }
  $('fileLabel').innerText = f.name + ' (' + Math.round(f.size/1024/1024*10)/10 + ' MB)';
}

function evtDragOver(e){ e.preventDefault(); $('drop').classList.add('drag'); }
function evtDragLeave(){ $('drop').classList.remove('drag'); }
function evtDrop(e){
  e.preventDefault();
  $('drop').classList.remove('drag');
  const f = e.dataTransfer.files[0];
  if(!f) return;
  $('template').files = e.dataTransfer.files;
  filePicked($('template'));
}

function loadSample(){
  const sample = `# Project Nova: Quarterly Update

## Highlights
- Revenue up 18% QoQ
- Churn reduced to 2.1%
- LA rollout completed

## Product
- Shipped AI summarize v2
- 12% latency improvement
- NPS +6 points

## GTM
- 3 enterprise wins (>$250k)
- New SE playbook live
- Partner-sourced pipeline +30%

## Next Quarter
- Launch self-serve billing
- PII redaction in Export
- India DC pilot`;
  $('text').value = sample; countChars(); showToast('Loaded sample content');
}

/* ---------- API key modal flow ---------- */
function precheckAndMaybeAskKey(){
  const key = $('api_key').value.trim();
  if(!key){
    openKeyModal();
    return;
  }
  generate(); // proceed if key present
}

function openKeyModal(){
  $('api_key_modal').value = '';
  $('keyBackdrop').style.display = 'flex';
  setTimeout(()=> $('api_key_modal').focus(), 50);
}

function closeKeyModal(){
  $('keyBackdrop').style.display = 'none';
}

function saveKeyAndGenerate(){
  const k = $('api_key_modal').value.trim();
  if(!k){ showToast('Please paste your API key', 'warn'); return; }
  $('api_key').value = k;
  closeKeyModal();
  generate();
}

/* ---------- Generation ---------- */
async function generate(){
  const text = $('text').value.trim();
  if(!text){ showToast('Please paste some content', 'warn'); return; }
  const f = $('template').files[0];
  if(!f){ showToast('Please upload a .pptx or .potx template', 'warn'); return; }
  const key = $('api_key').value.trim();
  if(!key){ showToast('API key is required', 'warn'); return; }

  $('submitBtn').disabled = true;
  $('status').textContent = 'Uploading template & generating…';
  $('bar').style.width = '5%';

  const fd = new FormData();
  fd.append('text', text);
  fd.append('guidance', $('guidance').value);
  fd.append('provider', $('provider').value);
  fd.append('model', $('model').value);
  fd.append('api_key', key); // now always present
  fd.append('speaker_notes', $('speaker_notes').checked ? 'true' : 'false');
  fd.append('template', f);

  try{
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/generate', true);
    xhr.responseType = 'blob';

    xhr.upload.onprogress = function(e){
      if(e.lengthComputable){
        const pct = Math.max(5, Math.min(90, Math.round((e.loaded/e.total)*90)));
        $('bar').style.width = pct + '%';
      }
    };
    xhr.onerror = function(){ throw new Error('Network error'); };
    xhr.onload = function(){
      if(xhr.status !== 200){
        const reader = new FileReader();
        reader.onload = () => {
          $('status').textContent = 'Error: ' + (reader.result || ('HTTP ' + xhr.status));
          showToast('Generation failed', 'error');
          $('submitBtn').disabled = false;
          $('bar').style.width = '0%';
        };
        reader.readAsText(xhr.response);
        return;
      }
      $('bar').style.width = '100%';
      const blob = xhr.response;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'generated.pptx';
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      $('status').textContent = 'Done ✅ Download should begin automatically.';
      showToast('Deck generated');
      $('submitBtn').disabled = false;
      setTimeout(()=>{$('bar').style.width='0%';}, 600);
    };
    xhr.send(fd);
  }catch(err){
    $('status').textContent = 'Error: ' + (err?.message || err);
    showToast('Generation failed', 'error');
    $('submitBtn').disabled = false;
    $('bar').style.width = '0%';
  }
}

// initialize
(function init(){
  hydrateModels();
  $('provider').addEventListener('change', updateKPIs);
  $('model').addEventListener('change', updateKPIs);
})();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse(INDEX_HTML)

@app.get("/health", response_class=PlainTextResponse)
def health():
    return "ok"

@app.post("/generate", response_class=FileResponse)
def generate_ppt(
    text: str = Form(...),
    guidance: str = Form(""),
    provider: str = Form(""),  # default handled in call_llm_struct
    model: str = Form(""),
    api_key: str = Form(""),
    speaker_notes: str = Form("false"),
    template: UploadFile = File(...)
):
    ext = (template.filename or "").lower()
    if not (ext.endswith(".pptx") or ext.endswith(".potx")):
        raise HTTPException(status_code=400, detail="Please upload a .pptx or .potx template.")
    template_bytes = template.file.read()
    if len(template_bytes) > 25 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="Template too large (max 25MB).")

    # Build slide structure (will error with 400 if no API key anywhere)
    struct = call_llm_struct(provider=provider, model=model, user_text=text, guidance=guidance, api_key=api_key)

    # Optional second-pass speaker notes (best-effort)
    try:
        if speaker_notes.lower() == "true" and isinstance(struct.get("slides"), list):
            have_notes = any(s.get("notes") for s in struct["slides"])
            if not have_notes:
                joined = "\n\n".join([
                    f"Title: {s.get('title','')}\nBullets:\n- " + "\n- ".join(s.get('bullets',[]))
                    for s in struct.get("slides",[])
                ])
                note_prompt = (
                    "For each slide, write 1–3 short speaker notes sentences.\n"
                    "Return STRICT JSON: {\"notes\": [\"...\", \"...\"]} matching slide order."
                )
                # Reuse same provider/model; call will still enforce API key
                notes_struct = call_llm_struct(provider or "gemini", model or "gemini-2.5-flash",
                                               f"{note_prompt}\n\nSlides:\n{joined}",
                                               guidance="speaker notes", api_key=api_key)
                notes = notes_struct.get("slides") or notes_struct.get("notes")
                if isinstance(notes, list):
                    flat = []
                    for n in notes:
                        if isinstance(n, dict) and "notes" in n: flat.append(n["notes"])
                        elif isinstance(n, str): flat.append(n)
                    for i, s in enumerate(struct.get("slides", [])):
                        if i < len(flat): s["notes"] = flat[i]
    except HTTPException:
        raise
    except Exception as e:
        logger.info(f"Speaker-notes enrichment skipped: {e}")

    # Build PPTX
    try:
        ppt_bytes = build_ppt_from_template(template_bytes, struct, reuse_images=True)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to build PPT: {e}")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp.write(ppt_bytes); tmp.flush(); tmp.close()
    filename = f"generated_{secrets.token_hex(3)}.pptx"
    return FileResponse(tmp.name,
                        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        filename=filename)
