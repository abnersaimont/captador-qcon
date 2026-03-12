# scrape_qconcursos_brave_marcadores_gui_cdp.py
# GUI sem barra de título + alça para arrastar; contadores em tempo real; reinjeção pós-filtro
from pathlib import Path
import ctypes
import json, os, re, time, subprocess, queue, socket, sys, shutil, traceback
from datetime import datetime
from urllib.parse import urlparse
from urllib.request import urlopen
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont  # [NOVO] para ajustar fonte das labels

# =============== BASE DIR (ao lado do .exe) ===============
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

# =============== CONFIG ===============
# [AJUSTE #0] Removido path hardcoded do Brave. Agora detectamos Brave/Chrome dinamicamente.
HOME_URL = "https://www.qconcursos.com/questoes-de-concursos/questoes"
LOGIN_CHECK_SELECTOR = "body[data-is-user-signed-in='true']"

# CDP
DEVTOOLS_HOST = "127.0.0.1"  # evita ::1
PREFERRED_DEVTOOLS_PORT = 9222

# >>> INÍCIO: escolha automática de porta livre <<<
def _pick_free_port(start=9222, end=9230, host=DEVTOOLS_HOST, timeout=0.2):
    """Retorna a primeira porta livre no intervalo [start, end]."""
    for port in range(start, end + 1):
        try:
            s = socket.create_connection((host, port), timeout=timeout)
            s.close()
        except Exception:
            return port  # conectar falhou: porta livre
    return start  # fallback (pode falhar depois, mas não altera o restante do fluxo)

DEVTOOLS_PORT = PREFERRED_DEVTOOLS_PORT
DEVTOOLS_URL = f"http://{DEVTOOLS_HOST}:{DEVTOOLS_PORT}"
# >>> FIM: escolha automática de porta livre <<<

# Posição inicial do painel (canto inferior direito, com margem)
PANEL_MARGIN_X = 24
PANEL_MARGIN_Y = 60

SELECTORS = {
    "card": ".js-question-item",
    "ano": ".q-question-info > span:nth-child(1)",
    "banca": ".q-question-info span:nth-child(2) > a",
    "enunciado": ".q-question-enunciation",
    "alternativas": ".q-radio-button",
    "altA_click": "div:nth-child(2) > .q-radio-button",
    "btn_responder": ".js-answer-btn",
    "gabarito_after_click": ".js-question-right-answer",
    # [NOVO QID] seletor do ID da questão no cabeçalho
    "qid": ".q-question-header .q-id a",
}

# [AJUSTE #2] Pasta de mídia ao lado do .exe
MEDIA_DIR = BASE_DIR / "midia"

# [AJUSTE #3] Perfil de automação separado (agora local à pasta do .exe)
AUTOMATION_USER_DATA_DIR = BASE_DIR / "perfil"

# [AJUSTE #4] Logs de erro apenas quando houver exceção
LOGS_DIR = BASE_DIR / "logs_erro"
TEMP_JSON_PATH = BASE_DIR / "questoes.temp.json"

# =============== UTILS (iguais, com pequenas adições) ===============
def clean(s: str) -> str:
    if not s: return ""
    return " ".join(s.replace("\xa0", " ").split())

def get_text_or_none(root, selector: str):
    try:
        el = root.query_selector(selector)
        if el:
            return clean(el.inner_text())
    except Exception:
        pass
    return None

def normalize_gabarito_text(g):
    if not g: return ""
    m = re.search(r"\b([A-E])\b", g, flags=re.I)
    if m: return m.group(1).upper()
    g = g.strip().upper()
    for ch in "ABCDE":
        if ch in g: return ch
    return ""

def ensure_dir(p: Path): p.mkdir(parents=True, exist_ok=True)

def guess_ext_from_url(url: str) -> str:
    try:
        path = urlparse(url).path
        ext = os.path.splitext(path)[1].lower()
        if ext in (".png",".jpg",".jpeg",".gif",".webp",".svg"): return ext
    except Exception: pass
    return ".png"

def save_binary(context, url: str, dest: Path) -> bool:
    try:
        resp = context.request.get(url, timeout=25000)
        if resp.ok:
            with open(dest, "wb") as f:
                f.write(resp.body())
            return True
    except Exception:
        pass
    return False

def save_json_atomic(path: Path, data):
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp_path.replace(path)

def load_partial_capture(path: Path):
    if not path.exists():
        return [], set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if not isinstance(dados, list):
            raise ValueError("JSON parcial invalido.")
        vistos = {str(item.get("QID")).strip() for item in dados if str(item.get("QID") or "").strip()}
        return dados, vistos
    except Exception:
        return [], set()

def persist_partial_capture(path: Path, dados):
    save_json_atomic(path, dados)

def remove_file_silent(path: Path):
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass

def find_img_by_src(root, target_src: str):
    if not root or not target_src:
        return None
    try:
        for img_el in root.query_selector_all("img"):
            try:
                if (img_el.get_attribute("src") or "") == target_src:
                    return img_el
            except Exception:
                pass
    except Exception:
        pass
    return None

def get_foreground_pid():
    if os.name != "nt":
        return None
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return None
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value or 0) or None
    except Exception:
        return None

def get_automation_browser_pids():
    if os.name != "nt":
        return set()
    try:
        profile = AUTOMATION_USER_DATA_DIR.resolve().as_posix().replace("'", "''")
        cmd = (
            "Get-CimInstance Win32_Process | "
            "Where-Object { $_.Name -match '^(brave|chrome)(\\.exe)?$' -and "
            "$_.CommandLine -like '*--remote-debugging-port=*' -and "
            f"$_.CommandLine -like '*--user-data-dir={profile}*' }} | "
            "Select-Object -ExpandProperty ProcessId"
        )
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return {int(line.strip()) for line in out.splitlines() if line.strip().isdigit()}
    except Exception:
        return set()

def serialize_node_with_markers(root_el):
    return root_el.evaluate(r"""
    (root) => {
      const norm = (s) => (s || '').replace(/[\s\u00A0]+/g, ' ');
      let clean = ''; const positions = []; const imagesMeta = [];
      function appendText(t){ t = norm(t); if(!t) return; if(clean && !/\s$/.test(clean)) clean+=' '; clean += t.replace(/^ +/,''); }
      function placeholderFor(img){
        const alt = img.getAttribute('alt') || '';
        const src = img.getAttribute('src') || '';
        const w = img.naturalWidth || img.width || parseInt(img.getAttribute('width')) || null;
        const h = img.naturalHeight || img.height || parseInt(img.getAttribute('height')) || null;
        let ph = null;
        if (/\.(png|jpe?g|gif|webp|svg)(\?|$)/i.test(alt)) ph = alt;
        if (!ph){
          let name = (src.split('/').pop() || '').split('?')[0] || 'IMG';
          let dims = (w && h) ? ` (${w}×${h})` : '';
          ph = name + dims;
        }
        return {placeholder: ph, src, alt, width: w, height: h};
      }
      function walk(n){
        if (n.nodeType===Node.TEXT_NODE){ appendText(n.textContent); return; }
        if (n.nodeType!==Node.ELEMENT_NODE) return;
        const tag = n.tagName.toLowerCase();
        if (tag==='img'){ const {placeholder,src,alt,width,height}=placeholderFor(n);
          positions.push({index:clean.length, placeholder, src, alt, width, height});
          imagesMeta.push({src,alt,placeholder,width,height}); return; }
        if (tag==='br'){ appendText(' '); return; }
        for (const c of n.childNodes) walk(c);
      }
      for (const c of root.childNodes) walk(c);
      positions.sort((a,b)=>a.index-b.index);
      let out='', last=0;
      for (const pos of positions){
        out += clean.slice(last,pos.index); out = out.replace(/\s+$/,'');
        out += (out?' ':'') + pos.placeholder + ' '; last = pos.index;
      }
      out += clean.slice(last);
      out = out.replace(/[ \t\u00A0]+/g,' ').replace(/^\s+|\s+$/g,'');
      return { text: out, images: imagesMeta };
    }""")

# [NOVO] Extrator de HTML saneado com placeholders (mesma lógica de placeholder)
def extract_html_with_placeholders(root_el):
    try:
        res = root_el.evaluate(r"""
        (root) => {
          function normText(s){ return (s||'').replace(/\u00A0/g,' '); }
          function placeholderFor(img){
            const alt=img.getAttribute('alt')||'';
            const src=img.getAttribute('src')||'';
            const w=img.naturalWidth||img.width||parseInt(img.getAttribute('width'))||null;
            const h=img.naturalHeight||img.height||parseInt(img.getAttribute('height'))||null;
            let ph=null;
            if (/\.(png|jpe?g|gif|webp|svg)(\?|$)/i.test(alt)) ph = alt;
            if (!ph){
              let name=(src.split('/').pop()||'').split('?')[0]||'IMG';
              let dims=(w&&h)?` (${w}×${h})`:''; 
              ph = name + dims;
            }
            return ph;
          }

          // Mantemos 'i' na lista por compatibilidade, mas nunca emitiremos <i>.
          const ALLOWED = new Set(['b','i','u','sup','sub','br','p']);

          function wrapInline(inner, {bold, underline}){
            // Italic foi removido de propósito.
            if (underline) inner = '<u>'+inner+'</u>';
            if (bold)      inner = '<b>'+inner+'</b>';
            return inner;
          }

          function toHTML(n){
            if (n.nodeType===Node.TEXT_NODE) return normText(n.textContent);
            if (n.nodeType!==Node.ELEMENT_NODE) return '';
            const tag=n.tagName.toLowerCase();

            if (tag==='img'){
              const ph = placeholderFor(n);
              return ph ? ph : '';
            }
            if (tag==='br') return '<br>';

            let inner = '';
            for (const c of n.childNodes) inner += toHTML(c);

            if (tag==='strong') return '<b>'+inner+'</b>';
            // >>> IGNORAR ITÁLICO <<<
            if (tag==='em' || tag==='i') return inner;

            if (tag==='span'){
              const st=(n.getAttribute('style')||'').toLowerCase();
              const underline = /text-decoration(?:-line)?\s*:\s*underline\b/.test(st);
              // Bold reconhece 'bold' e pesos 600–900
              const bold      = /font-weight\s*:\s*(bold|[6-9]00)\b/.test(st);
              // Italic é propositalmente ignorado (mesmo que exista no style)
              if (underline || bold){
                return wrapInline(inner, {bold, underline});
              }
              return inner;
            }

            if (ALLOWED.has(tag)) return '<'+tag+'>'+inner+'</'+tag+'>';
            if (tag==='div')      return '<p>'+inner+'</p>';
            return inner;
          }

          let html = '';
          for (const c of root.childNodes) html += toHTML(c);
          html = html.replace(/<p>\s*<\/p>/g,'');
          return {html};
        }""")
        return (res or {}).get("html") or ""
    except Exception:
        return ""


def reset_zoom(page):
    try:
        page.keyboard.down("Control"); page.keyboard.press("0"); page.keyboard.up("Control")
    except Exception: pass

def enforce_100_percent_css(page):
    try:
        page.add_style_tag(content="html, body { zoom: 1 !important; }")
    except Exception: pass

def normalize_view(page): reset_zoom(page); enforce_100_percent_css(page)

def marcar_cards_com_qid(page):
    page.evaluate(r"""
      () => {
        const cards = Array.from(document.querySelectorAll('.js-question-item'));
        cards.forEach((el, i) => el.setAttribute('data-qid', i));
      }
    """)

# ======== SELEÇÃO (borda verde; bloqueio por QID; DUPLICADOS com BLUR) ========
SELECAO_JS = r"""
(() => {
  const SEL_CARD = '.js-question-item';

  // CSS + helpers para blur em duplicadas
  if (!document.getElementById('qc-dup-style')) {
    const st = document.createElement('style');
    st.id = 'qc-dup-style';
    st.textContent = `
      .qc-dup-wrap { position: relative; }
      .qc-dup-blur { filter: blur(1.5px) saturate(0) brightness(0.5); transition: filter .15s ease; }
    `;
    document.head.appendChild(st);
  }
  function applyBlockedStyle(card) {
    card.classList.add('qc-dup-wrap', 'qc-dup-blur');
    card.title = 'Já captada (QID)';
  }
  function clearBlockedStyle(card) {
    card.classList.remove('qc-dup-wrap', 'qc-dup-blur');
    card.style.outline = '';
    card.style.opacity = '';
    card.removeAttribute('title');
  }

  try {
    if (window.__QC_obs2__) window.__QC_obs2__.disconnect();
    document.querySelectorAll('.qc-checkbox').forEach(x=>x.remove());
    Array.from(document.querySelectorAll(SEL_CARD)).forEach(c=> clearBlockedStyle(c));
  } catch(e){}

  const marcados = new Set();

  function qidFor(card){
    let qid = card.getAttribute('data-qid');
    if (qid == null || qid === '') {
      qid = Array.from(document.querySelectorAll(SEL_CARD)).indexOf(card);
      card.setAttribute('data-qid', qid);
    }
    return qid;
  }

  // [NOVO QID] extrai o QID visível no cabeçalho (ex.: "Q3577348")
  function qidFromCard(card){
    const el = card.querySelector('.q-question-header .q-id a');
    if (el) {
      const t = (el.innerText || '').trim();
      if (t) return t;
    }
    return null;
  }

  function prepararCard(card){
    if (!card || card.__qc_ready2__) return;
    card.__qc_ready2__ = true;
    const qid = qidFor(card);
    card.style.position = card.style.position || 'relative';

    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.className = 'qc-checkbox';
    Object.assign(chk.style, {
      position: 'absolute', top: '8px', left: '8px', transform: 'scale(1.3)',
      zIndex: 10, cursor: 'pointer', background: 'white'
    });
    card.prepend(chk);

    // bloquear se já visto (aplica apenas blur)
    try {
      const realId = qidFromCard(card);
      const dup = realId && window.__QC_vistos_qid && window.__QC_vistos_qid.has && window.__QC_vistos_qid.has(String(realId));
      if (dup) {
        chk.disabled = true;
        chk.checked = false;
        applyBlockedStyle(card);
      }
    } catch (e) {}

    card.addEventListener('click', (ev) => {
      const path = ev.composedPath && ev.composedPath();
      if (path && path.includes(chk)) return;
      if (chk.disabled) return; // não permite selecionar duplicado
      chk.checked = !chk.checked;
      chk.dispatchEvent(new Event('change', {bubbles:true}));
    });

    chk.addEventListener('change', () => {
      if (chk.checked) { marcados.add(qid); card.style.outline = '5px solid #22c55e'; }  // verde ao selecionar
      else { marcados.delete(qid); card.style.outline = ''; }
      if (window.pyUpdateSelectionCount) window.pyUpdateSelectionCount(marcados.size);
    });
  }

  Array.from(document.querySelectorAll(SEL_CARD)).forEach(prepararCard);
  const obs = new MutationObserver(() => {
    Array.from(document.querySelectorAll(SEL_CARD)).forEach(prepararCard);
  });
  obs.observe(document.body, {childList:true, subtree:true});
  window.__QC_obs2__ = obs;

  window.__QC_getSelecionados = () => Array.from(marcados).map(x=>parseInt(x,10)).filter(Number.isFinite).sort((a,b)=>a-b);
  window.__QC_limparSelecao = () => {
    document.querySelectorAll('.qc-checkbox').forEach(chk => {
      chk.checked = false; chk.dispatchEvent(new Event('change',{bubbles:true}));
    });
    marcados.clear();
    if (window.pyUpdateSelectionCount) window.pyUpdateSelectionCount(0);
  };

  // reprocessar duplicados quando Python atualizar o set
  window.__QC_markDuplicatesByQID = () => {
    document.querySelectorAll(SEL_CARD).forEach(card => {
      const chk = card.querySelector('.qc-checkbox');
      if (!chk) return;
      const realId = qidFromCard(card);
      const dup = realId && window.__QC_vistos_qid && window.__QC_vistos_qid.has && window.__QC_vistos_qid.has(String(realId));
      if (dup) {
        chk.disabled = true;
        chk.checked = false;
        applyBlockedStyle(card);   // apenas blur
      } else {
        clearBlockedStyle(card);
      }
    });
    const n = document.querySelectorAll('.qc-checkbox:checked').length;
    if (window.pyUpdateSelectionCount) window.pyUpdateSelectionCount(n);
  };
})();
"""

def obter_indices_selecionados(page):
    try:
        return page.evaluate("() => (window.__QC_getSelecionados && window.__QC_getSelecionados()) || []") or []
    except Exception:
        return []

# =============== DETECÇÃO DE NAVEGADOR E ABERTURA ===============
def _port_is_up(host: str, port: int, timeout: float = 0.3) -> bool:
    try:
        s = socket.create_connection((host, port), timeout=timeout)
        s.close()
        return True
    except Exception:
        return False

def _cdp_endpoint_ready(host: str, port: int, timeout: float = 1.0) -> bool:
    try:
        with urlopen(f"http://{host}:{port}/json/version", timeout=timeout) as resp:
            return resp.status == 200
    except Exception:
        return False

def _common_paths(exe_name: str):
    # Windows comuns
    pf = os.environ.get("PROGRAMFILES", r"C:\Program Files")
    pf86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
    local = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(pf) / "BraveSoftware" / "Brave-Browser" / "Application" / "brave.exe",
        Path(pf86) / "BraveSoftware" / "Brave-Browser" / "Application" / "brave.exe",
        Path(local) / "BraveSoftware" / "Brave-Browser" / "Application" / "brave.exe",
        Path(pf) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(pf86) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(local) / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]
    # Linux/mac (melhor esforço)
    candidates += [
        Path(shutil.which("brave") or "") if shutil.which("brave") else Path(),
        Path(shutil.which("brave-browser") or "") if shutil.which("brave-browser") else Path(),
        Path(shutil.which("google-chrome") or "") if shutil.which("google-chrome") else Path(),
        Path(shutil.which("chrome") or "") if shutil.which("chrome") else Path(),
        Path(shutil.which(exe_name) or "") if shutil.which(exe_name) else Path(),
    ]
    return [p for p in candidates if str(p)]

def _find_browser_exe(prefer: str = "brave"):
    prefer = prefer.lower()
    names = ["brave", "chrome"] if prefer == "brave" else ["chrome", "brave"]
    # tenta PATH
    for nm in names:
        for alias in (nm, f"{nm}.exe"):
            found = shutil.which(alias)
            if found and Path(found).is_file():
                return Path(found), nm
    # tenta locais comuns
    for nm in names:
        for p in _common_paths(nm):
            if p.name.lower().startswith(nm) and p.is_file():
                return p, nm
    return None, None

def start_external_browser_for_cdp(exe_path: Path, port: int):
    """Abre um navegador (Brave/Chrome) isolado com CDP ativo e user-data-dir dedicado."""
    ensure_dir(AUTOMATION_USER_DATA_DIR)

    # Porta já está de pé?
    if _cdp_endpoint_ready(DEVTOOLS_HOST, port):
        return
    if _port_is_up(DEVTOOLS_HOST, port):
        raise RuntimeError(f"A porta {port} ja esta em uso, mas nao respondeu como CDP valido.")

    # Lançar com user-data-dir dedicado
    cmd = [
        str(exe_path),
        f"--remote-debugging-port={port}",
        f"--user-data-dir={AUTOMATION_USER_DATA_DIR.as_posix()}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
    ]
    subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, close_fds=True)

    # Aguardar a porta subir
    deadline = time.time() + 20.0
    while time.time() < deadline:
        if _cdp_endpoint_ready(DEVTOOLS_HOST, port, timeout=0.5):
            return
        time.sleep(0.25)
    raise RuntimeError(f"Navegador com CDP não subiu na porta {port} em {DEVTOOLS_HOST}.")

def get_browser_context_page(p):
    """
    Tenta (1) Brave; (2) Chrome — ambos via CDP e perfil dedicado; (3) fallback Chromium do Playwright (perfil dedicado).
    Retorna: (browser_or_none, context, page)
    """
    if _cdp_endpoint_ready(DEVTOOLS_HOST, DEVTOOLS_PORT):
        browser = p.chromium.connect_over_cdp(DEVTOOLS_URL)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()
        return browser, context, page

    exe, which = _find_browser_exe(prefer="brave")
    if exe:
        try:
            start_external_browser_for_cdp(exe, DEVTOOLS_PORT)
            browser = p.chromium.connect_over_cdp(DEVTOOLS_URL)
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.pages[0] if context.pages else context.new_page()
            return browser, context, page
        except Exception:
            pass

    # tenta Chrome explicitamente se não achou Brave
    exe, which = _find_browser_exe(prefer="chrome")
    if exe:
        try:
            start_external_browser_for_cdp(exe, DEVTOOLS_PORT)
            browser = p.chromium.connect_over_cdp(DEVTOOLS_URL)
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.pages[0] if context.pages else context.new_page()
            return browser, context, page
        except Exception:
            pass

    fallback_port = _pick_free_port(9223, 9230)
    fallback_url = f"http://{DEVTOOLS_HOST}:{fallback_port}"
    for prefer in ("brave", "chrome"):
        exe, which = _find_browser_exe(prefer=prefer)
        if exe:
            try:
                start_external_browser_for_cdp(exe, fallback_port)
                browser = p.chromium.connect_over_cdp(fallback_url)
                context = browser.contexts[0] if browser.contexts else browser.new_context()
                page = context.pages[0] if context.pages else context.new_page()
                return browser, context, page
            except Exception:
                pass

    # Fallback: Chromium do Playwright (perfil persistente)
    ensure_dir(AUTOMATION_USER_DATA_DIR)
    context = p.chromium.launch_persistent_context(
        user_data_dir=AUTOMATION_USER_DATA_DIR.as_posix(),
        headless=False,
        args=["--start-maximized"],
    )
    page = context.pages[0] if context.pages else context.new_page()
    return None, context, page

# ========= GUI (sem barra de título + ALÇA) =========
class Painel:
    def __init__(self):
        self.actions = queue.Queue()
        self.root = tk.Tk()
        self.root.title("")
        self.root.overrideredirect(True)     # remove barra superior
        self.root.attributes("-topmost", False)
        self.root.resizable(False, False)
        self._topmost_state = False

        frm = ttk.Frame(self.root, padding=8)
        frm.grid(row=0, column=0, sticky="nsew")

        # ---- ALÇA DE ARRASTAR (grip) ----
        grip = ttk.Frame(frm, height=10)
        grip.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0,6))
        grip.columnconfigure(0, weight=1)
        grip.configure(cursor="fleur")
        self.btn_sair = ttk.Button(grip, text="X", width=3, command=lambda: self.actions.put("sair"))
        self.btn_sair.grid(row=0, column=1, sticky="e")
        self._drag_origin = (0, 0)
        self._win_origin = (0, 0)
        def start_move(e):
            self._drag_origin = (e.x_root, e.y_root)
            self._win_origin = (self.root.winfo_x(), self.root.winfo_y())
        def do_move(e):
            dx = e.x_root - self._drag_origin[0]
            dy = e.y_root - self._drag_origin[1]
            self.root.geometry(f"+{self._win_origin[0]+dx}+{self._win_origin[1]+dy}")
        grip.bind("<ButtonPress-1>", start_move)
        grip.bind("<B1-Motion>", do_move)

        # ---- Botões ----
        self.btn_filtrar = ttk.Button(frm, text="Selecionar", command=lambda: self.actions.put("filtrar_ok"))
        self.btn_captar   = ttk.Button(frm, text="Captar",   command=lambda: self.actions.put("captar"))
        self.btn_finalizar = ttk.Button(frm, text="Finalizar", command=lambda: self.actions.put("finalizar"))
        self.btn_filtrar.grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        self.btn_captar.grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self.btn_finalizar.grid(row=1, column=2, padx=4, pady=4, sticky="ew")

        ttk.Separator(frm, orient="horizontal").grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6,2))

        # Labels (fonte MAIOR, sem alterar botões)
        self.lbl_sel = ttk.Label(frm, text="Selecionadas nesta página: 0")
        self.lbl_tot = ttk.Label(frm, text="Total captadas: 0")
        self.lbl_sel.grid(row=3, column=0, columnspan=3, sticky="w")
        self.lbl_tot.grid(row=4, column=0, columnspan=3, sticky="w")

        # fonte maior para contadores
        try:
            _base = tkfont.nametofont("TkDefaultFont").copy()
            _base.configure(size=13)
            self.lbl_sel.configure(font=_base)
            self.lbl_tot.configure(font=_base)
        except Exception:
            pass
        try:
            self.root.after(60, self.position_bottom_right)
        except Exception:
            pass

        # ---- separador abaixo dos contadores ----
        ttk.Separator(frm, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(6,2)
        )

        # ---- texto da empresa ----
        self.lbl_empresa = ttk.Label(frm, text="Beckm", anchor="e")
        self.lbl_empresa.grid(row=6, column=0, columnspan=3, sticky="e")
        try:
            _small = tkfont.nametofont("TkDefaultFont").copy()
            _small.configure(size=5)
            self.lbl_empresa.configure(font=_small)
        except Exception:
            pass

        for i in range(3):
            frm.columnconfigure(i, weight=1)
        self.position_bottom_right()

        # posição inicial: canto inferior direito
        try:
            self.root.update_idletasks()
            w = max(self.root.winfo_reqwidth(), self.root.winfo_width())
            h = max(self.root.winfo_reqheight(), self.root.winfo_height())
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, sw - w - PANEL_MARGIN_X)
            y = max(0, sh - h - PANEL_MARGIN_Y)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

        # Captar começa DESABILITADO até ter seleção
        try:
            self.btn_captar.state(["disabled"])
        except Exception:
            pass

    # habilitar/desabilitar botão Captar conforme a contagem
    def update_selected(self, n: int):
        self.lbl_sel.config(text=f"Selecionadas nesta página: {n}")
        try:
            if n and int(n) > 0:
                self.btn_captar.state(["!disabled"])
            else:
                self.btn_captar.state(["disabled"])
        except Exception:
            pass

    def update_total(self, n: int):
        self.lbl_tot.config(text=f"Total captadas: {n}")

    def get_action_nonblocking(self):
        try: return self.actions.get_nowait()
        except queue.Empty: return None

    def loop_once(self):
        self.root.update_idletasks(); self.root.update()

    def set_filtrar_enabled(self, enabled: bool):
        if enabled:
            self.btn_filtrar.state(["!disabled"])
        else:
            self.btn_filtrar.state(["disabled"])

    def set_topmost(self, enabled: bool):
        enabled = bool(enabled)
        if enabled == self._topmost_state:
            return
        try:
            self.root.attributes("-topmost", enabled)
            self._topmost_state = enabled
        except Exception:
            pass

    def position_bottom_right(self):
        try:
            self.root.update_idletasks()
            w = max(self.root.winfo_reqwidth(), self.root.winfo_width())
            h = max(self.root.winfo_reqheight(), self.root.winfo_height())
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, sw - w - PANEL_MARGIN_X)
            y = max(0, sh - h - PANEL_MARGIN_Y)
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

# ========= helpers de página =========
def page_signature(page):
    """Assinatura leve da página de resultados para saber se mudou: (url, qtd_cards)."""
    try:
        url = page.url
    except Exception:
        url = ""
    try:
        count = page.evaluate("() => document.querySelectorAll('.js-question-item').length") or 0
    except Exception:
        count = 0
    return (url, int(count))

# ========= NOVO HELPER: revelar + capturar "Texto associado" =========
def revelar_e_capturar_texto_associado(card, context, q_folder, total):
    """
    Procura e, se preciso, expande o 'Texto associado' e captura seu conteúdo + imagens.
    Compatível com:
      - .q-question-text--print-hide a.q-link
      - .q-question-text a.q-link
    Também faz fallback direto para o painel .collapse[id^='question-'][id$='-text'].
    """
    def _serialize_and_download(content_el):
        texto_assoc_local = ""
        assoc_imagem_paths_local = []
        assoc_imagens_info_out_local = []
        texto_assoc_html_local = ""

        # Serializar texto + imagens
        try:
            res = serialize_node_with_markers(content_el)
            texto_assoc_local = (res.get("text") or "").strip()
            imagens_meta = res.get("images") or []
        except Exception:
            try:
                texto_assoc_local = clean(content_el.inner_text() or "")
                imagens_meta = []
            except Exception:
                imagens_meta = []
                texto_assoc_local = ""

        # HTML saneado
        try:
            texto_assoc_html_local = extract_html_with_placeholders(content_el) or ""
        except Exception:
            texto_assoc_html_local = ""

        # Baixar imagens
        ensure_dir(q_folder)
        for j, info in enumerate(imagens_meta, start=1):
            src = info.get("src") or ""
            placeholder = (info.get("placeholder") or "").strip()
            alt = info.get("alt") or ""
            width = info.get("width"); height = info.get("height")
            base_name = f"Questão {total+1} - Texto Associado" if j == 1 else f"Questão {total+1} - Texto Associado.{j}"
            try:
                ext = (os.path.splitext(placeholder.split(" ")[0])[1].lower()
                       if placeholder and re.search(r"\.(png|jpe?g|gif|webp|svg)", placeholder, flags=re.I)
                       else guess_ext_from_url(src))
            except Exception:
                ext = ".png"
            dest = q_folder / (base_name + ext)
            saved = False
            if src.startswith("http"):
                saved = save_binary(context, src, dest)
            if not saved:
                try:
                    img_el = find_img_by_src(content_el, src)
                    if img_el:
                        #img_el.scroll_into_view_if_needed()
                        img_el.screenshot(path=dest.as_posix())
                        saved = True
                except Exception:
                    pass
            if saved:
                path_str = dest.resolve().as_posix()
                assoc_imagem_paths_local.append(path_str)
                assoc_imagens_info_out_local.append({
                    "url": src,
                    "alt": alt,
                    "placeholder": placeholder,
                    "width": width,
                    "height": height,
                    "path": path_str,
                    "scope": "texto_associado"
                })

        return texto_assoc_local, assoc_imagem_paths_local, assoc_imagens_info_out_local, texto_assoc_html_local

    # 1) Tentar achar o link em ambas as variantes + fallback genérico
    link = None
    try:
        link = (card.query_selector(".q-question-text--print-hide a.q-link")
                or card.query_selector(".q-question-text a.q-link")
                or card.query_selector("a.q-link[data-toggle='collapse'][href^='#question-']"))
    except Exception:
        link = None

    # 2) Se não achar o link, tentar capturar diretamente o painel já presente
    if not link:
        try:
            content = card.query_selector("div.collapse[id^='question-'][id$='-text']")
            if content:
                return _serialize_and_download(content)
        except Exception:
            pass
        return "", [], [], ""

    # 3) Verificar expandido; se não, clicar e aguardar o painel surgir
    expanded = False
    try:
        expanded = (link.get_attribute("aria-expanded") or "").lower() == "true"
    except Exception:
        expanded = False

    target_sel = link.get_attribute("href") or ""
    if not expanded:
        try:
            link.scroll_into_view_if_needed()
        except Exception:
            pass
        clicked = False
        for _ in range(2):
            try:
                link.click()
                clicked = True
                break
            except Exception:
                try:
                    link.evaluate("el => el.click()")
                    clicked = True
                    break
                except Exception:
                    time.sleep(0.1)
        # Aguarda o painel existir/ficar visível (se tivermos id destino)
        if clicked and target_sel.startswith("#"):
            try:
                # existir
                card.wait_for_selector(target_sel, timeout=1000)
            except Exception:
                pass
            try:
                # visível (se suportado)
                card.wait_for_selector(target_sel, state="visible", timeout=500)
            except Exception:
                pass

    # 4) Encontrar o container do texto associado
    content = None
    if target_sel.startswith("#"):
        try:
            content = card.query_selector(target_sel)
        except Exception:
            content = None
    if not content:
        try:
            content = link.evaluate_handle("""
                (root, link) => {
                  const id = link.getAttribute('href');
                  const panel = id ? document.querySelector(id) : null;
                  if (panel) return panel;
                  let el = link.nextElementSibling;
                  while (el && !(el.classList && el.classList.contains('collapse'))) {
                    el = el.nextElementSibling;
                  }
                  return el || null;
                }
            """, link).as_element()
        except Exception:
            content = None
    if not content:
        try:
            content = card.query_selector("div.collapse[id^='question-'][id$='-text']")
        except Exception:
            content = None
    if not content:
        return "", [], [], ""

    # 5) Serializar + baixar imagens + HTML
    return _serialize_and_download(content)

# [NOVO QID] helper: extrair QID visível do card
def get_qid_from_card(card):
    try:
        el = card.query_selector(SELECTORS["qid"])
        if el:
            txt = (el.inner_text() or "").strip()
            return txt  # ex.: "Q3577348"
    except Exception:
        pass
    return None

# [AJUSTE #5] Helper leve de retry para wait_for_selector (usado após “Selecionar”)
def wait_for_selector_retry(page, selector: str, timeout: int = 20000, retries: int = 2, delay: float = 0.6):
    last_err = None
    for _ in range(max(1, retries)):
        try:
            return page.wait_for_selector(selector, timeout=timeout)
        except PWTimeoutError as e:
            last_err = e
            time.sleep(max(0.05, delay))
    if last_err:
        raise last_err

# ========= Principal (CDP + perfil de automação ou fallback Chromium) =========
def build_export_path():
    downloads = Path(os.path.join(os.path.expanduser("~"), "Downloads"))
    ensure_dir(downloads)
    out_path = downloads / "questoes.json"
    if out_path.exists():
        i = 2
        while True:
            cand = downloads / f"questoes ({i}).json"
            if not cand.exists():
                return cand
            i += 1
    return out_path

def detect_question_issues(registro):
    issues = []
    alternativas = [registro.get(f"Alternativa{letra}", "") for letra in "ABCDE"]
    if not registro.get("Ano"):
        issues.append("sem ano")
    if not registro.get("Banca"):
        issues.append("sem banca")
    if not registro.get("Enunciado"):
        issues.append("sem enunciado")
    if sum(1 for alt in alternativas if clean(alt)) < 2:
        issues.append("alternativas insuficientes")
    if not registro.get("Gabarito"):
        issues.append("sem gabarito")
    return issues

def main():
    print(f"Automação Qconcursos")

    dados, vistos_qid = load_partial_capture(TEMP_JSON_PATH)
    vistos_qid = set()  # set de QIDs já captados
    total = len(dados)
    issues_detected = []
    painel = Painel()
    if total > 0:
        painel.update_total(total)
        for idx, registro in enumerate(dados, start=1):
            critical_issues = detect_question_issues(registro)
            if critical_issues:
                qref = (registro.get("QID") or f"questao_{idx}")
                issues_detected.append(f"{qref}: {', '.join(critical_issues)}")
        print(f"ℹ️ Captação parcial recuperada: {total} questão(ões).")
    vistos_qid = {str(item.get("QID")).strip() for item in dados if str(item.get("QID") or "").strip()}

    with sync_playwright() as p:
        # [AJUSTE #0] Detecta Brave/Chrome; se nenhum, usa Chromium do Playwright.
        browser_or_none, context, page = get_browser_context_page(p)
        context.set_default_timeout(4000)
        context.set_default_navigation_timeout(8000)
        automation_pids = get_automation_browser_pids()
        last_pid_refresh = 0.0


        # Atualiza contagem na UI e habilita/desabilita "Captar"
        page.expose_function("pyUpdateSelectionCount", lambda n: painel.update_selected(int(n)))
        try:
            page.on("close", lambda *_: painel.actions.put("browser_closed"))
        except Exception:
            pass
        try:
            context.on("close", lambda *_: painel.actions.put("browser_closed"))
        except Exception:
            pass
        try:
            if browser_or_none:
                browser_or_none.on("disconnected", lambda *_: painel.actions.put("browser_closed"))
        except Exception:
            pass

        page.goto(HOME_URL, wait_until="domcontentloaded")
        normalize_view(page)
        try:
            page.wait_for_selector(LOGIN_CHECK_SELECTOR, timeout=5000)
            print("✅ Sessão ativa.")
        except PWTimeoutError:
            print("⚠️ Login não detectado neste perfil de automação (faça 1x e ficará salvo).")

        ensure_dir(MEDIA_DIR)

        running = True
        selecao_pronta = False
        last_poll = 0

        # para bloquear/desbloquear o botão "Filtragem OK"
        last_sig = page_signature(page)
        painel.set_filtrar_enabled(True)

        while running:
            painel.loop_once()
            act = painel.get_action_nonblocking()
            now = time.time()
            if now - last_pid_refresh > 1.0:
                refreshed_pids = get_automation_browser_pids()
                if refreshed_pids:
                    automation_pids = refreshed_pids
                last_pid_refresh = now
            painel.set_topmost(get_foreground_pid() in automation_pids if automation_pids else False)

            if act == "browser_closed":
                print("ℹ️ A janela da automação foi fechada. Encerrando o app para evitar erros.")
                running = False
                break

            # Detecta mudança de página/resultado para reabilitar o botão
            current_sig = page_signature(page)
            if selecao_pronta and current_sig != last_sig:
                selecao_pronta = False
                painel.set_filtrar_enabled(True)
                painel.update_selected(0)
                try:
                    page.evaluate("() => window.__QC_limparSelecao && window.__QC_limparSelecao()")
                except Exception:
                    pass
                last_sig = current_sig
            else:
                last_sig = current_sig

            if selecao_pronta:
                if now - last_poll > 0.2:
                    try:
                        n = page.evaluate("() => document.querySelectorAll('.qc-checkbox:checked').length")
                        painel.update_selected(int(n))
                    except Exception:
                        pass
                    last_poll = now

            if act == "filtrar_ok":
                painel.set_filtrar_enabled(False)

                try:
                    # [AJUSTE #5] retry leve
                    wait_for_selector_retry(page, SELECTORS["card"], timeout=8000, retries=2, delay=0.4)
                except PWTimeoutError:
                    print("⏳ Nenhum card encontrado; ajuste os filtros. Se isso persistir, o padrão do site pode ter mudado.")
                    selecao_pronta = False
                    painel.set_filtrar_enabled(True)
                    continue

                normalize_view(page)
                try:
                    page.evaluate("() => { try { delete window.__QC_MODO_SELECAO2__; } catch(e){} }")
                except Exception:
                    pass
                marcar_cards_com_qid(page)

                # disponibiliza QIDs já vistos no DOM antes de injetar o JS
                try:
                    page.evaluate("arr => { window.__QC_vistos_qid = new Set(arr.map(String)); }", list(vistos_qid))
                except Exception:
                    pass

                page.add_script_tag(content=SELECAO_JS)
                selecao_pronta = True
                painel.update_selected(0)
                last_sig = page_signature(page)
                print("🟩 Seleção habilitada — marque e clique em 'Captar'.")

            elif act == "captar":
                if not selecao_pronta:
                    print("⚠️ Clique primeiro em 'Filtragem OK'."); continue
                selecionados = obter_indices_selecionados(page)
                if not selecionados:
                    print("⚠️ Nenhuma questão marcada."); continue
                try:
                    cards = page.query_selector_all(SELECTORS["card"])
                except Exception:
                    cards = []
                alvo = [cards[i] for i in selecionados if 0 <= i < len(cards)]

                for card in alvo:
                    #try: card.scroll_into_view_if_needed()
                    #except Exception: pass

                    # pega QID e pula se já visto
                    qid_real = get_qid_from_card(card) or ""
                    if qid_real and qid_real in vistos_qid:
                        continue

                    ano = get_text_or_none(card, SELECTORS["ano"])
                    banca = get_text_or_none(card, SELECTORS["banca"])

                    enun_el = card.query_selector(SELECTORS["enunciado"])
                    enunciado_text = ""; images_meta_enun = []; enunciado_html = ""
                    if enun_el:
                        try:
                            res = serialize_node_with_markers(enun_el); enunciado_text = res.get("text") or ""; images_meta_enun = res.get("images") or []
                        except Exception:
                            enunciado_text = clean(enun_el.inner_text())
                        try:
                            enunciado_html = extract_html_with_placeholders(enun_el) or ""
                        except Exception:
                            enunciado_html = ""

                    alt_texts = ["", "", "", "", ""]; alt_images_meta = [[], [], [], [], []]
                    alt_htmls = ["", "", "", "", ""]
                    try:
                        alt_nodes = card.query_selector_all(SELECTORS["alternativas"])
                        for idx, node in enumerate(alt_nodes[:5]):
                            try:
                                res = serialize_node_with_markers(node)
                                alt_texts[idx] = res.get("text") or ""
                                alt_images_meta[idx] = res.get("images") or []
                            except Exception:
                                alt_texts[idx] = clean(node.inner_text()); alt_images_meta[idx] = []
                            # [NOVO] HTML: preferir o conteúdo real da alternativa (sem o rótulo A/B/C)
                            try:
                                sub = node.query_selector(".js-alternative-content") or node
                                alt_htmls[idx] = extract_html_with_placeholders(sub) or ""
                            except Exception:
                                alt_htmls[idx] = ""
                    except Exception: pass
                    altA, altB, altC, altD, altE = alt_texts
                    altA_html, altB_html, altC_html, altD_html, altE_html = alt_htmls

                    gabarito_raw = ""
                    try:
                        nodeA = card.query_selector(SELECTORS["altA_click"])
                        if nodeA: nodeA.scroll_into_view_if_needed(); nodeA.click()
                        btn = card.query_selector(SELECTORS["btn_responder"])
                        if btn:
                            btn.scroll_into_view_if_needed(); btn.click()
                            try: card.wait_for_selector(SELECTORS["gabarito_after_click"], timeout=1800)
                            except PWTimeoutError: pass
                        gabarito_raw = get_text_or_none(card, SELECTORS["gabarito_after_click"]) or ""
                    except Exception: pass
                    gabarito = normalize_gabarito_text(gabarito_raw)

                    q_folder = MEDIA_DIR / f"questao_{total+1:04d}"
                    ensure_dir(q_folder)
                    imagem_paths = []; mapa_imagens = {}; imagens_info_out = []

                    # ===== Imagens do enunciado =====
                    for idx_img, info in enumerate(images_meta_enun, start=1):
                        src = info.get("src") or ""
                        placeholder = (info.get("placeholder") or "").strip()
                        alt = info.get("alt") or ""
                        width = info.get("width"); height = info.get("height")
                        qnum = total + 1
                        base_name = f"Questão {qnum}" if idx_img == 1 else f"Questão {qnum}.{idx_img-1}"
                        try:
                            ext = (os.path.splitext(placeholder.split(" ")[0])[1].lower()
                                   if placeholder and re.search(r"\.(png|jpe?g|gif|webp|svg)", placeholder, flags=re.I)
                                   else guess_ext_from_url(src))
                        except Exception: ext = ".png"
                        dest = q_folder / (base_name + ext)
                        saved = False
                        if src.startswith("http"): saved = save_binary(context, src, dest)  # (NOTE: manter como estava no seu script original)
                        if not saved:
                            try:
                                if enun_el:
                                    img_el = find_img_by_src(enun_el, src)
                                    if img_el: img_el.scroll_into_view_if_needed(); img_el.screenshot(path=dest.as_posix()); saved = True
                            except Exception: pass
                        if saved:
                            path_str = dest.resolve().as_posix()
                            imagem_paths.append(path_str)
                            if placeholder: mapa_imagens[placeholder] = path_str
                            imagens_info_out.append({"url":src,"alt":alt,"placeholder":placeholder,"width":width,"height":height,"path":path_str,"scope":"enunciado"})

                    # ===== Texto associado =====
                    texto_assoc, assoc_imagem_paths, assoc_imagens_info, texto_assoc_html = revelar_e_capturar_texto_associado(
                        card, context, q_folder, total
                    )

                    # ===== Imagens das alternativas =====
                    letras = ["A","B","C","D","E"]
                    for idx_alt, images_meta in enumerate(alt_images_meta):
                        letra = letras[idx_alt]
                        alt_node = None
                        try: alt_node = card.query_selector_all(SELECTORS["alternativas"])[idx_alt]
                        except Exception: alt_node = None
                        for j, info in enumerate(images_meta, start=1):
                            src = info.get("src") or ""
                            placeholder = (info.get("placeholder") or "").strip()
                            alt = info.get("alt") or ""
                            width = info.get("width"); height = info.get("height")
                            qnum = total + 1
                            base_name = f"Questão {qnum} - Alt {letra}" if j == 1 else f"Questão {qnum} - Alt {letra}.{j}"
                            try:
                                ext = (os.path.splitext(placeholder.split(" ")[0])[1].lower()
                                       if placeholder and re.search(r"\.(png|jpe?g|gif|webp|svg)", placeholder, flags=re.I)
                                       else guess_ext_from_url(src))
                            except Exception: ext = ".png"
                            dest = q_folder / (base_name + ext)
                            saved = False
                            if src.startswith("http"): saved = save_binary(context, src, dest)
                            if not saved:
                                try:
                                    if alt_node:
                                        img_el = find_img_by_src(alt_node, src)
                                        if img_el: img_el.scroll_into_view_if_needed(); img_el.screenshot(path=dest.as_posix()); saved = True
                                except Exception: pass
                            if saved:
                                path_str = dest.resolve().as_posix()
                                imagem_paths.append(path_str)
                                if placeholder: mapa_imagens[placeholder] = path_str
                                imagens_info_out.append({"url":src,"alt":alt,"placeholder":placeholder,"width":width,"height":height,"path":path_str,"scope":f"alternativa_{letra}"})

                    registro = {
                        "Ano": ano or "",
                        "Banca": banca or "",
                        "Enunciado": enunciado_text or "",
                        "AlternativaA": altA or "",
                        "AlternativaB": altB or "",
                        "AlternativaC": altC or "",
                        "AlternativaD": altD or "",
                        "AlternativaE": altE or "",
                        "Gabarito": gabarito,
                        "Imagens": imagem_paths,
                        "MapaImagens": mapa_imagens,
                        "ImagensInfo": imagens_info_out,
                        # Campos do Texto associado (texto + imagens + HTML)
                        "TextoAssociado": texto_assoc or "",
                        "TextoAssociadoImagens": assoc_imagem_paths,
                        "TextoAssociadoMapaImagens": [{"path": p, "scope": "texto_associado"} for p in assoc_imagem_paths],
                        "TextoAssociadoImagensInfo": assoc_imagens_info,
                        "QID": qid_real,
                        # [NOVO] Campos HTML preservando formatação
                        "EnunciadoHTML": enunciado_html or "",
                        "AlternativaAHTML": altA_html or "",
                        "AlternativaBHTML": altB_html or "",
                        "AlternativaCHTML": altC_html or "",
                        "AlternativaDHTML": altD_html or "",
                        "AlternativaEHTML": altE_html or "",
                        "TextoAssociadoHTML": texto_assoc_html or "",
                    }

                    # ====== ALTERAÇÃO ÚNICA: deduplicar SOMENTE por QID ======
                    if qid_real:
                        vistos_qid.add(str(qid_real))
                    dados.append(registro)
                    total += 1
                    painel.update_total(total)
                    persist_partial_capture(TEMP_JSON_PATH, dados)
                    critical_issues = detect_question_issues(registro)
                    if critical_issues:
                        qref = qid_real or f"questao_{total}"
                        issues_detected.append(f"{qref}: {', '.join(critical_issues)}")
                    # ====== FIM DA ALTERAÇÃO ÚNICA ======

                # envia set atualizado para bloquear duplicados ainda visíveis
                try:
                    page.evaluate("arr => { window.__QC_vistos_qid = new Set(arr.map(String)); }", list(vistos_qid))
                    page.evaluate("() => window.__QC_markDuplicatesByQID && window.__QC_markDuplicatesByQID()")
                except Exception:
                    pass

                try:
                    page.evaluate("() => window.__QC_limparSelecao && window.__QC_limparSelecao()")
                except Exception: pass
                painel.update_selected(0)  # limpa seleção (desabilita "Captar")

            elif act == "finalizar":
                if total == 0 or not dados:
                    print("\nℹ️ Nenhuma questão captada nesta sessão — nenhum arquivo foi gerado.")
                    continue

                out_path = build_export_path()
                save_json_atomic(out_path, dados)
                remove_file_silent(TEMP_JSON_PATH)

                print(f"\n✅ Captação finalizada. Questões coletadas: {total}")
                print(f"🗂️ Exportado: {out_path.as_posix()}")
                print(f"🖼️ Mídias em: {MEDIA_DIR.resolve().as_posix()}")
                if issues_detected:
                    print(f"⚠️ Aviso: {len(issues_detected)} questão(ões) tiveram campos críticos faltando.")
                    for issue in issues_detected[:10]:
                        print(f"   - {issue}")
                    if len(issues_detected) > 10:
                        print(f"   - ... e mais {len(issues_detected) - 10}")

                dados = []
                vistos_qid = set()
                total = 0
                issues_detected = []
                selecao_pronta = False
                painel.update_total(0)
                painel.update_selected(0)
                painel.set_filtrar_enabled(True)
                try:
                    page.evaluate("() => window.__QC_limparSelecao && window.__QC_limparSelecao()")
                except Exception:
                    pass
                print("ℹ️ Sessão limpa. Você pode iniciar uma nova captação sem fechar o app.")
                continue
                # Se nada foi captado, NÃO gera JSON
                if total == 0 or not dados:
                    print("\nℹ️ Nenhuma questão captada — nenhum arquivo foi gerado.")
                    print(f"🖼️ Pasta de mídias (se criada): {MEDIA_DIR.resolve().as_posix()}")
                    running = False
                    break

                downloads = Path(os.path.join(os.path.expanduser("~"), "Downloads"))
                ensure_dir(downloads)
                # [AJUSTE #6] Nome fixo questoes.json (+ sufixos)
                out_path = downloads / "questoes.json"
                if out_path.exists():
                    i = 2
                    while True:
                        cand = downloads / f"questoes ({i}).json"
                        if not cand.exists(): out_path = cand; break
                        i += 1
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(dados, f, ensure_ascii=False, indent=2)
                print(f"\n✅ Finalizado. Questões coletadas: {total}")
                print(f"🗂️ Exportado: {out_path.as_posix()}")
                print(f"🖼️ Mídias em: {MEDIA_DIR.resolve().as_posix()}")
                running = False
                break

            elif act == "sair":
                running = False
                break

            time.sleep(0.05)

    try: painel.root.destroy()
    except Exception: pass

if __name__ == "__main__":
    # [AJUSTE #4] Geração de log APENAS se ocorrer erro
    try:
        main()
    except Exception:
        try:
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            ensure_dir(LOGS_DIR)  # cria só agora, no erro
            log_path = LOGS_DIR / f"erro-{ts}.txt"
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            print(f"\n❌ Ocorreu um erro. Um log foi salvo em: {log_path.as_posix()}")
        except Exception:
            pass
        # Opcionalmente, não relança para o .exe não encerrar com stack trace
        # raise
