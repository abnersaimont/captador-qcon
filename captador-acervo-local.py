# scrape_qconcursos_brave_marcadores_gui_cdp.py
# GUI sem barra de título + alça para arrastar; contadores em tempo real; reinjeção pós-filtro
from pathlib import Path
import json, os, re, time, subprocess, queue, socket
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont  # [NOVO] para ajustar fonte das labels
# [NOVO Grupo] askstring para solicitar o nome do lote
from tkinter import simpledialog

# =============== CONFIG ===============
BRAVE_PATH = r"C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe"

# Perfil de automação separado (não conflita com seu Brave do dia a dia):
AUTOMATION_USER_DATA_DIR = r"C:/Users/Abner/AppData/Local/BraveSoftware/Playwright-Automation"

HOME_URL = "https://www.qconcursos.com/questoes-de-concursos/questoes"
LOGIN_CHECK_SELECTOR = "body[data-is-user-signed-in='true']"

# CDP
DEVTOOLS_HOST = "127.0.0.1"  # evita ::1
DEVTOOLS_PORT = 9222
DEVTOOLS_URL = f"http://{DEVTOOLS_HOST}:{DEVTOOLS_PORT}"

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

MEDIA_DIR = Path("midia")

# =============== UTILS (iguais) ===============
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

def start_brave_for_cdp():
    """Abre um Brave isolado com porta CDP ativa, sem conflitar com outras janelas do Brave."""
    os.makedirs(AUTOMATION_USER_DATA_DIR, exist_ok=True)

    # Porta já está de pé?
    try:
        s = socket.create_connection((DEVTOOLS_HOST, DEVTOOLS_PORT), timeout=0.3)
        s.close()
        return
    except Exception:
        pass

    # Lançar Brave com user-data-dir dedicado
    cmd = [
        BRAVE_PATH,
        f"--remote-debugging-port={DEVTOOLS_PORT}",
        f"--user-data-dir={AUTOMATION_USER_DATA_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
    ]
    subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, close_fds=True)

    # Aguardar a porta subir
    deadline = time.time() + 20.0
    while time.time() < deadline:
        try:
            s = socket.create_connection((DEVTOOLS_HOST, DEVTOOLS_PORT), timeout=0.5)
            s.close()
            return
        except Exception:
            time.sleep(0.25)
    raise RuntimeError(f"Brave com CDP não subiu na porta {DEVTOOLS_PORT} em {DEVTOOLS_HOST}.")

# ========= GUI (sem barra de título + ALÇA) =========
class Painel:
    def __init__(self):
        self.actions = queue.Queue()
        self.root = tk.Tk()
        self.root.title("")
        self.root.overrideredirect(True)     # remove barra superior
        self.root.attributes("-topmost", True)
        self.root.resizable(False, False)

        frm = ttk.Frame(self.root, padding=8)
        frm.grid(row=0, column=0, sticky="nsew")

        # ---- ALÇA DE ARRASTAR (grip) ----
        grip = ttk.Frame(frm, height=10)
        grip.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0,6))
        grip.configure(cursor="fleur")
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
        self.btn_encerrar = ttk.Button(frm, text="Encerrar", command=lambda: self.actions.put("encerrar"))
        self.btn_filtrar.grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        self.btn_captar.grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self.btn_encerrar.grid(row=1, column=2, padx=4, pady=4, sticky="ew")

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

        # ---- separador abaixo dos contadores ----
        ttk.Separator(frm, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(6,2)
        )

        # ---- texto da empresa ----
        self.lbl_empresa = ttk.Label(frm, text="Beckm", anchor="e")
        self.lbl_empresa.grid(row=5, column=0, columnspan=3, sticky="e")
        try:
            _small = tkfont.nametofont("TkDefaultFont").copy()
            _small.configure(size=5)
            self.lbl_empresa.configure(font=_small)
        except Exception:
            pass

        for i in range(3):
            frm.columnconfigure(i, weight=1)

        # posição inicial: canto inferior direito
        try:
            self.root.update_idletasks()
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, sw - w - PANEL_MARGIN_X)
            y = max(0, sh - h - PANEL_MARGIN_Y)
            self.root.geometry(f"+{x}+{y}")
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
    Procura dentro do card por '.q-question-text--print-hide a.q-link'.
    - Se existir e estiver colapsado (aria-expanded != 'true'), clica para expandir.
    - Identifica o container pelo href (#question-XXXX-text) ou pelo próximo .collapse.
    - Serializa texto e imagens; baixa imagens (HTTP direto ou screenshot de fallback).
    Retorna: (texto_assoc, assoc_imagem_paths, assoc_imagens_info_out)
    """
    texto_assoc = ""
    assoc_imagem_paths = []
    assoc_imagens_info_out = []

    try:
        link = card.query_selector(".q-question-text--print-hide a.q-link")
        if not link:
            return "", [], []

        # Expandir se necessário
        try:
            expanded = (link.get_attribute("aria-expanded") or "").lower() == "true"
        except Exception:
            expanded = False

        if not expanded:
            try:
                link.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                link.click()
            except Exception:
                try:
                    link.evaluate("el => el.click()")
                except Exception:
                    pass
            time.sleep(0.15)

        # Achar container
        target_sel = link.get_attribute("href") or ""
        content = None
        if target_sel.startswith("#"):
            try:
                content = card.query_selector(target_sel)
            except Exception:
                content = None
        if not content:
            try:
                content = card.evaluate_handle("""
                    (root, link) => {
                      const sib = link.nextElementSibling;
                      if (sib && sib.classList && sib.classList.contains('collapse')) return sib;
                      return null;
                    }
                """, link).as_element()
            except Exception:
                content = None

        if not content:
            return "", [], []

        # Serializar texto + imagens
        try:
            res = serialize_node_with_markers(content)
            texto_assoc = (res.get("text") or "").strip()
            imagens_meta = res.get("images") or []
        except Exception:
            try:
                texto_assoc = clean(content.inner_text() or "")
                imagens_meta = []
            except Exception:
                imagens_meta = []
                texto_assoc = ""

        # Baixar imagens do texto associado
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
                    img_el = content.query_selector(f"img[src='{src}']")
                    if img_el:
                        img_el.scroll_into_view_if_needed()
                        img_el.screenshot(path=dest.as_posix())
                        saved = True
                except Exception:
                    pass
            if saved:
                path_str = dest.resolve().as_posix()
                assoc_imagem_paths.append(path_str)
                assoc_imagens_info_out.append({
                    "url": src,
                    "alt": alt,
                    "placeholder": placeholder,
                    "width": width,
                    "height": height,
                    "path": path_str,
                    "scope": "texto_associado"
                })

    except Exception:
        pass

    return texto_assoc, assoc_imagem_paths, assoc_imagens_info_out

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

# ========= Principal (CDP + perfil de automação) =========
def main():
    print(f"Automação Qconcursos")

    dados, vistos = [], set()
    vistos_qid = set()  # set de QIDs já captados
    total = 0
    painel = Painel()

    with sync_playwright() as p:
        start_brave_for_cdp()  # garante uma instância isolada com a porta aberta

        # Conecta ao Brave via CDP
        browser = p.chromium.connect_over_cdp(DEVTOOLS_URL)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()

        # Atualiza contagem na UI e habilita/desabilita "Captar"
        page.expose_function("pyUpdateSelectionCount", lambda n: painel.update_selected(int(n)))

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
                now = time.time()
                if now - last_poll > 0.08:
                    try:
                        n = page.evaluate("() => document.querySelectorAll('.qc-checkbox:checked').length")
                        painel.update_selected(int(n))
                    except Exception:
                        pass
                    last_poll = now

            if act == "filtrar_ok":
                painel.set_filtrar_enabled(False)

                try:
                    page.wait_for_selector(SELECTORS["card"], timeout=20000)
                except PWTimeoutError:
                    print("⏳ Nenhum card encontrado; ajuste filtros.")
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

                # [NOVO Grupo] solicitar o nome do grupo/lote para este clique em "Captar"
                grupo_nome = simpledialog.askstring(
                    "Nome do grupo",
                    "Digite um nome para este grupo de questões:",
                    parent=painel.root
                )
                if grupo_nome is None:
                    # usuário cancelou: não captura nada neste clique
                    print("ℹ️ Captura cancelada — nenhum nome informado.")
                    continue
                grupo_nome = grupo_nome.strip()
                # aceita vazio: grava como string vazia; se preferir, poderia usar um default

                try:
                    cards = page.query_selector_all(SELECTORS["card"])
                except Exception:
                    cards = []
                alvo = [cards[i] for i in selecionados if 0 <= i < len(cards)]

                for card in alvo:
                    try: card.scroll_into_view_if_needed()
                    except Exception: pass

                    # pega QID e pula se já visto
                    qid_real = get_qid_from_card(card) or ""
                    if qid_real and qid_real in vistos_qid:
                        continue

                    ano = get_text_or_none(card, SELECTORS["ano"])
                    banca = get_text_or_none(card, SELECTORS["banca"])

                    enun_el = card.query_selector(SELECTORS["enunciado"])
                    enunciado_text = ""; images_meta_enun = []
                    if enun_el:
                        try:
                            res = serialize_node_with_markers(enun_el); enunciado_text = res.get("text") or ""; images_meta_enun = res.get("images") or []
                        except Exception:
                            enunciado_text = clean(enun_el.inner_text())

                    alt_texts = ["", "", "", "", ""]; alt_images_meta = [[], [], [], [], []]
                    try:
                        alt_nodes = card.query_selector_all(SELECTORS["alternativas"])
                        for idx, node in enumerate(alt_nodes[:5]):
                            try:
                                res = serialize_node_with_markers(node)
                                alt_texts[idx] = res.get("text") or ""
                                alt_images_meta[idx] = res.get("images") or []
                            except Exception:
                                alt_texts[idx] = clean(node.inner_text()); alt_images_meta[idx] = []
                    except Exception: pass
                    altA, altB, altC, altD, altE = alt_texts

                    gabarito_raw = ""
                    try:
                        nodeA = card.query_selector(SELECTORS["altA_click"])
                        if nodeA: nodeA.scroll_into_view_if_needed(); nodeA.click()
                        btn = card.query_selector(SELECTORS["btn_responder"])
                        if btn:
                            btn.scroll_into_view_if_needed(); btn.click()
                            try: card.wait_for_selector(SELECTORS["gabarito_after_click"], timeout=4000)
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
                        if src.startswith("http"): saved = save_binary(context, src, dest)
                        if not saved:
                            try:
                                if enun_el:
                                    img_el = enun_el.query_selector(f"img[src='{src}']")
                                    if img_el: img_el.scroll_into_view_if_needed(); img_el.screenshot(path=dest.as_posix()); saved = True
                            except Exception: pass
                        if saved:
                            path_str = dest.resolve().as_posix()
                            imagem_paths.append(path_str)
                            if placeholder: mapa_imagens[placeholder] = path_str
                            imagens_info_out.append({"url":src,"alt":alt,"placeholder":placeholder,"width":width,"height":height,"path":path_str,"scope":"enunciado"})

                    # ===== Texto associado =====
                    texto_assoc, assoc_imagem_paths, assoc_imagens_info = revelar_e_capturar_texto_associado(
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
                                        img_el = alt_node.query_selector(f"img[src='{src}']")
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
                        # Campos do Texto associado
                        "TextoAssociado": texto_assoc or "",
                        "TextoAssociadoImagens": assoc_imagem_paths,
                        "TextoAssociadoMapaImagens": [{"path": p, "scope": "texto_associado"} for p in assoc_imagem_paths],
                        "TextoAssociadoImagensInfo": assoc_imagens_info,
                        # inclui QID no JSON
                        "QID": qid_real,
                        # [NOVO Grupo] nome do lote/grupo informado neste clique
                        "Grupo": grupo_nome,
                    }
                    chave = (registro["Enunciado"], registro["Banca"], registro["Ano"])
                    if chave not in vistos:
                        vistos.add(chave)
                        if qid_real:
                            vistos_qid.add(str(qid_real))
                        dados.append(registro)
                        total += 1
                        painel.update_total(total)

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

            elif act == "encerrar":
                # Se nada foi captado, NÃO gera JSON
                if total == 0 or not dados:
                    print("\nℹ️ Nenhuma questão captada — nenhum arquivo foi gerado.")
                    print(f"🖼️ Pasta de mídias (se criada): {MEDIA_DIR.resolve().as_posix()}")
                    running = False
                    break

                downloads = Path(os.path.join(os.path.expanduser("~"), "Downloads"))
                ensure_dir(downloads)
                out_path = downloads / "Questões.json"
                if out_path.exists():
                    i = 2
                    while True:
                        cand = downloads / f"Questões ({i}).json"
                        if not cand.exists(): out_path = cand; break
                        i += 1
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(dados, f, ensure_ascii=False, indent=2)
                print(f"\n✅ Finalizado. Questões coletadas: {total}")
                print(f"🗂️ Exportado: {out_path.as_posix()}")
                print(f"🖼️ Mídias em: {MEDIA_DIR.resolve().as_posix()}")
                running = False
                break

            time.sleep(0.03)

    try: painel.root.destroy()
    except Exception: pass

if __name__ == "__main__":
    main()
