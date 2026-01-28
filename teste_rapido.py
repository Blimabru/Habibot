import time
import json
from selenium.webdriver.common.by import By
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 INICIANDO MODO DE TESTE (INSPEÇÃO DETALHADA DO QUESTIONÁRIO)...")

def nova_logica_questionario(self) -> dict:
    """
    Varre cada div.question e coleta dados brutos sobre inputs/labels para diagnóstico.
    Retorna mapa normalizado pergunta -> resposta (com heurística) e imprime o RAW detalhado.
    """
    if not self.clicar_aba('Questionário'):
        return {}

    time.sleep(0.6)

    # JS que coleta evidências por pergunta (executado no browser)
    script_collect = r"""
    const out = [];
    // seleciona containers de pergunta (afine se necessário)
    const questions = Array.from(document.querySelectorAll('div.question, .question'));
    function clean(s){ return (s||'').toString().trim(); }
    for (let q of questions){
        try {
            // pergunta (tenta h6/h5/label/text)
            let title = '';
            try { title = clean((q.querySelector('h6')||q.querySelector('h5')||q.querySelector('h4')||q.querySelector('label')||q).innerText); } catch(e){}
            // inputs radios/checkboxes dentro do container
            const inputs = Array.from(q.querySelectorAll("input[type='radio'], input[type='checkbox']"));
            const options = [];
            for (let inp of inputs){
                try {
                    const id = inp.id || '';
                    const name = inp.name || '';
                    const checked_prop = !!inp.checked;
                    const attr_checked = (inp.getAttribute('checked')||'')+'';
                    const aria_checked = (inp.getAttribute('aria-checked')||'')+'';
                    const disabled = !!inp.disabled;
                    // find label local: label[for=id] within q, else closest label ancestor, else nextElementSibling label
                    let label_el = null;
                    try { if (id) label_el = q.querySelector("label[for='"+id+"']"); } catch(e){}
                    try { if (!label_el) label_el = inp.closest('label'); } catch(e){}
                    try { if (!label_el) { let s = inp.nextElementSibling; if (s && s.tagName && s.tagName.toLowerCase()==='label') label_el = s; } } catch(e){}
                    let label_text = label_el ? clean(label_el.innerText) : '';
                    let label_html = label_el ? label_el.outerHTML.slice(0,1000) : '';
                    let label_class = label_el ? (label_el.className || '') : '';
                    // computed styles (color/background) of label
                    let label_style = {};
                    try { if (label_el) { const cs = window.getComputedStyle(label_el); label_style.color = cs.color; label_style.background = cs.backgroundColor; label_style.fontWeight = cs.fontWeight; } } catch(e){}
                    // also capture inner icons/html to detect svg/i markers
                    let inner_html = '';
                    try { inner_html = label_el ? label_el.innerHTML.slice(0,800) : ''; } catch(e){}
                    options.push({
                        id, name, checked_prop, attr_checked, aria_checked, disabled,
                        label_text, label_class, label_html, label_style, inner_html
                    });
                } catch(e){}
            }

            // group-level: find input[name]:checked in the document for each distinct name found
            const group_checked = {};
            try {
                const names = [...new Set(inputs.map(i => i.name).filter(n=>n))];
                for (let nm of names){
                    const sel = document.querySelector("input[name='"+nm+"']:checked");
                    if (sel) group_checked[nm] = {id: sel.id || '', outerHTML: (sel.outerHTML||'').slice(0,600)};
                }
            } catch(e){}

            // also scan forward in document order within a small window for elements that indicate selection
            function scan_forward(node, steps){
                let cur = node;
                let acc = [];
                for (let i=0;i<steps;i++){
                    if (!cur) break;
                    if (cur.firstElementChild) cur = cur.firstElementChild;
                    else {
                        while (cur && !cur.nextElementSibling) cur = cur.parentElement;
                        if (!cur) break;
                        cur = cur.nextElementSibling;
                    }
                    if (!cur) break;
                    try {
                        const text = clean(cur.innerText||cur.textContent||'');
                        if (text) acc.push({tag: cur.tagName, text: text.slice(0,200), outer: (cur.outerHTML||'').slice(0,400)});
                    } catch(e){}
                }
                return acc;
            }

            out.push({title, options, group_checked, forward_neighbors: scan_forward(q, 10)});
        } catch(e){}
    }
    return out;
    """

    raw = []
    try:
        raw = self.driver.execute_script(script_collect)
    except Exception as e:
        print("Erro exec JS collect:", e)
        raw = []

    # Imprime RAW detalhado para você colar aqui
    print("\n--- RAW PERGUNTAS DETALHADO ---")
    print(json.dumps(raw, ensure_ascii=False, indent=2))
    print("--- END RAW ---\n")

    # Heurística de decisão (prioritária)
    saida = {}
    respostas_agregadas = []
    for q in raw:
        try:
            pergunta = q.get('title') or ''
            if not pergunta:
                # tenta extrair a partir do texto primeiro das neighbors
                pergunta = (q.get('forward_neighbors') and q['forward_neighbors'][0]['text']) or pergunta
            escolhido = None
            metodo = None

            # 1) procura option com checked_prop True
            for opt in q.get('options', []):
                if opt.get('checked_prop'):
                    escolhido = opt.get('label_text') or 'Sim'
                    metodo = 'checked_prop'
                    break

            # 2) se não, procura group_checked (input[name]:checked)
            if not escolhido:
                gch = q.get('group_checked') or {}
                if gch:
                    # pega o primeiro id marcado no grupo e localiza na options
                    for nm, info in gch.items():
                        gid = info.get('id')
                        if not gid: continue
                        for opt in q.get('options', []):
                            if opt.get('id') == gid:
                                escolhido = opt.get('label_text') or 'Sim'
                                metodo = 'group_checked'
                                break
                        if escolhido: break

            # 3) se não, procura label_class contendo active/selected/is-checked
            if not escolhido:
                for opt in q.get('options', []):
                    cls = (opt.get('label_class') or '').lower()
                    if any(k in cls for k in ('active','selected','is-checked','checked','btn--active','v-btn--is-active','bg-primary','text-success')):
                        escolhido = opt.get('label_text') or 'Sim'
                        metodo = 'label_class'
                        break

            # 4) se não, procura ícone/svgs em inner_html com classes típicas de seleção
            if not escolhido:
                for opt in q.get('options', []):
                    ih = (opt.get('inner_html') or '').lower()
                    if any(tok in ih for tok in ('ri-checkbox-circle-fill','ri-checkbox-fill','checked','is-checked','v-icon--active','svg')):
                        escolhido = opt.get('label_text') or 'Sim'
                        metodo = 'icon_marker'
                        break

            # 5) se não, usa cor computada do label (ex.: verde bootstrap = rgb(13, 110, 53) ou similar)
            if not escolhido:
                for opt in q.get('options', []):
                    style = opt.get('label_style') or {}
                    color = (style.get('color') or '').lower()
                    bg = (style.get('background') or '').lower()
                    if color and ('rgb' in color and ('13, 110, 53' in color or 'green' in color or '0, 128, 0' in color)) or ('rgb' in bg and ('13, 110, 53' in bg or 'green' in bg)):
                        escolhido = opt.get('label_text') or 'Sim'
                        metodo = 'computed_color'
                        break

            # 6) fallback: se houver 2 opções e ambas com label_text 'Sim'/'Não', tenta decidir por presence of attr_checked or attr 'aria-checked'
            if not escolhido:
                opts = q.get('options', [])
                if len(opts) >= 2:
                    # procura attr_checked or aria_checked truthy
                    for opt in opts:
                        ac = (opt.get('attr_checked') or '').lower()
                        ar = (opt.get('aria_checked') or '').lower()
                        if ac and ac not in ('', 'null', 'false') or ar in ('true','checked','1','on'):
                            escolhido = opt.get('label_text') or ('Sim' if 'sim' in (opt.get('label_text') or '').lower() else 'Não')
                            metodo = 'attr_checked'
                            break
                # se ainda nada, tenta fallback positional: procura label_text values and if first == 'Sim' else take actual texts
                if not escolhido and len(opts) >= 2:
                    lt0 = (opts[0].get('label_text') or '').lower()
                    lt1 = (opts[1].get('label_text') or '').lower()
                    if lt0 in ('sim','s') and lt1 in ('não','nao','n'):
                        escolhido = 'Sim'
                        metodo = 'positional'
                    elif lt1 in ('sim','s') and lt0 in ('não','nao','n'):
                        escolhido = 'Não'
                        metodo = 'positional'
                    else:
                        # can't decide
                        escolhido = opts[0].get('label_text') or None
                        metodo = 'fallback_first'

            if escolhido:
                # normaliza Sim/Não
                r = escolhido.strip()
                rl = r.lower()
                if rl in ('s','sim'): rnorm = 'Sim'
                elif rl in ('n','não','nao'): rnorm = 'Não'
                else: rnorm = r.capitalize()
                chave_norm = _norm_texto_chave(pergunta)
                saida[chave_norm] = rnorm
                respostas_agregadas.append(f"{pergunta} -> {rnorm} ({metodo})")
            else:
                # nada decidido
                chave_norm = _norm_texto_chave(pergunta)
                saida[chave_norm] = ''
                respostas_agregadas.append(f"{pergunta} -> (não detectado)")
        except Exception as e:
            print("Erro normalizando pergunta:", e)
            continue

    if respostas_agregadas:
        saida['Respostas do Questionário'] = "\n".join(respostas_agregadas)
    return saida

# BOT TESTE
class BotTeste(HabibotBot):
    def extrair_um(self) -> dict:
        extrator = ExtratorHabibot(self.driver, self.wait, debug=self.debug, debug_visual=self.debug_visual)
        row = {}
        print("   -> Lendo Identificação...")
        if extrator.clicar_aba('Titular'):
            try:
                time.sleep(1)
                row[_col_key('Titular', 'Dados Gerais', 'Nome')] = extrator._extrair_por_label_simples('Nome')
                row[_col_key('Titular', 'Dados Gerais', 'CPF/CNPJ')] = extrator._extrair_por_label_simples('CPF/CNPJ')
            except: pass

        print("   -> LENDO QUESTIONÁRIO (INSPEÇÃO DETALHADA)...")
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        extrator = ExtratorHabibot(self.driver, self.wait, debug=self.debug, debug_visual=self.debug_visual)
        qmap = extrator.extrair_questionario_mapa() or {}
        row.update(qmap)
        return row

if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE - INSPEÇÃO")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)