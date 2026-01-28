import time
import re
from selenium.webdriver.common.by import By
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 INICIANDO MODO DE TESTE (LÓGICA AVANÇADA + DEBUG)...")

def nova_logica_questionario(self) -> dict:
    """
    Extração do Questionário com varredura DOM em ordem document (proximidade) e debug.
    Melhorias:
      - procura label associado preferencialmente DENTRO do container da pergunta (não global)
      - fallback ordenado: container.querySelector -> input.closest('label') -> sibling -> document.querySelector
      - imprime RAW JS para depuração
    """
    if not self.clicar_aba('Questionário'):
        return {}

    time.sleep(0.6)  # espera rápida para render

    mapa = {
        'responsável pela unidade': '1. A mulher é a responsável pela unidade familiar?',
        'pessoa negra': '2. Há pessoa negra na composição familiar?',
        'deficiência': '3. Há pessoa com deficiência na composição familiar, comprovada por avaliação biopsicossocial (Lei nº 13.146/2015 e Decreto nº 11.063/2022)?',
        'idoso': '4. Há idoso na composição familiar, comprovado por documento civil com data de nascimento?',
        'criança': '5. Há criança ou adolescente na composição familiar, comprovado por certidão de nascimento, guarda ou tutela?',
        'câncer': '6. Há pessoa com câncer ou doença rara crônica e degenerativa na família, comprovado por laudo médico?',
        'violência': '7. Há mulheres vítimas de violência doméstica/familiar na família, comprovado por registro no Cadastro Nacional de Violência Doméstica (Lei Maria da Penha)?',
        'indígenas': '8. Há integrantes de povos indígenas ou quilombolas na família, declarados no CadÚnico?',
        'área de risco': '9. A família reside em área de risco (deslizamentos, inundações etc.), conforme mapeamento do PMRR, CPRM ou Defesa Civil?',
        'distratado': '10. O beneficiário teve contrato distratado ou rescindido involuntariamente, conforme normativo do Ente Público?',
        'socioassistenciais': '11. Atualmente é atendido pelas redes Socioassistenciais do Município?'
    }

    # JS que faz a varredura DOM a partir do nó da pergunta (document order traversal)
    script_js = r"""
    const mapa = arguments[0];

    function clean(s){ return (s||'').trim().toLowerCase(); }

    function nextNode(node) {
        if (!node) return null;
        if (node.firstElementChild) return node.firstElementChild;
        let n = node;
        while (n) {
            if (n.nextElementSibling) return n.nextElementSibling;
            n = n.parentElement;
        }
        return null;
    }

    function elementVisible(el){
        try {
            const rect = el.getBoundingClientRect();
            return !(rect.width === 0 && rect.height === 0);
        } catch(e){ return true; }
    }

    function findLocalLabelForInput(container, inp) {
        // Tentativas ordenadas de achar label relacionalmente dentro do container/ancestors
        try {
            const id = inp.id || '';
            if (id) {
                // 1) label within container
                try {
                    const lab = container.querySelector("label[for='"+id+"']");
                    if (lab) return lab;
                } catch(e){}
                // 2) input.closest('label')
                try {
                    const lab2 = inp.closest('label');
                    if (lab2) return lab2;
                } catch(e){}
                // 3) sibling label (next/previous)
                try {
                    let s = inp.nextElementSibling;
                    if (s && s.tagName.toLowerCase()==='label') return s;
                    s = inp.previousElementSibling;
                    if (s && s.tagName.toLowerCase()==='label') return s;
                } catch(e){}
                // 4) label inside the closest form or question ancestor
                try {
                    const ancestor = container.closest('form') || container.closest('.question') || container;
                    if (ancestor) {
                        const lab3 = ancestor.querySelector("label[for='"+id+"']");
                        if (lab3) return lab3;
                    }
                } catch(e){}
                // 5) global fallback (last resort)
                try {
                    return document.querySelector("label[for='"+id+"']");
                } catch(e){}
            }
        } catch(e){}
        return null;
    }

    function detectInNode(n) {
        try {
            // 1) inputs (radio/checkbox) near the node
            const inputs = Array.from(n.querySelectorAll("input[type='radio'], input[type='checkbox']"));
            for (let inp of inputs){
                try {
                    if (inp.checked){
                        // try find label text LOCAL to the container
                        const lab = findLocalLabelForInput(n, inp);
                        if (lab) {
                            // ensure label is within same question/container (or at least visible near)
                            if (n.contains(lab) || (lab.closest('.question') && lab.closest('.question') === n.closest('.question'))) {
                                return {method:'input_checked_label', text:clean(lab.innerText || lab.textContent), html:lab.outerHTML};
                            }
                        }
                        // sibling next label
                        try {
                            const sib = inp.nextElementSibling;
                            if (sib && clean(sib.innerText)) return {method:'input_checked_sibling', text:clean(sib.innerText), html:sib.outerHTML};
                        } catch(e){}
                        return {method:'input_checked', text:'sim', html:inp.outerHTML};
                    }
                    const ar = (inp.getAttribute('aria-checked')||'').toString().toLowerCase();
                    const ch = (inp.getAttribute('checked')||'').toString().toLowerCase();
                    if (ar==='true' || ch==='checked' || ch==='true') {
                        const lab = findLocalLabelForInput(n, inp);
                        if (lab && (n.contains(lab) || (lab.closest('.question') && lab.closest('.question') === n.closest('.question')))) {
                            return {method:'input_attr_checked', text:clean(lab.innerText || lab.textContent), html:lab.outerHTML};
                        }
                        try {
                            const sib = inp.nextElementSibling;
                            if (sib && clean(sib.innerText)) return {method:'input_attr_checked_sibling', text:clean(sib.innerText), html:sib.outerHTML};
                        } catch(e){}
                        return {method:'input_attr_checked', text:'sim', html:inp.outerHTML};
                    }
                } catch(e){}
            }

            // 2) look for elements that look like option buttons and check active classes
            const candidates = Array.from(n.querySelectorAll('button, label, .btn, .v-btn, .option, .radio, .choice, .option-item'));
            if (candidates.length >= 1){
                let textMap = [];
                for (let c of candidates){
                    try {
                        const t = clean(c.innerText || c.textContent || '');
                        textMap.push({el:c, text:t, cls:(c.className||'').toString().toLowerCase()});
                    } catch(e){}
                }
                // find active one inside the container only
                for (let it of textMap){
                    if (!n.contains(it.el)) continue;
                    if (it.cls.includes('active') || it.cls.includes('selected') || it.cls.includes('checked') || it.cls.includes('is-checked') || it.cls.includes('btn--active') || it.cls.includes('v-btn--is-active') || it.cls.includes('bg-primary') ){
                        return {method:'class_active', text:it.text || 'sim', html:it.el.outerHTML};
                    }
                }
                // if no active, but there are exactly two with 'sim' and 'não' texts, attempt icon-based detection within container
                const simNao = textMap.filter(x => n.contains(x.el) && (x.text==='sim' || x.text==='não' || x.text==='nao' || x.text==='s' || x.text==='n'));
                if (simNao.length === 2){
                    for (let it of simNao){
                        try {
                            const icon = it.el.querySelector('i, svg, span');
                            if (icon){
                                const cls = (icon.className||'').toString().toLowerCase();
                                if (cls.includes('checked') || cls.includes('selected') || cls.includes('ri-checkbox-circle-fill') || cls.includes('v-icon--active')) {
                                    return {method:'icon_checked', text:it.text, html:it.el.outerHTML};
                                }
                            }
                        } catch(e){}
                    }
                }
            }

            // 3) find explicit 'sim' or 'não' text nodes inside n
            try {
                const bodytxt = clean(n.innerText || n.textContent || '');
                const m = bodytxt.match(/\b(sim|s|não|nao|n)\b/i);
                if (m) return {method:'text_near', text:m[1].toLowerCase(), html:n.outerHTML};
            } catch(e){}
        } catch(e){}
        return null;
    }

    function findContainers(chave) {
        const nodes = Array.from(document.querySelectorAll('div, section, fieldset, li, form, article, .question, .row, .form-group'));
        const matches = [];
        for (let n of nodes){
            try {
                if (!elementVisible(n)) continue;
                const t = clean(n.innerText || n.textContent || '');
                if (t && t.includes(chave)) matches.push(n);
            } catch(e){}
        }
        return matches;
    }

    const results = {};
    for (const chave in mapa){
        try {
            const containers = findContainers(chave);
            let found = null;
            for (let c of containers){
                const d = detectInNode(c);
                if (d){ found = d; break; }
                // scan forward a few nodes in document order
                let cur = c;
                for (let i=0;i<40 && cur;i++){
                    cur = nextNode(cur);
                    if (!cur) break;
                    try {
                        if (!elementVisible(cur)) continue;
                        const d2 = detectInNode(cur);
                        if (d2){ found = d2; break; }
                    } catch(e){}
                }
                if (found) break;
            }
            if (found) {
                const txt = (found.text||'').toString().trim().toLowerCase();
                let norm = found.text;
                if (/^s($|im)/.test(txt)) norm = 'Sim';
                else if (/^n($|ão|ao)/.test(txt)) norm = 'Não';
                results[mapa[chave]] = {answer: norm, method: found.method, snippet: (found.html||'').slice(0,800)};
            } else {
                results[mapa[chave]] = {answer: null, method: 'not_found', snippet: '', debug: 'no containers matched or no markers'};
            }
        } catch(e){
            results[mapa[chave]] = {answer: null, method: 'exception', snippet: '', debug: String(e)};
        }
    }
    return results;
    """

    def run_js_context():
        try:
            return self.driver.execute_script(script_js, mapa) or {}
        except Exception as e:
            print("JS exec error:", e)
            return {}

    # 1) tenta no contexto atual
    raw = run_js_context()

    # 2) se vazio / tudo null, tenta em iframes
    any_found = any((v and v.get('answer')) for v in raw.values()) if isinstance(raw, dict) else False
    if not any_found:
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        except:
            iframes = []
        for f in iframes:
            try:
                self.driver.switch_to.frame(f)
                time.sleep(0.15)
                raw = run_js_context()
                try:
                    self.driver.switch_to.default_content()
                except:
                    pass
                any_found = any((v and v.get('answer')) for v in raw.values()) if isinstance(raw, dict) else False
                if any_found:
                    break
            except Exception:
                try:
                    self.driver.switch_to.default_content()
                except:
                    pass
                continue

    # Imprime o resultado cru para debug
    print("\n--- RAW JS Questionário Results ---")
    try:
        import json
        print(json.dumps(raw, ensure_ascii=False, indent=2) if raw else raw)
    except Exception:
        print(raw)
    print("--- END RAW ---\n")

    # Normaliza e preenche saida compatível com COL_SPECS
    saida = {}
    try:
        respostas_agregadas = []
        for pergunta_full, info in (raw or {}).items():
            try:
                resp = None
                if isinstance(info, dict):
                    resp = info.get('answer')
                else:
                    resp = info
                if resp is None:
                    continue
                r = str(resp).strip()
                if r.lower() in ('s', 'sim'):
                    rnorm = 'Sim'
                elif r.lower() in ('n', 'não', 'nao'):
                    rnorm = 'Não'
                else:
                    rnorm = r.capitalize()
                chave_norm = _norm_texto_chave(pergunta_full)
                saida[chave_norm] = rnorm
                respostas_agregadas.append(f"{pergunta_full} -> {rnorm}")
                for p_schema in COL_SPECS:
                    try:
                        if p_schema.aba == 'Questionário' and p_schema.campo == pergunta_full:
                            saida[p_schema.key] = rnorm
                            break
                    except:
                        continue
            except Exception as e:
                print("Erro normalizando item:", pergunta_full, e)
        if respostas_agregadas:
            saida['Respostas do Questionário'] = "\n".join(respostas_agregadas)
    except Exception as e:
        print("Erro ao montar saida:", e)

    # garante voltar do iframe
    try:
        self.driver.switch_to.default_content()
    except:
        pass

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
            except:
                pass

        print("   -> LENDO QUESTIONÁRIO (MÉTODO AVANÇADO COM DEBUG)...")
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        extrator = ExtratorHabibot(self.driver, self.wait, debug=self.debug, debug_visual=self.debug_visual)
        qmap = extrator.extrair_questionario_mapa() or {}
        row.update(qmap)
        return row

if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE - LÓGICA AVANÇADA")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)