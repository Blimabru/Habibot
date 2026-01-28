import time
import json
from selenium.webdriver.common.by import By
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 INICIANDO MODO DE TESTE (EXTRATOR DETERMINÍSTICO)...")

def nova_logica_questionario(self) -> dict:
    """
    Extrai cada div.question de forma determinística:
      - pergunta: h6/h5/text dentro do div.question
      - resposta: input[type=radio|checkbox]:checked dentro do mesmo div (preferível)
      - label: label[for=id] dentro do mesmo div (preferível) ou sibling
    Retorna mapa normalizado (chave -> 'Sim'|'Não' ou texto).
    """
    if not self.clicar_aba('Questionário'):
        return {}
    time.sleep(0.6)

    js = r"""
    const out = [];
    const questions = Array.from(document.querySelectorAll('div.question, .question'));
    function textOf(el){
        try { return (el && (el.innerText||el.textContent||'')).toString().trim(); } catch(e){ return ''; }
    }
    for (const q of questions){
        try {
            // pergunta
            let heading = '';
            const h = q.querySelector('h6,h5,h4,label');
            heading = textOf(h) || textOf(q).split('\n')[0] || '';
            // tenta achar input marcado DENTRO do container
            let checked = q.querySelector("input[type='radio']:checked, input[type='checkbox']:checked");
            let chosen = null;
            if (checked){
                const id = checked.id || '';
                let label = null;
                if (id) label = q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']");
                if (!label) label = checked.closest('label');
                if (!label) {
                    let sib = checked.nextElementSibling;
                    if (sib && (sib.tagName||'').toLowerCase() === 'label') label = sib;
                }
                chosen = {id: id, value: checked.value || '', label: label ? textOf(label) : '', labelHTML: label ? label.outerHTML : '', method: 'inside_checked'};
            } else {
                // se não tem checked dentro do container, tenta inspecionar inputs dentro do container procurando por property 'checked' (fallback)
                const inputs = Array.from(q.querySelectorAll("input[type='radio'], input[type='checkbox']"));
                for (const inp of inputs){
                    try {
                        if (inp.checked){
                            const id = inp.id || '';
                            let label = id ? (q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']")) : null;
                            if (!label) label = inp.closest('label');
                            if (!label) { let sib = inp.nextElementSibling; if (sib && (sib.tagName||'').toLowerCase()==='label') label = sib; }
                            chosen = {id: id, value: inp.value || '', label: label ? textOf(label) : '', labelHTML: label ? label.outerHTML : '', method: 'inside_property_checked'};
                            break;
                        }
                    } catch(e){}
                }
                // se ainda não, tenta input[name]:checked global mas garante que label belongs to this question
                if (!chosen){
                    const names = [...new Set(inputs.map(i=>i.name).filter(n=>n))];
                    for (const nm of names){
                        try {
                            const g = document.querySelector("input[name='"+nm+"']:checked");
                            if (g){
                                const id = g.id || '';
                                let label = id ? (q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']")) : null;
                                // verify label is inside this question or belongs to same question heading text
                                const ok = label && (q.contains(label) || (label.closest('.question') && label.closest('.question') === q));
                                if (ok){
                                    chosen = {id:id, value: g.value || '', label: textOf(label), labelHTML: label.outerHTML, method: 'global_checked_but_local_label'};
                                    break;
                                } else {
                                    // fallback: accept global checked but attach its label text if found
                                    if (id){
                                        const labGlobal = document.querySelector("label[for='"+id+"']");
                                        if (labGlobal) {
                                            chosen = {id:id, value:g.value||'', label:textOf(labGlobal), labelHTML: labGlobal.outerHTML, method:'global_checked_fallback'};
                                            break;
                                        }
                                    }
                                }
                            }
                        } catch(e){}
                    }
                }
            }
            // se ainda não achou, tenta deduzir pelos labels dentro do mesmo question e ver se um tem classe active
            if (!chosen){
                const labs = Array.from(q.querySelectorAll('label'));
                for (const L of labs){
                    const cls = (L.className||'').toString().toLowerCase();
                    if (cls.includes('active') || cls.includes('selected') || cls.includes('is-checked') || cls.includes('checked')){
                        chosen = {id:'', value:'', label:textOf(L), labelHTML:L.outerHTML, method:'label_class_active'};
                        break;
                    }
                }
            }

            out.push({heading: heading, chosen: chosen, rawLabels: Array.from(q.querySelectorAll('label')).map(l=>({text:textOf(l), html:l.outerHTML.slice(0,400)}))});
        } catch(e){}
    }
    return out;
    """

    try:
        collected = self.driver.execute_script(js)
    except Exception as e:
        print("Erro ao executar JS:", e)
        collected = []

    # imprime RAW para inspeção
    print("\n--- RAW COLLECTED ---")
    print(json.dumps(collected, ensure_ascii=False, indent=2))
    print("--- END RAW ---\n")

    # Monta mapa normalizado
    saida = {}
    respostas_agregadas = []
    for item in (collected or []):
        heading = item.get('heading') or ''
        chosen = item.get('chosen')
        resp_text = ''
        metodo = ''
        if chosen and chosen.get('label'):
            resp_text = chosen.get('label')
            metodo = chosen.get('method')
        elif chosen and chosen.get('value'):
            resp_text = chosen.get('value')
            metodo = chosen.get('method')
        else:
            # se não detectou, tenta checar rawLabels se existe label com inner 'Sim' e sibling input checked by name mapping:
            rawLabs = item.get('rawLabels') or []
            if rawLabs:
                # fallback: if first label text == 'Sim' and there is metadata in forward_neighbors maybe use positional, but prefer empty
                resp_text = ''
                metodo = 'not_detected'
            else:
                resp_text = ''
                metodo = 'not_detected'

        # normaliza Sim/Não
        rnorm = ''
        if resp_text:
            rl = resp_text.strip().lower()
            if rl in ('s','sim'):
                rnorm = 'Sim'
            elif rl in ('n','não','nao'):
                rnorm = 'Não'
            else:
                rnorm = resp_text.strip().capitalize()

        chave = _norm_texto_chave(heading or '')
        if chave:
            saida[chave] = rnorm
            respostas_agregadas.append(f"{heading} -> {rnorm} ({metodo})")
            # preenche valor na chave do schema também quando possível
            for p_schema in COL_SPECS:
                try:
                    if p_schema.aba == 'Questionário' and p_schema.campo == heading:
                        saida[p_schema.key] = rnorm
                        break
                except:
                    pass

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
            except:
                pass

        print("   -> LENDO QUESTIONÁRIO (DETERMINÍSTICO)...")
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        extrator = ExtratorHabibot(self.driver, self.wait, debug=self.debug, debug_visual=self.debug_visual)
        qmap = extrator.extrair_questionario_mapa() or {}
        row.update(qmap)
        return row

if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE - DETERMINÍSTICO")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)