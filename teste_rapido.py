import time
import re
from selenium.webdriver.common.by import By
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 INICIANDO MODO DE TESTE (LÓGICA POSICIONAL MELHORADA)...")

# ==============================================================================
# LÓGICA MELHORADA PARA EXTRACAO DO QUESTIONÁRIO
# ==============================================================================
def nova_logica_questionario(self) -> dict:
    """
    Extração resiliente do Questionário para testes rápidos.

    Estratégia:
      - Clica na aba "Questionário".
      - Tenta extrair no contexto atual (document) usando heurísticas DOM:
          * encontra nós que contenham o trecho-chave (case-insensitive)
          * procura container próximo (fieldset, .form-group, parent)
          * detecta radios (checked / aria-checked / atributo checked)
          * tenta labels por for/id, ancestor label, following-sibling label
          * fallback para botões com classe active/selected
      - Se nada for encontrado no contexto atual, itera por iframes e repete a extração em cada um.
      - Monta o dicionário com chaves normalizadas e também preenche as chaves do schema (COL_SPECS).
      - Retorna dict[normalized_question -> resposta]
    """
    if not self.clicar_aba('Questionário'):
        return {}

    time.sleep(0.6)  # pequena espera para render

    # mapa de busca (trecho curto -> pergunta completa do schema)
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

    saida = {}

    # JS utilitário que roda no contexto atual e retorna dicionário pergunta->resposta (quando possível)
    script_js = """
    const mapa = arguments[0];

    function clean(s) { return (s || '').trim().toLowerCase(); }

    function findContainersContaining(chave) {
        // procura nós sem filhos (leaf nodes) cujo texto contenha a chave
        const nodes = Array.from(document.querySelectorAll('body *'));
        const matches = [];
        for (let n of nodes) {
            try {
                if (!n.offsetParent && n.tagName !== 'LABEL') continue; // elemento não visível
            } catch(e) {}
            const txt = clean(n.innerText || n.textContent || '');
            if (txt && txt.includes(chave)) {
                matches.push(n);
            }
        }
        // para cada nó encontrado, retorna um container razoável (fieldset / .form-group / parent)
        const containers = [];
        for (let n of matches) {
            let c = n.closest('fieldset') || n.closest('.form-group') || n.closest('.row') || n.parentElement;
            if (c) containers.push(c);
            else containers.push(n);
        }
        return containers;
    }

    function detectAnswerInContainer(container) {
        // 1) radios/checkboxes
        const inputs = Array.from(container.querySelectorAll("input[type='radio'], input[type='checkbox']"));
        if (inputs.length >= 1) {
            // prefer checked input
            for (let i=0;i<inputs.length;i++) {
                try {
                    if (inputs[i].checked) {
                        // tenta pegar label associado
                        const id = inputs[i].id || '';
                        if (id) {
                            const lab = container.querySelector("label[for='" + id + "']");
                            if (lab && clean(lab.innerText)) return clean(lab.innerText);
                        }
                        // tenta next sibling label
                        let sib = inputs[i].nextElementSibling;
                        if (sib && clean(sib.innerText)) return clean(sib.innerText);
                        return 'sim';
                    }
                    const ch = (inputs[i].getAttribute('aria-checked') || inputs[i].getAttribute('checked') || '').toString().toLowerCase();
                    if (ch === 'true' || ch === 'checked' || ch === '1' || ch === 'on') {
                        let sib = inputs[i].nextElementSibling;
                        if (sib && clean(sib.innerText)) return clean(sib.innerText);
                        return 'sim';
                    }
                } catch(e){}
            }
            // se nenhum marcado, tenta inferir por posição: primeiro=true, segundo=false
            if (inputs.length >= 2) {
                // tenta ler labels dos dois primeiros
                try {
                    let l0 = inputs[0].nextElementSibling;
                    let l1 = inputs[1].nextElementSibling;
                    const t0 = l0 ? clean(l0.innerText) : '';
                    const t1 = l1 ? clean(l1.innerText) : '';
                    if (t0 && (t0==='sim' || t0==='não' || t0==='nao')) {
                        // se text labels existirem, mas nenhum marcado, não sabemos; retorna null
                    } else {
                        // fallback positional: se inputs[0] existimos consideramos como opção positiva
                        return null;
                    }
                } catch(e){}
            }
        }

        // 2) botões/labels com classes active/selected
        const btns = Array.from(container.querySelectorAll('button, .btn, .v-btn, label, a, .option'));
        for (let b of btns) {
            try {
                const cls = (b.className || '').toLowerCase();
                if (cls.includes('active') || cls.includes('selected') || cls.includes('checked') || cls.includes('is-checked') || cls.includes('btn--active') || cls.includes('v-btn--is-active')) {
                    const t = clean(b.innerText || b.textContent || '');
                    if (t) return t;
                }
            } catch(e){}
        }

        // 3) procura texto explícito 'sim'/'não' dentro do container (próximo à pergunta)
        try {
            const txt = clean(container.innerText || container.textContent || '');
            const m = txt.match(/\\b(sim|s|não|nao|n)\\b/);
            if (m) return m[1];
        } catch(e){}

        // 4) fallback: tenta labels dentro do container que contenham 'sim' ou 'não'
        try {
            const labs = Array.from(container.querySelectorAll('label'));
            for (let l of labs) {
                const t = clean(l.innerText || '');
                if (t==='sim' || t==='não' || t==='nao') return t;
            }
        } catch(e){}

        return null;
    }

    const resultados = {};
    for (const chave in mapa) {
        try {
            const containers = findContainersContaining(chave);
            let found = false;
            for (const c of containers) {
                const ans = detectAnswerInContainer(c);
                if (ans !== null && ans !== undefined && ans !== '') {
                    // normaliza resposta para Sim/Não em português quando aplicável
                    const a = (ans || '').toString().trim();
                    if (/^s($|im\\b)/i.test(a)) resultados[mapa[chave]] = 'Sim';
                    else if (/^n($|ão\\b|ao\\b)/i.test(a)) resultados[mapa[chave]] = 'Não';
                    else resultados[mapa[chave]] = a;
                    found = true;
                    break;
                }
            }
            // se containers vazios, tenta varredura por proximidade textual simples: procura 'chave' e pega janela de texto
            if (!found) {
                const bodytxt = clean(document.body.innerText || document.body.textContent || '');
                const idx = bodytxt.indexOf(chave);
                if (idx !== -1) {
                    const window = bodytxt.substring(idx, idx + 400);
                    const m = window.match(/\\b(sim|s|não|nao|n)\\b/i);
                    if (m) {
                        const g = m[1].toLowerCase();
                        if (g==='s' || g==='sim') resultados[mapa[chave]] = 'Sim';
                        else resultados[mapa[chave]] = 'Não';
                    }
                }
            }
        } catch(e){}
    }
    return resultados;
    """

    def run_in_current_context():
        try:
            return self.driver.execute_script(script_js, mapa) or {}
        except Exception:
            return {}

    # 1) tenta no contexto atual (document principal)
    dados = run_in_current_context()

    # 2) se vazio, itera iframes tentando em cada um
    if not dados:
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        except:
            iframes = []
        for idx, frame in enumerate(iframes):
            try:
                self.driver.switch_to.frame(frame)
                time.sleep(0.2)
                dados = run_in_current_context()
                # volta para o default antes de decidir
                try:
                    self.driver.switch_to.default_content()
                except:
                    pass
                if dados:
                    break
            except Exception:
                try:
                    self.driver.switch_to.default_content()
                except:
                    pass
                continue

    # Popula saida com normalização e chaves do schema
    try:
        respostas_agregadas = []
        for pergunta_full, resp in (dados or {}).items():
            # normaliza resposta para "Sim"/"Não" quando detectado
            r = (resp or '').strip()
            rl = r.lower()
            if rl in ('s', 'sim'):
                rnorm = 'Sim'
            elif rl in ('n', 'não', 'nao'):
                rnorm = 'Não'
            else:
                # capitaliza primeira letra
                rnorm = r.capitalize() if r else r

            chave_norm = _norm_texto_chave(pergunta_full)
            saida[chave_norm] = rnorm
            respostas_agregadas.append(f"{pergunta_full} -> {rnorm}")

            # tenta preencher a chave do schema correspondente
            for p_schema in COL_SPECS:
                try:
                    if p_schema.aba == 'Questionário' and p_schema.campo == pergunta_full:
                        saida[p_schema.key] = rnorm
                        break
                except:
                    continue

        if respostas_agregadas:
            saida['Respostas do Questionário'] = "\n".join(respostas_agregadas)
    except Exception as e:
        print(f"Erro ao normalizar dados JS: {e}")

    # garante voltar do iframe
    try:
        self.driver.switch_to.default_content()
    except:
        pass

    return saida

# ==============================================================================
# BOT TESTE
# ==============================================================================
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

        print("   -> LENDO QUESTIONÁRIO (MÉTODO MELHORADO)...")
        # substitui a implementação para este teste somente
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        extrator = ExtratorHabibot(self.driver, self.wait, debug=self.debug, debug_visual=self.debug_visual)
        qmap = extrator.extrair_questionario_mapa() or {}
        # atualiza row com respostas (as chaves já vêm normalizadas / com chaves do schema)
        row.update(qmap)
        return row

if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE - LÓGICA MELHORADA")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)