import time
import re
from selenium.webdriver.common.by import By
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 INICIANDO MODO DE TESTE (LÓGICA POSICIONAL)...")

# ==============================================================================
# LÓGICA POSICIONAL (PRIMEIRO É SIM, SEGUNDO É NÃO)
# ==============================================================================
def nova_logica_questionario(self) -> dict:
    """
    Ignora valores e labels.
    - Se o 1º Radio da linha está marcado -> SIM
    - Se o 2º Radio da linha está marcado -> NÃO
    """
    import time
    from selenium.webdriver.common.by import By

    if not self.clicar_aba('Questionário'):
        return {}
    
    time.sleep(3.0) 
    
    saida = {}
    usou_iframe = False
    try:
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            self.driver.switch_to.frame(iframes[0])
            usou_iframe = True
    except: pass

    try:
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

        # SCRIPT JAVASCRIPT: BUSCA POR POSIÇÃO (INDEX)
        script_js = """
        var mapa = arguments[0];
        var resultados = {};
        
        function clean(txt) { return (txt || "").trim().toLowerCase(); }

        // Pega containers de perguntas
        var linhas = document.querySelectorAll("tr, div.row, fieldset, .form-group, div");

        for (var chave in mapa) {
            var perguntaFull = mapa[chave];
            var respostaEncontrada = null;

            for (var i = 0; i < linhas.length; i++) {
                var el = linhas[i];
                if (clean(el.innerText).includes(chave)) {
                    
                    // 1. Procura TODOS os inputs do tipo radio nessa linha
                    var radios = el.querySelectorAll("input[type='radio'], input[type='checkbox']");
                    
                    if (radios.length >= 2) {
                        // LÓGICA DE POSIÇÃO:
                        // O primeiro radio costuma ser SIM (ou opção positiva/true)
                        // O segundo radio costuma ser NÃO (ou opção negativa/false)
                        
                        if (radios[0].checked) {
                            respostaEncontrada = "Sim";
                        } else if (radios[1].checked) {
                            respostaEncontrada = "Não";
                        }
                    } 
                    else if (radios.length === 1) {
                        // Se só tem 1 (tipo checkbox único), se tiver marcado é Sim
                        if (radios[0].checked) respostaEncontrada = "Sim";
                        else respostaEncontrada = "Não";
                    }

                    // 2. Fallback: Botões Visuais (Active Class) por Ordem
                    if (!respostaEncontrada) {
                        var botoes = el.querySelectorAll(".btn, .v-btn, .option");
                        // Filtra botões que parecem opções de sim/não
                        var opcoes = [];
                        for(var k=0; k<botoes.length; k++) {
                            var t = clean(botoes[k].innerText);
                            if(t==='sim' || t==='não' || t==='nao') opcoes.push(botoes[k]);
                        }
                        
                        if (opcoes.length >= 2) {
                            // Verifica qual tem classe 'active'
                            var c0 = (opcoes[0].className || "").toLowerCase();
                            var c1 = (opcoes[1].className || "").toLowerCase();
                            
                            if (c0.includes("active") || c0.includes("checked") || c0.includes("primary")) respostaEncontrada = "Sim";
                            else if (c1.includes("active") || c1.includes("checked") || c1.includes("primary")) respostaEncontrada = "Não";
                        }
                    }

                    if (respostaEncontrada) {
                        resultados[perguntaFull] = respostaEncontrada;
                        break; 
                    }
                }
            }
        }
        return resultados;
        """
        
        dados = self.driver.execute_script(script_js, mapa)
        
        if dados:
            for pergunta, resp in dados.items():
                print(f"✅ Posição detectou: {resp}")
                saida[_norm_texto_chave(pergunta)] = resp
                for p_schema in COL_SPECS:
                    if p_schema.aba == 'Questionário' and pergunta == p_schema.campo:
                        saida[p_schema.key] = resp

    except Exception as e:
        print(f"Erro JS: {e}")

    if usou_iframe:
        try: self.driver.switch_to.default_content()
        except: pass

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

        print("   -> LENDO QUESTIONÁRIO (LÓGICA POSICIONAL 1º=Sim, 2º=Não)...")
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        qmap = extrator.extrair_questionario_mapa() or {}
        row.update(qmap)
        return row

if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE - POSIÇÃO DOS BOTÕES")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)