import time
import os
import re
from selenium.webdriver.common.by import By

# Importa as classes do seu bot original
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 PREPARANDO MODO DE TESTE (DOM INTELIGENTE)...")

# ==============================================================================
# 1. NOVA LÓGICA: RASTREAMENTO VISUAL E DE INPUTS
# ==============================================================================
def nova_logica_questionario(self) -> dict:
    """
    Lógica inteligente: Encontra a pergunta e vasculha a linha inteira
    atrás de 'inputs marcados' ou 'botões acesos'.
    """
    # Imports locais para evitar erros
    import time
    import re
    from selenium.webdriver.common.by import By
    
    # 1. Clica na aba
    if not self.clicar_aba('Questionário'):
        print("⚠️ Não foi possível clicar na aba Questionário.")
        return {}
    
    # 2. Espera carregar
    time.sleep(3.0)
    
    saida = {}
    
    # 3. Iframe (Segurança)
    usou_iframe = False
    try:
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            self.driver.switch_to.frame(iframes[0])
            usou_iframe = True
    except: pass

    try:
        # Mapa de Perguntas (Palavra-chave -> Nome Completo)
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

        # Pega todos os elementos de texto visíveis que podem ser perguntas
        # (Label, Span, TD, P, Div, Strong)
        elementos_texto = self.driver.find_elements(By.CSS_SELECTOR, "label, span, td, p, div, strong, b")

        for chave, pergunta_full in mapa.items():
            resposta_encontrada = ""
            elemento_pergunta = None

            # 1. Encontra onde está a pergunta na tela
            for el in elementos_texto:
                try:
                    if not el.is_displayed(): continue
                    # Verifica se o texto do elemento contém a palavra-chave
                    if chave.lower() in (el.text or "").lower():
                        elemento_pergunta = el
                        break
                except: continue
            
            if elemento_pergunta:
                # 2. Varre os "ancestrais" (Pais) para achar o CONTAINER da pergunta (ex: a linha da tabela)
                container = elemento_pergunta
                for _ in range(5): # Sobe até 5 níveis
                    try:
                        container = container.find_element(By.XPATH, "./..")
                        
                        # --- ESTRATÉGIA A: INPUTS REAIS MARCADOS ---
                        inputs_checked = container.find_elements(By.CSS_SELECTOR, "input:checked")
                        if inputs_checked:
                            val = inputs_checked[0].get_attribute("value")
                            # Se o valor for explícito (Sim/Não)
                            if val and val.lower() in ['sim', 'não', 'nao', 's', 'n']:
                                resposta_encontrada = val
                                break
                            # Se for genérico (true/on/1), tenta achar o label vizinho
                            labels = container.find_elements(By.TAG_NAME, "label")
                            for lbl in labels:
                                # Se o label aponta para esse input ou está perto dele
                                if lbl.get_attribute("for") == inputs_checked[0].get_attribute("id"):
                                    resposta_encontrada = lbl.text
                                    break
                            if resposta_encontrada: break

                        # --- ESTRATÉGIA B: CLASSES VISUAIS (Active, Checked, Selected) ---
                        # Procura elementos com texto "Sim" ou "Não" dentro desse container
                        opcoes = container.find_elements(By.XPATH, ".//*[contains(translate(text(), 'SIMNAO', 'simnao'), 'sim') or contains(translate(text(), 'SIMNAO', 'simnao'), 'não') or contains(translate(text(), 'SIMNAO', 'simnao'), 'nao')]")
                        
                        for opcao in opcoes:
                            texto_opt = (opcao.text or "").strip().lower()
                            if texto_opt not in ['sim', 'não', 'nao']: continue # Ignora frases longas
                            
                            # Verifica se esse elemento (ou o pai dele) tem classe de "ativado"
                            # Classes comuns: active, checked, selected, btn-primary (vs btn-default), v-btn--active
                            classes_el = (opcao.get_attribute("class") or "").lower()
                            classes_pai = (opcao.find_element(By.XPATH, "./..").get_attribute("class") or "").lower()
                            
                            combo_classes = classes_el + " " + classes_pai
                            
                            # Palavras magicas que indicam seleção
                            if "active" in combo_classes or "checked" in combo_classes or "selected" in combo_classes or "btn-primary" in combo_classes or "btn-success" in combo_classes:
                                resposta_encontrada = texto_opt.capitalize().replace('nao', 'Não')
                                break
                        
                        if resposta_encontrada: break
                        
                    except: pass
            
            if resposta_encontrada:
                # Normaliza
                if resposta_encontrada.lower() in ['s', 'sim', '1', 'true']: resposta_encontrada = 'Sim'
                elif resposta_encontrada.lower() in ['n', 'não', 'nao', '0', 'false']: resposta_encontrada = 'Não'
                
                print(f"✅ {chave[:15]}... -> {resposta_encontrada}")
                
                # Salva
                saida[_norm_texto_chave(pergunta_full)] = resposta_encontrada
                for p_schema in COL_SPECS:
                    if p_schema.aba == 'Questionário' and pergunta_full == p_schema.campo:
                        saida[p_schema.key] = resposta_encontrada
                        break
            else:
                print(f"⚠️ Não achei resposta marcada para: {chave}")

    except Exception as e:
        print(f"Erro na varredura inteligente: {e}")

    # Sai do iframe
    if usou_iframe:
        try: self.driver.switch_to.default_content()
        except: pass

    return saida

# ==============================================================================
# 2. BOT TURBO (IGUAL AO ANTERIOR, SÓ MUDA A LÓGICA DO QUESTIONÁRIO)
# ==============================================================================
class BotTeste(HabibotBot):
    def extrair_um(self) -> dict:
        extrator = ExtratorHabibot(
            self.driver, 
            self.wait, 
            debug=self.debug, 
            debug_visual=self.debug_visual,
            ordem_logger=getattr(self, 'ordem_logger', None),
            acoes_logger=getattr(self, 'acoes_logger', None)
        )
        row = {}

        # 1. Identificação Rápida
        print("   -> Lendo Identificação...")
        if extrator.clicar_aba('Titular'):
            try:
                time.sleep(1)
                row[_col_key('Titular', 'Dados Gerais', 'Nome')] = extrator._extrair_por_label_simples('Nome')
                row[_col_key('Titular', 'Dados Gerais', 'CPF/CNPJ')] = extrator._extrair_por_label_simples('CPF/CNPJ')
            except: pass

        # 2. Pula o resto...

        # 3. Questionário com a NOVA lógica (Injetada)
        print("   -> Lendo Questionário (Smart DOM)...")
        # Injetamos a função diretamente aqui
        ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
        qmap = extrator.extrair_questionario_mapa() or {}
        row.update(qmap)

        return row

# ==============================================================================
# 3. EXECUÇÃO
# ==============================================================================
if __name__ == "__main__":
    print("\n🚀 INICIANDO TESTE (MODO DOM INTELIGENTE)")
    bot = BotTeste()
    bot.executar(debug=True, debug_visual=True)