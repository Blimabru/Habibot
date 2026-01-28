import time
import os
import re
from selenium.webdriver.common.by import By

# Importa as classes e configurações do seu bot original
from habibot import HabibotBot, ExtratorHabibot, COL_SPECS, _col_key, _norm_texto_chave

print("🔧 PREPARANDO MODO DE TESTE (FOCADO NO QUESTIONÁRIO)...")

# ==============================================================================
# 1. NOVA LÓGICA DO QUESTIONÁRIO (TEXTO BRUTO)
# ==============================================================================
def nova_logica_questionario(self) -> dict:
    """
    Substitui a função original.
    Lê o TEXTO visível na tela em vez de procurar botões escondidos.
    """
    
    # 1. Clica na aba
    if not self.clicar_aba('Questionário'):
        print("⚠️ Não foi possível clicar na aba Questionário.")
        return {}
    
    # 2. Espera 3 segundos (Fundamental para o texto carregar)
    time.sleep(3.0)
    
    saida = {}
    
    # 3. Segurança: verifica se o conteúdo está dentro de um Iframe
    usou_iframe = False
    try:
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            self.driver.switch_to.frame(iframes[0])
            usou_iframe = True
    except: 
        pass

    # 4. Lê o texto da página e procura as respostas
    try:
        # Pega TODO o texto visível da página
        corpo = self.driver.find_element(By.TAG_NAME, "body").text
        # Limpa espaços extras e deixa tudo minúsculo
        texto_limpo = " ".join(corpo.split()).lower()
        
        # Mapeia as perguntas
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

        for chave, pergunta_full in mapa.items():
            if chave in texto_limpo:
                # REGEX: Procura a pergunta + até 350 caracteres de lixo + a palavra Sim ou Não
                padrao = re.escape(chave) + r".{0,350}?\b(sim|não|nao)\b"
                match = re.search(padrao, texto_limpo)
                
                if match:
                    resp = match.group(1).capitalize()
                    if resp == "Nao": resp = "Não"
                    
                    # Salva nas chaves corretas para o Excel
                    
                    # 1. Chave normalizada (para uso interno)
                    saida[_norm_texto_chave(pergunta_full)] = resp
                    
                    # 2. Chave exata do Schema (Fundamental para gravar no Excel)
                    for p_schema in COL_SPECS:
                        if p_schema.aba == 'Questionário' and pergunta_full == p_schema.campo:
                            saida[p_schema.key] = resp
                            break
    except Exception as e:
        print(f"Erro na leitura do texto: {e}")

    # Sai do iframe se tiver entrado
    if usou_iframe:
        try: self.driver.switch_to.default_content()
        except: pass

    return saida

# ==============================================================================
# 2. FUNÇÃO RESUMIDA (Para ser rápido e testar só o Questionário)
# ==============================================================================
def extrair_apenas_questionario(self) -> dict:
    # Cria uma instância do extrator original passando os parâmetros corretos
    extrator = ExtratorHabibot(
        self.driver, 
        self.wait, 
        debug=self.debug, 
        debug_visual=self.debug_visual,
        ordem_logger=getattr(self, 'ordem_logger', None),
        acoes_logger=getattr(self, 'acoes_logger', None)
    )
    row = {}

    # 1. Pega apenas Nome/CPF para identificar o candidato
    print("   -> Lendo Nome/CPF...")
    titular = extrator.extrair_titular_mapa() or {}
    row.update(titular)

    # 2. PULA Renda, Endereço, etc. para ganhar tempo e focar no teste
    
    # 3. Pega o Questionário com a NOVA lógica
    print("   -> Lendo Questionário (Texto)...")
    qmap = extrator.extrair_questionario_mapa() or {}
    row.update(qmap)

    return row

# ==============================================================================
# 3. SUBSTITUI AS FUNÇÕES NO ROBÔ ORIGINAL (VACINA)
# ==============================================================================
# Aqui trocamos as peças do bot original pelas nossas novas versões
ExtratorHabibot.extrair_questionario_mapa = nova_logica_questionario
HabibotBot.extrair_um = extrair_apenas_questionario

# ==============================================================================
# 4. EXECUÇÃO
# ==============================================================================
if __name__ == "__main__":
    print("\n🚀 INICIANDO HABIBOT (MODO DE TESTE DE QUESTIONÁRIO)")
    
    bot = HabibotBot()
    
    # Executa o bot. Ele vai pedir LOGIN e SENHA no terminal.
    # Quando perguntar o modo, escolha "2" e cole o CPF do candidato.
    bot.executar(debug=True, debug_visual=True)