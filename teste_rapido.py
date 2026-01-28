import time
import re
import os
from selenium.webdriver.common.by import By
# Importa o bot original para usar as funções de login e credenciais dele
from habibot import HabibotBot, _norm_texto_chave 

# ==============================================================================
# 👇 AQUI FICA A LÓGICA NOVA QUE ESTAMOS TESTANDO (BOMBA ATÔMICA V3 - TEXTO)
# ==============================================================================
def minha_nova_logica_questionario(driver):
    # 1. Clica na aba (se precisar)
    try:
        abas = driver.find_elements(By.XPATH, "//a[contains(text(), 'Questionário')] | //button[contains(text(), 'Questionário')]")
        for aba in abas:
            if aba.is_displayed():
                aba.click()
                break
    except: pass
    
    # Espera o texto aparecer
    time.sleep(3) 

    print("🔎 Lendo a tela (Modo Texto)...")
    saida = {}

    try:
        # 1. Tenta focar no Iframe se existir (segurança extra)
        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if iframes:
                driver.switch_to.frame(iframes[0])
        except: pass

        # 2. Pega O TEXTO PURO da página (Ignora HTML complexo)
        corpo = driver.find_element(By.TAG_NAME, "body").text
        # Limpa espaços extras e deixa tudo minúsculo para facilitar a busca
        texto_limpo = " ".join(corpo.split()).lower()
        
        # Debug: Mostra um pedaço para você ver se o texto veio
        print(f"\n--- TEXTO CAPTURADO (AMOSTRA) ---\n{texto_limpo[:150]}...\n---------------------------------\n")

        # Mapa das perguntas (Palavra-chave -> Nome Curto para Teste)
        mapa = {
            'responsável pela unidade': 'A mulher é a responsável?',
            'pessoa negra': 'Pessoa negra?',
            'deficiência': 'Pessoa com deficiência?',
            'idoso': 'Idoso?',
            'criança': 'Criança/Adolescente?',
            'câncer': 'Câncer/Doença Rara?',
            'violência': 'Violência Doméstica?',
            'indígenas': 'Indígenas/Quilombolas?',
            'área de risco': 'Área de Risco?',
            'distratado': 'Contrato Distratado?',
            'socioassistenciais': 'Redes Socioassistenciais?'
        }

        for chave, nome_curto in mapa.items():
            if chave in texto_limpo:
                # REGEX MÁGICO:
                # Procura a frase da pergunta + qualquer coisa no meio + a palavra "Sim" ou "Não"
                # O .{0,300}? significa "aceite até 300 caracteres de lixo entre a pergunta e a resposta"
                padrao = re.escape(chave) + r".{0,300}?\b(sim|não|nao)\b"
                match = re.search(padrao, texto_limpo)
                
                if match:
                    resp = match.group(1).capitalize() # Pega o que achou (Sim ou Não)
                    if resp == "Nao": resp = "Não"
                    print(f"✅ {nome_curto} -> {resp}")
                    saida[nome_curto] = resp
                else:
                    print(f"❌ {nome_curto}: Achou pergunta, mas não achou Sim/Não perto.")
            else:
                print(f"⚠️ {nome_curto}: Não achou o texto da pergunta na tela.")

        # Volta do iframe se entrou
        try: driver.switch_to.default_content()
        except: pass

    except Exception as e:
        print(f"Erro no teste: {e}")

    return saida
# ==============================================================================


# --- SISTEMA DE TESTE (NÃO PRECISA MEXER) ---
if __name__ == "__main__":
    print("\n🛠️  MODO DE TESTE RÁPIDO DO HABIBOT")
    print("---------------------------------------")
    
    bot = HabibotBot()
    
    # 1. Solicita credenciais igual ao bot oficial (seguro)
    print("Por favor, faça login para iniciar os testes.")
    usuario, senha = bot.solicitar_credenciais()
    
    # 2. Inicia o navegador
    bot.iniciar(headless=False)
    
    try:
        # 3. Faz o Login
        if not bot.base_url:
            bot.base_url = "https://app.habisoft.com.br" # URL Padrão
            
        print("🔑 Logando no sistema...")
        bot.login(usuario, senha)
        
        print("\n✅ Login realizado!")
        print("\n👇 AGORA É COM VOCÊ:")
        print("1. No navegador aberto, VÁ MANUALMENTE até a página do candidato que você quer testar.")
        print("2. Quando estiver vendo o questionário na tela, VOLTE AQUI.")
        
        while True:
            input("\n🟢 Pressione ENTER aqui para rodar a extração do questionário...")
            print("\n" + "="*50)
            
            # Roda a função de teste
            resultado = minha_nova_logica_questionario(bot.driver)
            
            print("-" * 50)
            print("RESULTADO FINAL:", resultado)
            print("=" * 50)
            print("\n🔁 Pode navegar para outro candidato e apertar ENTER novamente.")
            
    except KeyboardInterrupt:
        print("\nEncerrando testes.")
    finally:
        bot.encerrar()