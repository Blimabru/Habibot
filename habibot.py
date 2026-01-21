"""
Bot de Automação Habibot
------------------------
Este bot realiza a extração automatizada de TODOS os candidatos cadastrados no Sistema de Habitação, salvando os dados em um relatório Excel.

Fluxo principal:
1. Solicita a URL do Sistema de Habitação
2. Solicita login e senha
3. Faz login e navega até a lista de candidatos
4. Extrai todos os dados de todos os candidatos, página por página
5. Salva os dados em um arquivo Excel

Saída:
- Pasta: "Extração de Dados" (no diretório onde o bot for executado)
- Nome: "Extração de Dados <Data> - <Hora>.xlsx"
"""


import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from getpass import getpass
from pathlib import Path
import textwrap
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException, SessionNotCreatedException
from openpyxl import Workbook, load_workbook


__app_name__ = "Bot de Automação Habibot"
__author__ = "Bruno Lima"
__language__ = "pt-BR"
__version__ = "1.0.0"
__release_date__ = "2026-01-18"  # AAAA-MM-DD
__description__ = (
    "Este bot de automação foi criado para suprir a necessidade de um relatório completo no Sistema de Habitação, que disponibilize todos os dados de todos os candidatos cadastrados."
)


#######################################################################
# Funções utilitárias para ambiente, terminal e manipulação de arquivos #
#######################################################################
def _diretorio_base_execucao() -> Path:
    """Retorna o diretório base do script/EXE.

    - Em modo EXE (PyInstaller), usa a pasta do executável.
    - Em modo .py, usa a pasta do arquivo.
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _encontrar_chromedriver_local() -> str | None:
    """Procura um chromedriver.exe local/embarcado para execução offline.

    Ordem:
    1) Variável de ambiente Habibot_CHROMEDRIVER
    2) Ao lado do EXE/script
    3) Pasta drivers/ ao lado do EXE/script
    4) Pasta drivers/ dentro do bundle do PyInstaller (_MEIPASS)
    """
    env = os.getenv('Habibot_CHROMEDRIVER')
    if env:
        try:
            p = Path(env).expanduser().resolve()
            if p.exists():
                return str(p)
        except:
            pass

    base = _diretorio_base_execucao()
    candidatos = [
        base / 'chromedriver.exe',
        base / 'drivers' / 'chromedriver.exe',
        base / 'assets' / 'drivers' / 'chromedriver.exe',
    ]

    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        try:
            candidatos.append(Path(meipass) / 'drivers' / 'chromedriver.exe')
            candidatos.append(Path(meipass) / 'assets' / 'drivers' / 'chromedriver.exe')
        except:
            pass

    for p in candidatos:
        try:
            if p.exists():
                return str(p)
        except:
            continue

    return None


def _terminal_width(padrao: int = 100) -> int:
    try:
        # columns, lines
        return max(60, shutil.get_terminal_size((padrao, 20)).columns)
    except:
        return max(60, padrao)


SEPARADOR_TELA = "─"


def _linha(char: str = SEPARADOR_TELA, largura: int | None = None) -> str:
    w = largura or _terminal_width()
    return char * w


def _print_bloco(titulo: str, linhas: list[str], largura: int | None = None):
    w = largura or _terminal_width()
    print(_linha(SEPARADOR_TELA, w))
    print(titulo.center(w))
    print(_linha(SEPARADOR_TELA, w))
    for l in linhas:
        if not l:
            print("")
            continue
        for parte in textwrap.wrap(l, width=w, break_long_words=False, break_on_hyphens=False):
            print(parte)


def mostrar_boas_vindas():
    w = _terminal_width()

    titulo = "🤖 🏠  Bot de Automação Habibot"
    linhas = [
        f"Criado por: {__author__}",
        f"Versão: {__version__}",
        "Atualizado em: 18/01/2026",
        "",
        "🤖 Este bot foi criado para suprir a necessidade de um relatório completo do sistema de Habitação, que disponibilizasse todos os dados de todos os candidatos cadastrados.",
        "",
        "⚠️  Atenção!",
        "Como o bot acessa cada candidato individualmente e salva os dados em uma planilha do Excel, o processo pode gerar muitas requisições em pouco tempo. Em alguns cenários, o servidor da Habibot pode interpretar isso como atividade suspeita e bloquear temporariamente o acesso.",
        "",
        "📌 Acompanhar o progresso",
        "Enquanto a extração estiver em andamento, você pode acompanhar o progresso abrindo o arquivo de visualização dentro da pasta \"Extração de Dados\" (arquivo com prefixo \"VISUALIZAÇÃO\"). Para atualizar, feche e abra novamente.",
        "",
    ]

    print("")
    print(_linha(SEPARADOR_TELA, w))
    print("")
    print(titulo.center(w))
    print(_linha(SEPARADOR_TELA, w))
    for l in linhas:
        if not l:
            print("")
            continue
        for parte in textwrap.wrap(l, width=w, break_long_words=False, break_on_hyphens=False):
            print(parte)
    print(_linha(SEPARADOR_TELA, w))
    print("")

COLUNAS_EXCEL = [
    'Nome',
    'CPF',
    'Telefone',
    'Email',
    'RG',
    'Data de nascimento',
    'Raça',
    'Gênero',
    'Aba Endereço',
    'Estado civil',
    'Cadastro preferencial',
    'Deficiência',
    'Doença',
    'Observação',
    'Diagnóstico social',
    'Tipo de beneficiário programa social',
    'Situação de emprego',
    'Informações do cadunico',
    'Aba Renda',
    'Aba Questionário',
    'Loteamento',
    'Grupo familiar',
    'Data de extração',
]



#########################################################
# Classe para manipulação do Excel em tempo real         #
#########################################################
class ExcelTempoReal:
    """Grava o Excel final e mantém um Excel de visualização (cópia) atualizado.

    O arquivo de visualização é apagado ao final (ou quando o usuário interromper).
    """

    def __init__(self, caminho_final: str, caminho_VISUALIZAÇÃO: str):
        """
        Inicializa o gerenciador de Excel, criando o arquivo final e a visualização.
        caminho_final: caminho do arquivo Excel principal (final)
        caminho_VISUALIZAÇÃO: caminho do arquivo de visualização (cópia temporária)
        """
        self.caminho_final = caminho_final
        self.caminho_VISUALIZAÇÃO_base = caminho_VISUALIZAÇÃO
        self.caminhos_VISUALIZAÇÃO_criados = []
        self._inicializar_arquivo(self.caminho_final)
        self._atualizar_VISUALIZAÇÃO()

    def _inicializar_arquivo(self, caminho_arquivo: str):
        """
        Cria o arquivo Excel se não existir e garante a estrutura da planilha.
        Se o arquivo já existe, garante que a aba e o cabeçalho estejam presentes.
        """
        pasta = os.path.dirname(caminho_arquivo)
        if pasta and not os.path.exists(pasta):
            os.makedirs(pasta)
        if not os.path.exists(caminho_arquivo):
            wb = Workbook()
            ws = wb.active
            ws.title = 'Dados'
            ws.append(COLUNAS_EXCEL)
            wb.save(caminho_arquivo)
            return
        wb = load_workbook(caminho_arquivo)
        ws = wb['Dados'] if 'Dados' in wb.sheetnames else wb.active
        if ws.max_row < 1:
            ws.append(COLUNAS_EXCEL)
        wb.save(caminho_arquivo)

    @staticmethod
    def _normalizar(valor, coluna):
        """
        Normaliza valores para escrita no Excel, preenchendo campos vazios.
        Para a coluna 'Data de extração', retorna string vazia se None.
        Para outros campos, retorna 'Não Informado' se vazio ou None.
        """
        if coluna == 'Data de extração':
            if valor is None:
                return ''
            return str(valor).strip()
        if valor is None:
            return 'Não Informado'
        if isinstance(valor, str):
            return valor.strip() if valor.strip() else 'Não Informado'
        texto = str(valor).strip()
        return texto if texto else 'Não Informado'

    def adicionar_linha(self, registro: dict):
        """
        Adiciona uma linha de dados ao Excel final e atualiza a visualização.
        registro: dicionário com os dados do candidato.
        Retorna True se sucesso, False se arquivo estiver bloqueado.
        """
        linha = [self._normalizar(registro.get(col), col) for col in COLUNAS_EXCEL]
        try:
            wb = load_workbook(self.caminho_final)
            ws = wb['Dados'] if 'Dados' in wb.sheetnames else wb.active
            ws.append(linha)
            wb.save(self.caminho_final)
            self._atualizar_VISUALIZAÇÃO()
            return True
        except PermissionError:
            print("❌ O arquivo FINAL está aberto no Excel e bloqueou a escrita.")
            print(f"   Feche o arquivo e tente novamente: {self.caminho_final}")
            return False

    def _atualizar_VISUALIZAÇÃO(self):
        """
        Copia o arquivo final para um arquivo de visualização (pode ser aberto durante a execução).
        Se o arquivo de visualização estiver bloqueado, cria cópias incrementais.
        """
        destino = self.caminho_VISUALIZAÇÃO_base
        pasta = os.path.dirname(destino)
        if pasta and not os.path.exists(pasta):
            os.makedirs(pasta)
        try:
            shutil.copyfile(self.caminho_final, destino)
            if destino not in self.caminhos_VISUALIZAÇÃO_criados:
                self.caminhos_VISUALIZAÇÃO_criados.append(destino)
            return destino
        except PermissionError:
            # Se o arquivo estiver aberto, cria uma cópia incremental
            base, ext = os.path.splitext(self.caminho_VISUALIZAÇÃO_base)
            for i in range(1, 1000):
                alt = f"{base}_{i:03d}{ext}"
                try:
                    shutil.copyfile(self.caminho_final, alt)
                    self.caminhos_VISUALIZAÇÃO_criados.append(alt)
                    print(f"⚠️  Visualização aberta/travada. Nova visualização: {alt}")
                    return alt
                except PermissionError:
                    continue
            print("⚠️  Não foi possível atualizar a visualização (arquivos travados).")
            return None

    def limpar_visualizacoes(self):
        """
        Remove todos os arquivos de visualização criados nesta execução.
        """
        for caminho in list(dict.fromkeys(self.caminhos_VISUALIZAÇÃO_criados)):
            try:
                if os.path.exists(caminho):
                    os.remove(caminho)
            except:
                pass



#########################################################
# Classe para extração dos dados de cada candidato       #
#########################################################
class ExtratorHabibot:
    def __init__(self, driver, wait):
        self.driver = driver
        self.wait = wait

    def aguardar(self, segundos):
        self.driver.implicitly_wait(segundos)

    def clicar_aba(self, nome_aba: str) -> bool:
        try:
            xpath = (
                f"//a[contains(text(), '{nome_aba}')] | "
                f"//button[contains(text(), '{nome_aba}')] | "
                f"//*[@role='tab' and contains(text(), '{nome_aba}')]"
            )
            abas = self.driver.find_elements(By.XPATH, xpath)
            for aba in abas:
                if aba.is_displayed():
                    self.driver.execute_script("arguments[0].click();", aba)
                    self.aguardar(1)
                    return True
            return False
        except:
            return False

    def extrair_por_label(self, texto_label: str) -> str:
        try:
            texto_busca = texto_label.lower().strip()
            xpath = (
                "//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
                f"'{texto_busca}')]"
            )
            labels = self.driver.find_elements(By.XPATH, xpath)
            for label in labels:
                try:
                    parent = label.find_element(By.XPATH, "..")
                    campos = parent.find_elements(By.CSS_SELECTOR, "input:not([type='hidden']), select, textarea")
                    for campo in campos:
                        if not campo.is_displayed():
                            continue

                        if campo.tag_name == 'select':
                            try:
                                opcao = campo.find_element(By.CSS_SELECTOR, "option:checked, option[selected]")
                                texto = (opcao.text or "").strip()
                                if texto and not texto.lower().startswith('selecione') and texto.lower() not in ['nenhum', 'nenhuma']:
                                    return texto
                            except:
                                pass
                            continue

                        valor = campo.get_attribute('value')
                        if valor and valor.strip():
                            return valor.strip()

                        if campo.tag_name == 'textarea':
                            texto = (campo.text or "").strip()
                            if texto:
                                return texto
                except:
                    continue
        except:
            pass
        return ""

    def extrair_dados_titular(self) -> dict:
        if not self.clicar_aba('Titular'):
            return {}

        return {
            'Loteamento': self.extrair_por_label('Loteamento'),
            'Nome': self.extrair_por_label('Nome'),
            'CPF': self.extrair_por_label('CPF/CNPJ'),
            'RG': self.extrair_por_label('Nº Documento'),
            'Data de nascimento': self.extrair_por_label('Data de Nascimento'),
            'Raça': self.extrair_por_label('Raça'),
            'Gênero': self.extrair_por_label('Gênero'),
            'Estado civil': self.extrair_por_label('Estado Civil'),
            'Cadastro preferencial': self.extrair_por_label('Cadastro Preferencial'),
            'Deficiência': self.extrair_por_label('Deficiência'),
            'Doença': self.extrair_por_label('Doença'),
            'Situação de emprego': self.extrair_por_label('Situação de Emprego'),
            'Tipo de beneficiário programa social': self.extrair_por_label('Tipo de Beneficiário Programa Social'),
            'Informações do cadunico': self.extrair_por_label('A família esta inscrita no CAD Único'),
            'Email': self.extrair_por_label('Email'),
            'Telefone': self.extrair_por_label('Telefone Celular'),
            'Observação': self.extrair_por_label('Observações'),
            'Diagnóstico social': self.extrair_por_label('Diagnóstico Social'),
        }

    def extrair_endereco(self) -> dict:
        if not self.clicar_aba('Endereço'):
            return {}
        return {
            'CEP': self.extrair_por_label('CEP'),
            'UF': self.extrair_por_label('UF'),
            'Cidade': self.extrair_por_label('Cidade'),
            'Bairro': self.extrair_por_label('Bairro'),
            'Endereço': self.extrair_por_label('Endereço'),
            'Número': self.extrair_por_label('Número'),
            'Complemento': self.extrair_por_label('Complemento'),
            'Lote': self.extrair_por_label('Lote'),
            'Quadra': self.extrair_por_label('Quadra'),
        }

    def extrair_renda(self) -> dict:
        if not self.clicar_aba('Renda'):
            return {"rendas": [], "renda_total": "", "renda_per_capita": ""}

        rendas = []
        try:
            linhas = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            for linha in linhas:
                cols = [c.text.strip() for c in linha.find_elements(By.CSS_SELECTOR, "td")]
                if not cols:
                    continue
                nome = cols[0] if len(cols) > 0 else ""
                tipo = cols[1] if len(cols) > 1 else ""
                valor = cols[2] if len(cols) > 2 else ""
                if nome or tipo or valor:
                    rendas.append({"nome": nome, "tipo": tipo, "valor": valor})
        except:
            pass

        renda_total = self.extrair_por_label('Renda Familiar')
        renda_pc = self.extrair_por_label('Renda Per Capita')

        return {
            "rendas": rendas,
            "renda_total": f"Renda total: {renda_total}" if renda_total else "",
            "renda_per_capita": f"Renda por integrante: {renda_pc}" if renda_pc else "",
        }

    def extrair_questionario(self) -> dict:
        if not self.clicar_aba('Questionário'):
            return {}
        respostas = {}
        try:
            marcados = self.driver.find_elements(By.CSS_SELECTOR, "input[type='radio']:checked")
            for idx, el in enumerate(marcados, start=1):
                val = (el.get_attribute('value') or '').strip()
                if val:
                    respostas[f"pergunta_{idx}"] = val
        except:
            pass
        return respostas

    def extrair_segundo_titular(self) -> list:
        if not self.clicar_aba('Segundo Titular'):
            return []
        # Sem estrutura padronizada garantida: retorna lista vazia se não houver tabela
        lista = []
        try:
            linhas = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            for linha in linhas:
                cols = [c.text.strip() for c in linha.find_elements(By.CSS_SELECTOR, "td")]
                if not cols:
                    continue
                nome = cols[0] if len(cols) > 0 else ""
                cpf = cols[1] if len(cols) > 1 else ""
                if nome or cpf:
                    lista.append({"nome": nome, "cpf": cpf})
        except:
            pass
        return lista

    def extrair_composicao_familiar(self) -> list:
        if not self.clicar_aba('Composição Familiar'):
            return []
        membros = []
        try:
            linhas = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            for linha in linhas:
                cols = [c.text.strip() for c in linha.find_elements(By.CSS_SELECTOR, "td")]
                if not cols:
                    continue
                nome = cols[0] if len(cols) > 0 else ""
                cpf = cols[1] if len(cols) > 1 else ""
                tipo = cols[2] if len(cols) > 2 else ""
                if nome or cpf or tipo:
                    membros.append({"nome": nome, "cpf": cpf, "tipo": tipo})
        except:
            pass
        return membros



#########################################################
# Classe principal do bot: fluxo de execução, login etc. #
#########################################################
class HabibotBot:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.excel = None
        self.caminho_excel_final = None
        self.caminho_excel_VISUALIZAÇÃO = None
        self.base_url = None

    @staticmethod
    def solicitar_modo_execucao() -> bool:
        print("\n🖥️  Modo de execução:\n")
        print("1) Visualizar bot trabalhando (abrir o navegador)")
        print("2) Segundo plano (sem janela do navegador)\n")
        escolha = input("Selecione (1/2) [1]: ").strip() or "1"

        if escolha == "2":
            print("\n🕶️  Modo selecionado: segundo plano (headless)\n\n")
            return True

        print("\n🖥️  Modo selecionado: visualização no navegador\n\n")
        return False

    def iniciar(self, headless: bool = False):
        chrome_options = Options()

        if headless:
            chrome_options.add_argument('--headless=new')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--disable-software-rasterizer')
        else:
            chrome_options.add_argument('--start-maximized')

        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        # Silencia logs internos do Chrome/Chromium (não afeta a automação)
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

        chromedriver_local = _encontrar_chromedriver_local()
        if chromedriver_local:
            print(f"🧩 Usando ChromeDriver local: {chromedriver_local}")
            try:
                service = Service(chromedriver_local, log_output=subprocess.DEVNULL)
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception as e:
                print("⚠️  Falha ao iniciar com ChromeDriver local. Tentando alternativa via download...")
                print(f"   Detalhe: {e}")
                service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
        else:
            service = Service(ChromeDriverManager().install(), log_output=subprocess.DEVNULL)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 30)

    def encerrar(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass

    def solicitar_base_url(self):
        print("\n🌐 Insira o link do sistema de Habitação\n")
        print("Exemplo: app.habibot.com.br\n")
        while True:
            url = input("URL do sistema: ").strip()
            if not url:
                print("Por favor, insira uma URL válida.\n")
                continue
            if not url.startswith("http"):
                url = "https://" + url
            url = url.rstrip("/")
            self.base_url = url
            break
        print(f"\nURL definida: {self.base_url}\n")

    def solicitar_credenciais(self):
        print("🔐 Credenciais do Habibot")
        print("")
        usuario = input("Usuário (e-mail): ").strip()
        senha = getpass("Senha: ")
        print("")
        if not usuario or not senha:
            raise ValueError('Usuário e senha são obrigatórios')
        return usuario, senha

    def login(self, usuario: str, senha: str):
        login_url = f"{self.base_url}/login"
        self.driver.get(login_url)
        campo_usuario = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        campo_senha = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        campo_senha.clear()
        campo_senha.send_keys(senha)

        botao = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit'], button.btn-primary")
        botao.click()
        self.wait.until(EC.url_changes(login_url))

    def navegar_candidatos(self):
        candidatos_url = f"{self.base_url}/candidatos"
        self.driver.get(candidatos_url)
        self.wait.until(lambda d: d.execute_script("return document.readyState === 'complete'"))
        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except:
            pass
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr, table tbody tr, .table tbody tr")))

    def obter_linhas(self):
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")))
        except:
            pass
        linhas = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        return [l for l in linhas if l.is_displayed()]

    def _pagina_ativa(self) -> str:
        try:
            el = self.driver.find_element(By.XPATH, "//ul[contains(@class,'pagination')]//li[contains(@class,'active')]//a | //ul[contains(@class,'pagination')]//li[contains(@class,'active')]//button")
            return (el.text or '').strip()
        except:
            return ''

    def proxima_pagina(self) -> bool:
        antes = self._pagina_ativa()

        candidatos_xpath = [
            "//ul[contains(@class,'pagination')]//li[contains(@class,'active')]/following-sibling::li[not(contains(@class,'disabled'))][1]//a",
            "//ul[contains(@class,'pagination')]//li[contains(@class,'active')]/following-sibling::li[not(contains(@class,'disabled'))][1]//button",
            "//button[contains(translate(@aria-label,'PRÓXIMA','próxima'),'próxima') or contains(translate(@aria-label,'NEXT','next'),'next')]",
            "//a[contains(translate(.,'PRÓXIMA','próxima'),'próxima') or contains(translate(.,'NEXT','next'),'next') or contains(.,'»') or contains(.,'›')]",
        ]

        botao = None
        for xp in candidatos_xpath:
            try:
                el = self.driver.find_element(By.XPATH, xp)
                if el.is_displayed() and el.is_enabled():
                    botao = el
                    break
            except:
                continue

        if not botao:
            return False

        assinatura_antes = ''
        try:
            assinatura_antes = (self.driver.find_element(By.CSS_SELECTOR, "tbody tr").text or '').strip()
        except:
            pass

        try:
            self.driver.execute_script("arguments[0].click();", botao)
        except:
            botao.click()

        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except:
            pass

        for _ in range(40):
            depois = self._pagina_ativa()
            if antes and depois and depois != antes:
                return True
            try:
                assinatura_depois = (self.driver.find_element(By.CSS_SELECTOR, "tbody tr").text or '').strip()
                if assinatura_antes and assinatura_depois and assinatura_depois != assinatura_antes:
                    return True
            except:
                pass
            time.sleep(0.5)

        return True

    def abrir_candidato(self, linha):
        btn = linha.find_element(By.CSS_SELECTOR, "button.btn-soft-info")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        time.sleep(0.3)
        try:
            btn.click()
        except:
            self.driver.execute_script("arguments[0].click();", btn)

        self.wait.until(EC.url_contains('/dadoscandidato'))
        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-overlay.is-active")))
        except:
            pass
        time.sleep(0.5)

    def voltar_lista(self):
        try:
            self.driver.back()
            self.wait.until(lambda d: '/candidatos' in d.current_url)
        except:
            self.navegar_candidatos()

        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except:
            pass

    def extrair_um(self) -> dict:
        extrator = ExtratorHabibot(self.driver, self.wait)

        registro = {}
        registro.update(extrator.extrair_dados_titular())

        end = extrator.extrair_endereco()
        partes_end = []
        for k in ["CEP", "UF", "Cidade", "Bairro", "Endereço", "Número", "Complemento", "Lote", "Quadra"]:
            v = (end.get(k) or '').strip()
            if v:
                partes_end.append(f"{k}: {v}")
        registro['Aba Endereço'] = ' | '.join(partes_end)

        renda = extrator.extrair_renda()
        partes_r = []
        if renda.get('renda_total'):
            partes_r.append(renda['renda_total'])
        if renda.get('renda_per_capita'):
            partes_r.append(renda['renda_per_capita'])
        for r in renda.get('rendas', []):
            nm = (r.get('nome') or '').strip()
            tp = (r.get('tipo') or '').strip()
            vl = (r.get('valor') or '').strip()
            if nm or tp or vl:
                partes_r.append(f"{nm} - {tp}: {vl}".strip())
        registro['Aba Renda'] = ' | '.join([p for p in partes_r if p])

        q = extrator.extrair_questionario()
        registro['Aba Questionário'] = ' | '.join([f"{k}: {v}" for k, v in q.items() if str(v).strip()])

        cot = extrator.extrair_segundo_titular()
        fam = extrator.extrair_composicao_familiar()
        partes_g = []
        if cot:
            partes_g.append("Cotitular: " + "; ".join([f"{(c.get('nome') or '').strip()} ({(c.get('cpf') or '').strip()})".strip() for c in cot if (c.get('nome') or '').strip() or (c.get('cpf') or '').strip()]))
        if fam:
            mems = []
            for m in fam:
                nome = (m.get('nome') or '').strip()
                cpf = (m.get('cpf') or '').strip()
                tipo = (m.get('tipo') or '').strip()
                base = f"{nome} ({cpf})".strip()
                if tipo:
                    base = f"{base} - {tipo}" if base else tipo
                if base:
                    mems.append(base)
            if mems:
                partes_g.append("Composição Familiar: " + "; ".join(mems))
        registro['Grupo familiar'] = ' | '.join([p for p in partes_g if p])

        registro['Data de extração'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        # Garante todas as colunas presentes
        for col in COLUNAS_EXCEL:
            registro.setdefault(col, '')

        return registro

    def executar(self):
        mostrar_boas_vindas()
        headless = self.solicitar_modo_execucao()
        self.solicitar_base_url()
        usuario, senha = self.solicitar_credenciais()

        self.iniciar(headless=headless)
        try:
            print("\n🔐 Fazendo login...")
            self.login(usuario, senha)
            print("✅ Login OK")

            print("\n📋 Navegando para candidatos...")
            self.navegar_candidatos()
            print("✅ Página de candidatos carregada")

            # Pasta de saída: "Extração de Dados" dentro da pasta onde o bot foi executado
            pasta_saida = os.path.join(os.getcwd(), 'Extração de Dados')
            os.makedirs(pasta_saida, exist_ok=True)

            agora = datetime.now()
            data_ext = agora.strftime('%d-%m-%Y')
            hora_ext = agora.strftime('%H-%M-%S')
            base_nome = f"Extração de Dados {data_ext} - {hora_ext}"

            self.caminho_excel_final = os.path.join(pasta_saida, f"{base_nome}.xlsx")
            self.caminho_excel_VISUALIZAÇÃO = os.path.join(pasta_saida, f"VISUALIZAÇÃO - {base_nome}.xlsx")

            self.excel = ExcelTempoReal(self.caminho_excel_final, self.caminho_excel_VISUALIZAÇÃO)
            print(f"\n📄 Excel FINAL: {self.caminho_excel_final}")
            print(f"👀 Excel VISUALIZAÇÃO (pode abrir): {self.caminho_excel_VISUALIZAÇÃO}")

            total = 0
            pagina = 1

            while True:
                linhas = self.obter_linhas()
                if not linhas:
                    print("⚠️ Nenhum candidato encontrado nesta página.")
                    break

                print(f"\n📄 PÁGINA {pagina} - {len(linhas)} candidato(s)")

                for idx in range(len(linhas)):
                    # Re-busca a lista pra evitar stale
                    linhas = self.obter_linhas()
                    if idx >= len(linhas):
                        break

                    total += 1
                    print(f"\n🔍 Candidato #{total} (Pág {pagina}, Item {idx+1}/{len(linhas)})")

                    try:
                        self.abrir_candidato(linhas[idx])
                        registro = self.extrair_um()

                        nome = (registro.get('Nome') or '').strip()
                        print(f"👤 {nome if nome else 'Nome não encontrado'}")

                        self.excel.adicionar_linha(registro)

                    except Exception as e:
                        print(f"❌ Erro no candidato #{total}: {e}")
                    finally:
                        self.voltar_lista()

                print("\n🔍 Verificando próxima página...")
                if not self.proxima_pagina():
                    break

                pagina += 1

            print("\n============================================================")
            print(f"✅ FINALIZADO! Total extraído: {total}")
            print(f"📄 Excel FINAL: {os.path.abspath(self.caminho_excel_final)}")
            print("============================================================\n")

        except KeyboardInterrupt:
            print("\n🛑 Execução interrompida pelo usuário (Ctrl+C).")
            if self.caminho_excel_final:
                print(f"📄 Excel FINAL (parcial): {os.path.abspath(self.caminho_excel_final)}")

        finally:
            # Ao final (ou interrupção), remove o(s) arquivo(s) de visualização
            try:
                if self.excel:
                    self.excel.limpar_visualizacoes()
            except:
                pass
            self.encerrar()



#########################################################
# Função principal de entrada do script                  #
#########################################################
def main():
    try:
        HabibotBot().executar()
    except SessionNotCreatedException as e:
        msg = str(e)
        print("\n❌ Não foi possível iniciar o navegador (incompatibilidade Chrome/ChromeDriver).")
        print("\nO que isso significa:")
        print("- O Chrome instalado não combina com o ChromeDriver disponível.")
        print("\nComo resolver:")
        print("- Atualize o Google Chrome e tente novamente.")
        print("- Ou gere um novo EXE nesta máquina (ele traz um driver compatível com o seu Chrome).")
        print("- Ou forneça um chromedriver compatível e configure Habibot_CHROMEDRIVER apontando para ele.")
        print(f"\nDetalhe técnico: {msg.splitlines()[0] if msg else 'SessionNotCreatedException'}")
    except WebDriverException as e:
        msg = str(e)
        msg_lower = msg.lower()

        print("\n❌ Falha ao iniciar/controlar o Google Chrome via automação.")

        if 'cannot find chrome binary' in msg_lower or 'chrome binary' in msg_lower:
            print("\nCausa provável: Google Chrome não está instalado (ou não foi encontrado).")
            print("\nComo resolver:")
            print("- Instale o Google Chrome (versão 64-bit) e tente novamente.")
        elif 'access is denied' in msg_lower or 'winerror 5' in msg_lower or 'permission' in msg_lower:
            print("\nCausa provável: permissão/antivírus bloqueando a execução do driver/navegador.")
            print("\nComo resolver:")
            print("- Execute como Administrador (teste).")
            print("- Verifique antivírus/EDR (Windows Defender, etc.) e libere o executável.")
            print("- Verifique se sua empresa bloqueia automação do navegador por política.")
        elif 'session not created' in msg_lower:
            print("\nCausa provável: incompatibilidade entre Chrome e ChromeDriver.")
            print("\nComo resolver:")
            print("- Atualize o Chrome e/ou gere novamente o EXE.")
        elif 'disconnected' in msg_lower or 'chrome not reachable' in msg_lower:
            print("\nCausa provável: o Chrome fechou/crashou durante a execução.")
            print("\nComo resolver:")
            print("- Feche instâncias do Chrome abertas e tente de novo.")
            print("- Reinicie o computador se continuar acontecendo.")
        else:
            print("\nCausa provável: bloqueio/instalação do Chrome/driver ou política do ambiente.")
            print("\nComo resolver:")
            print("- Confirme que o Chrome está instalado.")
            print("- Tente rodar em uma máquina sem restrições de antivírus/políticas.")

        print(f"\nDetalhe técnico: {msg.splitlines()[0] if msg else 'WebDriverException'}")
    except OSError as e:
        # Em alguns ambientes, bloqueios de antivírus/SmartScreen aparecem como OSError/WinError.
        texto = str(e)
        texto_lower = texto.lower()
        winerror = getattr(e, 'winerror', None)

        print("\n❌ Não foi possível iniciar um componente necessário (possível bloqueio do Windows/antivírus).")

        if winerror == 225 or 'winerror 225' in texto_lower or 'contains a virus' in texto_lower or 'potentially unwanted' in texto_lower:
            print("\nCausa provável: o antivírus/Windows Defender bloqueou o executável ou o chromedriver.")
            print("\nComo resolver:")
            print("- Abra Segurança do Windows → Proteção contra vírus e ameaças → Histórico de proteção e permita/restaure o bloqueio.")
            print("- Adicione uma exclusão para a pasta do bot (apenas se sua política permitir).")
            print("- Se for ambiente corporativo (EDR), solicite liberação ao TI.")
        elif winerror == 193 or 'winerror 193' in texto_lower or '%1 is not a valid win32 application' in texto_lower:
            print("\nCausa provável: incompatibilidade de arquitetura (32/64 bits) ou arquivo corrompido.")
            print("\nComo resolver:")
            print("- Confirme que o Windows é 64-bit.")
            print("- Baixe/copie novamente o EXE.")
            print("- Gere o EXE novamente nesta máquina (build limpo).")
        elif winerror == 740 or 'winerror 740' in texto_lower or 'elevation' in texto_lower:
            print("\nCausa provável: o Windows exigiu elevação (Executar como Administrador).")
            print("\nComo resolver:")
            print("- Clique com botão direito no EXE e selecione 'Executar como administrador'.")
        elif '0xc0000142' in texto_lower or '0xc000007b' in texto_lower:
            print("\nCausa provável: dependência do sistema/compatibilidade (DLL/VC++), ou bloqueio do ambiente.")
            print("\nComo resolver:")
            print("- Execute Windows Update.")
            print("- Instale/atualize o Microsoft Visual C++ Redistributable (x64), se permitido.")
            print("- Verifique se o antivírus/EDR está bloqueando a execução.")
        else:
            print("\nCausa provável: bloqueio por política/antivírus ou ambiente restrito.")
            print("\nComo resolver:")
            print("- Tente executar em uma pasta local (ex: Desktop) e não em rede.")
            print("- Verifique bloqueios do SmartScreen/antivírus.")
            print("- Se for empresa, valide com TI/segurança.")

        print(f"\nDetalhe técnico: {texto.splitlines()[0] if texto else 'OSError'}")
    except PermissionError as e:
        print("\n❌ Erro de permissão ao acessar arquivos.")
        print("\nCausa provável:")
        print("- Um arquivo Excel (principalmente o FINAL) está aberto no Excel e foi bloqueado para escrita.")
        print("\nComo resolver:")
        print("- Feche o Excel/arquivo e tente novamente.")
        print(f"\nDetalhe técnico: {e}")
    except Exception as e:
        print("\n❌ Ocorreu um erro inesperado.")
        print("\nComo resolver (rápido):")
        print("- Tente novamente.")
        print("- Se persistir, envie o detalhe técnico abaixo.")
        detalhe = traceback.format_exc().strip().splitlines()
        if detalhe:
            print("\nDetalhe técnico (resumo):")
            print(detalhe[-1])
        else:
            print(f"\nDetalhe técnico: {e}")


if __name__ == '__main__':
    main()
