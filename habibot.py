"""Habibot.

Este bot realiza a extração automatizada dos candidatos cadastrados no
Sistema de Habitação, salvando os dados em um relatório Excel.

"""


import os
import shutil
import subprocess
import sys
import time
import traceback
import argparse
import signal
import textwrap
import re

from datetime import datetime
from getpass import getpass
from pathlib import Path
from dataclasses import dataclass
import pandas as pd
import pwinput


class OrdemLeituraLogger:
    """Logger simples (append) para registrar a ordem de leitura de campos.

    Formato TSV (tab-separated) para fácil colar/filtrar no Excel.
    """

    def __init__(self, caminho: str):
        self.caminho = str(caminho)
        self._seq = 0
        self._t0 = time.monotonic()
        self._garantir_header()

    def _garantir_header(self):
        try:
            pasta = os.path.dirname(self.caminho)
            if pasta and not os.path.exists(pasta):
                os.makedirs(pasta, exist_ok=True)
        except:
            pass

        try:
            if not os.path.exists(self.caminho) or os.path.getsize(self.caminho) == 0:
                with open(self.caminho, 'a', encoding='utf-8', newline='') as f:
                    f.write("seq\tts\trel_ms\taba\tsecao\tcampo\tmetodo\tconsulta\tfound\tqtd_matches\tvalor\tobs\n")
        except:
            pass

    @staticmethod
    def _one_line(v: str, *, max_len: int = 180) -> str:
        t = (v or '')
        t = t.replace('\r', ' ').replace('\n', ' \\n ')
        t = re.sub(r"\s+", " ", t).strip()
        if len(t) > max_len:
            t = t[: max_len - 1] + "…"
        return t

    def log(
        self,
        *,
        aba: str,
        secao: str,
        campo: str,
        metodo: str,
        consulta: str,
        found: bool,
        qtd_matches: int | None = None,
        valor: str | None = None,
        obs: str | None = None,
    ):
        try:
            self._seq += 1
            ts = datetime.now().strftime('%d-%m-%Y %H:%M:%S.%f')[:-3]
            rel_ms = int((time.monotonic() - self._t0) * 1000)
            qtd = '' if qtd_matches is None else str(int(qtd_matches))
            val = '' if valor is None else self._one_line(str(valor))
            ob = '' if obs is None else self._one_line(str(obs), max_len=220)
            linha = (
                f"{self._seq}\t{ts}\t{rel_ms}\t{self._one_line(aba)}\t{self._one_line(secao)}\t{self._one_line(campo)}\t"
                f"{self._one_line(metodo)}\t{self._one_line(consulta)}\t{('1' if found else '0')}\t{qtd}\t{val}\t{ob}\n"
            )
            with open(self.caminho, 'a', encoding='utf-8', newline='') as f:
                f.write(linha)
        except:
            pass


class AcoesLogger:
    """Logger simples (append) para registrar ações do bot/extrator.

    Formato TSV (tab-separated) para fácil colar/filtrar no Excel.
    """

    def __init__(self, caminho: str):
        self.caminho = str(caminho)
        self._seq = 0
        self._t0 = time.monotonic()
        self._garantir_header()

    def _garantir_header(self):
        try:
            pasta = os.path.dirname(self.caminho)
            if pasta and not os.path.exists(pasta):
                os.makedirs(pasta, exist_ok=True)
        except:
            pass

        try:
            if not os.path.exists(self.caminho) or os.path.getsize(self.caminho) == 0:
                with open(self.caminho, 'a', encoding='utf-8', newline='') as f:
                    f.write("seq\tts\trel_ms\tarea\tacao\tdetalhe\turl\tok\tdur_ms\tobs\n")
        except:
            pass

    @staticmethod
    def _one_line(v: str, *, max_len: int = 240) -> str:
        t = (v or '')
        t = t.replace('\r', ' ').replace('\n', ' \\n ')
        t = re.sub(r"\s+", " ", t).strip()
        if len(t) > max_len:
            t = t[: max_len - 1] + "…"
        return t

    def log(
        self,
        *,
        area: str,
        acao: str,
        detalhe: str | None = None,
        url: str | None = None,
        ok: bool | None = None,
        dur_ms: int | None = None,
        obs: str | None = None,
    ):
        try:
            self._seq += 1
            ts = datetime.now().strftime('%d-%m-%Y %H:%M:%S.%f')[:-3]
            rel_ms = int((time.monotonic() - self._t0) * 1000)
            u = '' if url is None else self._one_line(str(url), max_len=320)
            det = '' if detalhe is None else self._one_line(str(detalhe), max_len=320)
            o = '' if obs is None else self._one_line(str(obs), max_len=380)
            okv = '' if ok is None else ('1' if ok else '0')
            dur = '' if dur_ms is None else str(int(dur_ms))
            linha = (
                f"{self._seq}\t{ts}\t{rel_ms}\t{self._one_line(area)}\t{self._one_line(acao)}\t{det}\t{u}\t{okv}\t{dur}\t{o}\n"
            )
            with open(self.caminho, 'a', encoding='utf-8', newline='') as f:
                f.write(linha)
        except:
            pass

_DEPENDENCIAS_FALTANDO: list[str] = []

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.common.exceptions import WebDriverException, SessionNotCreatedException, TimeoutException
    from selenium.webdriver import ActionChains
except ModuleNotFoundError:
    webdriver = None
    By = None
    Keys = None
    WebDriverWait = None
    EC = None
    Service = None
    Options = None
    ActionChains = None
    WebDriverException = type('WebDriverException', (Exception,), {})
    SessionNotCreatedException = type('SessionNotCreatedException', (Exception,), {})
    TimeoutException = type('TimeoutException', (Exception,), {})
    _DEPENDENCIAS_FALTANDO.append('selenium')

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ModuleNotFoundError:
    ChromeDriverManager = None
    _DEPENDENCIAS_FALTANDO.append('webdriver_manager')

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
    from openpyxl.drawing.image import Image as OpenpyxlImage
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError:
    Workbook = None
    load_workbook = None
    Alignment = None
    PatternFill = None
    Font = None
    Border = None
    Side = None
    OpenpyxlImage = None
    get_column_letter = None
    _DEPENDENCIAS_FALTANDO.append('openpyxl')


# --- Função para ler metadados do version_info.txt ---
import re
def ler_version_info(path=None):
    path = path or (Path(__file__).parent / 'assets' / 'build' / 'version_info.txt')
    info = {
        'app_name': None,
        'author': None,
        'language': None,
        'version': None,
        'release_date': None,
        'description': None,
    }
    try:
        with open(path, encoding='utf-8') as f:
            txt = f.read()
        # Extrai campos do version_info.txt
        def extrai(campo, regex):
            m = re.search(regex, txt)
            return m.group(1).strip() if m else None
        info['app_name'] = extrai('ProductName', r"StringStruct\('ProductName',\s*'([^']+)'\)")
        info['author'] = extrai('CompanyName', r"StringStruct\('CompanyName',\s*'([^']+)'\)")
        info['version'] = extrai('ProductVersion', r"StringStruct\('ProductVersion',\s*'([^']+)'\)")
        info['description'] = extrai('FileDescription', r"StringStruct\('FileDescription',\s*'([^']+)'\)")
        info['language'] = extrai('Comments', r"StringStruct\('Comments',\s*'Idioma: ([^']+)'\)") or 'pt-BR'
        # Data de release pode ser extraída de LegalCopyright ou setada manualmente
        info['release_date'] = None
        m = re.search(r"LegalCopyright', '© ([0-9]{4}) ([^']+)'", txt)
        if m:
            info['release_date'] = f"{m.group(1)}"
        return info
    except Exception:
        return info

_version_info = ler_version_info()
__app_name__ = _version_info['app_name'] or "Bot de Automação Habibot"
__author__ = _version_info['author'] or "Bruno Lima"
__language__ = _version_info['language'] or "pt-BR"
__version__ = _version_info['version'] or "1.1.0"
__release_date__ = _version_info['release_date'] or "2026"
__description__ = _version_info['description'] or (
    "Este bot de automação foi criado para suprir a necessidade de um relatório completo no Sistema de Habitação que disponibilize todos os dados dos candidatos cadastrados."
)


def _diretorio_base_execucao() -> Path:
    """Retorna o diretório base do script/EXE."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _encontrar_chromedriver_local() -> str | None:
    """Procura um chromedriver.exe local/embarcado para execução offline."""
    is_windows = os.name == 'nt'
    env = os.getenv('HABIBOT_CHROMEDRIVER')
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
        base / 'chromedriver',
        base / 'drivers' / 'chromedriver.exe',
        base / 'drivers' / 'chromedriver',
        base / 'assets' / 'drivers' / 'chromedriver.exe',
        base / 'assets' / 'drivers' / 'chromedriver',
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
            if (not is_windows) and str(p).lower().endswith('.exe'):
                continue
            if p.exists():
                return str(p)
        except:
            continue

    return None


def _terminal_width(padrao: int = 100) -> int:
    try:
        return max(60, shutil.get_terminal_size((padrao, 20)).columns)
    except:
        return max(60, padrao)


SEPARADOR_TELA = "─"


def _linha(char: str = SEPARADOR_TELA, largura: int | None = None) -> str:
    w = largura or _terminal_width()
    return char * w


def _extrair_nome_sistema(url: str) -> str:
    # Extrai o nome do sistema a partir do domínio (ex: app.habisoft.com.br -> HABISOFT)
    if not url:
        return "SISTEMA"
    url = url.lower().replace('https://', '').replace('http://', '').split('/')[0]
    partes = url.split('.')
    # Se for domínio com 3 ou mais partes, pega a antepenúltima (ex: app.habisoft.com.br -> habisoft)
    if len(partes) >= 3:
        nome = partes[-3]
    elif len(partes) == 2:
        nome = partes[0]
    else:
        nome = partes[0]
    return nome.strip().replace('-', ' ').replace('_', ' ').capitalize()


def mostrar_boas_vindas():
    w = _terminal_width()

    titulo = f"🤖 🏠  Bot de Automação Habibot 😁"
    linhas = [
        f"Criado por: {__author__}",
        f"Versão: {__version__}",
        F"Atualizado em: 22/01/2026",
        "",
        f"🤖 Este bot de automação foi criado para suprir a necessidade de um relatório completo no Sistema de Habitação que disponibilize todos os dados dos candidatos cadastrados.",
        "",
        "⚠️  Atenção!",
        "Como o bot acessa cada candidato individualmente, o processo pode gerar muitas requisições em pouco tempo. Em alguns cenários, o servidor pode interpretar isso como atividade suspeita e bloquear temporariamente o acesso.",
        "",
        "🔎 Acompanhar o progresso",
        "Enquanto a extração estiver em andamento, você pode acompanhar o progresso abrindo o arquivo de visualização dentro da pasta \"Extração de Dados\" (arquivo com prefixo \"VISUALIZAÇÃO\"). Para atualizar, feche e abra novamente.",
        "",
    ]

    print("")
    print(_linha(SEPARADOR_TELA, w))
    print("")
    print(titulo.center(w))
    print(_linha(SEPARADOR_TELA, w))
    print("")
    for l in linhas:
        if not l:
            print("")
            continue
        for parte in textwrap.wrap(l, width=w, break_long_words=False, break_on_hyphens=False):
            print(parte)
    print("")
    print(_linha(SEPARADOR_TELA, w))
    print("")


def _norm_texto_chave(texto: str) -> str:
    t = (texto or '').strip().lower()
    t = (
        t.replace('á', 'a').replace('à', 'a').replace('ã', 'a').replace('â', 'a')
        .replace('é', 'e').replace('ê', 'e')
        .replace('í', 'i')
        .replace('ó', 'o').replace('ô', 'o').replace('õ', 'o')
        .replace('ú', 'u')
        .replace('ç', 'c')
    )
    t = t.replace('*', ' ')
    t = t.replace(':', ' ')
    t = t.replace('º', ' ')
    t = t.replace('°', ' ')
    t = re.sub(r"\s*\(.*?\)\s*", " ", t)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


@dataclass(frozen=True)
class ColSpec:
    aba: str
    secao: str
    campo: str
    key: str


# Schema do Excel alinhado com "Sistema/Dados que precisam ser extraídos.txt" (ordem e nomes).
SCHEMA_ESTRUTURA: list[tuple[str, list[tuple[str, list[str]]]]] = [
    (
        'Titular',
        [
            ('Dados do Empreendimento', ['Loteamento']),
            (
                'Dados Gerais',
                [
                    'Nome',
                    'CPF/CNPJ',
                    'Naturalidade (Cidade)',
                    'UF Naturalidade (Estado)',
                    'Nacionalidade (País)',
                    'Data de Nascimento',
                    'Raça',
                    'Gênero',
                    'Contrato Distratado ou Rescindido Involutariamente',
                    'Recebe Atendimento Socio-assistencial do Município',
                    'Situação Atual do Domicílio',
                    'Tipo de Beneficiário Programa Social',
                    'Proprietário do Imóvel?',
                    'Recebeu/Participou de REURB?',
                ],
            ),
            (
                'Informações do CadÚnico',
                [
                    'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                    'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Programas CAD Único',
                    'Tratamento Oncológico/Hemodialise',
                ],
            ),
            (
                'Cadastro Preferencial',
                ['Cadastro Preferencial', 'Deficiência', 'Doença', 'Possui Microcefalia', 'Tem Veículo Próprio'],
            ),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)'],
            ),
            (
                'Situação Marital',
                ['Estado Civil', 'Data do Casamento/União', 'Regime do Casamento', 'Regime do Casamento (se Outro)'],
            ),
            (
                'CAD',
                [
                    'Data de Emissão',
                    'UF Título de Eleitor',
                    'Cidade do Título de Eleitor',
                ],
            ),
            (
                'Dados de Contato',
                ['Email', 'Telefone', 'Telefone Celular', 'Telefone Comercial', 'Telefone Para Recados'],
            ),
            ('Anotações', ['Parecer Social', 'Diagnóstico Social', 'Observações']),
        ],
    ),
    (
        'Segundo Titular',
        [
            (
                'Dados Gerais',
                [
                    'Nome',
                    'CPF/CNPJ',
                    'Naturalidade (Cidade)',
                    'UF Naturalidade (Estado)',
                    'Nacionalidade (País)',
                    'Data de Nascimento',
                    'Raça',
                    'Gênero',
                    'Contrato Distratado ou Rescindido Involutariamente',
                    'Recebe Atendimento Socio-assistencial do Município',
                    'Situação Atual do Domicílio',
                    'Tipo de Beneficiário Programa Social',
                    'Proprietário do Imóvel?',
                    'Recebeu/Participou de REURB?',
                ],
            ),
            (
                'Informações do CadÚnico',
                [
                    'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                    'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                    'Programas CAD Único',
                    'Tratamento Oncológico/Hemodialise',
                ],
            ),
            (
                'Cadastro Preferencial',
                ['Cadastro Preferencial', 'Deficiência', 'Doença', 'Possui Microcefalia', 'Tem Veículo Próprio'],
            ),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)'],
            ),
            (
                'Situação Marital',
                ['Estado Civil', 'Data do Casamento/União', 'Regime do Casamento', 'Regime do Casamento (se Outro)'],
            ),
            (
                'CAD',
                [
                    'Data de Emissão',
                    'UF Título de Eleitor',
                    'Cidade do Título de Eleitor',
                ],
            ),
            (
                'Dados de Contato',
                ['Email', 'Telefone', 'Telefone Celular', 'Telefone Comercial', 'Telefone Para Recados'],
            ),
            ('Anotações', ['Parecer Social', 'Diagnóstico Social', 'Observações']),
        ],
    ),
    (
        'Composição Familiar',
        [
            (
                'Dados Gerais',
                [
                    'Tipo',
                    'Nome',
                    'CPF',
                    'Naturalidade',
                    'UF Naturalidade',
                    'Nacionalidade',
                    'Data de Nascimento',
                    'Raça',
                    'Gênero',
                    'Telefone',
                    'Estado Civil',
                    'Pessoa Negra?',
                    'Pessoa Quilombola?',
                    'Pessoa LGBTQIAPN+?',
                    'Situação de Rua?',
                ],
            ),
            (
                'Educação e Profissão',
                ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço'],
            ),
            ('Cadastro Preferencial', ['Cadastro Preferencial', 'Deficiência', 'Doença']),
        ],
    ),
    (
        'Renda',
        [
            ('Dados da Renda', ['Tipo de Renda', 'Pessoa', 'Valor da Renda', 'Tipo se Apto Para Reurb']),
        ],
    ),
    (
        'Endereço',
        [
            (
                'Dados do Endereço',
                [
                    'Lote',
                    'Quadra',
                    'CEP',
                    'UF',
                    'Cidade',
                    'Bairro',
                    'Endereço',
                    'Número',
                    'Complemento',
                    'Tipo de Moradia',
                    'Possui outro imóvel?',
                    'Forma de aquisição',
                    'Relação com o imóvel',
                    'Uso do imóvel',
                    'Há famílias que habitam ou trabalham a, no máximo, 28 km de distância do Centro do Empreendimento?',
                    'Reside no município desde (ano)',
                ],
            ),
        ],
    ),
    (
        'Questionário',
        [
            (
                'Perguntas',
                [
                    '1. A mulher é a responsável pela unidade familiar?',
                    '2. Há pessoa negra na composição familiar?',
                    '3. Há pessoa com deficiência na composição familiar, comprovada por avaliação biopsicossocial (Lei nº 13.146/2015 e Decreto nº 11.063/2022)?',
                    '4. Há idoso na composição familiar, comprovado por documento civil com data de nascimento?',
                    '5. Há criança ou adolescente na composição familiar, comprovado por certidão de nascimento, guarda ou tutela?',
                    '6. Há pessoa com câncer ou doença rara crônica e degenerativa na família, comprovado por laudo médico?',
                    '7. Há mulheres vítimas de violência doméstica/familiar na família, comprovado por registro no Cadastro Nacional de Violência Doméstica (Lei Maria da Penha)?',
                    '8. Há integrantes de povos indígenas ou quilombolas na família, declarados no CadÚnico?',
                    '9. A família reside em área de risco (deslizamentos, inundações etc.), conforme mapeamento do PMRR, CPRM ou Defesa Civil?',
                    '10. O beneficiário teve contrato distratado ou rescindido involuntariamente, conforme normativo do Ente Público?',
                    '11. Atualmente é atendido pelas redes Socioassistenciais do Município?',
                ],
            ),
        ],
    ),
]


def _flatten_schema(schema: list[tuple[str, list[tuple[str, list[str]]]]]) -> list[ColSpec]:
    specs: list[ColSpec] = []
    for aba, secoes in schema:
        for secao, campos in secoes:
            for campo in campos:
                key = f"{aba}|||{secao}|||{campo}"
                specs.append(ColSpec(aba=aba, secao=secao, campo=campo, key=key))
    return specs


COL_SPECS: list[ColSpec] = _flatten_schema(SCHEMA_ESTRUTURA)
COL_KEY_BY_TRIPLE: dict[tuple[str, str, str], str] = {(c.aba, c.secao, c.campo): c.key for c in COL_SPECS}


_ABORT_REQUESTED = False


def _request_abort():
    global _ABORT_REQUESTED
    _ABORT_REQUESTED = True


def _check_abort():
    if _ABORT_REQUESTED:
        raise KeyboardInterrupt


def _col_key(aba: str, secao: str, campo: str) -> str:
    return COL_KEY_BY_TRIPLE.get((aba, secao, campo), f"{aba}|||{secao}|||{campo}")


class ExcelTempoReal:
    """Grava o Excel final formatado e mantém um arquivo de visualização em texto."""

    def __init__(self, caminho_final: str, caminho_visualizacao: str, nome_sistema: str = "SISTEMA", flush_cada_registros: int = 5, flush_interval_segundos: float = 10.0):
        self.caminho_final = caminho_final
        self.caminho_visualizacao_txt = caminho_visualizacao
        self.nome_sistema = nome_sistema
        self._contador_registros = 0

        self._wb = None
        self._ws = None

        self._flush_cada_registros = max(1, int(flush_cada_registros or 1))
        self._flush_interval_segundos = max(1.0, float(flush_interval_segundos or 10.0))
        self._ultimo_save = time.monotonic()
        self._pendente_salvar = False

        if not os.path.exists(self.caminho_final):
            self._inicializar_arquivo(self.caminho_final)

        self._wb = load_workbook(self.caminho_final)
        self._ws = self._wb['Dados'] if 'Dados' in self._wb.sheetnames else self._wb.active

        self._atualizar_visualizacao()

    def _inicializar_arquivo(self, caminho_arquivo: str):
        nome_sistema = self.nome_sistema or "SISTEMA"
        pasta = os.path.dirname(caminho_arquivo)
        if pasta and not os.path.exists(pasta):
            os.makedirs(pasta)

        wb = Workbook()
        ws = wb.active
        ws.title = 'Dados'

        # Deixa 5 linhas vazias para o cabeçalho personalizado
        for _ in range(5):
            ws.append([]) 

        # Cabeçalhos da tabela nas linhas 6, 7 e 8
        headers_aba = [c.aba for c in COL_SPECS]
        headers_secao = [c.secao for c in COL_SPECS]
        headers_campo = [c.campo for c in COL_SPECS]

        ws.append(headers_aba)    # Linha 6
        ws.append(headers_secao)  # Linha 7
        ws.append(headers_campo)  # Linha 8

        # --- MESCLAGEM DOS CABEÇALHOS DA TABELA ---
        
        # Mesclar Abas (Linha 6)
        start = 1
        cur_aba = headers_aba[0]
        for idx, aba in enumerate(headers_aba, start=1):
            if aba != cur_aba:
                if (idx - 1) > start:
                    ws.merge_cells(start_row=6, start_column=start, end_row=6, end_column=idx - 1)
                start = idx
                cur_aba = aba
        if len(headers_aba) >= start:
             ws.merge_cells(start_row=6, start_column=start, end_row=6, end_column=len(headers_aba))

        # Mesclar Seções (Linha 7)
        start = 1
        cur_sec = headers_secao[0]
        cur_aba_ref = headers_aba[0]
        for idx, secao in enumerate(headers_secao, start=1):
            aba_atual = headers_aba[idx-1]
            if secao != cur_sec or aba_atual != cur_aba_ref:
                if (idx - 1) > start:
                    ws.merge_cells(start_row=7, start_column=start, end_row=7, end_column=idx - 1)
                start = idx
                cur_sec = secao
                cur_aba_ref = aba_atual
        if len(headers_secao) >= start:
             ws.merge_cells(start_row=7, start_column=start, end_row=7, end_column=len(headers_secao))

        # --- CABEÇALHO PERSONALIZADO (Linhas 1-5) ---
        ws['A2'] = f"DADOS DOS CANDIDATOS CADASTRADOS - {nome_sistema.upper()}"
        ws['A3'] = "PREFEITURA MUNICIPAL DE VITÓRIA DA CONQUISTA"
        ws['A4'] = f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}. Fonte: {nome_sistema.title()}"

        # Mesclar cabeçalho principal
        ultima_coluna = get_column_letter(len(COL_SPECS))
        try:
            ws.merge_cells(f'A1:{ultima_coluna}1')
            ws.merge_cells(f'A2:{ultima_coluna}2')
            ws.merge_cells(f'A3:{ultima_coluna}3')
            ws.merge_cells(f'A4:{ultima_coluna}4')
            ws.merge_cells(f'A5:{ultima_coluna}5')
        except:
            pass

        if Font:
            ws['A2'].font = Font(name='Calibri', size=14, bold=True)
            ws['A3'].font = Font(name='Calibri', size=12, bold=True)
            ws['A4'].font = Font(name='Calibri', size=10, italic=True)
            ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
            ws['A3'].alignment = Alignment(horizontal='center', vertical='center')
            ws['A4'].alignment = Alignment(horizontal='center', vertical='center')

        # Inserir Logo
        caminho_logo = _diretorio_base_execucao() / "assets" / "images" / "logos" / "logo.png"
        if os.path.exists(caminho_logo) and OpenpyxlImage:
            try:
                img = OpenpyxlImage(str(caminho_logo))
                img.width = 180
                img.height = 70
                ws.add_image(img, 'A1')
            except Exception:
                pass

        # --- ESTILIZAÇÃO GERAL DOS CABEÇALHOS (Linhas 6-8) ---
        fill_azul = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid") if PatternFill else None
        fonte_branca = Font(bold=True, color="FFFFFF") if Font else None
        borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')) if Border else None
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True) if Alignment else None

        for row in ws.iter_rows(min_row=6, max_row=8, min_col=1, max_col=len(COL_SPECS)):
            for cell in row:
                if align_center: cell.alignment = align_center
                if fonte_branca: cell.font = fonte_branca
                if fill_azul: cell.fill = fill_azul
                if borda: cell.border = borda

        # Congelar painéis (título + cabeçalhos fixos)
        ws.freeze_panes = 'A9'

        # Configuração de página: A4, horizontal, margens 0,5cm, ajustar colunas em uma página
        try:
            ws.page_setup.orientation = 'landscape'
            ws.page_setup.paperSize = ws.page_setup.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1  # Ajusta todas as colunas em uma página
            ws.page_setup.fitToHeight = 0  # Altura livre (pode quebrar em várias páginas)
            margem = 0.5 / 2.54  # cm -> polegadas
            ws.page_margins.left = margem
            ws.page_margins.right = margem
            ws.page_margins.top = margem
            ws.page_margins.bottom = margem
        except Exception:
            pass

        # Ajuste de largura inicial
        for col_idx in range(1, len(COL_SPECS) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 25

        # Salva de forma atômica para reduzir risco de corrupção em encerramentos abruptos
        tmp = caminho_arquivo + ".tmp"
        try:
            wb.save(tmp)
            os.replace(tmp, caminho_arquivo)
        finally:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except:
                pass

    @staticmethod
    def _normalizar(valor):
        if valor is None:
            return ''
        if isinstance(valor, str):
            txt = valor.strip()
            if not txt:
                return ''
            norm = txt.lower()
            placeholders = {
                'selecione',
                'selecione um estado',
                'selecione o estado',
                'selecione a escolaridade',
                'selecione a situação de emprego',
                'selecione o tempo de serviço',
                'selecione o regime de casamento',
                'selecione o estado civil',
                'selecione um tipo',
                'selecione a uf',
                'selecione o estado civil',
                'selecione o tempo de serviço',
                'selecione o estado',
                'selecione a situacao de emprego',
                'selecione a situacao',
                'nenhum',
                'n/a',
            }
            if norm in placeholders or norm.startswith('selecione'):
                return ''
            return txt
        txt = str(valor).strip()
        return txt if txt else ''

    def adicionar_linha(self, registro: dict):
        linha = [self._normalizar(registro.get(c.key)) for c in COL_SPECS]

        try:
            # Append na próxima linha disponível
            self._ws.append(linha)
            self._contador_registros += 1

            # --- FORMATAÇÃO DA LINHA DE DADOS (ZEBRA + ALTURA + ALINHAMENTO) ---
            nova_linha_idx = self._ws.max_row
            
            # Altura 40px
            self._ws.row_dimensions[nova_linha_idx].height = 40

            # Zebra Striping
            fill_cinza = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            fill_branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            
            # Conta índice de dados (linha 9 é a primeira de dados, indice 0)
            idx_dado = nova_linha_idx - 9 
            fill_atual = fill_cinza if idx_dado % 2 == 0 else fill_branco

            align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            for cell in self._ws[nova_linha_idx]:
                if Alignment: cell.alignment = align_center
                if PatternFill: cell.fill = fill_atual
                if Border: cell.border = borda
                # Se texto muito grande, reduz fonte (opcional)
                if cell.value and len(str(cell.value)) > 100 and Font:
                     cell.font = Font(size=8)

            self._atualizar_visualizacao(linha=linha)
            self._pendente_salvar = True

            if self._deve_salvar_agora():
                self._salvar_tentativa()
            return True

        except PermissionError:
            print("❌ O arquivo FINAL está aberto no Excel e bloqueou a escrita.")
            print(f"   Feche o arquivo e tente novamente: {self.caminho_final}")
            return False
        except Exception as e:
            print(f"Erro ao adicionar linha no Excel: {e}")
            raise

    def _deve_salvar_agora(self) -> bool:
        if not self._pendente_salvar:
            return False
        if self._contador_registros % self._flush_cada_registros == 0:
            return True
        if (time.monotonic() - self._ultimo_save) >= self._flush_interval_segundos:
            return True
        return False

    def _salvar_tentativa(self) -> bool:
        try:
            if self._wb is None:
                return False
            # Salva em arquivo temporário e substitui (atomic replace)
            tmp = self.caminho_final + ".tmp"
            self._wb.save(tmp)
            os.replace(tmp, self.caminho_final)
            self._ultimo_save = time.monotonic()
            self._pendente_salvar = False
            return True
        except PermissionError:
            print("❌ O arquivo FINAL está aberto no Excel e bloqueou a escrita.")
            print(f"   Feche o arquivo e o bot tentará salvar novamente: {self.caminho_final}")
            return False
        except Exception:
            return False
        finally:
            try:
                tmp = self.caminho_final + ".tmp"
                if os.path.exists(tmp):
                    os.remove(tmp)
            except:
                pass

    def flush(self):
        if self._pendente_salvar:
            self._salvar_tentativa()

    def _atualizar_visualizacao(self, linha: list[str] | None = None):
        try:
            destino = self.caminho_visualizacao_txt
            pasta = os.path.dirname(destino)
            if pasta and not os.path.exists(pasta):
                os.makedirs(pasta, exist_ok=True)

            if not hasattr(self, '_visualizacao_registros'):
                self._visualizacao_registros = []

            if linha is not None:
                self._visualizacao_registros.append(linha)

            total = len(self._visualizacao_registros)
            conteudo = []
            conteudo.append("Habibot - Visualização (completa)\n")
            conteudo.append(f"Gerado em: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}\n")
            conteudo.append(f"Arquivo FINAL: {os.path.abspath(self.caminho_final)}\n\n")

            # Agrupa campos por aba e seção para melhor visualização
            from collections import defaultdict
            for idx, reg in enumerate(self._visualizacao_registros, 1):
                conteudo.append(f"Candidato {idx}:")
                campos_agrupados = defaultdict(lambda: defaultdict(list))
                for i, c in enumerate(COL_SPECS):
                    valor = reg[i] if i < len(reg) else ""
                    campos_agrupados[c.aba][c.secao].append((c.campo, valor))

                for aba in campos_agrupados:
                    conteudo.append(f"  [{aba}]")
                    for secao in campos_agrupados[aba]:
                        conteudo.append(f"    - {secao}:")
                        for campo, valor in campos_agrupados[aba][secao]:
                            conteudo.append(f"        {campo}: {valor}")
                conteudo.append("")

            tmp = destino + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write("\n".join(conteudo))
            os.replace(tmp, destino)
            return destino
        except Exception:
            return None

    def limpar_visualizacoes(self):
        try:
            if self.caminho_visualizacao_txt and os.path.exists(self.caminho_visualizacao_txt):
                os.remove(self.caminho_visualizacao_txt)
        except:
            pass
        try:
            pasta = os.path.dirname(self.caminho_visualizacao_txt or '')
            if pasta and os.path.exists(pasta):
                for nome in os.listdir(pasta):
                    if nome.upper().startswith('VISUALIZAÇÃO'):
                        try:
                            os.remove(os.path.join(pasta, nome))
                        except:
                            pass
        except:
            pass


class ExtratorHabibot:
    def __init__(
        self,
        driver,
        wait,
        *,
        debug: bool = False,
        debug_visual: bool = False,
        debug_delay_segundos: float = 0.25,
        ordem_logger: OrdemLeituraLogger | None = None,
        acoes_logger: AcoesLogger | None = None,
    ):
        self.driver = driver
        self.wait = wait
        self.debug = bool(debug)
        self.debug_visual = bool(debug_visual)
        self.ordem_logger = ordem_logger
        self.acoes_logger = acoes_logger
        try:
            self.debug_delay_segundos = max(0.0, float(debug_delay_segundos or 0.0))
        except:
            self.debug_delay_segundos = 0.25

    def _log_acao(self, acao: str, *, detalhe: str | None = None, ok: bool | None = None, dur_ms: int | None = None, obs: str | None = None):
        if not self.acoes_logger:
            return
        try:
            url = ''
            try:
                url = self.driver.current_url
            except:
                url = ''
            self.acoes_logger.log(
                area='EXTRATOR',
                acao=str(acao),
                detalhe=detalhe,
                url=url,
                ok=ok,
                dur_ms=dur_ms,
                obs=obs,
            )
        except:
            pass

    def _dbg(self, msg: str):
        if self.debug:
            try:
                print(msg)
            except:
                pass

    def _mark_read(self, el, *, rotulo: str | None = None):
        """Marca um elemento como 'lido' com outline verde persistente (modo debug_visual)."""
        if not self.debug_visual:
            return
        try:
            if not el:
                return
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
            js = """
                const el = arguments[0];
                const label = arguments[1] || '';
                el.setAttribute('data-habibot-read', '1');
                el.style.outline = '3px solid #00ff00';
                el.style.outlineOffset = '2px';
                if (label) {
                    el.setAttribute('data-habibot-debug-read', label);
                }
            """
            self.driver.execute_script(js, el, rotulo or "")
        except:
            return

    def _highlight(self, el, *, rotulo: str | None = None):
        if not self.debug_visual:
            return
        try:
            if not el:
                return
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
            js = """
                const el = arguments[0];
                const label = arguments[1] || '';
                const prevOutline = el.style.outline;
                const prevOutlineOffset = el.style.outlineOffset;
                const prevBg = el.style.backgroundColor;
                el.style.outline = '3px solid #ff00ff';
                el.style.outlineOffset = '2px';
                if (label) {
                    el.setAttribute('data-habibot-debug', label);
                }
                setTimeout(() => {
                    try {
                        // Se já foi marcado como lido, volta para o verde.
                        if (el.getAttribute('data-habibot-read') === '1') {
                            el.style.outline = '3px solid #00ff00';
                            el.style.outlineOffset = '2px';
                        } else {
                            el.style.outline = prevOutline;
                            el.style.outlineOffset = prevOutlineOffset;
                            el.style.backgroundColor = prevBg;
                        }
                    } catch (e) {}
                }, Math.max(0, arguments[2] || 250));
            """
            self.driver.execute_script(js, el, rotulo or "", int(self.debug_delay_segundos * 1000))
            if self.debug_delay_segundos:
                time.sleep(self.debug_delay_segundos)
        except:
            return

    def _extrair_por_label_simples_found_ex(self, texto_label: str) -> tuple[bool, str, int, str]:
        """Versão estendida e CORRIGIDA para Grids (Colunas)."""
        _check_abort()
        try:
            texto_busca_norm = _norm_texto_chave(texto_label)
            if not texto_busca_norm: return False, "", 0, ""

            # 1. Busca todos os labels na tela que parecem com o texto
            labels = self.driver.find_elements(By.TAG_NAME, "label")
            candidatos = []
            for label in labels:
                try:
                    if not label.is_displayed(): continue
                    norm_lbl = _norm_texto_chave(label.text or "")
                    if not norm_lbl: continue
                    
                    # Pontuação: 0=Igual, 1=Começa com, 2=Contém
                    if norm_lbl == texto_busca_norm: score = 0
                    elif norm_lbl.startswith(texto_busca_norm): score = 1
                    elif texto_busca_norm in norm_lbl: score = 2
                    else: continue
                    candidatos.append((score, len(norm_lbl), label))
                except: continue

            # Tenta também encontrar textos soltos (span/div) que funcionam como label
            # Útil para perguntas longas que não usam a tag <label>
            if not candidatos:
                outros = self.driver.find_elements(By.CSS_SELECTOR, ".row div, .col-md-8, .col-form-label")
                for el in outros:
                    try:
                        if not el.is_displayed(): continue
                        txt = _norm_texto_chave(el.text)
                        if txt == texto_busca_norm:
                            candidatos.append((3, len(txt), el))
                    except: continue

            candidatos.sort(key=lambda x: (x[0], x[1]))
            qtd_matches = len(candidatos)

            # 2. Para cada label encontrado, tenta achar o campo input/select
            for _, _, label in candidatos:
                try:
                    self._highlight(label, rotulo=f"LABEL: {texto_label}")

                    # A) Tenta pelo atributo 'for' (o jeito certo do HTML)
                    try:
                        for_attr = label.get_attribute('for')
                        if for_attr:
                            campo = self.driver.find_element(By.ID, for_attr)
                            self._highlight(campo, rotulo="ACHEI (FOR)")
                            return True, self._valor_campo(campo), qtd_matches, label.text
                    except: pass

                    # B) Tenta irmão vizinho (quando estão colados)
                    try:
                        siblings = label.find_elements(By.XPATH, "following-sibling::*[self::input or self::select or self::textarea][1]")
                        if siblings:
                            self._highlight(siblings[0], rotulo="ACHEI (VIZINHO)")
                            return True, self._valor_campo(siblings[0]), qtd_matches, label.text
                    except: pass

                    # C) Tenta container PAI (quando estão numa div juntos)
                    try:
                        pai = label.find_element(By.XPATH, "..")
                        campos = pai.find_elements(By.CSS_SELECTOR, "input, select, textarea")
                        for c in campos:
                            if c.is_displayed():
                                self._highlight(c, rotulo="ACHEI (PAI)")
                                return True, self._valor_campo(c), qtd_matches, label.text
                    except: pass

                    # D) Tenta container AVÔ (A CORREÇÃO MÁGICA PARA SEU CASO)
                    # Isso resolve quando tem: <Linha> <Coluna1>Label</Coluna1> <Coluna2>Select</Coluna2> </Linha>
                    try:
                        avo = label.find_element(By.XPATH, "../..")
                        # Procura selects ou inputs dentro da linha inteira do avô
                        campos = avo.find_elements(By.CSS_SELECTOR, "select, input:not([type='hidden']), textarea")
                        for c in campos:
                            # Pega o primeiro campo visível que encontrar na mesma linha
                            if c.is_displayed():
                                self._highlight(c, rotulo="ACHEI (AVÔ/LINHA)")
                                return True, self._valor_campo(c), qtd_matches, label.text
                    except: pass

                except: continue

            return False, "", qtd_matches, ""
        except:
            return False, "", 0, ""

    def _extrair_por_label_simples_found(self, texto_label: str) -> tuple[bool, str]:
        encontrou, valor, _, _ = self._extrair_por_label_simples_found_ex(texto_label)
        return encontrou, valor

    def _log_leitura(self, *, aba: str, secao: str, campo: str, metodo: str, consulta: str, found: bool, qtd_matches: int | None = None, valor: str | None = None, obs: str | None = None):
        try:
            if self.ordem_logger:
                self.ordem_logger.log(
                    aba=aba,
                    secao=secao,
                    campo=campo,
                    metodo=metodo,
                    consulta=consulta,
                    found=bool(found),
                    qtd_matches=qtd_matches,
                    valor=valor,
                    obs=obs,
                )
        except:
            pass

    def _ler_campo_por_id(self, field_id: str) -> tuple[bool, str]:
        try:
            el = self.driver.find_element(By.ID, field_id)
            if el.is_displayed():
                return True, self._valor_campo(el)
        except:
            pass
        return False, ""

    def ler_campo_ordenado(
        self,
        *,
        aba: str,
        secao: str,
        campo: str,
        tentativas: list[tuple[str, str]],
    ) -> tuple[bool, str]:
        """Executa tentativas em ordem e registra no log.

        tentativas: lista de (metodo, consulta)
          - metodo 'id': consulta é o id do campo
          - metodo 'label': consulta é o texto do label (aceita contains)
        """
        for metodo, consulta in (tentativas or []):
            _check_abort()
            if metodo == 'id':
                found, valor = self._ler_campo_por_id(consulta)
                self._log_leitura(
                    aba=aba,
                    secao=secao,
                    campo=campo,
                    metodo='id',
                    consulta=consulta,
                    found=found,
                    qtd_matches=None,
                    valor=valor,
                )
                if found:
                    return True, valor
                continue

            if metodo == 'label':
                found, valor, qtd, _lbl = self._extrair_por_label_simples_found_ex(consulta)
                obs = None
                if qtd and qtd > 1:
                    obs = f"⚠️ {qtd} labels candidatos"
                self._log_leitura(
                    aba=aba,
                    secao=secao,
                    campo=campo,
                    metodo='label',
                    consulta=consulta,
                    found=found,
                    qtd_matches=qtd,
                    valor=valor,
                    obs=obs,
                )
                if found:
                    return True, valor
                continue

            # método desconhecido
            self._log_leitura(
                aba=aba,
                secao=secao,
                campo=campo,
                metodo=str(metodo),
                consulta=str(consulta),
                found=False,
                qtd_matches=None,
                valor="",
                obs="Método não suportado",
            )

        return False, ""

    def aguardar(self, segundos):
        self.driver.implicitly_wait(segundos)

    def clicar_aba(self, nome_aba: str) -> bool:
        _check_abort()
        t0 = time.monotonic()
        try:
            xpath = (
                f"//a[contains(text(), '{nome_aba}')] | "
                f"//button[contains(text(), '{nome_aba}')] | "
                f"//*[@role='tab' and contains(text(), '{nome_aba}')]"
            )
            abas = self.driver.find_elements(By.XPATH, xpath)
            for aba in abas:
                if aba.is_displayed():
                    self._dbg(f"🧭 Clicando aba: {nome_aba}")
                    self._highlight(aba, rotulo=f"ABA: {nome_aba}")
                    self.driver.execute_script("arguments[0].click();", aba)
                    self.aguardar(1)
                    self._log_acao('CLICAR_ABA', detalhe=nome_aba, ok=True, dur_ms=int((time.monotonic() - t0) * 1000))
                    return True
            self._log_acao('CLICAR_ABA', detalhe=nome_aba, ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs='Aba não encontrada/visível')
            return False
        except:
            self._log_acao('CLICAR_ABA', detalhe=nome_aba, ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs='Exceção ao clicar na aba')
            return False

    def _texto_limpo(self, texto: str) -> str:
        t = (texto or '').strip()
        t = re.sub(r"\s+", " ", t)
        return t

    def _valor_campo(self, campo) -> str:
        try:
            if not campo.is_displayed(): return ""
        except: return ""

        try:
            if campo.tag_name == 'input':
                tp = (campo.get_attribute('type') or '').strip().lower()
                if tp in ['checkbox', 'radio']:
                    try:
                        return 'Sim' if campo.is_selected() else 'Não'
                    except:
                        chk = (campo.get_attribute('checked') or '').strip().lower()
                        return 'Sim' if chk in ['true', 'checked', '1', 'on'] else 'Não'
            
            if campo.tag_name == 'select':
                try:
                    # Tenta pegar o texto visível da opção selecionada
                    return self._texto_limpo(self.driver.execute_script("return arguments[0].options[arguments[0].selectedIndex].text;", campo))
                except:
                    return ""

            if campo.tag_name == 'textarea':
                v = self._texto_limpo(campo.get_attribute('value') or campo.text or '')
                return v

            v = self._texto_limpo(campo.get_attribute('value') or '')
            return v
        except:
            return ""

    def _extrair_por_label_simples(self, texto_label: str) -> str:
        encontrou, valor = self._extrair_por_label_simples_found(texto_label)
        return valor

    def _extrair_por_id(self, field_id: str) -> str:
        try:
            el = self.driver.find_element(By.ID, field_id)
            if el.is_displayed():
                return self._valor_campo(el)
        except:
            pass
        return ""

    # --- EXTRAÇÃO PROFUNDA COM MODAL ---
    def _extrair_tabela_com_modal_edicao(self, campos_mapeados):
        dados_agregados = []
        try:
            # Localiza tabelas visíveis
            tabelas = self.driver.find_elements(By.TAG_NAME, "table")
            tabela_alvo = None
            for t in tabelas:
                if t.is_displayed():
                    tabela_alvo = t
                    break
            
            if not tabela_alvo: return ""

            linhas = tabela_alvo.find_elements(By.CSS_SELECTOR, "tbody tr")
            qtd_linhas = len(linhas)
            if qtd_linhas == 0: return ""

            for i in range(qtd_linhas):
                try:
                    # Refaz busca para garantir elemento fresco
                    tabelas = self.driver.find_elements(By.TAG_NAME, "table")
                    tabela_atual = None
                    for t in tabelas:
                        if t.is_displayed():
                            tabela_atual = t
                            break
                    if not tabela_atual: break

                    linhas_atual = tabela_atual.find_elements(By.CSS_SELECTOR, "tbody tr")
                    if i >= len(linhas_atual): break
                    linha = linhas_atual[i]

                    # Busca botão editar
                    btn_editar = None
                    try:
                        btn_editar = linha.find_element(By.CSS_SELECTOR, "button.btn-soft-info, .ri-pencil-line")
                    except: continue 
                    
                    self.driver.execute_script("arguments[0].click();", btn_editar)
                    time.sleep(1.5) 
                    
                    # Extrai dados
                    dados_item = []
                    for label_busca, nome_campo in campos_mapeados.items():
                        valor = self._extrair_por_label_simples(label_busca)
                        if valor:
                            dados_item.append(f"{nome_campo}: {valor}")
                    
                    if dados_item:
                        dados_agregados.append(" | ".join(dados_item))

                    # Fecha modal
                    try:
                        btns_fechar = self.driver.find_elements(By.CSS_SELECTOR, ".modal.show button.btn-close, .modal.show button.btn-light, button[data-bs-dismiss='modal']")
                        fechou = False
                        for btn in btns_fechar:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                fechou = True
                                break
                        if not fechou:
                             ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    except:
                        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    
                    time.sleep(1) 
                    
                except Exception as e:
                    try: ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    except: pass
                    time.sleep(1)
                    continue

        except Exception:
            pass
        return "\n".join(dados_agregados)

    def _extrair_tabela_com_modal_edicao_itens(self, campos_mapeados: dict[str, str], *, colunas_tabela: dict[int, str] | None = None) -> list[dict[str, str]]:
        itens: list[dict[str, str]] = []
        try:
            # Tenta encontrar a tabela com RE-TENTATIVAS (Blindagem)
            tabela_alvo = None
            for _ in range(3):
                try:
                    tabelas = self.driver.find_elements(By.TAG_NAME, "table")
                    for t in tabelas:
                        if t.is_displayed() and t.find_elements(By.CSS_SELECTOR, "tbody tr"):
                            tabela_alvo = t
                            break
                    if tabela_alvo: break
                    time.sleep(1)
                except: pass
            
            if not tabela_alvo: return []

            # Conta quantas linhas tem para processar
            qtd_linhas = len(tabela_alvo.find_elements(By.CSS_SELECTOR, "tbody tr"))

            for i in range(qtd_linhas):
                item = {}
                _check_abort()
                
                # RECARREGA A LINHA A CADA PASSADA (Fundamental para evitar StaleElement)
                try:
                    tabelas = self.driver.find_elements(By.TAG_NAME, "table")
                    t_atual = [t for t in tabelas if t.is_displayed() and t.find_elements(By.CSS_SELECTOR, "tbody tr")][0]
                    linha = t_atual.find_elements(By.CSS_SELECTOR, "tbody tr")[i]
                except: continue

                # 1. Tenta pegar dados que já estão na tela (Nome/CPF) caso o clique falhe
                try:
                    cols = linha.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 2:
                        txt0 = self._texto_limpo(cols[0].text)
                        txt1 = self._texto_limpo(cols[1].text)
                        if campos_mapeados:
                            # Preenche preventivamente
                            for k, v in campos_mapeados.items():
                                if "Nome" in v or "nome" in k: item[v] = txt0
                                if "CPF" in v or "cpf" in k: item[v] = txt1
                except: pass

                # 2. Busca o botão de Editar (Lápis)
                btn = None
                try:
                    # Procura por qualquer coisa clicável que pareça "Editar" ou tenha ícone de lápis
                    btn = linha.find_element(By.XPATH, ".//*[contains(@class, 'pencil') or contains(@class, 'edit') or contains(text(), 'Editar')]")
                    # Se achou o ícone <i>, pega o botão pai <button> ou <a>
                    if btn.tag_name == 'i': 
                        btn = btn.find_element(By.XPATH, "./..")
                except: pass

                # 3. Clica e Extrai
                if btn:
                    try:
                        self.driver.execute_script("arguments[0].click();", btn)
                        time.sleep(2.0) # Espera o modal abrir
                        
                        # Extrai os dados de dentro do modal
                        if campos_mapeados:
                            primeiro = next(iter(campos_mapeados.values()))
                            if isinstance(primeiro, list): # Novo formato
                                for col, labels in campos_mapeados.items():
                                    for lbl in labels:
                                        ok, val = self._extrair_por_label_simples_found(lbl)
                                        if ok: 
                                            item[col] = val
                                            break
                            else: # Formato antigo
                                for lbl, col in campos_mapeados.items():
                                    item[col] = self._extrair_por_label_simples(lbl)
                        
                        # Fecha o Modal
                        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.5)
                    except:
                        # Se der erro, garante que fecha o modal com ESC
                        try: ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                        except: pass
                
                itens.append(item)

        except Exception:
            pass # Segue o baile se der erro, não para o robô
        return itens

    @staticmethod
    def _itens_para_colunas_multilinha(itens: list[dict[str, str]], colunas: list[str]) -> dict[str, str]:
        """Converte lista de itens em um dict coluna->string com linhas separadas por \n."""
        saida: dict[str, str] = {}
        for col in colunas:
            linhas = [(it.get(col) or "") for it in (itens or [])]
            saida[col] = "\n".join(linhas).strip("\n")
        return saida

    # --- ABA: TITULAR ---
    def extrair_titular_mapa(self) -> dict:
        self.clicar_aba('Titular')
        # Espera carregar o básico
        try:
            WebDriverWait(self.driver, 5, poll_frequency=0.2).until(
                lambda d: self._extrair_por_label_simples('Nome')
            )
        except: pass
        
        mapa: dict[str, str] = {}

        # LISTA EXATA DO SEU NOVO SCHEMA
        passos: list[tuple[str, str, str, list[tuple[str, str]]]] = [
            # Dados do Empreendimento
            ('Titular', 'Dados do Empreendimento', 'Loteamento', [('label', 'Loteamento')]),

            # Dados Gerais
            ('Titular', 'Dados Gerais', 'Nome', [('label', 'Nome')]),
            ('Titular', 'Dados Gerais', 'CPF/CNPJ', [('label', 'CPF/CNPJ')]),
            ('Titular', 'Dados Gerais', 'Naturalidade (Cidade)', [('label', 'Naturalidade (Cidade)'), ('label', 'Naturalidade')]),
            ('Titular', 'Dados Gerais', 'UF Naturalidade (Estado)', [('label', 'UF Naturalidade (Estado)'), ('label', 'UF Naturalidade')]),
            ('Titular', 'Dados Gerais', 'Nacionalidade (País)', [('label', 'Nacionalidade (País)'), ('label', 'Nacionalidade')]),
            ('Titular', 'Dados Gerais', 'Data de Nascimento', [('label', 'Data de Nascimento')]),
            ('Titular', 'Dados Gerais', 'Raça', [('label', 'Raça')]),
            ('Titular', 'Dados Gerais', 'Gênero', [('label', 'Gênero')]),
            ('Titular', 'Dados Gerais', 'Contrato Distratado ou Rescindido Involutariamente', [('label', 'Contrato Distratado')]),
            ('Titular', 'Dados Gerais', 'Recebe Atendimento Socio-assistencial do Município', [('label', 'Recebe Atendimento Socio-assistencial')]),
            ('Titular', 'Dados Gerais', 'Situação Atual do Domicílio', [('label', 'Situação Atual do Domicílio')]),
            ('Titular', 'Dados Gerais', 'Tipo de Beneficiário Programa Social', [('label', 'Tipo de Beneficiário Programa Social')]),
            ('Titular', 'Dados Gerais', 'Proprietário do Imóvel?', [('label', 'Proprietário do Imóvel')]),
            ('Titular', 'Dados Gerais', 'Recebeu/Participou de REURB?', [('label', 'REURB')]),

            # CadÚnico
            ('Titular', 'Informações do CadÚnico', 'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?', [('label', 'violência doméstica')]),
            ('Titular', 'Informações do CadÚnico', 'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?', [('label', 'pessoa(s) negra(s)')]),
            ('Titular', 'Informações do CadÚnico', 'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?', [('label', 'pessoa(s) LGBT')]),
            ('Titular', 'Informações do CadÚnico', 'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?', [('label', 'Quilombola')]),
            ('Titular', 'Informações do CadÚnico', 'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?', [('label', 'povos tradicionais')]),
            ('Titular', 'Informações do CadÚnico', 'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?', [('label', 'situação de rua')]),
            ('Titular', 'Informações do CadÚnico', 'Programas CAD Único', [('label', 'Programas CAD Único')]),
            ('Titular', 'Informações do CadÚnico', 'Tratamento Oncológico/Hemodialise', [('label', 'Tratamento Oncológico')]),

            # Cadastro Preferencial
            ('Titular', 'Cadastro Preferencial', 'Cadastro Preferencial', [('label', 'Cadastro Preferencial')]),
            ('Titular', 'Cadastro Preferencial', 'Deficiência', [('label', 'Deficiência')]),
            ('Titular', 'Cadastro Preferencial', 'Doença', [('label', 'Doença')]),
            ('Titular', 'Cadastro Preferencial', 'Possui Microcefalia', [('label', 'Possui Microcefalia')]),
            ('Titular', 'Cadastro Preferencial', 'Tem Veículo Próprio', [('label', 'Tem Veículo Próprio')]),

            # Educação e Profissão
            ('Titular', 'Educação e Profissão', 'Escolaridade', [('label', 'Escolaridade')]),
            ('Titular', 'Educação e Profissão', 'Escola', [('label', 'Escola')]),
            ('Titular', 'Educação e Profissão', 'Situação de Emprego', [('label', 'Situação de Emprego')]),
            ('Titular', 'Educação e Profissão', 'Profissão', [('label', 'Profissão')]),
            ('Titular', 'Educação e Profissão', 'Profissão (se Outro)', [('label', 'Profissão (se Outro)')]),

            # Situação Marital
            ('Titular', 'Situação Marital', 'Estado Civil', [('label', 'Estado Civil')]),
            ('Titular', 'Situação Marital', 'Data do Casamento/União', [('label', 'Data do Casamento')]),
            ('Titular', 'Situação Marital', 'Regime do Casamento', [('label', 'Regime do Casamento')]),
            ('Titular', 'Situação Marital', 'Regime do Casamento (se Outro)', [('label', 'Regime do Casamento (se Outro)')]),

            # CAD (Título de Eleitor)
            ('Titular', 'CAD', 'Data de Emissão', [('id', 'dtEmissaoTituloEleitor'), ('label', 'Data de Emissão')]),
            ('Titular', 'CAD', 'UF Título de Eleitor', [('label', 'UF Título de Eleitor')]),
            ('Titular', 'CAD', 'Cidade do Título de Eleitor', [('label', 'Cidade do Título de Eleitor')]),

            # Dados de Contato
            ('Titular', 'Dados de Contato', 'Email', [('label', 'Email')]),
            ('Titular', 'Dados de Contato', 'Telefone', [('label', 'Telefone')]),
            ('Titular', 'Dados de Contato', 'Telefone Celular', [('label', 'Telefone Celular')]),
            ('Titular', 'Dados de Contato', 'Telefone Comercial', [('label', 'Telefone Comercial')]),
            ('Titular', 'Dados de Contato', 'Telefone Para Recados', [('label', 'Telefone Para Recados')]),

            # Anotações
            ('Titular', 'Anotações', 'Parecer Social', [('label', 'Parecer Social')]), # <--- NOVO
            ('Titular', 'Anotações', 'Diagnóstico Social', [('label', 'Diagnóstico Social')]),
            ('Titular', 'Anotações', 'Observações', [('label', 'Observações')]),
        ]

        for aba, secao, campo, tentativas in passos:
            _, valor = self.ler_campo_ordenado(aba=aba, secao=secao, campo=campo, tentativas=tentativas)
            mapa[_col_key(aba, secao, campo)] = valor

        return mapa

    # --- ABAS COM MODAL ---
    def extrair_segundo_titular(self):
        if not self.clicar_aba('Segundo Titular'):
            return {}

        aba = 'Segundo Titular'
        especificacao: list[tuple[str, str, list[tuple[str, str]]]] = []

        # Dados Gerais
        gerais = [
            'Nome', 'CPF/CNPJ', 'Naturalidade (Cidade)', 'UF Naturalidade (Estado)', 
            'Nacionalidade (País)', 'Data de Nascimento', 'Raça', 'Gênero',
            'Contrato Distratado ou Rescindido Involutariamente', 
            'Recebe Atendimento Socio-assistencial do Município', 
            'Situação Atual do Domicílio', 'Tipo de Beneficiário Programa Social', 
            'Proprietário do Imóvel?', 'Recebeu/Participou de REURB?'
        ]
        for c in gerais: especificacao.append(('Dados Gerais', c, [('label', c)]))

        # CadÚnico
        cadunico = [
            'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
            'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
            'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
            'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
            'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
            'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
            'Programas CAD Único', 'Tratamento Oncológico/Hemodialise'
        ]
        for c in cadunico: especificacao.append(('Informações do CadÚnico', c, [('label', c)]))

        # Cadastro Preferencial
        for c in ['Cadastro Preferencial', 'Deficiência', 'Doença', 'Possui Microcefalia', 'Tem Veículo Próprio']:
            especificacao.append(('Cadastro Preferencial', c, [('label', c)]))

        # Educação e Profissão
        for c in ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)']:
            especificacao.append(('Educação e Profissão', c, [('label', c)]))

        # Situação Marital
        for c in ['Estado Civil', 'Data do Casamento/União', 'Regime do Casamento', 'Regime do Casamento (se Outro)']:
            especificacao.append(('Situação Marital', c, [('label', c)]))

        # CAD
        for c in ['Data de Emissão', 'UF Título de Eleitor', 'Cidade do Título de Eleitor']:
            especificacao.append(('CAD', c, [('label', c)]))

        # Contato
        for c in ['Email', 'Telefone', 'Telefone Celular', 'Telefone Comercial', 'Telefone Para Recados']:
            especificacao.append(('Dados de Contato', c, [('label', c)]))

        # Parecer
        especificacao.append(('Parecer Social', 'Parecer Social', [('label', 'Parecer Social')]))
        especificacao.append(('Parecer Social', 'Diagnóstico Social', [('label', 'Diagnóstico Social')]))
        especificacao.append(('Parecer Social', 'Observações', [('label', 'Observações')]))

        variantes = {}
        col_keys = []
        for secao, campo, tentativas in especificacao:
            k = _col_key(aba, secao, campo)
            col_keys.append(k)
            labels = [txt for (modo, txt) in (tentativas or []) if modo == 'label']
            variantes[k] = labels or [campo]

        # Usa a função que clica no botão EDITAR
        itens = self._extrair_tabela_com_modal_edicao_itens(variantes)
        return self._itens_para_colunas_multilinha(itens, col_keys)

    def extrair_composicao_familiar(self):
        if not self.clicar_aba('Composição Familiar'):
            return {}

        aba = 'Composição Familiar'
        especificacao: list[tuple[str, str, list[str]]] = []

        # Dados Gerais
        campos_gerais = [
            'Tipo', 'Nome', 'CPF', 'Naturalidade', 'UF Naturalidade', 'Nacionalidade',
            'Data de Nascimento', 'Raça', 'Gênero', 'Telefone', 'Estado Civil',
            'Pessoa Negra?', 'Pessoa Quilombola?', 'Pessoa LGBTQIAPN+?', 'Situação de Rua?'
        ]
        for c in campos_gerais: especificacao.append(('Dados Gerais', c, [c]))

        # Educação e Profissão
        campos_educ = ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço']
        for c in campos_educ: especificacao.append(('Educação e Profissão', c, [c]))

        # Cadastro Preferencial
        for c in ['Cadastro Preferencial', 'Deficiência', 'Doença']:
            especificacao.append(('Cadastro Preferencial', c, [c]))

        variantes = {}
        col_keys = []
        for secao, campo, labels in especificacao:
            k = _col_key(aba, secao, campo)
            col_keys.append(k)
            variantes[k] = labels or [campo]

        itens = self._extrair_tabela_com_modal_edicao_itens(variantes)
        return self._itens_para_colunas_multilinha(itens, col_keys)
    
    def extrair_renda(self):
        if not self.clicar_aba('Renda'):
            return {}

        aba = 'Renda'
        secao = 'Dados da Renda'
        
        pessoas = []
        tipos = []
        valores = []

        try:
            # Pequena pausa técnica apenas para a tabela desenhar na tela
            time.sleep(1)
            
            tabelas = self.driver.find_elements(By.TAG_NAME, "table")
            
            for tabela in tabelas:
                if not tabela.is_displayed(): continue
                
                linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")
                if not linhas: # Se não achou tbody, tenta tr direto
                     linhas = tabela.find_elements(By.TAG_NAME, "tr")

                for linha in linhas:
                    cols = linha.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 3:
                        nome = self._texto_limpo(cols[0].text)
                        
                        # Ignora se for cabeçalho
                        if not nome or nome.lower() == 'nome':
                            continue
                            
                        pessoas.append(nome)
                        # Garante que pega as outras colunas mesmo se estiverem vazias
                        tipos.append(self._texto_limpo(cols[1].text) if len(cols) > 1 else "")
                        valores.append(self._texto_limpo(cols[2].text) if len(cols) > 2 else "")

                # Se achou pelo menos uma pessoa, considera que é a tabela certa e para
                if pessoas:
                    break
                    
        except Exception:
            pass

        # Formatação SIMPLES (Só pula linha, sem enfeites)
        def formatar(lista):
            if not lista: return "" 
            return "\n".join(lista)

        row = {}
        row[_col_key(aba, secao, 'Pessoa')] = formatar(pessoas)
        row[_col_key(aba, secao, 'Tipo de Renda')] = formatar(tipos)
        row[_col_key(aba, secao, 'Valor da Renda')] = formatar(valores)
        row[_col_key(aba, secao, 'Tipo se Apto Para Reurb')] = "" 

        return row

    # --- DEMAIS ABAS ---
    def extrair_endereco(self):
        if not self.clicar_aba('Endereço'):
            return {}

        aba = 'Endereço'
        secao = 'Dados do Endereço'
        dados: dict[str, str] = {}

        campos_diretos = [
            'Lote',
            'Quadra',
            'CEP',
            'UF',
            'Cidade',
            'Bairro',
            'Endereço',
            'Número',
            'Complemento',
            'Tipo de Moradia',
            'Possui outro imóvel?',
        ]
        for c in campos_diretos:
            dados[_col_key(aba, secao, c)] = self._extrair_por_label_simples(c)
        
        dados[_col_key(aba, secao, 'Forma de aquisição')] = self._extrair_por_label_simples('Forma de aquisição')
        dados[_col_key(aba, secao, 'Relação com o imóvel')] = self._extrair_por_label_simples('Relação com o imóvel')
        dados[_col_key(aba, secao, 'Uso do imóvel')] = self._extrair_por_label_simples('Uso do imóvel')
        dados[_col_key(aba, secao, 'Há famílias que habitam ou trabalham a, no máximo, 28 km de distância do Centro do Empreendimento?')] = (
            self._extrair_por_label_simples('28 km de distância')
            or self._extrair_por_label_simples('Centro do Empreendimento')
        )
        dados[_col_key(aba, secao, 'Reside no município desde (ano)')] = (
            self._extrair_por_label_simples('Reside no município desde (ano)')
            or self._extrair_por_label_simples('Reside no município desde')
        )
        return dados

    def extrair_questionario_mapa(self) -> dict[str, str]:
        """Extrai o questionário procurando, por cada div.question, o input marcado e seu label local.

    Estratégia determinística:
      - clica na aba Questionário
      - espera curto (render)
      - executa um script no browser que itera cada div.question e:
          * pega heading (h6/h5/h4/label/text)
          * procura input[type=radio|checkbox]:checked DENTRO do mesmo container
          * se encontrado, tenta achar label[for=id] DENTRO do container (fallback: closest label, sibling, global)
          * retorna lista de {heading, chosen: {id, value, label, method}, rawLabels}
      - normaliza para 'Sim'/'Não' e preenche também as chaves do schema COL_SPECS
    """
    try:
        if not self.clicar_aba('Questionário'):
            return {}
        # pequena espera para a aba renderizar
        time.sleep(0.6)

        js = r"""
        const out = [];
        const questions = Array.from(document.querySelectorAll('div.question, .question'));
        function textOf(el){ try { return (el && (el.innerText||el.textContent||'')).toString().trim(); } catch(e){ return ''; } }
        for (const q of questions){
            try {
                const h = q.querySelector('h6,h5,h4,label');
                const heading = textOf(h) || textOf(q).split('\n')[0] || '';
                let chosen = null;
                // primeiro: input marcado dentro do container
                let checked = q.querySelector("input[type='radio']:checked, input[type='checkbox']:checked");
                if (checked) {
                    const id = checked.id || '';
                    let label = null;
                    if (id) label = q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']");
                    if (!label) label = checked.closest('label');
                    if (!label) {
                        let sib = checked.nextElementSibling;
                        if (sib && (sib.tagName||'').toLowerCase()==='label') label = sib;
                    }
                    chosen = { id: id, value: checked.value||'', label: label ? textOf(label) : '', labelHTML: label ? label.outerHTML : '', method: 'inside_checked' };
                } else {
                    // fallback: procura inputs dentro do container cuja propriedade .checked seja true
                    const inputs = Array.from(q.querySelectorAll("input[type='radio'], input[type='checkbox']"));
                    for (const inp of inputs){
                        try {
                            if (inp.checked) {
                                const id = inp.id || '';
                                let label = id ? (q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']")) : null;
                                if (!label) label = inp.closest('label');
                                if (!label) { let s = inp.nextElementSibling; if (s && (s.tagName||'').toLowerCase()==='label') label = s; }
                                chosen = { id: id, value: inp.value||'', label: label ? textOf(label) : '', labelHTML: label ? label.outerHTML : '', method: 'inside_property_checked' };
                                break;
                            }
                        } catch(e){}
                    }
                    // se ainda nada, tentar input[name]:checked global com label pertencente ao mesmo question
                    if (!chosen) {
                        const names = [...new Set(Array.from(q.querySelectorAll("input")).map(i => i.name).filter(n=>n))];
                        for (const nm of names){
                            try {
                                const g = document.querySelector("input[name='"+nm+"']:checked");
                                if (g) {
                                    const id = g.id || '';
                                    let label = id ? (q.querySelector("label[for='"+id+"']") || document.querySelector("label[for='"+id+"']")) : null;
                                    const ok = label && (q.contains(label) || (label.closest('.question') && label.closest('.question') === q));
                                    if (ok) {
                                        chosen = { id:id, value:g.value||'', label: textOf(label), labelHTML: label.outerHTML, method: 'global_checked_but_local_label' };
                                        break;
                                    } else if (id) {
                                        // fallback: attach global label if exists
                                        const labg = document.querySelector("label[for='"+id+"']");
                                        if (labg) {
                                            chosen = { id:id, value:g.value||'', label: textOf(labg), labelHTML: labg.outerHTML, method: 'global_checked_fallback' };
                                            break;
                                        }
                                    }
                                }
                            } catch(e){}
                        }
                    }
                }

                // coletar labels brutos (útil para debug/compat)
                const rawLabels = Array.from(q.querySelectorAll('label')).map(l => ({ text: textOf(l), html: (l.outerHTML||'').slice(0,400) }));
                out.push({ heading: heading, chosen: chosen, rawLabels: rawLabels });
            } catch(e){}
        }
        return out;
        """

        collected = []
        try:
            collected = self.driver.execute_script(js) or []
        except Exception:
            collected = []

        saida: dict[str, str] = {}
        respostas_agregadas = []
        for item in (collected or []):
            try:
                heading = (item.get('heading') or '')[:1000]
                chosen = item.get('chosen') or {}
                resp_text = ''
                metodo = ''
                if chosen and chosen.get('label'):
                    resp_text = chosen.get('label')
                    metodo = chosen.get('method') or 'chosen_label'
                elif chosen and chosen.get('value'):
                    resp_text = chosen.get('value')
                    metodo = chosen.get('method') or 'chosen_value'
                else:
                    resp_text = ''
                    metodo = 'not_detected'

                rnorm = ''
                if resp_text:
                    rl = resp_text.strip().lower()
                    if rl in ('s', 'sim'):
                        rnorm = 'Sim'
                    elif rl in ('n', 'não', 'nao'):
                        rnorm = 'Não'
                    else:
                        rnorm = resp_text.strip().capitalize()

                key_norm = _norm_texto_chave(heading)
                if key_norm:
                    saida[key_norm] = rnorm
                    respostas_agregadas.append(f"{heading} -> {rnorm} ({metodo})")
                    # preenche também a chave do schema correspondente
                    for p_schema in COL_SPECS:
                        try:
                            if p_schema.aba == 'Questionário' and p_schema.campo == heading:
                                saida[p_schema.key] = rnorm
                                break
                        except:
                            continue

        if respostas_agregadas:
            saida['Respostas do Questionário'] = "\n".join(respostas_agregadas)

        return saida

    except Exception:
        return {}

    def extrair_questionario(self) -> str:
        mp = self.extrair_questionario_mapa() or {}
        linhas = []
        for k, v in mp.items():
            linhas.append(f"{k} -> {v}")
        return "\n".join(linhas)

class HabibotBot:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.excel = None
        self.caminho_excel_final = None
        self.caminho_excel_visualizacao = None
        self.caminho_log_ordem = None
        self.caminho_log_acoes = None
        self.ordem_logger = None
        self.acoes_logger = None
        self.pasta_saida = None
        self.pasta_debug = None
        self.cooldown_segundos = 0
        self.debug = False
        self.debug_visual = False
        self._headless = False
        self.base_url = None

    def solicitar_base_url(self):
        print("🌐 Insira o link do sistema de Habitação")
        print("Exemplo: app.habibot.com.br")
        while True:
            url = input("\n👉 URL do sistema: ").strip()
            if not url:
                print("Por favor, insira uma URL válida.\n")
                continue
            if not url.startswith("http"):
                url = "https://" + url
            url = url.rstrip("/")
            self.base_url = url
            self.nome_sistema = _extrair_nome_sistema(self.base_url)
            break

    @property
    def login_url(self):
        return f"{self.base_url}/login" if self.base_url else None

    @property
    def candidatos_url(self):
        return f"{self.base_url}/candidatos" if self.base_url else None

    def _highlight(self, el, *, rotulo: str | None = None):
        """Highlight visual para telas fora do Extrator (ex.: login / busca candidatos)."""
        if not self.debug_visual or self._headless:
            return
        try:
            if not el:
                return
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
            js = """
                const el = arguments[0];
                const label = arguments[1] || '';
                const prevOutline = el.style.outline;
                const prevOutlineOffset = el.style.outlineOffset;
                el.style.outline = '3px solid #ff00ff';
                el.style.outlineOffset = '2px';
                if (label) {
                    el.setAttribute('data-habibot-debug', label);
                }
                setTimeout(() => {
                    try {
                        el.style.outline = prevOutline;
                        el.style.outlineOffset = prevOutlineOffset;
                    } catch (e) {}
                }, 350);
            """
            self.driver.execute_script(js, el, rotulo or "")
        except:
            return

    def _log_acao(self, acao: str, *, detalhe: str | None = None, ok: bool | None = None, dur_ms: int | None = None, obs: str | None = None):
        if not self.acoes_logger:
            return
        try:
            url = ''
            try:
                url = self.driver.current_url
            except:
                url = ''
            self.acoes_logger.log(
                area='BOT',
                acao=str(acao),
                detalhe=detalhe,
                url=url,
                ok=ok,
                dur_ms=dur_ms,
                obs=obs,
            )
        except:
            pass

    def _obter_linha_por_cpf_posicao(self, cpf_informado: str, posicao: int):
        try:
            qtd, indices = self.buscar_candidato_por_cpf(cpf_informado)
            if not indices or posicao >= len(indices):
                return None
            linhas = self.obter_linhas()
            idx = indices[posicao]
            if idx < len(linhas):
                return linhas[idx]
        except:
            pass
        return None

    @staticmethod
    def _cpf_digitos(valor: str) -> str:
        return ''.join(ch for ch in (valor or '') if ch.isdigit())

    @staticmethod
    def _cpf_mascarado(digitos11: str) -> str:
        d = digitos11
        return f"{d[0:3]}.{d[3:6]}.{d[6:9]}-{d[9:11]}"

    @staticmethod
    def normalizar_cpfs(valores: list[str]) -> list[str]:
        tokens: list[str] = []
        for c in valores:
            for parte in (c or '').replace(';', ' ').replace(',', ' ').split():
                p = (parte or '').strip()
                if p:
                    tokens.append(p)
        saida: list[str] = []
        for t in tokens:
            dig = HabibotBot._cpf_digitos(t)
            if t.isdigit() and len(dig) == 11:
                saida.append(HabibotBot._cpf_mascarado(dig))
            else:
                saida.append(t)
        return saida

    @staticmethod
    def solicitar_escopo_extracao() -> tuple[str, list[str]]:
        print("📌 O que você deseja extrair?")
        print("1) Extrair dados de TODOS os candidatos cadastrados")
        print("2) Extrair dados somente de CPFs informados")
        escolha = input("\nEscolha uma opção [Padrão = 1]: ").strip() or "1"
        if escolha != "2":
            print("\nSelecionado: TODOS os candidatos")
            print("\n──────────────────────────────────────────────────────────────────────────────────\n")
            return "todos", []
        print("\nSelecionado: SOMENTE CPFs informados")
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        print("📄 Cole abaixo a lista de CPFs. Quando terminar, pressione ENTER em uma linha vazia.\n")
        cpfs: list[str] = []
        while True:
            try:
                linha = input()
            except EOFError:
                break
            if linha is None:
                break
            linha = linha.strip()
            if not linha:
                break
            cpfs.append(linha)
        normalizados = HabibotBot.normalizar_cpfs(cpfs)
        if not normalizados:
            raise ValueError("Nenhum CPF foi informado.")
        print(f"Total de CPFs informados: {len(normalizados)}")
        return "cpfs", normalizados

    @staticmethod
    def solicitar_cooldown_segundos() -> int:
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        print("⏳  Cooldown entre extrações")
        print("Informe o tempo de espera (em segundos) entre a extração dos candidatos")
        
        while True:
            raw = input("\n👉 Cooldown (segundos) [Padrão = 0]: ").strip()
            if raw == "":
                print("\n──────────────────────────────────────────────────────────────────────────────────\n")
                return 0
            try:
                valor = int(float(raw.replace(',', '.')))
                if valor < 0:
                    print("Valor inválido: cooldown não pode ser negativo.")
                    continue
                print("")
                return valor
            except ValueError:
                print("Valor inválido. Informe um número.")

    @staticmethod
    def solicitar_modo_execucao() -> bool:
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        print("🖥️  Modo de execução:")
        print("1) Visualizar bot trabalhando (abrir o navegador)")
        print("2) Segundo plano (sem janela do navegador)")
        
        escolha = input("\n👉 Selecione (1/2) [Padrão: 1]: ").strip() or "1"
        if escolha == "2":
            print("\n🥷🏻 Modo selecionado: segundo plano (headless)")
            print("\n──────────────────────────────────────────────────────────────────────────────────\n")
            return True
        print("\n 👀 Modo selecionado: visualização no navegador")
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        return False

    def iniciar(self, headless: bool = False):
        self._headless = bool(headless)
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
            try:
                self.driver = None
            except:
                pass

    def solicitar_credenciais(self):
        nome_sistema = getattr(self, 'nome_sistema', None) or getattr(self, 'base_url', 'sistema')
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        print(f"🔐 Credenciais do {nome_sistema}")
        print("Para continuar, informe seu login e senha:")

        def _email_parece_valido(s: str) -> bool:
            s = (s or '').strip()
            if not s:
                return False
            if s.startswith('&') or s.lower().endswith('.ps1') or 'activate.ps1' in s.lower():
                return False
            if '@' not in s:
                return False
            # validação leve: precisa ter algo após o '@' com ponto
            local, _, dom = s.partition('@')
            if not local or not dom or '.' not in dom:
                return False
            return True

        while True:
            usuario = input("\n  📨 Usuário (e-mail): ").strip()
            if _email_parece_valido(usuario):
                break
            print("❌ Usuário inválido. Tente novamente\n")

        try:
            senha = pwinput.pwinput(prompt="  👉 Senha: ", mask="*")
        except Exception:
            senha = input("\n  🔑 Senha (visível): ")
        print("")
        if not senha:
            raise ValueError('Senha é obrigatória')
        print("\n──────────────────────────────────────────────────────────────────────────────────\n")
        return usuario, senha

    def _ler_total_resultados_lista(self) -> int | None:
        try:
            candidatos = self.driver.find_elements(
                By.XPATH,
                "//p[contains(@class,'text-muted') and contains(., 'Mostrando') and contains(., 'resultados')]",
            )
            el = None
            for c in candidatos:
                try:
                    if c.is_displayed():
                        el = c
                        break
                except:
                    continue
            if not el: return None
            try:
                txt = (el.text or '').strip()
                txt = ' '.join(txt.split())
                m = re.search(r"\bde\s+(\d+)\s+resultados\b", txt, flags=re.IGNORECASE)
                if m: return int(m.group(1))
            except:
                pass
            bs = el.find_elements(By.TAG_NAME, 'b')
            if len(bs) < 3: return None
            total_txt = (bs[2].text or '').strip()
            total_digits = ''.join(ch for ch in total_txt if ch.isdigit())
            if total_digits == '': return None
            return int(total_digits)
        except:
            return None

    def _aguardar_total_resultados_estavel(self, timeout_segundos: float = 8.0) -> int | None:
        fim = time.monotonic() + max(0.5, timeout_segundos)
        ultimo_total = None
        repeticoes = 0
        pausa = 0.06
        while time.monotonic() < fim:
            try:
                if self.driver.find_elements(By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active"):
                    time.sleep(pausa)
                    continue
            except:
                pass
            total = self._ler_total_resultados_lista()
            if total is None:
                time.sleep(pausa)
                continue
            if total == 1:
                try:
                    linhas = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                    if not any(l.is_displayed() for l in linhas):
                        time.sleep(pausa)
                        continue
                except:
                    time.sleep(pausa)
                    continue
            if total == ultimo_total:
                repeticoes += 1
            else:
                repeticoes = 0
                ultimo_total = total
            if repeticoes >= 1:
                return total
            time.sleep(pausa)
        return ultimo_total

    def _assinatura_primeira_linha_lista(self) -> str:
        try:
            el = self.driver.find_element(By.CSS_SELECTOR, "tbody tr")
            return (el.text or '').strip()
        except:
            return ''

    def _aguardar_lista_atualizar(self, total_antes: int | None = None, assinatura_antes: str | None = None, timeout_segundos: float = 10.0):
        overlay_sel = ".vld-background, .loading, .spinner, .vld-overlay.is-active"
        assinatura_antes = assinatura_antes if assinatura_antes is not None else ''
        try:
            WebDriverWait(self.driver, max(1.0, timeout_segundos), poll_frequency=0.1).until(
                lambda d: not d.find_elements(By.CSS_SELECTOR, overlay_sel)
            )
        except:
            pass
        
        def pronto(d):
            try:
                if d.find_elements(By.CSS_SELECTOR, overlay_sel): return False
            except:
                return False
            total = self._ler_total_resultados_lista()
            if total == 0: return True
            if total_antes is not None and total is not None and total != total_antes: return True
            assinatura = self._assinatura_primeira_linha_lista()
            if assinatura_antes and assinatura and assinatura != assinatura_antes: return True
            if total_antes is None and not assinatura_antes:
                try:
                    linhas = d.find_elements(By.CSS_SELECTOR, "tbody tr")
                    if any(l.is_displayed() for l in linhas): return True
                except:
                    pass
            return False
        
        try:
            WebDriverWait(self.driver, max(1.0, timeout_segundos), poll_frequency=0.1).until(pronto)
        except:
            pass

    def _disparar_busca_por_texto(self, campo, texto_busca: str):
        texto = (texto_busca or '').strip()
        self._highlight(campo, rotulo=f"BUSCA: {texto[:40]}")
        try: campo.click()
        except: pass
        try: campo.clear()
        except: pass
        try: campo.send_keys(texto)
        except: pass
        try: 
            if Keys: campo.send_keys(Keys.ENTER)
        except: pass

    def buscar_candidato_por_cpf(self, cpf_informado: str) -> tuple[int, list[int]]:
        t0 = time.monotonic()
        dig = HabibotBot._cpf_digitos(cpf_informado)
        if (cpf_informado or '').isdigit() and len(dig) == 11:
            texto_busca = HabibotBot._cpf_mascarado(dig)
        else:
            texto_busca = (cpf_informado or '').strip()

        try:
            campo = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#searchMemberList")))
            self._highlight(campo, rotulo="CAMPO BUSCA (#searchMemberList)")
            total_antes = self._ler_total_resultados_lista()
            assinatura_antes = self._assinatura_primeira_linha_lista()

            self._disparar_busca_por_texto(campo, texto_busca)
            self._aguardar_lista_atualizar(total_antes=total_antes, assinatura_antes=assinatura_antes, timeout_segundos=10.0)

            total_resumo = self._aguardar_total_resultados_estavel(timeout_segundos=8.0)
            if total_resumo == 0:
                self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs='0 resultados')
                return 0, []

            linhas = self.obter_linhas()
            if not linhas and total_resumo == 1:
                time.sleep(2)
                linhas = self.obter_linhas()
            
            if not linhas:
                self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs='Sem linhas visíveis')
                return 0, []

            alvo = dig if len(dig) == 11 else HabibotBot._cpf_digitos(texto_busca)
            indices = []
            if alvo:
                for idx, l in enumerate(linhas):
                    if alvo in HabibotBot._cpf_digitos(l.text):
                        indices.append(idx)
            else:
                indices = list(range(len(linhas)))

            qtd = len(indices)
            if total_resumo is not None:
                if total_resumo == 1 and qtd >= 1:
                    self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs=f"total_resumo={total_resumo}; matches={qtd}")
                    return 1, indices
                if total_resumo == 1 and qtd == 0:
                    self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs=f"total_resumo={total_resumo}; matches=0")
                    return 1, []
                self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs=f"total_resumo={total_resumo}; matches={qtd}")
                return total_resumo, indices

            self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=True, dur_ms=int((time.monotonic() - t0) * 1000), obs=f"matches={qtd}")
            return qtd, indices
        except Exception:
             self._log_acao('BUSCAR_CANDIDATO', detalhe=f"q={texto_busca}", ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs='Exceção')
             return 0, []

    def login(self, usuario: str, senha: str):
        t0 = time.monotonic()
        self.driver.get(self.login_url)
        try:
            campo_usuario = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
            self._highlight(campo_usuario, rotulo="LOGIN: usuário")
            campo_usuario.clear()
            campo_usuario.send_keys(usuario)
            self._log_acao('LOGIN_PREENCHER_USUARIO', ok=True)
            campo_senha = self.driver.find_element(By.CSS_SELECTOR, "input[type='password']")
            self._highlight(campo_senha, rotulo="LOGIN: senha")
            campo_senha.clear()
            campo_senha.send_keys(senha)
            self._log_acao('LOGIN_PREENCHER_SENHA', ok=True)
            botao = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit'], button.btn-primary")
            self._highlight(botao, rotulo="LOGIN: entrar")
            botao.click()
            self._log_acao('LOGIN_SUBMIT', ok=True)
            self.wait.until(EC.url_changes(self.login_url))
            self._log_acao('LOGIN_OK', ok=True, dur_ms=int((time.monotonic() - t0) * 1000))
        except Exception as e:
            self._log_acao('LOGIN_FALHA', ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs=str(e))
            raise RuntimeError(f"Erro no login: {e}")

    def navegar_candidatos(self):
        t0 = time.monotonic()
        self.driver.get(self.candidatos_url)
        self.wait.until(lambda d: d.execute_script("return document.readyState === 'complete'"))
        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except: pass
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr, table tbody tr, .table tbody tr")))
        self._log_acao('NAVEGAR_CANDIDATOS', ok=True, dur_ms=int((time.monotonic() - t0) * 1000))

    def garantir_lista_candidatos(self):
        try:
            if '/candidatos' not in (self.driver.current_url or ''):
                self.navegar_candidatos()
                return
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#searchMemberList")))
            try:
                self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
            except: pass
        except:
            self.navegar_candidatos()

    def obter_linhas(self):
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")))
        except: pass
        linhas = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        return [l for l in linhas if l.is_displayed()]

    def _pagina_ativa(self) -> str:
        try:
            el = self.driver.find_element(By.XPATH, "//ul[contains(@class,'pagination')]//li[contains(@class,'active')]//a | //ul[contains(@class,'pagination')]//li[contains(@class,'active')]//button")
            return (el.text or '').strip()
        except:
            return ''

    def proxima_pagina(self) -> bool:
        t0 = time.monotonic()
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
            except: continue
        if not botao:
            self._log_acao('PROXIMA_PAGINA', ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs='Botão não encontrado')
            return False

        assinatura_antes = ''
        try:
            assinatura_antes = (self.driver.find_element(By.CSS_SELECTOR, "tbody tr").text or '').strip()
        except: pass

        try: self.driver.execute_script("arguments[0].click();", botao)
        except: botao.click()

        try: self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except: pass

        def mudou(d):
            try:
                depois = self._pagina_ativa()
                if antes and depois and depois != antes: return True
            except: pass
            try:
                assinatura_depois = (d.find_element(By.CSS_SELECTOR, "tbody tr").text or '').strip()
                if assinatura_antes and assinatura_depois and assinatura_depois != assinatura_antes: return True
            except: pass
            return False

        try: WebDriverWait(self.driver, 12, poll_frequency=0.15).until(mudou)
        except: pass
        self._log_acao('PROXIMA_PAGINA', ok=True, dur_ms=int((time.monotonic() - t0) * 1000), detalhe=f"antes={antes} depois={self._pagina_ativa()}")
        return True

    def abrir_candidato(self, linha):
        t0 = time.monotonic()
        candidatos_css = [
            "button.btn-soft-info",
            "button[title*='Editar'], a[title*='Editar']",
            "button[aria-label*='Editar'], a[aria-label*='Editar']",
        ]
        btn = None
        for sel in candidatos_css:
            try:
                btn = linha.find_element(By.CSS_SELECTOR, sel)
                if btn and btn.is_displayed() and btn.is_enabled(): break
            except: btn = None
        
        if not btn:
            candidatos_xpath = [
                ".//*[self::button or self::a][contains(translate(@title,'EDITAR','editar'),'editar') or contains(translate(@aria-label,'EDITAR','editar'),'editar') or contains(translate(normalize-space(.),'EDITAR','editar'),'editar')]",
                ".//i[contains(@class,'pencil') or contains(@class,'edit') or contains(@class,'ri-pencil')]/ancestor::*[self::button or self::a][1]",
            ]
            for xp in candidatos_xpath:
                try:
                    btn = linha.find_element(By.XPATH, xp)
                    if btn and btn.is_displayed() and btn.is_enabled(): break
                except: btn = None

        if not btn:
            self._log_acao('ABRIR_CANDIDATO', ok=False, dur_ms=int((time.monotonic() - t0) * 1000), obs='Botão editar/abrir não encontrado')
            raise RuntimeError('Botão de editar/abrir candidato não encontrado na linha.')

        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        try: WebDriverWait(self.driver, 3, poll_frequency=0.1).until(lambda d: btn.is_displayed() and btn.is_enabled())
        except: pass
        try: btn.click()
        except: self.driver.execute_script("arguments[0].click();", btn)

        self.wait.until(EC.url_contains('/dadoscandidato'))
        try: self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-overlay.is-active")))
        except: pass
        try: self.wait.until(lambda d: d.execute_script("return document.readyState === 'complete'"))
        except: pass
        self._log_acao('ABRIR_CANDIDATO', ok=True, dur_ms=int((time.monotonic() - t0) * 1000))

    def voltar_lista(self):
        t0 = time.monotonic()
        # Só tenta voltar se driver e wait estão ativos e driver não está fechado
        if not self.driver or not self.wait:
            return
        try:
            # Verifica se o driver está ativo
            if hasattr(self.driver, 'session_id') and self.driver.session_id is None:
                return
            self.driver.back()
            self.wait.until(lambda d: '/candidatos' in d.current_url)
        except Exception:
            try:
                self.navegar_candidatos()
            except Exception:
                return
        try:
            self.wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, ".vld-background, .loading, .spinner, .vld-overlay.is-active")))
        except: pass
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#searchMemberList")))
        except: pass
        try:
            self._log_acao('VOLTAR_LISTA', ok=True, dur_ms=int((time.monotonic() - t0) * 1000))
        except: pass

    def _col(self, cat, sub, field):
        return _col_key(cat, sub, field)

    def extrair_um(self) -> dict:
        """Função Gerente: Só chama as abas que foram pedidas."""
        extrator = ExtratorHabibot(
            self.driver,
            self.wait,
            debug=bool(self.debug),
            debug_visual=bool(self.debug_visual) and (not self._headless),
            ordem_logger=self.ordem_logger,
            acoes_logger=self.acoes_logger,
        )
        row = {}

        # 1) Titular (Loteamento, Dados Pessoais, Contato, CadÚnico, etc.)
        # Já configuramos essa função para pegar só o essencial
        titular = extrator.extrair_titular_mapa() or {}
        row.update(titular)

        # 2) Segundo Titular (Clica na aba, abre modal, pega dados)
        st = extrator.extrair_segundo_titular() or {}
        row.update(st)

        # 3) Composição Familiar (Clica na aba, abre modal, pega dados)
        cf = extrator.extrair_composicao_familiar() or {}
        row.update(cf)

        # 4) Renda (Clica na aba, abre modal)
        renda = extrator.extrair_renda() or {}
        row.update(renda)

        # 5) Endereço (Clica na aba)
        end = extrator.extrair_endereco() or {}
        row.update(end)

        # 6) Questionário (Clica na aba, pega perguntas e respostas)
        # Nota: Agora pega do jeito certo (bolinha marcada)
        qmap = extrator.extrair_questionario_mapa() or {}
        
        # Joga para o Excel na coluna única 'Respostas do Questionário'
        if 'Respostas do Questionário' in qmap:
            key_quest = _col_key('Questionário', 'Perguntas', 'Respostas do Questionário')
            row[key_quest] = qmap['Respostas do Questionário']

        # --- REMOVIDO: Imóvel (Você não pediu na lista) ---
        # --- REMOVIDO: Documentos (Você não pediu na lista) ---

        return row

    def executar(
        self,
        *,
        modo_extracao: str | None = None,
        cpfs: list[str] | None = None,
        cooldown_segundos: int | None = None,
        headless: bool | None = None,
        usuario: str | None = None,
        senha: str | None = None,
        max_candidatos: int | None = None,
        non_interactive: bool = False,
        prompt_credenciais: bool = False,
        debug: bool = False,
        debug_visual: bool = False,
        url_sistema: str = None,
    ):

        self.debug = bool(debug)
        self.debug_visual = bool(debug_visual)

        if cpfs is None:
            cpfs = []
        self.max_candidatos = max_candidatos

        def _env_cred(*nomes: str) -> str | None:
            for n in nomes:
                v = os.getenv(n)
                if v and v.strip():
                    return v.strip()
            return None

        try:
            if modo_extracao is None:
                if non_interactive:
                    raise ValueError("Modo de extração não informado.")
                modo_extracao, cpfs = self.solicitar_escopo_extracao()

            if cooldown_segundos is None:
                if non_interactive:
                    self.cooldown_segundos = 0
                else:
                    self.cooldown_segundos = self.solicitar_cooldown_segundos()
            else:
                self.cooldown_segundos = cooldown_segundos

            # ...restante do método (não exibido aqui)...
            if headless is None:
                if non_interactive:
                    headless = True
                else:
                    headless = self.solicitar_modo_execucao()

            # Se o usuário escolheu abrir o navegador no modo interativo,
            # ativa o highlight automaticamente.
            if (headless is False) and (not self.debug_visual):
                self.debug_visual = True

            # Solicita a URL do sistema antes das credenciais e da extração do nome do sistema
            if not self.base_url:
                self.solicitar_base_url()

            # Extrai o nome do sistema da URL para uso em nomes de arquivos e cabeçalhos
            if url_sistema:
                nome_sistema = _extrair_nome_sistema(url_sistema)
            elif hasattr(self, 'base_url') and self.base_url:
                nome_sistema = _extrair_nome_sistema(self.base_url)
            else:
                nome_sistema = "SISTEMA"

            if usuario is None:
                usuario = _env_cred('HABIBOT_USER', 'HABIBOT_USUARIO')
            if senha is None:
                senha = _env_cred('HABIBOT_PASS', 'HABIBOT_SENHA')

            if not usuario or not senha:
                if non_interactive and (not prompt_credenciais):
                    raise ValueError("Credenciais não encontradas.")
                usuario, senha = self.solicitar_credenciais()
        except KeyboardInterrupt:
            print("\n🛑 Configuração cancelada pelo usuário.")
            return

        self.iniciar(headless=headless)
        
        total = 0
        pagina = 1

        # Permite Ctrl+C interromper de forma mais previsível no Windows,
        # inclusive tentando fechar o navegador para destravar chamadas pendentes.
        global _ABORT_REQUESTED
        _ABORT_REQUESTED = False
        prev_handler = None
        try:
            prev_handler = signal.getsignal(signal.SIGINT)
        except:
            prev_handler = None

        def _on_sigint(sig, frame):
            _request_abort()
            try:
                print("\n🛑 Encerrado pelo usuário. Finalizando...")
            except:
                pass
            try:
                # Tenta fechar o driver para quebrar waits travados
                self.encerrar()
            except:
                pass
            raise KeyboardInterrupt

        try:
            signal.signal(signal.SIGINT, _on_sigint)
        except:
            pass

        try:
            print("\n──────────────────────────────────────────────────────────────────────────────────\n")
            print("🔐 Fazendo login...")
            self.login(usuario, senha)
            print("✅ Login OK")
            print("\n──────────────────────────────────────────────────────────────────────────────────\n")

            print("🛥️ Navegando para candidatos...\n")
            self.navegar_candidatos()
            print("✅ Página de candidatos carregada!")
            print("\n──────────────────────────────────────────────────────────────────────────────────\n")

            self.pasta_saida = os.path.join(os.getcwd(), 'Habibot - Dados Extraídos')
            os.makedirs(self.pasta_saida, exist_ok=True)

            # Logs em pasta separada e somente quando rodando como .py (não no EXE)
            gerar_logs_debug = not bool(getattr(sys, 'frozen', False))
            if gerar_logs_debug:
                self.pasta_debug = os.path.join(os.getcwd(), 'debug')
                os.makedirs(self.pasta_debug, exist_ok=True)
            else:
                self.pasta_debug = None

            agora = datetime.now()
            # Nome limpo para arquivos: sem espaços, sem acentos, sem caracteres especiais
            def _limpar_nome(s):
                import re, unicodedata
                s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')
                s = re.sub(r'[^a-zA-Z0-9_\-]', '_', s)
                s = re.sub(r'_+', '_', s)
                return s.strip('_')

            nome_sistema_limpo = _limpar_nome(nome_sistema.title())
            data_hora = agora.strftime('%d-%m-%Y às %H-%M-%S')
            base_nome = f"{nome_sistema_limpo}_{data_hora}"
            self.caminho_excel_final = os.path.join(self.pasta_saida, f"Candidatos_{base_nome}.xlsx")
            self.caminho_excel_visualizacao = os.path.join(self.pasta_saida, f"VISUALIZACAO_{base_nome}.txt")

            if gerar_logs_debug and self.pasta_debug:
                self.caminho_log_ordem = os.path.join(self.pasta_debug, f"log.ordem.tsv")
                self.ordem_logger = OrdemLeituraLogger(self.caminho_log_ordem)
                self.caminho_log_acoes = os.path.join(self.pasta_debug, f"log.acoes.tsv")
                self.acoes_logger = AcoesLogger(self.caminho_log_acoes)
            else:
                self.caminho_log_ordem = None
                self.ordem_logger = None
                self.caminho_log_acoes = None
                self.acoes_logger = None

            self.excel = ExcelTempoReal(self.caminho_excel_final, self.caminho_excel_visualizacao, nome_sistema=nome_sistema)
            print(f"📄 Excel FINAL: {self.caminho_excel_final}")
            print(f"👀 Visualização (texto): {self.caminho_excel_visualizacao}")
            if self.caminho_log_ordem:
                print(f"👁️ Log de ordem de leitura: {self.caminho_log_ordem}")
            if self.caminho_log_acoes:
                print(f"👁️ Log de ações: {self.caminho_log_acoes}")



            if modo_extracao == "cpfs":
                print("\n──────────────────────────────────────────────────────────────────────────────────\n")
                print(f"🎯 Modo por CPF: {len(cpfs)} candidato(s) alvo")
                self.garantir_lista_candidatos()
                for i, cpf11 in enumerate(cpfs, start=1):
                    _check_abort()
                    if self.max_candidatos and total >= self.max_candidatos:
                        print(f"\n🛑 Limite atingido (--max-candidatos={self.max_candidatos}). Encerrando...")
                        break
                    print("\n──────────────────────────────────────────────────────────────────────────────────\n")
                    print(f"🔎 Buscando CPF: {cpf11} | ({i}/{len(cpfs)})")
                    abriu_candidato = False
                    nome_titular = ''
                    try:
                        self.garantir_lista_candidatos()
                        qtd, indices = self.buscar_candidato_por_cpf(cpf11)
                        if qtd == 0:
                            nome_titular = 'CPF não encontrado'
                            print(f"\n👤 Encontrado: {nome_titular}")
                            print("")
                            print("⚠️ Nenhum candidato encontrado.")
                            total += 1
                            registro = {c.key: '' for c in COL_SPECS}
                            registro[_col_key('Titular', 'Dados Gerais', 'Nome')] = nome_titular
                            registro[_col_key('Titular', 'Dados Gerais', 'CPF/CNPJ')] = cpf11
                            registro[_col_key('Titular', 'Parecer Social', 'Observações')] = 'CPF não encontrado na lista de candidatos.'
                        else:
                            # CORREÇÃO: Usa a lista JÁ VISÍVEL para não pesquisar de novo
                            linhas_tela = self.obter_linhas()
                            
                            # Clica e Extrai usando o índice encontrado na busca anterior
                            if indices:
                                pos = indices[0] # Pega o primeiro match
                                _check_abort()
                                if pos < len(linhas_tela):
                                    linha_atual = linhas_tela[pos]
                                    
                                    # Pega nome só pra mostrar no log
                                    try:
                                        tds = linha_atual.find_elements(By.TAG_NAME, "td")
                                        nome_log = tds[1].text if len(tds) > 1 else cpf11
                                        print(f"\n👤 Encontrado: {nome_log}")
                                    except: print(f"\n👤 Encontrado: {cpf11}")
                                    print("")

                                    total += 1
                                    self.abrir_candidato(linha_atual)
                                    abriu_candidato = True
                                    registro = self.extrair_um()

                        self.excel.adicionar_linha(registro)
                        continue

                        if qtd > 1:
                            print(f"⚠️ Busca retornou {qtd} resultados. Extraindo os dois primeiros registros...")
                            max_registros = min(2, max(1, len(indices)))
                        else:
                            max_registros = 1

                        if not indices:
                            total += 1
                            registro = {c.key: '' for c in COL_SPECS}
                            registro[_col_key('Titular', 'Dados Gerais', 'Nome')] = f'Busca ambígua ({qtd} resultados)'
                            registro[_col_key('Titular', 'Dados Gerais', 'CPF/CNPJ')] = cpf11
                            self.excel.adicionar_linha(registro)
                            continue

                        for pos in range(max_registros):
                            _check_abort()
                            linha = self._obter_linha_por_cpf_posicao(cpf11, pos)
                            if not linha:
                                continue
                            total += 1
                            self.abrir_candidato(linha)
                            abriu_candidato = True
                            registro = self.extrair_um()
                            nome = (registro.get(_col_key('Titular', 'Dados Gerais', 'Nome')) or '').strip()
                            nome_titular = nome if nome else 'Nome não encontrado'
                            
                            self.excel.adicionar_linha(registro)
                            self.voltar_lista()
                            abriu_candidato = False

                    except KeyboardInterrupt:
                        raise
                    except Exception as e:
                        print(f"❌ Erro ao processar CPF {cpf11}: {e}")
                    finally:
                        if abriu_candidato:
                            self.voltar_lista()

                    if self.cooldown_segundos > 0 and i < len(cpfs):
                        print(f"⏳ Aguardando {self.cooldown_segundos}s...")
                        time.sleep(self.cooldown_segundos)

            else:
                while True:
                    _check_abort()
                    linhas = self.obter_linhas()
                    if not linhas:
                        print("⚠️ Nenhum candidato encontrado nesta página.")
                        break

                    print(f"\n📄 PÁGINA {pagina} - {len(linhas)} candidato(s)")

                    for idx in range(len(linhas)):
                        _check_abort()
                        if self.max_candidatos and total >= self.max_candidatos: break
                        linhas = self.obter_linhas()
                        if idx >= len(linhas): break

                        total += 1
                        print(f"\n🔍 Candidato #{total} (Pág {pagina}, Item {idx+1}/{len(linhas)})")

                        try:
                            self.abrir_candidato(linhas[idx])
                            registro = self.extrair_um()
                            nome = (registro.get(_col_key('Titular', 'Dados Gerais', 'Nome')) or '').strip()
                            print(f"👤 {nome if nome else 'Nome não encontrado'}")
                            self.excel.adicionar_linha(registro)

                        except KeyboardInterrupt:
                            raise
                        except Exception as e:
                            print(f"❌ Erro no candidato #{total}: {e}")
                        finally:
                            self.voltar_lista()

                        if self.cooldown_segundos > 0:
                            time.sleep(self.cooldown_segundos)

                    if self.max_candidatos and total >= self.max_candidatos:
                        print(f"\n🛑 Limite atingido (--max-candidatos={self.max_candidatos}). Encerrando...")
                        break

                    print("\n🔍 Verificando próxima página...")
                    if not self.proxima_pagina():
                        break
                    pagina += 1

            print("\n============================================================")
            print(f"✅ FINALIZADO! Total extraído: {total}")
            print(f"📄 Excel FINAL: {os.path.abspath(self.caminho_excel_final)}")
            print("============================================================\n")

        except KeyboardInterrupt:
            print("\n🛑 EXTRAÇÃO CANCELADA PELO USUÁRIO (Ctrl+C).")
            print(f"📊 Total extraído até agora: {total}")
            if self.caminho_excel_final:
                print(f"📄 Excel FINAL (parcial): {os.path.abspath(self.caminho_excel_final)}")

        except Exception as e:
            print("\n❌ Erro crítico na execução:")
            traceback.print_exc()

        finally:
            print("⏳ Finalizando e salvando dados...")
            try:
                if self.excel:
                    self.excel.flush()
            except:
                pass

            # Apaga o arquivo de visualização ao encerrar (conforme solicitado)
            try:
                if self.excel:
                    self.excel.limpar_visualizacoes()
            except:
                pass

            self.encerrar()
            try:
                if prev_handler is not None:
                    signal.signal(signal.SIGINT, prev_handler)
            except:
                pass
            print("👋 Bot encerrado.")


def main():
    mostrar_boas_vindas()
    pause_on_exit = True
    try:
        parser = argparse.ArgumentParser(add_help=True)
        modo_navegador = parser.add_mutually_exclusive_group()
        modo_navegador.add_argument('--headless', action='store_true', help='Executa sem janela do navegador (headless).')
        modo_navegador.add_argument('--headed', action='store_true', help='Executa com janela do navegador (não-headless).')
        parser.add_argument('--cooldown', type=int, default=None, help='Cooldown em segundos entre extrações (modo não-interativo).')
        parser.add_argument('--todos', action='store_true', help='Extrai todos os candidatos (modo não-interativo).')
        parser.add_argument('--cpfs-file', type=str, default=None, help='Arquivo texto com um CPF por linha (modo não-interativo).')
        parser.add_argument('--max-candidatos', type=int, default=None, help='Limita quantos registros processar (útil para conferência).')
        parser.add_argument('--non-interactive', action='store_true', help='Não faz perguntas; usa flags e variáveis de ambiente.')
        parser.add_argument('--prompt-credenciais', action='store_true', help='Permite digitar usuário/senha no terminal, mesmo com --non-interactive.')
        parser.add_argument('--no-pause', action='store_true', help='Não espera ENTER no final (útil para automação).')
        parser.add_argument('--debug', action='store_true', help='Logs detalhados para depuração.')
        parser.add_argument('--debug-visual', action='store_true', help='Destaca no navegador (não-headless) o elemento em uso.')
        args, _ = parser.parse_known_args()
        pause_on_exit = (not args.no_pause) and (not args.non_interactive)

        def _env_flag(nome: str) -> bool:
            try:
                v = (os.getenv(nome) or '').strip().lower()
                return v in ['1', 'true', 'yes', 'y', 'sim', 'on']
            except:
                return False

        debug = bool(args.debug) or _env_flag('HABIBOT_DEBUG')
        debug_visual = bool(args.debug_visual) or _env_flag('HABIBOT_DEBUG_VISUAL')

        if _DEPENDENCIAS_FALTANDO:
            print(f"\n❌ Dependências faltando: {', '.join(_DEPENDENCIAS_FALTANDO)}")
            print("Instale: pip install -r requirements.txt")
            return

        modo_extracao = None
        cpfs: list[str] | None = None
        if args.todos:
            modo_extracao = 'todos'
            cpfs = []
        elif args.cpfs_file:
            modo_extracao = 'cpfs'
            try:
                with open(args.cpfs_file, 'r', encoding='utf-8') as f:
                    linhas = [ln.strip() for ln in f.read().splitlines()]
                linhas = [ln for ln in linhas if ln]
                cpfs = HabibotBot.normalizar_cpfs(linhas)
            except Exception as e:
                raise ValueError(f"Falha ao ler --cpfs-file: {e}")
            if not cpfs:
                raise ValueError('Nenhum CPF válido encontrado em --cpfs-file.')

        headless_flag = None
        if args.headless:
            headless_flag = True
        elif args.headed:
            headless_flag = False

        HabibotBot().executar(
            modo_extracao=modo_extracao,
            cpfs=cpfs,
            cooldown_segundos=args.cooldown,
            headless=headless_flag,
            max_candidatos=args.max_candidatos,
            non_interactive=bool(args.non_interactive),
            prompt_credenciais=bool(args.prompt_credenciais),
            debug=debug,
            debug_visual=debug_visual,
        )

        if pause_on_exit:
            input("\nPressione ENTER para sair...")

    except KeyboardInterrupt:
        print("\n🛑 Encerrado.")
        if pause_on_exit:
            try:
                input("\nPressione ENTER para sair...")
            except:
                pass
    except Exception as e:
        print(f"\n❌ Erro no main: {e}")
        if pause_on_exit:
            try:
                input("\nPressione ENTER para sair...")
            except:
                pass

if __name__ == '__main__':
    main()