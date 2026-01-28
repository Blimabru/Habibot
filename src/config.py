"""Módulo de configuração e funções utilitárias do Habibot.

Contém:
- Informações de versão
- Funções utilitárias (paths, terminal, texto)
- Constantes globais
"""

import os
import re
import shutil
import sys
import textwrap
from pathlib import Path


def ler_version_info(path=None):
    """Lê metadados do arquivo version_info.txt."""
    path = path or (Path(__file__).parent.parent / 'assets' / 'build' / 'version_info.txt')
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
    return Path(__file__).resolve().parent.parent


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
    """Retorna a largura do terminal."""
    try:
        return max(60, shutil.get_terminal_size((padrao, 20)).columns)
    except:
        return max(60, padrao)


SEPARADOR_TELA = "─"


def _linha(char: str = SEPARADOR_TELA, largura: int | None = None) -> str:
    """Retorna uma linha de caracteres para separação visual."""
    w = largura or _terminal_width()
    return char * w


def _extrair_nome_sistema(url: str) -> str:
    """Extrai o nome do sistema a partir do domínio (ex: app.habisoft.com.br -> HABISOFT)."""
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
    """Exibe mensagem de boas-vindas do bot."""
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
    """Normaliza texto para comparação (remove acentos, pontuação, etc)."""
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
