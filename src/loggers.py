"""Módulo de logging para o bot Habibot.

Contém classes para registrar:
- Ordem de leitura de campos (OrdemLeituraLogger)
- Ações do bot (AcoesLogger)
"""

import os
import re
import time
from datetime import datetime


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
