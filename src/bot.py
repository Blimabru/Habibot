"""Módulo principal do bot Habibot.

Contém a classe HabibotBot que coordena todo o processo de automação.
"""

import os
import re
import signal
import subprocess
import time
from datetime import datetime
from pathlib import Path

import pwinput

from .dependencies import (
    webdriver, By, Keys, WebDriverWait, EC, Service, Options, ActionChains,
    WebDriverException, SessionNotCreatedException, TimeoutException,
    ChromeDriverManager
)
from .config import _diretorio_base_execucao, _encontrar_chromedriver_local, _extrair_nome_sistema
from .loggers import AcoesLogger
from .excel_handler import ExcelTempoReal
from .extractor import ExtratorHabibot
from .schema import _request_abort, COL_SPECS

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
        extrator = ExtratorHabibot(
            self.driver,
            self.wait,
            debug=bool(self.debug),
            debug_visual=bool(self.debug_visual) and (not self._headless),
            ordem_logger=self.ordem_logger,
            acoes_logger=self.acoes_logger,
        )
        row = {}

        def _log_secao(aba: str, secao: str):
            try:
                print(f"🧩 Extraindo: {aba} > {secao}")
            except:
                pass
            try:
                self._log_acao('EXTRAIR_SECAO', detalhe=f"{aba} > {secao}")
            except:
                pass
        
        # 1) Titular
        _log_secao('Titular', 'Dados do Empreendimento')
        titular = extrator.extrair_titular_mapa() or {}
        row.update(titular)

        # 2) Segundo Titular (Modal)
        _log_secao('Segundo Titular', 'Dados Gerais')
        st = extrator.extrair_segundo_titular() or {}
        row.update(st)

        # 3) Composição Familiar (Modal)
        _log_secao('Composição Familiar', 'Dados Gerais')
        cf = extrator.extrair_composicao_familiar() or {}
        row.update(cf)

        # 4) Renda (Modal)
        _log_secao('Renda', 'Dados da Renda')
        renda = extrator.extrair_renda() or {}
        row.update(renda)

        # 5) Endereço
        _log_secao('Endereço', 'Dados do Endereço')
        end = extrator.extrair_endereco() or {}
        row.update(end)

        # 6) Imóvel
        _log_secao('Imóvel', 'Aquisição e Infraestrutura')
        im = extrator.extrair_imovel() or {}
        row.update(im)

        # 7) Documentos (usa o mapa de status para preencher TODAS as colunas de documentos do schema)
        _log_secao('Documentos', 'Titular')
        doc_map = extrator.extrair_documentos_mapa() or {}
        for spec in COL_SPECS:
            if spec.aba == 'Documentos':
                row[_col_key(spec.aba, spec.secao, spec.campo)] = doc_map.get(spec.campo, '')

        # 8) Questionário (colunas por pergunta)
        _log_secao('Questionário', 'Perguntas')
        qmap = extrator.extrair_questionario_mapa() or {}
        for spec in COL_SPECS:
            if spec.aba == 'Questionário' and spec.secao == 'Perguntas':
                row[_col_key(spec.aba, spec.secao, spec.campo)] = qmap.get(_norm_texto_chave(spec.campo), '')

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
                            # Extrai o nome do titular diretamente da lista antes de clicar em Editar
                            nome_titular = ''
                            for pos in indices:
                                _check_abort()
                                linha = self._obter_linha_por_cpf_posicao(cpf11, pos)
                                if not linha:
                                    continue
                                # Busca o <td> do nome do titular na linha
                                try:
                                    tds = linha.find_elements(By.TAG_NAME, "td")
                                    nome_coluna = ''
                                    # Tenta identificar a coluna do nome pelo cabeçalho da tabela
                                    thead = None
                                    try:
                                        thead = linha.find_element(By.XPATH, "../../thead")
                                    except Exception:
                                        pass
                                    idx_nome = 1  # fallback: segunda coluna
                                    if thead:
                                        ths = thead.find_elements(By.TAG_NAME, "th")
                                        for idx, th in enumerate(ths):
                                            thtxt = (th.text or '').strip().lower()
                                            if 'nome' in thtxt:
                                                idx_nome = idx
                                                break
                                    if tds and len(tds) > idx_nome:
                                        nome_coluna = (tds[idx_nome].text or '').strip()
                                    nome_titular = nome_coluna if nome_coluna else cpf11
                                except Exception:
                                    nome_titular = cpf11
                                break  # Só pega o primeiro encontrado
                            print(f"\n👤 Encontrado: {nome_titular}")
                            print("")
                            # Agora sim clica em Editar e extrai os dados
                            for pos in indices:
                                _check_abort()
                                linha = self._obter_linha_por_cpf_posicao(cpf11, pos)
                                if not linha:
                                    continue
                                total += 1
                                self.abrir_candidato(linha)
                                abriu_candidato = True
                                registro = self.extrair_um()
                                break  # Só pega o primeiro encontrado

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


