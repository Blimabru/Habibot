"""Módulo de extração de dados do Habibot.

Contém a classe ExtratorHabibot que realiza a extração dos dados
dos candidatos do sistema.
"""

import re
from .dependencies import By, WebDriverWait, EC, TimeoutException
from .loggers import OrdemLeituraLogger, AcoesLogger
from .schema import _check_abort, _col_key
from .config import _norm_texto_chave

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
        """Versão estendida: retorna (encontrou, valor, qtd_matches, label_encontrado).

        Importante: se o campo for encontrado mesmo vazio, retorna encontrou=True.
        """
        _check_abort()
        try:
            texto_busca_norm = _norm_texto_chave(texto_label)
            if not texto_busca_norm:
                return False, "", 0, ""

            labels = self.driver.find_elements(By.TAG_NAME, "label")
            candidatos = []
            for label in labels:
                try:
                    if not label.is_displayed():
                        continue
                    texto_lbl = label.text or ""
                    norm_lbl = _norm_texto_chave(texto_lbl)
                    if not norm_lbl:
                        continue
                    if norm_lbl == texto_busca_norm:
                        score = 0
                    elif norm_lbl.startswith(texto_busca_norm):
                        score = 1
                    elif texto_busca_norm in norm_lbl:
                        score = 2
                    else:
                        continue
                    candidatos.append((score, len(norm_lbl), label))
                except:
                    continue

            candidatos.sort(key=lambda x: (x[0], x[1]))

            qtd_matches = len(candidatos)

            for _, _, label in candidatos:
                try:
                    _check_abort()
                    self._highlight(label, rotulo=f"LABEL: {texto_label}")

                    # Tenta associação direta via atributo "for"
                    try:
                        for_attr = (label.get_attribute('for') or '').strip()
                        if for_attr:
                            campo = self.driver.find_element(By.ID, for_attr)
                            self._highlight(campo, rotulo=f"CAMPO: {texto_label}")
                            valor = self._valor_campo(campo)
                            # Marca como lido (verde)
                            self._mark_read(label, rotulo=f"LIDO: {texto_label}")
                            self._mark_read(campo, rotulo=f"LIDO: {texto_label}")
                            return True, valor, qtd_matches, (label.text or "")
                    except:
                        pass

                    # Tenta irmão imediato
                    try:
                        siblings = label.find_elements(By.XPATH, "following-sibling::input[1] | following-sibling::select[1] | following-sibling::textarea[1]")
                        if siblings:
                            self._highlight(siblings[0], rotulo=f"CAMPO: {texto_label}")
                            valor = self._valor_campo(siblings[0])
                            self._mark_read(label, rotulo=f"LIDO: {texto_label}")
                            self._mark_read(siblings[0], rotulo=f"LIDO: {texto_label}")
                            return True, valor, qtd_matches, (label.text or "")
                    except:
                        pass

                    # Tenta dentro do mesmo container (pai)
                    try:
                        container = label.find_element(By.XPATH, "..")
                        campos = container.find_elements(By.CSS_SELECTOR, "input, select, textarea")
                        for campo in campos:
                            try:
                                if campo.is_displayed():
                                    self._highlight(campo, rotulo=f"CAMPO: {texto_label}")
                                    valor = self._valor_campo(campo)
                                    self._mark_read(label, rotulo=f"LIDO: {texto_label}")
                                    self._mark_read(campo, rotulo=f"LIDO: {texto_label}")
                                    return True, valor, qtd_matches, (label.text or "")
                            except:
                                continue
                    except:
                        pass
                except:
                    continue

            # Fallback: algumas telas apresentam "perguntas" em <h6> (sem <label>).
            # Ex.: perguntas do Minha Casa Minha Vida.
            try:
                h6s = self.driver.find_elements(By.TAG_NAME, "h6")
                cand_h6 = []
                for h in h6s:
                    try:
                        if not h.is_displayed():
                            continue
                        txt = h.text or ""
                        n = _norm_texto_chave(txt)
                        if not n:
                            continue
                        if n == texto_busca_norm:
                            score = 0
                        elif n.startswith(texto_busca_norm):
                            score = 1
                        elif texto_busca_norm in n:
                            score = 2
                        else:
                            continue
                        cand_h6.append((score, len(n), h))
                    except:
                        continue
                cand_h6.sort(key=lambda x: (x[0], x[1]))

                def _tentar_ler_em_container(node):
                    try:
                        radios_checked = node.find_elements(By.CSS_SELECTOR, "input[type='radio']:checked")
                        if radios_checked:
                            el = radios_checked[0]
                            self._highlight(el, rotulo=f"CAMPO: {texto_label}")
                            val = ""
                            try:
                                lbl = el.find_element(By.XPATH, "./following-sibling::label")
                                val = self._texto_limpo(lbl.text)
                                if lbl:
                                    self._mark_read(lbl, rotulo=f"LIDO: {texto_label}")
                            except:
                                # fallback simples: radio checked
                                val = "Sim"
                            self._mark_read(el, rotulo=f"LIDO: {texto_label}")
                            return True, val
                    except:
                        pass

                    try:
                        checks = node.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                        if checks:
                            # Se houver pelo menos 1 checkbox visível, marca os selecionados.
                            marcados = []
                            for ck in checks:
                                try:
                                    if not ck.is_displayed():
                                        continue
                                    if not ck.is_selected():
                                        continue
                                    self._mark_read(ck, rotulo=f"LIDO: {texto_label}")
                                    try:
                                        lbl = ck.find_element(By.XPATH, "./following-sibling::label")
                                        t = self._texto_limpo(lbl.text)
                                        if t:
                                            marcados.append(t)
                                            self._mark_read(lbl, rotulo=f"LIDO: {texto_label}")
                                    except:
                                        marcados.append("Sim")
                                except:
                                    continue
                            if marcados:
                                return True, "\n".join(marcados)
                    except:
                        pass

                    try:
                        campos = node.find_elements(By.CSS_SELECTOR, "select, textarea, input:not([type='radio']):not([type='checkbox'])")
                        for c in campos:
                            try:
                                if not c.is_displayed():
                                    continue
                                self._highlight(c, rotulo=f"CAMPO: {texto_label}")
                                val = self._valor_campo(c)
                                self._mark_read(c, rotulo=f"LIDO: {texto_label}")
                                return True, val
                            except:
                                continue
                    except:
                        pass

                    return False, ""

                for _, _, h in cand_h6:
                    _check_abort()
                    self._highlight(h, rotulo=f"PERGUNTA: {texto_label}")
                    ok = False
                    val = ""
                    # tenta em níveis acima (row/section)
                    node = h
                    for _ in range(3):
                        try:
                            ok, val = _tentar_ler_em_container(node)
                            if ok:
                                break
                        except:
                            pass
                        try:
                            node = node.find_element(By.XPATH, "..")
                        except:
                            break

                    # Sempre marca o <h6> como lido, mesmo sem input
                    self._mark_read(h, rotulo=f"LIDO: {texto_label}")
                    # Se não achou input, tenta marcar opções "Não" ou "Selecione" se existirem como texto irmão
                    if not ok:
                        try:
                            # Busca elementos irmãos com texto "Não" ou "Selecione"
                            siblings = h.find_elements(By.XPATH, "following-sibling::*")
                            for sib in siblings:
                                txtsib = (sib.text or "").strip().lower()
                                if txtsib in ("não", "nao", "selecione"):
                                    self._mark_read(sib, rotulo=f"LIDO: {txtsib}")
                        except:
                            pass
                    return True, val, max(qtd_matches, len(cand_h6)), (h.text or "")
            except:
                pass

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
        """Extrai itens de uma tabela abrindo o modal de edição em cada linha.

        Retorna uma lista de dicionários (1 por linha), onde as chaves são os nomes
        finais de coluna (valores em `campos_mapeados` e em `colunas_tabela`).

        - `campos_mapeados`: {label_visivel: nome_coluna_saida}
        - `colunas_tabela`: {indice_td: nome_coluna_saida}
        """
        itens: list[dict[str, str]] = []
        try:
            tabelas = self.driver.find_elements(By.TAG_NAME, "table")
            tabela_alvo = None
            for t in tabelas:
                if t.is_displayed():
                    tabela_alvo = t
                    break
            if not tabela_alvo:
                return []

            linhas = tabela_alvo.find_elements(By.CSS_SELECTOR, "tbody tr")
            if not linhas:
                return []

            for i in range(len(linhas)):
                _check_abort()
                item: dict[str, str] = {}
                try:
                    # Rebusca a linha para evitar stale element
                    tabelas = self.driver.find_elements(By.TAG_NAME, "table")
                    tabela_atual = None
                    for t in tabelas:
                        if t.is_displayed():
                            tabela_atual = t
                            break
                    if not tabela_atual:
                        break
                    linhas_atual = tabela_atual.find_elements(By.CSS_SELECTOR, "tbody tr")
                    if i >= len(linhas_atual):
                        break
                    linha = linhas_atual[i]

                    if colunas_tabela:
                        try:
                            tds = linha.find_elements(By.TAG_NAME, "td")
                            for idx_td, nome_saida in colunas_tabela.items():
                                try:
                                    if 0 <= int(idx_td) < len(tds):
                                        item[nome_saida] = self._texto_limpo(tds[int(idx_td)].text)
                                except:
                                    continue
                        except:
                            pass

                    # Abre modal
                    try:
                        btn_editar = linha.find_element(By.CSS_SELECTOR, "button.btn-soft-info, .ri-pencil-line")
                    except:
                        btn_editar = None
                    if not btn_editar:
                        continue

                    self._dbg(f"📝 Abrindo modal de edição (linha {i+1})")
                    self._highlight(btn_editar, rotulo=f"EDITAR LINHA {i+1}")
                    self._log_acao('MODAL_ABRIR', detalhe=f"linha={i+1}")
                    self.driver.execute_script("arguments[0].click();", btn_editar)
                    time.sleep(1.2)

                    # Extrai campos dentro do modal.
                    # Aceita 2 formatos:
                    # 1) {label_visivel: nome_coluna_saida}
                    # 2) {nome_coluna_saida: [label1, label2, ...]}  (variantes, mantendo a ordem das colunas)
                    if campos_mapeados:
                        try:
                            primeiro_valor = next(iter(campos_mapeados.values()))
                        except:
                            primeiro_valor = None

                        if isinstance(primeiro_valor, list):
                            # Formato 2
                            for nome_saida, labels_busca in campos_mapeados.items():
                                try:
                                    if item.get(nome_saida) is not None and (item.get(nome_saida) or "") != "":
                                        continue
                                    for label_busca in (labels_busca or []):
                                        self._dbg(f"   🔎 {nome_saida} <= '{label_busca}'")
                                        ok, val = self._extrair_por_label_simples_found(label_busca)
                                        if ok:
                                            item[nome_saida] = val
                                            break
                                except:
                                    item[nome_saida] = item.get(nome_saida, "")
                        else:
                            # Formato 1 (compat)
                            for label_busca, nome_saida in (campos_mapeados or {}).items():
                                try:
                                    atual = item.get(nome_saida) or ""
                                    if atual:
                                        continue
                                    self._dbg(f"   🔎 {nome_saida} <= '{label_busca}'")
                                    item[nome_saida] = self._extrair_por_label_simples(label_busca)
                                except:
                                    item[nome_saida] = item.get(nome_saida, "")

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
                        try:
                            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                        except:
                            pass

                    self._log_acao('MODAL_FECHAR', detalhe=f"linha={i+1}")

                    time.sleep(0.8)
                    itens.append(item)
                except Exception:
                    try:
                        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    except:
                        pass
                    time.sleep(0.6)
                    continue
        except Exception:
            return []

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
        try:
            WebDriverWait(self.driver, 6, poll_frequency=0.2).until(
                lambda d: (
                    self._extrair_por_label_simples('CPF/CNPJ')
                    or self._extrair_por_label_simples('Nome')
                )
            )
        except:
            pass
        mapa: dict[str, str] = {}

        # Ordem baseada na tela (seções e campos). Cada campo é lido 1x (na sequência).
        # Chave do retorno: aba|||secao|||campo (evita colisões com campos repetidos).
        passos: list[tuple[str, str, str, list[tuple[str, str]]]] = [
            # Dados do Empreendimento
            ('Titular', 'Dados do Empreendimento', 'Tipo de Cadastro', [('label', 'Tipo de Cadastro')]),
            ('Titular', 'Dados do Empreendimento', 'Loteamento', [('label', 'Loteamento')]),

            # Minha Casa, Minha Vida (perguntas)
            (
                'Titular',
                'Minha Casa, Minha Vida',
                'Você ou alguém da sua família possui imóvel (como herança, doação ou aluguel vitalício) em qualquer lugar do Brasil?',
                [('label', 'Você ou alguém da sua família possui imóvel')],
            ),
            (
                'Titular',
                'Minha Casa, Minha Vida',
                'A renda familiar mensal ultrapassa R$ 2.850,00?',
                [('label', 'A renda familiar mensal ultrapassa')],
            ),
            (
                'Titular',
                'Minha Casa, Minha Vida',
                'Mora no Município há pelo menos 2 anos na mesma cidade?',
                [('label', 'Mora no Município há pelo menos 2 anos')],
            ),

            # Dados Gerais
            ('Titular', 'Dados Gerais', 'Nome', [('label', 'Nome')]),
            ('Titular', 'Dados Gerais', 'CPF/CNPJ', [('label', 'CPF/CNPJ')]),
            ('Titular', 'Dados Gerais', 'Tipo de Pessoa', [('label', 'Tipo de Pessoa')]),
            ('Titular', 'Dados Gerais', 'Tipo de Documento', [('label', 'Tipo de Documento')]),
            ('Titular', 'Dados Gerais', 'Nº Documento', [('label', 'Nº Documento'), ('label', 'RG')]),
            (
                'Titular',
                'Dados Gerais',
                'Data de Emissão',
                [('id', 'dtEmissaoRg'), ('label', 'Data de Emissão do RG'), ('label', 'Data de Emissão')],
            ),
            ('Titular', 'Dados Gerais', 'Órgão Emissor', [('label', 'Órgão Emissor do RG'), ('label', 'Órgão Emissor')]),
            ('Titular', 'Dados Gerais', 'UF', [('label', 'UF do RG'), ('label', 'UF')]),
            ('Titular', 'Dados Gerais', 'Data de Vencimento', [('label', 'Data de Vencimento')]),
            ('Titular', 'Dados Gerais', 'Naturalidade (Cidade)', [('label', 'Naturalidade (Cidade)'), ('label', 'Naturalidade')]),
            ('Titular', 'Dados Gerais', 'UF Naturalidade (Estado)', [('label', 'UF Naturalidade (Estado)'), ('label', 'UF Naturalidade')]),
            ('Titular', 'Dados Gerais', 'Nacionalidade (País)', [('label', 'Nacionalidade (País)'), ('label', 'Nacionalidade')]),
            ('Titular', 'Dados Gerais', 'Data de Nascimento', [('label', 'Data de Nascimento')]),
            ('Titular', 'Dados Gerais', 'Raça', [('label', 'Raça')]),
            ('Titular', 'Dados Gerais', 'Gênero', [('label', 'Gênero')]),
            ('Titular', 'Dados Gerais', 'Contrato Distratado ou Rescindido Involutariamente', [('label', 'Contrato Distratado')]),
            ('Titular', 'Dados Gerais', 'Recebe Atendimento Socio-assistencial do Município', [('label', 'Recebe Atendimento Socio-assistencial')]),
            ('Titular', 'Dados Gerais', 'Segurança frequente no bairro (Ronda GCM, PM)', [('label', 'Segurança frequente no bairro')]),
            ('Titular', 'Dados Gerais', 'Situação Atual do Domicílio', [('label', 'Situação Atual do Domicílio')]),
            ('Titular', 'Dados Gerais', 'Tipo de Beneficiário Programa Social', [('label', 'Tipo de Beneficiário Programa Social')]),
            ('Titular', 'Dados Gerais', 'Proprietário do Imóvel?', [('label', 'Proprietário do Imóvel')]),
            ('Titular', 'Dados Gerais', 'Recebeu/Participou de REURB?', [('label', 'REURB')]),

            # Informações do CadÚnico
            ('Titular', 'Informações do CadÚnico', 'A família esta inscrita no CAD Único?', [('label', 'A família esta inscrita no CAD Único')]),
            (
                'Titular',
                'Informações do CadÚnico',
                'O responsável familiar é mulher monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                [('label', 'mulher monoparental')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'O responsável familiar é homem monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                [('label', 'homem monoparental')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família unipessoal, sem dependente, declarada no Cadúnico residente no domícilio?',
                [('label', 'unipessoal')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                [('label', 'violência doméstica')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'pessoa(s) negra(s)')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'pessoa(s) LGBT')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'Quilombola')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'povos tradicionais')],
            ),
            (
                'Titular',
                'Informações do CadÚnico',
                'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'situação de rua')],
            ),
            ('Titular', 'Informações do CadÚnico', 'Programas CAD Único', [('label', 'Programas CAD Único')]),
            ('Titular', 'Informações do CadÚnico', 'Mulheres Empreendedoras', [('label', 'Mulheres Empreendedoras')]),
            ('Titular', 'Informações do CadÚnico', 'Tratamento Oncológico/Hemodialise', [('label', 'Tratamento Oncológico')]),

            # Cadastro Preferencial
            ('Titular', 'Cadastro Preferencial', 'Cadastro Preferencial', [('label', 'Cadastro Preferencial')]),
            ('Titular', 'Cadastro Preferencial', 'Deficiência', [('label', 'Deficiência')]),
            ('Titular', 'Cadastro Preferencial', 'Doença', [('label', 'Doença')]),
            ('Titular', 'Cadastro Preferencial', 'Possui Microcefalia', [('label', 'Possui Microcefalia')]),
            ('Titular', 'Cadastro Preferencial', 'Tem Veículo Próprio', [('label', 'Tem Veículo Próprio')]),

            # Dados dos Pais
            ('Titular', 'Dados dos Pais', 'Nome do Pai', [('label', 'Nome do Pai')]),
            ('Titular', 'Dados dos Pais', 'Nome da Mãe', [('label', 'Nome da Mãe')]),

            # Educação e Profissão
            ('Titular', 'Educação e Profissão', 'Escolaridade', [('label', 'Escolaridade')]),
            ('Titular', 'Educação e Profissão', 'Escola', [('label', 'Escola')]),
            ('Titular', 'Educação e Profissão', 'Situação de Emprego', [('label', 'Situação de Emprego')]),
            ('Titular', 'Educação e Profissão', 'Profissão', [('label', 'Profissão')]),
            ('Titular', 'Educação e Profissão', 'Profissão (se Outro)', [('label', 'Profissão (se Outro)'), ('label', 'Profissão (se outro)')]),
            ('Titular', 'Educação e Profissão', 'Tempo de Serviço', [('label', 'Tempo de Serviço')]),

            # Situação Marital
            ('Titular', 'Situação Marital', 'Estado Civil', [('label', 'Estado Civil')]),
            (
                'Titular',
                'Situação Marital',
                'Data do Casamento/União',
                [('label', 'Data do Casamento/União'), ('label', 'Data do Casamento')],
            ),
            ('Titular', 'Situação Marital', 'Regime do Casamento', [('label', 'Regime do Casamento')]),
            (
                'Titular',
                'Situação Marital',
                'Regime do Casamento (se Outro)',
                [('label', 'Regime do Casamento (se Outro)'), ('label', 'Regime do Casamento (se outro)')],
            ),

            # CAD
            ('Titular', 'CAD', 'NIS (PIS/PASEP)', [('label', 'NIS (PIS/PASEP)'), ('label', 'NIS')]),
            ('Titular', 'CAD', 'Nº da Carteira de Trabalho', [('label', 'Nº da Carteira de Trabalho'), ('label', 'Nº da Carteira')]),
            ('Titular', 'CAD', 'Nº de Série da Carteira de Trabalho', [('label', 'Nº de Série da Carteira de Trabalho'), ('label', 'Nº de Série')]),
            ('Titular', 'CAD', 'UF da Carteira de Trabalho', [('label', 'UF da Carteira de Trabalho'), ('label', 'UF da Carteira')]),
            ('Titular', 'CAD', 'Nº do Título de Eleitor', [('label', 'Nº do Título de Eleitor'), ('label', 'Título de Eleitor')]),
            ('Titular', 'CAD', 'Zona Eleitoral', [('label', 'Zona Eleitoral')]),
            ('Titular', 'CAD', 'Seção Eleitoral', [('label', 'Seção Eleitoral')]),
            ('Titular', 'CAD', 'Data de Emissão', [('id', 'dtEmissaoTituloEleitor'), ('label', 'Data de Emissão')]),
            ('Titular', 'CAD', 'UF Título de Eleitor', [('label', 'UF Título de Eleitor')]),
            ('Titular', 'CAD', 'Cidade do Título de Eleitor', [('label', 'Cidade do Título de Eleitor')]),

            # Dados de Contato
            ('Titular', 'Dados de Contato', 'Email', [('label', 'Email')]),
            ('Titular', 'Dados de Contato', 'Telefone', [('label', 'Telefone')]),
            ('Titular', 'Dados de Contato', 'Telefone Celular', [('label', 'Telefone Celular'), ('label', 'Celular')]),
            ('Titular', 'Dados de Contato', 'Telefone Comercial', [('label', 'Telefone Comercial'), ('label', 'Tel. Comercial')]),
            (
                'Titular',
                'Dados de Contato',
                'Telefone Para Recados',
                [('label', 'Telefone Para Recados'), ('label', 'Telefone para Recados'), ('label', 'Recado')],
            ),

            # Parecer Social
            ('Titular', 'Parecer Social', 'Diagnóstico Social', [('label', 'Diagnóstico Social')]),
            ('Titular', 'Parecer Social', 'Observações', [('label', 'Observações')]),
        ]

        for aba, secao, campo, tentativas in passos:
            self._dbg(f"🔎 {aba} > {secao} > {campo}")
            _, valor = self.ler_campo_ordenado(aba=aba, secao=secao, campo=campo, tentativas=tentativas)
            mapa[_col_key(aba, secao, campo)] = valor

        return mapa

    # --- ABAS COM MODAL ---
    def extrair_segundo_titular(self):
        if not self.clicar_aba('Segundo Titular'):
            return {}

        aba = 'Segundo Titular'

        # Monta variantes por CHAVE TRIPLA (aba|||secao|||campo) para evitar colisões (ex.: "Data de Emissão").
        especificacao: list[tuple[str, str, list[tuple[str, str]]]] = []

        # Triagem
        for pergunta in [
            'Você ou alguém da sua família possui imóvel (como herança, doação ou aluguel vitalício) em qualquer lugar do Brasil?',
            'A renda familiar mensal ultrapassa R$ 2.850,00?',
            'Mora no Município há pelo menos 2 anos na mesma cidade?',
        ]:
            especificacao.append(('Perguntas de Triagem', pergunta, [('label', pergunta)]))

        # Dados gerais
        especificacao += [
            ('Dados Gerais', 'Nome', [('label', 'Nome')]),
            ('Dados Gerais', 'CPF/CNPJ', [('label', 'CPF/CNPJ'), ('label', 'CPF')]),
            ('Dados Gerais', 'Tipo de Pessoa', [('label', 'Tipo de Pessoa')]),
            ('Dados Gerais', 'Tipo de Documento', [('label', 'Tipo de Documento')]),
            ('Dados Gerais', 'Nº Documento', [('label', 'Nº Documento'), ('label', 'RG')]),
            ('Dados Gerais', 'Data de Emissão', [('label', 'Data de Emissão do RG'), ('label', 'Data de Emissão')]),
            ('Dados Gerais', 'Órgão Emissor', [('label', 'Órgão Emissor do RG'), ('label', 'Órgão Emissor')]),
            ('Dados Gerais', 'UF', [('label', 'UF do RG'), ('label', 'UF')]),
            ('Dados Gerais', 'Data de Vencimento', [('label', 'Data de Vencimento')]),
            ('Dados Gerais', 'Naturalidade (Cidade)', [('label', 'Naturalidade (Cidade)'), ('label', 'Naturalidade')]),
            ('Dados Gerais', 'UF Naturalidade (Estado)', [('label', 'UF Naturalidade (Estado)'), ('label', 'UF Naturalidade')]),
            ('Dados Gerais', 'Nacionalidade (País)', [('label', 'Nacionalidade (País)'), ('label', 'Nacionalidade')]),
            ('Dados Gerais', 'Data de Nascimento', [('label', 'Data de Nascimento')]),
            ('Dados Gerais', 'Raça', [('label', 'Raça')]),
            ('Dados Gerais', 'Gênero', [('label', 'Gênero')]),
            ('Dados Gerais', 'Contrato Distratado ou Rescindido Involutariamente', [('label', 'Contrato Distratado')]),
            ('Dados Gerais', 'Recebe Atendimento Socio-assistencial do Município', [('label', 'Recebe Atendimento Socio-assistencial')]),
            ('Dados Gerais', 'Segurança frequente no bairro (Ronda GCM, PM)', [('label', 'Segurança frequente no bairro')]),
            ('Dados Gerais', 'Situação Atual do Domicílio', [('label', 'Situação Atual do Domicílio')]),
            ('Dados Gerais', 'Tipo de Beneficiário Programa Social', [('label', 'Tipo de Beneficiário Programa Social')]),
            ('Dados Gerais', 'Proprietário do Imóvel?', [('label', 'Proprietário do Imóvel')]),
            ('Dados Gerais', 'Recebeu/Participou de REURB?', [('label', 'REURB')]),
        ]

        # CadÚnico
        especificacao += [
            ('Informações do CadÚnico', 'A família esta inscrita no CAD Único?', [('label', 'A família esta inscrita no CAD Único')]),
            (
                'Informações do CadÚnico',
                'O responsável familiar é mulher monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                [('label', 'mulher monoparental')],
            ),
            (
                'Informações do CadÚnico',
                'O responsável familiar é homem monoparental e responsável por maior número de filhos de 0 a 17 anos declarado no Cadúnico?',
                [('label', 'homem monoparental')],
            ),
            (
                'Informações do CadÚnico',
                'Família unipessoal, sem dependente, declarada no Cadúnico residente no domícilio?',
                [('label', 'unipessoal')],
            ),
            (
                'Informações do CadÚnico',
                'A requerente é mulher em situação de violência doméstica/familiar, com medida protetiva de urgência declarada?',
                [('label', 'violência doméstica')],
            ),
            (
                'Informações do CadÚnico',
                'Família com pessoa(s) negra(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'pessoa(s) negra(s)')],
            ),
            (
                'Informações do CadÚnico',
                'Família com pessoa(s) LGBT na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'pessoa(s) LGBT')],
            ),
            (
                'Informações do CadÚnico',
                'Família com pessoa(s) Quilombola(s) na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'Quilombola')],
            ),
            (
                'Informações do CadÚnico',
                'Família com pessoa(s) de povos tradicionais na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'povos tradicionais')],
            ),
            (
                'Informações do CadÚnico',
                'Família com pessoa(s) em situação de rua na sua composição, declarada no Cadúnico residente no domícilio?',
                [('label', 'situação de rua')],
            ),
            ('Informações do CadÚnico', 'Programas CAD Único', [('label', 'Programas CAD Único')]),
            ('Informações do CadÚnico', 'Mulheres Empreendedoras', [('label', 'Mulheres Empreendedoras')]),
            ('Informações do CadÚnico', 'Tratamento Oncológico/Hemodialise', [('label', 'Tratamento Oncológico')]),
        ]

        # Cadastro Preferencial
        especificacao += [
            ('Cadastro Preferencial', 'Cadastro Preferencial', [('label', 'Cadastro Preferencial')]),
            ('Cadastro Preferencial', 'Deficiência', [('label', 'Deficiência')]),
            ('Cadastro Preferencial', 'Doença', [('label', 'Doença')]),
            ('Cadastro Preferencial', 'Possui Microcefalia', [('label', 'Possui Microcefalia')]),
            ('Cadastro Preferencial', 'Tem Veículo Próprio', [('label', 'Tem Veículo Próprio')]),
        ]

        # Pais
        especificacao += [
            ('Dados dos Pais', 'Nome do Pai', [('label', 'Nome do Pai')]),
            ('Dados dos Pais', 'Nome da Mãe', [('label', 'Nome da Mãe')]),
        ]

        # Educação e Profissão
        especificacao += [
            ('Educação e Profissão', 'Escolaridade', [('label', 'Escolaridade')]),
            ('Educação e Profissão', 'Escola', [('label', 'Escola')]),
            ('Educação e Profissão', 'Situação de Emprego', [('label', 'Situação de Emprego')]),
            ('Educação e Profissão', 'Profissão', [('label', 'Profissão')]),
            ('Educação e Profissão', 'Profissão (se Outro)', [('label', 'Profissão (se Outro)'), ('label', 'Profissão (se outro)')]),
            ('Educação e Profissão', 'Tempo de Serviço', [('label', 'Tempo de Serviço')]),
        ]

        # Situação Marital
        especificacao += [
            ('Situação Marital', 'Estado Civil', [('label', 'Estado Civil')]),
            ('Situação Marital', 'Data do Casamento/União', [('label', 'Data do Casamento/União'), ('label', 'Data do Casamento')]),
            ('Situação Marital', 'Regime do Casamento', [('label', 'Regime do Casamento')]),
            ('Situação Marital', 'Regime do Casamento (se Outro)', [('label', 'Regime do Casamento (se Outro)'), ('label', 'Regime do Casamento (se outro)')]),
        ]

        # CAD
        especificacao += [
            ('CAD', 'NIS (PIS/PASEP)', [('label', 'NIS (PIS/PASEP)'), ('label', 'NIS')]),
            ('CAD', 'Nº da Carteira de Trabalho', [('label', 'Nº da Carteira de Trabalho'), ('label', 'Nº da Carteira')]),
            ('CAD', 'Nº de Série da Carteira de Trabalho', [('label', 'Nº de Série da Carteira de Trabalho'), ('label', 'Nº de Série')]),
            ('CAD', 'UF da Carteira de Trabalho', [('label', 'UF da Carteira de Trabalho'), ('label', 'UF da Carteira')]),
            ('CAD', 'Nº do Título de Eleitor', [('label', 'Nº do Título de Eleitor'), ('label', 'Título de Eleitor')]),
            ('CAD', 'Zona Eleitoral', [('label', 'Zona Eleitoral')]),
            ('CAD', 'Seção Eleitoral', [('label', 'Seção Eleitoral')]),
            ('CAD', 'Data de Emissão', [('label', 'Data de Emissão')]),
            ('CAD', 'UF Título de Eleitor', [('label', 'UF Título de Eleitor')]),
            ('CAD', 'Cidade do Título de Eleitor', [('label', 'Cidade do Título de Eleitor')]),
        ]

        # Contato
        especificacao += [
            ('Dados de Contato', 'Email', [('label', 'Email')]),
            ('Dados de Contato', 'Telefone', [('label', 'Telefone')]),
            ('Dados de Contato', 'Telefone Celular', [('label', 'Telefone Celular'), ('label', 'Celular')]),
            ('Dados de Contato', 'Telefone Comercial', [('label', 'Telefone Comercial'), ('label', 'Tel. Comercial')]),
            ('Dados de Contato', 'Telefone Para Recados', [('label', 'Telefone Para Recados'), ('label', 'Telefone para Recados'), ('label', 'Recado')]),
        ]

        # Parecer
        especificacao += [
            ('Parecer Social', 'Diagnóstico Social', [('label', 'Diagnóstico Social')]),
            ('Parecer Social', 'Observações', [('label', 'Observações')]),
        ]

        # Documentos (no modal, quando existir)
        for doc in [
            'CPF/RG Frente',
            'CPF/RG Verso',
            'CPF/RG (Cônjuge) Frente',
            'CPF/RG (Cônjuge) Verso',
            'Certidão de Nascimento',
            'Certidão de Casamento/União Estável',
            'Foto do Candidato',
            'Comprovante de Renda',
        ]:
            especificacao.append(('Documentos', doc, [('label', doc)]))

        variantes: dict[str, list[str]] = {}
        col_keys: list[str] = []
        for secao, campo, tentativas in especificacao:
            k = _col_key(aba, secao, campo)
            col_keys.append(k)
            labels = [txt for (modo, txt) in (tentativas or []) if modo == 'label']
            variantes[k] = labels or [campo]

        itens = self._extrair_tabela_com_modal_edicao_itens(variantes)
        return self._itens_para_colunas_multilinha(itens, col_keys)

    def extrair_composicao_familiar(self):
        if not self.clicar_aba('Composição Familiar'):
            return {}

        aba = 'Composição Familiar'

        especificacao: list[tuple[str, str, list[str]]] = []

        # Dados Gerais
        for campo in [
            'Tipo',
            'Nome',
            'CPF',
            'Tem CPF?',
            'Tipo de Pessoa',
            'RG',
            'Data de Emissão do RG',
            'Órgão Emissor do RG',
            'UF do RG',
            'Naturalidade',
            'UF Naturalidade',
            'Nacionalidade',
            'Data de Nascimento',
            'Nº da Certidão de Nascimento',
            'Raça',
            'Gênero',
            'Telefone',
            'Estado Civil',
            'Pessoa Negra?',
            'Pessoa Quilombola?',
            'Pessoa LGBTQIAPN+?',
            'Situação de Rua?',
        ]:
            especificacao.append(('Dados Gerais', campo, [campo]))

        # Educação e Profissão
        for campo in ['Escolaridade', 'Escola', 'Situação de Emprego', 'Profissão', 'Profissão (se Outro)', 'Tempo de Serviço']:
            especificacao.append(('Educação e Profissão', campo, [campo]))

        # Cadastro Preferencial
        for campo in ['Cadastro Preferencial', 'Deficiência', 'Doença']:
            especificacao.append(('Cadastro Preferencial', campo, [campo]))

        # CAD
        for campo in [
            'NIS (PIS/PASEP)',
            'Nº da Carteira de Trabalho',
            'Nº de Série da Carteira de Trabalho',
            'UF da Carteira de Trabalho',
            'Família Inscrita no CAD Único',
            'Programas CAD Único',
            'Nº do Título de Eleitor',
            'Zona Eleitoral',
            'Seção Eleitoral',
            'Data de Emissão',
            'UF Título de Eleitor',
            'Cidade do Título de Eleitor',
        ]:
            especificacao.append(('CAD', campo, [campo]))

        # Documentos
        for campo in [
            'CPF/RG Frente',
            'CPF/RG Verso',
            'CPF/RG (Cônjuge) Frente',
            'CPF/RG (Cônjuge) Verso',
            'Certidão de Nascimento',
            'Certidão de Casamento/União Estável',
            'Foto do Candidato',
            'Comprovante de Renda',
        ]:
            especificacao.append(('Documentos', campo, [campo]))

        variantes: dict[str, list[str]] = {}
        col_keys: list[str] = []
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
        campos = ['Tipo de Renda', 'Pessoa', 'Valor da Renda', 'Tipo se Apto Para Reurb']

        variantes: dict[str, list[str]] = {}
        col_keys: list[str] = []
        for campo in campos:
            k = _col_key(aba, secao, campo)
            col_keys.append(k)
            if campo == 'Tipo se Apto Para Reurb':
                variantes[k] = [campo, 'Apto Para Reurb']
            else:
                variantes[k] = [campo]

        itens = self._extrair_tabela_com_modal_edicao_itens(variantes)
        return self._itens_para_colunas_multilinha(itens, col_keys)

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

    def extrair_imovel(self):
        if not self.clicar_aba('Imóvel'):
            return {}

        aba = 'Imóvel'
        dados: dict[str, str] = {}

        def _get(*labels: str) -> str:
            for lb in labels:
                v = self._extrair_por_label_simples(lb)
                if v:
                    return v
            return ''

        # Aquisição e Infraestrutura
        secao = 'Aquisição e Infraestrutura'
        dados[_col_key(aba, secao, 'Lote')] = _get('Lote')
        dados[_col_key(aba, secao, 'Quadra')] = _get('Quadra')
        dados[_col_key(aba, secao, 'Área total (m²)')] = _get('Área total (m²)', 'Área total')
        dados[_col_key(aba, secao, 'Área construída (m²)')] = _get('Área construída (m²)', 'Área construída')
        dados[_col_key(aba, secao, 'Tipo de terreno')] = _get('Tipo de terreno')
        dados[_col_key(aba, secao, 'Tipo de imóvel')] = _get('Tipo de imóvel')
        dados[_col_key(aba, secao, 'Nº de cadastro imobiliário (IPTU)')] = _get('Nº de cadastro imobiliário (IPTU)', 'Nº de cadastro imobiliário')
        dados[_col_key(aba, secao, 'CEP')] = _get('CEP')
        dados[_col_key(aba, secao, 'UF')] = _get('UF')
        dados[_col_key(aba, secao, 'Cidade')] = _get('Cidade')
        dados[_col_key(aba, secao, 'Bairro')] = _get('Bairro')
        dados[_col_key(aba, secao, 'Endereço')] = _get('Endereço')
        dados[_col_key(aba, secao, 'Número')] = _get('Número')
        dados[_col_key(aba, secao, 'Complemento')] = _get('Complemento')

        # Infraestrutura do entorno (pode ser label ou lista de checkboxes)
        infra = _get('Infraestrutura do entorno', 'Infraestrutura Disponível')
        if not infra:
            try:
                checks = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']:checked")
                selecionados = []
                for c in checks:
                    if c.is_displayed():
                        try:
                            selecionados.append(self._texto_limpo(c.find_element(By.XPATH, "./following-sibling::label").text))
                        except:
                            pass
                infra = "\n".join(selecionados)
            except:
                infra = ''
        dados[_col_key(aba, secao, 'Infraestrutura do entorno')] = infra

        # Informações do Imóvel
        secao = 'Informações do Imóvel'
        for campo in [
            'Data do início do contrato',
            'Data do fim do contrato',
            'Data de aquisição',
            'Valor de aquisição',
            'Nome do edifício',
            'Nº de matrícula',
            'Nome do cartório',
        ]:
            dados[_col_key(aba, secao, campo)] = _get(campo)

        return dados

    def extrair_documentos_mapa(self) -> dict[str, str]:
        """Extrai um mapa simples doc_nome -> status (Anexado/Pendente/...) na aba Documentos."""
        if not self.clicar_aba('Documentos'):
            return {}
        try:
            out: dict[str, str] = {}
            tabelas = self.driver.find_elements(By.TAG_NAME, "table")
            for tabela in tabelas:
                try:
                    if not tabela.is_displayed():
                        continue
                    rows = tabela.find_elements(By.CSS_SELECTOR, "tbody tr")
                    for row in rows:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) < 2:
                            continue
                        nome = self._texto_limpo(cols[0].text)
                        if not nome:
                            continue
                        status = ""
                        try:
                            html = cols[1].get_attribute('innerHTML')
                            if "text-danger" in html:
                                status = "Pendente"
                            elif "text-success" in html or "text-primary" in html:
                                status = "Anexado"
                            else:
                                status = self._texto_limpo(cols[1].text)
                        except:
                            status = self._texto_limpo(cols[1].text)
                        out[nome] = status or out.get(nome, '')
                except:
                    continue
            return out
        except:
            return {}

    def extrair_documentos(self) -> str:
        """Compat: retorna um resumo em texto (útil para debug humano)."""
        mp = self.extrair_documentos_mapa() or {}
        linhas = []
        for k, v in mp.items():
            linhas.append(f"{k}: {v}")
        return "\n".join(linhas)

    def extrair_questionario_mapa(self) -> dict[str, str]:
        if not self.clicar_aba('Questionário'):
            return {}
        saida: dict[str, str] = {}
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, ".row, .question")
            for row in rows:
                try:
                    if not row.is_displayed():
                        continue

                    pergunta = ""
                    pergunta_el = None
                    try:
                        pergunta_el = row.find_element(By.TAG_NAME, "h6")
                        pergunta = self._texto_limpo(pergunta_el.text)
                    except:
                        try:
                            pergunta_el = row.find_element(By.TAG_NAME, "label")
                            pergunta = self._texto_limpo(pergunta_el.text)
                        except:
                            pass
                    if not pergunta:
                        continue

                    # Marca a pergunta como lida (quando possível)
                    try:
                        if pergunta_el:
                            self._highlight(pergunta_el, rotulo=f"PERGUNTA: {pergunta[:45]}")
                            self._mark_read(pergunta_el, rotulo=f"LIDO: {pergunta[:45]}")
                    except:
                        pass

                    resposta = ""
                    radios = row.find_elements(By.CSS_SELECTOR, "input[type='radio']:checked")
                    if radios:
                        try:
                            resposta = self._texto_limpo(radios[0].find_element(By.XPATH, "./following-sibling::label").text)
                        except:
                            resposta = "Sim"
                        try:
                            self._mark_read(radios[0], rotulo=f"LIDO: {pergunta[:45]}")
                            try:
                                lbl = radios[0].find_element(By.XPATH, "./following-sibling::label")
                                self._mark_read(lbl, rotulo=f"LIDO: {pergunta[:45]}")
                            except:
                                pass
                        except:
                            pass
                    else:
                        campos = row.find_elements(By.CSS_SELECTOR, "select, input:not([type='radio']), textarea")
                        for c in campos:
                            if c.is_displayed():
                                val = self._valor_campo(c)
                                if val:
                                    resposta = val
                                    try:
                                        self._mark_read(c, rotulo=f"LIDO: {pergunta[:45]}")
                                    except:
                                        pass
                                    break

                    if resposta:
                        saida[_norm_texto_chave(pergunta)] = resposta
                except:
                    continue
        except:
            pass
        return saida

    def extrair_questionario(self) -> str:
        mp = self.extrair_questionario_mapa() or {}
        linhas = []
        for k, v in mp.items():
            linhas.append(f"{k} -> {v}")
        return "\n".join(linhas)

    def extrair_historico(self):
        if not self.clicar_aba('Histórico'): return ""
        try:
            logs = []
            cards = self.driver.find_elements(By.CSS_SELECTOR, ".card-body")
            for card in cards:
                txt = self._texto_limpo(card.text)
                if txt and ("alterado por" in txt.lower() or "adicionado" in txt.lower()):
                    logs.append(txt)
            return "\n---\n".join(logs)
        except: return ""

    def extrair_status(self):
        if self.clicar_aba('Status'):
            try: return self._texto_limpo(self.driver.find_element(By.CSS_SELECTOR, ".container").text)
            except: return ""
        return ""


