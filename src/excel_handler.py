"""Módulo de manipulação de arquivos Excel.

Contém a classe ExcelTempoReal para geração de relatórios.
"""

import os
import time
from pathlib import Path

from .dependencies import (
    Workbook, load_workbook, Alignment, PatternFill, Font, Border, Side,
    OpenpyxlImage, get_column_letter
)
from .schema import COL_SPECS

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
            return 'Não informado'
        if isinstance(valor, str):
            txt = valor.strip()
            if not txt:
                return 'Não informado'
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
                return 'Não informado'
            return txt
        txt = str(valor).strip()
        return txt if txt else 'Não informado'

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


