"""Habibot - Bot de Automação para Sistema de Habitação.

Este bot realiza a extração automatizada dos candidatos cadastrados no
Sistema de Habitação, salvando os dados em um relatório Excel.

Versão modularizada para melhor organização e manutenção do código.
"""

import argparse
import os

from src.config import mostrar_boas_vindas
from src.dependencies import get_missing_dependencies
from src.bot import HabibotBot


def main():
    """Função principal do bot."""
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

        # Verifica dependências
        missing = get_missing_dependencies()
        if missing:
            print(f"\n❌ Dependências faltando: {', '.join(missing)}")
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
