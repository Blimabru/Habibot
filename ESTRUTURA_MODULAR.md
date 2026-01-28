# Estrutura Modular do Código - Habibot

Este documento explica a nova organização modular do código do Habibot.

## 📋 Visão Geral

O código foi reorganizado de um único arquivo monolítico (`habibot.py` com 3822 linhas) para uma estrutura modular bem organizada com o diretório `src/` contendo módulos especializados.

## 📂 Estrutura de Diretórios

```
Habibot/
├── habibot.py                 # Ponto de entrada principal (117 linhas)
├── src/                       # Código fonte modularizado
│   ├── __init__.py           # Inicialização do pacote
│   ├── bot.py                # Automação principal (1100 linhas)
│   ├── config.py             # Configurações e utilidades (220 linhas)
│   ├── dependencies.py       # Gerenciamento de dependências (64 linhas)
│   ├── excel_handler.py      # Manipulação de Excel (357 linhas)
│   ├── extractor.py          # Extração de dados (1467 linhas)
│   ├── loggers.py            # Sistema de logging (156 linhas)
│   └── schema.py             # Estrutura de dados (482 linhas)
├── assets/                   # Recursos (drivers, imagens)
├── tools/                    # Scripts auxiliares
└── requirements.txt          # Dependências Python
```

## 🔧 Módulos

### 1. `habibot.py` - Ponto de Entrada
- **Responsabilidade**: Ponto de entrada principal do aplicativo
- **Linhas**: ~117 (reduzido de 3822!)
- **Funções**:
  - Processar argumentos da linha de comando
  - Inicializar e executar o bot
  - Tratamento de erros de alto nível

### 2. `src/config.py` - Configuração
- **Responsabilidade**: Configurações, constantes e funções utilitárias
- **Componentes**:
  - Informações de versão (`__version__`, `__author__`, etc.)
  - Funções utilitárias (`_diretorio_base_execucao()`, `_terminal_width()`, etc.)
  - Função de boas-vindas (`mostrar_boas_vindas()`)
  - Normalização de texto (`_norm_texto_chave()`)
  - Detecção de ChromeDriver local

### 3. `src/dependencies.py` - Dependências
- **Responsabilidade**: Gerenciamento de importações opcionais
- **Funções**:
  - Importação condicional de Selenium
  - Importação condicional de WebDriver Manager
  - Importação condicional de OpenPyXL
  - Rastreamento de dependências faltantes

### 4. `src/loggers.py` - Sistema de Logging
- **Responsabilidade**: Classes de logging especializadas
- **Classes**:
  - `OrdemLeituraLogger`: Registra ordem de leitura de campos (formato TSV)
  - `AcoesLogger`: Registra ações do bot/extrator (formato TSV)

### 5. `src/schema.py` - Estrutura de Dados
- **Responsabilidade**: Define a estrutura dos dados extraídos
- **Componentes**:
  - `ColSpec`: Especificação de colunas (dataclass)
  - `SCHEMA_ESTRUTURA`: Schema completo do Excel
  - `COL_SPECS`: Lista flat de especificações
  - Funções auxiliares de mapeamento

### 6. `src/excel_handler.py` - Manipulação de Excel
- **Responsabilidade**: Geração e formatação de arquivos Excel
- **Classe Principal**:
  - `ExcelTempoReal`: Grava Excel formatado em tempo real
- **Funcionalidades**:
  - Criação de arquivo com cabeçalhos mesclados
  - Inserção de logo personalizado
  - Formatação de células
  - Arquivo de visualização em TXT

### 7. `src/extractor.py` - Extração de Dados
- **Responsabilidade**: Extração de dados do sistema web
- **Classe Principal**:
  - `ExtratorHabibot`: Navega e extrai dados dos candidatos
- **Funcionalidades**:
  - Navegação por abas do sistema
  - Extração de campos específicos
  - Tratamento de diferentes tipos de input
  - Debug visual (destaque de elementos)

### 8. `src/bot.py` - Automação Principal
- **Responsabilidade**: Coordenação geral da automação
- **Classe Principal**:
  - `HabibotBot`: Orquestra todo o processo de automação
- **Funcionalidades**:
  - Configuração do WebDriver
  - Login no sistema
  - Listagem de candidatos
  - Coordenação de extração
  - Paginação automática
  - Tratamento de CPFs

## 🎯 Benefícios da Modularização

### Manutenibilidade
- ✅ Código organizado por responsabilidade
- ✅ Fácil localização de funcionalidades
- ✅ Redução de acoplamento entre componentes

### Legibilidade
- ✅ Arquivos menores e mais focados
- ✅ Imports claros e explícitos
- ✅ Melhor documentação por módulo

### Testabilidade
- ✅ Módulos independentes podem ser testados isoladamente
- ✅ Facilitação de mocks e stubs
- ✅ Dependências claramente definidas

### Reutilização
- ✅ Componentes podem ser reutilizados em outros projetos
- ✅ Loggers e utilitários são independentes
- ✅ Schema pode ser versionado separadamente

## 🔄 Como Usar

### Execução Normal
```bash
python habibot.py
```

### Com Opções
```bash
python habibot.py --todos --headless
python habibot.py --cpfs-file lista.txt --debug
```

### Importando Módulos
```python
from src.config import __version__
from src.loggers import OrdemLeituraLogger
from src.bot import HabibotBot

print(f"Versão: {__version__}")
bot = HabibotBot()
```

## 🏗️ Build do Executável

O script `build_exe.ps1` foi atualizado para incluir todos os módulos do diretório `src/` no executável final:

```powershell
.\build_exe.ps1
```

O PyInstaller agora usa `--hidden-import` para garantir que todos os módulos sejam incluídos.

## 📝 Compatibilidade

- ✅ Mantém 100% de compatibilidade com a versão anterior
- ✅ Todos os argumentos de linha de comando funcionam igual
- ✅ Mesma saída de arquivos e logs
- ✅ Comportamento idêntico

## 🎓 Próximos Passos Recomendados

1. **Testes Unitários**: Criar testes para cada módulo
2. **Type Hints**: Adicionar type hints completos em todos os módulos
3. **Documentação**: Expandir docstrings com exemplos
4. **CI/CD**: Configurar testes automatizados
5. **Logging Aprimorado**: Usar biblioteca de logging padrão do Python

## 📚 Referências

- Estrutura baseada em boas práticas Python
- Organização modular seguindo princípios SOLID
- Separação de responsabilidades (Separation of Concerns)
