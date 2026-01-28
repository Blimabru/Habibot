# 🤖 Bot de Automação Habibot

> 🔄 Versão **v1.1.0**

Bot para extração automatizada de **TODOS** os dados de candidatos do sistema Habibot com exportação para Excel.

## ✨ Funcionalidades

### 📊 Extração Completa
- ✅ Paginação automática (todas as páginas)
- ✅ Excel em tempo real (salva após cada candidato)
- ✅ Busca por lista de CPFs
- ✅ Exporta todos os campos do sistema, organizados por aba/seção
- ✅ Logs detalhados de ações e ordem de leitura (TSV)
- ✅ Visualização em tempo real em arquivo TXT
- ✅ Suporte a modo headless (sem abrir navegador)
- ✅ Destaca elementos no navegador (debug visual)
- ✅ Instalação fácil via requirements.txt
- ✅ Permite inserir logo personalizado no cabeçalho da planilha

### 🎯 Dados Extraídos

**Principais:** Nome, CPF, Telefone, Email, RG, Data de nascimento, Raça, Gênero, Estado civil, Cadastro preferencial, Deficiência, Doença, Observação, Diagnóstico social, Tipo de beneficiário, Situação de emprego, Informações CadÚnico, Loteamento

**Por Aba:**
- 🏠 Endereço: CEP, logradouro, número, complemento, bairro, cidade, estado
- 💰 Renda: Individual, familiar, per capita, fonte de renda
- 📋 Questionário: Todas as perguntas e respostas

**Grupo Familiar:**
- 👤 Titular
- 👥 Cotitular
- 👨‍👩‍👧‍👦 Composição Familiar

## 🚀 Instalação Rápida

### 1. Instalar Python
- Baixe em: https://www.python.org/downloads/
- ⚠️ Marque "Add Python to PATH"

### 2. Instalar Dependências
```bash
pip install -r requirements.txt
```

### 3. Credenciais
O bot solicita usuário e senha na execução.

## ▶️ Como Usar

### Executando pelo Python
```bash
python habibot.py
```

### Executando com opções avançadas
- Extrair todos os candidatos:
  ```bash
  python habibot.py --todos
  ```
Extrair por lista de CPFs:
  ```bash
  python habibot.py --cpfs-file lista_cpfs.txt
  ```

  > 📄 **Como usar o arquivo de CPFs**
  >
  > 1. **Crie um arquivo texto** (ex: `lista_cpfs.txt`) com um CPF por linha.
  > 2. **Salve preferencialmente na raiz do projeto** (mesma pasta do `habibot.py`).
  > 3. **Exemplo de caminho:**
  >    - `C:\Users\User\Desktop\Habibot\lista_cpfs.txt`
  > 4. **Você pode usar outro nome ou pasta**, basta informar corretamente no parâmetro `--cpfs-file`.
  >
  > **Como informar o parâmetro:**
  >
  > - Caminho relativo (arquivo na mesma pasta do habibot.py):
  >   ```bash
  >   python habibot.py --cpfs-file lista_cpfs.txt
  >   ```
  > - Caminho absoluto (arquivo em outro local):
  >   ```bash
  >   python habibot.py --cpfs-file "C:\Users\User\Desktop\Habibot\lista_cpfs.txt"
  >   ```
  > - Caminho para subpasta:
  >   ```bash
  >   python habibot.py --cpfs-file assets/listas/cpfs.txt
  >   ```
  >
  > **Formato do arquivo:**
  > ```
  > 123.456.789-01
  > 987.654.321-00
  > ...
  > ```
- Executar sem abrir navegador:
  ```bash
  python habibot.py --headless
  ```
- Limitar quantidade de candidatos:
  ```bash
  python habibot.py --max-candidatos 10
  ```
- Ativar debug visual:
  ```bash
  python habibot.py --debug-visual
  ```

### Parâmetros disponíveis
- `--todos` : Extrai todos os candidatos
- `--cpfs-file arquivo.txt` : Extrai apenas CPFs informados
- `--headless` : Executa sem abrir navegador
- `--headed` : Executa com navegador visível
- `--cooldown N` : Espera N segundos entre extrações
- `--max-candidatos N` : Limita o total extraído
- `--debug` : Ativa logs detalhados
- `--debug-visual` : Destaca elementos no navegador
- `--non-interactive` : Não faz perguntas, usa apenas flags/variáveis
- `--prompt-credenciais` : Permite digitar usuário/senha mesmo com --non-interactive
- `--no-pause` : Não espera ENTER ao final

## 📁 Estrutura do Projeto
```
Habibot/
├── habibot.py                # Script principal do bot (ponto de entrada)
├── build_exe.ps1             # Script de build do executável (Windows)
├── requirements.txt          # Dependências Python
├── README.md                 # Documentação
│
├───src/                      # Código fonte modularizado
│   ├── __init__.py           # Inicialização do pacote
│   ├── bot.py                # Classe HabibotBot (automação principal)
│   ├── config.py             # Configurações e utilitários
│   ├── dependencies.py       # Gerenciamento de dependências
│   ├── excel_handler.py      # Manipulação de arquivos Excel
│   ├── extractor.py          # Classe ExtratorHabibot (extração de dados)
│   ├── loggers.py            # Classes de logging (OrdemLeituraLogger, AcoesLogger)
│   └── schema.py             # Schema e estrutura dos dados
│
├───assets/
│   ├───build/
│   │       version_info.txt  # Metadados do executável
│   ├───drivers/
│   │       chromedriver.exe  # ChromeDriver incluso para modo offline
│   └───images/
│       └───icons/
│       └───logos/
│               logo.png # Logo padrão para cabeçalho da planilha
│               (adicione sua logo personalizada aqui)
│
├───debug/                    # Logs de execução (TSV)
│
├───Extração de Dados Habibot/
│       VISUALIZAÇÃO-...txt   # Visualização em tempo real
│
├───Sistema/
│       Dados que precisam ser extraídos.txt
│       Capturas de tela/
│
└───tools/
        make_windows_ico.py   # Script para gerar ícones Windows
```

## 🖼️ Como adicionar uma imagem de logo no cabeçalho da planilha

1. Prepare sua imagem no formato PNG (recomendado até 180x70px).
2. Renomeie para `logo.png` ou outro nome desejado.
3. Coloque o arquivo em `assets/images/logos/`.
   - Exemplo: `assets/images/logos/minha-logo.png`
4. O bot insere automaticamente a imagem encontrada em `assets/images/logos/logo.png` no cabeçalho da planilha gerada.
   - Para usar outro nome, renomeie sua imagem para `logo.png` ou altere o nome no código (classe `ExcelTempoReal`).
5. Se não houver imagem, o cabeçalho será gerado apenas com texto.

## 📁 Arquivo Excel Gerado

Pasta: `Habibot - Dados Extraídos/`

Nome: `Candidatos_<sistema>_AAAA-MM-DD_HH-MM-SS.xlsx`

- 1 planilha: `Dados`
- 1 candidato por linha
- Campos vazios são preenchidos com `Não Informado`
- Data de extração é a última coluna
- Cabeçalho com logo (se existir em `assets/images/logos/logo.png`)
- Cabeçalhos mesclados por aba/seção/campo
- Visualização em tempo real: arquivo TXT

Durante a execução existe um arquivo temporário de visualização:
- `VISUALIZACAO_<sistema>_AAAA-MM-DD_HH-MM-SS.txt`

Obs (Windows/Excel): se o arquivo de visualização estiver aberto no Excel, o Windows pode bloquear escrita.
Quando isso acontecer, o bot cria uma nova cópia numerada.
Ao final (ou Ctrl+C), os arquivos de visualização são apagados e fica apenas o arquivo final.

## 🧱 Gerar Executável (Windows)

Gera um .exe com Python e dependências embutidos via PyInstaller:

```powershell
./build_exe.ps1
```

Saída:
- `dist/Habibot_Dev.exe` (ou similar, dependendo da versão/tag)

### 🌐/🔌 ChromeDriver incluso

O build já embute um `chromedriver.exe` em `assets/drivers/chromedriver.exe` dentro do EXE, para rodar em um PC sem internet.

Observações:
- O Chrome precisa estar instalado na máquina de destino.
- Se o Chrome do outro PC estiver em uma versão muito diferente, pode dar incompatibilidade de driver.
- O ícone do executável é gerado automaticamente a partir de `assets/images/icons/icon.png` pelo script `tools/make_windows_ico.py`.

## 🔄 Paginação Automática

O bot detecta automaticamente quando há mais páginas e processa todos os candidatos disponíveis no sistema, não apenas os primeiros 30.

Progresso exibido:
```
📄 PROCESSANDO PÁGINA 1
📊 Candidatos nesta página: 30
🔄 Navegando para próxima página...

📄 PROCESSANDO PÁGINA 2
📊 Candidatos nesta página: 30
...

📊 TOTAL FINAL: 127 candidatos processados
📄 Total de páginas: 5
```

## 📊 Estrutura dos Dados

### Candidato Principal
```
tipo_registro: "CANDIDATO"
id_candidato: 1
nome, cpf, telefone, email, rg
data_nascimento, raca, genero
endereco_*, renda_*, questionario_*
data_extracao: timestamp
```

### Membros da Família
```
tipo_registro: "TITULAR" | "COTITULAR" | "COMPOSIÇÃO FAMILIAR"
id_candidato: referência ao candidato
nome_membro, cpf_membro, parentesco
renda_membro
data_extracao: timestamp
```

## 🐛 Solução de Problemas

### "Python não reconhecido"
- Reinstale Python marcando "Add to PATH"
- Reinicie o computador

### "ChromeDriver" / compatibilidade
- Se aparecer erro de incompatibilidade entre Chrome e ChromeDriver, atualize o Chrome ou gere novamente o EXE na máquina com internet (o build baixa um driver atualizado).

## 📝 Logs

Durante execução, o bot exibe:
- 🚀 Início de operações
- ✅ Sucesso
- ⚠️ Avisos
- ❌ Erros
- 📊 Progresso (candidatos/páginas)
- 👁️ Logs detalhados em TSV na pasta `debug/` (quando rodando como .py)

## 📚 Dependências

Veja `requirements.txt` para detalhes e instruções de instalação.

---