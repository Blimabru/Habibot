# 🤖 Bot de Automação Habibot

> 🔄 Versão **v1.0.0**

Bot para extração automatizada de **TODOS** os dados de candidatos do sistema Habibot com exportação para Excel.

## ✨ Funcionalidades

### 📊 Extração Completa
- ✅ **Paginação automática** - Processa TODOS os candidatos (todas as páginas)
- ✅ **Excel em tempo real**: salva após cada candidato
- ✅ Solicita **login e senha** no início (não salva credenciais em arquivo)

### 🎯 Dados Extraídos

**Dados Principais:**
Nome, CPF, Telefone, Email, RG, Data de nascimento, Raça, Gênero, Estado civil, Cadastro preferencial, Deficiência, Doença, Observação, Diagnóstico social, Tipo de beneficiário, Situação de emprego, Informações CadÚnico, Loteamento

**Por Aba:**
- 🏠 **Endereço**: CEP, logradouro, número, complemento, bairro, cidade, estado
- 💰 **Renda**: Individual, familiar, per capita, fonte de renda
- 📋 **Questionário**: Todas as perguntas e respostas

**Grupo Familiar:**
- 👤 TITULAR
- 👥 COTITULAR
- 👨‍👩‍👧‍👦 COMPOSIÇÃO FAMILIAR

## 🚀 Instalação Rápida

### 1. Instalar Python
- Baixe em: https://www.python.org/downloads/
- ⚠️ **IMPORTANTE**: Marque "Add Python to PATH"

### 2. Instalar Dependências
```bash
pip install -r requirements.txt
```

### 3. Credenciais
O bot solicita **usuário e senha** na execução.

## ▶️ Como Usar


### 📁 Estrutura do Projeto
```
Habibot/
├── bot_habibot.py           # Script principal do bot
├── build_exe.ps1            # Script de build do executável (Windows)
├── requirements.txt         # Dependências Python
├── README.md                # Documentação
├── .gitignore               # Arquivos ignorados pelo Git
│
├───.github/
│   └───workflows/
│           release.yml      # CI/CD
│
├───assets/
│   ├───build/
│   │       version_info.txt # Metadados do executável
│   ├───drivers/
│   │       chromedriver.exe # ChromeDriver incluso para modo offline
│   └───images/
│       └───icons/
│               icon.png     # Ícone base (PNG)
│
├───dist/
│       Habibot_<versão>.exe      # Executável gerado pelo build
│
└───tools/
	make_windows_ico.py  # Script para gerar ícones Windows
```


### Opção 1: Executável (Recomendado)
1. Execute: `dist/Habibot_Dev.exe` (ou o nome gerado pelo build)
2. Aguarde até aparecer a mensagem de conclusão
3. O arquivo Excel será gerado na pasta `Extração de Dados Habibot/`

### Opção 2: Linha de Comando (Python)
```bash
python bot_habibot.py
```

## 📁 Arquivo Excel Gerado

Pasta: `Extração de Dados Habibot/`

Nome: `Candidatos_Habibot_AAAA-MM-DD_HH-MM-SS.xlsx`

- **1 planilha**: `Dados`
- **1 candidato por linha**
- Campos vazios são preenchidos com `Não Informado`
- `Data de extração` é a última coluna

Durante a execução existe um arquivo temporário de visualização:
- `VISUALIZACAO-Candidatos_Habibot_AAAA-MM-DD_HH-MM-SS.xlsx`

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

O build já embute um `chromedriver.exe` em `assets/drivers/chromedriver.exe` dentro do EXE, para rodar em um PC **sem internet**.

Observações:
- O Chrome precisa estar instalado na máquina de destino.
- Se o Chrome do outro PC estiver em uma versão muito diferente, pode dar incompatibilidade de driver.
- O ícone do executável é gerado automaticamente a partir de `assets/images/icons/icon.png` pelo script `tools/make_windows_ico.py`.

## 🔄 Paginação Automática

O bot detecta automaticamente quando há mais páginas e processa **TODOS** os candidatos disponíveis no sistema, não apenas os primeiros 30.

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