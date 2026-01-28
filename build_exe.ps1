param(
  [string]$Entry = "habibot.py",
  [string]$Version = "" # Opção para forçar versão manual: .\build_exe.ps1 -Version "v2.0"
)

$ErrorActionPreference = "Stop"

# --- Lógica de Versão Dinâmica ---
if ($Version) {
    # 1. Usa a versão informada manualmente
    $VersionName = $Version
}
else {
    # 2. Verifica a versão automaticamente
    try {
        # Tentativa A: Método padrão (exige que a tag esteja no histórico do commit atual)
        $RawTag = git describe --tags --abbrev=0 2>$null
        
        # Tentativa B: Ppega a última tag criada, independente do commit
        if ([string]::IsNullOrWhiteSpace($RawTag)) {
             # --sort=-creatordate ordena do mais novo para o mais velho
             $RawTag = $(git tag --sort=-creatordate | Select-Object -First 1)
        }

        # Limpeza e validação
        if (-not [string]::IsNullOrWhiteSpace($RawTag)) {
            $VersionName = $RawTag.Trim()
        } else {
            $VersionName = "dev"
        }
    } catch {
        $VersionName = "dev"
    }
}

# Define o nome final do arquivo
$ExeName = "Habibot_$VersionName"
# --------------------------------

Write-Host "== Build EXE ($ExeName) ==" -ForegroundColor Cyan

# Garante que pyinstaller exista
python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller
python -m pip install -r requirements.txt
python -m pip install --upgrade pillow

# Prepara ChromeDriver local (para execução offline)
if (!(Test-Path .\assets\drivers)) { New-Item -ItemType Directory -Path .\assets\drivers | Out-Null }
if (!(Test-Path .\assets\drivers\chromedriver.exe)) {
  Write-Host "Baixando ChromeDriver para .\\assets\\drivers\\chromedriver.exe ..." -ForegroundColor Yellow
  python -c "from webdriver_manager.chrome import ChromeDriverManager; import shutil; from pathlib import Path; p=Path(ChromeDriverManager().install()); t=Path('assets')/'drivers'/'chromedriver.exe'; t.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(p, t); print('ChromeDriver:', t)"
}

# Gera ícone do executável (ICO multi-tamanho) a partir de assets\images\icons\icon.png
if (!(Test-Path .\assets\images\icons)) { New-Item -ItemType Directory -Path .\assets\images\icons | Out-Null }
if (!(Test-Path .\assets\images\icons\icon.png)) {
  throw "Ícone fonte não encontrado em .\\assets\\images\\icons\\icon.png"
}
python .\tools\make_windows_ico.py

# Limpa builds antigos
if (Test-Path .\build) { Remove-Item -Recurse -Force .\build }
$DistDir = ".\\dist"
if (Test-Path $DistDir)  {
  try {
    Remove-Item -Recurse -Force $DistDir -ErrorAction Stop
  } catch {
    # PowerShell 5.1 pode quebrar acentos sem BOM; manter mensagens em ASCII
    Write-Host "! Nao foi possivel limpar a pasta .\\dist (provavelmente o EXE esta aberto/em execucao)." -ForegroundColor Yellow
    Write-Host "  Farei o build em .\\dist_build. Feche o Habibot.exe e rode novamente se quiser sobrescrever .\\dist." -ForegroundColor Yellow

    $DistDir = ".\\dist_build"
    if (Test-Path $DistDir) { Remove-Item -Recurse -Force $DistDir }
  }
}
if (Test-Path .\*.spec){ Remove-Item -Force .\*.spec }

# --onefile inclui o Python e dependências no executável
# --add-binary embute o chromedriver.exe para execução offline
# --icon define o ícone do executável (ICO multi-tamanho)
# --version-file define metadados do arquivo no Windows
# --hidden-import garante que os módulos do pacote src sejam incluídos
python -m PyInstaller --clean --onefile --name $ExeName --distpath $DistDir `
  --icon "assets\images\icons\Habibot.ico" `
  --version-file "assets\build\version_info.txt" `
  --add-binary "assets\drivers\chromedriver.exe;assets\drivers" `
  --hidden-import src.bot `
  --hidden-import src.config `
  --hidden-import src.dependencies `
  --hidden-import src.excel_handler `
  --hidden-import src.extractor `
  --hidden-import src.loggers `
  --hidden-import src.schema `
  $Entry

Write-Host "\nEXE gerado em: $DistDir\\$ExeName.exe" -ForegroundColor Green
Write-Host "Obs: o Chrome precisa estar instalado na maquina." -ForegroundColor Yellow

# Limpa artefatos não necessários para distribuir o EXE
if (Test-Path .\build) { Remove-Item -Recurse -Force .\build }
if (Test-Path .\*.spec){ Remove-Item -Force .\*.spec }
