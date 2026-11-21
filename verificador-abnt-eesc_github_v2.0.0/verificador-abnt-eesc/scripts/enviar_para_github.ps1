# Envia o projeto para o GitHub (primeira vez e atualizacoes).
#
# Primeira vez: instala Git e GitHub CLI se faltarem (winget), faz login no GitHub pelo navegador,
# cria o repositorio (PRIVADO por padrao) e envia tudo.
# Depois: so registra as mudancas e envia.
#
# Uso (pela pasta do projeto ou com dois cliques em scripts\enviar_para_github.bat):
#   enviar_para_github.bat                                 cria/atualiza "verificador-abnt-eesc" (privado)
#   enviar_para_github.bat -Mensagem "Nova regra de siglas"
#   enviar_para_github.bat -Versao                         tambem publica a versao do arquivo VERSION
#                                                          (o GitHub gera o .zip de instalacao em Releases)
#   enviar_para_github.bat -Dono biblioteca-eesc           cria dentro de uma organizacao do GitHub
#   enviar_para_github.bat -Publico                        repositorio publico (pense antes: ver docs)
param(
    [string]$Repo = "verificador-abnt-eesc",
    [string]$Dono = "",
    [string]$Mensagem = "",
    [switch]$Publico,
    [switch]$Versao
)

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"
$Raiz = Split-Path -Parent $PSScriptRoot
Set-Location $Raiz

function Ok([string]$m)    { Write-Host "  [OK] $m" -ForegroundColor Green }
function Faz([string]$m)   { Write-Host "  [..] $m" -ForegroundColor Cyan }
function Aviso([string]$m) { Write-Host "  [!!] $m" -ForegroundColor Yellow }

function Atualizar-Path {
    $m = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $u = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$m;$u"
}

function Garantir([string]$exe, [string]$idWinget, [string]$nome, [string]$link) {
    if (Get-Command $exe -ErrorAction SilentlyContinue) { Ok $nome; return }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "$nome nao encontrado e o winget nao esta disponivel. Instale manualmente: $link"
    }
    Faz "Instalando $nome (pode pedir confirmacao do Windows)..."
    & winget install --id $idWinget -e --source winget --accept-source-agreements --accept-package-agreements
    Atualizar-Path
    foreach ($p in @("$env:ProgramFiles\Git\cmd", "$env:ProgramFiles\GitHub CLI", "$env:LOCALAPPDATA\Programs\GitHub CLI")) {
        if (Test-Path $p) { $env:Path = "$env:Path;$p" }
    }
    if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) {
        throw "$nome foi instalado, mas o Windows ainda nao o encontrou. Feche esta janela e rode de novo."
    }
    Ok "$nome instalado"
}

try {
    Write-Host ""
    Write-Host "=== Enviar o projeto para o GitHub ===" -ForegroundColor White

    Garantir "git" "Git.Git" "Git" "https://git-scm.com/download/win"
    Garantir "gh" "GitHub.cli" "GitHub CLI" "https://cli.github.com"

    # ---- login
    & gh auth status *> $null
    if ($LASTEXITCODE -ne 0) {
        Faz "Entre na sua conta do GitHub na janela do navegador (copie o codigo que aparecer aqui)..."
        & gh auth login --hostname github.com --git-protocol https --web
        if ($LASTEXITCODE -ne 0) { throw "Login no GitHub nao concluido." }
    }
    & gh auth setup-git *> $null
    $login = (& gh api user --jq .login).Trim()
    Ok "Conectado ao GitHub como $login"

    # ---- identidade dos commits
    if (-not (& git config user.name)) {
        $nomeGh = (& gh api user --jq ".name // .login").Trim()
        & git config --global user.name $nomeGh
    }
    if (-not (& git config user.email)) {
        $id = (& gh api user --jq .id).Trim()
        & git config --global user.email "$id+$login@users.noreply.github.com"
    }

    # ---- repositorio local
    if (-not (Test-Path (Join-Path $Raiz ".git"))) {
        Faz "Criando o repositorio local..."
        & git init -b main | Out-Null
        Ok "Repositorio local criado"
    }

    # ---- seguranca: nada de chaves nem trabalhos de alunos
    & git add -A
    $proibidos = & git diff --cached --name-only | Where-Object {
        $_ -match "config\.json$|publico\.json$|\.log$" -or
        ($_ -match "\.docx$" -and $_ -notmatch "^verificador-interno/(exemplos|assets/modelos)/")
    }
    if ($proibidos) {
        & git reset -q
        throw "Arquivos que nao podem ir para o GitHub (chaves ou trabalhos de alunos): $($proibidos -join ', ')"
    }
    $grandes = & git diff --cached --name-only | Where-Object { (Test-Path $_) -and (Get-Item $_).Length -gt 50MB }
    if ($grandes) { & git reset -q; throw "Arquivos maiores que 50 MB: $($grandes -join ', ')" }

    # ---- commit
    & git diff --cached --quiet
    if ($LASTEXITCODE -ne 0) {
        if (-not $Mensagem) { $Mensagem = "Atualizacao de " + (Get-Date -Format "dd/MM/yyyy HH:mm") }
        & git commit -q -m $Mensagem
        Ok "Mudancas registradas: $Mensagem"
    } else {
        Ok "Nenhuma mudanca nova para registrar"
    }

    # ---- remoto
    $nomeCompleto = if ($Dono) { "$Dono/$Repo" } else { "$login/$Repo" }
    $origem = (& git remote get-url origin 2>$null)
    if (-not $origem) {
        & gh repo view $nomeCompleto *> $null
        if ($LASTEXITCODE -eq 0) {
            Faz "O repositorio $nomeCompleto ja existe no GitHub; ligando a ele..."
            & git remote add origin "https://github.com/$nomeCompleto.git"
        } else {
            $vis = if ($Publico) { "--public" } else { "--private" }
            Faz "Criando $nomeCompleto no GitHub ($($vis.TrimStart('-')))..."
            & gh repo create $nomeCompleto $vis --source . --remote origin `
                --description "Normalizacao ABNT da Biblioteca EESC-USP: verificador de TCCs, dissertacoes e teses e assistente para alunos"
            if ($LASTEXITCODE -ne 0) { throw "Nao foi possivel criar o repositorio." }
            Ok "Repositorio criado"
        }
    }

    Faz "Enviando..."
    & git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        throw "O envio falhou. Se alguem mudou o repositorio pelo site, rode 'git pull --rebase' e tente de novo."
    }
    Ok "Projeto enviado"

    # ---- versao (gera o .zip de instalacao em Releases)
    if ($Versao) {
        $v = (Get-Content (Join-Path $Raiz "verificador-interno\VERSION") -Raw).Trim()
        $tag = "v$v"
        & git ls-remote --exit-code --tags origin $tag *> $null
        if ($LASTEXITCODE -eq 0) {
            Aviso "A versao $tag ja foi publicada. Aumente o numero em verificador-interno\VERSION e rode de novo."
        } else {
            & git tag -a $tag -m "Versao $v"
            & git push origin $tag
            Ok "Versao $tag enviada: o GitHub vai testar e publicar o .zip em Releases em alguns minutos"
        }
    }

    $url = (& gh repo view $nomeCompleto --json url --jq .url).Trim()
    Write-Host ""
    Write-Host "Pronto: $url" -ForegroundColor Green
    Start-Process $url
    exit 0
} catch {
    Write-Host ""
    Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
