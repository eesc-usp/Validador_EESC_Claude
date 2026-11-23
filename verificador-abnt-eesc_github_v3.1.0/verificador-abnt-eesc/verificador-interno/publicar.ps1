# Verificador ABNT - Biblioteca EESC-USP
# Publica o Verificador na internet a partir deste PC: um site com link para os alunos usarem.
#
# Como funciona: roda uma 2a copia do app so para a internet (porta 8502, fechada para a rede local),
# em MODO PUBLICO (sem configuracoes, com limites de uso e aviso de privacidade), e liga um tunel que
# da um endereco https publico sem abrir portas no roteador nem no firewall.
#
# Tuneis:
#   tailscale  (recomendado) link fixo e gratuito, ex.: https://biblioteca.tail1234.ts.net
#              precisa instalar o Tailscale uma vez (pede senha de administrador) e entrar com uma conta.
#   cloudflare link fixo no SEU dominio (ex.: verificador.seudominio.org); precisa de um token de tunel
#              criado no painel da Cloudflare (Zero Trust > Networks > Tunnels) e do -Endereco.
#   rapido     sem conta e sem instalar nada; o link (*.trycloudflare.com) MUDA a cada reinicio. Bom para testar.
#
# Uso:  publicar.bat                         (pergunta qual tunel)
#       publicar.ps1 -Tunel tailscale
#       publicar.ps1 -Tunel cloudflare -Token <token> -Endereco https://verificador.seudominio.org
#       publicar.ps1 -Tunel rapido
#       publicar.ps1 -Desativar              (tira o site do ar)
param(
    [ValidateSet("", "tailscale", "cloudflare", "rapido")][string]$Tunel = "",
    [string]$Token = "",
    [string]$Endereco = "",
    [switch]$Desativar,
    [switch]$Reiniciar,
    [switch]$Oculto
)

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Pasta     = $PSScriptRoot
$LogArq    = Join-Path $Pasta "publicar.log"
$CfgArq    = Join-Path $Pasta "publico.json"
$ConfigApp = Join-Path $Pasta "config.json"
$Venv      = Join-Path $Pasta ".venv"
$VPy       = Join-Path $Venv "Scripts\python.exe"
$Porta     = 8502
$Local     = "http://127.0.0.1:$Porta"
$Ferr      = Join-Path $Pasta "ferramentas"
$Cf        = Join-Path $Ferr "cloudflared.exe"
$CfUrl     = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
$LinkArq   = Join-Path $Pasta "compartilhar\link_publico.txt"
$Inicio    = Join-Path ([Environment]::GetFolderPath("Startup")) "Verificador ABNT - Site publico.lnk"

function Log([string]$msg, [string]$cor = "Gray") {
    if (-not $Oculto) { Write-Host $msg -ForegroundColor $cor }
    Add-Content -Path $LogArq -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg)
}
function Ok([string]$m)    { Log "  [OK] $m" "Green" }
function Faz([string]$m)   { Log "  [..] $m" "Cyan" }
function Aviso([string]$m) { Log "  [!!] $m" "Yellow" }

function Ler-Json([string]$arq) {
    if (Test-Path $arq) { try { return Get-Content $arq -Raw | ConvertFrom-Json } catch { } }
    return $null
}

function Salvar-Cfg($cfg) {
    ($cfg | ConvertTo-Json) | Set-Content $CfgArq -Encoding UTF8
}

function Site-Ativo {
    try { Invoke-WebRequest "$Local/_stcore/health" -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop | Out-Null; return $true }
    catch { return $false }
}

# ------------------------------------------------------------------ copia publica do app
function Parar-Publico {
    $pidArq = Join-Path $Pasta "publico.pid"
    if (Test-Path $pidArq) {
        New-Item -ItemType File -Force -Path (Join-Path $Pasta "publico.parar") | Out-Null
        for ($i = 0; $i -lt 15; $i++) { Start-Sleep 1; if (-not (Test-Path $pidArq)) { break } }
        if (Test-Path $pidArq) {
            $id = (Get-Content $pidArq -Raw).Trim()
            & taskkill /PID $id /T /F *> $null
            Remove-Item $pidArq -ErrorAction SilentlyContinue
        }
        Remove-Item (Join-Path $Pasta "publico.parar") -ErrorAction SilentlyContinue
    }
    Get-Process cloudflared -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $Cf } | Stop-Process -Force -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt 10; $i++) { if (-not (Site-Ativo)) { break }; Start-Sleep 1 }
}

function Iniciar-Publico {
    $pyw = Join-Path $Venv "Scripts\pythonw.exe"
    if (-not (Test-Path $pyw)) { $pyw = $VPy }
    Start-Process $pyw -ArgumentList "`"$(Join-Path $Pasta 'servidor.py')`"", "--publico" -WorkingDirectory $Pasta -WindowStyle Hidden
    for ($i = 0; $i -lt 90; $i++) { Start-Sleep 1; if (Site-Ativo) { break } }
    if (Site-Ativo) { Ok "Copia publica do app rodando (so para o tunel, porta $Porta)" }
    else { throw "A copia publica do app nao respondeu. Veja publico.log." }
}

function Criar-Inicializacao {
    $ps = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $sh = New-Object -ComObject WScript.Shell
    $a = $sh.CreateShortcut($Inicio)
    $a.TargetPath = $ps
    $a.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$(Join-Path $Pasta 'publicar.ps1')`" -Reiniciar -Oculto"
    $a.WorkingDirectory = $Pasta
    $a.WindowStyle = 7
    $a.IconLocation = "$(Join-Path $Pasta 'assets\icone.ico'),0"
    $a.Description = "Verificador ABNT - site publico"
    $a.Save()
    Ok "Site configurado para voltar sozinho quando o Windows iniciar"
}

function Gravar-Link([string]$url) {
    $dir = Split-Path $LinkArq
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Set-Content $LinkArq $url -Encoding ASCII
    Set-Content (Join-Path $dir "Verificador ABNT (site publico).url") "[InternetShortcut]`r`nURL=$url`r`n" -Encoding ASCII
}

# ------------------------------------------------------------------ tuneis
function Garantir-Cloudflared {
    if (Test-Path $Cf) { Ok "cloudflared"; return }
    New-Item -ItemType Directory -Force -Path $Ferr | Out-Null
    Faz "Baixando cloudflared (programa do tunel, sem instalacao)..."
    Invoke-WebRequest -Uri $CfUrl -OutFile $Cf -UseBasicParsing -ErrorAction Stop
    Ok "cloudflared baixado"
}

function Achar-Tailscale {
    foreach ($c in @("$env:ProgramFiles\Tailscale\tailscale.exe", "${env:ProgramFiles(x86)}\Tailscale\tailscale.exe")) {
        if ($c -and (Test-Path $c)) { return $c }
    }
    $cmd = Get-Command tailscale.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Garantir-Tailscale {
    $ts = Achar-Tailscale
    if (-not $ts) {
        Faz "Instalando o Tailscale (o Windows vai pedir a senha de administrador)..."
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            & winget install --id Tailscale.Tailscale -e --silent --accept-source-agreements --accept-package-agreements
        } else {
            $tmp = Join-Path $env:TEMP "tailscale-setup.exe"
            Invoke-WebRequest -Uri "https://pkgs.tailscale.com/stable/tailscale-setup-latest.exe" -OutFile $tmp -UseBasicParsing -ErrorAction Stop
            Start-Process $tmp -ArgumentList "/quiet" -Verb RunAs -Wait
        }
        Start-Sleep 5
        $ts = Achar-Tailscale
        if (-not $ts) {
            throw "Nao foi possivel instalar o Tailscale. Instale manualmente em https://tailscale.com/download/windows (precisa de administrador) e rode publicar.bat de novo."
        }
    }
    Ok "Tailscale"
    $st = (& $ts status --json 2>$null) -join "`n"
    $logado = $false
    try { $logado = (($st | ConvertFrom-Json).BackendState -eq "Running") } catch { }
    if (-not $logado) {
        Faz "Entre com a conta da biblioteca na janela do navegador que vai abrir (uma vez so)..."
        & $ts up
        if ($LASTEXITCODE -ne 0) { throw "Login no Tailscale nao concluido." }
    }
    Ok "Tailscale conectado"
    return $ts
}

function Ligar-Tailscale {
    $ts = Garantir-Tailscale
    Faz "Ligando o Funnel (link publico)..."
    $saida = (& $ts funnel --bg $Porta 2>&1) -join "`n"
    Log $saida
    if ($LASTEXITCODE -ne 0) {
        throw "O Funnel nao foi ligado. Se apareceu um link acima, abra-o para permitir o Funnel na conta e rode publicar.bat de novo."
    }
    $dns = ""
    try { $dns = ((& $ts status --json) -join "`n" | ConvertFrom-Json).Self.DNSName.TrimEnd(".") } catch { }
    if (-not $dns) { throw "Nao foi possivel descobrir o endereco do Tailscale." }
    return "https://$dns"
}

function Esperar-Link([int]$segundos = 60) {
    for ($i = 0; $i -lt $segundos; $i++) {
        $c = Ler-Json $CfgArq
        if ($c -and $c.endereco) { return $c.endereco }
        Start-Sleep 1
    }
    return ""
}

# ------------------------------------------------------------------ principal
try {
    Log ""
    Log "=== Verificador ABNT - site publico ===" "White"

    if ($Desativar) {
        $c = Ler-Json $CfgArq
        Parar-Publico
        if ($c -and $c.tunel -eq "tailscale") {
            $ts = Achar-Tailscale
            if ($ts) { & $ts funnel --bg $Porta off *> $null; & $ts funnel reset *> $null }
        }
        Remove-Item $Inicio, $LinkArq, (Join-Path (Split-Path $LinkArq) "Verificador ABNT (site publico).url") -ErrorAction SilentlyContinue
        if ($c) { $c | Add-Member -NotePropertyName ativo -NotePropertyValue $false -Force; Salvar-Cfg $c }
        Ok "Site publico desativado. O app interno da biblioteca continua funcionando."
        exit 0
    }

    if (-not (Test-Path $VPy)) { throw "O Verificador ainda nao foi instalado neste PC. Rode instalar.bat primeiro." }

    $cfg = Ler-Json $CfgArq
    if ($Reiniciar -and -not $Tunel) {
        if (-not $cfg -or -not $cfg.ativo) { Log "Site publico nao esta ativado. Nada a fazer."; exit 0 }
        $Tunel = $cfg.tunel
    }
    if (-not $Tunel) {
        Log ""
        Log "Qual tunel usar para o link publico?" "White"
        Log "  1) Tailscale  - link fixo e gratuito (recomendado; instala o Tailscale, pede administrador uma vez)"
        Log "  2) Cloudflare - link fixo no seu dominio (precisa de token do painel da Cloudflare)"
        Log "  3) Rapido     - sem conta; o link muda a cada reinicio (bom para testar)"
        $r = Read-Host "Digite 1, 2 ou 3"
        $Tunel = @{ "1" = "tailscale"; "2" = "cloudflare"; "3" = "rapido" }[$r.Trim()]
        if (-not $Tunel) { throw "Opcao invalida." }
    }
    if ($Tunel -eq "cloudflare") {
        if (-not $Token -and $cfg -and $cfg.token) { $Token = $cfg.token }
        if (-not $Endereco -and $cfg -and $cfg.endereco -and $cfg.tunel -eq "cloudflare") { $Endereco = $cfg.endereco }
        if (-not $Token) { $Token = Read-Host "Cole o token do tunel (Cloudflare Zero Trust > Networks > Tunnels)" }
        if (-not $Endereco) { $Endereco = Read-Host "Endereco publico configurado no tunel (ex.: https://verificador.seudominio.org)" }
        if (-not $Token -or -not $Endereco) { throw "Token e endereco sao obrigatorios para o tunel da Cloudflare." }
        if ($Endereco -notmatch "^https?://") { $Endereco = "https://$Endereco" }
    }

    $app = Ler-Json $ConfigApp
    if (-not $app -or -not $app.senha_admin) {
        Aviso "Nao ha senha de administrador. No site publico as configuracoes ficam bloqueadas (o que e seguro)."
        Aviso "Para poder ajustar pelo site, defina a senha no app deste PC: barra lateral > Senhas."
    }

    $nova = [ordered]@{ tunel = $Tunel; ativo = $true; endereco = ""; token = "" }
    if ($Tunel -eq "cloudflare") { $nova.token = $Token; $nova.endereco = $Endereco }
    Salvar-Cfg $nova

    Parar-Publico
    if ($Tunel -ne "tailscale") { Garantir-Cloudflared }
    Iniciar-Publico

    $link = ""
    if ($Tunel -eq "tailscale") {
        $link = Ligar-Tailscale
        $nova.endereco = $link
        Salvar-Cfg $nova
        Gravar-Link $link
    } elseif ($Tunel -eq "cloudflare") {
        $link = $Endereco
    } else {
        Faz "Aguardando o link do tunel rapido..."
        $link = Esperar-Link 60
        if (-not $link) { throw "O tunel nao informou o link. Veja publico.log." }
    }
    Criar-Inicializacao

    Log ""
    Log "Site publico no ar:" "Green"
    Log "   $link" "Green"
    if ($Tunel -eq "rapido") { Aviso "Esse link muda quando o PC reinicia. Para um link fixo, use Tailscale ou Cloudflare." }
    Log "O link tambem fica em compartilhar\link_publico.txt. O PC precisa ficar ligado e conectado."
    if (-not $Oculto) {
        try { Set-Clipboard -Value $link; Ok "Link copiado para a area de transferencia" } catch { }
        Start-Process $link
    }
    exit 0
} catch {
    Log ""
    Log "ERRO: $($_.Exception.Message)" "Red"
    exit 1
}
