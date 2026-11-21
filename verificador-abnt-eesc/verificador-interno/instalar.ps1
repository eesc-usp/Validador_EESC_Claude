# Verificador ABNT - Biblioteca EESC
# Instala, verifica e repara tudo o que o app precisa, sem administrador e sem reiniciar.
#
#   instalar.bat            -> verificacao completa: instala o que falta e corrige o que esta errado
#   iniciar.bat             -> verificacao rapida; se a versao mudou ou algo sumiu, faz a completa; abre o app
#
# Parametros:
#   -Rapido      so faz a verificacao completa se a versao mudou ou faltar algum componente
#   -Iniciar     ao final, abre o app (janela propria do aplicativo)
#   -Navegador   com -Iniciar, abre no navegador em vez da janela propria
#   -Oculto      usado pelo atalho: se precisar instalar/reparar, reabre esta janela visivel
#   -Servidor    deixa este PC como servidor da biblioteca: todos acessam pelo navegador na rede
#   -DesativarServidor  para o servidor e tira da inicializacao do Windows
#   -SemIA       nao instala Ollama nem modelo (so a verificacao por regras)
#   -Modelo X    forca um modelo (ex.: qwen2.5:3b)

param(
    [switch]$Rapido,
    [switch]$Iniciar,
    [switch]$SemIA,
    [switch]$Navegador,
    [switch]$Oculto,
    [switch]$Servidor,
    [switch]$DesativarServidor,
    [string]$Modelo = ""
)

# "Continue": no PowerShell 5, avisos de programas externos (pip, ollama) no stderr
# virariam erro fatal com "Stop". Erros sao checados por $LASTEXITCODE e -ErrorAction Stop.
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"   # downloads muito mais rapidos no PowerShell 5
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Pasta      = $PSScriptRoot
$LogArq     = Join-Path $Pasta "instalar.log"
$EstadoArq  = Join-Path $Pasta ".instalado.json"
$ConfigArq  = Join-Path $Pasta "config.json"
$Venv       = Join-Path $Pasta ".venv"
$VPy        = Join-Path $Venv "Scripts\python.exe"
$Versao     = (Get-Content (Join-Path $Pasta "VERSION") -Raw).Trim()

$PythonUrl  = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe"
$OllamaUrl  = "https://ollama.com/download/OllamaSetup.exe"
$OllamaMin  = [version]"0.5.0"
$OllamaApi  = "http://localhost:11434"
$Porta      = 8501
$AppUrl     = "http://localhost:$Porta"
$RegraFw    = "Verificador ABNT (porta $Porta)"
$Arquivos   = @("app.py", "requirements.txt", "VERSION", ".streamlit\config.toml",
                "verificador\__init__.py", "verificador\documento.py", "verificador\docx_util.py",
                "verificador\estrutura.py", "verificador\referencias.py", "verificador\ia_local.py",
                "verificador\relatorio.py", "verificador\modelo.py", "verificador\correcao.py",
                "verificador\visualizacao.py", "verificador\edicao.py", "componentes\area_trabalho\index.html",
                "desktop.py", "assets\icone.ico", "requirements-desktop.txt", "servidor.py",
                "verificador\ia_provedores.py", "verificador\configuracao.py", "verificador\tabelas.py", "verificador\anotacao.py", "publicar.ps1",
                "verificador\perfil.py", "verificador\extras.py", "verificador\trabalho.py", "verificador\elementos.py", "verificador\tipografia.py")

function Log([string]$msg, [string]$cor = "Gray") {
    Write-Host $msg -ForegroundColor $cor
    Add-Content -Path $LogArq -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg)
}
function Ok([string]$m)    { Log "  [OK] $m" "Green" }
function Faz([string]$m)   { Log "  [..] $m" "Cyan" }
function Aviso([string]$m) { Log "  [!!] $m" "Yellow" }

function Baixar([string]$url, [string]$destino) {
    Faz "Baixando $(Split-Path $url -Leaf) ..."
    Invoke-WebRequest -Uri $url -OutFile $destino -UseBasicParsing -ErrorAction Stop
}

function Ler-Json([string]$arq) {
    if (Test-Path $arq) { try { return Get-Content $arq -Raw | ConvertFrom-Json } catch { } }
    return $null
}

# ------------------------------------------------------------------ arquivos do app
function Verificar-Arquivos {
    $faltando = @($Arquivos | Where-Object { -not (Test-Path (Join-Path $Pasta $_)) })
    if ($faltando.Count) {
        throw "Pasta do app incompleta. Faltam: $($faltando -join ', '). Copie a pasta inteira de novo."
    }
    Ok "Arquivos do app (versao $Versao)"
}

# ------------------------------------------------------------------ Python
function Testar-Python([string]$exe, [string[]]$pre = @()) {
    try {
        $saida = & $exe @pre -c "import sys; print(sys.version_info[0]*100+sys.version_info[1]); print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $saida.Count -ge 2 -and [int]$saida[0] -ge 310) { return $saida[1].Trim() }
    } catch { }
    return $null
}

function Achar-Python {
    $p = Testar-Python "py" @("-3"); if ($p) { return $p }
    $p = Testar-Python "python";     if ($p) { return $p }
    $locais = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue |
              Sort-Object FullName -Descending
    foreach ($l in $locais) { $p = Testar-Python $l.FullName; if ($p) { return $p } }
    return $null
}

function Garantir-Python {
    $py = Achar-Python
    if ($py) { Ok "Python: $py"; return $py }
    Faz "Python 3.10+ nao encontrado. Instalando Python 3.12 (so para este usuario)..."
    $inst = Join-Path $env:TEMP "python-instalador.exe"
    Baixar $PythonUrl $inst
    $p = Start-Process $inst -Wait -PassThru -ArgumentList @(
        "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_launcher=1",
        "InstallLauncherAllUsers=0", "Include_test=0", "Include_doc=0")
    Remove-Item $inst -ErrorAction SilentlyContinue
    if ($p.ExitCode -ne 0) { throw "Instalador do Python terminou com codigo $($p.ExitCode)." }
    $py = Achar-Python
    if (-not $py) { throw "Python instalado, mas nao localizado. Feche e abra o instalar.bat de novo." }
    Ok "Python instalado: $py"
    return $py
}

# ------------------------------------------------------------------ ambiente e bibliotecas
function Garantir-Ambiente([string]$py) {
    $saudavel = $false
    if (Test-Path $VPy) {
        & $VPy -c "import sys" 2>$null
        $saudavel = ($LASTEXITCODE -eq 0)
    }
    if (-not $saudavel) {
        if (Test-Path $Venv) { Aviso "Ambiente Python danificado: recriando."; Remove-Item $Venv -Recurse -Force -ErrorAction Stop }
        Faz "Criando ambiente Python do app..."
        & $py -m venv $Venv
        if ($LASTEXITCODE -ne 0) { throw "Falha ao criar o ambiente Python (.venv)." }
    }
    Faz "Conferindo bibliotecas (instala o que falta, atualiza o que estiver em versao errada)..."
    & $VPy -m pip install --upgrade pip --quiet --disable-pip-version-check 2>$null
    & $VPy -m pip install -r (Join-Path $Pasta "requirements.txt") --upgrade --quiet --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar as bibliotecas. Verifique a internet e veja instalar.log." }
    & $VPy -m pip check *> $null
    $importa = & $VPy -c "import streamlit, docx, requests, pandas; print('ok')" 2>$null
    if ($LASTEXITCODE -ne 0 -or $importa -ne "ok") {
        Aviso "Bibliotecas inconsistentes: reinstalando."
        & $VPy -m pip install -r (Join-Path $Pasta "requirements.txt") --force-reinstall --quiet --disable-pip-version-check
        if ($LASTEXITCODE -ne 0) { throw "Falha ao reinstalar as bibliotecas." }
    }
    Ok "Bibliotecas do app"
}

# ------------------------------------------------------------------ janela do aplicativo
function Garantir-Janela {
    & $VPy -c "import webview" 2>$null
    if ($LASTEXITCODE -eq 0) { Ok "Janela do aplicativo"; return $true }
    Faz "Instalando a janela do aplicativo (pywebview)..."
    & $VPy -m pip install -r (Join-Path $Pasta "requirements-desktop.txt") --upgrade --quiet --disable-pip-version-check
    & $VPy -c "import webview" 2>$null
    if ($LASTEXITCODE -eq 0) { Ok "Janela do aplicativo"; return $true }
    Aviso "Janela propria indisponivel neste PC: o app vai abrir no navegador."
    return $false
}

# ------------------------------------------------------------------ Ollama
function Achar-Ollama {
    $cmd = Get-Command ollama -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $local = Join-Path $env:LOCALAPPDATA "Programs\Ollama\ollama.exe"
    if (Test-Path $local) { return $local }
    return $null
}

function Versao-Ollama([string]$exe) {
    try {
        $txt = (& $exe --version 2>&1) -join " "
        if ($txt -match "(\d+\.\d+\.\d+)") { return [version]$Matches[1] }
    } catch { }
    return $null
}

function Ollama-Ativo {
    try { Invoke-RestMethod "$OllamaApi/api/tags" -TimeoutSec 3 -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}

function Iniciar-Ollama([string]$exe) {
    if (Ollama-Ativo) { return $true }
    $app = Join-Path (Split-Path $exe) "ollama app.exe"
    if (Test-Path $app) { Start-Process $app } else { Start-Process $exe -ArgumentList "serve" -WindowStyle Hidden }
    for ($i = 0; $i -lt 30; $i++) { Start-Sleep 1; if (Ollama-Ativo) { return $true } }
    return $false
}

function Garantir-Ollama {
    $exe = Achar-Ollama
    $v = if ($exe) { Versao-Ollama $exe } else { $null }
    if (-not $exe -or -not $v -or $v -lt $OllamaMin) {
        if ($exe) { Aviso "Ollama $v desatualizado ou danificado: reinstalando." }
        else { Faz "Ollama nao encontrado. Instalando (download grande, pode levar alguns minutos)..." }
        $inst = Join-Path $env:TEMP "OllamaSetup.exe"
        Baixar $OllamaUrl $inst
        $p = Start-Process $inst -Wait -PassThru -ArgumentList "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES"
        Remove-Item $inst -ErrorAction SilentlyContinue
        if ($p.ExitCode -ne 0) { throw "Instalador do Ollama terminou com codigo $($p.ExitCode)." }
        $env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" +
                    [Environment]::GetEnvironmentVariable("Path", "Machine")
        $exe = Achar-Ollama
        if (-not $exe) { throw "Ollama instalado, mas nao localizado." }
        $v = Versao-Ollama $exe
    }
    Ok "Ollama $v"
    if (-not (Iniciar-Ollama $exe)) { throw "O Ollama nao iniciou. Abra 'Ollama' no menu Iniciar e rode de novo." }
    Ok "Ollama em execucao"
    return $exe
}

function Escolher-Modelo {
    if ($Modelo) { return $Modelo }
    $cfg = Ler-Json $ConfigArq
    if ($cfg -and $cfg.modelo) { return $cfg.modelo }   # mantem a escolha feita antes
    try { $ram = [math]::Round((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1GB) }
    catch { $ram = 8 }
    Log "  [--] Memoria RAM: $ram GB"
    if ($ram -ge 15) { return "qwen2.5:7b" } else { return "qwen2.5:3b" }
}

function Garantir-Modelo([string]$exe, [string]$nome) {
    $lista = (& $exe list 2>$null) -join "`n"
    $tem = $lista -match [regex]::Escape($nome)
    if ($tem) {
        & $exe show $nome *> $null
        if ($LASTEXITCODE -ne 0) { Aviso "Modelo $nome danificado: baixando de novo."; & $exe rm $nome *> $null; $tem = $false }
    }
    if (-not $tem) {
        Faz "Baixando o modelo $nome (2 a 5 GB, uma vez so)..."
        & $exe pull $nome
        if ($LASTEXITCODE -ne 0) { throw "Falha ao baixar o modelo $nome." }
    }
    Ok "Modelo de IA: $nome"
}

# ------------------------------------------------------------------ configuracao e atalho
function Salvar-Config([string]$modelo) {
    $cfg = Ler-Json $ConfigArq
    $dados = [ordered]@{}
    if ($cfg) { $cfg.PSObject.Properties | ForEach-Object { $dados[$_.Name] = $_.Value } }
    if ($modelo) { $dados["modelo"] = $modelo }
    if (-not $dados.Contains("nucleos")) { $dados["nucleos"] = [math]::Max(1, [Environment]::ProcessorCount - 2) }
    if (-not $dados.Contains("ollama_url")) { $dados["ollama_url"] = $OllamaApi }
    ($dados | ConvertTo-Json) | Set-Content $ConfigArq -Encoding UTF8 -ErrorAction Stop
    Ok "Configuracao (config.json)"
}

function Criar-Atalho([string]$lnk, [string]$modo = "-Iniciar") {
    $ps   = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $argumentos = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$(Join-Path $Pasta 'instalar.ps1')`" -Rapido $modo -Oculto"
    $icone = Join-Path $Pasta "assets\icone.ico"
    $sh = New-Object -ComObject WScript.Shell
    if (Test-Path $lnk) {
        $atual = $sh.CreateShortcut($lnk)
        if ($atual.TargetPath -eq $ps -and $atual.Arguments -eq $argumentos -and $atual.IconLocation -like "$icone*") { return $false }
    }
    New-Item -ItemType Directory -Force -Path (Split-Path $lnk) | Out-Null
    $a = $sh.CreateShortcut($lnk)
    $a.TargetPath = $ps
    $a.Arguments = $argumentos
    $a.WorkingDirectory = $Pasta
    $a.WindowStyle = 7
    $a.IconLocation = "$icone,0"
    $a.Description = "Verificador ABNT - Biblioteca EESC"
    $a.Save()
    return $true
}

function Garantir-Atalho {
    $nome = "Verificador ABNT EESC.lnk"
    $mesa = Join-Path ([Environment]::GetFolderPath("Desktop")) $nome
    $menu = Join-Path ([Environment]::GetFolderPath("Programs")) $nome
    $m1 = Criar-Atalho $mesa
    $m2 = Criar-Atalho $menu
    $mudou = $m1 -or $m2
    if ($mudou) { Ok "Atalhos criados/corrigidos (area de trabalho e menu Iniciar)" } else { Ok "Atalhos" }
}

# ------------------------------------------------------------------ servidor da biblioteca
function Parar-Servidor {
    $pidArq = Join-Path $Pasta "servidor.pid"
    if (-not (Test-Path $pidArq)) { return }
    New-Item -ItemType File -Force -Path (Join-Path $Pasta "servidor.parar") | Out-Null
    for ($i = 0; $i -lt 15; $i++) { Start-Sleep 1; if (-not (Test-Path $pidArq)) { break } }
    if (Test-Path $pidArq) {
        $id = (Get-Content $pidArq -Raw).Trim()
        & taskkill /PID $id /T /F *> $null
        Remove-Item $pidArq -ErrorAction SilentlyContinue
    }
    Remove-Item (Join-Path $Pasta "servidor.parar") -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt 10; $i++) { if (-not (App-Ativo)) { break }; Start-Sleep 1 }
    Ok "Servidor anterior parado"
}

function Iniciar-Servidor {
    if (App-Ativo) { Ok "Servidor em execucao"; return }
    $pyw = Join-Path $Venv "Scripts\pythonw.exe"
    if (-not (Test-Path $pyw)) { $pyw = $VPy }
    Start-Process $pyw -ArgumentList "`"$(Join-Path $Pasta 'servidor.py')`"" -WorkingDirectory $Pasta -WindowStyle Hidden
    for ($i = 0; $i -lt 90; $i++) { Start-Sleep 1; if (App-Ativo) { break } }
    if (App-Ativo) { Ok "Servidor em execucao" } else { throw "O servidor nao respondeu. Veja servidor.log." }
}

function Garantir-Firewall {
    $existe = $false
    try { $existe = [bool](Get-NetFirewallRule -DisplayName $RegraFw -ErrorAction Stop) } catch { }
    if ($existe) { Ok "Firewall: porta $Porta liberada"; return }
    Faz "Liberando a porta $Porta no firewall (o Windows vai pedir permissao de administrador)..."
    $cmd = "New-NetFirewallRule -DisplayName '$RegraFw' -Direction Inbound -Protocol TCP -LocalPort $Porta " +
           "-Action Allow -Profile Domain,Private | Out-Null"
    try {
        Start-Process powershell -Verb RunAs -Wait -WindowStyle Hidden -ErrorAction Stop `
            -ArgumentList "-NoProfile", "-Command", $cmd | Out-Null
    } catch { }
    try { Get-NetFirewallRule -DisplayName $RegraFw -ErrorAction Stop | Out-Null; Ok "Firewall: porta $Porta liberada" }
    catch {
        Aviso "Nao foi possivel liberar o firewall (precisa de administrador)."
        Aviso "Peca ao TI: liberar ENTRADA TCP na porta $Porta, redes de dominio/privada, neste computador."
    }
}

function Garantir-Inicializacao {
    $lnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Verificador ABNT - Servidor.lnk"
    if (Criar-Atalho $lnk "-Servidor") { Ok "Servidor configurado para iniciar com o Windows" }
    else { Ok "Inicia com o Windows" }
}

function Remover-Inicializacao {
    $lnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Verificador ABNT - Servidor.lnk"
    Remove-Item $lnk -ErrorAction SilentlyContinue
}

function Enderecos {
    $ips = @()
    try {
        $ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
               Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } |
               Select-Object -ExpandProperty IPAddress
    } catch { }
    return @("http://$($env:COMPUTERNAME):$Porta") + @($ips | ForEach-Object { "http://$($_):$Porta" })
}

function Gerar-AtalhoRede([string[]]$urls) {
    $dir = Join-Path $Pasta "compartilhar"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $icone = Join-Path $Pasta "assets\icone.ico"
    $conteudo = "[InternetShortcut]`r`nURL=$($urls[0])`r`n"
    Set-Content (Join-Path $dir "Verificador ABNT (rede).url") $conteudo -Encoding ASCII
    if ($urls.Count -gt 1) {
        Set-Content (Join-Path $dir "Verificador ABNT (rede - IP).url") "[InternetShortcut]`r`nURL=$($urls[1])`r`n" -Encoding ASCII
    }
    Copy-Item $icone (Join-Path $dir "icone.ico") -Force
    $leia = "Verificador ABNT - acesso pela rede da biblioteca`r`n`r`n" +
            "Copie o atalho 'Verificador ABNT (rede)' para a area de trabalho e abra com duplo clique.`r`n" +
            "Se nao abrir, use o atalho com IP. Enderecos:`r`n" + (($urls | ForEach-Object { "  $_" }) -join "`r`n") +
            "`r`n`r`nNada precisa ser instalado: o app roda no computador $($env:COMPUTERNAME)."
    Set-Content (Join-Path $dir "LEIA-ME.txt") $leia -Encoding UTF8
    Ok "Atalho de rede gerado na pasta 'compartilhar'"
}

function Configurar-Servidor {
    Log ""
    Log "--- Servidor da biblioteca ---" "White"
    Parar-Servidor
    Iniciar-Servidor
    Garantir-Firewall
    Garantir-Inicializacao
    $urls = Enderecos
    Gerar-AtalhoRede $urls
    Log ""
    Log "Todos na rede acessam pelo navegador:" "Green"
    $urls | ForEach-Object { Log "   $_" "Green" }
    Log "Configure as chaves de IA e a senha de administrador abrindo $AppUrl neste computador." "Green"
}

# ------------------------------------------------------------------ app
function App-Ativo {
    try { Invoke-WebRequest "$AppUrl/_stcore/health" -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop | Out-Null; return $true }
    catch { return $false }
}

function Abrir-App {
    $janela = Join-Path $Venv "Scripts\pythonw.exe"
    if (-not $Navegador -and (Test-Path $janela)) {
        Start-Process $janela -ArgumentList "`"$(Join-Path $Pasta 'desktop.py')`"" -WorkingDirectory $Pasta
        Ok "Aplicativo aberto"
        return
    }
    if (-not (App-Ativo)) {
        Faz "Iniciando o app..."
        Start-Process $VPy -ArgumentList "-m", "streamlit", "run", "app.py" -WorkingDirectory $Pasta -WindowStyle Minimized
        for ($i = 0; $i -lt 60; $i++) { Start-Sleep 1; if (App-Ativo) { break } }
    }
    if (App-Ativo) { Start-Process $AppUrl; Ok "App aberto em $AppUrl" }
    else { throw "O app nao respondeu. Veja a janela minimizada do app para o erro." }
}

# ------------------------------------------------------------------ principal
try {
    $estado = Ler-Json $EstadoArq
    if ($DesativarServidor) {
        Parar-Servidor
        Remover-Inicializacao
        if ($estado) { $estado.servidor = $false; $estado | ConvertTo-Json | Set-Content $EstadoArq -Encoding UTF8 }
        Log "Servidor desativado. O app continua funcionando neste PC pelo atalho." "Green"
        exit 0
    }
    $modoServidor = [bool]$Servidor -or [bool]($estado -and $estado.servidor)
    $desligarIA = [bool]$SemIA -or [bool]($estado -and $estado.sem_ia -and -not $Modelo)
    $completo = $true
    if ($Rapido -and $estado -and $estado.versao -eq $Versao -and (Test-Path $VPy) -and
        (Test-Path (Join-Path $Venv "Scripts\streamlit.exe"))) {
        $completo = $false
    }

    if ($completo -and $Oculto) {
        # aberto pelo atalho (janela escondida) mas precisa instalar/reparar: reabre visivel para mostrar o progresso
        $extra = @()
        if ($Iniciar) { $extra += "-Iniciar" }
        if ($Servidor) { $extra += "-Servidor" }
        Start-Process powershell -ArgumentList (@("-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            "`"$PSCommandPath`"") + $extra)
        exit 0
    }

    if ($completo) {
        $motivo = if (-not $estado) { "primeira instalacao" }
                  elseif ($estado.versao -ne $Versao) { "nova versao: $($estado.versao) -> $Versao" }
                  else { "verificacao completa" }
        Log ""
        Log "=== Verificador ABNT - Biblioteca EESC | $motivo ===" "White"
        Verificar-Arquivos
        $py = Garantir-Python
        Garantir-Ambiente $py
        $temJanela = Garantir-Janela
        $modeloFinal = $null
        $verOllama = $null
        if (-not $desligarIA) {
            try {
                $ollama = Garantir-Ollama
                $verOllama = "$(Versao-Ollama $ollama)"
                $modeloFinal = Escolher-Modelo
                Garantir-Modelo $ollama $modeloFinal
            } catch {
                Aviso "IA local nao configurada: $($_.Exception.Message)"
                Aviso "O app funciona sem IA. Rode instalar.bat de novo para tentar outra vez."
            }
        } else { Log "  [--] IA local desativada (-SemIA)" }
        Salvar-Config $modeloFinal
        Garantir-Atalho
        if ($modoServidor) { Configurar-Servidor }
        $pub = Ler-Json (Join-Path $Pasta "publico.json")
        if ($pub -and $pub.ativo) {
            Faz "Reiniciando o site publico com a versao nova..."
            & (Join-Path $Pasta "publicar.ps1") -Reiniciar -Oculto
            if ($LASTEXITCODE -eq 0) { Ok "Site publico atualizado" } else { Aviso "Site publico nao reiniciou; veja publicar.log" }
        }
        [ordered]@{ versao = $Versao; data = (Get-Date -Format "yyyy-MM-dd HH:mm"); python = $py; servidor = $modoServidor;
                    ollama = $verOllama; modelo = $modeloFinal; sem_ia = $desligarIA; janela = $temJanela } |
            ConvertTo-Json | Set-Content $EstadoArq -Encoding UTF8
        Log "=== Tudo pronto ===" "Green"
    } else {
        if (-not $desligarIA) {
            $ollama = Achar-Ollama
            if ($ollama -and -not (Ollama-Ativo)) { Iniciar-Ollama $ollama | Out-Null }
        }
        if ($modoServidor -and -not (App-Ativo)) { Iniciar-Servidor }
    }

    if ($Servidor -and -not $Iniciar -and -not $Oculto) { Start-Process $AppUrl }

    if ($Iniciar) { Abrir-App }
    exit 0
} catch {
    Log "ERRO: $($_.Exception.Message)" "Red"
    Log "Detalhes em $LogArq" "Red"
    exit 1
}
