# Função para reiniciar o script com privilégios elevados
function Run-AsAdmin {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command $arguments" -Verb RunAs
    exit
}

# Verificar se o script está sendo executado como administrador
if (-not (Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA")) {
    Write-Host "Este script precisa ser executado como administrador."
    exit
}

# Se o script não estiver rodando como administrador, reiniciar com privilégios elevados
if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA')) {
    Run-AsAdmin
}

# Mapeando a unidade da Rede que é onde o executavel do Office365 está
$networkPath = "\\plane\office365"
$driveLetter = "B:"

#Verificar se a unidade de rede já está mapeada
if (-not (Test-Path $driveLetter)) {
    #Mapeando a unidade de rede
    net use $driveLetter $networkPath
}

# Alterando para a unidade de rede mapeada
Set-Location -Path $driveLetter

# Executando o instalador do Office 365
Start-Process -FilePath "Setup.exe" -ArgumentList "/configure Office365.xml" -Waith