# setup_propostas.ps1
# -------------------
# 1) Cria a pasta "Propostas" com ícone “pasta.ico”
# 2) Gera atalhos na Área de Trabalho:
#    • Propostas.lnk → a pasta
#    • Sistema de Propostas.lnk → o seu .exe

# caminho do script
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# 1) cria pasta
$folderName = 'Propostas'
$folderPath = Join-Path $root $folderName
if (-not (Test-Path $folderPath)) {
  New-Item -ItemType Directory -Path $folderPath | Out-Null
}

# 2) desktop.ini p/ ícone da pasta
$iconSource   = Join-Path $root 'pasta.ico'
$desktopIni   = Join-Path $folderPath 'desktop.ini'

@"
[.ShellClassInfo]
IconResource=$iconSource,0
"@ | Out-File -FilePath $desktopIni -Encoding ASCII -Force

# 3) ajusta atributos (SYSTEM na pasta, HIDDEN+SYSTEM no desktop.ini)
$sys  = [System.IO.FileAttributes]::System
$ro   = [System.IO.FileAttributes]::ReadOnly
$hid  = [System.IO.FileAttributes]::Hidden

(Get-Item $folderPath).Attributes    = $sys -bor $ro
(Get-Item $desktopIni).Attributes    = $hid -bor $sys

# 4) shortcuts na Área de Trabalho
$desktop = [Environment]::GetFolderPath('Desktop')
$Wsh     = New-Object -ComObject WScript.Shell

# → pasta
$lnk1 = $Wsh.CreateShortcut((Join-Path $desktop 'Propostas.lnk'))
$lnk1.TargetPath      = $folderPath
$lnk1.WorkingDirectory= $folderPath
$lnk1.IconLocation    = $iconSource
$lnk1.Save()

# → executável
$exeName = 'GeProp.exe'      # ou troque pelo nome do seu .exe
$exePath = Join-Path $root $exeName
$lnk2 = $Wsh.CreateShortcut((Join-Path $desktop 'GeProp.lnk'))
$lnk2.TargetPath      = $exePath
$lnk2.WorkingDirectory= $root
$lnk2.IconLocation    = $exePath
$lnk2.Save()
