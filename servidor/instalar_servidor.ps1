<#
INSTALACION DE HERRAMIENTAS GE&PE EN EL PC DE LA OFICINA
=========================================================
Ejecutar en el servidor, en PowerShell COMO ADMINISTRADOR (ver GUIA_SERVIDOR.md):

    powershell -ExecutionPolicy Bypass -File servidor\instalar_servidor.ps1 `
        -Subredes "192.168.0.0/24,10.8.0.0/24" -Usuario "GEYPE\svc_herramientas"

Crea (o sustituye):
  1. Regla del cortafuegos: puerto de la app solo desde las subredes indicadas (oficina y VPN).
  2. Tarea "Herramientas GEYPE - servidor": arranca la plataforma al encender el servidor y la reinicia
     si se cae.
  3. Tareas de datos: descarga diaria del PVPC, mensual de componentes de ESIOS y copia de
     seguridad diaria de la base de datos.
No toca datos ni credenciales.
#>
param(
    [Parameter(Mandatory = $true)] [string] $Subredes,      # "192.168.0.0/24,10.8.0.0/24"
    [Parameter(Mandatory = $true)] [string] $Usuario,       # cuenta con la que corren las tareas
    [int]    $Puerto = 8501,
    [string] $Python = (Get-Command python -ErrorAction SilentlyContinue).Source,
    [string] $CarpetaApp = (Split-Path -Parent $PSScriptRoot),
    [string] $CarpetaEsios = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "Descarga_datos_ESIOS"),
    [string] $CarpetaCopias = (Join-Path (Split-Path -Parent $PSScriptRoot) "copias_seguridad")
)
$ErrorActionPreference = "Stop"

function Comprobar($condicion, $mensaje) { if (-not $condicion) { throw $mensaje } }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Comprobar $admin "Hay que ejecutar este script como administrador."
Comprobar ($Python -and (Test-Path $Python)) "No se encuentra python.exe. Indícalo con -Python."
Comprobar (Test-Path (Join-Path $CarpetaApp "lanzador.py")) "No está lanzador.py en $CarpetaApp."
Comprobar (Test-Path (Join-Path $CarpetaEsios "descarga_PVPC_diario_excel.py")) `
    "No está Descarga_datos_ESIOS en $CarpetaEsios. Indícalo con -CarpetaEsios."

$redes = $Subredes.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
Comprobar ($redes.Count -gt 0) "Indica al menos una subred en -Subredes."

Write-Host "Contraseña de la cuenta $Usuario (para que las tareas corran sin sesión iniciada):"
$clave = Read-Host -AsSecureString
$claveTexto = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($clave))

# 0. instalacion anterior y extractor antiguo -----------------------------------------
Get-ScheduledTask -TaskName "Revisor SSAA - *" -ErrorAction SilentlyContinue | ForEach-Object {
    Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false
    Write-Host "Quitada tarea antigua: $($_.TaskName)"
}
Get-NetFirewallRule -DisplayName "Revisor SSAA (puerto *" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
$ocupado = Get-NetTCPConnection -LocalPort $Puerto -State Listen -ErrorAction SilentlyContinue
if ($ocupado) {
    $proc = Get-Process -Id ($ocupado | Select-Object -First 1).OwningProcess -ErrorAction SilentlyContinue
    Write-Warning ("El puerto $Puerto lo usa ahora '$($proc.ProcessName)' (PID $($proc.Id)), " +
                   "seguramente el extractor de Gemweb antiguo. Ciérralo y quita su arranque " +
                   "automático (GUIA_SERVIDOR.md, punto 0) antes de seguir.")
    throw "Puerto $Puerto ocupado."
}

# 1. cortafuegos ------------------------------------------------------------------------
$nombreRegla = "Herramientas GEYPE (puerto $Puerto)"
Get-NetFirewallRule -DisplayName $nombreRegla -ErrorAction SilentlyContinue | Remove-NetFirewallRule
New-NetFirewallRule -DisplayName $nombreRegla -Direction Inbound -Protocol TCP -LocalPort $Puerto `
    -RemoteAddress $redes -Action Allow -Profile Domain, Private | Out-Null
Write-Host "Cortafuegos: puerto $Puerto abierto solo para $($redes -join ', ')."

# 2. tareas programadas ----------------------------------------------------------------
function Registrar($nombre, $programa, $argumentos, $carpeta, $disparador, $reiniciar) {
    $accion = New-ScheduledTaskAction -Execute $programa -Argument $argumentos -WorkingDirectory $carpeta
    $ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)
    if ($reiniciar) {
        $ajustes.RestartCount = 999
        $ajustes.RestartInterval = "PT1M"
    }
    Unregister-ScheduledTask -TaskName $nombre -Confirm:$false -ErrorAction SilentlyContinue
    Register-ScheduledTask -TaskName $nombre -Action $accion -Trigger $disparador -Settings $ajustes `
        -User $Usuario -Password $claveTexto -RunLevel Limited | Out-Null
    Write-Host "Tarea: $nombre"
}

Registrar "Herramientas GEYPE - servidor" $Python "lanzador.py --servidor --puerto $Puerto" $CarpetaApp `
    (New-ScheduledTaskTrigger -AtStartup) $true
Registrar "Herramientas GEYPE - descarga PVPC diaria" $Python "descarga_PVPC_diario_excel.py" $CarpetaEsios `
    (New-ScheduledTaskTrigger -Daily -At "08:30") $false
$semanal = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Monday -At "09:00"
Registrar "Herramientas GEYPE - descarga componentes ESIOS" $Python "descarga_componentes_precio_excel.py" `
    $CarpetaEsios $semanal $false
Registrar "Herramientas GEYPE - copia de seguridad" $Python "almacen.py --copia `"$CarpetaCopias`" --conservar 30" `
    $CarpetaApp (New-ScheduledTaskTrigger -Daily -At "23:00") $false

$claveTexto = $null
Start-ScheduledTask -TaskName "Herramientas GEYPE - servidor"
Write-Host ""
Write-Host "Listo. En unos 20 segundos la plataforma estará en http://$($env:COMPUTERNAME):$Puerto"
Write-Host "El primer usuario que entre creará el administrador."
