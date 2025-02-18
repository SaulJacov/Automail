# Crea un nuevo perfil de Outlook
New-OutlookProfile -Name "MiPerfil"

# Evitar que la ventana de PowerShell se cierre automáticamente
Write-Host "Presiona cualquier tecla para cerrar esta ventana..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null