$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace()

# Nombre del perfil (si no existe, Outlook lo creará)
$profileName = "MiNuevoPerfil"  # O el nombre que desees

# Datos de la cuenta
$EmailAddress = "soporte5@soluciona.com.mx"
$Password = "temp@2025"
$DisplayName = "Soporte Soluciona"
$POPServer = "mail.soluciona.com.mx"
$POPPort = 110
$SMTPServer = "mail.soluciona.com.mx"
$SMTPPort = 587

try {
    # Obtener la colección de cuentas del perfil (o crear el perfil si no existe)
    $profiles = $namespace.Profiles
    $profile = $null

    foreach ($p in $profiles) {
        if ($p.Name -eq $profileName) {
            $profile = $p
            break
        }
    }

    if ($profile -eq $null) {
        $profile = $profiles.Add($profileName)
    }

    $accounts = $profile.Accounts

    # Agregar la nueva cuenta POP3
    $account = $accounts.Add($EmailAddress)

    # Configurar las propiedades de la cuenta
    $account.DisplayName = $DisplayName
    $account.IncomingMailServer = $POPServer
    $account.IncomingPort = $POPPort
    $account.OutgoingMailServer = $SMTPServer
    $account.OutgoingPort = $SMTPPort
    $account.UserName = $EmailAddress
    $account.Password = $Password

    # Guardar los cambios (¡MUY IMPORTANTE!)
    $profile.Save()

    Write-Host "La cuenta de correo $EmailAddress ha sido configurada exitosamente en el perfil $profileName."

} catch {
    Write-Host "Ocurrió un error al configurar la cuenta de correo: $_"
}

# Evitar que la ventana de PowerShell se cierre automáticamente
Write-Host "Presiona cualquier tecla para cerrar esta ventana..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null