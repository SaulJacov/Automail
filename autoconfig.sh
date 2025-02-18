#!/bin/bash
# Variables
PROFILE_NAME="PerfilCorreo"
EMAIL="soporte5@soluciona.com.mx"
POP3_SERVER="mail.soluciona.com.mx"
SMTP_SERVER="mail.soluciona.com.mx"
POP3_PORT="110"
SMTP_PORT="587"
# Crear el archivo .prf
cat <<EOF > configuracion_outlook.prf
[General]
Custom=1
ProfileName=$PROFILE_NAME
OverwriteProfile=Yes

[Service1]
UniqueService=No
AccountName=$PROFILE_NAME
EmailAddress=$EMAIL
POP3Server=$POP3_SERVER
POP3UserName=$EMAIL
POP3UseSPA=0
POP3Port=$POP3_PORT
POP3UseSSL=0
SMTPServer=$SMTP_SERVER
SMTPUseAuth=1
SMTPAuthUserName=$EMAIL
SMTPUseSPA=0
SMTPPort=$SMTP_PORT
SMTPUseSSL=0
SMTPUseTLS=0
EOF

echo "Archivo configuracion_outlook.prf creado exitosamente."