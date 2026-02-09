# Script para generar certificados SSL para FEV RIPS (Guía Minsalud)
# Basado en: https://gist.github.com/javierllns/0fc9879253a74aa6c031e71c8e2f1158

$ErrorActionPreference = "Stop"

function Check-OpenSSL {
    try {
        $version = openssl version
        Write-Host "✅ OpenSSL detectado: $version" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "❌ OpenSSL no encontrado en el PATH." -ForegroundColor Red
        Write-Host "Por favor, instala OpenSSL para Windows:"
        Write-Host "   https://slproweb.com/products/Win32OpenSSL.html (Win64 OpenSSL v3.x Light)"
        Write-Host "Y asegúrate de agregarlo al PATH durante la instalación."
        return $false
    }
}

if (-not (Check-OpenSSL)) {
    exit
}

# 1. Configurar Ruta
$certPath = Read-Host "Ingrese la ruta para guardar los certificados (Default: C:\Certificates)"
if ([string]::IsNullOrWhiteSpace($certPath)) {
    $certPath = "C:\Certificates"
}

if (-not (Test-Path -Path $certPath)) {
    New-Item -ItemType Directory -Path $certPath | Out-Null
    Write-Host "✅ Carpeta creada: $certPath" -ForegroundColor Green
} else {
    Write-Host "ℹ️ Usando carpeta existente: $certPath" -ForegroundColor Yellow
}

Set-Location -Path $certPath

# 2. Datos del Certificado
$password = Read-Host "Ingrese una contraseña para el certificado (Úsela luego en docker-compose)" -AsSecureString
$passPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

if ([string]::IsNullOrWhiteSpace($passPlain)) {
    Write-Host "❌ La contraseña no puede estar vacía." -ForegroundColor Red
    exit
}

Write-Host "`nGenerando certificados... (Esto puede tardar unos segundos)`n"

try {
    # Paso 1: Generar clave privada
    openssl genrsa -out server.key 2048
    
    # Paso 2: Crear CSR (Configuración automática para evitar preguntas interactivas)
    $configContent = @"
[req]
distinguished_name = req_distinguished_name
prompt = no

[req_distinguished_name]
C = CO
ST = Bolivar
L = Cartagena
O = Compania SAS
OU = Admin
CN = localhost
emailAddress = admin@localhost.com
"@
    $configContent | Out-File -FilePath "openssl.cnf" -Encoding ASCII
    
    openssl req -new -key server.key -out server.csr -config openssl.cnf -passout pass:$passPlain

    # Paso 3: Generar Root CA Key y Certificado
    # Nota: La guía usa una forma simplificada de autofirma. Vamos a seguir la guía.
    # openssl req -x509 -newkey rsa:2048 -keyout rootCA.key -out rootCA.pem -days 825 -nodes -subj "/CN=My Root CA"
    openssl req -x509 -newkey rsa:2048 -keyout rootCA.key -out rootCA.pem -days 825 -nodes -subj "/CN=My Root CA"

    # Paso 4: Firmar el certificado
    openssl x509 -req -in server.csr -CA rootCA.pem -CAkey rootCA.key -CAcreateserial -out server.crt -days 825 -sha256

    # Paso 5: Exportar a PFX
    # Importante: Aquí se usa la contraseña
    openssl pkcs12 -export -out server.pfx -inkey server.key -in server.crt -certfile rootCA.pem -passout pass:$passPlain

    Write-Host "`n✅ ¡Certificados generados exitosamente en $certPath!" -ForegroundColor Green
    Write-Host "-----------------------------------------------------"
    Write-Host "Archivos creados:"
    Get-ChildItem . | Select-Object Name
    Write-Host "-----------------------------------------------------"
    Write-Host "⚠️  PASOS SIGUIENTES:"
    Write-Host "1. Edita tu archivo 'apilocal-dockercompose.Production.yml'."
    Write-Host "2. En 'volumes', asegúrate de tener: - $certPath:/certificates"
    Write-Host "3. En 'environment' -> 'ASPNETCORE_Kestrel__Certificates__Default__Password', pon la contraseña que acabas de elegir."
    Write-Host "4. En 'environment' -> 'ASPNETCORE_Kestrel__Certificates__Default__Path', pon: /certificates/server.pfx"
    Write-Host "   (Nota: Si antes se llamaba fevripsapilocal.pfx, cámbialo a server.pfx o renombra el archivo)"

} catch {
    Write-Host "❌ Error generando certificados: $_" -ForegroundColor Red
}
