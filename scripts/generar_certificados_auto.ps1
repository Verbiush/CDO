# Script AUTOMÁTICO para generar certificados SSL para FEV RIPS
# No requiere interacción del usuario. Usa contraseña por defecto "fevrips2024*" si no se provee.

param (
    [string]$Password = "fevrips2024*",
    [string]$CertPath = "C:\Certificates"
)

$ErrorActionPreference = "Stop"

function Check-OpenSSL {
    try {
        $version = openssl version
        Write-Host "✅ OpenSSL detectado en PATH: $version" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "⚠️ OpenSSL no encontrado en el PATH. Buscando en ubicaciones comunes..." -ForegroundColor Yellow
        
        # Lista de rutas comunes de Git/OpenSSL
        $commonPaths = @(
            "C:\Program Files\Git\usr\bin",
            "C:\Program Files\Git\mingw64\bin",
            "C:\Program Files\Git\bin",
            "C:\Program Files (x86)\Git\usr\bin",
            "C:\Program Files (x86)\Git\mingw64\bin",
            "C:\OpenSSL-Win64\bin"
        )
        
        foreach ($path in $commonPaths) {
            if (Test-Path "$path\openssl.exe") {
                Write-Host "✅ OpenSSL encontrado en: $path" -ForegroundColor Green
                $env:Path += ";$path"
                return $true
            }
        }
        
        Write-Host "❌ OpenSSL no encontrado." -ForegroundColor Red
        return $false
    }
}

if (-not (Check-OpenSSL)) {
    Write-Host "Instalando OpenSSL light..."
    # Aquí podríamos intentar instalarlo o simplemente fallar.
    # Por ahora, fallamos con mensaje claro.
    exit 1
}

# 1. Crear carpeta
if (-not (Test-Path -Path $CertPath)) {
    New-Item -ItemType Directory -Path $CertPath | Out-Null
    Write-Host "✅ Carpeta creada: $CertPath" -ForegroundColor Green
}

Set-Location -Path $CertPath

# 2. Generar Certificados
Write-Host "Generando certificados en $CertPath..."

try {
    # Paso 1: Generar clave privada
    openssl genrsa -out server.key 2048
    
    # Paso 2: Crear CSR
    $configContent = @"
[req]
distinguished_name = req_distinguished_name
prompt = no

[req_distinguished_name]
C = CO
ST = Bogota
L = Bogota
O = Minsalud
OU = FEVRIPS
CN = localhost
emailAddress = admin@localhost.com
"@
    $configContent | Out-File -FilePath "openssl.cnf" -Encoding ASCII
    
    openssl req -new -key server.key -out server.csr -config openssl.cnf -passout pass:$Password

    # Paso 3: Generar Root CA Key y Certificado
    openssl req -x509 -newkey rsa:2048 -keyout rootCA.key -out rootCA.pem -days 3650 -nodes -subj "/CN=FEVRIPS Local CA"

    # Paso 4: Firmar el certificado
    openssl x509 -req -in server.csr -CA rootCA.pem -CAkey rootCA.key -CAcreateserial -out server.crt -days 3650 -sha256

    # Paso 5: Exportar a PFX
    # Importante: Aquí se usa la contraseña
    openssl pkcs12 -export -out server.pfx -inkey server.key -in server.crt -certfile rootCA.pem -passout pass:$Password

    Write-Host "✅ CERTIFICADOS GENERADOS EXITOSAMENTE" -ForegroundColor Green
    Write-Host "Ruta: $CertPath\server.pfx"
    Write-Host "Contraseña: $Password"
} catch {
    Write-Host "❌ Error: $_" -ForegroundColor Red
    exit 1
}
