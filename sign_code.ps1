
param (
    [string]$TargetFile = ""
)

# Configuración
$certSubject = "CN=CDO_Organizer_SelfSigned"

Write-Host "Iniciando proceso de firmado de código..." -ForegroundColor Cyan

# 1. Buscar o Crear Certificado
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -eq $certSubject } | Select-Object -First 1

if (-not $cert) {
    Write-Host "Creando nuevo certificado autofirmado: $certSubject..." -ForegroundColor Yellow
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $certSubject -CertStoreLocation Cert:\CurrentUser\My
    Write-Host "Certificado creado exitosamente: $($cert.Thumbprint)" -ForegroundColor Green
} else {
    Write-Host "Usando certificado existente: $($cert.Thumbprint)" -ForegroundColor Green
}

# Función para firmar un archivo
function Sign-File {
    param ($path)
    if (Test-Path $path) {
        Write-Host "Firmando: $path ..." -ForegroundColor Yellow
        try {
            Set-AuthenticodeSignature -FilePath $path -Certificate $cert -ErrorAction Stop
            Write-Host "¡Firmado correctamente!" -ForegroundColor Green
        } catch {
            Write-Host "Error al firmar $path : $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Archivo no encontrado: $path" -ForegroundColor Red
    }
}

# 2. Lógica de Firmado
if ($TargetFile -ne "") {
    # Modo Directo (llamado desde Python/Build)
    Sign-File -path $TargetFile
} else {
    # Modo Batch (Por defecto, busca en dist/)
    $distPath = "$PSScriptRoot\dist"
    $executables = @("CDO_Cliente.exe", "CDO_Instalador.exe", "Instalador_CDO.exe")
    
    foreach ($exe in $executables) {
        $exePath = Join-Path $distPath $exe
        if (Test-Path $exePath) {
            Sign-File -path $exePath
        }
    }
}
