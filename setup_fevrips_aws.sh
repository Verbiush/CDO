#!/bin/bash
echo "--- Configuración de FEVRIPS en AWS ---"

# 1. Generar certificados
echo "1. Generando certificados SSL..."
if [ -f "./generate_certs.sh" ]; then
    chmod +x ./generate_certs.sh
    ./generate_certs.sh
else
    echo "Creando script de certificados..."
    mkdir -p ./certificates
    openssl req -x509 -newkey rsa:4096 -keyout ./certificates/server.key -out ./certificates/server.crt -days 365 -nodes -subj "/CN=localhost"
    openssl pkcs12 -export -out ./certificates/server.pfx -inkey ./certificates/server.key -in ./certificates/server.crt -passout pass:fevrips2024*
    chmod 644 ./certificates/server.pfx
    echo "Certificados generados en ./certificates"
fi

# 2. Instrucciones de Login
echo ""
echo "2. IMPORTANTE: Autenticación en Azure Container Registry"
echo "Ejecute el siguiente comando manualmente antes de iniciar Docker:"
echo "----------------------------------------------------------------"
echo "docker login -u puller -p v1GLVFn6pWoNrQWgEzmx7MYsf1r7TKJQo+kwadvffq+ACRA3mLxs fevripsacr.azurecr.io"
echo "----------------------------------------------------------------"

echo ""
echo "3. Iniciar Servicios"
echo "docker-compose -f docker-compose-aws.yml up -d --build"
