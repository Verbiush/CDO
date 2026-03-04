#!/bin/bash
mkdir -p ./certificates
openssl req -x509 -newkey rsa:4096 -keyout ./certificates/server.key -out ./certificates/server.crt -days 365 -nodes -subj "/CN=localhost"
openssl pkcs12 -export -out ./certificates/server.pfx -inkey ./certificates/server.key -in ./certificates/server.crt -passout pass:fevrips2024*
chmod 644 ./certificates/server.pfx
echo "Certificados generados en ./certificates"
