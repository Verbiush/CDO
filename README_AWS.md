# Guía de Despliegue en AWS (Amazon Web Services)

Esta guía detalla los pasos para desplegar la aplicación "Organizador de Archivos" en un servidor EC2 de AWS utilizando Docker.

## 1. Prerrequisitos

- Cuenta activa en AWS.
- Instancia EC2 lanzada (Recomendado: **Ubuntu 22.04 LTS** o Amazon Linux 2023).
  - Tipo de instancia sugerido: **t3.small** o superior (t2.micro puede quedarse corto de memoria RAM durante la instalación).
  - Espacio en disco: Al menos **20 GB**.
- Grupo de Seguridad (Security Group) configurado para permitir:
  - **SSH (Puerto 22)**: Para conectarse al servidor.
  - **TCP Personalizado (Puerto 8501)**: Para acceder a la aplicación web.

## 2. Preparación del Servidor

Conéctese a su instancia vía SSH y ejecute los siguientes comandos para instalar Docker:

```bash
# Actualizar sistema
sudo apt-get update
sudo apt-get upgrade -y

# Instalar Docker
sudo apt-get install -y docker.io docker-compose
sudo usermod -aG docker $USER
```

*Nota: Después de ejecutar el último comando, desconéctese (`exit`) y vuelva a conectarse para que los cambios de grupo surtan efecto.*

## 3. Instalación de la Aplicación

Puede clonar el repositorio o copiar los archivos manualmente.

### Opción A: Clonar desde GitHub (Recomendado)

```bash
git clone https://github.com/Verbiush/CDO.git
cd CDO
```

### Opción B: Copia Manual
Si no usa Git en el servidor, suba los archivos del proyecto a una carpeta `CDO`.

## 4. Configuración y Despliegue

1. Cree el directorio para la base de datos persistente:
   ```bash
   mkdir -p data
   # Asegúrese de que Docker tenga permisos para escribir
   sudo chmod 777 data
   ```

2. Inicie la aplicación usando Docker Compose:
   ```bash
   docker-compose -f docker-compose-aws.yml up -d --build
   ```

3. Verifique que el contenedor esté corriendo:
   ```bash
   docker ps
   ```

## 5. Acceso

Abra su navegador y vaya a:
`http://<IP-PUBLICA-DE-SU-INSTANCIA>:8501`

## 6. Mantenimiento

### Ver logs
```bash
docker-compose -f docker-compose-aws.yml logs -f
```

### Detener la aplicación
```bash
docker-compose -f docker-compose-aws.yml down
```

### Actualizar la aplicación
Si realiza cambios en el código (GitHub):
```bash
git pull origin main
docker-compose -f docker-compose-aws.yml up -d --build
```

## Notas Importantes

- **Persistencia de Datos**: La base de datos `users.db` se guardará en la carpeta `./data` del servidor. Si borra esta carpeta, perderá los usuarios y registros.
- **Archivos Temporales**: La aplicación limpia archivos temporales automáticamente, pero es recomendable reiniciar el contenedor periódicamente si nota un uso excesivo de disco.
