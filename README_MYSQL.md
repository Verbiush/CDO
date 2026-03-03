# Guía de Configuración MySQL para Organizador de Archivos CDO

Este documento describe cómo configurar la aplicación para utilizar una base de datos MySQL en lugar de SQLite (por defecto).

## Prerrequisitos

1.  **Servidor MySQL**: Debe tener acceso a un servidor MySQL (local o remoto).
2.  **Base de Datos**: Debe crear una base de datos vacía (ej: `cdo_organizer`).
3.  **Usuario**: Se recomienda crear un usuario específico con permisos sobre esa base de datos.

## Configuración

La aplicación detecta automáticamente si debe usar MySQL basándose en la presencia de ciertas variables de entorno. Si estas variables están configuradas, la aplicación intentará conectarse a MySQL. Si falta alguna, usará SQLite (`users.db`).

### Variables de Entorno Requeridas

Configure las siguientes variables de entorno en su sistema o contenedor Docker:

-   `MYSQL_HOST`: Dirección IP o hostname del servidor MySQL (ej: `localhost`, `192.168.1.10`, `db-server`).
-   `MYSQL_USER`: Nombre de usuario de MySQL.
-   `MYSQL_PASSWORD`: Contraseña del usuario.
-   `MYSQL_DATABASE`: Nombre de la base de datos a utilizar (ej: `cdo_organizer`).
-   `MYSQL_PORT`: Puerto de conexión (opcional, por defecto `3306`).

### Ejemplo de Ejecución (Windows PowerShell)

```powershell
$env:MYSQL_HOST="localhost"
$env:MYSQL_USER="cdo_user"
$env:MYSQL_PASSWORD="secure_password"
$env:MYSQL_DATABASE="cdo_organizer"

streamlit run src/app_web.py
```

### Ejemplo con Docker Compose

Modifique su archivo `docker-compose.yml`:

```yaml
version: '3.8'

services:
  app:
    image: cdo-organizer
    environment:
      - MYSQL_HOST=db
      - MYSQL_USER=cdo_user
      - MYSQL_PASSWORD=secure_password
      - MYSQL_DATABASE=cdo_organizer
    depends_on:
      - db
    ports:
      - "8501:8501"

  db:
    image: mysql:8.0
    environment:
      - MYSQL_ROOT_PASSWORD=root_password
      - MYSQL_DATABASE=cdo_organizer
      - MYSQL_USER=cdo_user
      - MYSQL_PASSWORD=secure_password
    volumes:
      - db_data:/var/lib/mysql

volumes:
  db_data:
```

## Migración de Datos

### Inicio Automático
Al iniciar la aplicación con las variables de entorno configuradas:
1.  Se crearán automáticamente las tablas necesarias (`users`, `tasks`, `pacientes`, `facturas`, `atenciones`, `document_records`) si no existen.
2.  Si la tabla `users` está vacía, se creará un usuario administrador por defecto (`admin` / `admin123`).
3.  Se aplicarán las actualizaciones de esquema necesarias (versión 2.1).

### Migración de Datos Existentes (SQLite -> MySQL)
Actualmente, la aplicación **no migra automáticamente** los datos existentes de SQLite a MySQL. Inicia con una base de datos MySQL limpia.

Si desea migrar datos, deberá exportar los datos de SQLite a CSV/SQL e importarlos a MySQL manualmente, o utilizar scripts de migración personalizados.
El sistema sí migra automáticamente usuarios desde `users.json` si la tabla `users` está vacía.

## Notas Importantes

-   **Respaldo**: La función de "Descargar Copia de Seguridad" en el panel de administración **no está disponible** para MySQL. Debe realizar los respaldos utilizando herramientas estándar de MySQL (`mysqldump`, MySQL Workbench).
-   **Rendimiento**: La conexión a una base de datos remota puede añadir latencia. Asegúrese de que el servidor de base de datos tenga buena conectividad con el servidor de la aplicación.
-   **Seguridad**: Nunca exponga el puerto 3306 de MySQL directamente a internet. Use VPN o redes privadas.
