# Guía de Actualización y Mantenimiento

Esta guía te ayudará a subir tus cambios a GitHub y actualizar tu servidor AWS con las últimas mejoras (incluyendo el soporte para MySQL).

## 1. Flujo de Trabajo (Resumen)

El ciclo de vida de una actualización es:
1.  **Local**: Haces cambios -> Guardas -> `git add` -> `git commit` -> `git push`.
2.  **Servidor**: Te conectas -> `git pull` -> Reconstruyes contenedores.

---

## 2. Pasos en tu PC Local (Windows)

Cada vez que realices modificaciones en el código o configuración:

### Paso 1: Guardar cambios en Git
Abre la terminal en la carpeta del proyecto (`D:\instalar\OrganizadorArchivos`) y ejecuta:

```powershell
# 1. Ver qué archivos cambiaron
git status

# 2. Añadir todos los cambios (el punto es importante)
git add .

# 3. Guardar los cambios con un mensaje descriptivo
git commit -m "Descripción de lo que hiciste (ej: Agregado soporte MySQL)"

# 4. Subir a GitHub
git push
```

*Si el `git push` te pide credenciales y falla, asegúrate de tener configurado tu acceso a GitHub.*

---

## 3. Pasos en el Servidor AWS

### Paso 1: Conectarse al Servidor
Usa tu cliente SSH (PuTTY, Terminal, etc.) para entrar a tu instancia EC2.

### Paso 2: Descargar los Cambios
Navega a la carpeta del proyecto y descarga la última versión:

```bash
cd CDO
git pull origin main
```

### Paso 3: Actualizar la Aplicación
Dependiendo de qué cambios hiciste, elige la opción adecuada:

#### Opción A: Actualización Completa (Recomendada)
Si cambiaste librerías (`requirements.txt`), configuración de Docker o bases de datos. Esta opción es la más segura.

```bash
# Detener, reconstruir e iniciar en segundo plano
docker-compose -f docker-compose-aws.yml up -d --build
```

#### Opción B: Reinicio Rápido
Si solo cambiaste código Python (`.py`) y NO agregaste nuevas librerías.

```bash
docker-compose -f docker-compose-aws.yml restart app
```

### Paso 4: Verificar
Revisa que todo esté corriendo bien:

```bash
docker ps
# O para ver los logs en tiempo real:
docker-compose -f docker-compose-aws.yml logs -f --tail=50
```

---

## 4. Configuración de MySQL en AWS (Nueva Funcionalidad)

Si deseas activar MySQL en el servidor después de esta actualización:

1.  Asegúrate de tener un servidor MySQL disponible (puede ser RDS o otro contenedor).
2.  Edita el archivo `docker-compose-aws.yml` en el servidor o configura las variables de entorno antes de lanzar el contenedor.

Ejemplo de cómo editar en el servidor con `nano`:
```bash
nano docker-compose-aws.yml
```
Descomenta o agrega las variables de entorno `MYSQL_HOST`, `MYSQL_USER`, etc., en la sección `environment`.

Guarda con `Ctrl+O`, `Enter`, y sal con `Ctrl+X`. Luego ejecuta la **Opción A** del Paso 3.
