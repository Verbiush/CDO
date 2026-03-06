# Instrucciones para Instalar y Ejecutar el Agente Local CDO

Este agente permite conectar su máquina local con el servidor central (AWS) para ejecutar tareas de automatización (ej. validación OVIDA, gestión de archivos).

## 1. Requisitos Previos

- Sistema Operativo: Windows 10/11
- Conexión a Internet (para conectar al servidor AWS)

## 2. Instalación

El método recomendado es usar el **Instalador Automático**.

1.  Descargue el archivo `Instalar_Agente.exe` desde la aplicación web (botón "Descargar Agente") o búsquelo en la carpeta raíz del proyecto.
2.  Ejecute `Instalar_Agente.exe`.
3.  El instalador le pedirá:
    *   **Usuario y Contraseña**: Use las mismas credenciales que para la web.
    *   **Carpeta de Instalación**: Por defecto en `AppData`.
4.  El instalador validará sus credenciales con el servidor y configurará el servicio automáticamente.

### Método Manual (Alternativo)

1.  Copie el archivo ejecutable `CDO_Agente.exe` (ubicado en `src/local_agent/dist/`) a la carpeta deseada.
2.  Ejecute el agente y configure `agent_config.json` manualmente como se indica abajo.

## 3. Configuración Inicial (Solo Manual)

El agente necesita saber a qué servidor conectarse.

1.  Ejecute el agente `CDO_Agente.exe` por primera vez.
2.  Se creará una carpeta de configuración en:
    `%LOCALAPPDATA%\CDO_Organizer`
    (Generalmente: `C:\Users\SuUsuario\AppData\Local\CDO_Organizer`)
3.  Abra el archivo `agent_config.json` creado en esa carpeta y edítelo con los siguientes datos:

```json
{
    "server_url": "http://3.142.164.128:8000",
    "username": "su_usuario",
    "password": "su_password"
}
```

**¡IMPORTANTE!**
*   **NO USE el puerto 8501** (ese es para la página web).
*   **USE SIEMPRE el puerto 8000** (ese es para la conexión del agente).
*   URL Correcta: `http://3.142.164.128:8000`

*Nota: Solicite su usuario y contraseña al administrador del sistema.*

## 4. Ejecución

1.  Haga doble clic en `CDO_Agente.exe`.
2.  Se abrirá una ventana de consola mostrando el estado de la conexión.
    - Si ve "Starting Polling Client...", el agente está conectado y esperando tareas.
    - Si ve errores de conexión, verifique la IP y su conexión a internet.

## 5. Uso

El agente funcionará en segundo plano. Cuando el servidor AWS asigne una tarea (ej. validar un lote de archivos), el agente la detectará y ejecutará automáticamente.

- **Logs:** Puede ver los registros de actividad en `%LOCALAPPDATA%\CDO_Organizer\agent.log`.

## 6. Solución de Problemas

- **Error de conexión:** Verifique que el servidor AWS (3.142.164.128) esté accesible y el puerto 8000 esté abierto.
- **Configuración no encontrada:** Asegúrese de haber editado `agent_config.json` correctamente.
