REQUISITOS PARA DESPLIEGUE EN SERVIDOR (WEB/CLOUD)
==================================================

Para ejecutar CDO Organizer como un servicio web centralizado con soporte multi-usuario real:

1. ARQUITECTURA
---------------
   - Servidor Central (Cloud/VPS/Intranet):
     - Ejecuta la Interfaz Web (Streamlit) en puerto 8501.
     - Ejecuta la API de Puente (FastAPI) en puerto 8000 (para comunicación con agentes remotos).
     - Aloja la Base de Datos SQLite (cdo_data.db).
   
   - Clientes (Usuarios):
     - Acceden vía Navegador Web (Chrome/Edge) a la IP del servidor.
     - NO requieren instalación local (excepto si necesitan control de hardware local).
   
   - Agentes Locales (Opcional):
     - Si el usuario necesita controlar impresoras locales o descargar archivos a SU equipo desde la web,
       debe descargar y ejecutar el "Agente Ligero" (disponible en la web).

2. REQUISITOS DEL SERVIDOR
--------------------------
   - Sistema Operativo: Windows Server (Recomendado) o Linux (Ubuntu/Debian).
   - Python: Versión 3.10 o superior.
   - Puertos:
     - 8501 (TCP): Para acceso de usuarios (Interfaz Web).
     - 8000 (TCP): Para conexión de Agentes (API).
   - RAM: Mínimo 4GB (8GB recomendados si hay muchos usuarios concurrentes).

3. INDEPENDENCIA DE USUARIOS
----------------------------
   El sistema ya ha sido configurado para garantizar independencia total:
   - Base de Datos: SQLite con bloqueo de hilos (thread-safe) para manejar sesiones concurrentes.
   - Sesiones: Cada pestaña del navegador es una sesión aislada.
   - Archivos: Cada usuario tiene su propia carpeta temporal y configuración.
   - Autenticación: Sistema de Login obligatorio con roles (Admin/User).

4. HERRAMIENTAS PARA EMULAR ENTORNO WEB (LOCAL)
-----------------------------------------------
   Para simular que estás subiendo el aplicativo a un servidor real desde tu PC:

   A. Docker Desktop (Recomendado)
      - Permite crear "contenedores" que simulan un servidor Linux.
      - Instalación: https://www.docker.com/products/docker-desktop/
      - Uso:
        1. Abrir terminal en la carpeta del proyecto.
        2. Ejecutar: docker-compose up --build
        3. Acceder a http://localhost:8501

   B. Ngrok (Para exponer a internet)
      - Permite que personas fuera de tu red accedan a tu servidor local.
      - Instalación: https://ngrok.com/download
      - Uso:
        1. Iniciar el servidor (con Docker o start_server.bat).
        2. Ejecutar en terminal: ngrok http 8501
        3. Ngrok te dará una URL pública (ej. https://x82h...ngrok-free.app) que puedes compartir.

5. INSTRUCCIONES DE PUESTA EN MARCHA (SERVIDOR REAL)
------------------------------------
   1. Copie la carpeta del proyecto al servidor.
   2. Instale Docker y Docker Compose en el servidor.
   3. Ejecute: docker-compose up -d --build
   4. Los usuarios ingresan a: http://IP-SERVIDOR:8501

5. NOTAS DE SEGURIDAD
---------------------
   - Se recomienda usar un Proxy Inverso (Nginx/Apache) con SSL (HTTPS) para producción.
   - Cambie la contraseña por defecto del usuario 'admin'.
