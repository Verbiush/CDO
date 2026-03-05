# Instrucciones para el Agente Local (Modo AWS)

Este instalador permite conectar su **PC Local** con la aplicación alojada en **AWS**, habilitando:
1.  **Acceso a Discos Duros Locales**: Para cargar archivos masivamente desde la nube.
2.  **Ejecución de Navegador Local**: Para automatizaciones que requieren su navegador (Bot Zeus, validadores web).
3.  **Modo Nativo**: Integración fluida entre la Web en AWS y su equipo.

## Pasos de Instalación

1.  Descargue el archivo `Instalador_Agente_CDO.exe` de este repositorio.
2.  Ejecútelo en su equipo local (Windows).
3.  Siga los pasos del asistente.
    *   El instalador configurará automáticamente la conexión al servidor AWS (`3.142.164.128`).
    *   Configurará el agente para **iniciar automáticamente** al encender el equipo.

## Verificación

1.  Una vez instalado, verá el icono del agente o el proceso `CDO_Agente.exe` en el Administrador de Tareas.
2.  Vaya a la aplicación web en AWS.
3.  En la configuración de "Agente Local", debería ver el estado **Conectado**.

## Solución de Problemas

*   **Si no conecta**: Verifique que no tenga firewall bloqueando el puerto 8000 o la salida a Internet.
*   **Si el navegador no abre**: Asegúrese de tener Chrome instalado en su equipo local.
