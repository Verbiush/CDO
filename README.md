# CDO Organizer & Web Agent

Sistema de gestión de archivos y automatización de procesos médicos, diseñado con una arquitectura híbrida (Web + Agente Local).

## 🚀 Arquitectura del Proyecto

El proyecto consta de dos partes principales:

1.  **Aplicación Web (Streamlit):**
    *   Interfaz de usuario principal.
    *   Manejo de lógica de negocio (RIPS, Validaciones, Bot Zeus).
    *   Ubicación: `src/app_web.py`.

2.  **Agente Local (FastAPI):**
    *   Servicio en segundo plano que permite a la web acceder al sistema de archivos local de forma segura.
    *   Se ejecuta en el puerto `8989`.
    *   Ubicación: `src/local_agent/main.py`.

## 📂 Estructura de Carpetas

```
OrganizadorArchivos/
├── src/
│   ├── app_web.py           # Punto de entrada Web
│   ├── local_agent/         # Código del Agente Local
│   │   ├── main.py
│   │   └── setup_agent.py
│   ├── modules/             # Lógica de negocio (Validadores, Procesadores)
│   └── run_native.py        # Lanzador de escritorio (Wrapper)
├── assets/                  # Imágenes y recursos estáticos
├── requirements.txt         # Dependencias Python
└── build_release.py         # Script de compilación a .exe
```

## ☁️ Despliegue en Nube (Testing)

Para montar el proyecto en un servidor gratuito para pruebas, se recomienda **Streamlit Community Cloud**.

### Pasos para desplegar:

1.  Subir este código a un repositorio de **GitHub**.
2.  Registrarse en [share.streamlit.io](https://share.streamlit.io/).
3.  Crear una "New App" y seleccionar el repositorio.
4.  Apuntar el "Main file path" a: `src/app_web.py`.

### ⚠️ Nota Importante sobre la Nube

Al desplegar en la nube, la **comunicación con el Agente Local no funcionará** por defecto.
*   La web estará en un servidor de EE.UU.
*   El agente estará en tu PC.
*   La web intentará buscar `localhost:8989` y buscará dentro del servidor de EE.UU., no en tu PC.

Para pruebas en la nube, solo funcionarán los módulos que no requieran acceso directo al disco duro (ej: Validadores de carga manual, Bots que no dependan de rutas locales absolutas).

## 🛠️ Desarrollo Local

1.  Instalar dependencias:
    ```bash
    pip install -r requirements.txt
    ```
2.  Ejecutar la web:
    ```bash
    streamlit run src/app_web.py
    ```
3.  Ejecutar el agente (en otra terminal):
    ```bash
    python src/local_agent/main.py
    ```
