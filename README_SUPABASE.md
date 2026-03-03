# Guía de Integración con Supabase (PostgreSQL)

Este proyecto ahora soporta conexión a bases de datos PostgreSQL, específicamente diseñado para trabajar con **Supabase**.

## 1. Prerrequisitos

Asegúrate de tener instaladas las dependencias necesarias:

```bash
pip install -r requirements.txt
# O manualmente:
pip install psycopg2-binary
```

## 2. Configuración de Variables de Entorno

Para conectar el proyecto a tu base de datos en Supabase, necesitas configurar las siguientes variables de entorno. Puedes hacerlo en un archivo `.env` en la raíz del proyecto o en las variables de entorno de tu sistema/servidor (AWS/Render/Heroku).

Obtén estos datos desde tu panel de Supabase: `Project Settings` -> `Database` -> `Connection parameters`.

```ini
# Configuración Supabase / PostgreSQL
SUPABASE_HOST=db.xbnkmstqkascpvrrhxrn.supabase.co
SUPABASE_USER=postgres
SUPABASE_PASSWORD=[TU_CONTRASEÑA_DE_BASE_DE_DATOS]
SUPABASE_DB=postgres
SUPABASE_PORT=5432
```

> **Nota:** Reemplaza `[TU_CONTRASEÑA_DE_BASE_DE_DATOS]` con la contraseña que creaste para tu base de datos al iniciar el proyecto en Supabase.

## 3. Migración de Usuarios (Reinicio de Base de Datos)

Si deseas "reiniciar" la base de datos (empezar con tablas limpias) pero conservar tus usuarios actuales desde la base de datos local (`users.db`), ejecuta el script de migración:

```bash
python src/migrate_to_supabase.py
```

**¿Qué hace este script?**
1. Conecta a tu base de datos local SQLite (`src/users.db`).
2. Conecta a tu base de datos Supabase.
3. Crea todas las tablas necesarias en Supabase (Pacientes, Facturas, Usuarios, etc.) si no existen.
4. Copia **SOLO** los usuarios de la base de datos local a Supabase.
5. Las demás tablas (Pacientes, Facturas, etc.) quedarán vacías, logrando el "reinicio" deseado.

## 4. Ejecutar la Aplicación con Supabase

Una vez configuradas las variables de entorno, la aplicación detectará automáticamente la configuración de Supabase y la utilizará en lugar de SQLite o MySQL.

```bash
streamlit run src/app_web.py
```

## 5. Solución de Problemas

- **Error de conexión:** Verifica que la contraseña sea correcta y que el host sea accesible desde tu red.
- **Librería faltante:** Si ves `ModuleNotFoundError: No module named 'psycopg2'`, ejecuta `pip install psycopg2-binary`.
- **Tablas no encontradas:** Ejecuta el script de migración `src/migrate_to_supabase.py` al menos una vez para crear la estructura de tablas.
