# Despliegue de RIVA / Inlearning en GitHub + Render

Este paquete está preparado para desplegarse con **Render Blueprint** usando `render.yaml`.

## 1. Qué NO debes subir a GitHub

Este paquete ya excluye mediante `.gitignore`:

- `database.db` y cualquier `*.db`
- `base.xlsx` y `base_docentes.xlsx`
- cualquier archivo `*.xlsx` o `*.xlsm`
- carpeta `uploads/`
- archivos `.env`
- entornos virtuales (`venv/`, `.venv/`)

Esto evita publicar DNI, teléfonos, bases de alumnos/docentes o secretos.

## 2. Subir el proyecto a GitHub

Abre CMD o PowerShell dentro de la carpeta del proyecto y ejecuta:

```bash
git init
git add .
git commit -m "Preparar Inlearning RIVA para Render"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git push -u origin main
```

Si el repositorio ya tiene Git configurado, usa solo:

```bash
git add .
git commit -m "Actualizar Inlearning RIVA"
git push
```

## 3. Crear el servicio en Render con Blueprint

1. Entra a Render.
2. Selecciona **New > Blueprint**.
3. Conecta tu cuenta de GitHub.
4. Selecciona el repositorio donde subiste este proyecto.
5. Render detectará automáticamente `render.yaml`.
6. Antes de crear los recursos, completa las variables secretas que Render solicite:
   - `ADMIN_PASSWORD`: contraseña inicial del administrador.
   - `OPERATIONAL_DEFAULT_PASSWORD`: contraseña inicial de las cuentas Seguridad, SAE, Comercial y Pantalla TV.
7. Confirma la creación del Web Service y PostgreSQL.

Render generará automáticamente `SECRET_KEY` y conectará `DATABASE_URL` con PostgreSQL.

## 4. Inicio de sesión

El correo administrativo por defecto es:

```text
admin@idat.edu.pe
```

La contraseña será la que hayas definido en `ADMIN_PASSWORD` al crear el Blueprint.

## 5. Base de datos

En local, la aplicación puede seguir usando SQLite.

En Render, `DATABASE_URL` hace que la aplicación utilice PostgreSQL automáticamente. No debes subir `database.db` a GitHub.

Cuando el sitio esté funcionando, importa alumnos y docentes desde el panel administrativo usando tus Excel locales.

## 6. Comandos usados por Render

Build:

```bash
pip install --upgrade pip && pip install -r requirements.txt
```

Start:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120
```

Health check:

```text
/health
```

## 7. Actualizaciones futuras

Cada vez que hagas cambios:

```bash
git add .
git commit -m "Descripción del cambio"
git push
```

Render volverá a desplegar el proyecto conectado al repositorio.

## Importante sobre el plan gratuito de Render

El Web Service gratuito sirve para pruebas y demostraciones. El PostgreSQL gratuito de Render expira después de 30 días. Para conservar información de visitas de forma permanente, cambia la base de datos a un plan PostgreSQL de pago antes de que expire.

## Corrección de compatibilidad PostgreSQL

Esta versión evita comparar campos INTEGER (`flag_new` y `flag_in_attention`) con cadenas vacías, lo que causaba `psycopg2.errors.InvalidTextRepresentation` en Render/PostgreSQL. Además, `render.yaml` fija `PYTHON_VERSION=3.13.5`.

Si el servicio ya existía antes de este cambio, sincroniza nuevamente el Blueprint o agrega manualmente `PYTHON_VERSION=3.13.5` en Render > Environment y ejecuta **Clear build cache & deploy**.
