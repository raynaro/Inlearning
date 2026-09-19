# RIVA · Registro de Visitas y Asistencias

Sistema Flask para Inlearning con control de Seguridad, SAE, Comercial, turnos, Pantalla TV y reportes diferenciados para IDAT y ZEGEL.

## Funciones principales

- Registro de visitas por IDAT o ZEGEL.
- Gestión por campus: Lima Sur, Lima Este y Lima Norte.
- Usuarios Seguridad, SAE (3 por campus), Comercial (4 por campus) y Pantalla TV (una TV SAE y una TV Comercial por campus).
- Turnos y módulos SAE tipo sistema bancario.
- Estados combinables: Visita nueva + En atención.
- Reportes separados de SAE y Comercial para IDAT y ZEGEL.
- Gráficos comparativos IDAT vs. ZEGEL.
- Panel administrativo y exportación a Excel.

## Ejecutar localmente en Windows

```cmd
py -m venv venv
venv\Scripts\activate
python -m pip install -r requirements.txt
python app.py
```

Abrir:

```text
http://127.0.0.1:5000
```

En desarrollo, si no defines variables de entorno, se mantienen las credenciales locales predeterminadas del proyecto. Para producción usa siempre variables de entorno seguras.

## Base de datos

- Local: SQLite (`database.db`).
- Producción: PostgreSQL mediante `DATABASE_URL`.

Los archivos locales de base de datos y Excel están excluidos de Git mediante `.gitignore` para evitar publicar información personal.

## Desplegar en Render

El repositorio incluye:

- `render.yaml`
- `Procfile`
- `.python-version`
- `.gitignore`
- `.env.example`
- endpoint `/health`

Consulta [DEPLOY_RENDER_GITHUB.md](DEPLOY_RENDER_GITHUB.md) para los pasos completos.

## Disponibilidad SAE y Pantalla TV
Cada asesor SAE puede marcar su módulo como Activo/Inactivo sin desactivar su cuenta. La Pantalla TV refleja únicamente los módulos disponibles y los llamados activos. Los gráficos administrativos excluyen domingos y trabajan de lunes a sábado.

## Pantallas TV por área
Cada cuenta Pantalla TV tiene un área fija (SAE o Comercial). La TV Comercial muestra solo los llamados de los 4 módulos Comercial del campus. Ver `CAMBIOS_RIVA_V13_COMERCIAL_TV.md`.
