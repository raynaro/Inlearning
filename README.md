# RIVA · Registro de Visitas y Asistencias

Sistema Flask para el control de visitas, seguridad, SAE y Comercial.

## Estructura actual

RIVA utiliza dos conceptos visibles:

- **Marca:** IDAT o ZEGEL.
- **Sede física:** Lima Sur, Lima Este o Lima Norte.

Equivalencias internas conservadas para no perder información histórica:

- `SJM` → **Lima Sur**
- `ATE` → **Lima Este**
- `SJL` → **Lima Norte**

IDAT y ZEGEL comparten la misma sede física. Por ejemplo, IDAT SJM y ZEGEL SJM se muestran como **Sede Lima Sur**.

## Seguridad

Cuando Seguridad registra un alumno, docente o una visita, puede seleccionar la **Marca**:

- IDAT
- ZEGEL

La sede queda determinada por la cuenta de Seguridad.

## SAE y Comercial

SAE y Comercial trabajan por sede física consolidada. En una misma bandeja pueden recibir visitas de IDAT y ZEGEL, y cada registro conserva su **Marca** para identificar el origen.

La campana de nuevas visitas considera ambas marcas dentro de la sede asignada.

## Administración

El administrador puede:

- Filtrar por Marca y sede.
- Ver tres sedes consolidadas: Lima Sur, Lima Este y Lima Norte.
- Revisar SAE y Comercial por separado en cada sede.
- Ver la Marca de cada visita.
- Exportar reportes con columnas **Marca** y **Sede**.
- Descargar y volver a importar informes mensuales acumulables.

## Ejecutar en Windows

```cmd
python -m venv venv
venv\Scripts\activate
python -m pip install -r requirements.txt
python app.py
```

Si `python` no funciona, usa `py`.

Abrir en el navegador:

```text
http://127.0.0.1:5000
```

## Administrador inicial

```text
Usuario: admin@idat.edu.pe
Contraseña: admin123
```

## Importación de alumnos y docentes

Selecciona primero la **Marca** y la **Sede**. El Excel puede usar `MARCA` o el encabezado antiguo `INSTITUTO`. Para la sede acepta tanto los nombres visibles como los códigos internos.

Ejemplos válidos:

```text
Marca: IDAT / ZEGEL
Sede: Lima Sur / Lima Este / Lima Norte
```

También siguen siendo reconocidos internamente `SJM`, `ATE` y `SJL` para compatibilidad con archivos anteriores.

## Base de datos

Localmente utiliza SQLite (`database.db`). En producción admite PostgreSQL mediante `DATABASE_URL`.
