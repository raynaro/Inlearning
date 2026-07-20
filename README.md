# PDR · Plataforma de Registro por sedes

Sistema Flask preparado para operar a mayor escala con separación estricta por **instituto, sede y área**.

## Estructura oficial

### IDAT
- IDAT SJM
- IDAT SJL
- IDAT ATE
- IDAT PT
- IDAT TV

### ZEGEL
- ZEGEL SJM
- ZEGEL SJL
- ZEGEL ATE

Cada sede funciona como una unidad independiente para:

- Seguridad
- SAE
- Comercial
- Alumnos y docentes
- Visitas y atenciones
- Solicitudes de registro
- Reportes Excel

## Solicitudes de registro por sede

Seguridad puede crear solicitudes de alta para alumnos o docentes. Cada solicitud se guarda únicamente en la bandeja de la sede del usuario.

El administrador puede:

- Filtrar solicitudes por instituto, sede y estado.
- Aprobar y registrar directamente al alumno o docente en la base de esa sede.
- Rechazar solicitudes.
- Ver quién solicitó y quién revisó.
- Exportar las solicitudes a Excel.

Estados disponibles:

- PENDIENTE
- APROBADA
- RECHAZADA

## Control administrativo

El administrador general puede:

- Ver un consolidado global de las ocho sedes.
- Abrir un panel independiente por sede.
- Gestionar usuarios de Seguridad, SAE y Comercial.
- Cambiar nombre, correo, contraseña, rol, instituto, sede y estado.
- Importar alumnos y docentes sin afectar otras sedes.
- Ver quién registró alumnos, docentes y visitas.
- Ver quién atendió y quién actualizó cada visita.
- Revisar auditoría de cambios.
- Exportar registros, visitas y solicitudes por sede.

## Ejecutar en Windows

```cmd
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir:

```text
http://127.0.0.1:5000
```

## Administrador inicial

```text
Usuario: admin@idat.edu.pe
Contraseña: admin123
```

## Importación de alumnos

Selecciona primero una combinación válida de instituto y sede. El Excel debe contener como mínimo:

- NombreCompleto
- DNI
- Codigo

También puede incluir `INSTITUTO` y `SEDE`.

Valores de sede permitidos:

```text
SJM, SJL, ATE, PT, TV
```

ZEGEL solo admite `SJM`, `SJL` y `ATE`.

## Migración incluida

Los registros anteriores se convierten automáticamente a la nueva estructura. La base existente queda en **IDAT PT**, porque anteriormente pertenecía a Lima Centro. Luego puede redistribuirse importando la base correspondiente en cada sede.

## Producción

El proyecto utiliza SQLite de forma local y admite PostgreSQL mediante `DATABASE_URL`. Incluye índices por instituto, sede, área, estado y fecha.
