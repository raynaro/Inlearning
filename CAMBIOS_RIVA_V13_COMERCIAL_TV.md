# RIVA V13 - 4 módulos Comercial y TV Comercial por campus

## Usuarios
- Cada campus (Lima Sur SJM, Lima Este ATE, Lima Norte SJL) tiene ahora **4 usuarios Comercial** (Módulo 1 a 4):
  `comercial.<campus>.01@pdr.local` ... `comercial.<campus>.04@pdr.local` (códigos `COM-<campus>-01..04`).
- Cada campus tiene **dos pantallas TV separadas**:
  - TV SAE: `pantalla.<campus>@pdr.local` (código `TV-<campus>-01`), muestra solo llamados de SAE.
  - TV Comercial (nueva): `pantalla.comercial.<campus>@pdr.local` (código `TVC-<campus>-01`), muestra solo llamados de Comercial.
- Contraseña inicial de las cuentas nuevas: la variable `OPERATIONAL_DEFAULT_PASSWORD` (por defecto `Riva1234`). Cambiarla en producción.

## Migración (bases ya desplegadas)
- Nueva migración `inlearning_users_v4`, se ejecuta una sola vez al iniciar la app.
- Solo agrega lo que falta: no borra, no reactiva cuentas desactivadas ni cambia contraseñas existentes.
- Agrega la columna `users.tv_area` (compatible SQLite y PostgreSQL). Las pantallas TV existentes quedan como `SAE`.

## Regla de la pantalla TV
- Una cuenta TV tiene un área fija (`tv_area`: SAE o COMERCIAL). Ya no depende de `?area=` en la URL:
  la TV SAE no puede mostrar llamados de Comercial ni al revés, aunque se edite la URL.
- El Administrador sigue eligiendo campus y área (selector "TV SAE / TV Comercial" en la barra superior de la TV
  y enlaces "Probar TV SAE / TV Comercial" en Administración TV).
- La TV muestra junto al campus el área (p. ej. "Lima Sur · Comercial") para identificar rápido cada pantalla.

## Administración de usuarios
- Al crear/editar un usuario con rol Pantalla TV se elige el **Área de la pantalla TV** (SAE o Comercial).
- El listado de usuarios indica el área de cada pantalla TV y su botón "Abrir" abre la TV del área correcta.

## Corrección
- El control "Ya tienes un visitante en atención" de Comercial consultaba el área SAE, por lo que nunca bloqueaba.
  Ahora consulta COMERCIAL: cada asesor Comercial atiende un visitante a la vez, y su módulo se muestra
  correctamente en la TV Comercial.

## Módulos en pantalla
- La TV muestra tantos módulos como asesores activos tenga el campus en su área: **3 en la TV SAE y 4 en la TV Comercial** (antes había 4 módulos fijos en ambas).
- La grilla se ajusta a la cantidad de módulos y se actualiza sola si se agrega o desactiva un asesor.
