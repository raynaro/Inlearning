# RIVA · Ajuste final SAE, TV y gráficos

## 1. Estados de visita simplificados
- Cada visita vuelve a usar un único selector con tres opciones: Visita nueva, En atención y Atendido.
- Se elimina de la interfaz la selección múltiple de estados por registro.
- La selección múltiple queda únicamente en los filtros de la lista.
- Los filtros permiten combinar Nuevas + En atención, En atención + Atendidas, Nuevas + Atendidas o las tres.

## 2. Disponibilidad operativa SAE
- Se agrega `service_active` a usuarios para separar disponibilidad de atención del estado de la cuenta.
- Cada usuario SAE conserva un módulo fijo según su orden de cuenta dentro del campus.
- El usuario SAE puede cambiar entre Activo e Inactivo desde su panel.
- Un SAE no puede ponerse Inactivo mientras tenga una visita en atención.
- Un SAE Inactivo no puede tomar una nueva visita.
- Cada SAE puede mantener solo una persona en atención simultáneamente.

## 3. Pantalla TV
- La TV muestra cuántos módulos SAE están activos respecto del total.
- Muestra individualmente Módulo 1, 2 y 3 como Activo o Inactivo según corresponda.
- Los llamados activos solo se muestran si el asesor SAE responsable está operativo.
- Se mantienen ocultos DNI, teléfono y marca del visitante.

## 4. Gráficos administrativos
- Todos los gráficos de actividad usan únicamente lunes a sábado.
- Los domingos quedan excluidos de las series y de los totales mostrados en los gráficos.
- La vista semanal se identifica como `Lun–Sáb`.
- La vista larga muestra 30 días operativos, excluyendo domingos.
- Se conservan las líneas comparativas de IDAT y ZEGEL.

## 5. Render / PostgreSQL
- Se conserva la corrección PostgreSQL para columnas INTEGER.
- Se mantiene `PYTHON_VERSION=3.13.5` en Render.
- La migración de `service_active` es compatible con bases PostgreSQL ya existentes.
