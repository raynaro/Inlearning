# Sistema InLearning - Control de ingresos

Proyecto actualizado con estilo InLearning, login profesional compacto, roles separados y panel administrador con gráficos por área.

## Cambios principales
- Se retiró la meta de logros por semana.
- Gráficos separados para SAE y Comercial.
- Gráficos en color amarillo.
- SAE solo ve su apartado de visitas que ingresan y puede exportar sus visitas.
- Comercial solo ve su apartado de visitas que ingresan y puede exportar sus visitas.
- Seguridad solo ve el panel de registros dinámicos: alumno, docente, visita SAE y visita Comercial.
- Administrador mantiene el control completo del sistema.

## Ejecutar en Windows
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir: http://127.0.0.1:5000

## Usuarios de prueba
- Admin: admin@idat.edu.pe / admin123
- SAE: sae@idat.edu.pe / sae123
- Comercial: comercial@idat.edu.pe / comercial123
- Seguridad: seguridad@idat.edu.pe / seguridad123

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```
