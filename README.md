# Control de Asistencia y Tareo PRIZE PRO

Aplicativo web Flask listo para GitHub y Render.

## Funciones
- Login admin/operador.
- Carga de trabajadores por Excel.
- Captura de fotocheck QR/código de barras por cámara o lector USB.
- Marcación de entrada/salida con foto evidencia opcional y geolocalización si el navegador permite.
- Registro de tareos por labor, fundo, lote, horas, cantidad y unidad.
- Dashboard y exportación Excel.
- PWA básica para instalar como app en celular.

## Usuario inicial
- Usuario: `admin`
- Clave: `admin123`

Cambiar la clave luego del primer despliegue.

## Render
Variables recomendadas:
- `SECRET_KEY`: una clave segura.
- `DATABASE_URL`: PostgreSQL de Render para persistencia real.
- `PERSIST_DIR`: opcional. En Render Free no depender de archivos locales.

Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```

## Plantilla trabajadores
Desde el sistema: **Cargar base > Plantilla Excel**.
El QR/código de barras del fotocheck debe contener el DNI de 8 dígitos.
