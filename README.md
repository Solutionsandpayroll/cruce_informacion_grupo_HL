# Como ejecutar el proyecto

## Introducción breve

Este proyecto es una aplicación web en Django para cargar archivos Excel, comparar información de empleados y generar una versión actualizada del archivo INFRA con el cruce de datos.

## Funcionamiento

1. El usuario entra a la ruta raíz `/` y la solicitud llega a `cruce_arl.urls`, que enruta hacia `views.index`.
2. Si la petición es `GET`, la vista muestra el formulario en `app/template/upload.html`.
3. Si la petición es `POST`, `views.index` recibe los archivos `reporte`, `trabajadores` e `infra`, valida que existan y delega el procesamiento a `utils.py`.
4. `utils.parse_reporte` y `utils.parse_trabajadores` leen los Excel, normalizan campos y construyen los datos base del cruce.
5. `utils.build_cruce_preview` y `utils.build_emp_preview` generan la previsualización que se muestra en `app/template/dashboard.html`.
6. El estado temporal se guarda en archivos del sistema temporal con `utils.save_state`, usando una cookie llamada `infra_token` para recuperar la sesión.
7. Cuando el usuario solicita la descarga, la ruta `/download-infra/` llama a `views.download_infra`, que recupera el estado con `utils.load_state` y construye el archivo final con `utils.generate_infra`.
8. La respuesta final es una descarga Excel con el nombre `INFRA_Cruce_ARL_actualizado.xlsx`.

## Estructura tecnica

| Archivo o carpeta | Rol |
| --- | --- |
| `app/manage.py` | Punto de entrada para comandos de Django. |
| `app/app/urls.py` | Enrutamiento principal del proyecto. Incluye `cruce_arl.urls` y `admin/`. |
| `app/cruce_arl/urls.py` | Define las rutas de la app. |
| `app/cruce_arl/views.py` | Endpoints y control de peticiones HTTP. |
| `app/cruce_arl/utils.py` | Lógica reutilizable para lectura, cruce, formateo y generación del archivo Excel. |
| `app/template/upload.html` | Vista de carga inicial de archivos. |
| `app/template/dashboard.html` | Vista de resultados y descarga. |
| `app/static/` | Archivos estáticos servidos por Django/WhiteNoise. |
| `app/app/settings.py` | Configuración del proyecto, templates, estáticos y middleware. |
| `app/app/wsgi.py` | Entrada WSGI para despliegue. |
| `api/index.py` | Entrada usada por Vercel para exponer la app Django. |
| `app/db.sqlite3` | Base SQLite local presente en el proyecto. |
| `app/cruce_arl/models.py` | Archivo presente sin modelos propios definidos. |

En esta arquitectura Django:

- `urls.py` se encarga del enrutamiento.
- `views.py` recibe y responde a las solicitudes.
- `utils.py` concentra la lógica reutilizable separada por funciones.
- `manage.py` es el punto de entrada para comandos de administración.

## Endpoints principales

| Ruta | Función | Propósito |
| --- | --- | --- |
| `/` | `views.index` | Muestra el formulario de carga y procesa los archivos Excel. |
| `/download-infra/` | `views.download_infra` | Genera y descarga el archivo INFRA actualizado. |
| `/admin/` | Django admin | Interfaz administrativa estándar de Django. |

## Requisitos

- Python recomendado: 3.12, de acuerdo con la configuración de despliegue en `vercel.json`.
- Entorno virtual: recomendado.
- Dependencias: `requirements.txt`.
- Variables de entorno obligatorias: no se detectaron variables obligatorias en el código revisado.

## Estructura recomendada de ejecucion

1. Abrir PowerShell en la raíz del proyecto.
2. Crear el entorno virtual:

```powershell
py -3.12 -m venv venv
```

3. Activar el entorno virtual:

```powershell
.\venv\Scripts\Activate.ps1
```

4. Entrar al directorio de la aplicación Django:

```powershell
Set-Location .\app
```

5. Instalar dependencias:

```powershell
pip install -r ..\requirements.txt
```

6. Preparar el proyecto con migraciones y estáticos:

```powershell
python manage.py migrate
python manage.py collectstatic --noinput
```

7. Ejecutar el servidor:

```powershell
python manage.py runserver
```

8. Abrir la aplicación en el navegador:

```text
http://127.0.0.1:8000/
```

## Opcion con script

El archivo `build_files.sh` instala dependencias con `pip`, ejecuta `collectstatic` y luego aplica migraciones. Está pensado para entornos tipo Unix o despliegues automatizados, no para uso directo en PowerShell.

## Notas

- La aplicación usa archivos Excel como entrada y salida.
- El estado temporal se guarda en el directorio temporal del sistema mediante archivos generados por `utils.py`.
- La plantilla base de la interfaz está en `app/template`.
- Los archivos estáticos se sirven desde `app/static`.
- La base SQLite local del proyecto está en `app/db.sqlite3`.
- Para despliegue en Vercel, `vercel.json` apunta a la app WSGI ubicada en `app/app/wsgi.py`.