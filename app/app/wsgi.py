import os
import sys
from django.core.wsgi import get_wsgi_application
from django.core.management import call_command

# Agregar el directorio raíz al path
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if path not in sys.path:
    sys.path.insert(0, path)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.settings')

application = get_wsgi_application()
app = application 

# --- AÑADE ESTO ---
# Ejecuta las migraciones en tiempo de ejecución si la DB no existe en /tmp
try:
    call_command('migrate', interactive=False)
except Exception as e:
    print(f"Error corriendo migraciones en runtime: {e}")