import uuid
import io
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_POST

from .utils import (
    parse_reporte,
    parse_trabajadores,
    build_cruce_preview,
    build_emp_preview,
    generate_infra,
    save_state,
    load_state,
    delete_state,
    _TOKEN_COOKIE,
)


def index(request):
    """
    GET  → muestra el formulario de carga.
    POST → procesa los tres archivos y muestra el dashboard con la previsualización.
    """
    if request.method == "GET":
        return render(request, "upload.html")

    # ── Leer archivos subidos ──────────────────────────────────────
    reporte_file     = request.FILES.get("reporte")
    trabajadores_file = request.FILES.get("trabajadores")
    infra_file        = request.FILES.get("infra")

    errors = []
    if not reporte_file:
        errors.append("Falta el archivo Reporte de empleados.")
    if not trabajadores_file:
        errors.append("Falta el archivo Trabajadores Vigentes.")
    if not infra_file:
        errors.append("Falta el archivo INFRA base.")

    if errors:
        return render(request, "upload.html", {"errors": errors})

    # ── Parsear ────────────────────────────────────────────────────
    try:
        df_rep = parse_reporte(reporte_file)
    except Exception as e:
        return render(request, "upload.html",
                      {"errors": [f"Error leyendo Reporte: {e}"]})

    try:
        df_trab = parse_trabajadores(trabajadores_file)
    except Exception as e:
        return render(request, "upload.html",
                      {"errors": [f"Error leyendo Trabajadores: {e}"]})

    # ── Persistir estado en archivos temporales ────────────────────
    infra_bytes = infra_file.read()

    # Reutilizar token existente o crear uno nuevo
    token = request.COOKIES.get(_TOKEN_COOKIE) or str(uuid.uuid4())
    save_state(token, df_rep, df_trab, infra_bytes)

    # ── Previsualización ───────────────────────────────────────────
    cruce_preview = build_cruce_preview(df_trab)
    emp_preview   = build_emp_preview(df_rep)

    context = {
        "cruce_preview":      cruce_preview[:20],
        "emp_preview":        emp_preview[:20],
        "total_trabajadores": len(df_trab),
        "total_empleados":    len(df_rep),
        "cruce_cols": list(cruce_preview[0].keys()) if cruce_preview else [],
        "emp_cols":   list(emp_preview[0].keys())   if emp_preview   else [],
    }

    response = render(request, "dashboard.html", context)
    # Cookie sin expiración → dura hasta que el navegador se cierre
    response.set_cookie(_TOKEN_COOKIE, token, httponly=True, samesite="Lax")
    return response


@require_POST
def download_infra(request):
    """
    Genera el archivo INFRA con los datos de Reporte y Trabajadores
    inyectados y lo devuelve como descarga.
    """
    token = request.COOKIES.get(_TOKEN_COOKIE)
    if not token:
        return JsonResponse(
            {"error": "No se encontró la sesión. Por favor vuelve a cargar los archivos."},
            status=400,
        )

    try:
        df_rep, df_trab, infra_bytes = load_state(token)
    except FileNotFoundError:
        return JsonResponse(
            {"error": "Los archivos temporales expiraron o no existen. "
                      "Por favor vuelve a cargar los archivos."},
            status=400,
        )
    except Exception as e:
        return JsonResponse({"error": f"Error cargando datos: {e}"}, status=500)

    try:
        output_bytes = generate_infra(infra_bytes, df_rep, df_trab)
    except Exception as e:
        return JsonResponse({"error": f"Error generando archivo: {e}"}, status=500)

    response = HttpResponse(
        output_bytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="INFRA_Cruce_ARL_actualizado.xlsx"'
    return response