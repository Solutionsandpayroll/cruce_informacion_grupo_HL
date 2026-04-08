import json
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
)

def index(request):
    """
    GET → muestra el formulario de carga.
    POST → procesa los tres archivos y muestra el dashboard con la previsualización.
    """
    if request.method == "GET":
        return render(request, "upload.html")

    # Leer archivos subidos
    reporte_file = request.FILES.get("reporte")
    trabajadores_file = request.FILES.get("trabajadores")
    infra_file = request.FILES.get("infra")

    errors = []
    if not reporte_file:
        errors.append("Falta el archivo Reporte de empleados.")
    if not trabajadores_file:
        errors.append("Falta el archivo Trabajadores Vigentes.")
    if not infra_file:
        errors.append("Falta el archivo INFRA base.")

    if errors:
        return render(request, "upload.html", {"errors": errors})

    # Parsear
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

    # Guardar INFRA en sesión (bytes)
    infra_bytes = infra_file.read()
    request.session["infra_bytes"] = list(infra_bytes)
    request.session["reporte_json"] = df_rep.to_json(
        orient="records", date_format="iso", default_handler=str
    )
    request.session["trabajadores_json"] = df_trab.to_json(
        orient="records", date_format="iso", default_handler=str
    )

    # Previsualización
    cruce_preview = build_cruce_preview(df_trab)
    emp_preview = build_emp_preview(df_rep)

    context = {
        "cruce_preview": cruce_preview[:20],
        "emp_preview": emp_preview[:20],
        "total_trabajadores": len(df_trab),
        "total_empleados": len(df_rep),
        "cruce_cols": list(cruce_preview[0].keys()) if cruce_preview else [],
        "emp_cols": list(emp_preview[0].keys()) if emp_preview else [],
    }
    return render(request, "dashboard.html", context)


@require_POST
def download_infra(request):
    """
    Genera el archivo INFRA con los datos de Reporte y Trabajadores
    inyectados y lo devuelve como descarga.
    """
    import pandas as pd

    infra_bytes_list = request.session.get("infra_bytes")
    reporte_json = request.session.get("reporte_json")
    trabajadores_json = request.session.get("trabajadores_json")

    if not infra_bytes_list or not reporte_json or not trabajadores_json:
        return JsonResponse(
            {"error": "Sesión expirada. Por favor vuelve a cargar los archivos."},
            status=400,
        )

    infra_bytes = bytes(infra_bytes_list)
    df_rep = pd.read_json(io.StringIO(reporte_json), orient="records")
    df_trab = pd.read_json(io.StringIO(trabajadores_json), orient="records")

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