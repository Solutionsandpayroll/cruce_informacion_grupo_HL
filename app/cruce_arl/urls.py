from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("download-infra/", views.download_infra, name="download_infra"),  # Cambiado de "descargar/" a "download-infra/"
]