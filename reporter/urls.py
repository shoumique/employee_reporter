from django.urls import path

from . import views

app_name = "reporter"

urlpatterns = [
    path("", views.upload_view, name="upload"),
    path("configure/", views.configure_view, name="configure"),
    path("export/", views.export_view, name="export"),
    path("export-docx/", views.export_docx_view, name="export_docx"),
    path("api/preset-columns/", views.preset_columns_view, name="preset_columns"),
]
