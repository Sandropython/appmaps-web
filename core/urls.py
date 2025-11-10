from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", views.index, name="index"),        # raiz /
    path("healthz", views.healthz, name="healthz"),
]
