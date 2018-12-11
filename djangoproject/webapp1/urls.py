
from django.urls import path, include
from . import views
from django.views.generic import TemplateView

urlpatterns = [
    path('hi/', views.index, name="index"),
    path('start/', views.start_app, name="start_app"),
    path('start/next/', views.next, name="next")

]
