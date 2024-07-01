from django.urls import path
from . import views

urlpatterns = [
    path('accueil/', views.accueil, name='accueil'),
    path('home/', views.home, name='home'),
    path('fichor/', views.fichor, name='fichor'),
    path('display_fichor/', views.display_fichor, name='display_fichor'),
    path('download_fich/', views.download_fich, name='download_fich'),
    path('list_ens/', views.list_ens, name='list_ens'),
    path('display_list/', views.display_list, name='display_list'),
 ]