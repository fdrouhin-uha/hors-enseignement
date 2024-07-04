from django.contrib import admin
from django.urls import path
from excel_app import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.upload_files, name='upload_files'),
    path('modify/', views.modify_data, name='modify_data'),
    path('save/', views.save_changes, name='save_changes'),
    path('download_files/', views.download_files, name='download_files'),
]
