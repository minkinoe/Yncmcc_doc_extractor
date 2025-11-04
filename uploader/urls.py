from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_file, name='upload_file'),
    path('result/', views.show_result, name='show_result'),
    path('history/', views.file_history, name='file_history'),
    path('detail/<int:file_id>/', views.file_detail, name='file_detail'),
    path('download/<int:file_id>/', views.download_file, name='download_file'),
]