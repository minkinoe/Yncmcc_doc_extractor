from django.urls import path
from . import views

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('dashboard/<int:file_id>/', views.dashboard, name='dashboard_with_id'),
    path('upload-mark/<int:file_id>/', views.toggle_upload_mark, name='toggle_upload_mark'),
    path('upload-unit/<int:file_id>/', views.update_construction_unit, name='update_construction_unit'),
    path('extracted-construction-order/<int:info_id>/', views.update_construction_order_code, name='update_construction_order_code'),
    path('extracted-construction-email/<int:info_id>/', views.update_construction_email_sent, name='update_construction_email_sent'),
    path('upload/', views.upload_file, name='upload_file'), # Keep for compatibility but redirects
    path('result/', views.show_result, name='show_result'),
    path('history/', views.file_history, name='file_history'),
    path('detail/<int:file_id>/', views.file_detail, name='file_detail'),
    path('download/<int:file_id>/', views.download_file, name='download_file'),
]
