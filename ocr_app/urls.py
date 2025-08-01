#ocr_app

from django.urls import path
from .views import upload_and_save_data
from . import views

urlpatterns = [
    path('upload/', upload_and_save_data, name='upload'),  # URL without store location
    path('upload/<str:store_location>/', upload_and_save_data, name='upload_with_location'),  # URL with store location
    path('upload/<str:store_location>/<str:photo_id>/', upload_and_save_data, name='upload_with_location_and_photo'),
    path('result/<str:store_location>/<str:photo_id>/', views.result_page, name='result'),
  # New URL with store location and photo ID
]
