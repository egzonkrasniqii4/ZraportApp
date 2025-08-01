# z_report_monthly/urls.py
from django.urls import path
from .views import upload_and_extract_text

urlpatterns = [
    path('upload/', upload_and_extract_text, name='upload'),  # URL without store location
    path('upload/<str:store_location>/', upload_and_extract_text, name='upload_with_location'),  # URL with store location
    path('upload/<str:store_location>/<str:photo_id>/', upload_and_extract_text, name='upload_location_and_photo'),  # URL with store location and photo ID
]
