# intro/urls.py
from django.urls import path
from .views import home_view, upload_and_extract_text

urlpatterns = [
    path('home/<str:company_code>/<str:store_code>/', home_view, name='home_view'),  # Home path
    path('z/upload/<str:store_location>/<str:photo_id>/', upload_and_extract_text, name='upload_with_location_and_photo'),  # Monthly report
    path('upload/<str:store_location>/', upload_and_extract_text, name='upload_with_location'),  # Daily report
]
