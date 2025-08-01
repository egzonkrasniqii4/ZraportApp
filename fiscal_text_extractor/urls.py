# fiscal_text_extractor/urls.py

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('z/', include('ocr_app.urls')),  # Daily report URLs
    path('zMonthly/', include('z_report_monthly.urls')),  # Monthly report URLs
    path('', include('intro.urls')),  # Home URLs
    path('Administrata/', include('administrata.urls')),  # Administrata
]

# Serve static files during development
if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
