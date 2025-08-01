# intro/views.py
from django.shortcuts import render
from django.urls import reverse

def home_view(request, company_code, store_code):
    store_location = company_code  # Use store_code as store_location
    photo_id = store_code  # Assuming you want to use the same ID for the photo

    # Generate URLs for daily and monthly reports
    daily_url = reverse('upload_with_location_and_photo', args=[store_location, photo_id])  # URL for daily report
    monthly_url = reverse('upload_location_and_photo', args=[store_location, photo_id])  # URL for monthly report

    context = {
        'daily_url': daily_url,
        'monthly_url': monthly_url,
    }
    return render(request, 'intro/intro.html', context)
def upload_and_extract_text(request, store_location, photo_id=None):
    # Your upload logic here, using store_code and photo_id
    return render(request, 'intro/upload.html')  # Example placeholder for the upload template
