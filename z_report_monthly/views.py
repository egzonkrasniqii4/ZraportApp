import os
from datetime import datetime
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
import pyodbc
from django.utils import timezone
from django.conf import settings

# Define the static folder path for saving images
SAVE_PATH = r'\\192.168.49.3\FashionApps\static\SCANAPP'

def upload_and_extract_text(request, store_location=None, photo_id=None):
    # Ensure store_location is properly formatted as two digits if provided
    if store_location:
        store_location = str(store_location).zfill(2)  # Ensure it is 2 digits (e.g., '01', '02')

    # Prepare to display the upload form
    error_message = None
    uploaded_successfully = False

    # Define the target folder path, including photo_id as a subfolder if provided
    target_folder = os.path.join(SAVE_PATH, store_location or '', photo_id or '')

    # Create the directories if they don't exist
    os.makedirs(target_folder, exist_ok=True)

    if request.method == 'POST':
        # Check if the receipt file is present
        uploaded_file = request.FILES.get('receipt')
        if uploaded_file:
            # Get the current timestamp
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')

            # Create the new filename using photo_id if available
            new_filename = f'ZPeriod_{store_location}_{photo_id}_{timestamp}.jpg'

            # Save the uploaded file to the target folder
            fs = FileSystemStorage(location=target_folder)
            filename = fs.save(new_filename, uploaded_file)
            uploaded_image_path = os.path.join(target_folder, filename)

            relative_image_path = os.path.relpath(uploaded_image_path, SAVE_PATH).replace(os.sep, '\\')

            # # Retrieve form data
            # total_value = request.POST.get('total_value')

            # # Convert numeric fields safely
            # total_value = try_parse_float(total_value)

            # Save to database using photo_id as location
            if save_to_database(store_location, photo_id,  uploaded_file, relative_image_path):
                uploaded_successfully = True

                # Render the result.html template
                return render(request, 'z_report_monthly/result.html', {
                    'message': 'RAPORTI PERIODIK U DERGUA ME SUKSES!',
                    'store_location': store_location,
                    'photo_id': photo_id,
                })
            else:
                error_message = "Failed to save data to the database."
        else:
            error_message = "No receipt uploaded."

    return render(request, 'z_report_monthly/upload.html', {
        'error': error_message,
        'uploaded_successfully': uploaded_successfully,
        'store_location': store_location,
        'photo_id': photo_id,
    })

def try_parse_float(value):
    """Safely convert a value to float, returning None if conversion fails."""
    try:
        return float(value) if value else None
    except ValueError:
        print(f"Error converting value {value} to float.")
        return None

def save_to_database(store_location, photo_id, uploaded_file, relative_image_path):
    # Database connection parameters
    conn_str = (
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=192.168.49.49;'
        'DATABASE=ZRaportsApp;'
        'UID=sa;'
        'PWD=sasa'
    )
    
    # Connect to SQL Server
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Debug output
        print("Store Location (Company):", store_location)  # Debugging store location
        print("Photo ID:", photo_id)

        # Get the previous month and year
        now = timezone.now()
        previous_month = now.month - 1 if now.month > 1 else 12
        year = now.year if previous_month != 12 else now.year - 1

        # Insert data into the ZPeriod table
        cursor.execute(""" 
            INSERT INTO ZPeriod (Company, Location, ImageData, ReceivedTime, ReceivedMonth, ReceivedYear) 
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            store_location,  # Use store_location for Company
            photo_id,  # Use photo_id for Location
            relative_image_path,  # Save image path as relative path
            now,  # Store the current time for ReceivedTime
            previous_month,  # Insert the previous month
            year  # Insert the year for the previous month
        ))
        conn.commit()
        return True  # Indicate success
    except Exception as e:
        print("Error inserting data into the database:", str(e))
        return False  # Indicate failure
    finally:
        cursor.close()
        conn.close()


