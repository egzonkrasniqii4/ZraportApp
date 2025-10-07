import os
from datetime import datetime
from django.shortcuts import render, redirect
from django.core.files.storage import FileSystemStorage
import pyodbc
from django.utils import timezone
from django.conf import settings

# Update SAVE_PATH to point to the static folder
SAVE_PATH = r''

def upload_and_save_data(request, store_location=None, photo_id=None):
    store_location = str(store_location).zfill(2) if store_location else None

    # Prepare to display the upload form
    error_message = None

    # Define the target folder path, including photo_id as a subfolder if provided
    target_folder = os.path.join(SAVE_PATH, store_location or '', photo_id or '')
    os.makedirs(target_folder, exist_ok=True)

    if request.method == 'POST':
        uploaded_file = request.FILES.get('receipt')
        if uploaded_file:
            # Create timestamped filename
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
            new_filename = f'{store_location}_{photo_id}_{timestamp}_Z.jpg'

            # Save the uploaded file to the target folder
            fs = FileSystemStorage(location=target_folder)
            filename = fs.save(new_filename, uploaded_file)
            uploaded_image_path = os.path.join(target_folder, filename)

            # Extract the relative path (from SAVE_PATH onwards)
            relative_image_path = os.path.relpath(uploaded_image_path, SAVE_PATH).replace(os.sep, '\\')

            # Retrieve form data
            date = request.POST.get('date')
            cash = try_parse_float(request.POST.get('cash'))
            card = try_parse_float(request.POST.get('card'))
            cupon = try_parse_float(request.POST.get('cupon'))

            # Calculate total value if cash and card are provided
            total_value = (cash or 0) + (card or 0) + (cupon or 0)

            # Save to database, including relative_image_path
            try:
                save_to_database(store_location, photo_id, date, total_value, cash, card, cupon, relative_image_path)
                # Redirect to the result page with a GET request
                return redirect('result', store_location=store_location, photo_id=photo_id)
            except Exception as db_error:
                error_message = f"Failed to save data to the database: {db_error}"
        else:
            error_message = "No receipt uploaded."

    return render(request, 'ocr_app/upload.html', {
        'error': error_message,
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

def save_to_database(store_location, location, date, total_value, cash, card, cupon, relative_image_path):
    # Database connection parameters
    conn_str = (
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=;'
        'DATABASE=;'
        'UID=;'
        'PWD='
    )

    # Connect to SQL Server
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Debug output
        print("Location (photo_id):", location)
        print("Date:", date)
        print("Total Value:", total_value)
        print("Cash:", cash)
        print("Card:", card)
        print("Cupon:", cupon)
        print("Image Path:", relative_image_path)

        cupon = cupon if cupon is not None else 0

        # Dynamically select the table based on store_location
        table_name = f"[{store_location}]"

        # Insert data into the dynamically selected table, including relative_image_path
        cursor.execute(f""" 
            INSERT INTO {table_name} (Location, Date, TotalValue, Cash, Card, Cupon, ImageData, ReceivedTime) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", (
            location,
            date, 
            total_value, 
            cash, 
            card, 
            cupon,
            relative_image_path,  # Store the relative path in the database
            timezone.now()
        ))
        conn.commit()
    except Exception as e:
        print("Error inserting data into the database:", str(e))
        raise  # Rethrow the exception to be caught in the upload_and_save_data function
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def result_page(request, store_location, photo_id):
    context = {
        'message': 'Z RAPORTI U DERGUA ME SUKSES!',
        'store_location': store_location,
        'photo_id': photo_id,
    }
    return render(request, 'ocr_app/result.html', context)
