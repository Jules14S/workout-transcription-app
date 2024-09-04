# Workout Transcription Backend

This backend processes images of workout logs using Google Cloud Vision's OCR API and converts them into structured Excel files. The app is built using Flask and 
handles multiple image uploads, extracting data such as date, exercise name, sets, and reps.

## Technologies Used

- **Flask**: A lightweight WSGI web application framework.
- **Google Cloud Vision API**: Used for OCR to extract text from images.
- **Pandas**: Used to structure and handle data.
- **Openpyxl**: Used to generate Excel files.
- **CORS**: Enabled to allow cross-origin requests from the frontend.
  
## Features

- Users can upload multiple images at once.
- Images are processed and converted to Excel tables.
- Each table contains the date, exercise name, sets, reps, and additional info.
- Excel file is returned with each table on a signel sheet.

      ```
## Project Structure

```bash
├── app.py                # Main Flask application
├── requirements.txt      # Python dependencies
├── uploads/              # Directory where uploaded images are temporarily saved
└── README.md             # This file

