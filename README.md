# Workout Transcription Backend

This is the Flask backend for the **Workout Transcription App**. It receives images of handwritten or printed workout logs, uses **Google Cloud Vision API** to extract text (OCR), parses the data, and returns a cleanly formatted **Excel file (.xlsx)**.

---

## Features

- Upload multiple workout log images
- OCR processing using Google Vision API
- Extracts:
  - Workout date and title
  - Exercise names
  - Sets and reps
- Outputs a structured Excel spreadsheet

---

## Tech Stack

- **Flask** – Python web server
- **Google Cloud Vision** – OCR service
- **Pandas** – data processing
- **OpenPyXL** – Excel generation
- **CORS** – allows frontend communication

---

## Project Structure

```
├── app.py                  # Main Flask app
├── google_credentials.json # Google Cloud credentials (not tracked in Git)
├── requirements.txt        # Python dependencies
├── .gitignore              # Git ignore rules
├── uploads/                # Temporary image storage (auto-created)
└── README.md               # Project documentation
```

---

## Getting Started

### 1. Clone the Repo

```bash
git clone https://github.com/yourusername/workout-transcription-backend.git
cd workout-transcription-backend
```

### 2. Set Up Environment

Install dependencies:

```bash
pip install -r requirements.txt
```


### 3. Add Google Vision Credentials

Place your `google_credentials.json` file in the root directory.  
This file must contain your Google Cloud service account key for Vision API.

Also, set the following environment variable:

```bash
export GOOGLE_APPLICATION_CREDENTIALS=google_credentials.json
```

(On Windows: `set GOOGLE_APPLICATION_CREDENTIALS=google_credentials.json`)

---

### 4. Run the Server

```bash
python app.py
```

Server will run locally on `http://127.0.0.1:5000/`

---

## Usage

### Upload Images via Frontend or CURL

Send a `POST` request to `/` with image files under the key `files[]`.

Example CURL:

```bash
curl -F "files[]=@workout1.jpg" -F "files[]=@workout2.jpg" http://127.0.0.1:5000/ --output workout.xlsx
```

---

## Output

Returns a `.xlsx` file where each image becomes a structured workout table with:

- Date and workout title
- Exercise names
- Set/rep breakdown
- Extra info column for notes or unmatched text

---

## Deployment

To deploy this backend:
- Host on **Render**, **Heroku**, or your own server
- Make sure to set up the credentials and CORS for frontend access
- Connect it to your frontend at `/` via POST requests

---

