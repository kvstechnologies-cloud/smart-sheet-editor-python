# Smart Sheet Editor Python

This simple Flask application allows uploading Excel `.xlsx` files from the browser and returns their contents as JSON. The app reads every sheet in the workbook using `pandas` and `openpyxl`.

## Setup

1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the application:
   ```bash
   python app.py
   ```
3. Open `http://localhost:5000/` in the browser to test uploading a file.

Uploaded files are stored temporarily in the `uploads/` directory and deleted immediately after being processed.
