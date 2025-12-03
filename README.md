# AI-Powered Document Extraction Dashboard

This project is an AI-powered tool that extracts structured data from PDF documents and converts it into an Excel format. It features a modern, user-friendly web dashboard for easy file uploading and data visualization.

## Features

*   **AI Extraction:** Uses Google's Gemini AI to intelligently extract key-value pairs and contextual comments from unstructured PDF text.
*   **Web Dashboard:** A professional, dark-themed web interface built with Flask.
*   **Drag & Drop:** Easy file upload support.
*   **Data Preview:** View the extracted data in a clean table format before downloading.
*   **Excel Export:** Download the structured data as a formatted Excel file (`.xlsx`).
*   **Smart Deduplication:** Automatically handles duplicate comments by referencing the original row.

## Prerequisites

*   Python 3.8 or higher
*   A Google Gemini API Key

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repo-url>
    cd <your-repo-name>
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Set up your API Key:**
    *   Create a file named `api key.txt` in the root directory.
    *   Paste your Google Gemini API key into this file.

## Usage

1.  **Start the application:**
    ```bash
    python dashboard/app.py
    ```

2.  **Open the dashboard:**
    *   Open your web browser and navigate to `http://127.0.0.1:5000`.

3.  **Process a file:**
    *   Enter your API Key (if not already loaded).
    *   Drag and drop a PDF file (e.g., `Data Input.pdf`) into the upload area.
    *   Click "Process".
    *   Wait for the extraction to complete.
    *   View the results and click "Download Excel" to save the file.

## Project Structure

*   `dashboard/`: Contains the Flask web application code.
    *   `app.py`: The main Flask server.
    *   `engine.py`: Core logic for PDF extraction, AI processing, and Excel generation.
    *   `static/`: CSS and JavaScript files.
    *   `templates/`: HTML templates.
*   `requirements.txt`: List of Python dependencies.
*   `README.md`: This documentation file.

## Technologies Used

*   **Backend:** Python, Flask
*   **AI Model:** Google Gemini (via `google-generativeai`)
*   **PDF Processing:** `pdfplumber`
*   **Data Handling:** `pandas`, `openpyxl`
*   **Frontend:** HTML5, CSS3, JavaScript

## License

[Choose a license, e.g., MIT]
