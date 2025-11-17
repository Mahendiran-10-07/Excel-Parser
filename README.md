🤖 Excel Parser Pro (with AI Vision)

This is a powerful Flask web application designed to automate and simplify complex data extraction tasks. It provides a clean web interface for processing unstructured data from various sources (Excel, Images, PDFs) and transforms it into clean, structured, and downloadable .xlsx files using Python, Pandas, and the Google Gemini AI API.

✨ Key Features

This tool provides five distinct data processing modules:

📄 Custom Excel Parser:

Parses complex, non-standard Excel files with unusual layouts.

Handles two different sheet structures ("Sheet1" and "Sheet2") with custom logic to extract data from merged cells, blocks of text, and standard rows.

🖼️ Image Data Extractor (AI-Powered):

Upload one or more images (JPG, PNG) of spreadsheets or tables.

Uses the Google Gemini 1.5 Flash model to perform advanced OCR and data extraction.

The AI analyzes the visual data and returns a structured JSON, which is then converted into a clean Excel table.

📱 WhatsApp Chat Extractor (AI-Powered):

Uploads screenshots of WhatsApp chats or contact lists.

Uses the Gemini AI to extract all text from the images.

Applies a regular expression (regex) to find, clean, and list all phone numbers found in the text.

📊 PDF Data Extractor (AI-Powered):

Uploads a multi-page PDF document.

Uses PyMuPDF to convert each PDF page into a high-resolution image (no external dependencies like Poppler needed).

Sends each page image to the Gemini AI to extract all tabular data.

Intelligently combines the data from all pages into a single, clean Excel file.

✉️ Email Verifier:

--------Work on Process---------

🚀 How to Set Up and Run This Project

Follow these steps to run the application on your local machine.

1. Prerequisites

Python 3.9+

Git

Google Gemini API Key: You must have a valid API key from the Google AI Studio.

2. Local Setup Instructions

1. Clone the Repository:
Open your terminal and clone this project.

git clone [https://github.com/your-username/Excel-Parser.git](https://github.com/your-username/Excel-Parser.git)
cd Excel-Parser


2. Create a Virtual Environment:
It's essential to use a virtual environment to manage dependencies.

python -m venv venv


3. Activate the Virtual Environment:

On Windows (PowerShell):

.\venv\Scripts\Activate


On macOS/Linux:

source venv/bin/activate


4. Install All Required Libraries:
Install all the necessary libraries from the requirements.txt file.

pip install -r requirements.txt


5. Create Your Secret .env File:

Create a new file in the project's root directory named .env

Open the file and add your Google Gemini API key:

GEMINI_API_KEY="AIzaSy...your...key...here..."


The .gitignore file is already set up to keep this file private and secure.

6. Run the Application:
With your virtual environment still active, run the Flask app.

python app.py


The application will be running at http://127.0.0.1:5000. Open this address in your web browser.
