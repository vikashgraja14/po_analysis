"""
This script initializes a Flask web application for handling file uploads, converting PDFs to images,
and managing an SQLite database. It includes the following components:

1. Imports:
   - base64: Encoding and decoding binary data using base64 encoding.
   - cv2 (OpenCV): Computer vision library for image processing tasks.
   - fitz: Python wrapper for the MuPDF library (PDF handling).
   - io: Core tools for working with I/O operations.
   - os: Interacting with the operating system (file and directory operations).
   - numpy: Numerical computing library for arrays and matrices.
   - pandas: Data manipulation library for tabular data.
   - pytesseract: Python wrapper for the Tesseract OCR engine.
   - re: Regular expression functionality for text manipulation.
   - sqlite3: Built-in Python module for SQLite database operations.
   - PIL (Pillow): Python Imaging Library for image processing.
   - Flask: Micro web framework for building web applications.
   - render_template, request, send_file: Flask functions for rendering HTML templates and handling HTTP requests.

2. Configuration:
   - Sets the Tesseract executable path.
   - Defines base directories for files, CSS, and PDF images.
   - Creates necessary directories if they don't exist.

3. Flask Application:
   - Initializes the Flask app.
   - Handles file uploads, PDF-to-image conversion, and database updates.

"""
import base64
import cv2
import fitz
import io
import os
import numpy as np
import pandas as pd
import pytesseract
import re
import sqlite3
from PIL import Image
from flask import Flask, render_template, request, send_file
from urllib.parse import unquote

# Set the Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Initialize the Flask application
app = Flask(__name__)

# Base directory where your files are located
base_directory = os.path.dirname(__file__)
os.chdir(base_directory)

# Directory where CSS file is located
css_directory = os.path.join(base_directory, 'static')

# Directory where PDF images are stored
base_directory = r"Policies"
app.config['base_directory'] = base_directory
dest_folder = r"pdfimages"
image_folder = r"pdfimages"
app.config['image_folder'] = image_folder

# Create the base directory if it doesn't exist
if not os.path.exists(base_directory):
    os.makedirs(base_directory)

# Create the image folder if it doesn't exist
if not os.path.exists(image_folder):
    os.makedirs(image_folder)


# CSS styles (moved to CSS file)
CSS_STYLES = """
<style>
    mark { 
        padding: 0;
        background-color: transparent;
        color: inherit;
    }
    mark.highlight { 
        background-color: #ff0;
        font-weight: bold;
    }
    table {
        width: 100%;
    }
</style>
"""


def download_file(base_path, category, file_name):
    """
    Generates a download link for a file based on the specified category and file name.

    Args:
        base_path (str): The base directory path where files are stored.
        category (str): The category of the file (e.g., 'Contracts', 'ISO', 'Policies', 'GDMS').
        file_name (str): The name of the file to be downloaded.

    Returns:
        str: A hyperlink (HTML anchor tag) that allows the user to download the specified file.

    Example:
        base_path = '/path/to/files'
        category = 'Contracts'
        file_name = 'contract.pdf'
        download_link = download_file(base_path, category, file_name)
        # The returned value will be an HTML anchor tag with the appropriate download link.
    """
    if not category:
        return "Category is empty."

    category_base_path = {
        'Contracts': os.path.join(base_path, 'Contracts'),
        'ISO': os.path.join(base_path, 'ISO'),
        'Policies': os.path.join(base_path, 'Policies'),
        'GDMS': os.path.join(base_path, 'GDMS')
    }
    if category not in category_base_path:
        return f"Invalid category: {category}"

    file_path = os.path.join(category_base_path[category], file_name)
    if os.path.exists(file_path):
        with open(file_path, "rb") as file:
            file_content = file.read()
            b64 = base64.b64encode(file_content).decode()
            href = (f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}"'
                    f'>Download {file_name}</a>')
            return href
    else:
        return f'File not found: {file_path}'


def view_page_link(category, filename, pagenumber):
    """
    Generates a hyperlink to view a specific page of an image associated with a PDF file.

    Args:
        category (str): The category of the PDF file (e.g., 'Contracts', 'ISO', 'Policies', 'GDMS').
        filename (str): The name of the PDF file (without the ".pdf" extension).
        pagenumber (int): The page number to view (corresponding to an image).

    Returns:
        str: An HTML anchor tag that opens the specified page image in a new tab when clicked.

    Example:
        category = 'Contracts'
        filename = 'contract'
        pagenumber = 3
        link = view_page_link(category, filename, pagenumber)
        # The returned value will be an HTML anchor tag with the appropriate link.
    """
    filename = re.sub(r'\.pdf$', '', filename)  # Remove the ".pdf" extension from the filename
    image_path = os.path.join(dest_folder, category, filename, f"{pagenumber}.png")
    if os.path.exists(image_path):
        return f'<a href="/view_image/{category}/{filename}/{pagenumber}" target="_blank">View Page</a>'
    else:
        return 'Image not found'


def search_data(db_name, keywords, category):
    """
    Searches for files in a database based on specified keywords and optional category.

    Args:
        db_name (str): The name of the SQLite database.
        keywords (str): Keywords to search for in the text or filename.
        category (str, optional): The category of files to search within ('All' by default).

    Returns:
        pandas.DataFrame: A DataFrame containing search results with columns: 'S.NO', 'filename', 'category',
                          'pagenumber', 'View Page', and 'Download'.

    Example:
        db_name = 'my_database.db'
        keywords = 'contract'
        category = 'Contracts'
        result_df = search_data(db_name, keywords, category)
        # The returned DataFrame will include relevant search results.
    """
    conn = sqlite3.connect(db_name)

    if category and category != 'All':
        query = (f"SELECT filename, '{category}' AS category, pagenumber FROM "
                 f"{category.lower()} WHERE text LIKE ? OR filename LIKE ?")
        params = ('%' + keywords + '%', '%' + keywords + '%')
    else:
        query = (
            "SELECT filename, 'Contracts' AS category, pagenumber FROM contracts WHERE text LIKE ? OR filename LIKE ? "
            "UNION SELECT filename, 'Policies' AS category, pagenumber FROM policies WHERE text LIKE ? OR filename LIKE ? "
            "UNION SELECT filename, 'GDMS' AS category, pagenumber FROM GDMS WHERE text LIKE ? OR filename LIKE ? "
            "UNION SELECT filename, 'ISO' AS category, pagenumber FROM iso WHERE text LIKE ? OR filename LIKE ?"
        )

        params = (
            '%' + keywords + '%', '%' + keywords + '%',
            '%' + keywords + '%', '%' + keywords + '%',
            '%' + keywords + '%', '%' + keywords + '%',
            '%' + keywords + '%', '%' + keywords + '%',
        )

    # Fetch data with category
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    # Check if DataFrame is empty
    if df.empty:
        return df
    df.reset_index(drop=True, inplace=True)
    df.index = df.index + 1
    df.rename_axis('S.NO', axis=1, inplace=True)
    # Add download link and view page link columns
    df['View Page'] = df.apply(lambda row: view_page_link(row['category'], row['filename'], row['pagenumber']), axis=1)
    df['Download'] = df.apply(lambda row: download_file(base_directory, row['category'], row['filename']), axis=1)
    return df


def get_df2(db_name):
    """
        Retrieves a DataFrame containing information about files with page number 1 from various categories.

        Args:
            db_name (str): The name of the SQLite database.

        Returns:
            pandas.DataFrame: A DataFrame with columns: 'S.NO', 'filename', and 'category'.

        Example:
            db_name = 'my_database.db'
            df2 = get_df2(db_name)
            # The returned DataFrame will include relevant file information.
        """
    conn = sqlite3.connect(db_name)
    # SQL query to select data with pagenumber 1 from all tables
    query = ("SELECT filename, 'Contracts' AS category, pagenumber "
             "FROM contracts WHERE pagenumber = 1 UNION SELECT filename, "
             "'Policies' AS category, pagenumber FROM policies WHERE pagenumber = 1 UNION SELECT filename, "
             "'ISO' AS category, pagenumber FROM iso WHERE pagenumber = 1 UNION SELECT filename, "
             "'GDMS' AS category, pagenumber FROM GDMS WHERE pagenumber = 1"
             )
    df2 = pd.read_sql_query(query, conn)
    conn.close()
    df2.reset_index(drop=True, inplace=True)
    df2 = df2.drop_duplicates(subset='filename', keep='first')
    df2.sort_values(by='category', inplace=True)
    df2.reset_index(drop=True, inplace=True)
    df2.drop('pagenumber', axis=1, inplace=True)
    df2.index = df2.index + 1
    df2.rename_axis('S.NO', axis=1, inplace=True)
    return df2


def generate_grouped_html_tables(grouped_df):
    """
    Generates an HTML representation of grouped data tables.

    Args:
        grouped_df (pandas.DataFrame): A DataFrame containing grouped data.

    Returns:
        str: HTML code representing the grouped tables.
    """
    html_code = ""
    for (filename, category), group in grouped_df.groupby(['filename', 'category']):
        filename = re.sub(r'\.pdf$', '', filename)
        # Add filename and category as the title of the group
        group_title = f"{filename} - {category}"
        html_code += f"<div class='group'><h2>{group_title}</h2>"

        # Generate HTML table for the group
        html_code += "<table>"
        # Add table headers
        html_code += "<tr>"
        for column in group.columns:
            html_code += f"<th>{column}</th>"
        html_code += "</tr>"
        # Add table rows
        for _, row in group.iterrows():
            html_code += "<tr>"
            for value in row:
                html_code += f"<td>{value}</td>"
            html_code += "</tr>"
        html_code += "</table></div>"

    return html_code


def correct_skewness(image):
    """
    Corrects skewness in an input image.

    Args:
        image (numpy.ndarray): The input image (BGR format).

    Returns:
        numpy.ndarray: The rotated image with corrected skew.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.bitwise_not(gray)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thresh > 0))
    angle = cv2.minAreaRect(coords)[-1]

    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    h, w = image.shape[:2]
    center = (w // 2, h // 2)
    m = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, m, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    return rotated


def extract_text_from_image(image):
    """
       Extracts text from an input image using optical character recognition (OCR).

       Args:
           image (numpy.ndarray): The input image (BGR format).

       Returns:
           str: Extracted text from the image (stripped of leading/trailing whitespace).
       """
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray_image)
    return text.strip()


def check_scanned_pdf(pdf_file):
    """
    Extracts images from a scanned PDF document.

    Args:
        pdf_file (str): Path to the input PDF file.
    Returns:
        list: A list of tuples, where each tuple contains an image (as a numpy array)
              and the corresponding page number.
    """

    images = []
    pdf_doc = fitz.open(pdf_file)
    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = pdf_doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = Image.open(io.BytesIO(image_bytes))
            image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            images.append((image, page_num + 1))
    return images


def extract_text_from_pdf(pdf_file):
    """
       Extracts text from a PDF file, handling both scanned images and regular text.

       Args:
           pdf_file (str): Path to the input PDF file.

       Returns:
           list: A list of tuples, where each tuple contains:
               - Page number (int)
               - Extracted text (str)
               - Source type ('Image' or 'PDF')
       """
    extracted_texts = []
    images = check_scanned_pdf(pdf_file)
    if images:
        for image, page_num in images:
            image = correct_skewness(image)
            text_from_image = extract_text_from_image(image)
            extracted_texts.append((page_num, text_from_image.strip(), 'Image'))
    else:
        pdf_doc = fitz.open(pdf_file)
        for index, page in enumerate(pdf_doc):
            text = page.get_text().strip()
            extracted_texts.append((index + 1, text, 'PDF'))
    return extracted_texts


def check_file_exists(db_name, filename):
    """
       Checks if a given filename exists in any of the specified database tables.

       Args:
           db_name (str): Path to the SQLite database file.
           filename (str): The filename to search for.

       Returns:
           tuple: A tuple containing:
               - bool: True if the file exists in any table, False otherwise.
               - str or None: The name of the table where the file was found (or None if not found).
       """
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    tables = ['contracts', 'policies', 'iso', 'GDMS']
    for table in tables:
        cursor.execute(f"SELECT filename FROM {table} WHERE filename=?", (filename,))
        result = cursor.fetchone()
        if result:
            conn.close()
            return True, table  # Return True and table name if file exists
    conn.close()
    return False, None  # Return False if file does not exist in any table


def create_table_if_not_exists(db_name, table_name):
    """
    Creates an SQLite table if it does not already exist.

    Args:
        db_name (str): Path to the SQLite database file.
        table_name (str): Name of the table to create.

    Returns:
        None
    """
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        category TEXT,
        pagenumber INTEGER,
        text TEXT
    )
    """)
    conn.commit()
    conn.close()


# PDF to Image Conversion Functions
def pdf_to_images(pdf_path, output_folder):
    """
    Converts each page of a PDF document into individual images and saves them to the specified folder.

    Args:
        pdf_path (str): Path to the input PDF file.
        output_folder (str): Path to the folder where images will be saved.

    Returns:
        None
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    pdf_document = fitz.open(pdf_path)
    for page_number in range(pdf_document.page_count):
        page = pdf_document.load_page(page_number)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_array = np.array(img)
        if img_array.shape[2] == 4:
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
        enhanced_image = cv2.convertScaleAbs(img_array, alpha=0.85, beta=0)
        kernel = np.array([[0, -0.5, 0], [-0.5, 3, -0.5], [0, -0.5, 0]])
        sharpened_image = cv2.filter2D(enhanced_image, -1, kernel)
        image_path = os.path.join(output_folder, f"{page_number + 1}.png")
        cv2.imwrite(image_path, sharpened_image)
    pdf_document.close()


def convert_pdfs_to_images(source_folder, dest_folder):
    """
    Converts PDF files in the source folder to individual images and saves them in the destination folder.

    Args:
        source_folder (str): Path to the folder containing PDF files.
        dest_folder (str): Path to the destination folder for saving images.

    Returns:
        None
    """
    for subfolder_name in os.listdir(source_folder):
        subfolder_path = os.path.join(source_folder, subfolder_name)
        if os.path.isdir(subfolder_path):
            dest_subfolder_path = os.path.join(dest_folder, subfolder_name)
            os.makedirs(dest_subfolder_path, exist_ok=True)
            for file_name in os.listdir(subfolder_path):
                file_path = os.path.join(subfolder_path, file_name)
                if file_name.endswith(".pdf"):
                    pdf_name = os.path.splitext(file_name)[0]
                    pdf_output_folder = os.path.join(dest_subfolder_path, pdf_name)
                    os.makedirs(pdf_output_folder, exist_ok=True)
                    pdf_to_images(file_path, pdf_output_folder)


@app.route("/")
def index():
    """
    Renders the main page (homepage) of the web application.

    Returns:
        str: HTML content for the main page.
    """
    return render_template('index.html')


@app.route("/index_page")
def display_df2():
    """
       Retrieves data from the 'Archivus.db' database and renders it in an HTML template.

       Returns:
           str: Rendered HTML content containing data from the 'df2' DataFrame.
       """
    df2 = get_df2("Archivus_database.db")
    return render_template('index_page.html', df2=df2)


@app.route("/search")
def search():
    """
    Handles search requests based on keywords and category.

    Retrieves data from the "Archivus.db" database based on the provided keywords and optional category.
    If matching data is found, it generates grouped HTML tables for display.

    Returns:
        str: Rendered HTML content for search results or an empty results page.
    """
    keywords = request.args.get('keywords')
    category = request.args.get('category')
    df = search_data("Archivus_database.db", keywords, category if category else None)

    if df.empty:
        return render_template('search_results_empty.html', keywords=keywords)
    else:
        grouped_html_tables = generate_grouped_html_tables(df)
        return render_template('search_results.html', grouped_html_tables=grouped_html_tables)


@app.route("/view_image/<category>/<filename>/<int:pagenumber>")
def view_image(category, filename, pagenumber):
    """
    Retrieves and serves an image file for viewing.

    Args:
        category (str): The category of the image.
        filename (str): The base filename (without extension) of the PDF.
        pagenumber (int): The page number of the image to retrieve.

    Returns:
        str or file: If the image exists, returns the image file (PNG format).
                     Otherwise, returns a message indicating that the image was not found.
    """
    # Decode the URL-encoded filename
    filename = unquote(filename)

    # Remove the ".pdf" extension
    filename = re.sub(r'\.pdf$', '', filename)

    image_path = os.path.join(dest_folder, category, filename, f"{pagenumber}.png")
    if os.path.exists(image_path):
        return send_file(image_path, mimetype='image/png')
    else:
        return f'Image not found at path: {image_path}'


@app.route("/indexupload")
def upload():
    """
    Renders the file upload page for users.

    Returns:
        str: HTML content for the file upload page.
    """
    return render_template('indexupload.html')


@app.route('/', methods=['GET', 'POST'])
def uploadindex():
    """
      Handles file uploads, converts PDFs to images, and updates an SQLite database.

      This function performs the following steps:
      1. Accepts a file upload with a specified category.
      2. Saves the uploaded file to the appropriate category folder.
      3. Converts the uploaded PDF to individual images and saves them.
      4. Checks if the file already exists in the database.
      5. Inserts extracted text from the PDF into the corresponding database table.

      Returns:
          str: A message indicating successful file processing or existing file status.
      """
    if request.method == 'POST':
        category = request.form['category']
        file = request.files['file']
        if file:
            filename = file.filename
            category_path = os.path.join(app.config['base_directory'], category)
            if not os.path.exists(category_path):
                os.makedirs(category_path)
            file_path = os.path.join(category_path, filename)
            file.save(file_path)

            # Convert the PDF to images
            pdf_name = os.path.splitext(filename)[0]
            pdf_output_folder = os.path.join(app.config['image_folder'], category, pdf_name)
            os.makedirs(pdf_output_folder, exist_ok=True)
            pdf_to_images(file_path, pdf_output_folder)

            db_name = "Archivus_database.db"
            file_exists, table_name = check_file_exists(db_name, filename)
            if file_exists:
                return f"File {filename} is already present in {table_name} table."
            else:
                create_table_if_not_exists(db_name, category)
                texts = extract_text_from_pdf(file_path)
                conn = sqlite3.connect(db_name)
                cursor = conn.cursor()
                for page_num, text, source in texts:
                    text_with_source = f"{text} [{source}]"
                    cursor.execute(f"INSERT INTO {category} (filename, category, pagenumber, text) VALUES (?, ?, ?, ?)",
                                   (filename, category, page_num, text_with_source))
                conn.commit()
                conn.close()
                return f"Sucessfully Inserted text from {filename} in category {category} into database."
    return render_template('indexupload.html')
    return uploadindex()


if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=True)
