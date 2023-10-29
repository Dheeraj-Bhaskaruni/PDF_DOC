from flask import Flask, render_template, request, send_from_directory
import os
from pdf2docx import Converter

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

# Define a maximum upload size, for example, 10MB
MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf_file"]
        if not file or file.filename == "":
            return "Please select a file"

        # Validate file size
        file.seek(0, os.SEEK_END)
        file_length = file.tell()
        if file_length > MAX_UPLOAD_SIZE:
            return "File size exceeds the allowed limit of 10MB."

        # Reset file pointer for further operations
        file.seek(0)

        # Validate MIME type
        if not file.mimetype == 'application/pdf':
            return "Invalid file type. Please upload a PDF."

        filename = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(filename)

        page_range = request.form.get("page_range")
        output_format = request.form.get("output_format", "docx")

        output_filename = os.path.join(app.config["OUTPUT_FOLDER"], file.filename.replace(".pdf", f".{output_format}"))

        # Convert the page range string to a list of integers
        pages = parse_page_range(page_range)

        cv = Converter(filename)
        cv.convert(output_filename, start=0, end=None, pages=pages)
        cv.close()

        return send_from_directory(app.config["OUTPUT_FOLDER"], file.filename.replace(".pdf", f".{output_format}"), as_attachment=True)

    return render_template("index.html")

def parse_page_range(page_range_str):
    """Convert a page range string (e.g., '1-3,5,7') into a list of page numbers (e.g., [1,2,3,5,7])."""
    if not page_range_str:
        return None

    pages = []
    for part in page_range_str.split(","):
        if "-" in part:
            start, end = map(int, part.split("-"))
            pages.extend(range(start, end+1))
        else:
            pages.append(int(part))
    return pages

if __name__ == "__main__":
    app.run(debug=True)
