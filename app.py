from flask import Flask, request, send_file, render_template_string
import os
import subprocess
from PIL import Image
import fitz  # PyMuPDF
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>File Converter</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        * { box-sizing: border-box; }

        body {
            margin: 0;
            font-family: Arial, sans-serif;
            background: #f4f7fb;
            color: #172033;
        }

        .container {
            max-width: 700px;
            margin: 70px auto;
            padding: 25px;
        }

        .card {
            background: white;
            padding: 35px;
            border-radius: 20px;
            box-shadow: 0 10px 35px rgba(0,0,0,.08);
            text-align: center;
        }

        h1 {
            margin-bottom: 10px;
            color: #2563eb;
        }

        .upload {
            border: 2px dashed #2563eb;
            border-radius: 15px;
            padding: 45px 20px;
            margin: 25px 0;
            cursor: pointer;
            background: #f8fbff;
        }

        input[type=file] {
            display: none;
        }

        select, button {
            width: 100%;
            padding: 14px;
            margin-top: 12px;
            border-radius: 10px;
            font-size: 16px;
        }

        select {
            border: 1px solid #ddd;
        }

        button {
            border: none;
            background: #2563eb;
            color: white;
            cursor: pointer;
            font-weight: bold;
        }

        button:hover {
            background: #1d4ed8;
        }

        #fileName {
            margin-top: 10px;
            color: #555;
        }

        .formats {
            margin-top: 25px;
            color: #666;
            font-size: 14px;
        }
    </style>
</head>

<body>

<div class="container">
    <div class="card">

        <h1>📄 File Converter</h1>
        <p>Convert your documents quickly and easily.</p>

        <form action="/convert" method="POST" enctype="multipart/form-data">

            <label class="upload">
                📁 Click to choose a file
                <input type="file" name="file" id="file" required>
                <div id="fileName"></div>
            </label>

            <select name="conversion" required>
                <option value="">Select conversion</option>
                <option value="word_pdf">Word → PDF</option>
                <option value="pdf_word">PDF → Word</option>
                <option value="image_pdf">Image → PDF</option>
            </select>

            <button type="submit">Convert File</button>

        </form>

        <div class="formats">
            Supported: DOCX, PDF, JPG, JPEG, PNG
        </div>

    </div>
</div>

<script>
const fileInput = document.getElementById("file");
const fileName = document.getElementById("fileName");

fileInput.addEventListener("change", () => {
    if (fileInput.files.length) {
        fileName.textContent = "Selected: " + fileInput.files[0].name;
    }
});
</script>

</body>
</html>
"""


@app.route("/")
def home():
    return render_template_string(HTML)


@app.route("/convert", methods=["POST"])
def convert():

    file = request.files.get("file")
    conversion = request.form.get("conversion")

    if not file:
        return "No file selected", 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_DIR, filename)
    file.save(input_path)

    # Word → PDF
    if conversion == "word_pdf":

        output_dir = os.path.abspath(OUTPUT_DIR)

        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_path
        ], check=True)

        output_file = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(OUTPUT_DIR, output_file)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_file
        )

    # PDF → Word
    elif conversion == "pdf_word":

        doc = fitz.open(input_path)

        output_file = os.path.splitext(filename)[0] + ".docx"
        output_path = os.path.join(OUTPUT_DIR, output_file)

        from docx import Document

        word = Document()

        for page in doc:
            text = page.get_text()
            word.add_paragraph(text)

        word.save(output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_file
        )

    # Image → PDF
    elif conversion == "image_pdf":

        image = Image.open(input_path)

        if image.mode != "RGB":
            image = image.convert("RGB")

        output_file = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(OUTPUT_DIR, output_file)

        image.save(output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_file
        )

    return "Invalid conversion", 400


if __name__ == "__main__":
    app.run(debug=True)
