import os
import uuid
import shutil
import subprocess
from pathlib import Path

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    jsonify
)
from werkzeug.utils import secure_filename

import fitz  # PyMuPDF
from docx import Document
from pdf2docx import Converter
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader


# ============================================================
# CONFIGURATION
# ============================================================

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB


ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx",
    "txt",
    "jpg",
    "jpeg",
    "png",
    "webp",
    "bmp",
    "xlsx",
    "xls",
    "ppt",
    "pptx"
}


# ============================================================
# HELPERS
# ============================================================

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


def extension(filename):
    return Path(filename).suffix.lower().replace(".", "")


def make_job_folder():
    job_id = uuid.uuid4().hex

    upload_dir = UPLOAD_FOLDER / job_id
    output_dir = OUTPUT_FOLDER / job_id

    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    return job_id, upload_dir, output_dir


def cleanup_job(job_id):
    upload_dir = UPLOAD_FOLDER / job_id
    output_dir = OUTPUT_FOLDER / job_id

    shutil.rmtree(upload_dir, ignore_errors=True)
    shutil.rmtree(output_dir, ignore_errors=True)


def find_libreoffice():
    possible = [
        "libreoffice",
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]

    for program in possible:
        if shutil.which(program):
            return program

        if os.path.exists(program):
            return program

    return None


def convert_with_libreoffice(input_file, output_dir, target_format):
    """
    Uses LibreOffice for:
        DOC/DOCX -> PDF
        XLS/XLSX -> PDF
        PPT/PPTX -> PDF
        TXT -> PDF
    """

    libreoffice = find_libreoffice()

    if not libreoffice:
        raise RuntimeError(
            "LibreOffice is not installed. "
            "Install LibreOffice and make sure it is available."
        )

    command = [
        libreoffice,
        "--headless",
        "--convert-to",
        target_format,
        "--outdir",
        str(output_dir),
        str(input_file)
    ]

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or result.stdout or "LibreOffice conversion failed."
        )

    files = list(output_dir.iterdir())

    if not files:
        raise RuntimeError("Conversion did not produce an output file.")

    return files[0]


# ============================================================
# CONVERSIONS
# ============================================================

def word_to_pdf(input_file, output_dir):
    return convert_with_libreoffice(
        input_file,
        output_dir,
        "pdf"
    )


def excel_to_pdf(input_file, output_dir):
    return convert_with_libreoffice(
        input_file,
        output_dir,
        "pdf"
    )


def powerpoint_to_pdf(input_file, output_dir):
    return convert_with_libreoffice(
        input_file,
        output_dir,
        "pdf"
    )


def txt_to_pdf(input_file, output_dir):
    return convert_with_libreoffice(
        input_file,
        output_dir,
        "pdf"
    )


def pdf_to_word(input_file, output_dir):
    output_file = output_dir / (
        Path(input_file).stem + ".docx"
    )

    converter = Converter(str(input_file))

    try:
        converter.convert(
            str(output_file),
            start=0,
            end=None
        )
    finally:
        converter.close()

    return output_file


def word_to_txt(input_file, output_dir):
    output_file = output_dir / (
        Path(input_file).stem + ".txt"
    )

    document = Document(str(input_file))

    with open(
        output_file,
        "w",
        encoding="utf-8"
    ) as f:

        for paragraph in document.paragraphs:
            f.write(paragraph.text)
            f.write("\n")

    return output_file


def images_to_pdf(input_files, output_dir):
    output_file = output_dir / "converted_images.pdf"

    images = []

    for file in input_files:

        image = Image.open(file)

        if image.mode != "RGB":
            image = image.convert("RGB")

        images.append(image)

    if not images:
        raise RuntimeError("No images were supplied.")

    first = images[0]
    remaining = images[1:]

    first.save(
        output_file,
        "PDF",
        resolution=100.0,
        save_all=True,
        append_images=remaining
    )

    for image in images:
        image.close()

    return output_file


def pdf_to_images(input_file, output_dir):
    pdf = fitz.open(str(input_file))

    created_files = []

    for page_number in range(len(pdf)):

        page = pdf[page_number]

        pix = page.get_pixmap(
            matrix=fitz.Matrix(2, 2),
            alpha=False
        )

        output_file = output_dir / (
            f"page_{page_number + 1}.png"
        )

        pix.save(str(output_file))

        created_files.append(output_file)

    pdf.close()

    # Put all images into a ZIP
    zip_path = output_dir / "pdf_images"

    shutil.make_archive(
        str(zip_path),
        "zip",
        str(output_dir)
    )

    return Path(str(zip_path) + ".zip")


def merge_pdfs(input_files, output_dir):
    output_file = output_dir / "merged.pdf"

    merged = fitz.open()

    for file in input_files:

        pdf = fitz.open(str(file))

        merged.insert_pdf(pdf)

        pdf.close()

    merged.save(str(output_file))
    merged.close()

    return output_file


def split_pdf(input_file, output_dir):
    pdf = fitz.open(str(input_file))

    created = []

    for page_number in range(len(pdf)):

        new_pdf = fitz.open()

        new_pdf.insert_pdf(
            pdf,
            from_page=page_number,
            to_page=page_number
        )

        output_file = output_dir / (
            f"page_{page_number + 1}.pdf"
        )

        new_pdf.save(str(output_file))
        new_pdf.close()

        created.append(output_file)

    pdf.close()

    zip_path = output_dir / "split_pages"

    shutil.make_archive(
        str(zip_path),
        "zip",
        str(output_dir)
    )

    return Path(str(zip_path) + ".zip")


def compress_pdf(input_file, output_dir):
    output_file = output_dir / "compressed.pdf"

    pdf = fitz.open(str(input_file))

    pdf.save(
        str(output_file),
        garbage=4,
        deflate=True,
        clean=True
    )

    pdf.close()

    return output_file


def rotate_pdf(input_file, output_dir, angle):
    output_file = output_dir / "rotated.pdf"

    pdf = fitz.open(str(input_file))

    angle = int(angle)

    if angle not in [90, 180, 270]:
        angle = 90

    for page in pdf:
        page.set_rotation(
            page.rotation + angle
        )

    pdf.save(str(output_file))

    pdf.close()

    return output_file


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():
    return render_template("index.html")


# ============================================================
# CONVERT
# ============================================================

@app.route("/convert", methods=["POST"])
def convert():

    files = request.files.getlist("files")
    operation = request.form.get("operation", "")

    if not files or not any(
        f.filename for f in files
    ):
        return jsonify({
            "success": False,
            "error": "Please select a file."
        }), 400

    job_id, upload_dir, output_dir = make_job_folder()

    saved_files = []

    try:

        # ----------------------------------------------------
        # Save uploaded files
        # ----------------------------------------------------

        for file in files:

            if not file.filename:
                continue

            if not allowed_file(file.filename):
                raise RuntimeError(
                    f"File type not supported: {file.filename}"
                )

            filename = secure_filename(
                file.filename
            )

            save_path = upload_dir / filename

            file.save(str(save_path))

            saved_files.append(save_path)

        if not saved_files:
            raise RuntimeError(
                "No valid files were uploaded."
            )

        # ----------------------------------------------------
        # Operations
        # ----------------------------------------------------

        if operation == "word_to_pdf":

            output = word_to_pdf(
                saved_files[0],
                output_dir
            )

        elif operation == "pdf_to_word":

            output = pdf_to_word(
                saved_files[0],
                output_dir
            )

        elif operation == "word_to_txt":

            output = word_to_txt(
                saved_files[0],
                output_dir
            )

        elif operation == "txt_to_pdf":

            output = txt_to_pdf(
                saved_files[0],
                output_dir
            )

        elif operation == "images_to_pdf":

            output = images_to_pdf(
                saved_files,
                output_dir
            )

        elif operation == "pdf_to_images":

            output = pdf_to_images(
                saved_files[0],
                output_dir
            )

        elif operation == "merge_pdf":

            output = merge_pdfs(
                saved_files,
                output_dir
            )

        elif operation == "split_pdf":

            output = split_pdf(
                saved_files[0],
                output_dir
            )

        elif operation == "compress_pdf":

            output = compress_pdf(
                saved_files[0],
                output_dir
            )

        elif operation == "rotate_pdf":

            angle = request.form.get(
                "angle",
                "90"
            )

            output = rotate_pdf(
                saved_files[0],
                output_dir,
                angle
            )

        elif operation == "excel_to_pdf":

            output = excel_to_pdf(
                saved_files[0],
                output_dir
            )

        elif operation == "powerpoint_to_pdf":

            output = powerpoint_to_pdf(
                saved_files[0],
                output_dir
            )

        else:

            raise RuntimeError(
                "Unknown conversion operation."
            )

        # ----------------------------------------------------
        # Download response
        # ----------------------------------------------------

        if not output.exists():

            raise RuntimeError(
                "Output file was not created."
            )

        return send_file(
            str(output),
            as_attachment=True,
            download_name=output.name
        )

    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

    finally:

        # Do not immediately remove the job folder here,
        # because send_file still needs the output.
        #
        # The cleanup endpoint is called by the frontend.
        pass


# ============================================================
# CLEANUP
# ============================================================

@app.route("/cleanup/<job_id>", methods=["POST"])
def cleanup(job_id):

    cleanup_job(job_id)

    return jsonify({
        "success": True
    })


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():

    return jsonify({
        "status": "ok",
        "service": "Document Converter"
    })


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
