from flask import Flask, render_template
from docx import Document
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
AUDIO_FOLDER = "static/audio"


def read_docx(file_path):

    document = Document(file_path)

    paragraphs = []

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:
            paragraphs.append(text)

    return "\n".join(paragraphs)


def safe_filename(filename):

    filename = os.path.splitext(filename)[0]

    filename = re.sub(r'[\\/*?:"<>|]', "", filename)

    filename = filename.replace(" ", "_")

    return filename


def get_all_articles():

    articles = []

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(AUDIO_FOLDER, exist_ok=True)

    for filename in os.listdir(UPLOAD_FOLDER):

        if filename.lower().endswith(".docx"):

            file_path = os.path.join(
                UPLOAD_FOLDER,
                filename
            )

            try:

                text = read_docx(file_path)

                audio_name = safe_filename(filename) + ".mp3"

                audio_path = os.path.join(
                    AUDIO_FOLDER,
                    audio_name
                )

                # Check whether audio exists
                audio_exists = os.path.exists(audio_path)

                articles.append({
                    "title": os.path.splitext(filename)[0],
                    "filename": filename,
                    "text": text,
                    "audio": audio_name,
                    "audio_exists": audio_exists
                })

            except Exception as e:

                print(
                    f"Error reading {filename}: {e}"
                )

    return articles


@app.route("/")
def home():

    articles = get_all_articles()

    return render_template(
        "index.html",
        articles=articles
    )


if __name__ == "__main__":

    app.run(
        debug=True
    )