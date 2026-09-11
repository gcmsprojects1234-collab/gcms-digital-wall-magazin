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

    # Get all available MP3 files
    audio_files = {}
    for audio_file in os.listdir(AUDIO_FOLDER):
        if audio_file.lower().endswith(".mp3"):
            base_name = os.path.splitext(audio_file)[0].lower()
            audio_files[base_name] = audio_file

    for filename in os.listdir(UPLOAD_FOLDER):

        if filename.lower().endswith(".docx"):

            file_path = os.path.join(UPLOAD_FOLDER, filename)

            try:
                text = read_docx(file_path)

                title = os.path.splitext(filename)[0]

                # Generate expected audio filename
                expected_audio = safe_filename(filename)
                expected_base = expected_audio.lower()

                # First try exact match
                audio_name = audio_files.get(expected_base)

                # If exact match is not found, try normalized matching
                if not audio_name:
                    normalized_expected = re.sub(
                        r'[^a-z0-9]+',
                        '_',
                        expected_base
                    ).strip('_')

                    for existing_base, existing_file in audio_files.items():

                        normalized_existing = re.sub(
                            r'[^a-z0-9]+',
                            '_',
                            existing_base
                        ).strip('_')

                        if normalized_existing == normalized_expected:
                            audio_name = existing_file
                            break

                audio_exists = audio_name is not None

                articles.append({
                    "title": title,
                    "filename": filename,
                    "text": text,
                    "audio": audio_name if audio_name else expected_audio + ".mp3",
                    "audio_exists": audio_exists
                })

            except Exception as e:
                print(f"Error reading {filename}: {e}")

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