from docx import Document
from gtts import gTTS
import os
import re

UPLOAD_FOLDER = "uploads"
AUDIO_FOLDER = "static/audio"

os.makedirs(AUDIO_FOLDER, exist_ok=True)


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

    # Special characters remove
    filename = re.sub(r'[\\/*?:"<>|]', "", filename)

    # Space replace
    filename = filename.replace(" ", "_")

    return filename


for filename in os.listdir(UPLOAD_FOLDER):

    if filename.lower().endswith(".docx"):

        file_path = os.path.join(UPLOAD_FOLDER, filename)

        text = read_docx(file_path)

        if not text.strip():
            print("Empty file:", filename)
            continue

        audio_name = safe_filename(filename) + ".mp3"

        audio_path = os.path.join(AUDIO_FOLDER, audio_name)

        print("Generating audio:", filename)

        try:
            tts = gTTS(
                text=text,
                lang="bn",
                slow=False
            )

            tts.save(audio_path)

            print("Done:", audio_name)

        except Exception as e:
            print("Error:", filename)
            print(e)


print("\nAll audio files generated!")