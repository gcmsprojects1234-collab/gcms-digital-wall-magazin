import os
import re
import asyncio
import subprocess
import tempfile

import edge_tts
import imageio_ffmpeg

from docx import Document
from docx.document import Document as _Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


# =========================================================
# FOLDER SETTINGS
# =========================================================

UPLOAD_FOLDER = "uploads"
AUDIO_FOLDER = "static/audio"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(AUDIO_FOLDER, exist_ok=True)


# =========================================================
# FFMPEG
# =========================================================

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()


# =========================================================
# NEURAL VOICES
# =========================================================

BENGALI_VOICE = "bn-IN-TanishaaNeural"
HINDI_VOICE = "hi-IN-SwaraNeural"
ENGLISH_VOICE = "en-IN-NeerjaNeural"


# =========================================================
# VOICE SETTINGS
# =========================================================

RATE = "-5%"
VOLUME = "+0%"
PITCH = "+0Hz"


# =========================================================
# DOCUMENT BLOCK READER
# Keeps paragraphs and tables in actual DOCX order
# =========================================================

def iter_block_items(parent):

    if isinstance(parent, _Document):
        parent_elm = parent.element.body

    elif isinstance(parent, _Cell):
        parent_elm = parent._tc

    else:
        raise ValueError(
            "Unsupported parent type"
        )

    for child in parent_elm.iterchildren():

        if isinstance(child, CT_P):

            yield Paragraph(
                child,
                parent
            )

        elif isinstance(child, CT_Tbl):

            yield Table(
                child,
                parent
            )


# =========================================================
# READ COMPLETE DOCX
# =========================================================

def read_docx(file_path):

    document = Document(file_path)

    content = []

    for block in iter_block_items(document):

        # -------------------------------------------------
        # PARAGRAPH
        # -------------------------------------------------

        if isinstance(block, Paragraph):

            text = block.text.strip()

            if text:
                content.append(text)

        # -------------------------------------------------
        # TABLE
        # -------------------------------------------------

        elif isinstance(block, Table):

            for row in block.rows:

                row_text = []

                for cell in row.cells:

                    cell_text = cell.text.strip()

                    if cell_text:
                        row_text.append(
                            cell_text
                        )

                if row_text:

                    content.append(
                        " ".join(row_text)
                    )

    return content


# =========================================================
# CLEAN TEXT
# =========================================================

def clean_text(text):

    # Remove bullet symbols
    text = re.sub(
        r"[•●▪■◆◇]",
        "",
        text
    )

    # Replace tabs with spaces
    text = text.replace(
        "\t",
        " "
    )

    # Replace multiple spaces
    text = re.sub(
        r"[ ]+",
        " ",
        text
    )

    # Clean excessive exclamation marks
    text = re.sub(
        r"!{2,}",
        "!",
        text
    )

    # Clean excessive question marks
    text = re.sub(
        r"\?{2,}",
        "?",
        text
    )

    # Clean excessive dots
    text = re.sub(
        r"\.{4,}",
        "...",
        text
    )

    return text.strip()


# =========================================================
# DETECT LANGUAGE
# =========================================================

def detect_language(text):

    bengali_count = len(
        re.findall(
            r"[\u0980-\u09FF]",
            text
        )
    )

    hindi_count = len(
        re.findall(
            r"[\u0900-\u097F]",
            text
        )
    )

    english_count = len(
        re.findall(
            r"[A-Za-z]",
            text
        )
    )

    # Bengali
    if (
        bengali_count > hindi_count
        and bengali_count >= english_count
    ):

        return "bn"

    # Hindi
    elif (
        hindi_count > bengali_count
        and hindi_count >= english_count
    ):

        return "hi"

    # English
    elif (
        english_count > bengali_count
        and english_count > hindi_count
    ):

        return "en"

    # Default
    return "en"


# =========================================================
# GET VOICE
# =========================================================

def get_voice(language):

    if language == "bn":

        return BENGALI_VOICE

    elif language == "hi":

        return HINDI_VOICE

    elif language == "en":

        return ENGLISH_VOICE

    return ENGLISH_VOICE


# =========================================================
# NATURAL TEXT PREPARATION
# =========================================================

def prepare_text(text, language):

    text = clean_text(text)

    if not text:
        return ""

    # Add natural ending punctuation
    if not re.search(
        r"[.!?।॥]$",
        text
    ):

        if language in ["bn", "hi"]:

            text += "।"

        else:

            text += "."

    return text


# =========================================================
# SPLIT LONG TEXT
#
# Tries to split at sentence boundaries first.
# =========================================================

def split_long_text(
    text,
    max_chars=1500
):

    # -----------------------------------------------------
    # First split by sentence
    # -----------------------------------------------------

    sentences = re.split(
        r"(?<=[.!?।॥])\s+",
        text
    )

    chunks = []

    current = ""

    for sentence in sentences:

        sentence = sentence.strip()

        if not sentence:
            continue

        # -------------------------------------------------
        # If adding sentence is safe
        # -------------------------------------------------

        test = (
            current + " " + sentence
        ).strip()

        if len(test) <= max_chars:

            current = test

        else:

            if current:

                chunks.append(
                    current
                )

            # -------------------------------------------------
            # Sentence itself too long
            # -------------------------------------------------

            if len(sentence) > max_chars:

                words = sentence.split()

                small_chunk = ""

                for word in words:

                    test_word = (
                        small_chunk
                        + " "
                        + word
                    ).strip()

                    if len(test_word) <= max_chars:

                        small_chunk = test_word

                    else:

                        if small_chunk:

                            chunks.append(
                                small_chunk
                            )

                        small_chunk = word

                if small_chunk:

                    current = small_chunk

                else:

                    current = ""

            else:

                current = sentence

    if current:

        chunks.append(
            current
        )

    return chunks


# =========================================================
# GENERATE ONE CHUNK
# =========================================================

async def generate_chunk(
    text,
    output_file,
    language
):

    voice = get_voice(
        language
    )

    print(
        f"       Voice: {voice}"
    )

    communicate = edge_tts.Communicate(
        text=text,
        voice=voice,
        rate=RATE,
        volume=VOLUME,
        pitch=PITCH
    )

    await communicate.save(
        output_file
    )


# =========================================================
# CHECK AUDIO FILE
# =========================================================

def audio_file_is_valid(
    file_path
):

    return (
        os.path.exists(file_path)
        and os.path.getsize(file_path) > 1000
    )


# =========================================================
# MERGE AUDIO FILES USING FFMPEG
#
# No pydub.
# No audioop.
# Works with Python 3.14.
# =========================================================

def merge_audio_files(
    audio_files,
    output_file
):

    if not audio_files:

        raise RuntimeError(
            "No audio files to merge."
        )

    # -----------------------------------------------------
    # Create temporary concat list
    # -----------------------------------------------------

    concat_file = os.path.join(
        AUDIO_FOLDER,
        "_concat_list.txt"
    )

    with open(
        concat_file,
        "w",
        encoding="utf-8"
    ) as f:

        for audio_file in audio_files:

            absolute_path = os.path.abspath(
                audio_file
            )

            # FFmpeg concat format
            absolute_path = (
                absolute_path
                .replace("\\", "/")
                .replace("'", "'\\''")
            )

            f.write(
                f"file '{absolute_path}'\n"
            )

    # -----------------------------------------------------
    # FFmpeg command
    # -----------------------------------------------------

    command = [

        FFMPEG,

        "-y",

        "-f",
        "concat",

        "-safe",
        "0",

        "-i",
        concat_file,

        "-vn",

        "-acodec",
        "libmp3lame",

        "-b:a",
        "128k",

        output_file
    ]

    print(
        "\nRunning FFmpeg..."
    )

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    # -----------------------------------------------------
    # Delete concat list
    # -----------------------------------------------------

    try:

        os.remove(
            concat_file
        )

    except Exception:
        pass

    # -----------------------------------------------------
    # Check result
    # -----------------------------------------------------

    if result.returncode != 0:

        print(
            "\nFFmpeg ERROR:"
        )

        print(
            result.stderr
        )

        raise RuntimeError(
            "FFmpeg failed while merging audio."
        )

    if not audio_file_is_valid(
        output_file
    ):

        raise RuntimeError(
            "Final MP3 was not created correctly."
        )


# =========================================================
# PROCESS ONE DOCX
# =========================================================

async def process_document(
    filename
):

    print("\n")
    print(
        "=" * 70
    )

    print(
        "PROCESSING:",
        filename
    )

    print(
        "=" * 70
    )

    file_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )

    # =====================================================
    # READ COMPLETE DOCUMENT
    # =====================================================

    paragraphs = read_docx(
        file_path
    )

    if not paragraphs:

        print(
            "No text found."
        )

        return

    print(
        "Total sections found:",
        len(paragraphs)
    )

    # =====================================================
    # TEMP DIRECTORY
    # =====================================================

    temp_directory = tempfile.mkdtemp(
        prefix="wall_magazine_",
        dir=AUDIO_FOLDER
    )

    temp_files = []

    chunk_number = 0

    # =====================================================
    # PROCESS EVERY SECTION
    # =====================================================

    for paragraph_number, paragraph in enumerate(
        paragraphs,
        start=1
    ):

        paragraph = clean_text(
            paragraph
        )

        if not paragraph:

            continue

        language = detect_language(
            paragraph
        )

        paragraph = prepare_text(
            paragraph,
            language
        )

        if not paragraph:

            continue

        print("\n")
        print(
            f"Paragraph {paragraph_number}"
        )

        print(
            "Language:",
            language
        )

        print(
            "Characters:",
            len(paragraph)
        )

        # =================================================
        # SPLIT LONG PARAGRAPH
        # =================================================

        chunks = split_long_text(
            paragraph,
            max_chars=1500
        )

        print(
            "Chunks:",
            len(chunks)
        )

        # =================================================
        # GENERATE EACH CHUNK
        # =================================================

        for chunk_index, chunk in enumerate(
            chunks,
            start=1
        ):

            chunk_number += 1

            temp_file = os.path.join(
                temp_directory,
                f"chunk_{chunk_number:04d}.mp3"
            )

            print(
                f"   Generating chunk "
                f"{chunk_number}..."
            )

            try:

                await generate_chunk(
                    chunk,
                    temp_file,
                    language
                )

                # -----------------------------------------
                # Verify generated MP3
                # -----------------------------------------

                if audio_file_is_valid(
                    temp_file
                ):

                    temp_files.append(
                        temp_file
                    )

                    print(
                        "   ✓ Chunk generated"
                    )

                else:

                    print(
                        "   ✗ Audio file appears empty."
                    )

            except Exception as e:

                print(
                    "   ✗ ERROR generating chunk:"
                )

                print(
                    e
                )

    # =====================================================
    # CHECK GENERATED FILES
    # =====================================================

    if not temp_files:

        print(
            "\nNo audio files were generated."
        )

        try:

            os.rmdir(
                temp_directory
            )

        except Exception:
            pass

        return

    print("\n")
    print(
        "=" * 70
    )

    print(
        "GENERATED AUDIO CHUNKS:",
        len(temp_files)
    )

    print(
        "=" * 70
    )

    # =====================================================
    # FINAL FILE NAME
    # =====================================================

    base_name = os.path.splitext(
        filename
    )[0]

    # Remove invalid filename characters
    base_name = re.sub(
        r'[\\/*?:"<>|]',
        "",
        base_name
    )

    # Replace spaces
    base_name = re.sub(
        r"\s+",
        "_",
        base_name
    )

    output_file = os.path.join(
        AUDIO_FOLDER,
        base_name + ".mp3"
    )

    # =====================================================
    # DELETE OLD AUDIO
    # =====================================================

    if os.path.exists(
        output_file
    ):

        try:

            os.remove(
                output_file
            )

        except Exception as e:

            print(
                "Could not delete old audio:",
                e
            )

            return

    # =====================================================
    # MERGE
    # =====================================================

    print("\n")
    print(
        "=" * 70
    )

    print(
        "MERGING ALL AUDIO CHUNKS..."
    )

    print(
        "=" * 70
    )

    try:

        merge_audio_files(
            temp_files,
            output_file
        )

    except Exception as e:

        print(
            "\n✗ MERGING FAILED"
        )

        print(
            e
        )

        # Cleanup
        for temp_file in temp_files:

            try:

                os.remove(
                    temp_file
                )

            except Exception:
                pass

        try:

            os.rmdir(
                temp_directory
            )

        except Exception:
            pass

        return

    # =====================================================
    # DELETE TEMP FILES
    # =====================================================

    print(
        "\nCleaning temporary files..."
    )

    for temp_file in temp_files:

        try:

            if os.path.exists(
                temp_file
            ):

                os.remove(
                    temp_file
                )

        except Exception:
            pass

    try:

        os.rmdir(
            temp_directory
        )

    except Exception:
        pass

    # =====================================================
    # FINAL RESULT
    # =====================================================

    file_size_mb = (
        os.path.getsize(
            output_file
        )
        / (1024 * 1024)
    )

    print("\n")
    print(
        "=" * 70
    )

    print(
        "✓ AUDIO GENERATED SUCCESSFULLY"
    )

    print(
        "=" * 70
    )

    print(
        "File:",
        output_file
    )

    print(
        f"Size: {file_size_mb:.2f} MB"
    )

    print(
        "=" * 70
    )


# =========================================================
# MAIN
# =========================================================

async def main():

    print("\n")
    print(
        "=" * 70
    )

    print(
        "        WALL MAGAZINE NEURAL TTS"
    )

    print(
        "=" * 70
    )

    print(
        "\nUsing Neural Voices:"
    )

    print(
        "Bengali :",
        BENGALI_VOICE
    )

    print(
        "Hindi   :",
        HINDI_VOICE
    )

    print(
        "English :",
        ENGLISH_VOICE
    )

    print(
        "\nPython 3.14 compatible mode"
    )

    print(
        "Pydub: DISABLED"
    )

    print(
        "FFmpeg: ENABLED"
    )

    # =====================================================
    # FIND DOCX FILES
    # =====================================================

    files = sorted(
        [
            filename
            for filename in os.listdir(
                UPLOAD_FOLDER
            )
            if filename.lower().endswith(
                ".docx"
            )
        ]
    )

    if not files:

        print(
            "\nNo DOCX files found in uploads folder."
        )

        return

    print(
        f"\nFound {len(files)} DOCX file(s)."
    )

    print(
        "\nFiles:"
    )

    for index, filename in enumerate(
        files,
        start=1
    ):

        print(
            f"  {index}. {filename}"
        )

    # =====================================================
    # PROCESS EVERY DOCX
    # =====================================================

    for filename in files:

        try:

            await process_document(
                filename
            )

        except Exception as e:

            print("\n")
            print(
                "=" * 70
            )

            print(
                "ERROR processing:",
                filename
            )

            print(
                e
            )

            print(
                "=" * 70
            )

    # =====================================================
    # COMPLETE
    # =====================================================

    print("\n")
    print(
        "=" * 70
    )

    print(
        "        ALL FILES COMPLETED"
    )

    print(
        "=" * 70
    )

    print(
        "\nCheck the static/audio folder."
    )

    print()


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )