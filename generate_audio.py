
import os
import re
import asyncio
import subprocess
import tempfile
import shutil

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
AUDIO_FOLDER = os.path.join("static", "audio")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(AUDIO_FOLDER, exist_ok=True)

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()


# =========================================================
# NEURAL VOICES
# =========================================================

BENGALI_VOICE = "bn-IN-TanishaaNeural"
HINDI_VOICE = "hi-IN-SwaraNeural"
ENGLISH_VOICE = "en-IN-NeerjaNeural"


# =========================================================
# NATURAL NARRATION SETTINGS
# =========================================================

# A slightly slower pace for accessibility.
# These settings do not guarantee human-level emotion.
RATE = "-5%"
VOLUME = "+0%"
PITCH = "+0Hz"

# Keep chunks reasonably sized for reliable TTS.
MAX_CHARS = 1000
MAX_RETRIES = 3

# Silence inserted between paragraphs.
PARAGRAPH_PAUSE_SECONDS = 0.65

# A small pause between consecutive sentences.
SENTENCE_PAUSE_SECONDS = 0.18

MIN_AUDIO_BYTES = 1000


# =========================================================
# DOCUMENT BLOCK READER
# Preserves paragraph/table order in the DOCX body.
# =========================================================

def iter_block_items(parent):

    if isinstance(parent, _Document):
        parent_element = parent.element.body

    elif isinstance(parent, _Cell):
        parent_element = parent._tc

    else:
        raise ValueError("Unsupported DOCX parent type")

    for child in parent_element.iterchildren():

        if isinstance(child, CT_P):
            yield Paragraph(child, parent)

        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =========================================================
# READ DOCX
# =========================================================

def read_docx(file_path):

    document = Document(file_path)
    paragraphs = []

    for block in iter_block_items(document):

        if isinstance(block, Paragraph):

            text = block.text.strip()

            if text:
                paragraphs.append(text)

        elif isinstance(block, Table):

            for row in block.rows:

                cell_texts = []
                seen_cells = set()

                for cell in row.cells:

                    # Avoid repeating merged cells.
                    cell_id = id(cell._tc)

                    if cell_id in seen_cells:
                        continue

                    seen_cells.add(cell_id)

                    text = cell.text.strip()

                    if text:
                        cell_texts.append(text)

                if cell_texts:
                    paragraphs.append(
                        ". ".join(cell_texts)
                    )

    return paragraphs


# =========================================================
# CLEAN TEXT
# Preserve punctuation, numbers and meaningful symbols.
# =========================================================

def clean_text(text):

    text = text.replace("\u00a0", " ")
    text = text.replace("\t", " ")

    # Remove decorative bullets, not the accompanying words.
    text = re.sub(r"[•●▪■◆◇]", "", text)

    # Collapse repeated whitespace.
    text = re.sub(r"\s+", " ", text)

    # Reduce excessive punctuation while preserving emphasis.
    text = re.sub(r"!{2,}", "!", text)
    text = re.sub(r"\?{2,}", "?", text)
    text = re.sub(r"\.{4,}", "...", text)

    return text.strip()


# =========================================================
# LANGUAGE DETECTION
#
# Detects Bengali and Hindi scripts first.
# Latin-script sentences use the English voice.
#
# Limitation:
# Romanized Bengali/Hindi cannot be reliably distinguished
# from English using script detection alone.
# =========================================================

def detect_language(text, fallback="en"):

    bengali_count = len(
        re.findall(r"[\u0980-\u09FF]", text)
    )

    hindi_count = len(
        re.findall(r"[\u0900-\u097F]", text)
    )

    english_count = len(
        re.findall(r"[A-Za-z]", text)
    )

    if bengali_count > 0 and bengali_count >= hindi_count:
        return "bn"

    if hindi_count > 0:
        return "hi"

    if english_count > 0:
        return "en"

    return fallback


# =========================================================
# VOICE SELECTION
# =========================================================

def get_voice(language):

    voices = {
        "bn": BENGALI_VOICE,
        "hi": HINDI_VOICE,
        "en": ENGLISH_VOICE,
    }

    return voices.get(language, ENGLISH_VOICE)


# =========================================================
# NATURAL TEXT PREPARATION
# =========================================================

def prepare_text(text, language):

    text = clean_text(text)

    if not text:
        return ""

    # Add an ending only if punctuation is missing.
    if not re.search(
        r"""[.!?।॥…]["'”’)\]]*$""",
        text
    ):

        if language in ("bn", "hi"):
            text += "।"
        else:
            text += "."

    return text


# =========================================================
# SENTENCE SPLITTING
#
# Retains sentence-ending punctuation so TTS can use it
# to produce more natural intonation.
# =========================================================

def split_sentences(text):

    text = text.strip()

    if not text:
        return []

    # Avoid splitting at ordinary decimal points.
    protected = re.sub(
        r"(?<=\d)\.(?=\d)",
        "<DECIMAL_POINT>",
        text
    )

    # Split after sentence-ending punctuation.
    # This is a practical rule, not a full NLP parser.
    sentences = re.split(
        r'(?<=[.!?।॥…])\s+',
        protected
    )

    result = []

    for sentence in sentences:

        sentence = sentence.replace(
            "<DECIMAL_POINT>", "."
        ).strip()

        if sentence:
            result.append(sentence)

    return result


# =========================================================
# BUILD LANGUAGE-AWARE CHUNKS
#
# Keeps consecutive sentences of the same detected
# language together, avoiding a new voice for every word.
# =========================================================

def build_chunks(text, max_chars=MAX_CHARS):

    sentences = split_sentences(text)

    if not sentences and text.strip():
        sentences = [text.strip()]

    chunks = []
    current_language = None
    current_text = ""

    def flush():

        nonlocal current_language, current_text

        if current_text.strip():
            chunks.append((
                current_language,
                current_text.strip()
            ))

        current_language = None
        current_text = ""

    for sentence in sentences:

        language = detect_language(sentence)

        # Split unusually long sentences at word boundaries.
        pieces = []

        if len(sentence) <= max_chars:
            pieces = [sentence]

        else:

            words = sentence.split()
            piece = ""

            for word in words:

                # A single exceptionally long token must
                # also be handled without an infinite loop.
                if len(word) > max_chars:

                    if piece:
                        pieces.append(piece)
                        piece = ""

                    for start in range(0, len(word), max_chars):
                        pieces.append(
                            word[start:start + max_chars]
                        )

                    continue

                candidate = (piece + " " + word).strip()

                if len(candidate) <= max_chars:
                    piece = candidate

                else:

                    if piece:
                        pieces.append(piece)

                    piece = word

            if piece:
                pieces.append(piece)

        for piece in pieces:

            # A language change begins a new TTS chunk.
            if current_language != language:
                flush()
                current_language = language

            candidate = (current_text + " " + piece).strip()

            if len(candidate) <= max_chars:
                current_text = candidate

            else:
                flush()
                current_language = language
                current_text = piece

    flush()

    return chunks


# =========================================================
# AUDIO VALIDATION
# =========================================================

def audio_file_is_valid(file_path):

    return (
        os.path.isfile(file_path)
        and os.path.getsize(file_path) > MIN_AUDIO_BYTES
    )


def validate_audio_with_ffmpeg(file_path):

    if not audio_file_is_valid(file_path):
        raise RuntimeError(
            f"Audio file missing or too small: {file_path}"
        )

    command = [
        FFMPEG,
        "-v", "error",
        "-i", file_path,
        "-f", "null",
        "-"
    ]

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            "Audio validation failed: "
            + result.stderr[-1500:]
        )


# =========================================================
# GENERATE ONE TTS CHUNK WITH RETRIES
# A failed chunk stops that article from being published.
# =========================================================

async def generate_chunk(text, output_file, language):

    voice = get_voice(language)
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):

        try:

            if os.path.exists(output_file):
                os.remove(output_file)

            print(
                f"      Voice: {voice} | "
                f"Attempt {attempt}/{MAX_RETRIES}"
            )

            communicate = edge_tts.Communicate(
                text=text,
                voice=voice,
                rate=RATE,
                volume=VOLUME,
                pitch=PITCH
            )

            await communicate.save(output_file)

            validate_audio_with_ffmpeg(output_file)

            print("      Audio chunk verified.")
            return

        except Exception as error:

            last_error = error

            print(
                f"      TTS attempt {attempt} failed: {error}"
            )

            try:

                if os.path.exists(output_file):
                    os.remove(output_file)

            except OSError:
                pass

            if attempt < MAX_RETRIES:
                await asyncio.sleep(attempt * 2)

    raise RuntimeError(
        f"Chunk failed after {MAX_RETRIES} attempts: "
        f"{last_error}"
    )


# =========================================================
# GENERATE SILENCE USING FFMPEG
# =========================================================

def generate_silence(output_file, duration_seconds):

    command = [
        FFMPEG,
        "-y",
        "-f", "lavfi",
        "-i", "anullsrc=r=24000:cl=mono",
        "-t", str(duration_seconds),
        "-c:a", "libmp3lame",
        "-b:a", "128k",
        output_file
    ]

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            "Silence generation failed: "
            + result.stderr[-1500:]
        )

    validate_audio_with_ffmpeg(output_file)


# =========================================================
# MERGE AUDIO FILES
# =========================================================

def merge_audio_files(audio_files, output_file):

    if not audio_files:
        raise RuntimeError("No audio files to merge.")

    concat_file = os.path.join(
        os.path.dirname(output_file),
        "concat_list.txt"
    )

    try:

        with open(
            concat_file,
            "w",
            encoding="utf-8"
        ) as file:

            for audio_file in audio_files:

                absolute_path = os.path.abspath(
                    audio_file
                ).replace("\\", "/")

                # Escape single quotes for concat-list syntax.
                absolute_path = absolute_path.replace(
                    "'", "'\\''"
                )

                file.write(
                    f"file '{absolute_path}'\n"
                )

        command = [
            FFMPEG,
            "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", concat_file,
            "-vn",
            "-c:a", "libmp3lame",
            "-b:a", "128k",
            output_file
        ]

        print("Merging audio with FFmpeg...")

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        if result.returncode != 0:
            raise RuntimeError(
                "FFmpeg merge failed: "
                + result.stderr[-3000:]
            )

        validate_audio_with_ffmpeg(output_file)

    finally:

        try:

            if os.path.exists(concat_file):
                os.remove(concat_file)

        except OSError:
            pass


# =========================================================
# OUTPUT FILENAME
# =========================================================

def make_output_filename(filename):

    base_name = os.path.splitext(filename)[0]

    base_name = re.sub(
        r'[\\/*?:"<>|]',
        "_",
        base_name
    )

    base_name = re.sub(
        r"\s+",
        "_",
        base_name
    ).strip("._ ")

    if not base_name:
        base_name = "article"

    return base_name + ".mp3"


# =========================================================
# PROCESS ONE DOCX ARTICLE
# =========================================================

async def process_document(filename):

    print("\n" + "=" * 65)
    print("PROCESSING ARTICLE:", filename)
    print("=" * 65)

    file_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )

    paragraphs = read_docx(file_path)

    paragraphs = [
        clean_text(text)
        for text in paragraphs
        if clean_text(text)
    ]

    if not paragraphs:
        raise RuntimeError(
            "No readable text found in the DOCX body."
        )

    print("Text blocks:", len(paragraphs))

    temp_directory = tempfile.mkdtemp(
        prefix="wall_magazine_",
        dir=AUDIO_FOLDER
    )

    audio_sequence = []
    chunk_number = 0

    try:

        for paragraph_index, paragraph in enumerate(
            paragraphs,
            start=1
        ):

            print(
                f"\nParagraph {paragraph_index}/"
                f"{len(paragraphs)}"
            )

            chunks = build_chunks(paragraph)

            if not chunks:
                continue

            for language, chunk_text in chunks:

                chunk_number += 1

                chunk_file = os.path.join(
                    temp_directory,
                    f"chunk_{chunk_number:05d}.mp3"
                )

                print(
                    f"  Chunk {chunk_number} | "
                    f"Language: {language} | "
                    f"Characters: {len(chunk_text)}"
                )

                await generate_chunk(
                    prepare_text(chunk_text, language),
                    chunk_file,
                    language
                )

                audio_sequence.append(chunk_file)

                # Short pauses between TTS chunks.
                # Sentence punctuation itself also affects
                # the voice's natural rhythm.
                if SENTENCE_PAUSE_SECONDS > 0:

                    pause_file = os.path.join(
                        temp_directory,
                        f"pause_{chunk_number:05d}.mp3"
                    )

                    generate_silence(
                        pause_file,
                        SENTENCE_PAUSE_SECONDS
                    )

                    audio_sequence.append(pause_file)

            # A longer pause between document paragraphs.
            if paragraph_index < len(paragraphs):

                pause_file = os.path.join(
                    temp_directory,
                    f"paragraph_pause_{paragraph_index:05d}.mp3"
                )

                generate_silence(
                    pause_file,
                    PARAGRAPH_PAUSE_SECONDS
                )

                audio_sequence.append(pause_file)

        if not audio_sequence:
            raise RuntimeError(
                "No audio chunks were generated."
            )

        # Remove the trailing pause, if present.
        last_file = audio_sequence[-1]

        if os.path.basename(last_file).startswith(
            "paragraph_pause_"
        ) or os.path.basename(last_file).startswith("pause_"):

            audio_sequence.pop()

        if not audio_sequence:
            raise RuntimeError("No speech audio was generated.")

        output_filename = make_output_filename(filename)

        final_output = os.path.join(
            AUDIO_FOLDER,
            output_filename
        )

        # Build the complete file separately. A previous
        # valid MP3 remains untouched if generation fails.
        temporary_output = os.path.join(
            temp_directory,
            "complete_article.mp3"
        )

        merge_audio_files(
            audio_sequence,
            temporary_output
        )

        # Only publish after successful validation.
        os.replace(
            temporary_output,
            final_output
        )

        size_mb = (
            os.path.getsize(final_output)
            / (1024 * 1024)
        )

        print("\n" + "=" * 65)
        print("ARTICLE AUDIO COMPLETED")
        print("Output:", final_output)
        print(f"Size: {size_mb:.2f} MB")
        print("=" * 65)

        return final_output

    finally:

        shutil.rmtree(
            temp_directory,
            ignore_errors=True
        )


# =========================================================
# MAIN
# =========================================================

async def main():

    print("\n" + "=" * 65)
    print("       WALL MAGAZINE NATURAL TTS")
    print("=" * 65)

    print("Bengali:", BENGALI_VOICE)
    print("Hindi  :", HINDI_VOICE)
    print("English:", ENGLISH_VOICE)
    print("FFmpeg :", FFMPEG)

    files = sorted(
        filename
        for filename in os.listdir(UPLOAD_FOLDER)
        if (
            filename.lower().endswith(".docx")
            and not filename.startswith("~$")
        )
    )

    if not files:
        print("No DOCX files found in uploads.")
        return

    print(f"\nFound {len(files)} DOCX files.")

    successful = []
    failed = []

    for index, filename in enumerate(files, start=1):

        print(f"\nARTICLE {index}/{len(files)}")

        try:

            output = await process_document(filename)

            successful.append((filename, output))

        except Exception as error:

            print("\nARTICLE FAILED:", filename)
            print("Reason:", error)

            failed.append((filename, str(error)))

            # Continue with the next article.
            continue

    print("\n" + "=" * 65)
    print("PROCESSING SUMMARY")
    print("=" * 65)

    print("Total:", len(files))
    print("Successful:", len(successful))
    print("Failed:", len(failed))

    if successful:

        print("\nSUCCESSFUL ARTICLES:")

        for filename, output in successful:
            print(f"{filename} -> {output}")

    if failed:

        print("\nFAILED ARTICLES:")

        for filename, error in failed:
            print(f"{filename}: {error}")

    print("\nCheck static/audio for the completed MP3 files.")


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    asyncio.run(main())
