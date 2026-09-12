# ================================================================
# AI EDUCATOR PRO
# ================================================================
# Single-file desktop application
#
# FEATURES
# ---------------------------------------------------------------
# 1. SQLite study library database
# 2. OpenAI API integration (text + image + PDF analysis)
# 3. Auto-install of required Python packages on startup
# 4. Read TXT / MD / CSV / DOCX / PDF
# 5. Image OCR and scanned-PDF OCR
# 6. AI Summary / AI Explanation / AI Analyze Source
# 7. Image → PDF conversion
# 8. Read Aloud (Windows speech)
# 9. Export to Word / Excel / PDF
# 10. Search, multiple learner profiles, settings, API key, model selection
#
# HOW TO RUN
# ---------------------------------------------------------------
#   python AI_Educator_Pro.py
#
# On first run the app installs missing Python packages:
#   pillow, pytesseract, python-docx, pypdf, pymupdf,
#   xlsxwriter, reportlab, openai
#
# For local OCR you also need the Tesseract OCR program:
#   Windows: https://github.com/UB-Mannheim/tesseract/wiki
#   macOS:   brew install tesseract
#   Linux:   sudo apt install tesseract-ocr
#
# OpenAI key: Settings → API Settings, or set OPENAI_API_KEY
# ================================================================

import sys
import os
import sqlite3
import subprocess
import importlib
import threading
import json
import base64
import re
import time
import mimetypes
import shutil
from pathlib import Path
from datetime import datetime

# ================================================================
# AUTO INSTALL PYTHON PACKAGES
# ================================================================
# Installs only what is needed for reading, OCR, PDF, export, AI.

REQUIRED_PACKAGES = {
    "PIL": "pillow",
    "pytesseract": "pytesseract",
    "docx": "python-docx",
    "pypdf": "pypdf",
    "fitz": "pymupdf",
    "xlsxwriter": "xlsxwriter",
    "reportlab": "reportlab",
    "openai": "openai",
}


def _module_available(module_name):
    try:
        importlib.import_module(module_name)
        return True
    except Exception:
        return False


def install_required_packages():
    """Install missing packages, then verify imports."""
    missing = []
    for module_name, package_name in REQUIRED_PACKAGES.items():
        if not _module_available(module_name):
            missing.append(package_name)

    if not missing:
        print("All required Python packages are already installed.")
        return True

    print("=" * 60)
    print("AI Educator Pro - installing required packages")
    print("Needed for: reading, OCR, PDF analysis, AI, export")
    print("=" * 60)
    print("Packages:", ", ".join(missing))
    print("=" * 60)

    ok = True
    for package in missing:
        print(f"Installing {package} ...")
        try:
            result = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "--disable-pip-version-check",
                    "--upgrade",
                    package,
                ],
                capture_output=True,
                text=True,
                timeout=600,
            )
            if result.returncode != 0:
                ok = False
                print(f"FAILED: {package}")
                if result.stderr:
                    print(result.stderr[-2000:])
            else:
                print(f"OK: {package}")
        except Exception as exc:
            ok = False
            print(f"Installation error for {package}: {exc}")

    # Clear cached failed imports so re-import can succeed.
    for module_name in list(REQUIRED_PACKAGES.keys()):
        root = module_name.split(".")[0]
        if root in sys.modules:
            # Keep already-working modules; only drop if previously missing.
            pass

    still_missing = [
        package
        for module_name, package in REQUIRED_PACKAGES.items()
        if not _module_available(module_name)
    ]

    print("=" * 60)
    if still_missing:
        print("Some packages could not be imported after install:")
        print(", ".join(still_missing))
        print("Try manually:")
        print(f"  {sys.executable} -m pip install " + " ".join(still_missing))
        ok = False
    else:
        print("Package installation completed successfully.")
    print("=" * 60)
    return ok


install_required_packages()


# ================================================================
# IMPORTS
# ================================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    from PIL import Image, ImageTk, ImageEnhance, ImageFilter
except Exception as exc:
    raise SystemExit(
        "Pillow (PIL) is required but could not be imported.\n"
        f"Install with: {sys.executable} -m pip install pillow\n"
        f"Details: {exc}"
    )

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except Exception:
    Document = None
    Pt = None
    WD_ALIGN_PARAGRAPH = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import fitz  # PyMuPDF - scanned PDF page rendering for OCR
except Exception:
    fitz = None

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        PageBreak,
    )
except Exception:
    SimpleDocTemplate = None
    colors = None
    TA_CENTER = None
    A4 = None
    getSampleStyleSheet = None
    ParagraphStyle = None
    mm = None
    Paragraph = None
    Spacer = None
    Table = None
    TableStyle = None
    PageBreak = None


def configure_tesseract_path():
    """Find the Tesseract OCR executable and configure pytesseract."""
    if pytesseract is None:
        return False

    found = shutil.which("tesseract")
    if found and os.path.isfile(found):
        pytesseract.pytesseract.tesseract_cmd = found
        return True

    candidates = []
    if os.name == "nt":
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        local = os.environ.get("LOCALAPPDATA", "")
        user = os.path.expanduser("~")
        candidates = [
            os.path.join(pf, "Tesseract-OCR", "tesseract.exe"),
            os.path.join(pf86, "Tesseract-OCR", "tesseract.exe"),
            os.path.join(local, "Tesseract-OCR", "tesseract.exe"),
            os.path.join(local, "Programs", "Tesseract-OCR", "tesseract.exe"),
            os.path.join(user, "AppData", "Local", "Programs", "Tesseract-OCR", "tesseract.exe"),
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
    else:
        candidates = [
            "/usr/bin/tesseract",
            "/usr/local/bin/tesseract",
            "/opt/homebrew/bin/tesseract",
        ]

    for executable in candidates:
        if executable and os.path.isfile(executable):
            pytesseract.pytesseract.tesseract_cmd = executable
            return True
    return False


def install_tesseract_engine():
    """Try to install the native Tesseract OCR engine on Windows using winget.

    The pytesseract Python package is installed by install_required_packages().
    OCR also needs the separate native Tesseract program.
    """
    if configure_tesseract_path():
        return True, "Tesseract OCR is already installed."

    if os.name != "nt":
        return False, (
            "Automatic Tesseract installation is currently available on Windows only.\n\n"
            "Linux: sudo apt install tesseract-ocr\n"
            "macOS: brew install tesseract"
        )

    winget = shutil.which("winget")
    if not winget:
        return False, (
            "Windows Package Manager (winget) was not found.\n\n"
            "Install Tesseract OCR manually, then restart the application."
        )

    package_ids = ["tesseract-ocr.tesseract", "UB-Mannheim.TesseractOCR"]
    errors = []
    for package_id in package_ids:
        try:
            result = subprocess.run(
                [
                    winget, "install", "--id", package_id, "--exact",
                    "--accept-package-agreements", "--accept-source-agreements",
                    "--silent",
                ],
                capture_output=True,
                text=True,
                timeout=900,
            )
            if result.returncode == 0:
                # Refresh process PATH with common winget link directory.
                winget_links = os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Microsoft", "WinGet", "Links"
                )
                if winget_links and winget_links not in os.environ.get("PATH", ""):
                    os.environ["PATH"] = winget_links + os.pathsep + os.environ.get("PATH", "")
                if configure_tesseract_path():
                    return True, "Tesseract OCR was installed successfully."
                return True, (
                    "Tesseract installation completed. Restart this application before using OCR."
                )
            errors.append((result.stderr or result.stdout or f"Exit code {result.returncode}")[-1200:])
        except Exception as exc:
            errors.append(str(exc))

    return False, "Tesseract OCR could not be installed automatically.\n\n" + "\n".join(errors[-2:])


configure_tesseract_path()


# ================================================================
# APPLICATION PATHS
# ================================================================

APP_FOLDER = os.path.join(
    os.path.expanduser("~"),
    "AI_Educator_Pro"
)

os.makedirs(
    APP_FOLDER,
    exist_ok=True
)

DATABASE_FILE = os.path.join(
    APP_FOLDER,
    "educator.db"
)


# ================================================================
# FILE READING / TEXT EXTRACTION
# ================================================================

def ocr_image(image):
    """Run local OCR after light preprocessing."""
    if pytesseract is None:
        raise Exception("pytesseract is not installed.")
    gray = image.convert("L")
    gray = ImageEnhance.Contrast(gray).enhance(1.7)
    gray = gray.filter(ImageFilter.SHARPEN)
    return pytesseract.image_to_string(gray, config="--psm 6").strip()

def extract_text_from_file(file_path):
    if not file_path:
        return ""
    path = Path(file_path)
    extension = path.suffix.lower()
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if extension in [".txt", ".md", ".csv", ".log"]:
        for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
            try:
                return path.read_text(encoding=encoding).strip()
            except UnicodeDecodeError:
                continue
        return path.read_text(errors="ignore").strip()
    if extension == ".docx":
        if Document is None:
            raise Exception("python-docx is not installed.")
        doc = Document(file_path)
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                values = [cell.text.strip() for cell in row.cells]
                if any(values):
                    parts.append(" | ".join(values))
        return "\n".join(parts).strip()
    if extension == ".pdf":
        if PdfReader is None:
            raise Exception("pypdf is not installed.")
        reader = PdfReader(file_path)
        pages = []
        for page_number, page in enumerate(reader.pages, start=1):
            text = (page.extract_text() or "").strip()
            if text:
                pages.append(f"[Page {page_number}]\n{text}")
        extracted = "\n\n".join(pages).strip()
        if not extracted and fitz is not None and pytesseract is not None and configure_tesseract_path():
            pdf = fitz.open(file_path)
            try:
                ocr_pages = []
                for page_number, page in enumerate(pdf, start=1):
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text = ocr_image(image)
                    if text:
                        ocr_pages.append(f"[Page {page_number}]\n{text}")
                extracted = "\n\n".join(ocr_pages).strip()
            finally:
                pdf.close()
        return extracted
    if extension in [".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"]:
        return ocr_image(Image.open(file_path))
    raise Exception(f"Unsupported file type: {extension}")

def clean_text_for_export(text):
    return re.sub(r"\n{3,}", "\n\n", (text or "")).strip()


# DATABASE MANAGER
# ================================================================

class Database:

    def __init__(self):

        self.db_file = DATABASE_FILE

        self.create_tables()

    # ------------------------------------------------------------
    # CONNECTION
    # ------------------------------------------------------------

    def connect(self):

        connection = sqlite3.connect(
            self.db_file
        )

        connection.row_factory = sqlite3.Row

        return connection

    # ------------------------------------------------------------
    # TABLES
    # ------------------------------------------------------------

    def create_tables(self):

        conn = self.connect()

        cursor = conn.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                path TEXT,
                file_type TEXT,
                text TEXT DEFAULT '',
                summary TEXT DEFAULT '',
                explanation TEXT DEFAULT '',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS profile (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                name TEXT DEFAULT '',
                age TEXT DEFAULT '',
                user_type TEXT DEFAULT 'Student',
                qualification TEXT DEFAULT '',
                hobbies TEXT DEFAULT '',
                mood TEXT DEFAULT 'Happy'
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS learner_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT DEFAULT '',
                age TEXT DEFAULT '',
                user_type TEXT DEFAULT 'Student',
                qualification TEXT DEFAULT '',
                hobbies TEXT DEFAULT '',
                mood TEXT DEFAULT 'Happy',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS exports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER,
                format TEXT,
                output_path TEXT,
                created_at TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER,
                action TEXT,
                result TEXT,
                created_at TEXT
            )
            """
        )

        cursor.execute(
            """
            INSERT OR IGNORE INTO profile
            (id, name, age, user_type, qualification, hobbies, mood)
            VALUES
            (1, '', '', 'Student', '', '', 'Happy')
            """
        )

        # Migrate the original single profile on the first upgraded run.
        profile_count = cursor.execute(
            "SELECT COUNT(*) AS count FROM learner_profiles"
        ).fetchone()[0]
        if profile_count == 0:
            old = cursor.execute("SELECT * FROM profile WHERE id=1").fetchone()
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                """
                INSERT INTO learner_profiles
                (name, age, user_type, qualification, hobbies, mood, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (old["name"] if old and old["name"] else "Default Learner"),
                    (old["age"] if old else ""),
                    (old["user_type"] if old else "Student"),
                    (old["qualification"] if old else ""),
                    (old["hobbies"] if old else ""),
                    (old["mood"] if old else "Happy"),
                    now,
                    now,
                ),
            )
            cursor.execute(
                "INSERT OR REPLACE INTO settings(key, value) VALUES('active_profile_id', ?)",
                (str(cursor.lastrowid),),
            )

        conn.commit()

        conn.close()

    # ------------------------------------------------------------
    # SETTINGS
    # ------------------------------------------------------------

    def set_setting(
        self,
        key,
        value
    ):

        conn = self.connect()

        conn.execute(
            """
            INSERT INTO settings(key, value)
            VALUES (?, ?)
            ON CONFLICT(key)
            DO UPDATE SET value=excluded.value
            """,
            (
                key,
                value
            )
        )

        conn.commit()
        conn.close()

    def get_setting(
        self,
        key,
        default=""
    ):

        conn = self.connect()

        row = conn.execute(
            """
            SELECT value
            FROM settings
            WHERE key=?
            """,
            (key,)
        ).fetchone()

        conn.close()

        if row:
            return row["value"]

        return default

    # ------------------------------------------------------------
    # DOCUMENTS
    # ------------------------------------------------------------

    def add_document(
        self,
        name,
        path,
        file_type,
        text=""
    ):

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn = self.connect()

        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO documents
            (
                name,
                path,
                file_type,
                text,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                path,
                file_type,
                text,
                now,
                now
            )
        )

        document_id = cursor.lastrowid

        conn.commit()
        conn.close()

        return document_id

    def update_document_text(
        self,
        document_id,
        text
    ):

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn = self.connect()

        conn.execute(
            """
            UPDATE documents
            SET text=?,
                updated_at=?
            WHERE id=?
            """,
            (
                text,
                now,
                document_id
            )
        )

        conn.commit()
        conn.close()

    def update_ai_result(
        self,
        document_id,
        field,
        value
    ):

        allowed = [
            "summary",
            "explanation"
        ]

        if field not in allowed:
            return

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn = self.connect()

        conn.execute(
            f"""
            UPDATE documents
            SET {field}=?,
                updated_at=?
            WHERE id=?
            """,
            (
                value,
                now,
                document_id
            )
        )

        conn.commit()
        conn.close()

    def get_document(
        self,
        document_id
    ):

        conn = self.connect()

        row = conn.execute(
            """
            SELECT *
            FROM documents
            WHERE id=?
            """,
            (document_id,)
        ).fetchone()

        conn.close()

        return row

    def get_documents(
        self,
        search=""
    ):

        conn = self.connect()

        if search.strip():

            rows = conn.execute(
                """
                SELECT *
                FROM documents
                WHERE
                    name LIKE ?
                    OR text LIKE ?
                    OR summary LIKE ?
                    OR explanation LIKE ?
                ORDER BY id DESC
                """,
                (
                    f"%{search}%",
                    f"%{search}%",
                    f"%{search}%",
                    f"%{search}%"
                )
            ).fetchall()

        else:

            rows = conn.execute(
                """
                SELECT *
                FROM documents
                ORDER BY id DESC
                """
            ).fetchall()

        conn.close()

        return rows

    def delete_document(
        self,
        document_id
    ):

        conn = self.connect()

        conn.execute(
            """
            DELETE FROM documents
            WHERE id=?
            """,
            (document_id,)
        )

        conn.commit()
        conn.close()

    # ------------------------------------------------------------
    # HISTORY
    # ------------------------------------------------------------

    def add_history(
        self,
        document_id,
        action,
        result
    ):

        conn = self.connect()

        conn.execute(
            """
            INSERT INTO history
            (
                document_id,
                action,
                result,
                created_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                document_id,
                action,
                result,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )
        )

        conn.commit()
        conn.close()

    # ------------------------------------------------------------
    # EXPORTS
    # ------------------------------------------------------------

    def add_export(self, document_id, export_format, output_path):
        conn = self.connect()
        conn.execute(
            """
            INSERT INTO exports(document_id, format, output_path, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (document_id, export_format, output_path, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        conn.close()

    def get_history(self, document_id):
        conn = self.connect()
        rows = conn.execute(
            """
            SELECT * FROM history
            WHERE document_id=?
            ORDER BY id DESC
            """,
            (document_id,)
        ).fetchall()
        conn.close()
        return rows

    def get_exports(self, document_id=None):
        conn = self.connect()
        if document_id is None:
            rows = conn.execute("SELECT * FROM exports ORDER BY id DESC").fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM exports WHERE document_id=? ORDER BY id DESC",
                (document_id,)
            ).fetchall()
        conn.close()
        return rows

    # ------------------------------------------------------------
    # PROFILE
    # ------------------------------------------------------------

    def get_profiles(self):
        conn = self.connect()
        rows = conn.execute(
            "SELECT * FROM learner_profiles ORDER BY name COLLATE NOCASE, id"
        ).fetchall()
        conn.close()
        return rows

    def get_profile(self, profile_id=None):
        conn = self.connect()
        if profile_id is None:
            value = conn.execute(
                "SELECT value FROM settings WHERE key='active_profile_id'"
            ).fetchone()
            profile_id = int(value["value"]) if value and str(value["value"]).isdigit() else None

        row = None
        if profile_id is not None:
            row = conn.execute(
                "SELECT * FROM learner_profiles WHERE id=?", (profile_id,)
            ).fetchone()
        if row is None:
            row = conn.execute(
                "SELECT * FROM learner_profiles ORDER BY id LIMIT 1"
            ).fetchone()
        conn.close()
        return row

    def set_active_profile(self, profile_id):
        self.set_setting("active_profile_id", str(profile_id))

    def add_profile(self, name="New Learner"):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = self.connect()
        cursor = conn.execute(
            """
            INSERT INTO learner_profiles
            (name, age, user_type, qualification, hobbies, mood, created_at, updated_at)
            VALUES (?, '', 'Student', '', '', 'Happy', ?, ?)
            """,
            (name, now, now),
        )
        profile_id = cursor.lastrowid
        conn.commit()
        conn.close()
        self.set_active_profile(profile_id)
        return profile_id

    def save_profile(self, name, age, user_type, qualification, hobbies, mood, profile_id=None):
        if profile_id is None:
            active = self.get_profile()
            profile_id = active["id"] if active else self.add_profile(name or "New Learner")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = self.connect()
        conn.execute(
            """
            UPDATE learner_profiles
            SET name=?, age=?, user_type=?, qualification=?, hobbies=?, mood=?, updated_at=?
            WHERE id=?
            """,
            (name, age, user_type, qualification, hobbies, mood, now, profile_id),
        )
        conn.commit()
        conn.close()
        self.set_active_profile(profile_id)

    def delete_profile(self, profile_id):
        profiles = self.get_profiles()
        if len(profiles) <= 1:
            raise ValueError("At least one learner profile must remain.")
        conn = self.connect()
        conn.execute("DELETE FROM learner_profiles WHERE id=?", (profile_id,))
        replacement = conn.execute(
            "SELECT id FROM learner_profiles ORDER BY id LIMIT 1"
        ).fetchone()
        conn.commit()
        conn.close()
        if replacement:
            self.set_active_profile(replacement["id"])



# ================================================================
# OPENAI AI MANAGER
# ================================================================

class AIManager:
    def __init__(self, database):
        self.database = database

    def get_client(self):
        if OpenAI is None:
            raise Exception("OpenAI package is not installed. Run: pip install -U openai")
        api_key = self.database.get_setting("openai_api_key") or os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise Exception("OpenAI API key is not configured. Go to Settings → API Settings or set OPENAI_API_KEY.")
        return OpenAI(api_key=api_key)

    def get_model(self):
        return self.database.get_setting("openai_model", "gpt-5")

    def profile_context(self):
        profile = self.database.get_profile()
        if not profile:
            return ""
        return (
            "ACTIVE LEARNER PROFILE:\n"
            f"Name: {profile['name'] or 'Learner'}\n"
            f"Age: {profile['age'] or 'Not specified'}\n"
            f"Type: {profile['user_type'] or 'Student'}\n"
            f"Qualification: {profile['qualification'] or 'Not specified'}\n"
            f"Hobbies: {profile['hobbies'] or 'Not specified'}\n"
            f"Preferred style/mood: {profile['mood'] or 'Clear'}\n"
            "Adapt the wording and examples to this learner while preserving accuracy.\n\n"
        )

    def generate(self, instruction, text):
        client = self.get_client()
        prompt = (
            self.profile_context()
            + f"{instruction}\n\nSTUDY MATERIAL:\n"
            + "--------------------------------\n"
            + f"{(text or '')[:100000]}\n"
            + "--------------------------------\n"
        )
        response = client.responses.create(model=self.get_model(), input=prompt)
        return response.output_text or ""

    def analyze_file(self, file_path, instruction):
        """Analyze an image or PDF directly with the OpenAI Responses API."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(str(path))
        client = self.get_client()
        instruction = self.profile_context() + instruction
        extension = path.suffix.lower()
        if extension == ".pdf":
            with open(path, "rb") as handle:
                uploaded = client.files.create(file=handle, purpose="user_data")
            content = [{"type": "input_text", "text": instruction}, {"type": "input_file", "file_id": uploaded.id}]
        elif extension in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}:
            mime = mimetypes.guess_type(path.name)[0] or "image/jpeg"
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            content = [
                {"type": "input_text", "text": instruction},
                {"type": "input_image", "image_url": f"data:{mime};base64,{encoded}", "detail": "high"},
            ]
        else:
            return self.generate(instruction, extract_text_from_file(file_path))
        response = client.responses.create(model=self.get_model(), input=[{"role": "user", "content": content}])
        return response.output_text or ""


# ================================================================
# WINDOWS TEXT TO SPEECH
# ================================================================

def windows_speak(
    text
):

    if not text.strip():
        return

    if os.name != "nt":

        messagebox.showinfo(
            "Read Aloud",
            "Windows built-in speech is used "
            "by this application."
        )

        return

    def worker():

        try:

            text = text[:10000]

            safe = (
                text
                .replace(
                    "`",
                    "``"
                )
                .replace(
                    '"',
                    '`"'
                )
                .replace(
                    "$",
                    "`$"
                )
            )

            command = (
                "Add-Type -AssemblyName System.Speech; "
                "$s = New-Object "
                "System.Speech.Synthesis.SpeechSynthesizer; "
                f'$s.Speak("{safe}");'
            )

            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    command
                ],
                creationflags=subprocess.CREATE_NO_WINDOW
            )

        except Exception as e:

            print(
                "Speech error:",
                e
            )

    threading.Thread(
        target=worker,
        daemon=True
    ).start()


# ================================================================
# MAIN APPLICATION
# ================================================================

class EducatorApp:

    def __init__(
        self,
        root
    ):

        self.root = root

        self.root.title(
            "AI Educator Pro"
        )

        self.root.geometry(
            "1450x850"
        )

        self.root.minsize(
            1100,
            700
        )

        self.database = Database()

        self.ai = AIManager(
            self.database
        )

        self.selected_document_id = None

        self.colors = {
            "sidebar": "#173F5F",
            "accent": "#00A6A6",
            "background": "#F5F8FC",
            "white": "#FFFFFF",
            "text": "#263238"
        }

        self.create_style()

        self.create_header()

        self.create_layout()

        self.create_status()

        self.load_library()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close_application
        )

    # ============================================================
    # STYLE
    # ============================================================

    def create_style(self):

        style = ttk.Style()

        try:
            style.theme_use(
                "clam"
            )
        except Exception:
            pass

        style.configure(
            "TButton",
            font=(
                "Segoe UI",
                10
            ),
            padding=8
        )

        style.configure(
            "Treeview",
            rowheight=32,
            font=(
                "Segoe UI",
                10
            )
        )

        style.configure(
            "Treeview.Heading",
            font=(
                "Segoe UI",
                10,
                "bold"
            )
        )

    # ============================================================
    # HEADER
    # ============================================================

    def create_header(self):

        header = tk.Frame(
            self.root,
            bg=self.colors["sidebar"],
            height=75
        )

        header.pack(
            fill="x"
        )

        header.pack_propagate(
            False
        )

        tk.Label(
            header,
            text="🎓 AI EDUCATOR PRO",
            bg=self.colors["sidebar"],
            fg="white",
            font=(
                "Segoe UI",
                23,
                "bold"
            )
        ).pack(
            side="left",
            padx=25
        )

        tk.Label(
            header,
            text="AI-powered study library",
            bg=self.colors["sidebar"],
            fg="white",
            font=(
                "Segoe UI",
                11
            )
        ).pack(
            side="left"
        )

        ttk.Button(
            header,
            text="👤 Profile",
            command=self.open_profile
        ).pack(
            side="right",
            padx=20,
            pady=18
        )

    # ============================================================
    # LAYOUT
    # ============================================================

    def create_layout(self):

        main = tk.Frame(
            self.root,
            bg=self.colors["background"]
        )

        main.pack(
            fill="both",
            expand=True
        )

        # --------------------------------------------------------
        # SIDEBAR
        # --------------------------------------------------------

        sidebar = tk.Frame(
            main,
            bg=self.colors["sidebar"],
            width=230
        )

        sidebar.pack(
            side="left",
            fill="y"
        )

        sidebar.pack_propagate(
            False
        )

        self.side_button(
            sidebar,
            "📁 Add to Library",
            self.add_to_library
        )

        self.side_button(
            sidebar,
            "🔍 OCR Image / Scanned PDF",
            self.ocr_selected
        )

        self.side_button(
            sidebar,
            "🧰 Install / Check OCR",
            self.install_ocr_support
        )

        self.side_button(
            sidebar,
            "🧠 AI Analyze Source",
            self.ai_analyze_source
        )

        self.side_button(
            sidebar,
            "🖼 Image → PDF",
            self.convert_selected_image_to_pdf
        )

        self.side_button(
            sidebar,
            "📝 Save Edited Text",
            self.save_selected_text
        )

        self.side_button(
            sidebar,
            "🤖 AI Summary",
            self.ai_summary
        )

        self.side_button(
            sidebar,
            "🎓 AI Explanation",
            self.ai_explanation
        )

        self.side_button(
            sidebar,
            "🔊 Read Aloud",
            self.read_selected_aloud
        )

        self.side_button(
            sidebar,
            "📄 Export to Word",
            self.export_selected_word
        )

        self.side_button(
            sidebar,
            "📊 Export to Excel",
            self.export_selected_excel
        )

        self.side_button(
            sidebar,
            "📕 Export to PDF",
            self.export_selected_pdf
        )

        self.side_button(
            sidebar,
            "📦 Export All",
            self.export_selected_all
        )

        self.side_button(
            sidebar,
            "🗑 Delete Selected",
            self.delete_selected
        )

        self.side_button(
            sidebar,
            "⚙ Settings",
            self.open_settings
        )

        # --------------------------------------------------------
        # CONTENT
        # --------------------------------------------------------

        content = tk.Frame(
            main,
            bg=self.colors["background"]
        )

        content.pack(
            side="left",
            fill="both",
            expand=True,
            padx=15,
            pady=15
        )

        # --------------------------------------------------------
        # SEARCH BAR
        # --------------------------------------------------------

        search_frame = tk.Frame(
            content,
            bg=self.colors["background"]
        )

        search_frame.pack(
            fill="x"
        )

        tk.Label(
            search_frame,
            text="Search Library:",
            bg=self.colors["background"],
            font=(
                "Segoe UI",
                10,
                "bold"
            )
        ).pack(
            side="left"
        )

        self.search_var = tk.StringVar()

        search_entry = ttk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=45
        )

        search_entry.pack(
            side="left",
            padx=10
        )

        ttk.Button(
            search_frame,
            text="Search",
            command=self.search_library
        ).pack(
            side="left"
        )

        ttk.Button(
            search_frame,
            text="Show All",
            command=self.load_library
        ).pack(
            side="left",
            padx=5
        )

        # --------------------------------------------------------
        # LIBRARY TREE
        # --------------------------------------------------------

        library_frame = tk.LabelFrame(
            content,
            text="📚 Study Library",
            bg="white",
            font=(
                "Segoe UI",
                11,
                "bold"
            )
        )

        library_frame.pack(
            fill="both",
            expand=True,
            pady=(12, 8)
        )

        columns = (
            "id",
            "name",
            "type",
            "updated"
        )

        self.library_tree = ttk.Treeview(
            library_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        self.library_tree.heading(
            "id",
            text="ID"
        )

        self.library_tree.heading(
            "name",
            text="File / Study Material"
        )

        self.library_tree.heading(
            "type",
            text="Type"
        )

        self.library_tree.heading(
            "updated",
            text="Last Updated"
        )

        self.library_tree.column(
            "id",
            width=60,
            anchor="center"
        )

        self.library_tree.column(
            "name",
            width=500
        )

        self.library_tree.column(
            "type",
            width=100
        )

        self.library_tree.column(
            "updated",
            width=180
        )

        self.library_tree.pack(
            side="left",
            fill="both",
            expand=True
        )

        tree_scroll = ttk.Scrollbar(
            library_frame,
            orient="vertical",
            command=self.library_tree.yview
        )

        tree_scroll.pack(
            side="right",
            fill="y"
        )

        self.library_tree.configure(
            yscrollcommand=tree_scroll.set
        )

        self.library_tree.bind(
            "<<TreeviewSelect>>",
            self.library_selected
        )

        self.library_tree.bind(
            "<Double-1>",
            self.library_selected
        )

        # --------------------------------------------------------
        # DOCUMENT DETAILS
        # --------------------------------------------------------

        details = tk.LabelFrame(
            content,
            text="Selected Study Material",
            bg="white",
            font=(
                "Segoe UI",
                11,
                "bold"
            )
        )

        details.pack(
            fill="both",
            expand=True
        )

        self.selected_label = tk.Label(
            details,
            text="No document selected",
            bg="white",
            fg="#777",
            anchor="w",
            font=(
                "Segoe UI",
                11,
                "bold"
            )
        )

        self.selected_label.pack(
            fill="x",
            padx=10,
            pady=5
        )

        self.document_text = tk.Text(
            details,
            wrap="word",
            font=(
                "Segoe UI",
                11
            )
        )

        self.document_text.pack(
            side="left",
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        text_scroll = ttk.Scrollbar(
            details,
            orient="vertical",
            command=self.document_text.yview
        )

        text_scroll.pack(
            side="right",
            fill="y"
        )

        self.document_text.configure(
            yscrollcommand=text_scroll.set
        )

    # ============================================================
    # SIDEBAR BUTTON
    # ============================================================

    def side_button(
        self,
        parent,
        text,
        command
    ):

        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=self.colors["sidebar"],
            fg="white",
            activebackground=self.colors["accent"],
            activeforeground="white",
            relief="flat",
            anchor="w",
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            padx=18,
            pady=13
        )

        button.pack(
            fill="x"
        )

    # ============================================================
    # STATUS
    # ============================================================

    def create_status(self):

        self.status_var = tk.StringVar(
            value="Ready"
        )

        tk.Label(
            self.root,
            textvariable=self.status_var,
            bg="#E8EEF5",
            fg="#333",
            anchor="w",
            padx=10
        ).pack(
            fill="x"
        )

    def status(
        self,
        message
    ):

        self.status_var.set(
            message
        )

        self.root.update_idletasks()

    # ============================================================
    # LIBRARY
    # ============================================================

    def add_to_library(self):

        files = filedialog.askopenfilenames(
            title="Add Study Material",
            filetypes=[
                (
                    "Supported Files",
                    "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf *.docx *.txt"
                ),
                (
                    "Images",
                    "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"
                ),
                (
                    "Documents",
                    "*.pdf *.docx *.txt"
                ),
                (
                    "All Files",
                    "*.*"
                )
            ]
        )

        if not files:
            return

        added = 0

        for file_path in files:

            name = Path(
                file_path
            ).name

            extension = Path(
                file_path
            ).suffix.lower()

            # Avoid duplicates
            existing = self.database.get_documents()

            duplicate = False

            for row in existing:

                if row["path"] == file_path:

                    duplicate = True
                    break

            if duplicate:
                continue

            self.database.add_document(
                name=name,
                path=file_path,
                file_type=extension
            )

            added += 1

        self.load_library()

        messagebox.showinfo(
            "Library",
            f"{added} file(s) added to the library."
        )

    # ============================================================
    # LOAD LIBRARY
    # ============================================================

    def load_library(
        self
    ):

        if not hasattr(
            self,
            "library_tree"
        ):
            return

        for item in self.library_tree.get_children():

            self.library_tree.delete(
                item
            )

        rows = self.database.get_documents()

        for row in rows:

            self.library_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["id"],
                    row["name"],
                    row["file_type"],
                    row["updated_at"]
                )
            )

        self.status(
            f"{len(rows)} document(s) in Library."
        )

    # ============================================================
    # SEARCH
    # ============================================================

    def search_library(self):

        query = self.search_var.get()

        for item in self.library_tree.get_children():

            self.library_tree.delete(
                item
            )

        rows = self.database.get_documents(
            query
        )

        for row in rows:

            self.library_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                values=(
                    row["id"],
                    row["name"],
                    row["file_type"],
                    row["updated_at"]
                )
            )

        self.status(
            f"{len(rows)} result(s) found."
        )

    # ============================================================
    # SELECT LIBRARY DOCUMENT
    # ============================================================

    def library_selected(
        self,
        event=None
    ):

        selected = self.library_tree.selection()

        if not selected:
            return

        document_id = int(
            selected[0]
        )

        self.selected_document_id = (
            document_id
        )

        document = self.database.get_document(
            document_id
        )

        if not document:
            return

        self.selected_label.config(
            text=(
                f"Selected: {document['name']}  |  "
                f"{document['file_type']}"
            ),
            fg=self.colors["text"]
        )

        self.document_text.delete(
            "1.0",
            "end"
        )

        current_text = document["text"] or ""

        if not current_text.strip() and document["path"]:
            try:
                self.status("Reading selected study material...")
                current_text = extract_text_from_file(document["path"])
                if current_text:
                    self.database.update_document_text(document_id, current_text)
                    self.database.add_history(document_id, "READ", "Text extracted from selected file.")
            except Exception as e:
                messagebox.showerror("Read Error", str(e))

        self.document_text.insert(
            "1.0",
            current_text
        )

        if current_text.strip():
            self.status(f"Selected and read: {document['name']}")
        else:
            self.status(f"Selected: {document['name']} - no readable text found")

        # Automatically analyze newly read material when an API key is configured.
        # Existing saved AI results are reused instead of calling the API again.
        if current_text.strip() and self.database.get_setting("openai_api_key"):
            if not (document["summary"] or "").strip():
                self.run_ai(document_id, "summary", current_text)
            if not (document["explanation"] or "").strip():
                self.run_ai(document_id, "explanation", current_text)

    # ============================================================
    # GET SELECTED DOCUMENT
    # ============================================================

    def get_selected_document(self):

        if not self.selected_document_id:

            messagebox.showwarning(
                "Select Document",
                "Please select a document from the Library first."
            )

            return None

        document = self.database.get_document(
            self.selected_document_id
        )

        if not document:

            messagebox.showerror(
                "Error",
                "Selected document could not be found."
            )

            return None

        return document

    # ============================================================
    # SAVE TEXT
    # ============================================================

    def save_selected_text(
        self
    ):

        document = self.get_selected_document()

        if not document:
            return

        text = self.document_text.get(
            "1.0",
            "end"
        ).strip()

        self.database.update_document_text(
            document["id"],
            text
        )

        self.status(
            "Edited text saved to database."
        )

        messagebox.showinfo(
            "Saved",
            "Study text saved successfully."
        )

    # ============================================================
    # OCR
    # ============================================================

    def install_ocr_support(self):
        """Check OCR and offer automatic native-engine installation."""
        if pytesseract is None:
            messagebox.showerror(
                "OCR Setup",
                "The pytesseract Python package could not be imported. Restart the app "
                "while connected to the internet so it can be installed automatically."
            )
            return
        if configure_tesseract_path():
            try:
                version = pytesseract.get_tesseract_version()
            except Exception:
                version = "installed"
            messagebox.showinfo("OCR Setup", f"Tesseract OCR is ready. Version: {version}")
            return
        if not messagebox.askyesno(
            "Install Tesseract OCR",
            "The Python package is installed, but the Tesseract OCR program is missing.\n\n"
            "Install it automatically now? Internet access may be required."
        ):
            return

        def worker():
            self.root.after(0, lambda: self.status("Installing Tesseract OCR..."))
            ok, details = install_tesseract_engine()
            self.root.after(0, lambda: self.status("OCR is ready." if ok else "OCR installation failed."))
            self.root.after(
                0,
                lambda: messagebox.showinfo("OCR Setup", details)
                if ok else messagebox.showerror("OCR Setup", details)
            )

        threading.Thread(target=worker, daemon=True).start()

    def ocr_selected(self):

        document = self.get_selected_document()

        if not document:
            return

        extension = (
            document["file_type"] or ""
        ).lower()

        ocr_extensions = [".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".pdf"]

        if extension not in ocr_extensions:

            messagebox.showwarning(
                "OCR",
                "OCR is available for image files and scanned PDFs."
            )

            return

        if pytesseract is None:

            messagebox.showerror(
                "OCR",
                "pytesseract Python package is not installed.\n\n"
                "Restart the app to auto-install packages, or run:\n"
                f"{sys.executable} -m pip install pytesseract"
            )

            return

        if not configure_tesseract_path():
            if messagebox.askyesno(
                "OCR Program Missing",
                "pytesseract is installed, but the Tesseract OCR program is missing.\n\n"
                "Would you like the application to install it now?"
            ):
                self.install_ocr_support()
            return

        thread = threading.Thread(
            target=self.ocr_worker,
            args=(document,),
            daemon=True
        )

        thread.start()

    def ocr_worker(
        self,
        document
    ):

        try:

            self.root.after(
                0,
                lambda: self.status(
                    "OCR processing..."
                )
            )

            if (document["file_type"] or "").lower() == ".pdf":
                text = extract_text_from_file(document["path"])
            else:
                text = ocr_image(Image.open(document["path"]))

            self.database.update_document_text(
                document["id"],
                text.strip()
            )

            self.database.add_history(
                document["id"],
                "OCR",
                text.strip()
            )

            def update():

                self.document_text.delete(
                    "1.0",
                    "end"
                )

                self.document_text.insert(
                    "1.0",
                    text.strip()
                )

                self.status(
                    "OCR completed and saved."
                )

                messagebox.showinfo(
                    "OCR Complete",
                    "Text has been extracted and saved "
                    "to the Library database."
                )

            self.root.after(
                0,
                update
            )

        except Exception as e:

            err = str(e)
            if "tesseract" in err.lower():
                err = (
                    err
                    + "\n\nInstall the Tesseract OCR program separately:\n"
                    + "Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
                    + "macOS: brew install tesseract\n"
                    + "Linux: sudo apt install tesseract-ocr"
                )
            self.root.after(
                0,
                lambda msg=err: messagebox.showerror(
                    "OCR Error",
                    msg
                )
            )

    # ============================================================
    # DIRECT SOURCE ANALYSIS / IMAGE TO PDF
    # ============================================================

    def ai_analyze_source(self):
        document = self.get_selected_document()
        if not document or not document["path"]:
            return
        instruction = (
            "Analyze this study source carefully. Extract visible or embedded text, "
            "identify the topic, summarize the main ideas, list key definitions and "
            "facts, note formulas or tables, and explain anything difficult. "
            "Use headings and bullet points. If the source is unclear, say what is uncertain. "
            "Do not invent information."
        )
        def worker():
            try:
                self.root.after(0, lambda: self.status("AI is analyzing the source file..."))
                result = self.ai.analyze_file(document["path"], instruction)
                self.database.update_ai_result(document["id"], "explanation", result)
                self.database.add_history(document["id"], "SOURCE_ANALYSIS", result)
                self.root.after(0, lambda: self.show_ai_result(document["id"], "explanation", result))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("AI Source Analysis Error", str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def convert_selected_image_to_pdf(self):
        document = self.get_selected_document()
        if not document:
            return
        image_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
        if (document["file_type"] or "").lower() not in image_extensions:
            messagebox.showwarning("Image to PDF", "Select an image file first.")
            return
        output = filedialog.asksaveasfilename(
            title="Convert Image to PDF", defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=Path(document["name"]).stem + ".pdf"
        )
        if not output:
            return
        try:
            image = Image.open(document["path"]).convert("RGB")
            image.save(output, "PDF", resolution=150.0)
            new_id = self.database.add_document(Path(output).name, output, ".pdf", "")
            self.database.add_history(new_id, "IMAGE_TO_PDF", output)
            self.load_library()
            self.status("Image converted to PDF and added to the Library.")
            messagebox.showinfo("Image to PDF", "PDF created and added to the Library.")
        except Exception as e:
            messagebox.showerror("Image to PDF Error", str(e))

    # ============================================================
    # AI SUMMARY
    # ============================================================

    def ai_summary(self):

        document = self.get_selected_document()

        if not document:
            return

        text = document["text"]

        # Use current edited text
        current_text = self.document_text.get(
            "1.0",
            "end"
        ).strip()

        if current_text:
            text = current_text

            self.database.update_document_text(
                document["id"],
                text
            )

        if not text.strip():

            messagebox.showwarning(
                "AI Summary",
                "This document does not contain text.\n\n"
                "For an image, use OCR first."
            )

            return

        self.run_ai(
            document["id"],
            "summary",
            text
        )

    # ============================================================
    # AI EXPLANATION
    # ============================================================

    def ai_explanation(self):

        document = self.get_selected_document()

        if not document:
            return

        text = document["text"]

        current_text = self.document_text.get(
            "1.0",
            "end"
        ).strip()

        if current_text:
            text = current_text

            self.database.update_document_text(
                document["id"],
                text
            )

        if not text.strip():

            messagebox.showwarning(
                "AI Explanation",
                "This document does not contain text.\n\n"
                "For an image, use OCR first."
            )

            return

        self.run_ai(
            document["id"],
            "explanation",
            text
        )

    # ============================================================
    # RUN AI
    # ============================================================

    def run_ai(
        self,
        document_id,
        operation,
        text
    ):

        if operation == "summary":

            instruction = """
Create a high-quality study summary of the supplied material.

Requirements:
- Identify the main concepts.
- Use clear headings.
- Use bullet points where useful.
- Keep important facts, definitions and examples.
- Remove unnecessary repetition.
- Make it useful for exam revision.
- Do not invent information that is not present.
- Write in clear professional English.
"""

        else:

            instruction = """
Explain the supplied study material as an excellent teacher.

Requirements:
- Explain the concepts in simple language.
- Preserve technical accuracy.
- Break difficult concepts into smaller parts.
- Give examples where useful.
- Explain important terminology.
- Highlight important points for exams.
- Do not invent facts that are not supported by the material.
"""

        def worker():

            try:

                self.root.after(
                    0,
                    lambda: self.status(
                        f"AI {operation} in progress..."
                    )
                )

                result = self.ai.generate(
                    instruction,
                    text
                )

                self.database.update_ai_result(
                    document_id,
                    operation,
                    result
                )

                self.database.add_history(
                    document_id,
                    operation.upper(),
                    result
                )

                self.root.after(
                    0,
                    lambda: self.show_ai_result(
                        document_id,
                        operation,
                        result
                    )
                )

            except Exception as e:

                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "AI Error",
                        str(e)
                    )
                )

                self.root.after(
                    0,
                    lambda: self.status(
                        "AI operation failed."
                    )
                )

        threading.Thread(
            target=worker,
            daemon=True
        ).start()

    # ============================================================
    # SHOW AI RESULT
    # ============================================================

    def show_ai_result(
        self,
        document_id,
        operation,
        result
    ):

        title = (
            "AI Summary"
            if operation == "summary"
            else "AI Explanation"
        )

        self.status(
            f"{title} completed."
        )

        win = tk.Toplevel(
            self.root
        )

        win.title(
            title
        )

        win.geometry(
            "950x700"
        )

        text = tk.Text(
            win,
            wrap="word",
            font=(
                "Segoe UI",
                12
            )
        )

        text.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=15
        )

        text.insert(
            "1.0",
            result
        )

        bottom = tk.Frame(
            win
        )

        bottom.pack(
            fill="x",
            padx=15,
            pady=10
        )

        ttk.Button(
            bottom,
            text="🔊 Read Aloud",
            command=lambda: windows_speak(
                text.get(
                    "1.0",
                    "end"
                )
            )
        ).pack(
            side="left"
        )

        ttk.Button(
            bottom,
            text="Save Changes",
            command=lambda: self.save_ai_window(
                document_id,
                operation,
                text,
                win
            )
        ).pack(
            side="left",
            padx=10
        )

        ttk.Button(
            bottom,
            text="Close",
            command=win.destroy
        ).pack(
            side="right"
        )

    # ============================================================
    # SAVE AI RESULT
    # ============================================================

    def save_ai_window(
        self,
        document_id,
        operation,
        text_widget,
        window
    ):

        value = text_widget.get(
            "1.0",
            "end"
        ).strip()

        self.database.update_ai_result(
            document_id,
            operation,
            value
        )

        messagebox.showinfo(
            "Saved",
            f"{operation.title()} saved to database."
        )

        window.destroy()

    # ============================================================
    # READ SELECTED
    # ============================================================

    def read_selected_aloud(self):

        document = self.get_selected_document()

        if not document:
            return

        current_text = self.document_text.get(
            "1.0",
            "end"
        ).strip()

        text = current_text or document["text"]

        if not text:

            messagebox.showwarning(
                "Read Aloud",
                "There is no text available."
            )

            return

        windows_speak(
            text
        )

        self.status(
            "Reading selected study material..."
        )

    # ============================================================
    # EXPORT SELECTED DOCUMENT
    # ============================================================

    def export_selected_word(self):

        document = self.get_selected_document()

        if not document:
            return

        if Document is None:

            messagebox.showerror(
                "Word Export",
                "python-docx is not installed."
            )

            return

        current_text = self.document_text.get(
            "1.0",
            "end"
        ).strip()

        text = (
            current_text
            or document["text"]
            or ""
        )

        summary = (
            document["summary"]
            or ""
        )

        explanation = (
            document["explanation"]
            or ""
        )

        if not text and not summary and not explanation:

            messagebox.showwarning(
                "Export",
                "There is no content to export."
            )

            return

        output = filedialog.asksaveasfilename(
            title="Export Study Material",
            defaultextension=".docx",
            filetypes=[
                (
                    "Microsoft Word",
                    "*.docx"
                )
            ],
            initialfile=(
                Path(
                    document["name"]
                ).stem
                + "_AI_Educator.docx"
            )
        )

        if not output:
            return

        try:

            doc = Document()

            # ----------------------------------------------------
            # TITLE
            # ----------------------------------------------------

            title = doc.add_paragraph()

            title.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            run = title.add_run(
                "AI EDUCATOR PRO"
            )

            run.bold = True
            run.font.size = Pt(22)

            subtitle = doc.add_paragraph()

            subtitle.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            subtitle.add_run(
                "Study Material & AI Learning Notes"
            ).bold = True

            doc.add_paragraph("")

            # ----------------------------------------------------
            # FILE INFORMATION
            # ----------------------------------------------------

            p = doc.add_paragraph()

            p.add_run(
                "Source File: "
            ).bold = True

            p.add_run(
                document["name"]
            )

            p = doc.add_paragraph()

            p.add_run(
                "Generated: "
            ).bold = True

            p.add_run(
                datetime.now().strftime(
                    "%d-%m-%Y %H:%M"
                )
            )

            doc.add_paragraph("")

            # ----------------------------------------------------
            # ORIGINAL / OCR TEXT
            # ----------------------------------------------------

            heading = doc.add_paragraph()

            heading.add_run(
                "1. STUDY MATERIAL"
            ).bold = True

            doc.add_paragraph(
                text
            )

            # ----------------------------------------------------
            # SUMMARY
            # ----------------------------------------------------

            if summary:

                heading = doc.add_paragraph()

                heading.add_run(
                    "2. AI SUMMARY"
                ).bold = True

                doc.add_paragraph(
                    summary
                )

            # ----------------------------------------------------
            # EXPLANATION
            # ----------------------------------------------------

            if explanation:

                heading = doc.add_paragraph()

                heading.add_run(
                    "3. AI EXPLANATION"
                ).bold = True

                doc.add_paragraph(
                    explanation
                )

            # ----------------------------------------------------
            # PROFILE
            # ----------------------------------------------------

            profile = self.database.get_profile()

            if profile:

                doc.add_paragraph("")

                heading = doc.add_paragraph()

                heading.add_run(
                    "LEARNER PROFILE"
                ).bold = True

                doc.add_paragraph(
                    f"Name: {profile['name']}\n"
                    f"Type: {profile['user_type']}\n"
                    f"Age: {profile['age']}\n"
                    f"Class / Qualification: "
                    f"{profile['qualification']}\n"
                    f"Hobbies: {profile['hobbies']}"
                )

            doc.save(
                output
            )

            self.database.add_export(document["id"], "DOCX", output)
            self.database.add_history(document["id"], "EXPORT_DOCX", output)

            self.status(
                "Word document exported successfully."
            )

            messagebox.showinfo(
                "Export Complete",
                "Word document created successfully:\n\n"
                + output
            )

        except Exception as e:

            messagebox.showerror(
                "Export Error",
                str(e)
            )

    # ============================================================
    # SHARED EXPORT DATA
    # ============================================================

    def get_selected_export_data(self):

        document = self.get_selected_document()
        if not document:
            return None

        current_text = self.document_text.get("1.0", "end").strip()
        text = current_text or document["text"] or ""
        summary = document["summary"] or ""
        explanation = document["explanation"] or ""
        profile = self.database.get_profile()
        history = self.database.get_history(document["id"])

        if not text and document["path"]:
            try:
                text = extract_text_from_file(document["path"])
                if text:
                    self.database.update_document_text(document["id"], text)
            except Exception:
                pass

        return {
            "document": document,
            "text": clean_text_for_export(text),
            "summary": clean_text_for_export(summary),
            "explanation": clean_text_for_export(explanation),
            "profile": profile,
            "history": history,
        }

    def export_selected_excel(self):

        data = self.get_selected_export_data()
        if not data:
            return

        if xlsxwriter is None:
            messagebox.showerror("Excel Export", "xlsxwriter is not installed.")
            return

        document = data["document"]
        output = filedialog.asksaveasfilename(
            title="Export Study Package to Excel",
            defaultextension=".xlsx",
            filetypes=[("Microsoft Excel", "*.xlsx")],
            initialfile=Path(document["name"]).stem + "_AI_Educator.xlsx"
        )
        if not output:
            return

        try:
            workbook = xlsxwriter.Workbook(output)
            title_fmt = workbook.add_format({"bold": True, "font_size": 18, "align": "center", "valign": "vcenter"})
            section_fmt = workbook.add_format({"bold": True, "font_size": 13, "bg_color": "#173F5F", "font_color": "white"})
            header_fmt = workbook.add_format({"bold": True, "bg_color": "#E8EEF5", "border": 1})
            wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
            meta_fmt = workbook.add_format({"border": 1, "valign": "top"})

            overview = workbook.add_worksheet("Study Package")
            overview.merge_range("A1:F1", "AI EDUCATOR PRO - Study Package", title_fmt)
            overview.set_column("A:A", 24)
            overview.set_column("B:F", 24)
            overview.write_row("A3", ["Field", "Value"], header_fmt)
            meta_rows = [
                ["Document ID", document["id"]],
                ["File Name", document["name"]],
                ["File Type", document["file_type"]],
                ["Source Path", document["path"] or ""],
                ["Created", document["created_at"]],
                ["Updated", document["updated_at"]],
            ]
            overview.write("A4", "Document Metadata", section_fmt)
            row = 4
            for key, value in meta_rows:
                row += 1
                overview.write(row, 0, key, meta_fmt)
                overview.write(row, 1, value, wrap_fmt)
            row += 2
            overview.write(row, 0, "Study Material", section_fmt)
            overview.write(row + 1, 0, data["text"], wrap_fmt)
            overview.merge_range(row + 1, 0, row + max(3, min(40, data["text"].count("\n") + 3)), 5, data["text"], wrap_fmt)

            summary_ws = workbook.add_worksheet("AI Analysis")
            summary_ws.set_column("A:A", 20)
            summary_ws.set_column("B:B", 100)
            summary_ws.write_row("A1", ["Analysis", "AI Output"], header_fmt)
            summary_ws.write("A2", "Summary", section_fmt)
            summary_ws.write("B2", data["summary"], wrap_fmt)
            summary_ws.write("A3", "Explanation", section_fmt)
            summary_ws.write("B3", data["explanation"], wrap_fmt)

            db_ws = workbook.add_worksheet("Database")
            db_ws.set_column("A:A", 12)
            db_ws.set_column("B:C", 18)
            db_ws.set_column("D:D", 80)
            db_ws.set_column("E:E", 22)
            db_ws.write_row("A1", ["History ID", "Document ID", "Action", "Result", "Created At"], header_fmt)
            for i, item in enumerate(data["history"], start=1):
                db_ws.write(i, 0, item["id"], meta_fmt)
                db_ws.write(i, 1, item["document_id"], meta_fmt)
                db_ws.write(i, 2, item["action"], meta_fmt)
                db_ws.write(i, 3, item["result"], wrap_fmt)
                db_ws.write(i, 4, item["created_at"], meta_fmt)
            if data["history"]:
                db_ws.add_table(0, 0, len(data["history"]), 4, {"name": "HistoryTable", "style": "Table Style Medium 2", "columns": [
                    {"header": "History ID"}, {"header": "Document ID"}, {"header": "Action"}, {"header": "Result"}, {"header": "Created At"}
                ]})

            exports = self.database.get_exports(document["id"])
            exp_start = max(len(data["history"]) + 4, 5)
            db_ws.write(exp_start, 0, "Exports", section_fmt)
            db_ws.write_row(exp_start + 1, 0, ["Export ID", "Format", "Output Path", "Created At"], header_fmt)
            for i, item in enumerate(exports, start=exp_start + 2):
                db_ws.write(i, 0, item["id"], meta_fmt)
                db_ws.write(i, 1, item["format"], meta_fmt)
                db_ws.write(i, 2, item["output_path"], wrap_fmt)
                db_ws.write(i, 3, item["created_at"], meta_fmt)

            workbook.close()
            self.database.add_export(document["id"], "XLSX", output)
            self.database.add_history(document["id"], "EXPORT_XLSX", output)
            self.status("Excel workbook exported successfully.")
            messagebox.showinfo("Export Complete", "Excel workbook created successfully:\n\n" + output)
        except Exception as e:
            messagebox.showerror("Excel Export Error", str(e))

    def export_selected_pdf(self):

        data = self.get_selected_export_data()
        if not data:
            return

        if SimpleDocTemplate is None:
            messagebox.showerror("PDF Export", "reportlab is not installed.")
            return

        document = data["document"]
        output = filedialog.asksaveasfilename(
            title="Export Study Package to PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=Path(document["name"]).stem + "_AI_Educator.pdf"
        )
        if not output:
            return

        try:
            styles = getSampleStyleSheet()
            styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER, spaceAfter=10))
            styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], spaceBefore=10, spaceAfter=6))
            body = styles["BodyText"]
            body.leading = 14

            doc_pdf = SimpleDocTemplate(output, pagesize=A4, rightMargin=16 * mm, leftMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm)
            story = [
                Paragraph("AI EDUCATOR PRO", styles["CenterTitle"]),
                Paragraph("Study Material & AI Learning Notes", styles["CenterTitle"]),
                Spacer(1, 8),
                Paragraph("Document Information", styles["Section"]),
            ]

            meta = [
                ["Document ID", str(document["id"])],
                ["File Name", document["name"]],
                ["File Type", document["file_type"]],
                ["Generated", datetime.now().strftime("%d-%m-%Y %H:%M")],
            ]
            table = Table(meta, colWidths=[38 * mm, 130 * mm])
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E8EEF5")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ]))
            story.extend([table, Spacer(1, 10)])

            def add_text_section(title, text):
                story.append(Paragraph(title, styles["Section"]))
                if not text:
                    story.append(Paragraph("No content available.", body))
                    return
                for block in text.split("\n\n"):
                    safe = (block.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>"))
                    story.append(Paragraph(safe, body))
                    story.append(Spacer(1, 5))

            add_text_section("1. Study Material", data["text"])
            add_text_section("2. AI Summary", data["summary"])
            add_text_section("3. AI Explanation", data["explanation"])

            history_text = []
            for item in data["history"][:50]:
                history_text.append(f"{item['created_at']} | {item['action']} | {item['result']}")
            add_text_section("4. Database Activity History", "\n\n".join(history_text))

            doc_pdf.build(story)
            self.database.add_export(document["id"], "PDF", output)
            self.database.add_history(document["id"], "EXPORT_PDF", output)
            self.status("PDF exported successfully.")
            messagebox.showinfo("Export Complete", "PDF created successfully:\n\n" + output)
        except Exception as e:
            messagebox.showerror("PDF Export Error", str(e))

    def export_selected_all(self):

        data = self.get_selected_export_data()
        if not data:
            return

        folder = filedialog.askdirectory(title="Choose Export Folder")
        if not folder:
            return

        stem = Path(data["document"]["name"]).stem
        outputs = []

        # Word
        try:
            if Document is not None:
                word_path = os.path.join(folder, stem + "_AI_Educator.docx")
                # Reuse the currently selected document export workflow by building the Word document directly.
                doc = Document()
                title = doc.add_paragraph()
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = title.add_run("AI EDUCATOR PRO")
                run.bold = True
                run.font.size = Pt(22)
                doc.add_paragraph(data["document"]["name"]).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_heading("1. Study Material", level=1)
                doc.add_paragraph(data["text"] or "No content available.")
                doc.add_heading("2. AI Summary", level=1)
                doc.add_paragraph(data["summary"] or "No summary available.")
                doc.add_heading("3. AI Explanation", level=1)
                doc.add_paragraph(data["explanation"] or "No explanation available.")
                doc.add_heading("4. Database Activity", level=1)
                for item in data["history"][:50]:
                    doc.add_paragraph(f"{item['created_at']} | {item['action']} | {item['result']}")
                doc.save(word_path)
                self.database.add_export(data["document"]["id"], "DOCX", word_path)
                outputs.append(word_path)
        except Exception:
            pass

        # Excel - write a lightweight, self-contained package without opening another dialog.
        if xlsxwriter is not None:
            try:
                xlsx_path = os.path.join(folder, stem + "_AI_Educator.xlsx")
                workbook = xlsxwriter.Workbook(xlsx_path)
                title_fmt = workbook.add_format({"bold": True, "font_size": 18, "align": "center"})
                header_fmt = workbook.add_format({"bold": True, "bg_color": "#E8EEF5", "border": 1})
                wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
                ws = workbook.add_worksheet("Study Package")
                ws.merge_range("A1:F1", "AI EDUCATOR PRO - Study Package", title_fmt)
                ws.set_column("A:A", 22)
                ws.set_column("B:F", 24)
                ws.write_row("A3", ["Section", "Content"], header_fmt)
                rows = [["File Name", data["document"]["name"]], ["File Type", data["document"]["file_type"]], ["Study Material", data["text"]], ["AI Summary", data["summary"]], ["AI Explanation", data["explanation"]]]
                for r, pair in enumerate(rows, start=3):
                    ws.write(r, 0, pair[0], header_fmt)
                    ws.merge_range(r, 1, r, 5, pair[1], wrap_fmt)
                db = workbook.add_worksheet("Database")
                db.write_row("A1", ["ID", "Document ID", "Action", "Result", "Created At"], header_fmt)
                for r, item in enumerate(data["history"], start=1):
                    db.write_row(r, 0, [item["id"], item["document_id"], item["action"], item["result"], item["created_at"]])
                db.set_column("D:D", 80)
                workbook.close()
                self.database.add_export(data["document"]["id"], "XLSX", xlsx_path)
                outputs.append(xlsx_path)
            except Exception:
                pass

        # PDF
        if SimpleDocTemplate is not None:
            try:
                pdf_path = os.path.join(folder, stem + "_AI_Educator.pdf")
                styles = getSampleStyleSheet()
                story = [Paragraph("AI EDUCATOR PRO", styles["Title"]), Paragraph(data["document"]["name"], styles["Heading2"])]
                def add_pdf(title, text):
                    story.append(Paragraph(title, styles["Heading2"]))
                    story.append(Paragraph((text or "No content available.").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>"), styles["BodyText"]))
                add_pdf("1. Study Material", data["text"])
                add_pdf("2. AI Summary", data["summary"])
                add_pdf("3. AI Explanation", data["explanation"])
                add_pdf("4. Database Activity", "\n\n".join(f"{x['created_at']} | {x['action']} | {x['result']}" for x in data["history"][:50]))
                SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=16 * mm, leftMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm).build(story)
                self.database.add_export(data["document"]["id"], "PDF", pdf_path)
                outputs.append(pdf_path)
            except Exception:
                pass

        if outputs:
            self.database.add_history(data["document"]["id"], "EXPORT_ALL", "\n".join(outputs))
            self.status("Word, Excel and PDF exports completed.")
            messagebox.showinfo("Export Complete", "Files created:\n\n" + "\n".join(outputs))
        else:
            messagebox.showerror("Export", "No export format could be created. Check installed packages.")

    # ============================================================
    # DELETE
    # ============================================================

    def delete_selected(self):

        document = self.get_selected_document()

        if not document:
            return

        answer = messagebox.askyesno(
            "Delete Document",
            (
                f"Delete '{document['name']}' "
                "from the Library database?"
            )
        )

        if not answer:
            return

        self.database.delete_document(
            document["id"]
        )

        self.selected_document_id = None

        self.document_text.delete(
            "1.0",
            "end"
        )

        self.selected_label.config(
            text="No document selected",
            fg="#777"
        )

        self.load_library()

    # ============================================================
    # SETTINGS
    # ============================================================

    def open_settings(self):

        win = tk.Toplevel(
            self.root
        )

        win.title(
            "AI Educator Settings"
        )

        win.geometry(
            "600x500"
        )

        win.transient(
            self.root
        )

        frame = tk.Frame(
            win,
            bg="white"
        )

        frame.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=25
        )

        tk.Label(
            frame,
            text="⚙ AI Settings",
            bg="white",
            font=(
                "Segoe UI",
                20,
                "bold"
            )
        ).pack(
            pady=(0, 20)
        )

        # --------------------------------------------------------
        # API KEY
        # --------------------------------------------------------

        tk.Label(
            frame,
            text="OpenAI API Key",
            bg="white",
            font=(
                "Segoe UI",
                10,
                "bold"
            )
        ).pack(
            anchor="w"
        )

        api_var = tk.StringVar(
            value=self.database.get_setting(
                "openai_api_key"
            )
        )

        api_entry = tk.Entry(
            frame,
            textvariable=api_var,
            show="*",
            font=(
                "Segoe UI",
                10
            )
        )

        api_entry.pack(
            fill="x",
            pady=8
        )

        tk.Label(
            frame,
            text=(
                "The API key is stored in the local SQLite "
                "database on this computer."
            ),
            bg="white",
            fg="#666"
        ).pack(
            anchor="w"
        )

        # --------------------------------------------------------
        # MODEL
        # --------------------------------------------------------

        tk.Label(
            frame,
            text="AI Model",
            bg="white",
            font=(
                "Segoe UI",
                10,
                "bold"
            )
        ).pack(
            anchor="w",
            pady=(20, 0)
        )

        model_var = tk.StringVar(
            value=self.database.get_setting(
                "openai_model",
                "gpt-5"
            )
        )

        model_combo = ttk.Combobox(
            frame,
            textvariable=model_var,
            values=[
                "gpt-5",
                "gpt-4.1",
                "gpt-4o"
            ],
            state="normal"
        )

        model_combo.pack(
            fill="x",
            pady=8
        )

        # --------------------------------------------------------
        # TEST API
        # --------------------------------------------------------

        def test_api():

            key = api_var.get().strip()

            if not key:

                messagebox.showwarning(
                    "API",
                    "Please enter an API key."
                )

                return

            try:

                if OpenAI is None:

                    raise Exception(
                        "OpenAI package is unavailable."
                    )

                client = OpenAI(
                    api_key=key
                )

                response = client.responses.create(
                    model=model_var.get().strip(),
                    input=(
                        "Reply with exactly: "
                        "AI Educator API connection successful."
                    )
                )

                messagebox.showinfo(
                    "API Test",
                    response.output_text
                )

            except Exception as e:

                messagebox.showerror(
                    "API Test Failed",
                    str(e)
                )

        ttk.Button(
            frame,
            text="🔌 Test API Connection",
            command=test_api
        ).pack(
            pady=15
        )

        # --------------------------------------------------------
        # SAVE
        # --------------------------------------------------------

        def save():

            self.database.set_setting(
                "openai_api_key",
                api_var.get().strip()
            )

            self.database.set_setting(
                "openai_model",
                model_var.get().strip()
            )

            messagebox.showinfo(
                "Settings Saved",
                "AI settings saved successfully."
            )

            win.destroy()

        ttk.Button(
            frame,
            text="Save Settings",
            command=save
        ).pack(
            pady=10
        )

        tk.Label(
            frame,
            text=(
                "API usage may incur charges according to "
                "your OpenAI account and selected model."
            ),
            bg="white",
            fg="#777",
            wraplength=500,
            justify="left"
        ).pack(
            pady=15
        )

    # ============================================================
    # PROFILE
    # ============================================================

    def open_profile(self):
        win = tk.Toplevel(self.root)
        win.title("Learner Profiles")
        win.geometry("620x700")
        win.transient(self.root)

        frame = tk.Frame(win, bg="white")
        frame.pack(fill="both", expand=True, padx=25, pady=20)

        tk.Label(
            frame, text="👥 Learner Profiles", bg="white",
            font=("Segoe UI", 20, "bold")
        ).pack(pady=(0, 12))

        selector_frame = tk.Frame(frame, bg="white")
        selector_frame.pack(fill="x", pady=(0, 12))
        tk.Label(selector_frame, text="Select profile:", bg="white").pack(side="left")

        profile_var = tk.StringVar()
        selector = ttk.Combobox(selector_frame, textvariable=profile_var, state="readonly", width=38)
        selector.pack(side="left", padx=8, fill="x", expand=True)

        form = tk.Frame(frame, bg="white")
        form.pack(fill="both", expand=True)
        fields = {}

        def entry_field(label):
            tk.Label(form, text=label, bg="white").pack(anchor="w")
            variable = tk.StringVar()
            ttk.Entry(form, textvariable=variable).pack(fill="x", pady=(3, 9))
            fields[label] = variable
            return variable

        name = entry_field("Name")
        age = entry_field("Age")
        qualification = entry_field("Class / Qualification")
        hobbies = entry_field("Hobbies")

        tk.Label(form, text="Profile Type", bg="white").pack(anchor="w")
        user_type = tk.StringVar()
        ttk.Combobox(
            form, textvariable=user_type,
            values=["Student", "Professional", "Teacher", "Researcher", "Parent", "Other"],
            state="readonly"
        ).pack(fill="x", pady=(3, 9))

        tk.Label(form, text="Mood / Learning Style", bg="white").pack(anchor="w")
        mood = tk.StringVar()
        ttk.Combobox(
            form, textvariable=mood,
            values=["Happy", "Calm", "Professional", "Focused", "Beginner", "Detailed", "Exam Revision"],
            state="readonly"
        ).pack(fill="x", pady=(3, 9))

        profile_map = {}

        def refresh(select_id=None):
            nonlocal profile_map
            profiles = self.database.get_profiles()
            profile_map = {f"{row['name'] or 'Unnamed'} (ID {row['id']})": row["id"] for row in profiles}
            selector["values"] = list(profile_map.keys())
            active = self.database.get_profile(select_id)
            if active:
                display = next((key for key, value in profile_map.items() if value == active["id"]), "")
                profile_var.set(display)
                load_profile(active["id"])

        def load_profile(profile_id=None):
            if profile_id is None:
                profile_id = profile_map.get(profile_var.get())
            row = self.database.get_profile(profile_id)
            if not row:
                return
            self.database.set_active_profile(row["id"])
            name.set(row["name"] or "")
            age.set(row["age"] or "")
            qualification.set(row["qualification"] or "")
            hobbies.set(row["hobbies"] or "")
            user_type.set(row["user_type"] or "Student")
            mood.set(row["mood"] or "Happy")
            win.title(f"Learner Profiles - Active: {row['name'] or 'Unnamed'}")

        def selected(event=None):
            profile_id = profile_map.get(profile_var.get())
            if profile_id:
                load_profile(profile_id)

        selector.bind("<<ComboboxSelected>>", selected)

        def save():
            profile_id = profile_map.get(profile_var.get())
            if not profile_id:
                return
            if not name.get().strip():
                messagebox.showwarning("Profile", "Please enter a profile name.")
                return
            self.database.save_profile(
                name.get().strip(), age.get().strip(), user_type.get(),
                qualification.get().strip(), hobbies.get().strip(), mood.get(), profile_id
            )
            refresh(profile_id)
            messagebox.showinfo("Profile", "Profile saved and selected as active.")

        def add_new():
            new_name = simpledialog.askstring("New Profile", "Enter the learner's name:", parent=win)
            if new_name and new_name.strip():
                profile_id = self.database.add_profile(new_name.strip())
                refresh(profile_id)

        def delete_current():
            profile_id = profile_map.get(profile_var.get())
            if not profile_id:
                return
            if not messagebox.askyesno("Delete Profile", "Delete this learner profile?", parent=win):
                return
            try:
                self.database.delete_profile(profile_id)
                refresh()
            except Exception as exc:
                messagebox.showerror("Delete Profile", str(exc), parent=win)

        buttons = tk.Frame(frame, bg="white")
        buttons.pack(fill="x", pady=12)
        ttk.Button(buttons, text="＋ New Profile", command=add_new).pack(side="left")
        ttk.Button(buttons, text="Save / Set Active", command=save).pack(side="left", padx=8)
        ttk.Button(buttons, text="Delete", command=delete_current).pack(side="left")
        ttk.Button(buttons, text="Close", command=win.destroy).pack(side="right")

        refresh()

    # ============================================================
    # CLOSE
    # ============================================================

    def close_application(self):

        self.root.destroy()


# ================================================================
# START
# ================================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = EducatorApp(
        root
    )

    root.mainloop()
