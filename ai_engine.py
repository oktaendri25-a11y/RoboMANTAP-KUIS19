import os
import io
import json
import re
import time
import hashlib
import subprocess
import tempfile
import pandas as pd
try:
    import streamlit as st  # tersedia di aplikasi Streamlit utama
except ImportError:
    class _StreamlitCompat:
        secrets = {}
        @staticmethod
        def error(message):
            print(f"[STREAMLIT COMPAT] {message}")
        @staticmethod
        def connection(*args, **kwargs):
            raise RuntimeError("Set DATABASE_URL pada environment variables untuk deployment FastAPI.")
    st = _StreamlitCompat()
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Setup Matplotlib Headless untuk Server Cloud (Render/Streamlit)
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
# Import Python-Docx & XML Parser untuk Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from datetime import datetime, timezone, timedelta


try:
    from sqlalchemy import create_engine, text
    from sqlalchemy.orm import sessionmaker
except ImportError:
    pass

load_dotenv()

# ==============================================================================
# PEMBACAAN API KEY AMAN (DUAL COMPATIBILITY: STREAMLIT & FASTAPI/RENDER)
# ==============================================================================
api_keys = []
raw_keys = os.getenv("GEMINI_API_KEYS") or os.getenv("GEMINI_API_KEY")

# Fallback ke st.secrets jika dipanggil di dalam lingkungan Streamlit
if not raw_keys:
    try:
        if hasattr(st, "secrets"):
            if "GEMINI_API_KEYS" in st.secrets:
                raw_keys = st.secrets["GEMINI_API_KEYS"]
            elif "GEMINI_API_KEY" in st.secrets:
                raw_keys = st.secrets["GEMINI_API_KEY"]
    except Exception:
        raw_keys = None

if raw_keys:
    if isinstance(raw_keys, list):
        api_keys = raw_keys
    else:
        raw_keys_str = str(raw_keys).strip()
        if raw_keys_str.startswith("["):
            try:
                api_keys = json.loads(raw_keys_str)
            except Exception:
                api_keys = [k.strip(' "\'') for k in raw_keys_str.strip("[]").split(",") if k.strip()]
        elif "," in raw_keys_str:
            api_keys = [k.strip(' "\'') for k in raw_keys_str.split(",") if k.strip()]
        else:
            api_keys = [raw_keys_str.strip(' "\'')]

# Helper untuk menampilkan error yang aman di kedua lingkungan (Streamlit & Render)
def show_error(msg: str):
    print(f"[AI ENGINE ERROR] {msg}")
    try:
        st.error(msg)
    except Exception:
        pass

# Model untuk pembuatan soal
QUIZ_MODELS = ("gemini-3.5-flash-lite", "gemini-3.1-flash-lite")

# Model khusus interaksi LIVE: prioritaskan latency rendah.
STREAM_MODELS = ("gemini-3.5-flash-lite", "gemini-3.1-flash-lite")

STREAM_HINT_MAX_TOKENS = 9000
STREAM_SOLUTION_MAX_TOKENS = 9000
STREAM_TIMEOUT_MS = 90_000

_clients_cache = None

def get_gemini_clients():
    """
    Reuse koneksi Gemini antar request.
    Menggunakan caching global yang kompatibel untuk Streamlit maupun Uvicorn FastAPI.
    """
    global _clients_cache
    if _clients_cache is not None:
        return _clients_cache

    clients = []
    for key in api_keys:
        try:
            clients.append(
                genai.Client(
                    api_key=key,
                    http_options=types.HttpOptions(
                        timeout=STREAM_TIMEOUT_MS,
                        retry_options=types.HttpRetryOptions(attempts=1),
                    ),
                )
            )
        except Exception:
            continue

    _clients_cache = clients
    return clients


def _stream_config(model_name: str, max_output_tokens: int):
    if model_name.startswith("gemini-3."):
        return types.GenerateContentConfig(
            max_output_tokens=max_output_tokens,
            thinking_config=types.ThinkingConfig(
                thinking_level="high"
            ),
        )

    return types.GenerateContentConfig(
        max_output_tokens=max_output_tokens,
        thinking_config=types.ThinkingConfig(
            thinking_budget=0,
            include_thoughts=False,
        ),
    )


def _buffer_stream_text(source, min_chars: int = 2, flush_seconds: float = 0.01):
    buffer = []
    size = 0
    last_flush = time.monotonic()

    for chunk in source:
        if not chunk:
            continue

        buffer.append(chunk)
        size += len(chunk)

        now = time.monotonic()
        if size >= min_chars or (now - last_flush) >= flush_seconds:
            yield "".join(buffer)
            buffer.clear()
            size = 0
            last_flush = now

    if buffer:
        yield "".join(buffer)


def _stream_from_clients(prompt: str, max_output_tokens: int):
    clients = get_gemini_clients()

    if not clients:
        yield "⚠️ Tidak ada koneksi yang aktif nih. Coba Kamu klik lagi.."
        return

    for client in clients:
        for model_name in STREAM_MODELS:
            try:
                response = client.models.generate_content_stream(
                    model=model_name,
                    contents=prompt,
                    config=_stream_config(model_name, max_output_tokens),
                )

                emitted = False

                def raw_stream():
                    for chunk in response:
                        chunk_text = getattr(chunk, "text", None)
                        if chunk_text:
                            yield chunk_text

                for piece in _buffer_stream_text(raw_stream()):
                    emitted = True
                    yield piece

                if emitted:
                    return

            except Exception:
                continue

    yield (
        "⚠️ Maaf ya, koneksi sedang bermasalah atau kuota sedang penuh nih. "
        "Silakan coba klik lagi ya!"
    )

def format_latex_options(options):
    formatted = []
    for opt in options:
        opt = str(opt).replace(r"\frac", r"\tfrac")
        if "\\" in opt and "$" not in opt:
            parts = opt.split(". ", 1)
            opt = f"{parts[0]}. ${parts[1]}$" if len(parts) == 2 else f"${opt}$"
        formatted.append(opt)
    return formatted


# ==============================================================================
# KEBIJAKAN OPSI JAWABAN & PROFIL BAHASA MADRASAH
# ==============================================================================
OPTION_LABELS_MTS = ("A", "B", "C", "D")
OPTION_LABELS_MA = ("A", "B", "C", "D", "E")


def option_labels_for_jenjang(jenjang: str) -> tuple[str, ...]:
    """MTs -> A-D, MA -> A-E. Default aman mengikuti MTs untuk jenjang tidak dikenal."""
    value = str(jenjang or "").strip().lower()
    if "ma" in value or "aliyah" in value:
        return OPTION_LABELS_MA
    return OPTION_LABELS_MTS


def option_count_for_jenjang(jenjang: str) -> int:
    return len(option_labels_for_jenjang(jenjang))


def normalize_custom_timer_config(config: dict | None) -> dict:
    """Normalisasi durasi kuis custom ke satu sumber kebenaran: timer_seconds.

    Tetap mempertahankan timer_h/timer_m/timer_s untuk kompatibilitas UI lama,
    tetapi engine CBT harus membaca total detik agar durasi 1-23 jam tidak
    terpotong menjadi timer_m saja.
    """
    cfg = dict(config or {})
    try:
        stored_seconds = cfg.get("timer_seconds")
        stored_total = int(stored_seconds or 0) if stored_seconds is not None else 0
        h = max(0, int(cfg.get("timer_h", 0) or 0))
        m = max(0, int(cfg.get("timer_m", 0) or 0))
        sec = max(0, int(cfg.get("timer_s", 0) or 0))
        component_total = h * 3600 + m * 60 + sec
        # timer_seconds menjadi sumber utama bila bernilai positif. Bila 0
        # tetapi komponen jam/menit/detik berisi nilai, pulihkan dari komponen
        # agar konfigurasi lama yang tidak konsisten tidak berubah menjadi
        # "tanpa batas waktu".
        total = stored_total if stored_total > 0 else component_total
    except (TypeError, ValueError):
        total = 0

    cfg["timer_seconds"] = total
    cfg["timer_h"] = total // 3600
    cfg["timer_m"] = (total % 3600) // 60
    cfg["timer_s"] = total % 60
    return cfg


def _split_option_label(value: str, fallback_index: int = 0):
    text = str(value or "").strip()
    match = re.match(r"^\s*([A-Ea-e])\s*[\.\)\:\-]\s*(.*)$", text, flags=re.DOTALL)
    if match:
        return match.group(1).upper(), match.group(2).strip()
    fallback = chr(65 + fallback_index)
    return fallback, text


def normalize_quiz_options(options, jenjang: str):
    """Normalisasi label opsi ke A. ... / B. ... tanpa mengubah isi pilihan."""
    expected = option_labels_for_jenjang(jenjang)
    normalized = []
    for idx, raw in enumerate(options or []):
        label, body = _split_option_label(raw, idx)
        normalized.append(f"{label}. {body}" if body else f"{label}.")
    return normalized, expected


def _is_arabic_subject(mapel: str, bahasa: str = "") -> bool:
    hay = f"{mapel or ''} {bahasa or ''}".lower()
    keys = (
        "bahasa arab", "nahwu", "sharaf", "qawaid", "muhadatsah",
        "insya", "balaghah", "mufradat", "qiraah", "qiroah", "imla"
    )
    return any(k in hay for k in keys)


def _is_religious_subject(mapel: str, konteks: str = "") -> bool:
    hay = f"{mapel or ''} {konteks or ''}".lower()
    keys = (
        "fiqih", "fikih", "aqidah", "akidah", "akhlak", "al-qur", "alqur",
        "hadis", "hadits", "ski", "sejarah kebudayaan islam", "keislaman",
        "ushul", "tarikh", "waris", "zakat", "ibadah"
    )
    return any(k in hay for k in keys)


def build_language_guidance(mapel: str, bahasa: str = "", konteks: str = "", source_pack: dict | None = None) -> str:
    """Instruksi bahasa yang tegas tetapi hanya mengaktifkan Arab penuh pada konteks yang relevan."""
    arabic_subject = _is_arabic_subject(mapel, bahasa)
    religious_subject = _is_religious_subject(mapel, konteks)
    if arabic_subject:
        return (
            "MODE BAHASA ARAB MADRASAH: Karena mata pelajaran/konfigurasi berhubungan langsung dengan Bahasa Arab, "
            "gunakan Bahasa Arab Fusha/Modern Standard Arabic yang natural, gramatikal, dan modern. Pertahankan seluruh "
            "teks soal, pilihan jawaban, dan solution_basis dalam aksara Arab; jangan gunakan transliterasi Latin. "
            "Gunakan harakat secara selektif pada kosakata/struktur yang berpotensi ambigu. Pertahankan simbol matematika "
            "dan satuan secara jelas. Untuk kutipan keagamaan, jangan mengarang teks Arab dan jangan memparafrasekan sebagai kutipan langsung."
        )
    if religious_subject:
        return (
            "MODE MATERI KEAGAMAAN: Bahasa utama mengikuti pilihan guru. Gunakan istilah, ungkapan, atau kutipan Arab asli "
            "hanya saat memang relevan dengan konsep/materi. Jika menyertakan teks Al-Qur'an, Hadis, doa, atau istilah Arab, "
            "tulis aksara Arab dengan benar dan jangan membuat kutipan Arab yang tidak didukung sumber. Sertakan penjelasan "
            "yang mudah dipahami dalam bahasa utama. Jangan memaksa seluruh soal menjadi Bahasa Arab."
        )
    return (
        "MODE UMUM MADRASAH: Gunakan bahasa utama yang dipilih guru. Bahasa Arab hanya digunakan untuk istilah yang memang "
        "relevan, bukan sebagai hiasan. Jangan menggunakan transliterasi Latin jika suatu istilah Arab asli memang diperlukan."
    )

def clean_json_text(text: str) -> str:
    if not text:
        return ""
    
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\n?", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()

    def replace_slash(match):
        g = match.group(0)
        if g in (r'\\', r'\"'):
            return g 
        return r'\\' 

    return re.sub(r'\\\\|\\"|\\', replace_slash, text)

def call_gemini_with_rotation(prompt: str, is_json: bool = False):
    clients = get_gemini_clients()
    if not clients:
        return None

    for client in clients:
        for model_name in QUIZ_MODELS:
            try:
                config_kwargs = {}

                if is_json:
                    config_kwargs["response_mime_type"] = "application/json"

                if model_name.startswith("gemini-3."):
                    config_kwargs["thinking_config"] = types.ThinkingConfig(
                        thinking_level="high"
                    )
                else:
                    config_kwargs["thinking_config"] = types.ThinkingConfig(
                        thinking_budget=0,
                        include_thoughts=False,
                    )

                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(**config_kwargs),
                )

                if response and response.text:
                    return response.text

            except Exception:
                continue

    return None

def stream_ai_text(prompt: str, max_output_tokens: int = STREAM_HINT_MAX_TOKENS):
    yield from _stream_from_clients(
        prompt,
        max_output_tokens=max_output_tokens,
    )

# ============================================================================
# MATERIAL HUB • FILE DROP • SOURCE GROUNDING • MEDIA AJAR STUDIO
# ============================================================================

SUPPORTED_MATERIAL_EXTENSIONS = {
    ".pdf": "application/pdf",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".ppt": "application/vnd.ms-powerpoint",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}


def normalize_material_trigger(value: str) -> bool:
    normalized = re.sub(r"\s+", " ", str(value or "").strip().lower())
    return normalized == "materi dilampirkan"


def _trim_text(value: str, max_chars: int = 30000) -> str:
    value = str(value or "").strip()
    return value if len(value) <= max_chars else value[:max_chars] + "\n...[dipotong untuk menjaga konteks AI]"


def _extract_docx_content(raw_bytes: bytes) -> str:
    from docx import Document as _Document
    doc = _Document(io.BytesIO(raw_bytes))
    blocks = []
    for paragraph in doc.paragraphs:
        text_value = paragraph.text.strip()
        if text_value:
            blocks.append(text_value)
    for idx, table in enumerate(doc.tables, start=1):
        rows = []
        for row in table.rows:
            rows.append(" | ".join(cell.text.strip() for cell in row.cells))
        if rows:
            blocks.append(f"[TABEL {idx}]\n" + "\n".join(rows))
    return _trim_text("\n\n".join(blocks), 40000)


def _convert_legacy_ppt_to_pptx(raw_bytes: bytes):
    workdir = tempfile.mkdtemp(prefix="robomantap_ppt_")
    source_path = os.path.join(workdir, "source.ppt")
    with open(source_path, "wb") as handle:
        handle.write(raw_bytes)
    try:
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pptx", "--outdir", workdir, source_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=45,
            check=False,
        )
        output_path = os.path.join(workdir, "source.pptx")
        if result.returncode == 0 and os.path.exists(output_path):
            with open(output_path, "rb") as handle:
                return handle.read()
    except Exception:
        pass
    return None


def _extract_pptx_content(raw_bytes: bytes):
    from pptx import Presentation as _Presentation
    prs = _Presentation(io.BytesIO(raw_bytes))
    slides = []
    images = []
    for slide_no, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text_value = shape.text.strip()
                if text_value:
                    texts.append(text_value)
            try:
                if getattr(shape, "shape_type", None) == 13 and getattr(shape, "image", None):
                    images.append({
                        "slide": slide_no,
                        "mime_type": shape.image.content_type or "image/png",
                        "bytes": shape.image.blob,
                    })
            except Exception:
                pass
        notes = ""
        try:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            pass
        slides.append({"slide": slide_no, "text": "\n".join(texts), "notes": notes})
    return {"slides": slides, "images": images}


def _extract_pdf_content(raw_bytes: bytes):
    import fitz
    pdf = fitz.open(stream=raw_bytes, filetype="pdf")
    pages = []
    image_pages = []
    for page_no, page in enumerate(pdf, start=1):
        text_value = page.get_text("text").strip()
        pages.append({"page": page_no, "text": _trim_text(text_value, 7000)})
        if len(text_value) < 80 and len(image_pages) < 6:
            pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
            image_pages.append({"page": page_no, "mime_type": "image/png", "bytes": pix.tobytes("png")})
    return {"pages": pages, "images": image_pages}


def _call_gemini_contents(contents, is_json: bool = False, max_output_tokens: int = 12000):
    clients = get_gemini_clients()
    if not clients:
        return None
    for client in clients:
        for model_name in QUIZ_MODELS:
            try:
                config_kwargs = {"max_output_tokens": max_output_tokens}
                if is_json:
                    config_kwargs["response_mime_type"] = "application/json"
                if model_name.startswith("gemini-3."):
                    config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_level="high")
                else:
                    config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=0, include_thoughts=False)
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(**config_kwargs),
                )
                if response and response.text:
                    return response.text
            except Exception:
                continue
    return None


def _vision_describe(image_bytes: bytes, mime_type: str, source_label: str) -> str:
    prompt = f"""
Anda adalah Vision Reader RoboMANTAP.
Sumber visual: {source_label}
Analisis fakta yang benar-benar terlihat pada gambar. Baca teks, angka, label, tabel, diagram, grafik, rumus, dan struktur visual yang relevan. Jangan mengarang bagian yang tidak terlihat. Gunakan Bahasa Indonesia yang rapi dan detail yang cukup untuk menjadi sumber soal serta media ajar.
"""
    try:
        part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        response = _call_gemini_contents([prompt, part], is_json=False, max_output_tokens=5000)
        return _trim_text(response or "", 8000)
    except Exception as exc:
        print(f"Vision material gagal: {exc}")
        return ""


def build_material_knowledge_pack(uploaded_files) -> dict:
    files = []
    for item in uploaded_files or []:
        name = str(getattr(item, "name", "material"))
        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED_MATERIAL_EXTENSIONS:
            continue
        raw = item.getvalue() if hasattr(item, "getvalue") else bytes(item)
        record = {
            "name": name,
            "extension": ext,
            "mime_type": SUPPORTED_MATERIAL_EXTENSIONS[ext],
            "size": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
            "content": "",
            "locator_map": [],
            "visual_analysis": [],
        }
        try:
            if ext == ".docx":
                record["content"] = _extract_docx_content(raw)
                record["locator_map"] = [{"type": "document", "locator": "paragraphs/tables"}]
            elif ext in {".pptx", ".ppt"}:
                pptx_bytes = raw if ext == ".pptx" else _convert_legacy_ppt_to_pptx(raw)
                if not pptx_bytes:
                    record["content"] = "Format .ppt lama diterima, tetapi runtime tidak menyediakan konverter LibreOffice. Gunakan .pptx untuk ekstraksi otomatis."
                else:
                    parsed = _extract_pptx_content(pptx_bytes)
                    blocks = []
                    for slide in parsed["slides"]:
                        block = f"[SLIDE {slide['slide']}]\n{slide['text']}"
                        if slide.get("notes"):
                            block += f"\n[CATATAN PEMBICARA]\n{slide['notes']}"
                        blocks.append(block)
                        record["locator_map"].append({"type": "slide", "locator": slide["slide"]})
                    record["content"] = _trim_text("\n\n".join(blocks), 50000)
                    for image in parsed["images"][:6]:
                        analysis = _vision_describe(image["bytes"], image["mime_type"], f"{name}, slide {image['slide']}")
                        if analysis:
                            record["visual_analysis"].append({"locator": f"slide {image['slide']}", "analysis": analysis})
            elif ext == ".pdf":
                parsed = _extract_pdf_content(raw)
                pages = []
                for page in parsed["pages"]:
                    pages.append(f"[HALAMAN {page['page']}]\n{page['text']}")
                    record["locator_map"].append({"type": "page", "locator": page["page"]})
                record["content"] = _trim_text("\n\n".join(pages), 50000)
                for image in parsed["images"][:6]:
                    analysis = _vision_describe(image["bytes"], image["mime_type"], f"{name}, halaman {image['page']}")
                    if analysis:
                        record["visual_analysis"].append({"locator": f"halaman {image['page']}", "analysis": analysis})
            else:
                analysis = _vision_describe(raw, record["mime_type"], name)
                record["content"] = analysis
                if analysis:
                    record["visual_analysis"].append({"locator": "image", "analysis": analysis})
        except Exception as exc:
            record["content"] = f"Ekstraksi gagal untuk {name}: {exc}"
        files.append(record)

    source_blocks = []
    for record in files:
        source_blocks.append(
            f"### FILE: {record['name']}\n"
            f"Format: {record['extension']}\n"
            f"SHA-256: {record['sha256']}\n"
            f"Konten:\n{record.get('content','')}"
        )
        for visual in record.get("visual_analysis", []):
            source_blocks.append(f"Visual {visual['locator']}:\n{visual['analysis']}")
    return {
        "files": files,
        "file_count": len(files),
        "source_text": _trim_text("\n\n".join(source_blocks), 110000),
    }


def save_material_bundle_to_db(bundle_code: str, bundle: dict, config: dict | None = None) -> bool:
    conn = init_db_connection()
    if not conn:
        return False
    query = """
    INSERT INTO material_hub (bundle_code, config, bundle_data, created_at, updated_at)
    VALUES (:code, :cfg, :data, NOW() AT TIME ZONE 'Asia/Jakarta', NOW() AT TIME ZONE 'Asia/Jakarta')
    ON CONFLICT (bundle_code) DO UPDATE SET
        config = EXCLUDED.config,
        bundle_data = EXCLUDED.bundle_data,
        updated_at = NOW() AT TIME ZONE 'Asia/Jakarta';
    """
    try:
        with conn.session as session:
            session.execute(text(query), {
                "code": bundle_code.strip().upper(),
                "cfg": json.dumps(config or {}, ensure_ascii=False, default=str),
                "data": json.dumps(bundle, ensure_ascii=False, default=str),
            })
            session.commit()
        return True
    except Exception as exc:
        print(f"Error save_material_bundle_to_db: {exc}")
        return False


def get_material_bundle_from_db(bundle_code: str):
    conn = init_db_connection()
    if not conn:
        return None
    try:
        with conn.session as session:
            row = session.execute(
                text("SELECT bundle_data, config FROM material_hub WHERE UPPER(bundle_code)=UPPER(:code)"),
                {"code": bundle_code.strip()},
            ).fetchone()
        if not row:
            return None
        bundle = row[0] if isinstance(row[0], dict) else json.loads(row[0])
        config = row[1] if isinstance(row[1], dict) else json.loads(row[1] or "{}")
        bundle["config"] = config
        bundle["bundle_code"] = bundle_code.strip().upper()
        return bundle
    except Exception as exc:
        print(f"Error get_material_bundle_from_db: {exc}")
        return None


def _source_context_for_prompt(source_pack: dict | None, max_chars: int = 90000) -> str:
    if not source_pack:
        return "Tidak ada sumber materi terlampir."
    return _trim_text(source_pack.get("source_text", ""), max_chars)


def validate_and_repair_quiz(quiz: list, source_pack: dict | None, config: dict) -> list:
    if not quiz:
        return []
    count = len(quiz)
    jenjang = str(config.get("jenjang", "MTs"))
    mapel = str(config.get("mapel", ""))
    bahasa = str(config.get("bahasa", ""))
    konteks = str(config.get("konteks", ""))
    expected_labels = option_labels_for_jenjang(jenjang)
    option_count = len(expected_labels)
    language_guidance = build_language_guidance(mapel, bahasa, konteks, source_pack)
    source_locator_rule = "Tambahkan source_locator bila sumber tersedia; jangan memalsukan nomor halaman/slide."
    prompt = f"""
Anda adalah QA Validator RoboMANTAP untuk kuis madrasah.
Periksa lalu perbaiki draft soal. Pertahankan tepat {count} soal.
Konfigurasi: {json.dumps(config, ensure_ascii=False)}

KEBIJAKAN OPSI BERDASARKAN JENJANG:
- Jenjang MTs: tepat 4 opsi, label A-D.
- Jenjang MA: tepat 5 opsi, label A-E.
- Opsi harus berupa pilihan yang masuk akal dan hanya satu yang benar.
- Bila draft MA masih berisi 4 opsi, WAJIB tambahkan satu distraktor berkualitas menjadi E tanpa mengubah kebenaran jawaban.
- Bila draft memiliki format label yang tidak rapi, normalkan menjadi A. ... hingga label terakhir.

ATURAN BAHASA:
{language_guidance}

SUMBER UTAMA:
{_source_context_for_prompt(source_pack)}

DRAFT:
{json.dumps(quiz, ensure_ascii=False)}

ATURAN KUALITAS:
- Tepat {count} soal.
- Tepat {option_count} opsi per soal: {', '.join(expected_labels)}.
- correct_answer harus persis sama dengan salah satu opsi lengkap.
- Hindari ambiguitas, petunjuk jawaban yang terlalu mudah, duplikasi, dan pilihan yang tumpang tindih.
- Untuk soal numerik, solution_basis harus memuat proses hitungan inti dan hasil akhir.
- Jangan mengarang fakta inti yang tidak didukung sumber ketika sumber tersedia.
- {source_locator_rule}
- Keluarkan JSON murni.

FORMAT:
{{"quiz":[{{"id":1,"question":"...","options":["{expected_labels[0]}. ...", "{expected_labels[1]}. ...", "{expected_labels[2]}. ...", "{expected_labels[3]}. ..."{', "' + expected_labels[4] + '. ...' if option_count == 5 else ''}],"correct_answer":"{expected_labels[0]}. ...","solution_basis":"...","source_locator":"..."}}]}}
"""
    raw = call_gemini_with_rotation(prompt, is_json=True)
    if not raw:
        return quiz
    try:
        payload = json.loads(clean_json_text(raw), strict=False)
        validated = payload.get("quiz", [])
        return validated if isinstance(validated, list) and len(validated) == count else quiz
    except Exception:
        return quiz

def generate_media_ajar_ai(*, mapel: str, jenjang: str, kelas: str, topik: str, jumlah_slide: int,
                           durasi_menit: int, bahasa: str = "Bahasa Indonesia", gaya: str = "Premium UPN",
                           mode_presentasi: str = "Interactive", source_pack: dict | None = None):
    prompt = f"""
Anda adalah RoboMANTAP Media Architect. Rancang storyboard PPTX siap kelas.
Mapel: {mapel}
Jenjang: {jenjang}
Kelas: {kelas}
Topik: {topik}
Jumlah slide: tepat {jumlah_slide}
Durasi: {durasi_menit} menit
Bahasa: {bahasa}
Gaya: {gaya}
Mode presenter: {mode_presentasi}
SUMBER UTAMA:
{_source_context_for_prompt(source_pack)}
Prinsip: gunakan sumber sebagai fakta utama; jangan memenuhi slide dengan paragraf; variasikan konsep, visual/diagram, contoh, aktivitas, HOTS/diskusi, mini quiz, rangkuman/refleksi; setiap slide non-cover memiliki ide visual shape sederhana; speaker_notes harus memberi arahan guru.
JSON murni:
{{"deck_title":"...","deck_subtitle":"...","slides":[{{"slide":1,"type":"cover","title":"...","subtitle":"...","body":[],"visual_type":"hero","visual_content":"...","speaker_notes":"...","source_locator":"..."}}]}}
"""
    raw = call_gemini_with_rotation(prompt, is_json=True)
    if not raw:
        return None
    try:
        data = json.loads(clean_json_text(raw), strict=False)
        slides = data.get("slides", [])
        if not isinstance(slides, list) or len(slides) != jumlah_slide:
            return None
        for index, slide in enumerate(slides, start=1):
            slide["slide"] = index
            slide.setdefault("body", [])
            slide.setdefault("visual_type", "concept")
            slide.setdefault("visual_content", "")
            slide.setdefault("speaker_notes", "")
            slide.setdefault("source_locator", "")
        return data
    except Exception:
        return None


def build_media_ajar_pptx(storyboard: dict, config: dict, logo_path: str | None = None) -> bytes:
    from pptx import Presentation
    from pptx.dml.color import RGBColor as PptRGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches as PptInches, Pt as PptPt

    prs = Presentation()
    prs.slide_width = PptInches(13.333333)
    prs.slide_height = PptInches(7.5)
    NAVY = PptRGBColor(7, 15, 35)
    NAVY2 = PptRGBColor(15, 23, 42)
    EMERALD = PptRGBColor(16, 185, 129)
    MINT = PptRGBColor(167, 243, 208)
    WHITE = PptRGBColor(248, 250, 252)
    SLATE = PptRGBColor(148, 163, 184)
    GOLD = PptRGBColor(245, 158, 11)
    BLUE = PptRGBColor(59, 130, 246)

    def text_box(slide, value, x, y, w, h, size=20, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
        shape = slide.shapes.add_textbox(PptInches(x), PptInches(y), PptInches(w), PptInches(h))
        frame = shape.text_frame
        frame.clear(); frame.word_wrap = True
        paragraph = frame.paragraphs[0]
        paragraph.text = str(value or "")
        paragraph.alignment = align
        run = paragraph.runs[0]
        run.font.size = PptPt(size); run.font.bold = bold; run.font.color.rgb = color
        return shape

    def background(slide, color):
        fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = color

    def card(slide, value, x, y, w, h, accent=EMERALD, size=14):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptInches(x), PptInches(y), PptInches(w), PptInches(h))
        shape.fill.solid(); shape.fill.fore_color.rgb = NAVY2; shape.line.color.rgb = PptRGBColor(51,65,85)
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptInches(x), PptInches(y), PptInches(0.08), PptInches(h))
        strip.fill.solid(); strip.fill.fore_color.rgb = accent; strip.line.fill.background()
        text_box(slide, value, x+0.22, y+0.15, w-0.4, h-0.25, size, False, WHITE)

    for data in storyboard.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_type = str(data.get("type", "concept")).lower()
        title = data.get("title", "RoboMANTAP Media")
        if slide_type == "cover":
            background(slide, NAVY)
            text_box(slide, "U.PROJECT NEXUS • ROBOMANTAP", 0.7, 0.7, 7.5, 0.3, 12, True, MINT)
            text_box(slide, title, 0.7, 1.7, 10.4, 1.35, 34, True, WHITE)
            text_box(slide, data.get("subtitle", ""), 0.72, 3.15, 9.8, 0.75, 18, False, SLATE)
            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptInches(0.72), PptInches(4.3), PptInches(2.7), PptInches(0.52))
            pill.fill.solid(); pill.fill.fore_color.rgb = EMERALD; pill.line.fill.background()
            text_box(slide, f"{config.get('kelas','')} • {config.get('mapel','')}", 0.86, 4.43, 2.45, 0.2, 10, True, NAVY)
        else:
            background(slide, NAVY)
            text_box(slide, "ROBOMANTAP • MEDIA STUDIO", 0.55, 0.25, 4.8, 0.25, 10, True, MINT)
            text_box(slide, title, 0.55, 0.75, 11.9, 0.72, 25, True, WHITE)
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptInches(0.55), PptInches(1.52), PptInches(1.1), PptInches(0.05))
            line.fill.solid(); line.fill.fore_color.rgb = EMERALD; line.line.fill.background()
            visual_type = str(data.get("visual_type", "concept")).lower()
            visual = str(data.get("visual_content", ""))
            body = [str(x) for x in (data.get("body") or []) if str(x).strip()]
            if visual_type in {"diagram", "flow", "process"}:
                steps = body[:5] or [visual]
                width = 11.7 / max(1, len(steps))
                for i, item in enumerate(steps):
                    card(slide, item, 0.7 + i*width, 2.15, width-0.22, 2.0, EMERALD if i % 2 == 0 else GOLD, 13)
            elif visual_type == "table":
                rows = [r for r in visual.splitlines() if r.strip()] or body
                y = 2.05
                for row in rows[:6]:
                    parts = [p.strip() for p in row.split("|")]
                    if len(parts) <= 1:
                        card(slide, row, 0.8, y, 11.7, 0.6, EMERALD, 13)
                    else:
                        cell_w = 11.7 / min(3, len(parts))
                        for c, value in enumerate(parts[:3]):
                            card(slide, value, 0.8 + c*cell_w, y, cell_w-0.12, 0.6, EMERALD if c == 0 else BLUE, 12)
                    y += 0.74
            elif visual_type in {"activity", "discussion", "quiz"}:
                card(slide, visual or (body[0] if body else "Diskusikan dengan kelompok."), 0.75, 2.0, 11.7, 1.55, GOLD, 18)
                for i, item in enumerate(body[1:5]):
                    card(slide, item, 0.75 + (i % 2)*5.95, 3.9 + (i//2)*1.0, 5.65, 0.8, EMERALD, 13)
            else:
                text_box(slide, visual, 0.8, 2.0, 11.55, 1.25, 23, True, MINT)
                y = 3.35
                for i, item in enumerate(body[:5]):
                    card(slide, item, 0.8, y, 11.55, 0.56, EMERALD if i % 2 == 0 else BLUE, 12)
                    y += 0.67
            source = data.get("source_locator", "")
            if source:
                text_box(slide, f"Sumber: {source}", 9.1, 7.02, 3.6, 0.2, 8, False, SLATE, PP_ALIGN.RIGHT)
        try:
            slide.notes_slide.notes_text_frame.text = data.get("speaker_notes", "") or ""
        except Exception:
            pass

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def generate_quiz_batch(jenjang: str, mapel: str, stage: str, selected_submateri: list):
    submateri_text = ", ".join(selected_submateri) if selected_submateri else "Semua Submateri Terintegrasi"
    option_labels = option_labels_for_jenjang(jenjang)
    option_count = len(option_labels)
    language_guidance = build_language_guidance(mapel, "", "OMI / Olimpiade")

    stage_descriptions = {
        "Internal": "Internal: Fokus pada diagnostik, pemetaan bidang, dan penguatan konsep dasar.",
        "Kab/Kota": "Kab/Kota: Fokus pada pilihan ganda terstandar CBT, HOTS, dan analisis data.",
        "Provinsi": "Provinsi: Fokus pada analisis lintas konsep, pilihan ganda kompleks, serta keterkaitan sains, teknologi, dan nilai keislaman.",
        "Nasional": "Nasional: Fokus pada tingkat lanjutan (High-Level HOTS), eksplorasi problem solving, analisis eksperimen, dan penalaran ilmiah mendalam."
    }
    stage_description = stage_descriptions.get(stage, "Fokus pada penguatan konsep OMI.")

    if _is_arabic_subject(mapel):
        arabic_distribution = "Semua soal dapat menggunakan Bahasa Arab Fusha yang natural; jangan membatasi jumlah soal berbahasa Arab."
    elif _is_religious_subject(mapel):
        arabic_distribution = "Gunakan istilah/kutipan Arab asli hanya pada bagian yang memang berkaitan dengan materi keagamaan dan didukung sumber/konsep."
    else:
        arabic_distribution = "Bahasa utama tetap Bahasa Indonesia; Bahasa Arab hanya untuk istilah yang relevan."

    system_prompt = f"""
Anda adalah Pelatih Utama Bina Prestasi OMI 2026 (Olimpiade Sains & Matematika Al Irsyad) untuk tingkat {jenjang}.
Rancang 1 paket latihan CBT berisi TEPAT 10 SOAL PILIHAN GANDA yang orisinal, presisi, dan tematik OMI.

Spesifikasi:
- Jenjang: {jenjang}
- Bidang / Mata Pelajaran: {mapel}
- Tahap Pembinaan: {stage} ({stage_description})
- Cakupan Submateri: {submateri_text}
- Jumlah opsi WAJIB: {option_count} opsi, dengan label {', '.join(option_labels)}.

INTEGRASI TEMATIK & BAHASA MADRASAH:
- Konteks boleh mengaitkan Lingkungan, Teknologi, Kehidupan Sehari-hari, atau Nilai-Nilai Keislaman bila relevan.
- {language_guidance}
- {arabic_distribution}
- Jangan menggunakan transliterasi Latin untuk istilah Arab yang memang harus ditulis dalam aksara Arab.

ATURAN FORMATTING:
- Jangan sertakan field hint atau solution di sini.
- Angka biasa, nominal uang, satuan, dan jam ditulis sebagai teks biasa tanpa backslash.
- LaTeX $...$ hanya untuk rumus matematika asli, pecahan, akar, dan variabel.
- Jangan menaruh kalimat bahasa Indonesia di dalam delimiter matematika.
- Jangan membuat perintah LaTeX ilegal seperti '\\60.000.000' atau '\\14'.

Format keluaran JSON murni:
{{
  "quiz": [
    {{
      "id": 1,
      "question": "...",
      "options": ["{option_labels[0]}. ...", "{option_labels[1]}. ...", "{option_labels[2]}. ...", "{option_labels[3]}. ..."{', "' + option_labels[4] + '. ...' if option_count == 5 else ''}],
      "correct_answer": "{option_labels[0]}. ..."
    }}
  ]
}}
"""

    raw_response = call_gemini_with_rotation(system_prompt, is_json=True)
    if not raw_response:
        show_error("⚠️ Waduh kuota sedang penuh nih. Silakan coba klik lagi ya...")
        return []

    try:
        cleaned_response = clean_json_text(raw_response)
        data = json.loads(cleaned_response, strict=False)
        quiz_list = data.get("quiz", [])
        if not isinstance(quiz_list, list) or len(quiz_list) != 10:
            return []
        normalized = []
        for idx, q in enumerate(quiz_list, start=1):
            options = q.get("options", []) if isinstance(q, dict) else []
            norm_options, _ = normalize_quiz_options(options, jenjang)
            if len(norm_options) != option_count:
                return []
            answer = str(q.get("correct_answer", "")).strip()
            answer_label, answer_body = _split_option_label(answer, 0)
            answer_match = next((opt for opt in norm_options if opt.startswith(f"{answer_label}.")), None)
            if answer_match is None:
                answer_match = next((opt for opt in norm_options if opt == answer), "")
            if not answer_match:
                return []
            normalized.append({
                "id": idx,
                "question": str(q.get("question", "")).strip(),
                "options": format_latex_options(norm_options),
                "correct_answer": answer_match,
            })
        return normalized
    except Exception as e:
        show_error(f"Gagal memproses format soal: {e}")
        return []

def get_ai_hint_stream(question: str, user_attempt: str, mapel: str = "Umum"):
    prompt = f"""
    Kamu adalah 'RoboMANTAP', teman belajar dan asisten AI yang ramah, santai, ceria, dan sangat suportif dari MTs & MA Al Irsyad Putri Bondowoso (MANTAP).
    Gunakan gaya bahasa memberi sapaan 'aku' dan 'kamu' yang bersahabat namun tetap edukatif.

    Mata Pelajaran: {mapel}
    Soal OMI: {question}
    Ide Pengerjaan Siswa: {user_attempt}

    Instruksi:
    - Jangan berikan salam pembuka yang berlebihan.
    - Berikan petunjuk atau bimbingan logika interaktif yang menyemangati dan memuji usaha siswa.
    - Bantu siswa menemukan celah penyelesaian soal bidang {mapel} ini secara natural, runtut, dan analitis step-by-step tanpa membocorkan jawaban akhir.
    - Gunakan format LaTeX $...$ HANYA jika terdapat notasi matematika/sains.
    """
    return stream_ai_text(prompt, max_output_tokens=STREAM_HINT_MAX_TOKENS)

def get_ai_solution_stream(question: str, correct_answer: str, mapel: str = "Umum"):
    prompt = f"""
    Kamu adalah Pembina OMI 2026. Berikan pembahasan komprehensif, runtut, dan analitis step-by-step untuk soal berikut.

    Bidang: {mapel}
    Soal:
    {question}

    Kunci Jawaban yang Benar: {correct_answer}

    Instruksi Pembahasan:
    - Jangan berikan salam pembuka yang berlebihan.
    - Jelaskan secara natural, tajam, dan edukatif mengapa jawaban tersebut benar.
    - Jika ada unsur Bahasa Arab, terjemahkan dan kupas secara runtut.
    - Jika ada hitungan, tunjukkan proses rumusnya dengan jelas.
    - WAJIB gunakan format LaTeX $...$ untuk semua notasi matematika/simbol fisika-kimia.
    """
    return stream_ai_text(prompt, max_output_tokens=STREAM_SOLUTION_MAX_TOKENS)


# ==============================================================================
# INTEGRASI DATABASE SUPABASE POSTGRESQL (STREAMLIT & RENDER READY)
# ==============================================================================
class DBWrapper:
    """Wrapper kompatibel untuk SQLAlchemy agar memiliki interface .session"""
    def __init__(self, engine):
        self.engine = engine
        self.SessionMaker = sessionmaker(bind=self.engine)

    @property
    def session(self):
        return self.SessionMaker()

    def query(self, sql_query: str, ttl: int = 0):
        with self.engine.connect() as connection:
            return pd.read_sql(text(sql_query), connection)

_db_conn_cache = None
def init_db_connection():
    global _db_conn_cache
    if _db_conn_cache is not None:
        return _db_conn_cache

    # 1. Cek Environment Variables (Render.com / .env)
    db_url = os.getenv("DATABASE_URL") or os.getenv("POSTGRES_URL")

    # 2. Fallback ke Streamlit Secrets
    if not db_url:
        try:
            if hasattr(st, "secrets") and "DATABASE_URL" in st.secrets:
                db_url = st.secrets["DATABASE_URL"]
        except Exception:
            pass

    if db_url:
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql+psycopg2://", 1)
        try:
            engine = create_engine(
                db_url,
                pool_pre_ping=True,
                pool_size=5,          # Batasi jumlah koneksi aktif
                max_overflow=10,       # Toleransi lonjakan koneksi sementara
                pool_recycle=1800,     # Refresh koneksi setiap 30 menit
                connect_args={"connect_timeout": 5}  # Batas waktu tunggu koneksi maksimal 5 detik
            )
            _db_conn_cache = DBWrapper(engine)
            return _db_conn_cache
        except Exception as e:
            print(f"Gagal koneksi SQLAlchemy: {e}")

    # 3. Fallback Native Streamlit Connection (Jika dijalankan di Streamlit Cloud)
    try:
        conn = st.connection("postgresql", type="sql")
        _db_conn_cache = conn
        return _db_conn_cache
    except Exception:
        pass

    return None

def create_table_if_not_exists():
    conn = init_db_connection()
    if not conn: 
        return
    
    query = """
    CREATE TABLE IF NOT EXISTS sesi_ujian (
        id_sesi VARCHAR(100) PRIMARY KEY,
        nama_siswa VARCHAR(100) NOT NULL,
        jenjang VARCHAR(50),
        mapel VARCHAR(50),
        soal_sekarang INT DEFAULT 1,
        detail_jawaban JSONB DEFAULT '[]'::jsonb,
        jumlah_benar INT DEFAULT 0,
        jumlah_salah INT DEFAULT 0,
        nilai_akhir INT DEFAULT 0,
        status VARCHAR(20) DEFAULT 'BERJALAN',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS kuis_custom (
        kode_kuis VARCHAR(20) PRIMARY KEY,
        config JSONB NOT NULL,
        quiz_data JSONB NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS material_hub (
        bundle_code VARCHAR(32) PRIMARY KEY,
        config JSONB DEFAULT '{}'::jsonb,
        bundle_data JSONB NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """
    try:
        with conn.session as s:
            s.execute(text(query))
            s.commit()
    except Exception as e:
        print(f"Error create_table_if_not_exists: {e}")

def publish_custom_quiz_to_db(kode_kuis: str, config: dict, quiz_data: list) -> bool:
    config = normalize_custom_timer_config(config)
    conn = init_db_connection()
    if not conn: 
        return False

    query = """
    INSERT INTO kuis_custom (kode_kuis, config, quiz_data, created_at)
    VALUES (:kode, :cfg, :quiz, NOW() AT TIME ZONE 'Asia/Jakarta')
    ON CONFLICT (kode_kuis) DO UPDATE SET
        config = EXCLUDED.config,
        quiz_data = EXCLUDED.quiz_data;
    """
    try:
        with conn.session as s:
            s.execute(text(query), {
                "kode": kode_kuis.strip().upper(),
                "cfg": json.dumps(config, default=str),
                "quiz": json.dumps(quiz_data, default=str)
            })
            s.commit()
            return True
    except Exception as e:
        print(f"Error publish_custom_quiz_to_db: {e}")
        return False

def get_custom_quiz_from_db(kode_kuis: str):
    conn = init_db_connection()
    if not conn: 
        return None

    query = "SELECT config, quiz_data FROM kuis_custom WHERE UPPER(kode_kuis) = UPPER(:kode)"
    try:
        with conn.session as s:
            result = s.execute(text(query), {"kode": kode_kuis.strip()}).fetchone()
            if result:
                cfg = result[0] if isinstance(result[0], dict) else json.loads(result[0])
                cfg = normalize_custom_timer_config(cfg)
                quiz = result[1] if isinstance(result[1], list) else json.loads(result[1])
                return {"config": cfg, "quiz": quiz}
    except Exception as e:
        print(f"Error get_custom_quiz_from_db: {e}")
    return None

def update_progress_siswa(
    session_id: str,
    nama: str,
    jenjang: str,
    mapel: str,
    soal_sekarang: int,
    detail_jawaban: list,
    status: str = "BERJALAN",
    is_custom: bool = False,
    user_answers_dict: dict = None,
    quiz_data_list: list = None,
    anti_cheat: dict = None
):
    """
    Menyimpan progress CBT dengan proteksi state final.

    Prinsip penting:
    - created_at hanya dibuat saat INSERT pertama.
    - Sesi SELESAI / TRIAL / ARCHIVED tidak boleh hidup kembali menjadi BERJALAN.
    - Metadata anti-cheat dipertahankan walau save-answer berikutnya datang terlambat.
    - Format detail_jawaban tetap kompatibel dengan format list lama.
    """
    conn = init_db_connection()
    if not conn:
        return False

    mapel_db = f"{mapel} (Quiz)" if (is_custom and "(Quiz)" not in mapel) else mapel

    detail_jawaban = detail_jawaban if isinstance(detail_jawaban, list) else []
    total_soal = len(detail_jawaban) if detail_jawaban else (
        len(quiz_data_list) if isinstance(quiz_data_list, list) and quiz_data_list else 10
    )
    jumlah_benar = sum(1 for x in detail_jawaban if x is True)
    jumlah_salah = sum(1 for x in detail_jawaban if x is False)

    if is_custom:
        nilai_akhir = int(round((jumlah_benar / total_soal) * 100)) if total_soal > 0 else 0
    else:
        nilai_akhir = (jumlah_benar * 4) - (jumlah_salah * 1)

    # Ambil metadata anti-cheat sebelumnya agar tidak hilang karena save-answer biasa.
    existing_anti_cheat = None
    try:
        with conn.session as s:
            existing = s.execute(
                text("""
                    SELECT detail_jawaban
                    FROM sesi_ujian
                    WHERE id_sesi = :id_sesi
                    LIMIT 1
                """),
                {"id_sesi": session_id},
            ).fetchone()

        if existing:
            raw_existing = existing[0]
            if isinstance(raw_existing, str):
                try:
                    raw_existing = json.loads(raw_existing)
                except Exception:
                    raw_existing = None

            if isinstance(raw_existing, dict):
                existing_anti_cheat = raw_existing.get("anti_cheat")
    except Exception as e:
        print(f"[DB WARN] Gagal membaca metadata anti-cheat lama: {e}")

    effective_anti_cheat = anti_cheat if isinstance(anti_cheat, dict) else existing_anti_cheat

    # Payload JSONB.
    # Tanpa metadata tambahan, format list lama tetap dipertahankan untuk kompatibilitas.
    if user_answers_dict is not None or quiz_data_list is not None or effective_anti_cheat is not None:
        payload = {
            "detail_boolean": detail_jawaban,
            "user_answers": user_answers_dict or {},
            "quiz_data": quiz_data_list or [],
        }

        if effective_anti_cheat is not None:
            payload["anti_cheat"] = effective_anti_cheat

        detail_json = json.dumps(payload, default=str)
    else:
        detail_json = json.dumps(detail_jawaban)

    query = """
    INSERT INTO sesi_ujian (
        id_sesi, nama_siswa, jenjang, mapel, soal_sekarang, detail_jawaban,
        jumlah_benar, jumlah_salah, nilai_akhir, status, created_at, updated_at
    )
    VALUES (
        :id_sesi, :nama, :jenjang, :mapel, :soal, :detail,
        :benar, :salah, :nilai, :status,
        NOW() AT TIME ZONE 'Asia/Jakarta',
        NOW() AT TIME ZONE 'Asia/Jakarta'
    )
    ON CONFLICT (id_sesi) DO UPDATE SET
        nama_siswa = EXCLUDED.nama_siswa,
        jenjang = EXCLUDED.jenjang,
        mapel = EXCLUDED.mapel,
        soal_sekarang = EXCLUDED.soal_sekarang,
        detail_jawaban = EXCLUDED.detail_jawaban,
        jumlah_benar = EXCLUDED.jumlah_benar,
        jumlah_salah = EXCLUDED.jumlah_salah,
        nilai_akhir = EXCLUDED.nilai_akhir,
    
        -- created_at adalah WAKTU MULAI.
        -- Jangan pernah mengambil created_at dari request/update berikutnya.
        created_at = sesi_ujian.created_at,
    
        -- Status final tidak boleh hidup kembali menjadi BERJALAN.
        status = CASE
            WHEN sesi_ujian.status IN ('SELESAI', 'TRIAL', 'ARCHIVED')
                THEN sesi_ujian.status
            ELSE EXCLUDED.status
        END,
    
        -- Untuk sesi aktif: heartbeat/save-answer memperbarui updated_at.
        -- Untuk sesi final: waktu selesai dikunci.
        updated_at = CASE
            WHEN sesi_ujian.status IN ('SELESAI', 'TRIAL', 'ARCHIVED')
                THEN sesi_ujian.updated_at
            ELSE
                NOW() AT TIME ZONE 'Asia/Jakarta'
        END;
    """

    try:
        with conn.session as s:
            s.execute(
                text(query),
                {
                    "id_sesi": session_id,
                    "nama": nama,
                    "jenjang": jenjang,
                    "mapel": mapel_db,
                    "soal": int(soal_sekarang or 1),
                    "detail": detail_json,
                    "benar": jumlah_benar,
                    "salah": jumlah_salah,
                    "nilai": nilai_akhir,
                    "status": status,
                },
            )
            s.commit()
        return True
    except Exception as e:
        print(f"Error update_progress_siswa: {e}")
        return False

def touch_session_heartbeat(session_id: str) -> bool:
    """Memperbarui heartbeat sesi aktif tanpa mengubah status final."""
    conn = init_db_connection()
    if not conn:
        return False

    query = """
    UPDATE sesi_ujian
    SET updated_at = NOW() AT TIME ZONE 'Asia/Jakarta'
    WHERE id_sesi = :id_sesi
      AND status = 'BERJALAN';
    """

    try:
        with conn.session as s:
            s.execute(text(query), {"id_sesi": session_id})
            s.commit()
        return True
    except Exception as e:
        print(f"Error touch_session_heartbeat: {e}")
        return False

def generate_lkpd_content(mapel: str, kelas: str, topik: str, source_pack: dict | None = None):
    source_context = _source_context_for_prompt(source_pack, 65000)
    prompt = f"""
    Anda adalah Tim Ahli Kurikulum Lembaga Pendidikan Al-Irsyad Al-Islamiyah Putri Bondowoso.
    Rancanglah isi Lembar Kerja Peserta Didik (LKPD) berbasis HOTS dan Terintegrasi Keislaman.

    Spesifikasi LKPD:
    - Mata Pelajaran: {mapel}
    - Kelas / Jenjang: {kelas}
    - Topik / Materi Utama: {topik}

    SUMBER MATERI UTAMA (gunakan sebagai dasar fakta bila tersedia):
    {source_context}

    ATURAN NOTASI MATEMATIKA, FISIKA, KIMIA & LATEX (SANGAT PENTING):
    1. DILARANG KERAS menggunakan simbol dollar ($) atau backslash (\\) untuk rumus/variabel!
    2. Untuk angka pangkat atau indeks, HANYA gunakan simbol Unicode atau HTML sederhana:
       - Pangkat/Eksponen: Gunakan Unicode (x², x³, t²) atau <sup>2</sup>, <sup>3</sup>.
       - Indeks/Bawah: Gunakan Unicode (H₂O, CO₂) atau <sub>2</sub>.
       - Simbol Matematika: Gunakan simbol langsung seperti '≠', 'π', '√', '±', '≤', '≥', '°C'.
    3. Contoh Penulisan Rumus yang Benar di dalam teks:
       - "ax² + bx + c = 0 dengan a ≠ 0"
       - "h(t) = -5t² + 40t"
       - "Luas kolam adalah x² meter dan panjangnya x + 6 meter"
    4. UNTUK PECAHAN (SANGAT PENTING):
       - WAJIB gunakan simbol Unicode Pecahan Tegak untuk semua pecahan umum.
       - DILARANG KERAS menulis pecahan biasa dengan garis miring seperti '1/4', '3/8', atau '1/2'!
    5. Untuk matriks/array, WAJIB gunakan blok $$...$$
       - Jangan menulis environment matriks tanpa delimiter matematika.
    
    Instruksi Penyusunan Konten:
    1. Tujuan Pembelajaran: Buatkan 3 poin tujuan berbasis indikator HOTS.
    2. Apersepsi & Ringkasan Konsep: Sajikan materi singkat, tajam, dan korelasikan dengan nilai-nilai Keislaman/Tadabbur Sains.
    3. Tugas Eksplorasi Mandiri: Buat 5 soal studi kasus/problem solving HOTS yang melatih logika nalar santri/siswi.
    4. Refleksi Keislaman: Tuliskan 1 kalimat hikmah/perenungan dari mempelajari materi {topik}.

    Format keluaran WAJIB objek JSON murni:
    {{
        "tujuan": ["Poin tujuan 1", "Poin tujuan 2", "Poin tujuan 3"],
        "ringkasan": "Teks ringkasan konsep dan keislaman...",
        "soal_1": "Pertanyaan eksplorasi HOTS nomor 1",
        "soal_2": "Pertanyaan eksplorasi HOTS nomor 2",
        "soal_3": "Pertanyaan eksplorasi HOTS nomor 3",
        "soal_4": "Pertanyaan eksplorasi HOTS nomor 4",
        "soal_5": "Pertanyaan eksplorasi HOTS nomor 5",
        "refleksi": "Kalimat hikmah/refleksi..."
    }}
    """

    raw_response = call_gemini_with_rotation(prompt, is_json=True)
    if not raw_response:
        return None

    try:
        cleaned_response = clean_json_text(raw_response)
        data = json.loads(cleaned_response, strict=False)
        return data
    except Exception:
        return None

def generate_custom_quiz_ai(
    *,
    mapel: str,
    jenjang: str,
    kelas: str,
    materi: str,
    submateri: str,
    jumlah_soal: int,
    kesulitan: str,
    tipe_soal: str,
    bahasa: str,
    konteks: str,
    timer_seconds: int,
    source_pack: dict | None = None,
):
    option_labels = option_labels_for_jenjang(jenjang)
    option_count = len(option_labels)
    source_context = _source_context_for_prompt(source_pack, 90000)
    language_guidance = build_language_guidance(mapel, bahasa, konteks, source_pack)
    option_example = ', '.join([f'"{label}. ..."' for label in option_labels])
    prompt = f"""
Anda adalah Question Architect RoboMANTAP untuk guru madrasah.
Buat tepat {jumlah_soal} soal pilihan ganda berkualitas tinggi untuk pembelajaran.

KONFIGURASI:
- Mata Pelajaran: {mapel}
- Jenjang: {jenjang}
- Kelas: {kelas}
- Materi: {materi}
- Submateri: {submateri or 'Tidak ditentukan / semua yang relevan'}
- Tingkat Kesulitan: {kesulitan}
- Tipe Soal: {tipe_soal}
- Bahasa: {bahasa}
- Konteks: {konteks}
- Batas Waktu Sesi: {timer_seconds} detik

KEBIJAKAN OPSI WAJIB:
- {jenjang}: tepat {option_count} pilihan jawaban.
- Label wajib: {', '.join(option_labels)}.
- MTs = A-D (4 pilihan).
- MA = A-E (5 pilihan).
- Jangan pernah menghasilkan opsi tambahan di luar label yang ditentukan.
- Hanya satu opsi benar.

KEBIJAKAN BAHASA:
{language_guidance}

SUMBER MATERI TERLAMPIR / SOURCE GROUNDING:
{source_context}

ATURAN KUALITAS:
0. Bila sumber materi terlampir tersedia, jadikan sumber tersebut sebagai sumber fakta utama. Jangan mengarang fakta inti di luar sumber.
1. Tepat {jumlah_soal} soal, jangan kurang dan jangan lebih.
2. Setiap soal memiliki tepat {option_count} opsi: {', '.join(option_labels)}.
3. Hanya satu opsi yang benar.
4. correct_answer harus persis sama dengan salah satu opsi lengkap.
5. Hindari ambiguitas, data yang kurang, dan asumsi yang tidak disebutkan.
6. Untuk soal numerik, solution_basis harus memuat proses hitungan inti dan hasil akhir.
7. Untuk HOTS/olimpiade, gunakan penalaran yang benar-benar relevan dengan level.
8. Jangan memasukkan jawaban atau pembahasan yang saling bertentangan.
9. Jika menggunakan LaTeX, gunakan $...$ dan escape backslash secara valid untuk JSON.
10. Jangan menambahkan markdown atau teks pembuka di luar JSON.
11. Untuk matriks/array, gunakan blok $$...$$ dan jangan menulis environment matriks tanpa delimiter.
12. Hindari transliterasi Latin untuk kosakata Arab yang memang harus ditulis dalam aksara Arab.
13. Untuk Bahasa Arab, gunakan Bahasa Arab Fusha/Modern Standard Arabic yang natural, dengan harakat selektif jika membantu kejelasan. Jangan membuat kutipan agama yang tidak didukung sumber.

OUTPUT JSON MURNI:
{{
  "quiz": [
    {{
      "id": 1,
      "question": "...",
      "options": [{option_example}],
      "correct_answer": "{option_labels[0]}. ...",
      "solution_basis": "..."
    }}
  ],
  "config": {{"duration_seconds": {timer_seconds}}}
}}
"""

    raw_response = call_gemini_with_rotation(prompt, is_json=True)
    if not raw_response:
        return []

    try:
        cleaned = clean_json_text(raw_response)
        data = json.loads(cleaned, strict=False)
    except Exception:
        match = re.search(r"\{.*\}", raw_response, flags=re.DOTALL)
        if not match:
            return []
        try:
            data = json.loads(match.group(0), strict=False)
        except Exception:
            return []

    quiz = data.get("quiz", [])
    if not isinstance(quiz, list) or len(quiz) != jumlah_soal:
        return []

    normalized = []
    for idx, item in enumerate(quiz, start=1):
        if not isinstance(item, dict):
            return []
        question = str(item.get("question", "")).strip()
        raw_options = item.get("options", [])
        answer = str(item.get("correct_answer", "")).strip()
        solution_basis = str(item.get("solution_basis", "")).strip()
        if not question or not isinstance(raw_options, list) or not solution_basis:
            return []

        options, _ = normalize_quiz_options(raw_options, jenjang)
        # Beri kesempatan validator memperbaiki draft yang jumlah opsinya belum tepat.
        if len(options) not in (option_count, 4, 5):
            return []
        answer_label, _ = _split_option_label(answer, 0)
        answer_match = next((x for x in options if x.startswith(f"{answer_label}.")), None)
        if answer_match is None and answer in options:
            answer_match = answer
        if answer_match is None:
            return []
        normalized.append({
            "id": idx,
            "question": question,
            "options": format_latex_options(options),
            "correct_answer": answer_match,
            "solution_basis": solution_basis,
            "source_locator": str(item.get("source_locator", "")).strip(),
        })

    validation_config = {
        "mapel": mapel,
        "jenjang": jenjang,
        "kelas": kelas,
        "materi": materi,
        "submateri": submateri,
        "kesulitan": kesulitan,
        "tipe_soal": tipe_soal,
        "bahasa": bahasa,
        "konteks": konteks,
        "option_count": option_count,
        "option_labels": option_labels,
    }
    validated = validate_and_repair_quiz(normalized, source_pack, validation_config)

    final_items = []
    for idx, item in enumerate(validated, start=1):
        if not isinstance(item, dict):
            continue
        opts, _ = normalize_quiz_options(item.get("options", []), jenjang)
        if len(opts) != option_count:
            continue
        opts = format_latex_options(opts)
        ans = str(item.get("correct_answer", "")).strip()
        ans_label, _ = _split_option_label(ans, 0)
        matching = next((x for x in opts if x.startswith(f"{ans_label}.")), None)
        if matching is None:
            exact = next((x for x in opts if x == ans), None)
            matching = exact
        if matching is None:
            continue
        final_items.append({
            "id": idx,
            "question": str(item.get("question", "")).strip(),
            "options": opts,
            "correct_answer": matching,
            "solution_basis": str(item.get("solution_basis", "")).strip(),
            "source_locator": str(item.get("source_locator", "")).strip(),
        })

    if len(final_items) == jumlah_soal:
        return final_items
    # Hanya boleh fallback ke raw draft jika raw draft sudah memenuhi aturan jenjang secara utuh.
    if all(len(x.get("options", [])) == option_count for x in normalized):
        return normalized
    return []

def check_active_session_from_db(nama_siswa: str, mapel: str):
    conn = init_db_connection()
    if not conn: 
        return None

    query = """
    SELECT id_sesi, detail_jawaban, created_at, soal_sekarang
    FROM sesi_ujian
    WHERE LOWER(TRIM(nama_siswa)) = LOWER(TRIM(:nama))
      AND (
          LOWER(TRIM(mapel)) = LOWER(TRIM(:mapel))
          OR LOWER(mapel) LIKE LOWER(:mapel_like)
      )
      AND status = 'BERJALAN'
    ORDER BY created_at DESC LIMIT 1;
    """
    try:
        with conn.session as s:
            res = s.execute(text(query), {
                "nama": nama_siswa.strip(), 
                "mapel": mapel.strip(),
                "mapel_like": f"%{mapel.strip()}%"
            }).fetchone()
            
            if res:
                detail_ans = res[1]
                if isinstance(detail_ans, str):
                    try:
                        detail_ans = json.loads(detail_ans)
                    except Exception:
                        detail_ans = []

                anti_cheat = None
                if isinstance(detail_ans, dict):
                    anti_cheat = detail_ans.get("anti_cheat")
                    detail_boolean = detail_ans.get("detail_boolean", [])
                    user_answers = detail_ans.get("user_answers", {})
                else:
                    detail_boolean = detail_ans
                    user_answers = {}

                return {
                    "id_sesi": res[0],
                    "detail_jawaban": detail_boolean if isinstance(detail_boolean, list) else [],
                    "user_answers": user_answers if isinstance(user_answers, dict) else {},
                    "anti_cheat": anti_cheat if isinstance(anti_cheat, dict) else None,
                    "created_at": res[2],
                    "soal_sekarang": res[3] if len(res) > 3 and res[3] is not None else 1
                }
    except Exception as e:
        print(f"Error check_active_session_from_db: {e}")
    return None

def load_session_review_from_db(session_id: str):
    """Membaca data sesi ujian dari Supabase untuk ditampilkan di Streamlit."""
    conn = init_db_connection()
    if not conn:
        return None

    query = "SELECT nama_siswa, jenjang, mapel, detail_jawaban, nilai_akhir FROM sesi_ujian WHERE id_sesi = :id"
    try:
        with conn.session as s:
            res = s.execute(text(query), {"id": session_id.strip()}).fetchone()
            if res:
                nama, jenjang, mapel, detail_raw, nilai = res
                
                # Parsing detail_jawaban (apakah berupa dict payload baru atau list boolean lama)
                if isinstance(detail_raw, dict):
                    payload = detail_raw
                elif isinstance(detail_raw, str):
                    try:
                        payload = json.loads(detail_raw)
                    except Exception:
                        payload = {}
                else:
                    payload = {}

                if isinstance(payload, dict):
                    raw_user_answers = payload.get("user_answers", {})
                    quiz_data = payload.get("quiz_data", [])
                else:
                    raw_user_answers = {}
                    quiz_data = []

                # Format ulang key dictionary ke integer agar cocok dengan state Streamlit
                formatted_user_answers = {}
                for k, v in raw_user_answers.items():
                    try:
                        formatted_user_answers[int(k)] = v
                    except ValueError:
                        formatted_user_answers[k] = v

                raw_mapel = str(mapel or "")
                return {
                    "nama": nama,
                    "jenjang": jenjang or "MA",
                    "mapel": raw_mapel.replace(" (Quiz)", "").replace(" (TKA)", "") if raw_mapel else "Kuis",
                    "activity_type": "TKA" if "(TKA)" in raw_mapel else ("Quiz" if "(Quiz)" in raw_mapel else "OMI"),
                    "quiz_data": quiz_data,
                    "user_answers": formatted_user_answers,
                    "nilai": nilai
                }
    except Exception as e:
        print(f"Error load_session_review_from_db: {e}")
    return None

def generate_individual_analysis_ai(nama_siswa: str, mapel: str, jenjang: str, nilai: int, detail_jawaban_list: list, quiz_data: list = None, user_answers: dict = None):
    """Generates a deep, prescriptive pedagogical analysis for an individual student."""
    
    total_soal = len(detail_jawaban_list) if detail_jawaban_list else 10
    benar = sum(1 for x in detail_jawaban_list if x is True)
    salah = sum(1 for x in detail_jawaban_list if x is False)
    kosong = total_soal - (benar + salah)

    # Susun detail analisis per nomor jika ada data quiz
    detail_soal_text = ""
    if quiz_data and user_answers:
        for idx, q in enumerate(quiz_data):
            u_ans = user_answers.get(idx, user_answers.get(str(idx), "Tidak Diisi"))
            c_ans = q.get("correct_answer", "-")
            status_item = "✅ BENAR" if u_ans == c_ans else "❌ SALAH"
            detail_soal_text += f"\n- Soal {idx+1} [{status_item}]: {q.get('question','')[:80]}... | Jwb Siswa: {u_ans} | Kunci: {c_ans}"

    prompt = f"""
    Anda adalah Pakar Evaluasi Pendidikan & Konsultan Pedagogi OMI/CBT untuk Madrasah Al Irsyad.
    Berikan Analisis Diagnosis Preskriptif yang tajam, akurat, dan sangat berguna bagi Guru mengenai hasil kuis santri berikut:

    DATA SISWA:
    - Nama Santri: {nama_siswa}
    - Mata Pelajaran: {mapel} ({jenjang})
    - Skor Akhir: {nilai} / 100
    - Transkrip Performa: {benar} Benar, {salah} Salah, {kosong} Kosong (Total {total_soal} Soal)
    {detail_soal_text}

    FORMAT KELUARAN (Gunakan Markdown rapi dengan Icon & Elemen Visual):

    ### 🎯 1. Profil Kognitif & Diagnosa Cepat
    - **Tingkat Penguasaan:** [Sangat Mahir / Berkembang / Perlu Intervensi Khusus / Miskonsepsi Berat]
    - **Pola Pengerjaan:** [Analisis singkat apakah siswa paham konsep dasar, sering salah hitung, atau asal tebak]

    ### 🔍 2. Bedah Akar Miskonsepsi
    - **Akar Masalah Utama:** [Jelaskan secara tajam di mana letak kelemahan logika/konsep siswa]
    - **Tipe Kesalahan:** [Konseptual / Prosedural-Kalkulasi / Penalaran HOTS]

    ### 💡 3. Rekomendasi Intervensi Guru (Preskriptif)
    - **Tindakan Remedial Guru:** [Langkah praktis yang harus dilakukan guru dalam 1-on-1 coaching]
    - **Fokus Latihan Lanjutan:** [Beri saran 2-3 topik/submateri spesifik yang harus diberikan ke siswa ini]

    ### 📲 4. Umpan Balik Santri & Ortu (WA-Ready)
    *(Pesan singkat, hangat, bernuansa keislaman & menyemangati yang bisa dicopy guru untuk dikirim via WA)*
    """

    return call_gemini_with_rotation(prompt, is_json=False)

# ==============================================================================
# HELPER PARSER MATHEMATICA & OMML WORD EQUATION
# ==============================================================================
def clean_math_string(text: str) -> str:
    """Pembersih notasi matematika/LaTeX untuk output Word dan dokumen."""
    if not text:
        return ""
    text = str(text)

    # Operator penting pada fungsi/invers, relasi, kalkulus, dll.
    replacements = {
        r"\rightarrow": "→", r"\to": "→", r"\Rightarrow": "⇒",
        r"\leftarrow": "←", r"\leftrightarrow": "↔",
        r"\circ": "∘", r"\circl": "∘",
        r"\times": "×", r"\cdot": "·", r"\div": "÷",
        r"\neq": "≠", r"\leq": "≤", r"\geq": "≥",
        r"\le": "≤", r"\ge": "≥", r"\pm": "±", r"\mp": "∓",
        r"\infty": "∞", r"\pi": "π", r"\alpha": "α", r"\beta": "β",
        r"\gamma": "γ", r"\delta": "δ", r"\theta": "θ",
        r"\lambda": "λ", r"\mu": "μ", r"\sigma": "σ",
        r"\subset": "⊂", r"\subseteq": "⊆", r"\supset": "⊃", r"\supseteq": "⊇",
        r"\in": "∈", r"\notin": "∉", r"\forall": "∀", r"\exists": "∃",
        r"\emptyset": "∅", r"\angle": "∠", r"\perp": "⊥", r"\parallel": "∥",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)

    # Jika AI mengirim token mentah "circl" dari \circ yang terpotong, normalkan juga.
    text = re.sub(r"(?<![A-Za-z])circl(?![A-Za-z])", "∘", text)

    text = re.sub(r'\\left\b\s*[\(\[\{\.\|]?', '(', text)
    text = re.sub(r'\\right\b\s*[\)\]\}\.\|]?', ')', text)
    text = re.sub(r'\\(?:dots|cdots|ldots)', '…', text)
    text = re.sub(r'\\sqrt\{([^}]+)\}', r'√(\1)', text)
    text = re.sub(r'\\sqrt\s*([a-zA-Z0-9_]+)', r'√\1', text)

    sup_map = str.maketrans("0123456789+-=()nxyi", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿˣʸⁱ")
    sub_map = str.maketrans("0123456789+-=()nixy", "₀₁₂₃₄⁵₆₇₈₉₊₋₌₍₎ₙᵢₓᵧ")
    text = re.sub(r'\^\{([^}]+)\}|\^([\-0-9a-zA-Z])', lambda m: (m.group(1) or m.group(2)).translate(sup_map), text)
    text = re.sub(r'\_\{([^}]+)\}|\_([0-9a-zA-Z])', lambda m: (m.group(1) or m.group(2)).translate(sub_map), text)

    text = text.replace("$", "")
    text = text.replace("left(", "(").replace("right)", ")").replace("dots", "…")
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r'\\([a-zA-Z]+)', r'\1', text).replace("\\", "")
    return re.sub(r'\s+', ' ', text).strip()

clean_math_text = clean_math_string

def add_omml_fraction(paragraph, num_text: str, den_text: str):
    """Menyisipkan struktur Pecahan Tegak Resmi Microsoft Word (Equation)."""
    num_clean = clean_math_string(num_text)
    den_clean = clean_math_string(den_text)
    
    omml_xml = (
        f'<m:oMath {nsdecls("m")}>'
        f'  <m:f>'
        f'    <m:num><m:r><m:t>{num_clean}</m:t></m:r></m:num>'
        f'    <m:den><m:r><m:t>{den_clean}</m:t></m:r></m:den>'
        f'  </m:f>'
        f'</m:oMath>'
    )
    paragraph._p.append(parse_xml(omml_xml))

def add_omml_matrix(paragraph, matrix_type: str, content: str):
    """Menyisipkan struktur Matriks 2D Bertingkat Resmi Microsoft Word (Equation)."""
    beg_chr, end_chr = "(", ")"
    if matrix_type == "bmatrix":
        beg_chr, end_chr = "[", "]"
    elif matrix_type in ["vmatrix", "Vmatrix"]:
        beg_chr, end_chr = "|", "|"
    elif matrix_type == "matrix":
        beg_chr, end_chr = "", ""

    rows = [r.strip() for r in re.split(r'\\\\|\\cr', content) if r.strip()]
    matrix_xml_rows = []
    for r in rows:
        cols = [c.strip() for c in r.split('&')]
        cols_xml = []
        for c in cols:
            cleaned_c = html.escape(clean_math_string(c))
            cols_xml.append(f'<m:e><m:r><m:t>{cleaned_c}</m:t></m:r></m:e>')
        matrix_xml_rows.append(f'<m:mr>{"".join(cols_xml)}</m:mr>')

    inner_matrix = f'<m:m>{"".join(matrix_xml_rows)}</m:m>'
    if beg_chr or end_chr:
        omml_xml = (
            f'<m:oMath {nsdecls("m")}>'
            f'  <m:d>'
            f'    <m:dPr>'
            f'      <m:begChr m:val="{beg_chr}"/>'
            f'      <m:endChr m:val="{end_chr}"/>'
            f'    </m:dPr>'
            f'    <m:e>{inner_matrix}</m:e>'
            f'  </m:d>'
            f'</m:oMath>'
        )
    else:
        omml_xml = f'<m:oMath {nsdecls("m")}>{inner_matrix}</m:oMath>'

    paragraph._p.append(parse_xml(omml_xml))

def append_text_with_fractions(paragraph, text: str, is_bold: bool = False, color_rgb: RGBColor = None):
    """Membagi paragraf: teks biasa, pecahan \\frac, dan matriks \\begin{...matrix}."""
    if not text:
        return

    math_pattern = re.compile(
        r'\\begin\{(?P<mtype>[pbvV]?matrix)\}(?P<mcontent>.*?)\\end\{(?P=mtype)\}|\\(?:f|tf)rac\{(?P<num>[^}]+)\}\{(?P<den>[^}]+)\}',
        re.DOTALL
    )
    last_idx = 0

    for match in math_pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            plain_part = clean_math_string(text[last_idx:start])
            if plain_part:
                run = paragraph.add_run(plain_part + " ")
                run.bold = is_bold
                if color_rgb:
                    run.font.color.rgb = color_rgb

        if match.group('mtype'):
            add_omml_matrix(paragraph, match.group('mtype'), match.group('mcontent'))
        elif match.group('num'):
            add_omml_fraction(paragraph, match.group('num'), match.group('den'))
        
        run_space = paragraph.add_run(" ")
        run_space.bold = is_bold
        last_idx = end

    if last_idx < len(text):
        plain_part = clean_math_string(text[last_idx:])
        if plain_part:
            run = paragraph.add_run(plain_part)
            run.bold = is_bold
            if color_rgb:
                run.font.color.rgb = color_rgb
# ==============================================================================
# HELPER PARSER MARKDOWN & OMML UNTUK DOKUMEN KEDINASAN
# ==============================================================================
def append_markdown_formatted_text(doc, md_text: str):
    """
    Mengurai teks Markdown AI (### header, **bold**, bullet -) 
    menjadi paragraf Word berstandar kedinasan yang rapi dan elegan.
    """
    if not md_text:
        return

    lines = md_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue

        # 1. Sub-Header (###)
        if line_str.startswith('###'):
            p_h = doc.add_paragraph()
            p_h.paragraph_format.space_before = Pt(8)
            p_h.paragraph_format.space_after = Pt(3)
            h_text = line_str.lstrip('#').strip()
            run_h = p_h.add_run(h_text)
            run_h.bold = True
            run_h.font.size = Pt(11)
            run_h.font.color.rgb = RGBColor(6, 78, 59) # Emerald Kedinasan

        # 2. Poin Bullet (- atau *)
        elif line_str.startswith(('-', '*')):
            p_b = doc.add_paragraph()
            p_b.paragraph_format.left_indent = Inches(0.25)
            p_b.paragraph_format.space_after = Pt(2)
            
            content = line_str.lstrip('-*').strip()
            parts = re.split(r'(\*\*.*?\*\*)', content)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    clean_part = part[2:-2]
                    append_text_with_fractions(p_b, clean_part, is_bold=True, color_rgb=RGBColor(6, 78, 59))
                else:
                    append_text_with_fractions(p_b, part, is_bold=False)

        # 3. Paragraf Biasa
        else:
            p_gen = doc.add_paragraph()
            p_gen.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_gen.paragraph_format.space_after = Pt(4)
            
            parts = re.split(r'(\*\*.*?\*\*)', line_str)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    clean_part = part[2:-2]
                    append_text_with_fractions(p_gen, clean_part, is_bold=True)
                else:
                    append_text_with_fractions(p_gen, part, is_bold=False)

# ==============================================================================
# GENERATOR GRAFIK VISUAL MATPLOTLIB (IN-MEMORY BUFFER)
# ==============================================================================
def generate_class_visual_charts(data_siswa: list):
    """Menghasilkan 2 gambar grafik statistik (Donut Chart & Horizontal Bar Chart)."""
    if not data_siswa:
        return None, None

    df = pd.DataFrame(data_siswa)
    
    # 1. Klasifikasi Kognitif Kelas
    mahir = len(df[df['nilai_akhir'] >= 80])
    berkembang = len(df[(df['nilai_akhir'] >= 50) & (df['nilai_akhir'] < 80)])
    intervensi = len(df[df['nilai_akhir'] < 50])

    # --- GRAFIK 1: DONUT CHART KATEGORI PENGUASAAN ---
    fig1, ax1 = plt.subplots(figsize=(5, 3.5), dpi=200)
    labels = ['Sangat Mahir (≥80)', 'Berkembang (50-79)', 'Perlu Intervensi (<50)']
    sizes = [mahir, berkembang, intervensi]
    colors = ['#10B981', '#F59E0B', '#EF4444']
    
    # Filter kategori bernilai > 0
    active_labels = [l for l, s in zip(labels, sizes) if s > 0]
    active_sizes = [s for s in sizes if s > 0]
    active_colors = [c for c, s in zip(colors, sizes) if s > 0]

    wedges, texts, autotexts = ax1.pie(
        active_sizes, labels=active_labels, colors=active_colors,
        autopct='%1.1f%%', startangle=140, pctdistance=0.75,
        textprops=dict(color="#1E293B", fontsize=8, weight="bold")
    )
    
    # Lubang Donut Tengah
    centre_circle = plt.Circle((0, 0), 0.50, fc='white')
    fig1.gca().add_artist(centre_circle)
    ax1.set_title("Distribusi Tingkat Penguasaan Kognitif Santri", fontsize=10, fontweight='bold', pad=12, color='#064E3B')
    plt.tight_layout()

    buf1 = io.BytesIO()
    plt.savefig(buf1, format='png', bbox_inches='tight')
    plt.close(fig1)
    buf1.seek(0)

    # --- GRAFIK 2: BAR CHART NILAI PER-SISWA ---
    fig2, ax2 = plt.subplots(figsize=(6, max(3, len(df) * 0.4)), dpi=200)
    df_sorted = df.sort_values(by='nilai_akhir', ascending=True)

    names = [n[:18] + '...' if len(n) > 18 else n for n in df_sorted['nama_siswa']]
    scores = df_sorted['nilai_akhir']
    bar_colors = ['#10B981' if s >= 80 else ('#F59E0B' if s >= 50 else '#EF4444') for s in scores]

    bars = ax2.barh(names, scores, color=bar_colors, height=0.6)
    ax2.set_xlim(0, 100)
    ax2.set_xlabel("Skor Akhir Ujian", fontsize=8, fontweight='bold', color='#475569')
    ax2.set_title("Peringkat & Performa Skor Santri", fontsize=10, fontweight='bold', pad=12, color='#064E3B')
    
    # Label Nilai di Ujung Bar
    for bar in bars:
        width = bar.get_width()
        ax2.text(width + 1.5, bar.get_y() + bar.get_height()/2, f'{int(width)}',
                 va='center', ha='left', fontsize=8, fontweight='bold', color='#1E293B')

    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()

    buf2 = io.BytesIO()
    plt.savefig(buf2, format='png', bbox_inches='tight')
    plt.close(fig2)
    buf2.seek(0)

    return buf1, buf2

# ==============================================================================
# MAIN DOCX REPORT GENERATOR (STANDAR RESMI KEDINASAN & VISUAL ANALYTICS)
# ==============================================================================
def generate_corporate_executive_docx_report(config: dict, data_siswa: list, collective_ai_summary: str = None) -> bytes:
    """
    Menghasilkan Dokumen Word (.docx) Laporan Rekapitulasi Ujian & Diagnostik Eksekutif
    Presisi Kedinasan: Kop Instansi, Garis XML Solid, Tabel Metadata 3-Kolom, Chart Visual, & Diagnosis Individu.
    """
    doc = Document()

    # Set Margin Halaman Standard 1 Inci
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # --- 1. KOP INSTITUSI RESMI (PERSIS FORMAT ASLI) ---
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = False
    
    cells = header_table.rows[0].cells
    cells[0].width = Inches(1.6)
    cells[1].width = Inches(4.7)
    cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Logo Instansi
    p_logo = cells[0].paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(1.5))
    else:
        r_logo = p_logo.add_run("[LOGO MANTAP]")
        r_logo.bold = True
        r_logo.font.size = Pt(12)
        r_logo.font.color.rgb = RGBColor(6, 78, 59)

    # Teks Kop Instansi
    p_title = cells[1].paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    
    r1 = p_title.add_run("LAPORAN REKAPITULASI & DIAGNOSIS CBT\n")
    r1.bold = True
    r1.font.size = Pt(13)
    r1.font.color.rgb = RGBColor(6, 78, 59) # Warna Hijau Edukasi Kedinasan
    
    r2 = p_title.add_run("Madrasah Aliyah dan Tsanawiyah Al-Irsyad Al-Islamiyah Putri Bondowoso\n")
    r2.bold = True
    r2.font.size = Pt(10.5)
    
    # Tanggal Presisi WIB dengan Nama Bulan Bahasa Indonesia
    now_wib = datetime.utcnow() + timedelta(hours=7)
    nama_bulan = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    tgl_presisi = f"{now_wib.day} {nama_bulan[now_wib.month]} {now_wib.year}"
    
    r3 = p_title.add_run(f"Tanggal Penerbitan: {tgl_presisi} WIB")
    r3.italic = True
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(100, 100, 100)

    # --- 2. GARIS PEMBATAS HIJAU SOLID (XML EMBED) ---
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(10)
    pBdr = parse_xml(
        r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:bottom w:val="single" w:sz="20" w:space="1" w:color="064E3B"/>'
        r'</w:pBdr>'
    )
    p_div._p.get_or_add_pPr().append(pBdr)

    # --- 3. JUDUL EXECUTIVE REPORT ---
    p_paket = doc.add_paragraph()
    p_paket.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_paket.paragraph_format.space_after = Pt(10)
    r_paket = p_paket.add_run(f"EVALUASI UJIAN {config.get('mapel', 'Mata Pelajaran').upper()}")
    r_paket.bold = True
    r_paket.font.size = Pt(13)
    r_paket.font.color.rgb = RGBColor(6, 78, 59)

    # --- 4. METADATA KUIS (TABEL 3 KOLOM PRESISI & LURUS SEJAJAR) ---
    meta_items = [
        ("Jenjang / Kelas", f"{config.get('jenjang', '-')} ({config.get('kelas', 'Semua Kelas')})"),
        ("Materi Utama", f"{clean_math_string(config.get('materi', 'Ujian Terintegrasi'))}"),
        ("Total Peserta Ujian", f"{len(data_siswa)} Santri Terdaftar")
    ]

    meta_table = doc.add_table(rows=0, cols=3)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    for label, val in meta_items:
        row_cells = meta_table.add_row().cells
        
        # Kolom 1: Label
        p0 = row_cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        row_cells[0].width = Inches(1.5)
        
        # Kolom 2: Titik Dua
        p1 = row_cells[1].paragraphs[0]
        r1 = p1.add_run(":")
        r1.bold = True
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        row_cells[1].width = Inches(0.2)
        
        # Kolom 3: Nilai
        p2 = row_cells[2].paragraphs[0]
        p2.add_run(val)
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        row_cells[2].width = Inches(4.8)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- 5. RINGKASAN KPI KELAS ---
    total = len(data_siswa)
    if total > 0:
        scores = [s.get("nilai_akhir", 0) for s in data_siswa]
        avg_score = sum(scores) / total
        max_score = max(scores)
        min_score = min(scores)
    else:
        avg_score = max_score = min_score = 0

    kpi_table = doc.add_table(rows=2, cols=4)
    kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi_table.style = 'Table Grid'
    
    headers = ["Total Santri", "Rata-Rata Kelas", "Skor Tertinggi", "Skor Terendah"]
    values = [str(total), f"{avg_score:.1f}", str(max_score), str(min_score)]

    for i, h in enumerate(headers):
        cell = kpi_table.cell(0, i)
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, v in enumerate(values):
        cell = kpi_table.cell(1, i)
        cell.text = v
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(12.5)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(6, 78, 59)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # --- 6. VISUAL ANALYTICS (MATPLOTLIB IN-DOCX CHART) ---
    chart1_buf, chart2_buf = generate_class_visual_charts(data_siswa)
    if chart1_buf and chart2_buf:
        doc.add_heading("1. Analisis Visual Performa Kelas", level=2)
        p_charts = doc.add_paragraph()
        p_charts.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_charts.add_run().add_picture(chart1_buf, width=Inches(3.1))
        p_charts.add_run("   ")
        p_charts.add_run().add_picture(chart2_buf, width=Inches(3.3))
        
        doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- 7. DIAGNOSIS KOLEKTIF KELAS ---
    doc.add_heading("2. Diagnosis Pedagogis Kolektif (AI Engine)", level=2)
    summary_text = collective_ai_summary or "Berdasarkan evaluasi kuis, tingkat penguasaan konsep santri berada pada kategori baik."
    append_markdown_formatted_text(doc, summary_text)
    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # --- 8. TABEL REKAPITULASI DETAIL SISWA ---
    doc.add_heading("3. Tabel Rekapitulasi Nilai Seluruh Santri", level=2)
    
    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    headers_text = ["No", "Nama Santri", "Kelas", "Benar", "Salah", "Nilai Akhir"]
    for i, text in enumerate(headers_text):
        hdr_cells[i].text = text
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, s in enumerate(data_siswa, start=1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(s.get("nama_siswa", "-"))
        row_cells[2].text = str(s.get("jenjang", "-"))
        row_cells[3].text = str(s.get("jumlah_benar", 0))
        row_cells[4].text = str(s.get("jumlah_salah", 0))
        row_cells[5].text = str(s.get("nilai_akhir", 0))
        
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[5].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    # --- 9. LAMPIRAN DIAGNOSIS PRESKRIPTIF PER-SANTRI ---
    doc.add_page_break()
    
    p_lamp = doc.add_paragraph()
    p_lamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_lamp = p_lamp.add_run("LAMPIRAN: DIAGNOSIS PRESKRIPTIF INDIVIDU SANTRI")
    r_lamp.bold = True
    r_lamp.font.size = Pt(13)
    r_lamp.font.color.rgb = RGBColor(6, 78, 59)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    for idx, s in enumerate(data_siswa, start=1):
        p_indiv = doc.add_paragraph()
        r_indiv = p_indiv.add_run(f"Santri #{idx}: {s.get('nama_siswa', '-')} (Skor Akhir: {s.get('nilai_akhir', 0)})")
        r_indiv.bold = True
        r_indiv.font.size = Pt(11)
        r_indiv.font.color.rgb = RGBColor(6, 78, 59)

        detail_raw = s.get("detail_jawaban", [])
        quiz_data = None
        user_answers = None
        if isinstance(detail_raw, dict):
            detail_boolean = detail_raw.get("detail_boolean", [])
            user_answers = detail_raw.get("user_answers", {})
            quiz_data = detail_raw.get("quiz_data", [])
        elif isinstance(detail_raw, list):
            detail_boolean = detail_raw
        else:
            detail_boolean = []

        # Diagnosis AI Tanpa Nomor Soal
        indiv_report_text = generate_individual_analysis_ai(
            nama_siswa=s.get("nama_siswa", "-"),
            mapel=config.get("mapel", "Kuis"),
            jenjang=s.get("jenjang", "MA"),
            nilai=s.get("nilai_akhir", 0),
            detail_jawaban_list=detail_boolean,
            quiz_data=quiz_data,
            user_answers=user_answers
        )

        # Render Teks Markdown AI ke Paragraf Kedinasan
        append_markdown_formatted_text(doc, indiv_report_text)
        doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # Simpan ke BytesIO Buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
