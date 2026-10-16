"""Runtime-only shared helpers for RoboMANTAP HTMX service.

This file intentionally excludes Streamlit/admin/report generators from the main
RoboMANTAP ai_engine.py. Public helper names used by this service are preserved.
"""
import json
import os
import re
import time
from datetime import datetime, timezone, timedelta

from google import genai
from google.genai import types
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker

def _load_gemini_api_keys():
    """Read one key or a JSON/comma-separated list without logging secret values."""
    raw = os.getenv("GEMINI_API_KEYS") or os.getenv("GEMINI_API_KEY") or ""
    if isinstance(raw, (list, tuple)):
        values = list(raw)
    else:
        raw = str(raw).strip()
        values = []
        if raw.startswith("[") and raw.endswith("]"):
            try:
                parsed = json.loads(raw)
                if isinstance(parsed, list):
                    values = parsed
            except (ValueError, TypeError):
                # Tolerate multiline values with backslash line continuations.
                cleaned = raw.replace("\\\r\n", "").replace("\\\n", "")
                try:
                    parsed = json.loads(cleaned)
                    if isinstance(parsed, list):
                        values = parsed
                except (ValueError, TypeError):
                    values = []
        if not values:
            values = raw.split(",")
    result = []
    for value in values:
        key = str(value).strip().strip("\"' ")
        if key and key not in result:
            result.append(key)
    return result

api_keys = _load_gemini_api_keys()
print(f"[HTMX] Gemini configuration: {len(api_keys)} key(s) detected")
QUIZ_MODELS = ("gemini-3.5-flash-lite", "gemini-3.1-flash-lite")
STREAM_MODELS = ("gemini-3.5-flash-lite", "gemini-3.1-flash-lite")
STREAM_HINT_MAX_TOKENS = 9000
STREAM_SOLUTION_MAX_TOKENS = 9000
STREAM_TIMEOUT_MS = 90_000
_clients_cache = None
_db_conn_cache = None

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

def stream_ai_text(prompt: str, max_output_tokens: int = STREAM_HINT_MAX_TOKENS):
    yield from _stream_from_clients(
        prompt,
        max_output_tokens=max_output_tokens,
    )

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

class DBWrapper:
    """Small SQLAlchemy session wrapper used by the HTMX runtime."""
    def __init__(self, engine):
        self.engine = engine
        self.SessionMaker = sessionmaker(bind=self.engine)

    @property
    def session(self):
        return self.SessionMaker()

def init_db_connection():
    """Create/reuse PostgreSQL connection from Vercel Environment Variables."""
    global _db_conn_cache
    if _db_conn_cache is not None:
        return _db_conn_cache
    db_url = os.getenv("DATABASE_URL") or os.getenv("POSTGRES_URL")
    if not db_url:
        print("[DB CONFIG ERROR] DATABASE_URL/POSTGRES_URL belum diatur.")
        return None
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif db_url.startswith("postgresql://") and "+" not in db_url.split(":", 1)[0]:
        db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
    try:
        engine = create_engine(
            db_url,
            pool_pre_ping=True,
            pool_size=1,
            max_overflow=0,
            pool_recycle=300,
            connect_args={"connect_timeout": 5},
        )
        _db_conn_cache = DBWrapper(engine)
        return _db_conn_cache
    except Exception as exc:
        # Do not print credentials/connection URL to logs.
        print(f"[DB INIT ERROR] {type(exc).__name__}: gagal membuat koneksi PostgreSQL.")
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

