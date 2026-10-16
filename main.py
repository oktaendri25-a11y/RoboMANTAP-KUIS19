import io
import os
import json
import random
import uuid
from datetime import datetime, timezone, timedelta
from typing import Dict, Any
from fastapi import FastAPI, Request, Form, HTTPException, status
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from ai_engine import (
    create_table_if_not_exists,
    get_custom_quiz_from_db,
    normalize_custom_timer_config,
    update_progress_siswa,
    touch_session_heartbeat,
    get_ai_hint_stream
)
ai_hint_cache = {}
app = FastAPI(title="RoboMANTAP CBT Engine")

# Path absolut agar folder templates selalu terdeteksi di Linux Render
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# Inisialisasi Tabel DB
try:
    create_table_if_not_exists()
except Exception as e:
    print(f"[DB INIT WARN] {e}")

# In-Memory Session Storage
STUDENT_SESSIONS: Dict[str, Dict[str, Any]] = {}

STREAMLIT_URL = "https://robomantap-intelligence.streamlit.app/" # Sesuaikan dengan URL Streamlit app.py kamu

@app.get("/", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="student_login.html", 
        context={"error": None}
    )

def parse_wib_datetime(dt_str):
    """Helper untuk membaca format ISO dari Supabase dan mengonversinya ke datetime WIB."""
    if not dt_str:
        return None
    try:
        s = str(dt_str).strip().replace(" ", "T")
        dt_raw = datetime.fromisoformat(s.replace('Z', '+00:00'))
        if dt_raw.tzinfo is not None:
            return dt_raw.astimezone(timezone(timedelta(hours=7))).replace(tzinfo=None)
        return dt_raw
    except Exception as e:
        print(f"[DATE PARSE ERROR] Gagal parse tanggal '{dt_str}': {e}")
        return None


def normalize_jenjang(value: str) -> str:
    value = str(value or "MA").strip()
    upper = value.upper()
    if "MTS" in upper:
        return "MTs"
    if "MA" in upper:
        return "MA"
    return value


def build_detail_answers(sess: dict):
    """Bangun status benar/salah/kosong dari jawaban sesi saat ini."""
    quiz = sess.get("quiz", []) if isinstance(sess, dict) else []
    answers = sess.get("answers", {}) if isinstance(sess, dict) else {}
    detail = []

    for idx, item in enumerate(quiz):
        user_ans = answers.get(idx) or answers.get(str(idx))
        if user_ans is None:
            detail.append(None)
            continue

        correct_ans = item.get("correct_answer") or item.get("key")
        detail.append(user_ans == correct_ans)

    return detail


def normalized_anti_cheat(sess: dict) -> dict:
    current = sess.get("anti_cheat") if isinstance(sess, dict) else None
    if not isinstance(current, dict):
        current = {}

    return {
        "detected": bool(current.get("detected", False)),
        "reason": str(current.get("reason", "") or ""),
        "violation_count": max(0, int(current.get("violation_count", 0) or 0)),
        "max_violations": max(1, int(current.get("max_violations", 3) or 3)),
    }


@app.post("/verify-token", response_class=HTMLResponse)
async def verify_token(
    request: Request,
    nama: str = Form(...),
    kelas: str = Form(...),
    absen: str = Form(...),
    token: str = Form(...)
):
    clean_token = token.strip().upper()
    nama_clean = nama.strip()
    kelas_clean = kelas.strip()
    absen_clean = absen.strip()

    quiz_package = get_custom_quiz_from_db(clean_token)

    # 1. Cek keberadaan token
    if not quiz_package:
        return templates.TemplateResponse(
            request=request,
            name="student_login.html",
            context={"error": "❌ Kode Kuis tidak ditemukan atau belum diterbitkan!"}
        )

    config = normalize_custom_timer_config(quiz_package.get("config", {}) or {})

    # 2. Waktu saat ini dalam WIB
    now_wib = datetime.now(timezone.utc).astimezone(
        timezone(timedelta(hours=7))
    ).replace(tzinfo=None)

    active_from_raw = config.get("active_from")
    active_until_raw = config.get("active_until")

    # 3. Validasi masa aktif kuis
    if active_from_raw and active_until_raw:
        dt_from = parse_wib_datetime(active_from_raw)
        dt_until = parse_wib_datetime(active_until_raw)

        if dt_from and now_wib < dt_from:
            time_start = config.get("time_start_str", dt_from.strftime("%H:%M"))
            return templates.TemplateResponse(
                request=request,
                name="student_login.html",
                context={
                    "error": f"⏰ Kuis Belum Dibuka! Kuis baru dapat diakses pada pukul {time_start} WIB."
                }
            )

        if dt_until and now_wib > dt_until:
            time_end = config.get("time_end_str", dt_until.strftime("%H:%M"))
            return templates.TemplateResponse(
                request=request,
                name="student_login.html",
                context={
                    "error": f"❌ Kode Kuis Sudah Kedaluwarsa! Masa aktif kuis ini telah berakhir pada pukul {time_end} WIB."
                }
            )

    # 4. Ambil soal dan buat sesi ujian.
    master_quiz = quiz_package.get("quiz", []) or []
    packages = config.get("packages") or [master_quiz]
    selected_quiz = random.choice(packages) if packages else master_quiz

    session_id = str(uuid.uuid4())[:8]
    start_time_utc = datetime.now(timezone.utc)

    anti_cheat = {
        "detected": False,
        "reason": "",
        "violation_count": 0,
        "max_violations": 3,
    }

    STUDENT_SESSIONS[session_id] = {
        "nama": nama_clean,
        "kelas": kelas_clean,
        "absen": absen_clean,
        "token": clean_token,
        "config": config,
        "quiz": selected_quiz,
        "answers": {},
        "current_index": 0,
        "start_time": start_time_utc,
        "anti_cheat": anti_cheat,
    }

    # 5. Catat sesi ke database SEBELUM redirect.
    # created_at di sini = saat siswa benar-benar menekan Mulai.
    try:
        update_progress_siswa(
            session_id=session_id,
            nama=nama_clean,
            jenjang=normalize_jenjang(config.get("jenjang", "MA")),
            mapel=config.get("mapel", "Kuis"),
            soal_sekarang=1,
            detail_jawaban=[None] * len(selected_quiz),
            status="BERJALAN",
            is_custom=True,
            anti_cheat=anti_cheat,
        )
    except Exception as e:
        print(f"⚠️ Warning Sync Supabase (Start Exam): {e}")

    return RedirectResponse(
        url=f"/exam/{session_id}",
        status_code=status.HTTP_303_SEE_OTHER
    )

# ==============================================================================
# ROUTE WORKSPACE EXAM (MENGGABUNGKAN GET & POST KE SINGLE-FILE STUDENT_EXAM.HTML)
# ==============================================================================
# ==============================================================================
# HELPER FUNCTION & ROUTE WORKSPACE EXAM (PENGGANTI /start-exam LAMA)
# ==============================================================================

async def render_exam_workspace(request: Request, session_id: str):
    sess = STUDENT_SESSIONS.get(session_id)
    if not sess:
        return RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)

    # Sesi final tidak boleh dibuka kembali melalui URL /exam.
    if sess.get("finished") or normalized_anti_cheat(sess).get("detected"):
        review_url = f"{STREAMLIT_URL.rstrip('/')}/?review_session={session_id}"
        return RedirectResponse(
            url=review_url,
            status_code=status.HTTP_303_SEE_OTHER
        )

    # 1. Catat Waktu Mulai Pertama Kali (Mencegah Reset Timer saat Refresh)
    if "start_time" not in sess:
        sess["start_time"] = datetime.now(timezone.utc)

    # 2. Hitung Sisa Waktu Ujian REAL-TIME dari satu sumber kebenaran.
    # Jangan membaca timer_m saja: itu akan memotong kuis 2j/3j menjadi 30 menit.
    config = normalize_custom_timer_config(sess.get("config", {}) or {})
    sess["config"] = config
    duration_seconds = int(config.get("timer_seconds", 0) or 0)
    now_utc = datetime.now(timezone.utc)
    start_time = sess["start_time"]
    if start_time.tzinfo is None:
        start_time = start_time.replace(tzinfo=timezone.utc)

    elapsed_s = max(0.0, (now_utc - start_time).total_seconds())
    remaining_s = max(0, int(duration_seconds - elapsed_s)) if duration_seconds > 0 else 0

    # 3. Render Single-File student_exam.html
    return templates.TemplateResponse(
        request=request,
        name="student_exam.html", 
        context={
            "session_id": session_id,
            "sess": sess,
            "mapel": sess.get("config", {}).get("mapel", "Kuis RoboMANTAP"),
            "materi": sess.get("config", {}).get("materi", "Umum"),
            "nama": sess.get("nama", "Siswa"),
            "kelas": sess.get("kelas", "-"),
            "absen": sess.get("absen", "-"),
            "jumlah_soal": len(sess.get("quiz", [])),
            "durasi_menit": (duration_seconds + 59) // 60 if duration_seconds > 0 else 0,
            "duration_seconds": duration_seconds,
            "durasi_label": (
                f"{duration_seconds // 3600} jam {(duration_seconds % 3600) // 60} menit"
                if duration_seconds >= 3600 and duration_seconds % 60 == 0
                else f"{duration_seconds // 3600} jam {(duration_seconds % 3600) // 60} menit {duration_seconds % 60} detik"
                if duration_seconds >= 3600
                else f"{duration_seconds // 60} menit {duration_seconds % 60} detik"
                if duration_seconds > 0
                else "Tanpa batas waktu"
            ),
            "quiz_json": json.dumps(sess.get("quiz", [])),
            "answers_json": json.dumps(sess.get("answers", {})),
            "remaining_seconds": remaining_s,
            "anti_cheat_json": json.dumps(normalized_anti_cheat(sess))
        }
    )

# Route 1: Menerima submit form POST dari Login
@app.post("/start-exam", response_class=HTMLResponse)
async def start_exam_post(request: Request, session_id: str = Form(...)):
    return await render_exam_workspace(request, session_id)

# Route 2: Menerima akses langsung via URL GET / refresh browser
@app.get("/exam/{session_id}", response_class=HTMLResponse)
async def start_exam_get(request: Request, session_id: str):
    return await render_exam_workspace(request, session_id)

# ==============================================================================
# 1. API SAVE ANSWER (MENYIMPAN JAWABAN REAL-TIME)
# ==============================================================================
@app.post("/api/save-answer")
async def save_answer(
    session_id: str = Form(...),
    q_index: int = Form(...),
    answer: str = Form(...)
):
    sess = STUDENT_SESSIONS.get(session_id)

    if not sess:
        return HTMLResponse(content="", status_code=404)

    if sess.get("finished"):
        return HTMLResponse(content="", status_code=409)

    # Sumber kebenaran durasi tetap di server. Jawaban baru ditolak setelah
    # batas waktu tercapai, sehingga timer tidak hanya bergantung pada browser.
    config = normalize_custom_timer_config(sess.get("config", {}) or {})
    sess["config"] = config
    duration_seconds = int(config.get("timer_seconds", 0) or 0)
    if duration_seconds > 0:
        start_time = sess.get("start_time")
        if start_time:
            if start_time.tzinfo is None:
                start_time = start_time.replace(tzinfo=timezone.utc)
            if (datetime.now(timezone.utc) - start_time).total_seconds() >= duration_seconds:
                return HTMLResponse(content="TIMEOUT", status_code=409)

    quiz = sess.get("quiz", []) or []
    if q_index < 0 or q_index >= len(quiz):
        return HTMLResponse(content="", status_code=400)

    if "answers" not in sess:
        sess["answers"] = {}

    answer_clean = str(answer).strip()
    sess["answers"][q_index] = answer_clean
    sess["answers"][str(q_index)] = answer_clean
    sess["current_index"] = q_index

    detail_ans = build_detail_answers(sess)
    anti_cheat = normalized_anti_cheat(sess)

    try:
        update_progress_siswa(
            session_id=session_id,
            nama=sess.get("nama", "Siswa"),
            jenjang=normalize_jenjang(sess.get("config", {}).get("jenjang", "Kuis")),
            mapel=sess.get("config", {}).get("mapel", "Kuis"),
            soal_sekarang=q_index + 1,
            detail_jawaban=detail_ans,
            status="BERJALAN",
            is_custom=True,
            anti_cheat=anti_cheat,
        )
    except Exception as e:
        print(f"⚠️ Warning Sync Supabase (Save Answer): {e}")

    return HTMLResponse(content="", status_code=200)


@app.post("/api/session-heartbeat")
async def session_heartbeat(session_id: str = Form(...)):
    sess = STUDENT_SESSIONS.get(session_id)
    if not sess:
        return HTMLResponse(content="", status_code=404)

    if sess.get("finished"):
        return HTMLResponse(content="", status_code=204)

    try:
        touch_session_heartbeat(session_id)
    except Exception as e:
        print(f"⚠️ Warning Heartbeat: {e}")

    return HTMLResponse(content="", status_code=204)


@app.post("/api/anti-cheat")
async def anti_cheat_event(
    session_id: str = Form(...),
    violation_count: int = Form(0),
    reason: str = Form("Pindah tab")
):
    sess = STUDENT_SESSIONS.get(session_id)
    if not sess:
        return HTMLResponse(content="", status_code=404)

    anti_cheat = normalized_anti_cheat(sess)
    incoming_count = max(0, int(violation_count or 0))
    server_count = anti_cheat["violation_count"]

    # Server adalah sumber hitungan minimum. Nilai client yang lebih tinggi boleh
    # dipakai untuk menangkap event yang baru saja terjadi.
    effective_count = max(server_count, incoming_count)
    effective_count = min(effective_count, anti_cheat["max_violations"])

    anti_cheat["violation_count"] = effective_count

    if effective_count > 0:
        anti_cheat["reason"] = str(reason or "Pindah tab").strip()[:100] or "Pindah tab"

    is_forced_stop = effective_count >= anti_cheat["max_violations"]
    anti_cheat["detected"] = is_forced_stop

    sess["anti_cheat"] = anti_cheat

    detail_ans = build_detail_answers(sess)

    try:
        update_progress_siswa(
            session_id=session_id,
            nama=sess.get("nama", "Siswa"),
            jenjang=normalize_jenjang(sess.get("config", {}).get("jenjang", "MA")),
            mapel=sess.get("config", {}).get("mapel", "Kuis"),
            soal_sekarang=max(1, int(sess.get("current_index", 0)) + 1),
            detail_jawaban=detail_ans,
            status="SELESAI" if is_forced_stop else "BERJALAN",
            is_custom=True,
            user_answers_dict=sess.get("answers", {}) if is_forced_stop else None,
            quiz_data_list=sess.get("quiz", []) if is_forced_stop else None,
            anti_cheat=anti_cheat,
        )
    except Exception as e:
        print(f"⚠️ Warning Sync Supabase (Anti-Cheat): {e}")

    if is_forced_stop:
        sess["finished"] = True
        return HTMLResponse(content="STOP", status_code=200)

    return HTMLResponse(content="WARN", status_code=200)

# ==============================================================================
# 2. API SUBMIT EXAM (KALKULASI NILAI & Halaman student_result.html)
# ==============================================================================
@app.post("/submit-exam", response_class=HTMLResponse)
async def submit_exam(
    request: Request,
    session_id: str = Form(...),
    anti_cheat_detected: str = Form("0"),
    anti_cheat_reason: str = Form(""),
    anti_cheat_count: int = Form(0),
):
    sess = STUDENT_SESSIONS.get(session_id)
    if not sess:
        return RedirectResponse(url="/", status_code=status.HTTP_303_SEE_OTHER)

    quiz = sess.get("quiz", []) or []
    answers = sess.get("answers", {}) or {}

    # Ekstraksi nama depan/panggilan siswa
    nama_lengkap = sess.get("nama", "").strip()
    nama_depan = nama_lengkap.split()[0] if nama_lengkap else "Santri MANTAP"

    benar = 0
    salah = 0
    kosong = 0
    detail_ans = []

    for idx, item in enumerate(quiz):
        user_ans = answers.get(idx) or answers.get(str(idx))
        correct_ans = item.get("correct_answer") or item.get("key")

        if not user_ans:
            kosong += 1
            detail_ans.append(False)
        elif user_ans == correct_ans:
            benar += 1
            detail_ans.append(True)
        else:
            salah += 1
            detail_ans.append(False)

    total_soal = len(quiz)
    skor = int(round((benar / total_soal) * 100)) if total_soal > 0 else 0

    # Ambil metadata anti-cheat yang sudah disimpan server.
    anti_cheat = normalized_anti_cheat(sess)

    form_detected = str(anti_cheat_detected).strip().lower() in {"1", "true", "yes", "on"}
    form_count = max(0, int(anti_cheat_count or 0))

    if form_detected or form_count > 0:
        anti_cheat["violation_count"] = min(
            anti_cheat["max_violations"],
            max(anti_cheat["violation_count"], form_count),
        )
        if anti_cheat["violation_count"] > 0:
            anti_cheat["reason"] = (
                str(anti_cheat_reason or "Pindah tab").strip()[:100]
                or "Pindah tab"
            )

        anti_cheat["detected"] = (
            form_detected
            or anti_cheat["violation_count"] >= anti_cheat["max_violations"]
        )

    sess["anti_cheat"] = anti_cheat

    try:
        update_progress_siswa(
            session_id=session_id,
            nama=nama_lengkap,
            jenjang=normalize_jenjang(sess.get("config", {}).get("jenjang", "MA")),
            mapel=sess.get("config", {}).get("mapel", "Matematika"),
            soal_sekarang=total_soal,
            detail_jawaban=detail_ans,
            status="SELESAI",
            is_custom=True,
            user_answers_dict=answers,
            quiz_data_list=quiz,
            anti_cheat=anti_cheat,
        )
    except Exception as e:
        print(f"⚠️ Warning Sync Supabase (Submit Exam): {e}")

    sess["finished"] = True

    base_url = STREAMLIT_URL.rstrip("/")
    target_streamlit_url = f"{base_url}/?review_session={session_id}"

    return templates.TemplateResponse(
        request=request,
        name="student_result.html",
        context={
            "nama": nama_depan,
            "skor": skor,
            "benar": benar,
            "salah": salah,
            "kosong": kosong,
            "total": total_soal,
            "streamlit_url": target_streamlit_url,
        }
    )

# ==============================================================================
# 3. API HINT AI (ROBOMANTAP CONSULTATION)
# ==============================================================================
@app.post("/api/hint", response_class=HTMLResponse)
async def handle_hint_request(
    curr_idx: int = Form(...),
    mapel: str = Form(""),
    question: str = Form(""),
    attempt_input: str = Form("")
):
    attempt_str = attempt_input.strip()

    # Validasi input kosong
    if not attempt_str:
        return """
        <div class="p-2.5 bg-blue-950/40 border border-blue-500/30 text-blue-300 rounded-lg text-xs mt-2">
            💡 Tolong ketik sedikit ide kamu dulu ya, biar RoboMANTAP bisa kasih petunjuk yang pas!
        </div>
        """

    # Cek cache lokal
    hint_key = (mapel, curr_idx, question, attempt_str)

    if hint_key in ai_hint_cache:
        hint_text = ai_hint_cache[hint_key]
    else:
        try:
            # Panggil AI Stream dari ai_engine.py
            hint_chunks = [chunk for chunk in get_ai_hint_stream(question, attempt_str, mapel)]
            hint_text = "".join(hint_chunks)

            # Simpan ke cache jika tidak error
            if hint_text and "⚠️" not in hint_text:
                ai_hint_cache[hint_key] = hint_text
        except Exception as e:
            hint_text = f"⚠️ Maaf, RoboMANTAP sedang sibuk sebentar. Coba tekan tombol diskusi lagi ya! ({e})"

    return f"""
    <div class="p-3 bg-slate-900 border border-slate-800 rounded-lg text-slate-200 text-xs leading-relaxed mt-2 animate-fade-in">
        <div class="font-bold text-emerald-400 mb-1">🧕🏼 RoboMANTAP:</div>
        <div>{hint_text}</div>
    </div>
    """
