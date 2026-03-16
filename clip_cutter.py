"""
Video Clip Cutter - ตัดคลิปจากไฟล์ MP4 ตามข้อมูลใน Google Sheets
อ่านข้อมูลจาก Google Sheets (sheet: Hoshi) แล้วตัดคลิปด้วย FFmpeg
รองรับทั้งไฟล์ MP4 ในเครื่องและ YouTube URL (ผ่าน yt-dlp)
อัปเดต status กลับไปใน Sheet ผ่าน Google Apps Script Web API

Columns:
  A: status | B: ประเภท | C: ชื่อคลิป | D: ช่วงเริ่มต้น (URL)
  E: ช่วงเริ่มต้น | F: ช่วงสิ้นสุด | G: Path | H: Note
"""

import json
import os
import re
import subprocess
import sys
import urllib.request
import urllib.parse
from pathlib import Path
from datetime import date

# ==================== CONFIG ====================
API_KEY = "AIzaSyDKnNvb5bzkR2_-IDKJoBWX9lvYUaLDYzc"
SHEETS_ID = "1jNUCGulTY3mlpZkTTcCBlp-yq_7xYbqyJRjGBGPqwKY"
SHEET_NAME = "Hoshi"
OUTPUT_DIR = Path(__file__).parent / "output" / date.today().strftime("%Y-%m-%d")
FFMPEG_CMD = "ffmpeg"
YTDLP_CMD = "yt-dlp"

# ⚠️ ใส่ URL ของ Google Apps Script Web App ที่ deploy แล้วตรงนี้
# (ดูวิธี deploy ในไฟล์ google_apps_script.js)
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwZqYVUz92M65Yi8LaCvbR8w5lgdgYxciupT2FvE44K5vTlopddjrE5p5sgbt0tw6szAw/exec"
# ================================================


def fetch_sheet_data():
    """ดึงข้อมูลจาก Google Sheets API"""
    encoded_sheet = urllib.parse.quote(SHEET_NAME)
    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{SHEETS_ID}"
        f"/values/{encoded_sheet}?key={API_KEY}"
    )
    print(f"📊 กำลังดึงข้อมูลจาก Google Sheets (sheet: {SHEET_NAME})...")

    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode("utf-8"))
    except Exception as e:
        print(f"❌ ไม่สามารถดึงข้อมูลจาก Google Sheets ได้: {e}")
        sys.exit(1)

    rows = data.get("values", [])
    if len(rows) < 2:
        print("⚠️ ไม่พบข้อมูลคลิปใน Google Sheets")
        sys.exit(0)

    headers = rows[0]
    clips = []
    for i, row in enumerate(rows[1:], start=2):
        row_padded = row + [""] * (len(headers) - len(row))
        clip = dict(zip(headers, row_padded))
        clip["_row_number"] = i
        clips.append(clip)

    return clips


def update_sheet_cells(updates_list):
    """
    อัปเดตหลาย cells พร้อมกันผ่าน Google Apps Script Web API
    updates_list: [{"cell": "A2", "value": "success"}, {"cell": "H2", "value": "error msg"}, ...]
    """
    if not APPS_SCRIPT_URL:
        print("   ⚠️ ยังไม่ได้ตั้งค่า APPS_SCRIPT_URL — ข้ามการอัปเดต Sheet")
        return False

    payload = json.dumps({
        "sheet_name": SHEET_NAME,
        "updates": updates_list
    }).encode("utf-8")

    try:
        req = urllib.request.Request(
            APPS_SCRIPT_URL,
            data=payload,
            method="POST",
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req) as response:
            result = json.loads(response.read().decode("utf-8"))
            if result.get("success"):
                return True
            else:
                print(f"   ⚠️ Apps Script error: {result.get('error')}")
                return False
    except Exception as e:
        print(f"   ⚠️ ไม่สามารถอัปเดต Sheet: {e}")
        return False


def update_status_success(row_number):
    """อัปเดต status เป็น success ในคอลัมน์ A"""
    updates = [{"cell": f"A{row_number}", "value": "success"}]
    if update_sheet_cells(updates):
        print(f"   📝 อัปเดต Sheet: A{row_number} → success")


def update_status_fail(row_number, error_message):
    """อัปเดต status เป็น fail (A) และเขียน error (H)"""
    updates = [
        {"cell": f"A{row_number}", "value": "fail"},
        {"cell": f"H{row_number}", "value": error_message}
    ]
    if update_sheet_cells(updates):
        print(f"   📝 อัปเดต Sheet: A{row_number} → fail, H{row_number} → {error_message}")


def sanitize_filename(name):
    """ทำความสะอาดชื่อไฟล์ให้ใช้ได้บน Windows"""
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'[【】]', '_', name)
    name = re.sub(r'_+', '_', name)
    name = name.strip(' _.')
    if len(name) > 150:
        name = name[:150]
    return name


def time_to_seconds(time_str):
    """แปลงเวลาจาก H:MM:SS หรือ MM:SS เป็นวินาที"""
    parts = time_str.strip().split(":")
    parts = [int(p) for p in parts]
    if len(parts) == 3:
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    elif len(parts) == 2:
        return parts[0] * 60 + parts[1]
    else:
        return parts[0]


def is_youtube_url(path):
    """ตรวจสอบว่าเป็น YouTube URL หรือไม่"""
    return "youtube.com" in path or "youtu.be" in path


def clean_path(path_str):
    """ทำความสะอาด path string (ลบ quotes ออก)"""
    return path_str.strip().strip('"').strip("'")


def cut_local_file(input_path, output_path, start_time, end_time):
    """ตัดคลิปจากไฟล์ MP4 ในเครื่องด้วย FFmpeg"""
    if not os.path.exists(input_path):
        return False, f"ไม่พบไฟล์: {input_path}"

    cmd = [
        FFMPEG_CMD,
        "-y",
        "-ss", start_time,
        "-to", end_time,
        "-i", input_path,
        "-c", "copy",
        "-avoid_negative_ts", "make_zero",
        str(output_path)
    ]

    print(f"   🔧 FFmpeg: ตัดจาก {start_time} ถึง {end_time}")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300
        )
        if result.returncode != 0:
            return False, f"FFmpeg error: {result.stderr[-300:]}"
        return True, None
    except subprocess.TimeoutExpired:
        return False, "FFmpeg timeout (เกิน 5 นาที)"
    except FileNotFoundError:
        return False, "ไม่พบคำสั่ง ffmpeg"


def download_and_cut_youtube(url, output_path, start_time, end_time):
    """ดาวน์โหลดจาก YouTube แล้วตัดคลิป"""
    print(f"   📥 กำลังดาวน์โหลดจาก YouTube...")
    cmd_download = [
        YTDLP_CMD,
        "--download-sections", f"*{start_time}-{end_time}",
        "-f", "bv*[height<=1080]+ba/b[height<=1080]/bv*+ba/b",
        "--merge-output-format", "mp4",
        "-o", str(output_path),
        "--no-playlist",
        "--force-keyframes-at-cuts",
        url
    ]

    try:
        result = subprocess.run(
            cmd_download,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=600
        )
        if result.returncode != 0:
            return False, f"yt-dlp error: {result.stderr[-300:]}"
        return True, None
    except subprocess.TimeoutExpired:
        return False, "yt-dlp timeout (เกิน 10 นาที)"
    except FileNotFoundError:
        return False, "ไม่พบคำสั่ง yt-dlp — กรุณาติดตั้ง: pip install yt-dlp"


def main():
    print("=" * 60)
    print("🎬 Video Clip Cutter — ตัดคลิปจาก Google Sheets")
    print("=" * 60)

    if not APPS_SCRIPT_URL:
        print("⚠️  APPS_SCRIPT_URL ยังไม่ได้ตั้งค่า — จะไม่อัปเดต status ใน Sheet")
        print("   (ดูวิธี deploy ในไฟล์ google_apps_script.js)\n")

    # Create output directory
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"📁 Output directory: {OUTPUT_DIR}")

    # Fetch data from Google Sheets
    clips = fetch_sheet_data()

    # Filter only "wait" status
    wait_clips = [c for c in clips if c.get("status", "").strip().lower() == "wait"]
    print(f"\n📋 พบ {len(wait_clips)} คลิปที่ต้องตัด (status = wait)\n")

    if not wait_clips:
        print("✅ ไม่มีคลิปที่ต้องตัด")
        return

    results = {"success": [], "failed": []}

    for i, clip in enumerate(wait_clips, start=1):
        clip_name = clip.get("ชื่อคลิป", f"clip_{i}").strip()
        raw_path = clip.get("Path", "").strip()             # Column G
        start_time = clip.get("ช่วงเริ่มต้น", "").strip()     # Column E
        end_time = clip.get("ช่วงสิ้นสุด", "").strip()        # Column F
        row_num = clip["_row_number"]

        print(f"{'─' * 60}")
        print(f"🎬 [{i}/{len(wait_clips)}] {clip_name}")
        print(f"   📂 Source: {raw_path}")
        print(f"   ⏱️  ช่วงเวลา: {start_time} → {end_time}")

        # --- Validate: no path ---
        if not raw_path:
            error_msg = "ไม่มี Path ระบุ (คอลัมน์ G ว่าง)"
            print(f"   ❌ {error_msg}")
            update_status_fail(row_num, error_msg)
            results["failed"].append(clip_name)
            continue

        # --- Validate: no start/end time ---
        if not start_time or not end_time:
            error_msg = "ไม่มีช่วงเวลาเริ่มต้น/สิ้นสุด"
            print(f"   ❌ {error_msg}")
            update_status_fail(row_num, error_msg)
            results["failed"].append(clip_name)
            continue

        safe_name = sanitize_filename(clip_name)
        output_path = OUTPUT_DIR / f"{safe_name}.mp4"

        if is_youtube_url(raw_path):
            success, error_msg = download_and_cut_youtube(raw_path, output_path, start_time, end_time)
        else:
            input_path = clean_path(raw_path)
            # --- Validate: file not found ---
            if not os.path.exists(input_path):
                error_msg = f"ไม่พบไฟล์: {input_path}"
                print(f"   ❌ {error_msg}")
                update_status_fail(row_num, error_msg)
                results["failed"].append(clip_name)
                continue
            success, error_msg = cut_local_file(input_path, output_path, start_time, end_time)

        if success and output_path.exists():
            size_mb = output_path.stat().st_size / (1024 * 1024)
            print(f"   ✅ สำเร็จ! → {output_path.name} ({size_mb:.1f} MB)")
            update_status_success(row_num)
            results["success"].append(clip_name)
        else:
            print(f"   ❌ ล้มเหลว!")
            update_status_fail(row_num, error_msg or "ตัดคลิปล้มเหลว (ไม่ทราบสาเหตุ)")
            results["failed"].append(clip_name)

    # Summary
    print(f"\n{'=' * 60}")
    print(f"📊 สรุปผลการตัดคลิป")
    print(f"{'=' * 60}")
    print(f"   ✅ สำเร็จ: {len(results['success'])} คลิป")
    for name in results["success"]:
        print(f"      • {name}")
    if results["failed"]:
        print(f"   ❌ ล้มเหลว: {len(results['failed'])} คลิป")
        for name in results["failed"]:
            print(f"      • {name}")
    print(f"\n📁 ไฟล์คลิปอยู่ที่: {OUTPUT_DIR}")
    print("🎉 เสร็จสิ้น!")


if __name__ == "__main__":
    main()
