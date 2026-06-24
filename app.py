import os
import uuid
import threading
import time
import re

from flask import Flask, request, jsonify, send_file, render_template
import openpyxl
import requests
from jusho import Jusho

app = Flask(__name__)
_base_dir = os.environ.get("RENDER", None)
_data_root = "/tmp" if _base_dir is not None else os.path.dirname(__file__)
app.config["UPLOAD_FOLDER"] = os.path.join(_data_root, "uploads")
app.config["RESULT_FOLDER"] = os.path.join(_data_root, "results")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["RESULT_FOLDER"], exist_ok=True)

# Job tracking
jobs = {}

# jusho DB will be created per-thread to avoid SQLite threading issues

SCHOOL_API_URL = "https://school.teraren.com/schools.json"

KANJI_NORMALIZE = {
    "鷗": "鴎", "國": "国", "學": "学", "藝": "芸", "櫻": "桜",
    "澤": "沢", "邊": "辺", "齋": "斎", "齊": "斉", "廣": "広",
    "髙": "高", "﨑": "崎",
}


def normalize_school_name(name):
    if not name:
        return name
    result = name
    result = result.replace("\uFF0D", "\u30FC")
    result = result.replace("－", "ー")
    for old, new in KANJI_NORMALIZE.items():
        result = result.replace(old, new)
    return result


def address_to_zipcode(address, jusho_db):
    if not address:
        return ""

    addr = address
    if not re.match(r"^(東京都|北海道|(?:京都|大阪)府|.{2,3}県)", addr):
        city_pref_map = {
            "仙台市": "宮城県", "札幌市": "北海道", "横浜市": "神奈川県",
            "名古屋市": "愛知県", "大阪市": "大阪府", "京都市": "京都府",
            "神戸市": "兵庫県", "福岡市": "福岡県", "広島市": "広島県",
            "さいたま市": "埼玉県", "千葉市": "千葉県", "川崎市": "神奈川県",
            "北九州市": "福岡県", "堺市": "大阪府", "浜松市": "静岡県",
            "新潟市": "新潟県", "熊本市": "熊本県", "岡山市": "岡山県",
            "静岡市": "静岡県", "相模原市": "神奈川県",
        }
        for city, pref in city_pref_map.items():
            if addr.startswith(city):
                addr = pref + addr
                break

    NUM_CHARS = r"[\d０-９一二三四五六七八九十]"
    patterns = [
        rf"^(東京都|北海道|(?:京都|大阪)府|.{{2,3}}県)(.+?市.+?区)(.+?)(?:{NUM_CHARS}|$)",
        rf"^(東京都)(.+?区)(.+?)(?:{NUM_CHARS}|$)",
        rf"^(東京都|北海道|(?:京都|大阪)府|.{{2,3}}県)(.+?市)(.+?)(?:{NUM_CHARS}|$)",
        rf"^(東京都|北海道|(?:京都|大阪)府|.{{2,3}}県)(.+?郡.+?(?:町|村))(.+?)(?:{NUM_CHARS}|$)",
        rf"^(東京都|北海道|(?:京都|大阪)府|.{{2,3}}県)(.+?(?:町|村))(.+?)(?:{NUM_CHARS}|$)",
    ]

    for pat in patterns:
        m = re.match(pat, addr)
        if m:
            pref, city, town = m.group(1), m.group(2), m.group(3)
            town_clean = re.sub(r"[０-９0-9一二三四五六七八九十丁目番地号の\-－ー・]+$", "", town).strip()
            town_clean = re.sub(r"^(?:大字|字)", "", town_clean)
            if not town_clean:
                continue

            results = jusho_db.search_addresses(town_clean)
            best_match = None
            for r in results:
                addr_str = str(r)
                if pref not in addr_str:
                    continue
                city_matched = False
                if city in addr_str:
                    city_matched = True
                else:
                    city_parts = re.findall(r"[^市区町村郡]+[市区町村]", city)
                    if city_parts and all(part in addr_str for part in city_parts):
                        city_matched = True
                if not city_matched:
                    continue
                zm = re.search(r"〒(\d{3}-\d{4})", addr_str)
                if zm:
                    zipcode = zm.group(1).replace("-", "")
                    if town_clean + "(" in addr_str or f" {town_clean}" in addr_str:
                        return zipcode
                    if best_match is None:
                        best_match = zipcode
            if best_match:
                return best_match
    return ""


def get_school_address(school_name, original_zipcode=None, original_address=None):
    if not school_name:
        return ""

    normalized = normalize_school_name(school_name)
    search_terms = []

    if school_name != normalized:
        search_terms.append(school_name)

    search_terms.append(normalized)

    if normalized != school_name:
        half_normalized = school_name.replace("\uFF0D", "\u30FC").replace("－", "ー")
        if half_normalized != normalized and half_normalized != school_name:
            search_terms.append(half_normalized)

    if " " in normalized or "\u3000" in normalized:
        parts = re.split(r"[ \u3000]+", normalized)
        search_terms.append(parts[-1])

    short = re.sub(r"(高等学校|高等部|中学校高等学校|中等教育学校)$", "", normalized)
    if short != normalized:
        search_terms.append(short)

    no_suffix = re.sub(r"[・].*$", "", normalized)
    if no_suffix != normalized and len(no_suffix) >= 3:
        search_terms.append(no_suffix)

    no_campus = re.sub(r"(?:校舎|キャンパス)$", "", normalized)
    if no_campus != normalized:
        search_terms.append(no_campus)

    original_pref = ""
    if original_address:
        pm = re.match(r"^(東京都|北海道|(?:京都|大阪)府|.{2,3}県)", original_address)
        if pm:
            original_pref = pm.group(1)

    for term in search_terms:
        try:
            r = requests.get(SCHOOL_API_URL, params={"s": term}, timeout=10)
            results = r.json()
            if not results:
                continue
            if original_zipcode is not None:
                zip_str = str(int(original_zipcode)).zfill(7)
                for s in results:
                    if s.get("postal_code") == zip_str:
                        return s.get("location", "")
            if original_pref:
                for s in results:
                    if s.get("location", "").startswith(original_pref):
                        return s.get("location", "")
            return results[0].get("location", "")
        except Exception:
            continue
    return ""


def process_excel(job_id, input_path, output_path):
    try:
        jobs[job_id]["status"] = "processing"
        # Create per-thread jusho instance to avoid SQLite threading issues
        thread_jusho = Jusho()

        wb = openpyxl.load_workbook(input_path)
        ws = wb.active

        ws["B1"] = "高校名"
        ws["C1"] = "郵便番号（元データ）"
        ws["D1"] = "郵便番号（住所から逆引き）"
        ws["E1"] = "住所（元データ）"
        ws["F1"] = "住所（学校名APIから取得）"

        total = ws.max_row - 1
        jobs[job_id]["total"] = total

        for row_idx in range(2, ws.max_row + 1):
            school_name = ws.cell(row=row_idx, column=2).value
            zipcode = ws.cell(row=row_idx, column=3).value
            original_address = ws.cell(row=row_idx, column=4).value

            reverse_zipcode = address_to_zipcode(original_address, thread_jusho)
            school_address = get_school_address(school_name, zipcode, original_address)

            ws.cell(row=row_idx, column=4).value = reverse_zipcode
            ws.cell(row=row_idx, column=5).value = original_address
            ws.cell(row=row_idx, column=6).value = school_address

            jobs[job_id]["progress"] = row_idx - 1
            jobs[job_id]["current_school"] = school_name or ""
            time.sleep(0.3)

        wb.save(output_path)
        jobs[job_id]["status"] = "done"
        jobs[job_id]["progress"] = total

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.endswith(".xlsx"):
        return jsonify({"error": ".xlsx ファイルを選択してください"}), 400

    job_id = str(uuid.uuid4())[:8]
    input_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{job_id}_input.xlsx")
    output_path = os.path.join(app.config["RESULT_FOLDER"], f"{job_id}_output.xlsx")
    f.save(input_path)

    jobs[job_id] = {
        "status": "queued",
        "progress": 0,
        "total": 0,
        "current_school": "",
        "output_path": output_path,
        "original_filename": f.filename,
    }

    thread = threading.Thread(target=process_excel, args=(job_id, input_path, output_path))
    thread.daemon = True
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "ジョブが見つかりません"}), 404
    job = jobs[job_id]
    return jsonify({
        "status": job["status"],
        "progress": job["progress"],
        "total": job["total"],
        "current_school": job.get("current_school", ""),
        "error": job.get("error", ""),
    })


@app.route("/download/<job_id>")
def download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "ジョブが見つかりません"}), 404
    job = jobs[job_id]
    if job["status"] != "done":
        return jsonify({"error": "処理がまだ完了していません"}), 400

    original = job.get("original_filename", "output.xlsx")
    name_base = os.path.splitext(original)[0]
    download_name = f"{name_base}_processed.xlsx"

    return send_file(job["output_path"], as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
