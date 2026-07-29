import os
import uuid
import threading
import time
import re
import sqlite3

from flask import Flask, request, jsonify, send_file, render_template
import openpyxl
from jusho import Jusho
import posuto

# posuto は日本郵便KEN_ALLの最新データを同梱しており、
# jusho（同梱DBが古い）で見つからない町名のフォールバックに使う
POSUTO_DB = os.path.join(os.path.dirname(posuto.__file__), "postaldata.db")

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


WARD_REMAP = {
    "浜松市中央区": ["浜松市中区", "浜松市東区", "浜松市南区"],
    "浜松市浜名区": ["浜松市西区", "浜松市北区", "浜松市浜北区"],
}

# 住所表記で相互に揺れる異体字・仮名
KANJI_ADDR_VARIANTS = {
    "州": "洲", "洲": "州", "磯": "礒", "礒": "磯",
    "槙": "槇", "槇": "槙", "諌": "諫", "諫": "諌",
    "繩": "縄", "縄": "繩", "鴎": "鷗", "鷗": "鴎",
    "一": "壱", "壱": "一",
}

KE_GROUP = ["\u30f6", "\u30b1", "\u304c", "\u30ac"]  # ヶ ケ が ガ


def _char_variants(s):
    variants = [s]

    # ヶ/ケ/が/ガ は地名で相互に揺れる
    for i, ch in enumerate(s):
        if ch in KE_GROUP:
            expanded = []
            for v in variants:
                for repl in KE_GROUP:
                    expanded.append(v[:i] + repl + v[i + 1:])
            variants = expanded

    # ノ/の の揺れ、および省略
    no_expanded = []
    for v in variants:
        no_expanded.append(v)
        if "\u30ce" in v or "\u306e" in v:
            no_expanded.append(v.replace("\u30ce", "\u306e"))
            no_expanded.append(v.replace("\u306e", "\u30ce"))
            no_expanded.append(v.replace("\u30ce", "").replace("\u306e", ""))
    variants = no_expanded

    # 異体字
    kanji_expanded = []
    for v in variants:
        kanji_expanded.append(v)
        for old, new in KANJI_ADDR_VARIANTS.items():
            if old in v:
                kanji_expanded.append(v.replace(old, new))
    variants = kanji_expanded

    seen = set()
    result = []
    for v in variants:
        if v and v not in seen:
            seen.add(v)
            result.append(v)
    return result


def _digit_variants(s):
    half = "0123456789"
    full = "\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19"
    to_full = s.translate(str.maketrans(half, full))
    to_half = s.translate(str.maketrans(full, half))
    out = [s]
    for v in (to_full, to_half):
        if v not in out:
            out.append(v)
    return out


def _build_city_variants(city):
    variants = [city]
    for new_ward, old_wards in WARD_REMAP.items():
        if city == new_ward:
            variants.extend(old_wards)
            break
    # 郡下の町村が市に昇格したケース（例: 岩手郡滝沢村 → 滝沢市）
    if "\u90e1" in city:
        gm = re.search(r"\u90e1(.+?)[\u753a\u6751]$", city)
        if gm:
            core = gm.group(1)
            variants.extend([core + "\u5e02", core + "\u753a", core + "\u6751"])
    # ヶ/ケ 等の揺れ
    expanded = []
    for v in variants:
        expanded.extend(_char_variants(v))
    seen = set()
    out = []
    for v in expanded:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out


def _search_with_city_match(town_clean, pref, city, jusho_db):
    results = jusho_db.search_addresses(town_clean)
    city_variants = _build_city_variants(city)

    exact_match = None
    sonota_match = None
    subarea_match = None
    best_match = None
    for r in results:
        addr_str = str(r)
        if pref not in addr_str:
            continue
        city_matched = False
        for cv in city_variants:
            if cv in addr_str:
                city_matched = True
                break
        if not city_matched:
            city_parts = re.findall(r"[^市区町村郡]+[市区町村]", city)
            if city_parts and all(part in addr_str for part in city_parts):
                city_matched = True
        if not city_matched:
            continue
        zm = re.search(r"〒(\d{3}-\d{4})", addr_str)
        if not zm:
            continue
        zipcode = zm.group(1).replace("-", "")
        # 町名の直後がローマ字開き括弧 "(" = 小字なしの完全一致
        # 町名の直後が末尾 = 完全一致
        if f" {town_clean}(" in addr_str or addr_str.endswith(f" {town_clean}"):
            if exact_match is None:
                exact_match = zipcode
        # 「（その他）」= 小字を特定できない場合の代表郵便番号
        elif f" {town_clean}\uff08\u305d\u306e\u4ed6\uff09" in addr_str:
            if sonota_match is None:
                sonota_match = zipcode
        # 町名の直後が全角括弧 = 小字・丁目範囲付き（その他が無い場合の候補）
        elif f" {town_clean}\uff08" in addr_str:
            if subarea_match is None:
                subarea_match = zipcode
        if best_match is None:
            best_match = zipcode

    return exact_match or sonota_match or subarea_match or best_match


def _posuto_city_match(town_clean, pref, city, conn):
    rows = conn.execute(
        "SELECT code, city, neighborhood, data FROM postal_data "
        "WHERE prefecture = ? AND neighborhood LIKE ?",
        (pref, "%" + town_clean + "%"),
    ).fetchall()
    city_variants = _build_city_variants(city)

    exact_full = None    # 町名完全一致かつ非分割（その町域＝単一郵便番号の確定値）
    sonota_match = None   # koazabanchi=true =「（その他）」相当の代表郵便番号
    exact_partial = None  # 町名完全一致だが小字/丁目で分割
    subarea_match = None
    best_match = None
    for code, rcity, neighborhood, data in rows:
        city_matched = any(cv in rcity for cv in city_variants)
        if not city_matched:
            city_parts = re.findall(r"[^市区町村郡]+[市区町村]", city)
            if city_parts and all(part in rcity for part in city_parts):
                city_matched = True
        if not city_matched:
            continue
        if neighborhood == town_clean:
            if '"koazabanchi": true' in data:
                if sonota_match is None:
                    sonota_match = code
            elif '"partial": false' in data:
                if exact_full is None:
                    exact_full = code
            elif exact_partial is None:
                exact_partial = code
        elif neighborhood.startswith(town_clean):
            if subarea_match is None:
                subarea_match = code
        if best_match is None:
            best_match = code

    return exact_full or sonota_match or exact_partial or subarea_match or best_match


def _find_zipcode(address, search_fn):
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

    NUM_CHARS = r"[\d０-９]"
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
            # 市名に「市」が二つ含まれる誤分割を補正（例: 四日市市/野々市市 → town="市桜町"）
            town = re.sub(r"^市(?=.)", "", town)
            town_clean = re.sub(r"[０-９0-9一二三四五六七八九十丁目番地号の\-－ー・]+$", "", town).strip()
            town_clean = re.sub(r"^(?:大字|字)", "", town_clean)
            if not town_clean and not town:
                continue

            base_towns = []
            if town_clean:
                base_towns.append(town_clean)
            # 元の町名（末尾の丁目等を除去する前）も試す（例: 六番丁）
            raw_town = re.sub(r"^(?:大字|字)", "", town)
            if raw_town and raw_town not in base_towns:
                base_towns.append(raw_town)
            # 市名が町名に重複して残るケースを除去（例: 富谷市富谷町成田 → 成田）
            city_core = re.sub(r"^.+郡", "", city)
            city_core = re.sub(r"[市区町村]$", "", city_core)
            if city_core and town_clean.startswith(city_core):
                deduped = re.sub(r"^[市区町村]", "", town_clean[len(city_core):])
                if deduped and deduped != town_clean:
                    base_towns.append(deduped)
            if not town_clean:
                town_clean = raw_town
            # 「字」「大字」の処理: 除去 / 分割（例: 脇町大字脇町 → 脇町脇町、滝沢字牧野林 → 牧野林）
            if "字" in town_clean:
                base_towns.append(town_clean.replace("大字", "").replace("字", ""))
                aza_split = [p for p in re.split(r"大字|字", town_clean) if p]
                base_towns.extend(aza_split)
            # 京都の通り名住所（例: 壬生通八条下ル東寺町 → 東寺町）
            kyoto_split = re.split(r"(?:上る|下る|上ル|下ル|東入る|西入る|東入ル|西入ル|東入|西入)", town_clean)
            if len(kyoto_split) > 1 and kyoto_split[-1]:
                base_towns.append(kyoto_split[-1])
            # 「町」境界での区切り（例: 明大寺町伝馬 → 明大寺町）
            machi_match = re.match(r"^(.+?町)", town_clean)
            if machi_match and machi_match.group(1) != town_clean:
                base_towns.append(machi_match.group(1))
            # 末尾「町」の除去（例: 佐藤町 → 佐藤）
            if town_clean.endswith("町") and len(town_clean) > 2:
                base_towns.append(town_clean[:-1])
            # 複合地名のフォールバック: 末尾から1文字ずつ短縮（例: 西今宿阿弥陀寺 → 西今宿）
            for cut in range(len(town_clean) - 1, 1, -1):
                base_towns.append(town_clean[:cut])
            # 最終手段: 先頭の余分な1〜2文字を除去（例: 元データ誤記「立柏の森」→「柏の森」）
            for drop in (1, 2):
                if len(town_clean) - drop >= 2:
                    base_towns.append(town_clean[drop:])

            seen = set()
            for base in base_towns:
                for tv in _char_variants(base):
                    for dv in _digit_variants(tv):
                        if dv in seen or not dv:
                            continue
                        seen.add(dv)
                        result = search_fn(dv, pref, city)
                        if result:
                            return result
    return ""


def address_to_zipcode(address, jusho_db, posuto_conn=None):
    # 常に最新の日本郵便KEN_ALL（posuto）を主データ源として優先
    if posuto_conn is not None:
        result = _find_zipcode(
            address, lambda dv, pref, city: _posuto_city_match(dv, pref, city, posuto_conn)
        )
        if result:
            return result
    # posutoで見つからない場合のみ jusho（旧DB）にフォールバック
    return _find_zipcode(
        address, lambda dv, pref, city: _search_with_city_match(dv, pref, city, jusho_db)
    )


def process_excel(job_id, input_path, output_path):
    try:
        jobs[job_id]["status"] = "processing"
        # Create per-thread jusho instance to avoid SQLite threading issues
        thread_jusho = Jusho()
        thread_posuto = sqlite3.connect(POSUTO_DB, check_same_thread=False)

        wb = openpyxl.load_workbook(input_path)
        ws = wb.active

        ws["B1"] = "高校名"
        ws["C1"] = "郵便番号（元データ）"
        ws["D1"] = "郵便番号（住所から逆引き）"
        ws["E1"] = "住所（元データ）"

        total = ws.max_row - 1
        jobs[job_id]["total"] = total

        for row_idx in range(2, ws.max_row + 1):
            school_name = ws.cell(row=row_idx, column=2).value
            zipcode = ws.cell(row=row_idx, column=3).value
            original_address = ws.cell(row=row_idx, column=4).value

            reverse_zipcode = address_to_zipcode(original_address, thread_jusho, thread_posuto)

            ws.cell(row=row_idx, column=4).value = reverse_zipcode
            ws.cell(row=row_idx, column=5).value = original_address

            jobs[job_id]["progress"] = row_idx - 1
            jobs[job_id]["current_school"] = school_name or ""

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
    # jobs はプロセス内メモリのため、ワーカー再起動などで失われることがある。
    # その場合でも結果ファイルがディスクに残っていれば配信できるようにする。
    default_path = os.path.join(
        app.config["RESULT_FOLDER"], f"{job_id}_output.xlsx"
    )
    job = jobs.get(job_id)

    if job is not None:
        if job["status"] != "done":
            return jsonify({"error": "処理がまだ完了していません"}), 400
        output_path = job["output_path"]
        original = job.get("original_filename", "output.xlsx")
    else:
        if not os.path.exists(default_path):
            return jsonify({
                "error": "時間が経過したため処理結果が失われました。"
                         "お手数ですが、もう一度アップロードして処理してください。"
            }), 404
        output_path = default_path
        original = "output.xlsx"

    name_base = os.path.splitext(original)[0]
    download_name = f"{name_base}_processed.xlsx"

    return send_file(output_path, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
