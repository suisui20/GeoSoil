# app.py - Flask 主程序

import uuid
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from city_data import CITIES
import geo_utils

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

UPLOADS_DIR = Path(__file__).parent / "static" / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    cities = list(CITIES.keys())
    return render_template("index.html", cities=cities)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json()
    city = data.get("city", "").strip()
    try:
        count = int(data.get("count", 1000))
    except (ValueError, TypeError):
        return jsonify({"error": "数量必须为整数"}), 400

    if city not in CITIES:
        return jsonify({"error": f"不支持的城市：{city}"}), 400
    if not 10 <= count <= 20000:
        return jsonify({"error": "数量范围：10 ~ 20000"}), 400

    try:
        excel_path, summary = geo_utils.generate_land_points(city, count)
        filename = Path(excel_path).name
        return jsonify({
            "filename": filename,
            "summary": summary,
            "total": sum(s["count"] for s in summary),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/map", methods=["POST"])
def api_map():
    # 支持两种输入：上传文件 OR 已有文件名
    if "file" in request.files and request.files["file"].filename:
        f = request.files["file"]
        if not f.filename.lower().endswith(".xlsx"):
            return jsonify({"error": "仅支持 .xlsx 文件"}), 400
        fname = f"{uuid.uuid4().hex[:8]}_{f.filename}"
        save_path = UPLOADS_DIR / fname
        f.save(save_path)
        excel_path = str(save_path)
    else:
        filename = request.form.get("filename") or (
            request.get_json(silent=True) or {}
        ).get("filename")
        if not filename:
            return jsonify({"error": "请上传文件或提供已生成的文件名"}), 400
        excel_path = str(UPLOADS_DIR / filename)
        if not Path(excel_path).exists():
            return jsonify({"error": "文件不存在，请重新生成"}), 404

    try:
        image_rel = geo_utils.generate_map(excel_path)
        return jsonify({"image_url": f"/static/{image_rel}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/<filename>")
def download(filename):
    path = UPLOADS_DIR / filename
    if not path.exists():
        return "文件不存在", 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
