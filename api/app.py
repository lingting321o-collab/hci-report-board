from flask import Flask, Response, jsonify, request

from api.convert import INDEX_HTML, build_docx, parse_report

app = Flask(__name__)


@app.get("/")
def home():
    return Response(INDEX_HTML, content_type="text/html; charset=utf-8")


@app.post("/api/convert")
def convert_report():
    uploaded = request.files.get("file")
    if uploaded is None:
        return jsonify({"error": "没有找到上传文件"}), 400

    data = uploaded.read()
    if not data:
        return jsonify({"error": "上传文件为空"}), 400

    try:
        output = build_docx(parse_report(data))
    except Exception as exc:
        return jsonify({"error": f"生成失败：{exc}"}), 400

    return Response(
        output,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="HCI-remediation-board.docx"'},
    )


@app.route("/api/convert", methods=["OPTIONS"])
def convert_options():
    return Response(status=204, headers={
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
    })
