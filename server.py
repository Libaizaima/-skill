# -*- coding: utf-8 -*-
"""
流水分析系统 — Web 服务器
Flask + SQLite + SSE 实时日志 + 多文件下载
"""

import io
import os
import sys
import json
import uuid
import queue
import sqlite3
import zipfile
import threading
import subprocess
from datetime import datetime
from functools import wraps

from flask import (
    Flask, request, session, redirect, url_for,
    render_template, jsonify, send_file, Response, abort
)

# ──────────────────────────────────────────────
# 配置
# ──────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_CFG_PATH = os.path.join(BASE_DIR, "web_config.json")
DB_PATH = os.path.join(BASE_DIR, "web", "db", "jobs.db")
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

with open(WEB_CFG_PATH, "r", encoding="utf-8") as f:
    web_cfg = json.load(f)

app = Flask(__name__, template_folder="web/templates", static_folder="web/static")
app.secret_key = web_cfg.get("secret_key", "dev-secret")

# 全局 SSE 队列: task_id -> queue.Queue
_sse_queues: dict = {}
_sse_lock = threading.Lock()


# ──────────────────────────────────────────────
# 数据库
# ──────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                company_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                docx_path TEXT,
                output_dir TEXT,
                log TEXT DEFAULT ''
            )
        """)
        # 给旧表加 output_dir 列（兼容升级）
        try:
            db.execute("ALTER TABLE jobs ADD COLUMN output_dir TEXT")
        except Exception:
            pass
        db.commit()


init_db()


# ──────────────────────────────────────────────
# 认证
# ──────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        u = request.form.get("username", "")
        p = request.form.get("password", "")
        if u == web_cfg["username"] and p == web_cfg["password"]:
            session["logged_in"] = True
            return redirect(url_for("index"))
        error = "用户名或密码错误"
    return render_template("login.html", error=error)


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


# ──────────────────────────────────────────────
# 主页
# ──────────────────────────────────────────────
@app.route("/")
@login_required
def index():
    return render_template("dashboard.html")


# ──────────────────────────────────────────────
# 上传 & 触发分析
# ──────────────────────────────────────────────
@app.route("/upload", methods=["POST"])
@login_required
def upload():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".zip"):
        return jsonify({"error": "请上传 ZIP 文件"}), 400

    task_id = str(uuid.uuid4())[:8]
    company_name = os.path.splitext(f.filename)[0]
    zip_path = os.path.join(INPUT_DIR, f"{task_id}_{f.filename}")
    f.save(zip_path)

    with get_db() as db:
        db.execute(
            "INSERT INTO jobs (id, company_name, created_at, status) VALUES (?,?,?,?)",
            (task_id, company_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "running")
        )
        db.commit()

    q = queue.Queue()
    with _sse_lock:
        _sse_queues[task_id] = q

    threading.Thread(
        target=_run_analysis,
        args=(task_id, zip_path, q),
        daemon=True
    ).start()

    return jsonify({"task_id": task_id, "company_name": company_name})


def _run_analysis(task_id: str, zip_path: str, q: queue.Queue):
    """后台执行分析，日志写入 SSE 队列和 DB"""
    log_lines = []
    docx_path = None
    output_dir = None
    status = "error"

    try:
        proc = subprocess.Popen(
            [sys.executable, os.path.join(BASE_DIR, "src", "main.py"), zip_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            cwd=BASE_DIR
        )

        for raw in proc.stdout:
            line = raw.rstrip()
            log_lines.append(line)
            q.put(("log", line))

        proc.wait()

        if proc.returncode == 0:
            status = "done"
            # 找最新修改的 DOCX 和其所在目录
            newest_mtime = 0
            for root, dirs, files in os.walk(OUTPUT_DIR):
                for fn in files:
                    if fn.endswith(".docx"):
                        fp = os.path.join(root, fn)
                        mt = os.path.getmtime(fp)
                        if mt > newest_mtime:
                            newest_mtime = mt
                            docx_path = fp
                            output_dir = root  # 报告所在目录即公司输出目录
            q.put(("done", docx_path or ""))
        else:
            q.put(("error", f"分析进程退出码: {proc.returncode}"))

    except Exception as e:
        q.put(("error", str(e)))

    full_log = "\n".join(log_lines)
    with get_db() as db:
        db.execute(
            "UPDATE jobs SET status=?, docx_path=?, output_dir=?, log=? WHERE id=?",
            (status, docx_path, output_dir, full_log, task_id)
        )
        db.commit()

    with _sse_lock:
        _sse_queues.pop(task_id, None)


# ──────────────────────────────────────────────
# SSE 实时日志
# ──────────────────────────────────────────────
@app.route("/progress/<task_id>")
@login_required
def progress(task_id):
    def generate():
        with get_db() as db:
            row = db.execute("SELECT log, status FROM jobs WHERE id=?", (task_id,)).fetchone()
        if row and row["log"]:
            for line in row["log"].split("\n"):
                yield f"data: {json.dumps({'type':'log','text':line})}\n\n"
        if row and row["status"] in ("done", "error"):
            yield f"data: {json.dumps({'type': row['status'], 'text':''})}\n\n"
            return

        with _sse_lock:
            q = _sse_queues.get(task_id)
        if not q:
            yield f"data: {json.dumps({'type':'error','text':'任务不存在或已结束'})}\n\n"
            return

        while True:
            try:
                msg_type, text = q.get(timeout=30)
                yield f"data: {json.dumps({'type': msg_type, 'text': text})}\n\n"
                if msg_type in ("done", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type':'ping','text':''})}\n\n"

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ──────────────────────────────────────────────
# 历史记录
# ──────────────────────────────────────────────
@app.route("/history")
@login_required
def history():
    with get_db() as db:
        rows = db.execute(
            "SELECT id, company_name, created_at, status, docx_path, output_dir FROM jobs ORDER BY created_at DESC"
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/delete/<task_id>", methods=["DELETE"])
@login_required
def delete_job(task_id):
    with get_db() as db:
        db.execute("DELETE FROM jobs WHERE id=?", (task_id,))
        db.commit()
    return jsonify({"ok": True})


@app.route("/clear-all", methods=["DELETE"])
@login_required
def clear_all():
    with get_db() as db:
        db.execute("DELETE FROM jobs")
        db.commit()
    return jsonify({"ok": True})


# ──────────────────────────────────────────────
# 文件列表（列出输出目录下可下载的文件）
# ──────────────────────────────────────────────
# 可下载文件扩展名白名单
_ALLOWED_EXTS = {".docx", ".csv", ".json", ".xlsx", ".xls", ".pdf"}


def _find_output_dir_by_name(company_name: str) -> str:
    """根据公司名在 output/ 目录中匹配输出文件夹"""
    if not os.path.isdir(OUTPUT_DIR):
        return ""
    # 优先精确匹配，其次模糊匹配（company_name 是文件名，目录名可能含前缀）
    candidates = []
    for d in os.listdir(OUTPUT_DIR):
        dpath = os.path.join(OUTPUT_DIR, d)
        if not os.path.isdir(dpath):
            continue
        if d == company_name:
            return dpath  # 精确匹配
        if company_name in d or d in company_name:
            candidates.append(dpath)
    # 返回最近修改的候选
    if candidates:
        return max(candidates, key=os.path.getmtime)
    return ""


def _list_output_files(job_row) -> list:
    """列出某个任务的可下载文件"""
    files = []
    out_dir = job_row.get("output_dir") if isinstance(job_row, dict) else job_row["output_dir"]

    if not out_dir or not os.path.isdir(out_dir):
        # 降级1：从 docx_path 推导
        docx = job_row.get("docx_path") if isinstance(job_row, dict) else job_row["docx_path"]
        if docx and os.path.exists(docx):
            out_dir = os.path.dirname(docx)
        else:
            # 降级2：按公司名匹配 output/ 子目录（旧历史记录兼容）
            company = job_row.get("company_name") if isinstance(job_row, dict) else job_row["company_name"]
            out_dir = _find_output_dir_by_name(company or "")
        if not out_dir or not os.path.isdir(out_dir):
            return files

    # DOCX 报告（在 out_dir 里）
    for fn in os.listdir(out_dir):
        ext = os.path.splitext(fn)[1].lower()
        if ext in _ALLOWED_EXTS:
            files.append({"name": fn, "path": ".", "size": os.path.getsize(os.path.join(out_dir, fn))})

    # 02_解析结果 子目录
    parsed_dir = os.path.join(out_dir, "02_解析结果")
    if os.path.isdir(parsed_dir):
        for fn in sorted(os.listdir(parsed_dir)):
            ext = os.path.splitext(fn)[1].lower()
            if ext in _ALLOWED_EXTS:
                files.append({
                    "name": fn,
                    "path": "02_解析结果",
                    "size": os.path.getsize(os.path.join(parsed_dir, fn))
                })

    return files


@app.route("/files/<task_id>")
@login_required
def list_files(task_id):
    with get_db() as db:
        row = db.execute(
            "SELECT output_dir, docx_path, company_name FROM jobs WHERE id=?", (task_id,)
        ).fetchone()
    if not row:
        abort(404)
    return jsonify(_list_output_files(row))


# ──────────────────────────────────────────────
# 下载单个文件
# ──────────────────────────────────────────────
@app.route("/download/<task_id>")
@login_required
def download(task_id):
    """下载主 DOCX 报告（兼容旧接口）"""
    with get_db() as db:
        row = db.execute("SELECT docx_path, company_name FROM jobs WHERE id=?", (task_id,)).fetchone()
    if not row or not row["docx_path"] or not os.path.exists(row["docx_path"]):
        abort(404)
    return send_file(row["docx_path"], as_attachment=True,
                     download_name=f"客户分析_{row['company_name']}.docx")


@app.route("/download-file/<task_id>")
@login_required
def download_file(task_id):
    """下载任意单个输出文件，?path=02_解析结果&name=xxx.csv"""
    sub = request.args.get("path", ".")
    name = request.args.get("name", "")
    if not name or ".." in name or ".." in sub:
        abort(400)

    with get_db() as db:
        row = db.execute(
            "SELECT output_dir, docx_path, company_name FROM jobs WHERE id=?", (task_id,)
        ).fetchone()
    if not row:
        abort(404)

    out_dir = row["output_dir"]
    if not out_dir and row["docx_path"]:
        out_dir = os.path.dirname(row["docx_path"])
    if not out_dir:
        out_dir = _find_output_dir_by_name(row["company_name"] or "")
    if not out_dir or not os.path.isdir(out_dir):
        abort(404)

    base = out_dir if sub == "." else os.path.join(out_dir, sub)
    fpath = os.path.realpath(os.path.join(base, name))
    # 安全校验：不能走出 OUTPUT_DIR
    if not fpath.startswith(os.path.realpath(OUTPUT_DIR)):
        abort(403)
    if not os.path.isfile(fpath):
        abort(404)
    ext = os.path.splitext(name)[1].lower()
    if ext not in _ALLOWED_EXTS:
        abort(403)
    return send_file(fpath, as_attachment=True, download_name=name)


# ──────────────────────────────────────────────
# 打包下载多个文件
# ──────────────────────────────────────────────
@app.route("/download-zip/<task_id>", methods=["POST"])
@login_required
def download_zip(task_id):
    """
    Body JSON: {"files": [{"path": ".", "name": "xxx.docx"}, ...]}
    返回 zip 包
    """
    data = request.get_json(force=True) or {}
    selected = data.get("files", [])
    if not selected:
        abort(400)

    with get_db() as db:
        row = db.execute("SELECT output_dir, docx_path, company_name FROM jobs WHERE id=?",
                         (task_id,)).fetchone()
    if not row:
        abort(404)

    out_dir = row["output_dir"]
    if not out_dir and row["docx_path"]:
        out_dir = os.path.dirname(row["docx_path"])
    if not out_dir:
        abort(404)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in selected:
            sub = item.get("path", ".")
            name = item.get("name", "")
            if not name or ".." in name or ".." in sub:
                continue
            base = out_dir if sub == "." else os.path.join(out_dir, sub)
            fpath = os.path.realpath(os.path.join(base, name))
            if not fpath.startswith(os.path.realpath(OUTPUT_DIR)):
                continue
            if not os.path.isfile(fpath):
                continue
            arcname = name if sub == "." else f"{sub}/{name}"
            zf.write(fpath, arcname)

    buf.seek(0)
    company = row["company_name"]
    return send_file(buf, as_attachment=True,
                     download_name=f"{company}_文件包.zip",
                     mimetype="application/zip")


# ──────────────────────────────────────────────
# 查看历史日志
# ──────────────────────────────────────────────
@app.route("/log/<task_id>")
@login_required
def get_log(task_id):
    with get_db() as db:
        row = db.execute("SELECT log FROM jobs WHERE id=?", (task_id,)).fetchone()
    if not row:
        abort(404)
    return jsonify({"log": row["log"] or ""})


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────
if __name__ == "__main__":
    print("🚀 流水分析系统 Web 服务启动")
    print("   访问地址: http://localhost:7963")
    print("   账号/密码: 见 web_config.json")
    app.run(host="0.0.0.0", port=7963, debug=False, threaded=True)
