import os
import uuid
import shutil
from pathlib import Path

from flask import (
    Flask, request, render_template, send_from_directory,
    redirect, url_for, flash, abort,
)
from werkzeug.utils import secure_filename

from topico1_builder import build_topico1, convert_pptx_to_pdf

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "tmp_uploads"
OUTPUT_DIR = BASE_DIR / "tmp_outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

MAX_CONTENT_LENGTH = 150 * 1024 * 1024  # 150 MB
MAX_JOB_AGE_SECONDS = 2 * 60 * 60  # limpa saídas com mais de 2h


def _limpar_jobs_antigos():
    import time
    agora = time.time()
    for pasta in (UPLOAD_DIR, OUTPUT_DIR):
        for job_dir in pasta.glob("*"):
            try:
                if job_dir.is_dir() and (agora - job_dir.stat().st_mtime) > MAX_JOB_AGE_SECONDS:
                    shutil.rmtree(job_dir, ignore_errors=True)
            except FileNotFoundError:
                pass


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "chave-dev-troque-em-producao")
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    _limpar_jobs_antigos()
    arquivo = request.files.get("licao")
    tipo = request.form.get("tipo", "").strip().lower()

    if not arquivo or arquivo.filename == "":
        flash("Selecione o arquivo .pptx da lição completa.")
        return redirect(url_for("index"))

    if not arquivo.filename.lower().endswith(".pptx"):
        flash("O arquivo precisa ser um .pptx.")
        return redirect(url_for("index"))

    if tipo not in ("jovens", "adultos"):
        flash("Selecione o tipo da lição (jovens ou adultos).")
        return redirect(url_for("index"))

    job_id = uuid.uuid4().hex[:12]
    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id
    job_upload_dir.mkdir(parents=True, exist_ok=True)
    job_output_dir.mkdir(parents=True, exist_ok=True)

    safe_name = secure_filename(arquivo.filename) or "licao.pptx"
    src_path = job_upload_dir / safe_name
    arquivo.save(src_path)

    base_name = Path(safe_name).stem
    out_pptx_name = f"Topico-1-{base_name}.pptx"
    out_pptx_path = job_output_dir / out_pptx_name

    try:
        info = build_topico1(str(src_path), tipo, str(out_pptx_path))
        pdf_path = convert_pptx_to_pdf(str(out_pptx_path), str(job_output_dir))
    except Exception as exc:
        shutil.rmtree(job_upload_dir, ignore_errors=True)
        shutil.rmtree(job_output_dir, ignore_errors=True)
        flash(f"Não foi possível gerar o Tópico 1: {exc}")
        return redirect(url_for("index"))

    # Não precisamos mais guardar o upload original
    shutil.rmtree(job_upload_dir, ignore_errors=True)

    return render_template(
        "resultado.html",
        job_id=job_id,
        pptx_name=Path(out_pptx_path).name,
        pdf_name=Path(pdf_path).name,
        info=info,
    )


@app.route("/download/<job_id>/<path:filename>")
def download(job_id, filename):
    job_output_dir = OUTPUT_DIR / secure_filename(job_id)
    if not job_output_dir.exists():
        abort(404)
    return send_from_directory(job_output_dir, filename, as_attachment=True)


@app.route("/healthz")
def healthz():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
