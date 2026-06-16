# -*- coding: utf-8 -*-
"""
Validador de E-mails - versao Web (para rodar na VPS via Coolify).

Sobe uma planilha (CSV/Excel), processa em segundo plano (MX em tudo + SMTP so
nos dominios corporativos) e disponibiliza a planilha limpa .xlsx para download,
no MESMO modelo da original.

Protegido por senha (HTTP Basic). Configure por variaveis de ambiente:
  APP_USER       (padrao: admin)
  APP_PASSWORD   (OBRIGATORIA; sem ela o app recusa iniciar)
  DATA_DIR       (padrao: /data)  -> volume persistente no Coolify
  MAX_UPLOAD_MB  (padrao: 2048)
  SMTP_RATE      (padrao: 300)    -> sondas SMTP por minuto (dominios diversos)
  DNS_RATE       (padrao: 50)     -> consultas DNS por segundo
  WORKERS        (padrao: 8)      -> paralelismo (I/O-bound; ok em VPS fraca)
"""

import os
import io
import time
import uuid
import threading
import functools

from flask import (Flask, request, redirect, url_for, Response,
                   send_file, abort, render_template_string)
from werkzeug.utils import secure_filename

import validator_core as core

# --------------------------------------------------------------------------- #
#  Configuracao
# --------------------------------------------------------------------------- #
APP_USER = os.environ.get("APP_USER", "admin")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
DATA_DIR = os.environ.get("DATA_DIR", "/data")
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "2048"))
SMTP_RATE = int(os.environ.get("SMTP_RATE", "300"))
DNS_RATE = int(os.environ.get("DNS_RATE", "50"))
WORKERS = int(os.environ.get("WORKERS", "8"))
# SMTP profundo (catch-all + greylisting) e identidade do remetente
DEEP = os.environ.get("DEEP", "1") not in ("0", "false", "False", "")
GREYLIST_RETRIES = int(os.environ.get("GREYLIST_RETRIES", "0"))
GREYLIST_DELAY = int(os.environ.get("GREYLIST_DELAY", "20"))

# Identidade SMTP: em producao aponte para um dominio REAL com rDNS/SPF.
if os.environ.get("SMTP_FROM"):
    core.eva.SMTP_FROM = os.environ["SMTP_FROM"]
if os.environ.get("SMTP_HELO_HOST"):
    core.eva.SMTP_HELO_HOST = os.environ["SMTP_HELO_HOST"]
core.eva.SMTP_USE_STARTTLS = os.environ.get("SMTP_STARTTLS", "1") not in ("0", "false", "False", "")

UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
OUTPUT_DIR = os.path.join(DATA_DIR, "outputs")
SUPPRESSION_PATH = os.path.join(DATA_DIR, "suppression.txt")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXT = {".xlsx", ".xlsm", ".xls", ".csv", ".txt"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# --------------------------------------------------------------------------- #
#  Estado dos jobs (em memoria) + fila de 1 por vez (VPS fraca)
# --------------------------------------------------------------------------- #
_jobs = {}                       # id -> dict
_jobs_lock = threading.Lock()
_run_lock = threading.Lock()     # garante 1 processamento por vez


def _new_job(name, in_path, mode, dedup, excluir):
    jid = uuid.uuid4().hex[:12]
    job = {
        "id": jid, "name": name, "in_path": in_path,
        "mode": mode, "dedup": dedup, "excluir": excluir,
        "status": "fila", "counters": {}, "files": [], "error": "",
        "t0": time.time(), "t_end": None,
        "stop": threading.Event(),
    }
    with _jobs_lock:
        _jobs[jid] = job
    return job


def _run_job(job):
    out_base = os.path.join(OUTPUT_DIR, job["id"] + "_LIMPA.xlsx")
    with _run_lock:                       # serializa: 1 job pesado por vez
        if job["stop"].is_set():
            job["status"] = "cancelado"
            return
        job["status"] = "processando"

        def progress(c):
            job["counters"] = c
            job["files"] = c.get("arquivos", [])

        try:
            suppression = core.load_suppression(SUPPRESSION_PATH)
            c = core.process_file(
                job["in_path"], out_base,
                mode=job["mode"], dedup=job["dedup"],
                excluir_arriscados=job["excluir"],
                deep=DEEP, suppression=suppression,
                greylist_retries=GREYLIST_RETRIES, greylist_delay=GREYLIST_DELAY,
                dns_rate=DNS_RATE, smtp_rate=SMTP_RATE, workers=WORKERS,
                progress_cb=progress, stop_event=job["stop"])
            job["counters"] = c
            job["files"] = c.get("arquivos", [])
            job["status"] = "cancelado" if job["stop"].is_set() else "concluido"
        except Exception as exc:
            job["error"] = str(exc)
            job["status"] = "erro"
        finally:
            job["t_end"] = time.time()


# --------------------------------------------------------------------------- #
#  Autenticacao (HTTP Basic)
# --------------------------------------------------------------------------- #
def _check_auth(u, p):
    return u == APP_USER and p == APP_PASSWORD and APP_PASSWORD != ""


def requires_auth(f):
    @functools.wraps(f)
    def wrapper(*a, **kw):
        auth = request.authorization
        if not auth or not _check_auth(auth.username, auth.password):
            return Response(
                "Acesso restrito.", 401,
                {"WWW-Authenticate": 'Basic realm="Validador"'})
        return f(*a, **kw)
    return wrapper


# --------------------------------------------------------------------------- #
#  Paginas
# --------------------------------------------------------------------------- #
PAGE = """
<!doctype html><html lang="pt-br"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Validador de E-mails</title>
<style>
 body{font-family:system-ui,Segoe UI,Arial,sans-serif;max-width:880px;margin:24px auto;padding:0 14px;color:#1b1b1b}
 h1{font-size:20px} h2{font-size:16px;margin-top:28px}
 .card{border:1px solid #ddd;border-radius:10px;padding:16px;margin:12px 0;background:#fafafa}
 label{display:block;margin:8px 0 2px;font-size:14px}
 input[type=file]{margin:6px 0}
 .btn{background:#2563eb;color:#fff;border:0;padding:9px 16px;border-radius:8px;font-size:14px;cursor:pointer}
 .btn.gray{background:#6b7280}
 table{border-collapse:collapse;width:100%;font-size:13px}
 td,th{border:1px solid #e3e3e3;padding:6px 8px;text-align:left}
 .pill{padding:2px 8px;border-radius:20px;font-size:12px;color:#fff}
 .fila{background:#9ca3af}.processando{background:#2563eb}.concluido{background:#16a34a}
 .erro{background:#dc2626}.cancelado{background:#b45309}
 small{color:#666} a{color:#2563eb}
 .bar{height:8px;background:#e5e7eb;border-radius:6px;overflow:hidden;margin:4px 0}
 .bar>span{display:block;height:100%;background:#2563eb}
 .btn.mini{padding:4px 10px;font-size:12px}
 .sup{background:#fffbea;border-color:#f5e08c}
</style></head><body>
<h1>Validador de E-mails &mdash; limpeza de listas</h1>
<div class="card">
 <form method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
  <label><b>1) Planilha de leads</b> (.xlsx / .csv)</label>
  <input type="file" name="file" accept=".xlsx,.xlsm,.xls,.csv,.txt" required>
  <label><input type="checkbox" name="dedup" checked> Remover e-mails duplicados</label>
  <label><input type="checkbox" name="excluir"> Remover automaticos (noreply@, newsletter@)</label>
  <label><input type="checkbox" name="rapido"> Modo rapido (so MX, <b>sem</b> SMTP)</label>
  <p><small>Padrao: MX em tudo + <b>SMTP profundo so nos dominios corporativos</b>
   (deteta catch-all; gmail/hotmail/yahoo sao aceitos por MX). Remove tambem os
   e-mails que ja estao na lista de supressao. Saida em .xlsx no mesmo modelo
   da original; acima de 1 milhao de linhas e dividida em partes.</small></p>
  <button class="btn" type="submit">Enviar e processar</button>
 </form>
</div>

<div class="card sup">
 <form method="post" action="{{ url_for('suppression_upload') }}" enctype="multipart/form-data">
  <label><b>2) Lista de supressao</b> &mdash; bounces / Do-Not-Contact do Mautic</label>
  <p><small>Suba aqui o export de e-mails que <b>bouncaram</b> ou estao em
   Do-Not-Contact. Eles passam a ser removidos automaticamente de toda lista
   nova. <b>Na supressao agora: {{ '{:,}'.format(sup_count).replace(',','.') }} e-mails.</b></small></p>
  <input type="file" name="file" accept=".xlsx,.xlsm,.xls,.csv,.txt" required>
  <button class="btn gray" type="submit">Adicionar a supressao</button>
 </form>
</div>

<h2>Processamentos</h2>
{% if jobs %}
<table>
 <tr><th>Arquivo</th><th>Status</th><th>Progresso</th><th>Resultado</th></tr>
 {% for j in jobs %}
 {% set c = j.counters %}
 <tr>
  <td>{{ j.name }}<br><small>{{ j.id }}</small></td>
  <td><span class="pill {{ j.status }}">{{ j.status }}</span>
   {% if j.status in ['fila','processando'] %}
     <form method="post" action="{{ url_for('cancel', jid=j.id) }}" style="margin-top:6px">
       <button class="btn gray mini" type="submit">Cancelar</button></form>
   {% endif %}
  </td>
  <td>
   {% if c %}
     {% set proc = c.get('processadas',0) %}{% set tot = c.get('total_estimado',0) %}
     {% if tot %}{% set pct = (proc*100//tot) if tot else 0 %}
       <div class="bar"><span style="width:{{ pct }}%"></span></div>{{ pct }}%<br>{% endif %}
     {{ '{:,}'.format(proc).replace(',','.') }}{% if tot %} / {{ '{:,}'.format(tot).replace(',','.') }}{% endif %}
     <br><small>mantidas: <b>{{ '{:,}'.format(c.get('mantidas',0)).replace(',','.') }}</b>
     | corp.: {{ '{:,}'.format(c.get('corporativos',0)).replace(',','.') }}
     | catch-all: {{ '{:,}'.format(c.get('catch_all',0)).replace(',','.') }}
     | invalidos: {{ '{:,}'.format(c.get('invalido',0)).replace(',','.') }}
     | suprimidos: {{ '{:,}'.format(c.get('suprimidos',0)).replace(',','.') }}
     | dupl.: {{ '{:,}'.format(c.get('duplicados',0)).replace(',','.') }}</small>
   {% else %}<small>aguardando...</small>{% endif %}
  </td>
  <td>
   {% if j.status == 'concluido' or (j.files and j.status in ['cancelado']) %}
     {% for f in j.files %}
       <a href="{{ url_for('download', jid=j.id, idx=loop.index0) }}">baixar parte {{ loop.index }}</a><br>
     {% endfor %}
   {% elif j.status == 'erro' %}<small style="color:#dc2626">{{ j.error }}</small>
   {% elif j.status == 'processando' %}<small>em andamento...</small>
   {% else %}<small>-</small>{% endif %}
  </td>
 </tr>
 {% endfor %}
</table>
<p><small>Esta pagina atualiza sozinha a cada 8s enquanto houver algo processando.</small></p>
{% else %}<p><small>Nenhum processamento ainda.</small></p>{% endif %}

{% if refresh %}<script>setTimeout(function(){location.reload()},8000)</script>{% endif %}
</body></html>
"""


def _suppression_count():
    try:
        return len(core.load_suppression(SUPPRESSION_PATH))
    except Exception:
        return 0


@app.route("/")
@requires_auth
def index():
    with _jobs_lock:
        jobs = sorted(_jobs.values(), key=lambda j: j["t0"], reverse=True)
    refresh = any(j["status"] in ("fila", "processando") for j in jobs)
    return render_template_string(PAGE, jobs=jobs, refresh=refresh,
                                  sup_count=_suppression_count())


@app.route("/upload", methods=["POST"])
@requires_auth
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        abort(400, "Nenhum arquivo enviado.")
    name = secure_filename(f.filename)
    ext = os.path.splitext(name)[1].lower()
    if ext not in ALLOWED_EXT:
        abort(400, "Formato nao suportado. Use .xlsx ou .csv.")
    jid_part = uuid.uuid4().hex[:12]
    in_path = os.path.join(UPLOAD_DIR, jid_part + "_" + name)
    f.save(in_path)                       # werkzeug grava em disco (stream)

    dedup = request.form.get("dedup") == "on"
    excluir = request.form.get("excluir") == "on"
    mode = "rapido" if request.form.get("rapido") == "on" else "completo"

    job = _new_job(name, in_path, mode, dedup, excluir)
    threading.Thread(target=_run_job, args=(job,), daemon=True).start()
    return redirect(url_for("index"))


@app.route("/suppression/upload", methods=["POST"])
@requires_auth
def suppression_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        abort(400, "Nenhum arquivo enviado.")
    name = secure_filename(f.filename)
    ext = os.path.splitext(name)[1].lower()
    if ext not in ALLOWED_EXT:
        abort(400, "Formato nao suportado. Use .xlsx ou .csv.")
    tmp = os.path.join(UPLOAD_DIR, "sup_" + uuid.uuid4().hex[:8] + "_" + name)
    f.save(tmp)
    try:
        emails = core.extract_emails_from_file(tmp)
        core.add_to_suppression(SUPPRESSION_PATH, emails)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    return redirect(url_for("index"))


@app.route("/download/<jid>/<int:idx>")
@requires_auth
def download(jid, idx):
    with _jobs_lock:
        job = _jobs.get(jid)
    if not job or idx >= len(job["files"]):
        abort(404)
    path = job["files"][idx]
    if not os.path.isfile(path):
        abort(404)
    base = os.path.splitext(job["name"])[0]
    suffix = "" if idx == 0 else f"_parte{idx + 1}"
    dl_name = f"{base}_LIMPA{suffix}.xlsx"
    return send_file(path, as_attachment=True, download_name=dl_name)


@app.route("/cancel/<jid>", methods=["POST"])
@requires_auth
def cancel(jid):
    with _jobs_lock:
        job = _jobs.get(jid)
    if job:
        job["stop"].set()
    return redirect(url_for("index"))


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    if not APP_PASSWORD:
        raise SystemExit("Defina APP_PASSWORD (variavel de ambiente) antes de iniciar.")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), threaded=True)
