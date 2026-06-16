# Imagem do Validador de E-mails (versao web) para a VPS / Coolify.
# Build context = RAIZ do projeto (a pasta 'leads'), pois precisamos copiar
# tanto 'app-desktop/email_validator_app.py' quanto a pasta 'app-web/'.
#
# No Coolify:
#   - Build Pack: Dockerfile
#   - Dockerfile Location: app-web/Dockerfile
#   - Base Directory / Build Context: / (raiz do repo)
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    DATA_DIR=/data

WORKDIR /app

# Dependencias Python (sem tkinter: o import e opcional no nucleo)
COPY app-web/requirements.txt /app/app-web/requirements.txt
RUN pip install --no-cache-dir -r /app/app-web/requirements.txt

# Codigo: o nucleo reutilizado (em app-desktop) + o app web (app-web).
# Mantemos a MESMA estrutura de pastas do projeto para o import funcionar igual
# local e no container (validator_core procura ../app-desktop).
COPY app-desktop/email_validator_app.py /app/app-desktop/email_validator_app.py
COPY app-web/ /app/app-web/

# Volume persistente para uploads e resultados
RUN mkdir -p /data/uploads /data/outputs
VOLUME ["/data"]

WORKDIR /app/app-web
EXPOSE 8000

# 1 worker (jobs ficam em memoria) + varias threads; sem timeout para uploads
# grandes e processamentos longos.
CMD ["gunicorn", "--workers", "1", "--threads", "16", "--timeout", "0", \
     "--bind", "0.0.0.0:8000", "webapp:app"]
