# Imagem do Validador de E-mails (versao web) para a VPS / Coolify.
# Fica na RAIZ do repo de proposito: o build copia tanto o nucleo
# (app-desktop/email_validator_app.py) quanto o app web (app-web/).
#
# No Coolify (Build Pack: Dockerfile):
#   - Base Directory:     /          (NAO coloque o caminho do Dockerfile aqui!)
#   - Dockerfile Location: Dockerfile (na raiz; este arquivo)
#   - Port:               8000
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
