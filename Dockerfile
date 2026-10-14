# ============================================================
# FICHIER 2/3 : Dockerfile
# (copier sans les commentaires de cette section, uniquement le bloc ci-dessous)
# ============================================================

FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY workflow.py ./

CMD ["python", "-m", "mistralai.workflows.worker", "workflow:SuperPdpCollectWorkflow"]
