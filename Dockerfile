FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    TIMEZONE=Europe/Moscow \
    DB_PATH=/data/sprashivatel.db \
    BACKUP_DIR=/data/backups

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# непривилегированный пользователь; /data — единственное место, где бот хранит данные
RUN useradd --create-home --uid 1000 bot \
    && mkdir -p /data /app/exports \
    && chown -R bot:bot /data /app

COPY --chown=bot:bot bot.py .

USER bot
VOLUME ["/data"]

CMD ["python", "bot.py"]
