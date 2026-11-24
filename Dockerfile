FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    DATA_DIR=/data

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY wsgi.py ./

RUN mkdir -p /data/uploads /data/exports
VOLUME ["/data"]

EXPOSE 5000

# 单进程多线程即可满足演示；SQLite 写入量小
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--threads", "4", "wsgi:app"]
