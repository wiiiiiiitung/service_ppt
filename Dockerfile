# Container image for the Sunday worship PPT generator.
#
# The one thing this image exists for is LibreOffice. The app converts the
# legacy .ppt / .doc files the church still circulates into .pptx / .docx by
# shelling out to `soffice` (see file_converter.py). Vercel's @vercel/python
# runtime has no LibreOffice, so on that platform the conversion silently did
# nothing and every legacy file had to be re-saved by hand in PowerPoint
# first. Deploying this image instead removes that step.
#
#   docker build -t service-ppt .
#   docker run --rm -p 5001:5001 -e FLASK_SECRET_KEY=... service-ppt
#
# The CJK fonts matter too: 標楷體 / DFKai-SB are not in the image, and
# text metrics drive where slides break, so a close metric-compatible
# substitute keeps server-side rendering honest.
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-writer \
        fonts-arphic-ukai \
        fonts-arphic-uming \
        fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY . .

# Persisted learned Bible page numbers; mount a volume here to keep them.
ENV BIBLE_PAGES_PATH=/data/bible_pages.json
RUN mkdir -p /data

ENV PORT=5001
EXPOSE 5001
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --timeout 300 --workers 2 app:app"]
