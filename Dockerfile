FROM python:3.11-slim

# Install LibreOffice (headless), python3-uno (UNO bridge) and fonts
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        python3-uno \
        fonts-liberation \
        fonts-dejavu-core \
        && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Make start.sh executable
RUN chmod +x start.sh

ENV PORT=10000

EXPOSE 10000

# start.sh launches unoserver daemon then Flask
CMD ["./start.sh"]
