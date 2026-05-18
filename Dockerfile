# ── Base image ──────────────────────────────────────────────────────────────
FROM node:20-slim

# ── Install LibreOffice + Python (for pypdf PDF parsing) ────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    python3 \
    python3-pip \
    python3-pypdf2 \
    fonts-liberation \
    fonts-dejavu \
    && pip3 install --break-system-packages pypdf \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# ── App setup ────────────────────────────────────────────────────────────────
WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

# ── Runtime ──────────────────────────────────────────────────────────────────
ENV PORT=3002
EXPOSE 3002

CMD ["node", "server.js"]
