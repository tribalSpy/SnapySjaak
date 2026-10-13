FROM node:22-bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive
ENV NODE_ENV=production
ENV PYTHONUNBUFFERED=1
ENV VIRTUAL_ENV=/opt/venv
ENV PATH="/opt/venv/bin:${PATH}"
ENV PYTHON=python3

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 python3-pip python3-venv postgresql-client fontconfig fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./requirements.txt
RUN python3 -m venv "$VIRTUAL_ENV" \
    && pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY shadow-app/package.json shadow-app/package-lock.json ./shadow-app/
RUN cd shadow-app && npm ci

COPY . .
RUN cd shadow-app && npm run build

EXPOSE 4174

# Render instance: 2 GB. Node's own default heap limit was about 1 GB --
# the Inkoop Controle match over thousands of invoices hit it. 1400 MB
# leaves room for the Python workers next to it.
CMD ["node", "--max-old-space-size=1400", "shadow-app/server/index.js"]
