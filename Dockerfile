FROM node:22-bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive
ENV NODE_ENV=production
ENV PYTHONUNBUFFERED=1
ENV PYTHON=python3

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 python3-pip \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./requirements.txt
RUN pip3 install --no-cache-dir -r requirements.txt

COPY shadow-app/package.json shadow-app/package-lock.json ./shadow-app/
RUN cd shadow-app && npm ci

COPY . .
RUN cd shadow-app && npm run build

EXPOSE 4174

CMD ["node", "shadow-app/server/index.js"]
