# PastePro — single container running the Express backend + serving the
# static frontend (index.html) from the same origin.

FROM node:20-bookworm-slim

# ffmpeg: needed for audio extraction (-x) and video re-encode (--recode-video)
# curl/ca-certificates: needed only to fetch the yt-dlp binary below
RUN apt-get update && \
    apt-get install -y --no-install-recommends ffmpeg curl ca-certificates && \
    rm -rf /var/lib/apt/lists/*

# yt-dlp — standalone Linux binary (no Python needed). resolveYtdlpCommand()
# in server.js already falls back to a plain `yt-dlp` on PATH, which this
# satisfies.
RUN curl -L https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp \
      -o /usr/local/bin/yt-dlp && \
    chmod a+rx /usr/local/bin/yt-dlp

WORKDIR /app

# Install deps first so this layer is cached across code-only changes
COPY package*.json ./
RUN npm install --omit=dev

# Only what the running app actually needs — NOT .env, cookies.txt,
# firebase-service-account.json, or the Windows yt-dlp.exe/ffmpeg.exe
# (see .dockerignore).
COPY server.js index.html ./

ENV PORT=3002
EXPOSE 3002

CMD ["node", "server.js"]
