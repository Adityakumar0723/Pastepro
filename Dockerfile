# PastePro — single container running the Express backend + serving the
# static frontend (index.html) from the same origin.

FROM node:20-bookworm-slim

# ffmpeg: needed for audio extraction (-x) and video re-encode (--recode-video)
# curl/ca-certificates: needed only to fetch the yt-dlp binary below
RUN apt-get update && \
    apt-get install -y --no-install-recommends ffmpeg curl ca-certificates && \
    rm -rf /var/lib/apt/lists/*

# yt-dlp — MUST be the "yt-dlp_linux" asset, not the plain "yt-dlp" one.
# The plain asset is a Python zipapp (needs a real python3 interpreter on
# PATH, which this minimal image doesn't have); yt-dlp_linux is a genuine
# PyInstaller-frozen standalone ELF binary, no Python required at runtime.
# resolveYtdlpCommand() in server.js falls back to a plain `yt-dlp` on PATH,
# which this satisfies. -f/--fail is required on curl: without it, an HTTP
# error response (rate limit, redirect hiccup, etc.) is written into the
# file as if it were the binary, and the build still "succeeds" with a
# broken file. The trailing `yt-dlp --version` makes the BUILD itself fail
# loudly if that ever happens again, instead of deploying silently broken.
RUN curl -fL https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp_linux \
      -o /usr/local/bin/yt-dlp && \
    chmod a+rx /usr/local/bin/yt-dlp && \
    yt-dlp --version

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
