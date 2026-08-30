# PastePro — single container running the Express backend + serving the
# static frontend (index.html) from the same origin.

FROM node:20-bookworm-slim

# ffmpeg: needed for audio extraction (-x) and video re-encode (--recode-video)
# curl/ca-certificates/unzip: needed only to fetch the yt-dlp/deno binaries below
RUN apt-get update && \
    apt-get install -y --no-install-recommends ffmpeg curl ca-certificates unzip && \
    rm -rf /var/lib/apt/lists/*

# Deno — the JS runtime yt-dlp uses (auto-detected on PATH, no flag needed)
# to solve YouTube's JS bot-check/PoToken challenges. Without ANY JS runtime
# present, yt-dlp is far more likely to hit "Sign in to confirm you're not
# a bot" even with valid cookies — this was silently missing before and is
# very likely why deployed requests failed while local ones (which have
# Deno installed) worked.
RUN curl -fL https://github.com/denoland/deno/releases/latest/download/deno-x86_64-unknown-linux-gnu.zip \
      -o /tmp/deno.zip && \
    unzip -o /tmp/deno.zip -d /usr/local/bin && \
    rm /tmp/deno.zip && \
    chmod a+rx /usr/local/bin/deno && \
    deno --version

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
#
# Using the nightly-builds channel instead of the stable release: YouTube's
# anti-bot measures change often enough that yt-dlp's nightly builds (nearly
# daily) regularly ship countermeasures well before they reach a stable tag
# (verified: nightly here is ~10 days newer than the current stable release).
RUN curl -fL https://github.com/yt-dlp/yt-dlp-nightly-builds/releases/latest/download/yt-dlp_linux \
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
