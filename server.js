// ════════════════════════════════════════════════════════════
//  PastePro Backend Server
//  Node.js + Express + yt-dlp
//  Requirements: npm install express cors mongoose bcryptjs jsonwebtoken dotenv
//  System requirement: yt-dlp installed (pip install yt-dlp)
// ════════════════════════════════════════════════════════════

const express   = require('express');
const cors      = require('cors');
const { exec, execFile } = require('child_process');
const path      = require('path');
const fs        = require('fs');
const os        = require('os');
const mongoose  = require('mongoose');
const bcrypt    = require('bcryptjs');
const jwt       = require('jsonwebtoken');
const pdfParse  = require('pdf-parse');
const mammoth   = require('mammoth');
const XLSX      = require('xlsx');
const AdmZip    = require('adm-zip');
require('dotenv').config();

const app  = express();
// Port 3001 is occupied by VS Code's local webview service on this machine.
const PORT = process.env.PORT || 3002;

// ─── MongoDB ──────────────────────────────────────────────
// Override this in .env when needed. MongoDB must be running locally.
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/pastepro';
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-development-secret';

// ─── OpenRouter (Playground Query section) ─────────────────
// Key sirf yahan, server-side, rehta hai — kabhi bhi frontend/index.html mein
// mat daalna, warna koi bhi visitor DevTools se churaake use kar sakta hai.
// Default free models — verified directly against the live API before
// picking these: several other free models on OpenRouter (gemma-4-31b,
// glm-5.2, nemotron nano) were hitting shared-pool 429s/502s at the time,
// these two were not. OPENROUTER_MODEL retries once with
// OPENROUTER_FALLBACK_MODEL if the primary is upstream-rate-limited,
// since free models on OpenRouter share a pool that can get busy.
const OPENROUTER_API_KEY       = process.env.OPENROUTER_API_KEY || '';
const OPENROUTER_MODEL         = process.env.OPENROUTER_MODEL || 'minimax/minimax-m3:free';
const OPENROUTER_FALLBACK_MODEL = process.env.OPENROUTER_FALLBACK_MODEL || 'cohere/north-mini-code:free';
// Image attachments need a vision-capable model — verified directly against
// the live API with a real photo (not just going by OpenRouter's listed
// "input_modalities"): minimax-m3 and gemma-4-31b both refused/rate-limited
// on actual image input despite being listed as vision-capable, this one
// correctly described the test photo.
const OPENROUTER_VISION_MODEL  = process.env.OPENROUTER_VISION_MODEL || 'dots-studio/dots-3-note-preview:free';
const OPENROUTER_BASE_URL      = 'https://openrouter.ai/api/v1';

const userSchema = new mongoose.Schema({
  name: { type: String, required: true, trim: true, maxlength: 80 },
  email: { type: String, required: true, trim: true, lowercase: true, unique: true },
  password: { type: String, required: true }
}, { timestamps: true });

const downloadSchema = new mongoose.Schema({
  user: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
  url: { type: String, required: true },
  filename: { type: String, required: true },
  type: { type: String, enum: ['video', 'audio'], required: true },
  quality: { type: String, required: true }
}, { timestamps: true });

// Per-user activity log — "kis user ne kab kya kiya, kahan se" ka record.
// details Mixed hai kyunki har action ka shape alag hota hai (download vs
// search vs page_view). Kabhi bhi Auth password/token store nahi hota yahan.
const activityLogSchema = new mongoose.Schema({
  user:      { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  action:    { type: String, required: true }, // 'login' | 'signup' | 'page_view' | 'download' | 'search' | 'convert' | 'playground_query'
  details:   { type: mongoose.Schema.Types.Mixed, default: {} },
  ip:        String,
  userAgent: String,
}, { timestamps: true });

const User = mongoose.model('User', userSchema);
const Download = mongoose.model('Download', downloadSchema);
const ActivityLog = mongoose.model('ActivityLog', activityLogSchema);

// Best-effort — logging kabhi bhi asal feature ko fail nahi karna chahiye,
// isliye caller ko await karne ki bhi zaroorat nahi (fire-and-forget),
// lekin agar await kiya jaaye toh bhi ye khud kabhi throw nahi karta.
async function logActivity(req, action, details = {}) {
  try {
    await ActivityLog.create({
      user: req.user?.id,
      action,
      details,
      ip: (req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.socket?.remoteAddress || '',
      userAgent: req.headers['user-agent'] || '',
    });
  } catch (e) {
    console.error(`Activity log failed (${action}):`, e.message);
  }
}

// ─── Middleware ───────────────────────────────────────────
// No cookie credentials are used; JWTs are sent in the Authorization header.
// Allow the frontend both from a local dev server and when index.html is opened directly (Origin: null).
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
// Default express.json() limit is 100kb — way too small for base64-encoded
// attachments on the Playground Query. Multiple files can now be attached
// to one message (combined cap 20MB raw, enforced in /api/query) — 30mb
// comfortably covers that once base64's ~33% size overhead is added.
app.use(express.json({ limit: '30mb' }));

// Downloads folder — files yahan save honge.
// Kept OUTSIDE the project directory on purpose: if index.html is opened via
// VS Code Live Server, Live Server watches the whole project folder and
// reloads the page the moment a new file appears in it. Writing downloads
// inside the project would make Live Server refresh the tab mid-download,
// killing the in-flight fetch before the UI ever sees the response.
const DOWNLOADS_DIR = path.join(os.tmpdir(), 'pastepro-downloads');
if (!fs.existsSync(DOWNLOADS_DIR)) fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });

// COOKIES_PATH may point at a read-only mount (e.g. Render Secret Files,
// which are mounted read-only at /etc/secrets/<file>). yt-dlp always tries
// to save the cookie jar back to whatever --cookies path it was given when
// it exits — on a read-only mount that crashes with
// "OSError: Read-only file system" on every single run, even though the
// cookies themselves load and work fine. Copy the source file once at
// startup into a writable path (inside DOWNLOADS_DIR, already writable)
// and use that copy for every yt-dlp invocation instead.
const RAW_COOKIES_PATH = process.env.COOKIES_PATH || path.join(__dirname, 'cookies.txt');
const COOKIES_PATH = path.join(DOWNLOADS_DIR, 'cookies.txt');
try {
  if (fs.existsSync(RAW_COOKIES_PATH)) {
    fs.copyFileSync(RAW_COOKIES_PATH, COOKIES_PATH);
    // copyFileSync can inherit the source's read-only mode on the copy too
    // (confirmed on Windows; Linux behavior can vary by filesystem) — force
    // it writable explicitly rather than relying on that.
    fs.chmodSync(COOKIES_PATH, 0o644);
    console.log(`cookies.txt copied from ${RAW_COOKIES_PATH} to writable path ${COOKIES_PATH}`);
  }
} catch (e) {
  console.error('Could not copy cookies.txt to a writable path:', e.message);
}

// Optional proxy for yt-dlp's outbound requests (search/download/captions).
// Cookies + Deno + player-client retries reduce YouTube's bot-check on a
// datacenter IP, but can't fully eliminate it — that's an IP-reputation
// block, not something any yt-dlp flag can fully undo. Routing through a
// residential/mobile proxy is the only way to guarantee it stops happening.
// Set YTDLP_PROXY to a proxy URL (e.g. http://user:pass@host:port or
// socks5://user:pass@host:port) from any proxy provider to enable this —
// left unset, everything behaves exactly as before (direct connection).
const YTDLP_PROXY = process.env.YTDLP_PROXY || '';
function proxyFlag() {
  return YTDLP_PROXY ? `--proxy "${YTDLP_PROXY}"` : '';
}
function proxyArgs() {
  return YTDLP_PROXY ? ['--proxy', YTDLP_PROXY] : [];
}

// Local speech-to-text (whisper.cpp), baked into the Docker image at
// /opt/whisper — this is what makes word-by-word captions work when the
// source has no caption data of its own (Instagram never has any;
// Twitter/TikTok only sometimes; and now also the Search-page live
// preview, when YouTube itself has no auto-captions for that video).
// Multilingual model (not the English-only "tiny.en") on purpose — many
// PastePro users search/preview Hindi videos, and an English-only model
// would mangle non-English speech instead of transcribing it. Optional by
// design: if the binary/model aren't present (e.g. local dev on this
// Windows machine, where only the Docker image has them), this fallback
// simply doesn't run — same graceful behavior as before it existed.
const WHISPER_BIN   = process.env.WHISPER_BIN   || '/opt/whisper/whisper-cli';
const WHISPER_MODEL = process.env.WHISPER_MODEL || '/opt/whisper/ggml-tiny.bin';
const WHISPER_READY = fs.existsSync(WHISPER_BIN) && fs.existsSync(WHISPER_MODEL);

// Only pass --cookies when the file genuinely exists and is readable. If
// COOKIES_PATH points at a path whose directory doesn't exist (e.g. a
// misconfigured Secret File mount), yt-dlp doesn't just skip cookies — it
// crashes on exit trying to save the cookie jar back, breaking EVERY
// yt-dlp call (search, download, captions) at once. Checking here means a
// bad cookies setup degrades to "no cookies" instead of taking everything down.
function cookiesFlag() {
  try {
    return fs.existsSync(COOKIES_PATH) ? `--cookies "${COOKIES_PATH}"` : '';
  } catch (e) {
    return '';
  }
}
function cookiesArgs() {
  try {
    return fs.existsSync(COOKIES_PATH) ? ['--cookies', COOKIES_PATH] : [];
  } catch (e) {
    return [];
  }
}

// Serve the frontend itself (only index.html — not the whole directory,
// so .env / cookies.txt / node_modules never become web-accessible).
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Static serve for downloaded files
app.use('/files', (req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Range');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Range, Content-Length, Accept-Ranges');
  // Only force a Save-As dialog for explicit downloads (?dl=1) — a bare
  // request (e.g. from a <video>/<audio> preview player) should play inline.
  if (req.query.dl) res.setHeader('Content-Disposition', 'attachment');
  next();
}, express.static(DOWNLOADS_DIR));

// ─── Auth Middleware ──────────────────────────────────────
async function requireAuth(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Login karo pehle' });
  }
  const token = authHeader.split('Bearer ')[1];
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Token invalid hai. Dobara login karo' });
  }
}

function createToken(user) {
  return jwt.sign({ id: user._id.toString(), email: user.email, name: user.name }, JWT_SECRET, { expiresIn: '7d' });
}

// ─── Authentication routes ────────────────────────────────
app.post('/api/auth/signup', async (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const email = String(req.body.email || '').trim().toLowerCase();
    const password = String(req.body.password || '');
    if (!name || !email || !password) return res.status(400).json({ error: 'Name, email aur password zaroori hain' });
    if (!/^\S+@\S+\.\S+$/.test(email)) return res.status(400).json({ error: 'Valid email daalo' });
    if (password.length < 6) return res.status(400).json({ error: 'Password 6+ characters ka hona chahiye' });
    if (await User.exists({ email })) return res.status(409).json({ error: 'Email already registered hai' });

    const user = await User.create({ name, email, password: await bcrypt.hash(password, 12) });
    req.user = { id: user._id.toString() };
    logActivity(req, 'signup', { email: user.email, name: user.name });
    res.status(201).json({ token: createToken(user), user: { id: user._id, name: user.name, email: user.email } });
  } catch (error) {
    console.error('Signup error:', error);
    res.status(500).json({ error: 'Account create nahi ho saka. Dobara try karo' });
  }
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const email = String(req.body.email || '').trim().toLowerCase();
    const password = String(req.body.password || '');
    const user = await User.findOne({ email });
    if (!user || !(await bcrypt.compare(password, user.password))) {
      return res.status(401).json({ error: 'Email ya password galat hai' });
    }
    req.user = { id: user._id.toString() };
    logActivity(req, 'login', { email: user.email });
    res.json({ token: createToken(user), user: { id: user._id, name: user.name, email: user.email } });
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ error: 'Login nahi ho saka. Dobara try karo' });
  }
});

// ─── Supported URL check ──────────────────────────────────
function isSupportedUrl(url) {
  const patterns = [
    /youtube\.com/, /youtu\.be/,
    /instagram\.com/,
    /twitter\.com/, /x\.com/,
    /tiktok\.com/,
    /facebook\.com/,
    /vimeo\.com/,
    /dailymotion\.com/,
    /reddit\.com/
  ];
  return patterns.some(p => p.test(url));
}

// ─── Download-time format choices (yt-dlp khud handle karta hai) ──
const VALID_AUDIO_DL_FORMATS = ['mp3', 'm4a', 'wav', 'flac', 'opus'];
const VALID_VIDEO_DL_FORMATS = ['mp4', 'mkv', 'webm', 'mov', 'avi'];

function resolveDlFormat(type, format) {
  if (type === 'audio') return VALID_AUDIO_DL_FORMATS.includes(format) ? format : 'mp3';
  return VALID_VIDEO_DL_FORMATS.includes(format) ? format : 'mp4';
}

// ─── Quality to yt-dlp format ─────────────────────────────
function getYtdlpFormat(quality, type, format) {
  if (type === 'audio') {
    const audioQualityMap = {
      'best':    '0',
      '320kbps': '320K',
      '192kbps': '192K',
      '128kbps': '128K',
    };
    const aq = audioQualityMap[quality] || '0';
    return `-x --audio-format ${format} --audio-quality ${aq}`;
  }
  const heightMap = { '1080p': 1080, '720p': 720, '480p': 480, '360p': 360 };
  const targetRes  = heightMap[quality] || 720;
  // -S "res:X" (format SORTING) instead of -f "bestvideo[height<=X]..."
  // (format FILTERING): a hard height<=X filter fails outright for portrait
  // video (Instagram Reels, TikTok, YouTube Shorts) — their pixel *height*
  // is the long/vertical side (e.g. a 720-wide reel reports height=1280),
  // so "height<=720" excludes every format and yt-dlp errors with
  // "Requested format is not available". Sorting by res:X picks whichever
  // dimension is the short side, so it works for both orientations, and
  // gracefully falls back to the closest available quality instead of
  // hard-failing when the exact tier isn't offered.
  // --recode-video re-encodes only if the container actually needs it,
  // so mp4 (the common case) stays a fast remux.
  return `-S "res:${targetRes}" -f "bv*+ba/b" --recode-video ${format}`;
}

function quoteIfPath(value) {
  if (typeof value !== 'string') return value;
  const looksLikePath = value.includes(path.sep) || value.endsWith('.exe') || value.endsWith('/yt-dlp');
  return looksLikePath && value.includes(' ') ? `"${value}"` : value;
}

function testYtdlpCommand(candidate) {
  return new Promise(resolve => {
    const command = `${quoteIfPath(candidate)} --version`;
    exec(command, { shell: true, timeout: 10000 }, (err) => {
      resolve(!err);
    });
  });
}

async function resolveYtdlpCommand() {
  const localExe = path.join(__dirname, 'yt-dlp.exe');
  const localBin = path.join(__dirname, 'yt-dlp');
  const candidates = [];

  if (process.env.YTDLP_PATH) candidates.push(process.env.YTDLP_PATH);
  if (fs.existsSync(localExe)) candidates.push(localExe);
  if (fs.existsSync(localBin)) candidates.push(localBin);
  candidates.push('py -m yt_dlp', 'python -m yt_dlp', 'python3 -m yt_dlp', 'yt-dlp');

  for (const candidate of candidates) {
    try {
      if (await testYtdlpCommand(candidate)) {
        return candidate;
      }
    } catch (e) {
      // ignore and try next candidate
    }
  }
  return null;
}

// ─── ffmpeg resolve (format conversion ke liye) ───────────
async function resolveFfmpegCommand() {
  const localExe = path.join(__dirname, 'ffmpeg.exe');
  const candidates = [];

  if (process.env.FFMPEG_PATH) candidates.push(process.env.FFMPEG_PATH);
  if (fs.existsSync(localExe)) candidates.push(localExe);
  candidates.push('ffmpeg');

  for (const candidate of candidates) {
    try {
      const works = await new Promise(resolve => {
        exec(`${quoteIfPath(candidate)} -version`, { shell: true, timeout: 10000 }, (err) => resolve(!err));
      });
      if (works) return candidate;
    } catch (e) {
      // ignore and try next candidate
    }
  }
  return null;
}

// ─── Convert ke liye supported formats ────────────────────
const AUDIO_CONVERT_FORMATS = {
  mp3:  '-vn -codec:a libmp3lame -q:a 2',
  m4a:  '-vn -codec:a aac -b:a 192k',
  wav:  '-vn -codec:a pcm_s16le',
  flac: '-vn -codec:a flac',
  opus: '-vn -codec:a libopus -b:a 192k',
};
const VIDEO_CONVERT_FORMATS = {
  mp4:  '-c:v libx264 -preset veryfast -crf 23 -c:a aac',
  mkv:  '-c:v libx264 -preset veryfast -crf 23 -c:a aac',
  webm: '-c:v libvpx-vp9 -crf 32 -b:v 0 -c:a libopus',
  mov:  '-c:v libx264 -preset veryfast -crf 23 -c:a aac',
  avi:  '-c:v mpeg4 -q:v 5 -c:a libmp3lame',
};

// ─── Clean old files (30 min se purani) ──────────────────
function cleanOldFiles() {
  const now = Date.now();
  try {
    fs.readdirSync(DOWNLOADS_DIR).forEach(file => {
      const filepath = path.join(DOWNLOADS_DIR, file);
      const stat = fs.statSync(filepath);
      if (now - stat.mtimeMs > 30 * 60 * 1000) {
        fs.unlinkSync(filepath);
        console.log('Deleted old file:', file);
      }
    });
  } catch(e) { console.error('Cleanup error:', e); }
}
setInterval(cleanOldFiles, 10 * 60 * 1000);

// ─── Auto Captions — word-by-word sync (YouTube ke auto-caption VTT se) ──
// YouTube ke auto-caption VTT "rolling" style mein aata hai: har cue purane
// resolved words + naye words (inline <HH:MM:SS.mmm> timestamps ke saath)
// repeat karta hai. Hum sirf NAYE words nikalte hain (already-seen prefix ko
// hata kar) taaki ek clean, non-duplicate, time-ordered word list bane.
function vttTimeToSeconds(t) {
  const [h, m, sRest] = t.split(':');
  return parseInt(h, 10) * 3600 + parseInt(m, 10) * 60 + parseFloat(sRest);
}

function parseVttWords(raw) {
  const blocks = raw.split(/\r?\n\r?\n+/).map(b => b.trim()).filter(Boolean);
  const words = [];
  let lastResolvedWords = [];

  for (const block of blocks) {
    const lines = block.split(/\r?\n/);
    const timingLineIdx = lines.findIndex(l => l.includes('-->'));
    if (timingLineIdx === -1) continue; // WEBVTT/Kind/Language header block

    const m = lines[timingLineIdx].match(/(\d{2}:\d{2}:\d{2}\.\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}\.\d{3})/);
    if (!m) continue;

    const cueStart = vttTimeToSeconds(m[1]);
    const cueEnd   = vttTimeToSeconds(m[2]);
    if (cueEnd - cueStart < 0.1) continue; // rolling "settle" cue — no new content, ignore

    const textLines = lines.slice(timingLineIdx + 1).join('\n');
    if (!textLines.trim()) continue;

    const resolvedText  = textLines.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    const resolvedWords = resolvedText.split(' ').filter(Boolean);

    let common = 0;
    while (common < lastResolvedWords.length && common < resolvedWords.length && lastResolvedWords[common] === resolvedWords[common]) common++;
    const newWordCount = resolvedWords.length - common;
    if (newWordCount <= 0) { lastResolvedWords = resolvedWords; continue; }

    // Text ko inline timestamp tags par split karo, har tag ke baad wala
    // text usi tag ke time par bola gaya hota hai.
    const parts = textLines.split(/(<\d{2}:\d{2}:\d{2}\.\d{3}>)/g);
    let currentTime = cueStart;
    const timedWords = [];
    for (const part of parts) {
      const tagMatch = part.match(/^<(\d{2}:\d{2}:\d{2}\.\d{3})>$/);
      if (tagMatch) { currentTime = vttTimeToSeconds(tagMatch[1]); continue; }
      const cleanPart = part.replace(/<\/?c>/g, '');
      cleanPart.split(/\s+/).map(w => w.trim()).filter(Boolean).forEach(w => {
        timedWords.push({ time: Math.round(currentTime * 100) / 100, text: w });
      });
    }

    // Is cue mein sirf NAYE words chahiye (jo pehle emit nahi hue).
    timedWords.slice(-newWordCount).forEach(w => words.push(w));
    lastResolvedWords = resolvedWords;
  }
  return words;
}

// Same idea as fetchAutoCaptions but for an arbitrary language code — used
// as a fallback when the direct timedtext URL (from --dump-json metadata)
// gets rate-limited by YouTube itself (verified directly: repeated direct
// fetches to the same video's caption URL can 429 independently of
// yt-dlp's own bot-check). yt-dlp fetching it fresh avoids that.
function fetchCaptionsViaYtdlpLang(ytdlpCmd, videoId, langCode) {
  return new Promise((resolve) => {
    const subBase = path.join(DOWNLOADS_DIR, `cap_${videoId}_${langCode}`);
    const url = `https://www.youtube.com/watch?v=${videoId}`;
    const command = `${quoteIfPath(ytdlpCmd)} --write-auto-sub --sub-lang ${langCode} --skip-download --sub-format vtt ${cookiesFlag()} ${proxyFlag()} --paths "temp:${DOWNLOADS_DIR}" --output "${subBase}.%(ext)s" "${url}"`;
    exec(command, { timeout: 20 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error) => {
      const vttPath = `${subBase}.${langCode}.vtt`;
      if (error || !fs.existsSync(vttPath)) return resolve([]);
      try {
        const raw   = fs.readFileSync(vttPath, 'utf8');
        const words = parseVttWords(raw);
        fs.unlinkSync(vttPath);
        resolve(words);
      } catch (e) {
        resolve([]);
      }
    });
  });
}

function fetchAutoCaptions(url, ytdlpCmd, safeId) {
  return new Promise((resolve) => {
    const subBase = path.join(DOWNLOADS_DIR, safeId);
    const command  = `${quoteIfPath(ytdlpCmd)} --write-auto-sub --sub-lang en --skip-download --sub-format vtt ${cookiesFlag()} ${proxyFlag()} --paths "temp:${DOWNLOADS_DIR}" --output "${subBase}.%(ext)s" "${url}"`;
    exec(command, { timeout: 30 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error, stdout, stderr) => {
      const vttPath = `${subBase}.en.vtt`;
      if (error || !fs.existsSync(vttPath)) {
        console.log(`[captions] no auto-sub for ${url}${error ? ` (${(stderr || error.message || '').slice(0, 200)})` : ''}`);
        return resolve([]);
      }
      try {
        const raw   = fs.readFileSync(vttPath, 'utf8');
        const words = parseVttWords(raw);
        fs.unlinkSync(vttPath);
        resolve(words);
      } catch (e) {
        console.error('[captions] auto-sub VTT parse failed:', e.message);
        resolve([]);
      }
    });
  });
}

// Fallback for every platform that doesn't hand us real caption data:
// transcribe the file we just downloaded, ourselves, with a local
// whisper.cpp model. `-ml 2` caps each VTT cue at ~2 words so the
// existing parseVttWords() (built for YouTube's word-timed cues) gets
// near-word-level timing here too, instead of one giant per-sentence cue.
// Bounded and best-effort like fetchAutoCaptions: any failure (missing
// binary, ffmpeg hiccup, timeout) just resolves to [] rather than
// touching the download response.
function fetchWhisperCaptions(mediaPath, ffmpegCmd) {
  return new Promise((resolve) => {
    if (!WHISPER_READY) {
      console.log(`[whisper] skipped — WHISPER_READY is false (bin exists: ${fs.existsSync(WHISPER_BIN)}, model exists: ${fs.existsSync(WHISPER_MODEL)})`);
      return resolve([]);
    }
    try {
      const stat = fs.statSync(mediaPath);
      // Rough proxy for "too long to transcribe in a reasonable time" on a
      // free-tier CPU — skip rather than risk a very long-running process.
      if (stat.size > 150 * 1024 * 1024) {
        console.log(`[whisper] skipped — ${mediaPath} too large (${Math.round(stat.size / 1024 / 1024)}MB)`);
        return resolve([]);
      }
    } catch (e) {
      console.error(`[whisper] media file missing: ${mediaPath}`);
      return resolve([]);
    }

    const base    = mediaPath.replace(/\.[^/.]+$/, '') + '.whisper';
    const wavPath = `${base}.wav`;
    const vttPath = `${base}.vtt`;
    const extractCmd = `${quoteIfPath(ffmpegCmd)} -y -i "${mediaPath}" -vn -ar 16000 -ac 1 -c:a pcm_s16le "${wavPath}"`;

    exec(extractCmd, { timeout: 60 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (ffErr, ffStdout, ffStderr) => {
      if (ffErr || !fs.existsSync(wavPath)) {
        console.error(`[whisper] ffmpeg audio-extract failed for ${mediaPath}:`, (ffStderr || ffErr?.message || '').slice(0, 300));
        return resolve([]);
      }

      const whisperCmd = `${quoteIfPath(WHISPER_BIN)} -m "${WHISPER_MODEL}" -f "${wavPath}" -ml 2 -ovtt -of "${base}" -np`;
      exec(whisperCmd, { timeout: 120 * 1000, shell: true, cwd: DOWNLOADS_DIR, env: { ...process.env, LD_LIBRARY_PATH: path.dirname(WHISPER_BIN) } }, (wErr, wStdout, wStderr) => {
        try { fs.unlinkSync(wavPath); } catch (e) {}
        if (wErr || !fs.existsSync(vttPath)) {
          console.error(`[whisper] transcription failed for ${mediaPath}:`, (wStderr || wErr?.message || '').slice(0, 300));
          return resolve([]);
        }
        try {
          const raw   = fs.readFileSync(vttPath, 'utf8');
          const words = parseVttWords(raw);
          fs.unlinkSync(vttPath);
          console.log(`[whisper] transcribed ${mediaPath}: ${words.length} words`);
          resolve(words);
        } catch (e) {
          console.error('[whisper] VTT parse failed:', e.message);
          resolve([]);
        }
      });
    });
  });
}

// Search-page preview fallback: when YouTube itself has no auto-captions
// for a video (confirmed happens — not every video has speech, or ASR
// data), transcribe it ourselves instead of leaving captions empty. There's
// no already-downloaded file to reuse here (this runs before the user
// decides to download anything), so this pulls just the audio stream —
// not the full video — then reuses fetchWhisperCaptions() on it. Capped by
// duration (not just file size like the download-flow version) since a
// live preview shouldn't kick off a multi-minute transcription for a
// 2-hour livestream — it degrades to "no captions" instead, same as any
// other unavailable-captions case.
const MAX_WHISPER_PREVIEW_SECONDS = 20 * 60;

function fetchWhisperCaptionsForSearch(videoId, durationSeconds, ytdlpCmd, ffmpegCmd) {
  return new Promise((resolve) => {
    if (!WHISPER_READY) return resolve([]);
    if (durationSeconds && durationSeconds > MAX_WHISPER_PREVIEW_SECONDS) return resolve([]);

    const audioBase = path.join(DOWNLOADS_DIR, `whisper_preview_${videoId}`);
    const audioPath = `${audioBase}.mp3`;
    const url = `https://www.youtube.com/watch?v=${videoId}`;
    // -f bestaudio pulls only the audio stream (not the full video) —
    // -x --audio-format mp3 then gives fetchWhisperCaptions() a plain file
    // it already knows how to handle exactly like a downloaded file.
    const command = `${quoteIfPath(ytdlpCmd)} -f bestaudio -x --audio-format mp3 ${cookiesFlag()} ${proxyFlag()} --paths "temp:${DOWNLOADS_DIR}" --output "${audioBase}.%(ext)s" "${url}"`;

    exec(command, { timeout: 90 * 1000, shell: true, cwd: DOWNLOADS_DIR }, async (error) => {
      if (error || !fs.existsSync(audioPath)) return resolve([]);
      const words = await fetchWhisperCaptions(audioPath, ffmpegCmd);
      try { fs.unlinkSync(audioPath); } catch (e) {}
      resolve(words);
    });
  });
}

// ─── MAIN DOWNLOAD ROUTE ──────────────────────────────────
// uid -> { child, cancelled } for whichever yt-dlp process is currently
// running on their behalf — lets /api/download/cancel actually kill it
// instead of only dropping the frontend's own fetch.
const activeDownloadProcesses = new Map();

app.post('/api/download', requireAuth, async (req, res) => {
  const { url, quality = '720p', type = 'video', format } = req.body;
  const ext = resolveDlFormat(type, format);
  console.log("Received download request:", { url, quality, type, format: ext });
  if (!url) return res.status(400).json({ error: 'URL missing hai' });
  if (!isSupportedUrl(url)) return res.status(400).json({ error: 'Yeh platform supported nahi hai' });

  const uid       = req.user.id;
  const timestamp = Date.now();
  const safeId    = `${uid.slice(0,8)}_${timestamp}`;
  const outFile   = path.join(DOWNLOADS_DIR, `${safeId}.${ext}`);
  const fmtFlag   = getYtdlpFormat(quality, type, ext);

  // Rate limit: ek user ek waqt mein ek hi download.
  // Self-healing: agar lock 6 minute se purana hai (jaise server restart ke
  // beech mein ek download atak gaya ho aur cleanup kabhi chala hi nahi), toh
  // use stale maan kar naya download allow karo — warna user hamesha ke liye
  // "already downloading" mein fas jaata.
  const activeFile = path.join(DOWNLOADS_DIR, `${uid.slice(0,8)}_active`);
  if (fs.existsSync(activeFile)) {
    const lockAge = Date.now() - fs.statSync(activeFile).mtimeMs;
    if (lockAge < 6 * 60 * 1000) {
      return res.status(429).json({ error: 'Aapka download pehle se chal raha hai. Thoda rukho' });
    }
    console.log(`[${uid.slice(0,8)}] Stale lock (${Math.round(lockAge/1000)}s old) mila, hata kar naya download shuru kar rahe hain`);
  }
  fs.writeFileSync(activeFile, timestamp.toString());

  // Tracked so /api/download/cancel can actually kill the yt-dlp process
  // (not just let the frontend drop its own fetch) — child gets filled in
  // once the exec() below actually starts.
  const procEntry = { child: null, cancelled: false };
  activeDownloadProcesses.set(uid, procEntry);

  const ytdlpCmd = await resolveYtdlpCommand();
  if (!ytdlpCmd) {
    try { fs.unlinkSync(activeFile); } catch {}
    activeDownloadProcesses.delete(uid);
    return res.status(500).json({
      error: 'yt-dlp install nahi hai. Install karo ya YTDLP_PATH set karo',
      detail: 'Try `winget install yt-dlp`, download yt-dlp.exe into this folder, or `pip install yt-dlp` then set YTDLP_PATH.'
    });
  }

  const commandPath = quoteIfPath(ytdlpCmd);
  const baseArgs = `${fmtFlag} ${cookiesFlag()} ${proxyFlag()} --paths "temp:${DOWNLOADS_DIR}" --no-playlist --max-filesize 200m --output "${outFile}" "${url}"`;
  const primaryCommand = `${commandPath} ${baseArgs}`;
  // Fallback if the default client hits YouTube's bot-check: retry once
  // with a broad list of alternate player clients — verified valid for this
  // yt-dlp build (visionos is currently yt-dlp's own preferred default,
  // usually a sign it's least-restricted at the moment) — plus
  // formats=missing_pot so a format isn't silently discarded just because
  // it lacks a proof-of-origin token. Not forced on every request — ios/
  // android alone have their own current issue (YouTube's SABR streaming
  // experiment drops some of their formats, verified directly), so this
  // wide combination is only a retry, never the default path.
  const retryCommand = `${commandPath} --extractor-args "youtube:player_client=visionos,tv_simply,ios,android,mweb,web_creator,web,tv;formats=missing_pot" ${baseArgs}`;

  console.log(`[${uid.slice(0,8)}] Using yt-dlp command: ${ytdlpCmd}`);
  console.log(`[${uid.slice(0,8)}] Downloading: ${url} | ${type} | ${quality}`);

  const handleResult = async (error, stdout, stderr) => {
    // /api/download/cancel already killed the process and told the frontend
    // (which also aborted its own fetch, so it won't read this) — but this
    // original request is still open server-side and must be finalized,
    // otherwise the connection just hangs forever instead of closing.
    if (procEntry.cancelled) {
      try { if (!res.headersSent) res.status(499).json({ error: 'Download cancel kar diya gaya', cancelled: true }); } catch (e) {}
      return;
    }
    if (error) {
      console.error('yt-dlp full error:', stderr || error.message, stdout);
      const lowerErr = (stderr || error.message || '').toLowerCase();
      const msg = lowerErr.includes('sign in to confirm')  ? 'YouTube bot-check lag gaya hai. Server par real cookies.txt set karo (COOKIES_PATH)' :
                  lowerErr.includes('429') || lowerErr.includes('too many requests')
                                                        ? 'YouTube rate limit lag gaya hai (bahut requests). Thodi der baad try karo' :
                  lowerErr.includes('private video')      ? 'Yeh private video hai' :
                  lowerErr.includes('not available') || lowerErr.includes('unavailable')
                                                        ? 'Video available nahi hai ya private hai' :
                  lowerErr.includes('max-filesize')       ? 'File bahut badi hai (200MB limit)' :
                  lowerErr.includes('unsupported url')    ? 'Yeh URL supported nahi hai' :
                  lowerErr.includes('access is forbidden') || lowerErr.includes('winerror 10013')
                                                        ? 'Network access blocked hai. Internet/firewall settings check karo' :
                  lowerErr.includes('not recognized')    ? 'yt-dlp install nahi hai ya PATH mein nahi hai' :
                  lowerErr.includes('enoent')            ? 'yt-dlp install nahi hai ya PATH mein nahi hai' :
                  'Download fail ho gaya. Dobara try karo';
      return res.status(500).json({ error: msg });
    }

    // Find actual output file (yt-dlp sometimes adds extra extension)
    let finalFile = outFile;
    if (!fs.existsSync(finalFile)) {
      const files = fs.readdirSync(DOWNLOADS_DIR).filter(f => f.startsWith(safeId));
      if (files.length) finalFile = path.join(DOWNLOADS_DIR, files[0]);
      else return res.status(500).json({ error: 'File create nahi hui. Retry karo' });
    }

    const filename = path.basename(finalFile);
    // Return a relative URL so localhost/127.0.0.1 (and local dev servers) cannot disagree.
    const fileUrl = `/files/${filename}`;

    // Get video title from stdout
    const titleMatch = stdout.match(/\[download\] (.+?) has already/i) ||
                       stdout.match(/Destination: (.+)/i);
    const title = titleMatch ? path.basename(titleMatch[1]) : `download.${ext}`;

    console.log(`[${uid.slice(0,8)}] Done: ${filename}`);

    // Keep a lightweight record in MongoDB for the signed-in user.
    try {
      await Download.create({ user: uid, url, filename, type, quality });
    } catch (dbError) {
      console.error('Could not save download history:', dbError);
    }
    logActivity(req, 'download', { url, type, quality, format: ext, filename });

    // Captions ab yahan block nahi karte — pehle inline hi fetch karte the
    // (auto-subs, phir na milne par whisper.cpp transcribe), jisme YouTube
    // ke alawa har platform (Instagram/TikTok/Twitter/Facebook/Vimeo/
    // Dailymotion/Reddit — inke paas apna caption data kabhi hota hi nahi)
    // whisper fallback tak pahunch jaata tha: ffmpeg extract (60s timeout)
    // + whisper transcribe (120s timeout), sab kuch download response se
    // PEHLE khatam hone ka intezaar karte hue — isi wajah se in platforms
    // par "download" bahut slow lagta tha, jabki asli file kabhi ki taiyaar
    // ho chuki hoti thi. Ab file ready hote hi turant respond karte hain;
    // frontend captions ko alag se /api/download-captions se, video already
    // dikhne ke baad, background mein maangta hai (search page mein bhi
    // yahi pattern already istemaal ho raha hai).
    res.json({
      success:  true,
      fileUrl,
      filename,
      title,
      type,
      quality,
      format: ext,
      captions: []
    });
  };

  // detached:true makes yt-dlp (and anything it spawns, like ffmpeg for
  // merging) the leader of its own process group on POSIX — that's what
  // lets /api/download/cancel kill the whole group (negative pid), not
  // just the top-level shell exec() actually launches. Harmless no-op-ish
  // on Windows local dev, where cancel falls back to killing just the
  // direct child instead.
  const primaryChild = exec(primaryCommand, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR, detached: true }, (error, stdout, stderr) => {
    // Clean active lock
    try { fs.unlinkSync(activeFile); } catch{}

    const hitBotCheck = error && (stderr || '').toLowerCase().includes('sign in to confirm');
    if (hitBotCheck && !procEntry.cancelled) {
      console.log(`[${uid.slice(0,8)}] Bot-check on default client, retrying with alternate player clients...`);
      const retryChild = exec(retryCommand, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR, detached: true }, (error2, stdout2, stderr2) => {
        activeDownloadProcesses.delete(uid);
        handleResult(error2, stdout2, stderr2);
      });
      procEntry.child = retryChild;
      return;
    }
    activeDownloadProcesses.delete(uid);
    handleResult(error, stdout, stderr);
  });
  procEntry.child = primaryChild;
});

// User apni current download cancel kar sake — sirf frontend fetch abort
// karna kaafi nahi, warna yt-dlp server par chalta rehta aur agla download
// stale lock (6 minute tak) ki wajah se atka rehta.
app.post('/api/download/cancel', requireAuth, (req, res) => {
  const uid = req.user.id;
  const entry = activeDownloadProcesses.get(uid);
  const activeFile = path.join(DOWNLOADS_DIR, `${uid.slice(0,8)}_active`);

  if (entry && entry.child && entry.child.pid) {
    entry.cancelled = true;
    if (process.platform === 'win32') {
      // On Windows, exec()'s child is cmd.exe /c <command> — killing just
      // that leaves the actual yt-dlp.exe (and any ffmpeg it spawns for
      // merging) running as orphans. taskkill /t walks the whole process
      // tree; /f forces it. Verified directly: without this, yt-dlp.exe
      // kept running (and the download response never returned) even
      // after entry.child.kill() reported success.
      execFile('taskkill', ['/pid', String(entry.child.pid), '/t', '/f'], () => {});
    } else {
      try {
        process.kill(-entry.child.pid, 'SIGKILL'); // whole process group
      } catch (e) {
        try { entry.child.kill('SIGKILL'); } catch (e2) {}
      }
    }
    activeDownloadProcesses.delete(uid);
  }

  try { fs.unlinkSync(activeFile); } catch (e) {}
  logActivity(req, 'download_cancelled', {});
  res.json({ success: true });
});

// Captions ko /api/download se decouple kar diya gaya (wahan blocking tha —
// dekho us route ke comment mein wajah). Frontend video dikhne ke baad,
// background mein, is route se caption maangta hai — best-effort, fail ho
// jaaye ya na milein toh bhi [] hi resolve hota hai.
app.post('/api/download-captions', requireAuth, async (req, res) => {
  const url      = String(req.body.url || '').trim();
  const filename = String(req.body.filename || '').trim();
  if (!url || !filename) return res.json({ captions: [] });

  const uid = req.user.id;
  const safeName = path.basename(filename); // path traversal guard
  if (!safeName.startsWith(`${uid.slice(0,8)}_`)) {
    return res.status(403).json({ error: 'Yeh file aapki nahi hai' });
  }
  const finalFile = path.join(DOWNLOADS_DIR, safeName);
  if (!fs.existsSync(finalFile)) return res.json({ captions: [] });

  const safeId   = safeName.replace(/\.[^/.]+$/, '');
  const ytdlpCmd = await resolveYtdlpCommand();
  let captions = ytdlpCmd ? await fetchAutoCaptions(url, ytdlpCmd, safeId) : [];
  if (!captions.length && WHISPER_READY) {
    const ffmpegCmd = await resolveFfmpegCommand();
    if (ffmpegCmd) captions = await fetchWhisperCaptions(finalFile, ffmpegCmd);
  }
  res.json({ captions });
});

// ─── CONVERT ROUTE — already-downloaded file ko doosre format mein badlo ──
app.post('/api/convert', requireAuth, async (req, res) => {
  const { filename, format } = req.body;
  if (!filename || !format) return res.status(400).json({ error: 'Filename aur format zaroori hain' });

  const uid = req.user.id;
  const safeName = path.basename(String(filename)); // path traversal guard
  if (!safeName.startsWith(`${uid.slice(0,8)}_`)) {
    return res.status(403).json({ error: 'Yeh file aapki nahi hai' });
  }

  const inputPath = path.join(DOWNLOADS_DIR, safeName);
  if (!fs.existsSync(inputPath)) return res.status(404).json({ error: 'File nahi mili. Pehle dobara download karo' });

  const targetFormat  = String(format).toLowerCase();
  const audioArgs     = AUDIO_CONVERT_FORMATS[targetFormat];
  const videoArgs     = VIDEO_CONVERT_FORMATS[targetFormat];
  const conversionArgs = audioArgs || videoArgs;
  if (!conversionArgs) return res.status(400).json({ error: 'Yeh format supported nahi hai' });

  const ffmpegCmd = await resolveFfmpegCommand();
  if (!ffmpegCmd) {
    return res.status(500).json({
      error: 'ffmpeg install nahi hai. Convert karne ke liye ffmpeg chahiye',
      detail: 'https://ffmpeg.org/download.html se install karo ya FFMPEG_PATH set karo.'
    });
  }

  const outName    = `${path.basename(safeName, path.extname(safeName))}_${Date.now()}.${targetFormat}`;
  const outputPath = path.join(DOWNLOADS_DIR, outName);
  const command     = `${quoteIfPath(ffmpegCmd)} -y -i "${inputPath}" ${conversionArgs} "${outputPath}"`;

  console.log(`[${uid.slice(0,8)}] Converting: ${safeName} -> ${targetFormat}`);

  exec(command, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error, stdout, stderr) => {
    if (error || !fs.existsSync(outputPath)) {
      console.error('ffmpeg convert error:', stderr || (error && error.message));
      return res.status(500).json({ error: 'Convert fail ho gaya. Dobara try karo' });
    }
    console.log(`[${uid.slice(0,8)}] Converted: ${outName}`);
    logActivity(req, 'convert', { sourceFilename: safeName, targetFormat, outName });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, format: targetFormat });
  });
});

// ─── PLAYGROUND QUERY — OpenRouter se streamed jawab ──────
// Ek waqt mein ek user ki ek hi query process hoti hai (jaise download wala rate limit).
const activeQueries = new Set();

function queryOpenRouter(model, content) {
  return fetch(`${OPENROUTER_BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
    },
    body: JSON.stringify({
      model,
      stream: true,
      messages: [
        { role: 'system', content: 'Tum PastePro Assistant ho. Seedha, sahi aur helpful jawab do. Jab Hindi/Hinglish mein poocha jaaye, usi mein jawab do.' },
        { role: 'user', content },
      ],
    }),
  });
}

const MAX_ATTACHMENT_BYTES     = 15 * 1024 * 1024; // base64 decoded size
const MAX_EXTRACTED_TEXT_CHARS = 15000; // itna hi prompt mein bhejte hain, poora document nahi — bahut lambi file se query hi itni badi ho jaati ki free model reject/timeout kar de
const SUPPORTED_DOC_FORMATS    = ['docx', 'xlsx', 'pptx', 'md', 'html', 'odt', 'rtf', 'epub', 'txt'];

// Zip-based Office/ODF/EPUB formats me sirf plain text nikaalne ke liye —
// koi bhi tag hata ke saadi text reh jaati hai, entities decode ho jaate
// hain. Layout/formatting kho jaata hai, par AI ke liye content kaafi hai.
function stripXmlTags(xml) {
  return xml
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n\n')
    .trim();
}

// RTF ek plain-text format hai jisme control words ({\rtf1... \par \b ...})
// hote hain — koi standalone parser add karne ke bajaye ek simple regex
// strip se kaam chal jaata hai (formatting kho jaati hai, text reh jaata hai).
function stripRtf(rtf) {
  return rtf
    .replace(/\\par[d]?/g, '\n')
    .replace(/\{\\[^}]*\}/g, '')
    .replace(/\\[a-zA-Z]+-?\d*\s?/g, '')
    .replace(/[{}]/g, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n\n')
    .trim();
}

function escapeHtmlServer(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Raw <a:t> text nikaalte waqt XML entities (&amp; &lt; wagera) decode
// nahi hote — isse pehle "&amp;" jaisa literal text hi reh jaata tha, aur
// baad mein client-side escapeHtml() usse dobara escape kar deta tha
// (&amp;amp; ban jaata). Yahan decode karke asli character wapas milta hai.
function decodeXmlEntities(s) {
  return String(s ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// Sab non-PDF document formats (Playground attachment) ke liye ek hi jagah
// se text extraction — har format ka apna tarika hai kyunki koi bhi free AI
// model in files ko "as-is" accept nahi karta, sabko plain text banana
// padta hai jo phir prompt mein context ki tarah fold hota hai. XLSX/PPTX
// ke liye ek extra "preview" shape (html/slides) bhi dete hain — browser
// inhe natively render nahi kar sakta, isliye attachment modal mein plain
// CSV/paragraph text ke bajaye asli table/slide jaisa dikhta hai.
async function extractDocText(buffer, format) {
  switch (format) {
    case 'docx': {
      const result = await mammoth.extractRawText({ buffer });
      return { text: (result.value || '').trim() };
    }
    case 'xlsx': {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      const text = wb.SheetNames.map(name => {
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
        return `--- Sheet: ${name} ---\n${csv}`.trim();
      }).join('\n\n').trim();
      // Apna khud ka escaping karte hain (SheetJS ke built-in sheet_to_html
      // par bharosa karne ke bajaye) taaki koi bhi cell value HTML mein
      // safely render ho, chahe usme "<"/"&" jaise characters hi kyun na hon.
      const html = wb.SheetNames.map(name => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' });
        const rowsHtml = rows.slice(0, 500).map(row =>
          `<tr>${row.map(cell => `<td>${escapeHtmlServer(cell)}</td>`).join('')}</tr>`
        ).join('');
        const truncNote = rows.length > 500 ? `<div class="pg-doc-sheet-note">(pehli 500 rows dikh rahi hain, sheet mein ${rows.length} hain)</div>` : '';
        return `<div class="pg-doc-sheet"><div class="pg-doc-sheet-name">${escapeHtmlServer(name)}</div><table>${rowsHtml}</table>${truncNote}</div>`;
      }).join('');
      return { text, html };
    }
    case 'pptx': {
      const zip = new AdmZip(buffer);
      const slideEntries = zip.getEntries()
        .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a, b) => parseInt(a.entryName.match(/(\d+)/)[1], 10) - parseInt(b.entryName.match(/(\d+)/)[1], 10));
      // Har <a:p> (paragraph/bullet) apni line par — pehle poore slide ke
      // <a:t> ko ek saath jod dete the jisse text bina line-breaks ke ek
      // hi paragraph mein jumble ho jaata tha, ab har bullet alag dikhta hai.
      const slides = slideEntries.map((e, i) => {
        const xml = e.getData().toString('utf8');
        const lines = [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
          .map(p => [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => decodeXmlEntities(m[1])).join(''))
          .filter(line => line.trim());
        return { number: i + 1, lines };
      });
      const text = slides.map(s => `--- Slide ${s.number} ---\n${s.lines.join('\n')}`).join('\n\n').trim();
      return { text, slides };
    }
    case 'odt': {
      const zip = new AdmZip(buffer);
      const entry = zip.getEntry('content.xml');
      return { text: entry ? stripXmlTags(entry.getData().toString('utf8')) : '' };
    }
    case 'epub': {
      const zip = new AdmZip(buffer);
      const htmlEntries = zip.getEntries().filter(e => /\.(x?html|htm)$/i.test(e.entryName));
      return { text: htmlEntries.map(e => stripXmlTags(e.getData().toString('utf8'))).join('\n\n').trim() };
    }
    case 'rtf':
      return { text: stripRtf(buffer.toString('utf8')) };
    case 'html':
      return { text: stripXmlTags(buffer.toString('utf8')) };
    case 'md':
    case 'txt':
      return { text: buffer.toString('utf8').trim() };
    default:
      return { text: '' };
  }
}

const MAX_ATTACHMENTS_PER_QUERY   = 5;
const MAX_TOTAL_ATTACHMENT_BYTES  = 20 * 1024 * 1024; // sab non-video attachments combined (base64 decoded)

app.post('/api/query', requireAuth, async (req, res) => {
  const query = String(req.body.query || '').trim();
  const attachments = Array.isArray(req.body.attachments) ? req.body.attachments : [];

  if (!query) return res.status(400).json({ error: 'Pehle kuch likho' });
  if (query.length > 4000) return res.status(400).json({ error: 'Query bahut lambi hai (max 4000 characters)' });
  if (!OPENROUTER_API_KEY) return res.status(500).json({ error: 'OpenRouter API key server par configure nahi hai' });
  if (attachments.length > MAX_ATTACHMENTS_PER_QUERY) {
    return res.status(400).json({ error: `Ek saath max ${MAX_ATTACHMENTS_PER_QUERY} files attach kar sakte ho` });
  }

  // Pehle hi sab attachments ko validate kar lete hain (format/size) taaki
  // koi bhi OpenRouter call ya text-extraction shuru hone se pehle hi galat
  // input reject ho jaaye.
  let totalBytes = 0;
  for (const att of attachments) {
    if (!att || typeof att !== 'object') return res.status(400).json({ error: 'Attachment data galat hai' });
    if (att.kind === 'image') {
      if (typeof att.dataUrl !== 'string' || !/^data:image\/(png|jpe?g|webp|gif);base64,/.test(att.dataUrl)) {
        return res.status(400).json({ error: 'Sirf PNG/JPEG/WEBP/GIF images support hain' });
      }
      const bytes = att.dataUrl.length * 0.75;
      if (bytes > MAX_ATTACHMENT_BYTES) return res.status(400).json({ error: 'Image bahut badi hai (max 15MB)' });
      totalBytes += bytes;
    } else if (att.kind === 'pdf') {
      if (typeof att.base64 !== 'string') return res.status(400).json({ error: 'PDF data galat hai' });
      const bytes = att.base64.length * 0.75;
      if (bytes > MAX_ATTACHMENT_BYTES) return res.status(400).json({ error: 'PDF bahut badi hai (max 15MB)' });
      totalBytes += bytes;
    } else if (att.kind === 'doc') {
      if (typeof att.base64 !== 'string' || !SUPPORTED_DOC_FORMATS.includes(att.format)) {
        return res.status(400).json({ error: 'Ye file format support nahi hai' });
      }
      const bytes = att.base64.length * 0.75;
      if (bytes > MAX_ATTACHMENT_BYTES) return res.status(400).json({ error: 'File bahut badi hai (max 15MB)' });
      totalBytes += bytes;
    } else if (att.kind !== 'video') {
      return res.status(400).json({ error: 'Attachment type galat hai' });
    }
  }
  if (totalBytes > MAX_TOTAL_ATTACHMENT_BYTES) {
    return res.status(400).json({ error: 'Sab files ka total size bahut zyada hai (max 20MB combined)' });
  }

  const uid = req.user.id;
  if (activeQueries.has(uid)) {
    return res.status(429).json({ error: 'Pehli query abhi process ho rahi hai, thoda ruko' });
  }
  activeQueries.add(uid);
  logActivity(req, 'playground_query', { query, attachmentCount: attachments.length, kinds: attachments.map(a => a.kind) });

  try {
    // Har attachment type ka apna extraction tarika hai kyunki koi bhi free
    // AI model in files ko "as-is" accept nahi karta — PDF/doc/video sab
    // plain text context ban ke prompt mein fold ho jaate hain, aur image
    // (multiple bhi) seedha vision model ko multimodal content ki tarah
    // jaate hain.
    const imageDataUrls = [];
    let extractedContext = '';

    for (const att of attachments) {
      if (att.kind === 'image') {
        imageDataUrls.push(att.dataUrl);
      } else if (att.kind === 'video') {
        // No free model here can actually watch a video, so the raw file
        // never leaves the browser (only kept in IndexedDB for playback)
        // — we just let the model know one was attached, by name.
        const name = String(att.name || 'video').slice(0, 200);
        extractedContext += `\n\n(User ne ek video file attach ki hai: "${name}". Video ka content dekhna abhi possible nahi hai, sirf filename pata hai — agar zaroori ho to user se pucho video mein kya hai.)`;
      } else if (att.kind === 'pdf') {
        const name = String(att.name || 'document.pdf').slice(0, 200);
        try {
          const buffer = Buffer.from(att.base64, 'base64');
          const parsed = await pdfParse(buffer);
          const extracted = (parsed.text || '').trim().slice(0, MAX_EXTRACTED_TEXT_CHARS);
          extractedContext += extracted
            ? `\n\nYeh ek PDF ("${name}") ka content hai:\n\n${extracted}`
            : `\n\n(PDF "${name}" se koi readable text nahi mila — shayad ye scanned/image-based PDF hai.)`;
        } catch (e) {
          console.error('PDF parse error:', e.message);
          activeQueries.delete(uid);
          return res.status(400).json({ error: `PDF "${name}" read nahi ho payi. Kya ye ek valid PDF file hai?` });
        }
      } else if (att.kind === 'doc') {
        const name = String(att.name || 'document').slice(0, 200);
        try {
          const buffer = Buffer.from(att.base64, 'base64');
          const { text } = await extractDocText(buffer, att.format);
          const extracted = (text || '').trim().slice(0, MAX_EXTRACTED_TEXT_CHARS);
          extractedContext += extracted
            ? `\n\nYeh ek ${att.format.toUpperCase()} file ("${name}") ka content hai:\n\n${extracted}`
            : `\n\n(File "${name}" se koi readable text nahi mila.)`;
        } catch (e) {
          console.error('Doc parse error:', e.message);
          activeQueries.delete(uid);
          return res.status(400).json({ error: `File "${name}" read nahi ho payi. Kya ye ek valid file hai?` });
        }
      }
    }

    const finalQuery = extractedContext
      ? `${extractedContext.trim()}\n\n---\n\nUpar wale content ke baare mein sawaal: ${query}`
      : query;

    // Image attachment(s): needs a vision-capable model — verified directly,
    // not every model listed as "vision-capable" on OpenRouter actually
    // handles image input correctly, so this is a fixed, separately-tested
    // model rather than the usual text model + fallback pair. Multiple
    // images can go in the same multimodal content array.
    const content = imageDataUrls.length
      ? [{ type: 'text', text: finalQuery }, ...imageDataUrls.map(url => ({ type: 'image_url', image_url: { url } }))]
      : finalQuery;
    const primaryModel = imageDataUrls.length ? OPENROUTER_VISION_MODEL : OPENROUTER_MODEL;

    let upstream = await queryOpenRouter(primaryModel, content);

    // Free OpenRouter models share a rate-limited pool that occasionally
    // gets busy (verified directly — some free models 429 consistently,
    // others only intermittently) — one retry with a different free model
    // covers most of those cases instead of failing the query outright.
    // No verified-working second vision model yet, so this retry only
    // applies to plain text queries.
    if (upstream.status === 429 && !imageDataUrls.length && OPENROUTER_FALLBACK_MODEL !== OPENROUTER_MODEL) {
      console.log('OpenRouter primary model rate-limited, retrying with fallback model...');
      upstream = await queryOpenRouter(OPENROUTER_FALLBACK_MODEL, content);
    }

    if (!upstream.ok || !upstream.body) {
      const errText = await upstream.text().catch(() => '');
      console.error('OpenRouter error:', upstream.status, errText);
      const msg = upstream.status === 401 ? 'OpenRouter API key invalid hai. .env mein OPENROUTER_API_KEY check karo' :
                  upstream.status === 403 ? 'OpenRouter account mein credits ya license nahi hai. openrouter.ai par billing check karo' :
                  upstream.status === 429 ? 'AI models abhi busy hain (free tier). Thodi der baad dobara try karo' :
                  `AI se jawab nahi mil paya (upstream HTTP ${upstream.status}). Dobara try karo.`;
      return res.status(upstream.status >= 400 && upstream.status < 500 ? upstream.status : 502).json({ error: msg });
    }

    // Ab yahan se stream shuru — client ko plain text chunks milte rahenge.
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache');

    const reader  = upstream.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let streamDone = false;

    // "[DONE]" khud hi authoritative signal hai ki stream khatam ho gayi —
    // isko dekhte hi turant rukna chahiye, upstream connection ke apne aap
    // band hone ka intezaar nahi karna chahiye. Verified directly: agar
    // sirf `if (done) break` par depend karo (jo reader.read() se aata hai,
    // matlab underlying connection band), toh OpenRouter/network kabhi-kabhi
    // "[DONE]" bhejne ke baad bhi connection turant close nahi karta —
    // isse poori request hamesha ke liye latak jaati thi (jawab poora aa
    // chuka hota tha lekin UI kabhi "complete" state par nahi pahunchta,
    // askBtn hamesha disabled reh jaata, aur session mein assistant ka
    // message kabhi save hi nahi hota).
    while (!streamDone) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop(); // incomplete line — agle chunk ke saath jodo

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;
        const payload = trimmed.slice(5).trim();
        if (payload === '[DONE]') { streamDone = true; break; }
        try {
          const token = JSON.parse(payload).choices?.[0]?.delta?.content;
          if (token) res.write(token);
        } catch (e) {
          // malformed chunk — ignore karo, stream continue rahega
        }
      }
    }
    try { reader.cancel(); } catch (e) {}
    res.end();
  } catch (error) {
    console.error('Query stream error:', error);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Kuch error aa gaya. Dobara try karo.' });
    } else {
      res.end();
    }
  } finally {
    activeQueries.delete(uid);
  }
});

// Browsers can't natively render DOCX/XLSX/PPTX/ODT/RTF/EPUB (unlike
// images/video/PDF/HTML, which iframe/img/video already handle) — this
// on-demand endpoint lets the Playground attachment modal show the same
// extracted plain text as a readable preview when the user clicks the chip.
// No OpenRouter call here, so it doesn't touch activeQueries/rate limits.
app.post('/api/extract-doc-text', requireAuth, async (req, res) => {
  const docBase64 = typeof req.body.docBase64 === 'string' ? req.body.docBase64 : null;
  const docFormat = typeof req.body.docFormat === 'string' ? req.body.docFormat : null;
  if (!docBase64 || !docFormat) return res.status(400).json({ error: 'File data missing' });
  if (!SUPPORTED_DOC_FORMATS.includes(docFormat)) return res.status(400).json({ error: 'Ye file format support nahi hai' });
  const approxBytes = docBase64.length * 0.75;
  if (approxBytes > MAX_ATTACHMENT_BYTES) return res.status(400).json({ error: 'File bahut badi hai (max 15MB)' });

  try {
    const buffer = Buffer.from(docBase64, 'base64');
    const result = await extractDocText(buffer, docFormat);
    res.json({
      text: (result.text || '').slice(0, 50000), // AI-context cap se zyada, preview ke liye — phir bhi bounded
      html: result.html,     // xlsx only — actual <table> markup for a real spreadsheet-like preview
      slides: result.slides, // pptx only — per-slide paragraph lines for a readable outline-style preview
    });
  } catch (e) {
    console.error('Doc preview extract error:', e.message);
    res.status(400).json({ error: 'File read nahi ho payi. Kya ye ek valid file hai?' });
  }
});

// ─── YOUTUBE SEARCH ────────────────────────────────────────
// execFile use kiya hai (exec + string-concat nahi) taaki query mein koi bhi
// character (", `, $, ; wagera) ho, woh shell syntax ki tarah interpret na ho —
// yeh command-injection se bachata hai. resolveYtdlpCommand() kabhi-kabhi
// compound string deta hai (jaise "py -m yt_dlp") — execFile (bina shell)
// ke liye use command + args mein todna padta hai.
function splitYtdlpCommand(ytdlpCmd) {
  const knownCompound = ['py -m yt_dlp', 'python -m yt_dlp', 'python3 -m yt_dlp'];
  if (knownCompound.includes(ytdlpCmd)) {
    const parts = ytdlpCmd.split(' ');
    return { cmd: parts[0], extraArgs: parts.slice(1) };
  }
  return { cmd: ytdlpCmd, extraArgs: [] }; // exe path (spaces allowed) ya plain 'yt-dlp'
}

app.post('/api/search', requireAuth, async (req, res) => {
  const query = String(req.body.query || '').trim();
  if (!query) return res.status(400).json({ error: 'Kuch search toh karo' });
  if (query.length > 200) return res.status(400).json({ error: 'Query bahut lambi hai' });

  const ytdlpCmd = await resolveYtdlpCommand();
  if (!ytdlpCmd) {
    return res.status(500).json({ error: 'yt-dlp install nahi hai. Install karo ya YTDLP_PATH set karo' });
  }

  const { cmd, extraArgs } = splitYtdlpCommand(ytdlpCmd);
  const baseArgs    = ['--flat-playlist', '--dump-json', ...cookiesArgs(), ...proxyArgs(), `ytsearch40:${query}`];
  const primaryArgs = [...extraArgs, ...baseArgs];
  // Same rationale as /api/download: cookies don't guarantee immunity from
  // YouTube's bot-check on a datacenter IP, so retry once with alternate
  // player clients if the default client gets blocked.
  const retryArgs   = [...extraArgs, '--extractor-args', 'youtube:player_client=visionos,tv_simply,ios,android,mweb,web_creator,web,tv;formats=missing_pot', ...baseArgs];

  const handleSearchResult = (error, stdout, stderr, res) => {
    if (error) {
      console.error('Search error:', stderr || error.message);
      return res.status(500).json({ error: 'Search fail ho gaya. Dobara try karo' });
    }

    const results = stdout.trim().split('\n').filter(Boolean).map(line => {
      try {
        const item = JSON.parse(line);
        const thumbs = item.thumbnails || [];
        const thumb  = thumbs.length ? thumbs[thumbs.length - 1].url : null;
        return {
          id: item.id,
          title: item.title,
          channel: item.channel || item.uploader || '',
          duration: item.duration_string || '',
          views: item.view_count || 0,
          thumbnail: thumb,
          url: item.webpage_url || `https://www.youtube.com/watch?v=${item.id}`
        };
      } catch (e) {
        return null;
      }
    }).filter(Boolean);

    logActivity(req, 'search', { query, resultsCount: results.length });
    res.json({ success: true, results });
  };

  execFile(cmd, primaryArgs, { timeout: 30 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
    const hitBotCheck = error && (stderr || '').toLowerCase().includes('sign in to confirm');
    if (hitBotCheck) {
      console.log('Search bot-check on default client, retrying with alternate player clients...');
      return execFile(cmd, retryArgs, { timeout: 30 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 10 * 1024 * 1024 }, (error2, stdout2, stderr2) => {
        handleSearchResult(error2, stdout2, stderr2, res);
      });
    }
    handleSearchResult(error, stdout, stderr, res);
  });
});

// ─── SEARCH PAGE — live word-by-word captions while previewing ───
// YouTube ke auto-captions do cheezein deti hain: original spoken language
// ka ASR track (metadata mein "<lang>-orig" key se marked) aur usi track ka
// ~100 languages mein auto-translate (jaise "en"). Verified directly (real
// Hindi-original + real English-original videos): dono tracks hi rolling-
// style VTT mein aate hain with inline per-word timestamps — matlab
// parseVttWords() (jo pehle se YouTube ke liye bana hua hai) dono par
// waisi hi word-level sync deta hai, translated track par bhi.
//
// Metadata (--dump-json) thoda bhaari hai, isliye per-video 5 minute cache
// — same video par language switch (orig <-> en) dobara yt-dlp nahi chalata.
const videoCaptionsMetaCache = new Map(); // videoId -> { data, expiresAt }
const VIDEO_META_CACHE_MS = 5 * 60 * 1000;

function fetchVideoCaptionsMeta(videoId) {
  const cached = videoCaptionsMetaCache.get(videoId);
  if (cached && cached.expiresAt > Date.now()) return Promise.resolve(cached.data);

  return resolveYtdlpCommand().then(ytdlpCmd => {
    if (!ytdlpCmd) throw new Error('yt-dlp install nahi hai');
    const { cmd, extraArgs } = splitYtdlpCommand(ytdlpCmd);
    const url = `https://www.youtube.com/watch?v=${videoId}`;
    const baseArgs  = ['--skip-download', '--dump-json', ...cookiesArgs(), ...proxyArgs(), url];
    const primaryArgs = [...extraArgs, ...baseArgs];
    const retryArgs   = [...extraArgs, '--extractor-args', 'youtube:player_client=visionos,tv_simply,ios,android,mweb,web_creator,web,tv;formats=missing_pot', ...baseArgs];

    return new Promise((resolve, reject) => {
      const handle = (error, stdout) => {
        if (error) return reject(error);
        try {
          const info = JSON.parse(stdout);
          const data = { automaticCaptions: info.automatic_captions || {}, duration: info.duration || 0 };
          videoCaptionsMetaCache.set(videoId, { data, expiresAt: Date.now() + VIDEO_META_CACHE_MS });
          resolve(data);
        } catch (e) {
          reject(e);
        }
      };
      execFile(cmd, primaryArgs, { timeout: 20 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 20 * 1024 * 1024 }, (error, stdout, stderr) => {
        const hitBotCheck = error && (stderr || '').toLowerCase().includes('sign in to confirm');
        if (hitBotCheck) {
          return execFile(cmd, retryArgs, { timeout: 20 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 20 * 1024 * 1024 }, (error2, stdout2) => handle(error2, stdout2));
        }
        handle(error, stdout);
      });
    });
  });
}

app.post('/api/search-captions', requireAuth, async (req, res) => {
  const videoId = String(req.body.videoId || '').trim();
  const lang     = String(req.body.lang || 'orig').trim();
  if (!/^[\w-]{5,20}$/.test(videoId)) return res.status(400).json({ error: 'Invalid video id' });

  try {
    const { automaticCaptions: autoCaps, duration } = await fetchVideoCaptionsMeta(videoId);
    const origKey      = Object.keys(autoCaps).find(k => k.endsWith('-orig'));
    const origLangCode = origKey ? origKey.replace(/-orig$/, '') : null;
    const origLangName = origKey
      ? (autoCaps[origKey][0]?.name || origLangCode || '').replace(/\s*\(Original\)\s*$/i, '')
      : null;

    let targetKey;
    if (lang === 'orig') {
      targetKey = origKey || (autoCaps.en ? 'en' : Object.keys(autoCaps)[0]);
    } else {
      targetKey = autoCaps[lang] ? lang : (autoCaps[`${lang}-orig`] ? `${lang}-orig` : null);
    }

    const activeLang = targetKey ? targetKey.replace(/-orig$/, '') : null;
    let captions = [];

    if (targetKey && autoCaps[targetKey]) {
      const vttEntry = autoCaps[targetKey].find(e => e.ext === 'vtt');
      if (vttEntry) {
        const vttRes = await fetch(vttEntry.url);
        if (vttRes.ok) captions = parseVttWords(await vttRes.text());
      }
      // Direct timedtext URL can occasionally get rate-limited by YouTube
      // itself (independent of yt-dlp's own bot-check, verified directly) —
      // fall back to fetching it fresh through yt-dlp instead of failing.
      if (!captions.length) {
        const ytdlpCmd = await resolveYtdlpCommand();
        if (ytdlpCmd) captions = await fetchCaptionsViaYtdlpLang(ytdlpCmd, videoId, activeLang);
      }
    }

    // YouTube ke paas is video ke liye koi caption data hai hi nahi (kayi
    // videos mein ye hota hai) — sirf tab khud transcribe karo jab user ne
    // "original" maanga ho: whisper original spoken language mein hi
    // transcribe karta hai, translate nahi kar sakta, isliye explicit "en"
    // request par isko chalane ka koi matlab nahi.
    let usedWhisper = false;
    if (!captions.length && lang === 'orig') {
      const ytdlpCmd  = await resolveYtdlpCommand();
      const ffmpegCmd = await resolveFfmpegCommand();
      if (ytdlpCmd && ffmpegCmd) {
        captions = await fetchWhisperCaptionsForSearch(videoId, duration, ytdlpCmd, ffmpegCmd);
        if (captions.length) usedWhisper = true;
      }
    }

    res.json({
      success: true,
      captions,
      origLang: usedWhisper ? null : origLangCode,
      origLangName: usedWhisper ? null : origLangName,
      activeLang: usedWhisper ? null : activeLang,
      source: usedWhisper ? 'whisper' : 'youtube'
    });
  } catch (error) {
    console.error('search-captions error:', error.message);
    res.status(500).json({ error: 'Captions load nahi ho paaye' });
  }
});

// ─── Page-visit tracking ───────────────────────────────────
// Yeh ek SPA hai — page navigation (search/docs/playground/category) sirf
// client-side hota hai, server ko pata nahi chalta. Frontend har navigation
// par yeh route call karta hai taaki "user kahan visit kar raha hai" bhi log ho.
app.post('/api/log-visit', requireAuth, async (req, res) => {
  const page = String(req.body.page || '').slice(0, 100);
  if (!page) return res.status(400).json({ error: 'page zaroori hai' });
  await logActivity(req, 'page_view', { page, extra: req.body.extra || undefined });
  res.json({ success: true });
});

// ─── Apna activity history dekhna ──────────────────────────
// Sirf apne hi logs — kisi aur user ka data yahan se nahi dikhta.
app.get('/api/activity', requireAuth, async (req, res) => {
  try {
    const limit = Math.min(parseInt(req.query.limit, 10) || 50, 200);
    const logs = await ActivityLog.find({ user: req.user.id })
      .sort({ createdAt: -1 })
      .limit(limit)
      .lean();
    res.json({ success: true, logs });
  } catch (error) {
    console.error('Activity fetch error:', error);
    res.status(500).json({ error: 'Activity load nahi ho saka' });
  }
});

// ─── Settings page ka account overview + usage stats ──────
// Ek hi jagah se: profile info, kitne downloads/searches/conversions/
// playground-queries hue, aur recent download history — sab existing
// Download/ActivityLog data se derive hota hai, koi nayi tracking nahi.
app.get('/api/account/stats', requireAuth, async (req, res) => {
  try {
    const uid = req.user.id;
    const user = await User.findById(uid).select('name email createdAt').lean();
    if (!user) return res.status(404).json({ error: 'User nahi mila' });

    const totalDownloads = await Download.countDocuments({ user: uid });

    const actionCounts = await ActivityLog.aggregate([
      { $match: { user: new mongoose.Types.ObjectId(uid) } },
      { $group: { _id: '$action', count: { $sum: 1 } } }
    ]);
    const counts = {};
    actionCounts.forEach(a => { counts[a._id] = a.count; });

    const recentDownloads = await Download.find({ user: uid })
      .sort({ createdAt: -1 })
      .limit(10)
      .select('url filename type quality createdAt')
      .lean();

    res.json({
      success: true,
      account: { name: user.name, email: user.email, memberSince: user.createdAt },
      stats: {
        totalDownloads,
        totalSearches:           counts.search            || 0,
        totalPlaygroundQueries:  counts.playground_query   || 0,
        totalConverts:           counts.convert            || 0,
        totalPageViews:          counts.page_view          || 0,
        totalCancelled:          counts.download_cancelled || 0,
        totalLogins:             counts.login              || 0,
      },
      recentDownloads
    });
  } catch (error) {
    console.error('Account stats error:', error);
    res.status(500).json({ error: 'Stats load nahi ho saka' });
  }
});

// ─── Health check ─────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// ─── Start server ─────────────────────────────────────────
async function startServer() {
  try {
    await mongoose.connect(MONGODB_URI);
    console.log(`MongoDB connected: ${MONGODB_URI}`);
  } catch (error) {
    console.error(`MongoDB connection failed (${MONGODB_URI}):`, error.message);
    process.exit(1);
  }

  app.listen(PORT, () => {
  console.log(`
  ╔══════════════════════════════════════╗
  ║   PastePro Backend chalu hai ✓       ║
  ║   http://localhost:${PORT}              ║
  ╚══════════════════════════════════════╝
  `);
  });
}

startServer();
