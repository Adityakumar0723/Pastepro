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
require('dotenv').config();

const app  = express();
// Port 3001 is occupied by VS Code's local webview service on this machine.
const PORT = process.env.PORT || 3002;

// ─── MongoDB ──────────────────────────────────────────────
// Override this in .env when needed. MongoDB must be running locally.
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/pastepro';
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-development-secret';

// ─── xAI / Grok (Playground Query section) ─────────────────
// Key sirf yahan, server-side, rehta hai — kabhi bhi frontend/index.html mein
// mat daalna, warna koi bhi visitor DevTools se churaake use kar sakta hai.
const XAI_API_KEY  = process.env.XAI_API_KEY || '';
const XAI_MODEL    = process.env.XAI_MODEL || 'grok-4-fast';
const XAI_BASE_URL = 'https://api.x.ai/v1';

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

const User = mongoose.model('User', userSchema);
const Download = mongoose.model('Download', downloadSchema);

// ─── Middleware ───────────────────────────────────────────
// No cookie credentials are used; JWTs are sent in the Authorization header.
// Allow the frontend both from a local dev server and when index.html is opened directly (Origin: null).
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json());

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

function fetchAutoCaptions(url, ytdlpCmd, safeId) {
  return new Promise((resolve) => {
    const subBase = path.join(DOWNLOADS_DIR, safeId);
    const command  = `${quoteIfPath(ytdlpCmd)} --write-auto-sub --sub-lang en --skip-download --sub-format vtt ${cookiesFlag()} ${proxyFlag()} --paths "temp:${DOWNLOADS_DIR}" --output "${subBase}.%(ext)s" "${url}"`;
    exec(command, { timeout: 30 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error) => {
      const vttPath = `${subBase}.en.vtt`;
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

// ─── MAIN DOWNLOAD ROUTE ──────────────────────────────────
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

  const ytdlpCmd = await resolveYtdlpCommand();
  if (!ytdlpCmd) {
    try { fs.unlinkSync(activeFile); } catch {}
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

    // Best-effort: YouTube ke auto-captions se word-by-word timed transcript.
    // Fail ho jaaye toh chup-chaap [] — caption na hona download ko fail nahi karta.
    const captions = await fetchAutoCaptions(url, ytdlpCmd, safeId);

    res.json({
      success:  true,
      fileUrl,
      filename,
      title,
      type,
      quality,
      format: ext,
      captions
    });
  };

  exec(primaryCommand, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error, stdout, stderr) => {
    // Clean active lock
    try { fs.unlinkSync(activeFile); } catch{}

    const hitBotCheck = error && (stderr || '').toLowerCase().includes('sign in to confirm');
    if (hitBotCheck) {
      console.log(`[${uid.slice(0,8)}] Bot-check on default client, retrying with alternate player clients...`);
      return exec(retryCommand, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR }, handleResult);
    }
    handleResult(error, stdout, stderr);
  });
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
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, format: targetFormat });
  });
});

// ─── PLAYGROUND QUERY — xAI/Grok se streamed jawab ────────
// Ek waqt mein ek user ki ek hi query process hoti hai (jaise download wala rate limit).
const activeQueries = new Set();

app.post('/api/query', requireAuth, async (req, res) => {
  const query = String(req.body.query || '').trim();
  if (!query) return res.status(400).json({ error: 'Pehle kuch likho' });
  if (query.length > 4000) return res.status(400).json({ error: 'Query bahut lambi hai (max 4000 characters)' });
  if (!XAI_API_KEY) return res.status(500).json({ error: 'xAI API key server par configure nahi hai' });

  const uid = req.user.id;
  if (activeQueries.has(uid)) {
    return res.status(429).json({ error: 'Pehli query abhi process ho rahi hai, thoda ruko' });
  }
  activeQueries.add(uid);

  try {
    const upstream = await fetch(`${XAI_BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${XAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: XAI_MODEL,
        stream: true,
        messages: [
          { role: 'system', content: 'Tum PastePro Assistant ho. Seedha, sahi aur helpful jawab do. Jab Hindi/Hinglish mein poocha jaaye, usi mein jawab do.' },
          { role: 'user', content: query },
        ],
      }),
    });

    if (!upstream.ok || !upstream.body) {
      const errText = await upstream.text().catch(() => '');
      console.error('xAI error:', upstream.status, errText);
      const msg = upstream.status === 401 ? 'xAI API key invalid hai. .env mein XAI_API_KEY check karo' :
                  upstream.status === 403 ? 'xAI account/team mein credits ya license nahi hai. console.x.ai par billing add karo' :
                  upstream.status === 429 ? 'Rate limit lag gaya hai. Thoda ruk kar dobara try karo' :
                  `AI se jawab nahi mil paya (upstream HTTP ${upstream.status}). Dobara try karo.`;
      return res.status(upstream.status >= 400 && upstream.status < 500 ? upstream.status : 502).json({ error: msg });
    }

    // Ab yahan se stream shuru — client ko plain text chunks milte rahenge.
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache');

    const reader  = upstream.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop(); // incomplete line — agle chunk ke saath jodo

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;
        const payload = trimmed.slice(5).trim();
        if (payload === '[DONE]') continue;
        try {
          const token = JSON.parse(payload).choices?.[0]?.delta?.content;
          if (token) res.write(token);
        } catch (e) {
          // malformed chunk — ignore karo, stream continue rahega
        }
      }
    }
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
  const baseArgs    = ['--flat-playlist', '--dump-json', ...cookiesArgs(), ...proxyArgs(), `ytsearch12:${query}`];
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
