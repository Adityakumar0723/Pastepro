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
const crypto    = require('crypto');
const nodemailer = require('nodemailer');
const pdfParse  = require('pdf-parse');
const mammoth   = require('mammoth');
const XLSX      = require('xlsx');
const AdmZip    = require('adm-zip');
const multer    = require('multer');
const PDFDocument = require('pdfkit');
const { Document: DocxDocument, Packer: DocxPacker, Paragraph: DocxParagraph, HeadingLevel: DocxHeadingLevel, TextRun: DocxTextRun, ImageRun: DocxImageRun } = require('docx');
const PptxGenJS = require('pptxgenjs');
const { imageSize } = require('image-size');
const sharp = require('sharp');
const { PDFDocument: PdfLibDocument, StandardFonts: PdfLibStandardFonts, rgb: pdfLibRgb, degrees: pdfLibDegrees } = require('@cantoo/pdf-lib');
const { chromium } = require('playwright');
const { diffLines } = require('diff');
const dns = require('dns').promises;
const net = require('net');
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
  password: { type: String, required: true },
  // Forgot-password: sirf hashed token store hota hai (raw token kabhi nahi,
  // password ki tarah hi) — email mein jo link jaata hai wahi raw token
  // le jaata hai, yahan sirf uska SHA-256 hash match karne ke liye rakha hai.
  resetPasswordTokenHash: { type: String, default: null },
  resetPasswordExpires:   { type: Date,   default: null }
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

// PDF Tools "Workflow" — user kai PDF tools ko ek saved chain (steps array)
// mein jod deta hai (jaise iLovePDF), phir kisi bhi file par ek click mein
// pura chain run kar sakta hai. options ka exact shape har tool ke hisaab
// se alag hota hai isliye Mixed hai — WORKFLOW_TOOLS registry hi validate
// karta hai ki kaunse fields valid hain.
const workflowSchema = new mongoose.Schema({
  user: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  name: { type: String, required: true, trim: true, maxlength: 80 },
  steps: [{
    tool: { type: String, required: true },
    options: { type: mongoose.Schema.Types.Mixed, default: {} },
  }],
}, { timestamps: true });

const User = mongoose.model('User', userSchema);
const Download = mongoose.model('Download', downloadSchema);
const ActivityLog = mongoose.model('ActivityLog', activityLogSchema);
const Workflow = mongoose.model('Workflow', workflowSchema);

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

// Documentation page screenshots — plain static files, no auth needed.
app.use('/docs-assets', express.static(path.join(__dirname, 'docs-assets')));

// Static serve for downloaded files
app.use('/files', (req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Range');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Range, Content-Length, Accept-Ranges');
  // Only force a Save-As dialog for explicit downloads (?dl=1) — a bare
  // request (e.g. from a <video>/<audio> preview player) should play inline.
  // Crucially, only when the file actually exists: setting this header on a
  // 404 (file expired/deleted, e.g. after a server restart wiped the tmp
  // downloads dir) made Chrome treat the tiny 404 HTML page as a broken
  // "attachment" download during a full-page navigation, surfacing as
  // ERR_INVALID_RESPONSE instead of a normal 404 page.
  if (req.query.dl) {
    const resolved = path.normalize(path.join(DOWNLOADS_DIR, decodeURIComponent(req.path)));
    if (resolved.startsWith(DOWNLOADS_DIR) && fs.existsSync(resolved) && fs.statSync(resolved).isFile()) {
      res.setHeader('Content-Disposition', 'attachment');
    }
  }
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

const RESET_TOKEN_TTL_MS = 10 * 60 * 1000; // 10 minute — jitna user ne maanga
const APP_URL = (process.env.APP_URL || `http://localhost:${PORT}`).replace(/\/$/, '');

const mailTransporter = (process.env.GMAIL_USER && process.env.GMAIL_APP_PASSWORD)
  ? nodemailer.createTransport({
      service: 'gmail',
      auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_APP_PASSWORD },
    })
  : null;

async function sendPasswordResetEmail(toEmail, resetUrl) {
  if (!mailTransporter) {
    // Local dev / credentials na set hone par — link console mein taaki
    // pura flow bina real email ke bhi test ho sake.
    console.log(`[password-reset] GMAIL_USER/GMAIL_APP_PASSWORD set nahi hain — link: ${resetUrl}`);
    return;
  }
  await mailTransporter.sendMail({
    from: `"PastePro" <${process.env.GMAIL_USER}>`,
    to: toEmail,
    subject: 'PastePro — Password Reset Karo',
    html: `
      <div style="font-family:sans-serif; max-width:480px; margin:0 auto;">
        <h2 style="color:#e63946;">PastePro Password Reset</h2>
        <p>Aapne apna PastePro password reset karne ke liye request kiya hai. Neeche wale button se naya password set karo:</p>
        <p style="margin:24px 0;">
          <a href="${resetUrl}" style="background:linear-gradient(135deg,#e63946,#f4a261); color:#fff; padding:12px 24px; border-radius:8px; text-decoration:none; font-weight:bold;">Password Reset Karo</a>
        </p>
        <p style="color:#666; font-size:13px;">Ye link sirf <strong>10 minute</strong> ke liye valid hai. Agar aapne ye request nahi kiya, toh is email ko ignore kar do — aapka password nahi badlega.</p>
        <p style="color:#999; font-size:12px;">Agar button kaam na kare, ye link copy karke browser mein paste karo:<br>${resetUrl}</p>
      </div>
    `,
  });
}

app.post('/api/auth/forgot-password', async (req, res) => {
  try {
    const email = String(req.body.email || '').trim().toLowerCase();
    if (!/^\S+@\S+\.\S+$/.test(email)) return res.status(400).json({ error: 'Valid email daalo' });

    // User exist karta hai ya nahi — kisi bhi case mein response same rehta
    // hai (generic success), warna ye endpoint attacker ko bata deta ki
    // koi email PastePro par registered hai ya nahi (user enumeration).
    const user = await User.findOne({ email });
    if (user) {
      const rawToken  = crypto.randomBytes(32).toString('hex');
      const tokenHash = crypto.createHash('sha256').update(rawToken).digest('hex');
      user.resetPasswordTokenHash = tokenHash;
      user.resetPasswordExpires   = new Date(Date.now() + RESET_TOKEN_TTL_MS);
      await user.save();

      const resetUrl = `${APP_URL}/?resetToken=${rawToken}`;
      try {
        await sendPasswordResetEmail(user.email, resetUrl);
      } catch (mailErr) {
        console.error('Password reset email send failed:', mailErr.message);
        // Email fail hone par bhi client ko generic success hi milta hai —
        // par server logs mein asli wajah dikhti hai diagnose karne ke liye.
      }
      req.user = { id: user._id.toString() };
      logActivity(req, 'forgot_password_requested', { email: user.email });
    }

    res.json({ success: true, message: 'Agar ye email registered hai, reset link bhej diya gaya hai' });
  } catch (error) {
    console.error('Forgot-password error:', error);
    res.status(500).json({ error: 'Kuch error aa gaya. Dobara try karo' });
  }
});

app.post('/api/auth/reset-password', async (req, res) => {
  try {
    const token       = String(req.body.token || '').trim();
    const newPassword = String(req.body.newPassword || '');
    if (!token) return res.status(400).json({ error: 'Reset link invalid hai' });
    if (newPassword.length < 6) return res.status(400).json({ error: 'Password 6+ characters ka hona chahiye' });

    const tokenHash = crypto.createHash('sha256').update(token).digest('hex');
    const user = await User.findOne({
      resetPasswordTokenHash: tokenHash,
      resetPasswordExpires: { $gt: new Date() },
    });
    if (!user) return res.status(400).json({ error: 'Ye link invalid ya expire ho chuka hai (10 minute ke baad expire hota hai) — dobara reset request karo' });

    user.password = await bcrypt.hash(newPassword, 12);
    user.resetPasswordTokenHash = null;
    user.resetPasswordExpires   = null;
    await user.save();

    req.user = { id: user._id.toString() };
    logActivity(req, 'password_reset_completed', { email: user.email });
    res.json({ success: true });
  } catch (error) {
    console.error('Reset-password error:', error);
    res.status(500).json({ error: 'Password reset nahi ho saka. Dobara try karo' });
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
    /reddit\.com/,
    /loom\.com/
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
    exec(command, { shell: true, timeout: 10000 }, (err, stdout, stderr) => {
      if (err) console.log(`[yt-dlp resolve] "${candidate}" failed: ${(stderr || err.message || '').trim().slice(0, 200)}`);
      resolve(!err);
    });
  });
}

let ytdlpResolveLoggedOnce = false;

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
  // Har request par dobara na chhape (search/download dono is function ko
  // baar-baar call karte hain) — sirf pehli baar, taaki logs mein spam na ho,
  // par diagnosis ke liye PATH aur candidate list saaf dikh jaaye.
  if (!ytdlpResolveLoggedOnce) {
    ytdlpResolveLoggedOnce = true;
    console.error(`[yt-dlp resolve] koi bhi candidate kaam nahi kiya. Tried: ${JSON.stringify(candidates)} | PATH=${process.env.PATH}`);
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

// ─── ffprobe resolve (Video Editor mein fade/trim-duration maths ke liye) ──
async function resolveFfprobeCommand() {
  const localExe = path.join(__dirname, 'ffprobe.exe');
  const candidates = [];

  if (process.env.FFPROBE_PATH) candidates.push(process.env.FFPROBE_PATH);
  if (fs.existsSync(localExe)) candidates.push(localExe);
  candidates.push('ffprobe');

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

function getMediaDuration(ffprobeCmd, filePath) {
  return new Promise((resolve) => {
    const cmd = `${quoteIfPath(ffprobeCmd)} -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "${filePath}"`;
    exec(cmd, { shell: true, timeout: 15000 }, (error, stdout) => {
      const val = parseFloat(stdout);
      resolve(!error && Number.isFinite(val) ? val : null);
    });
  });
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

// Fetches a video's own auto-sub track for an arbitrary language code —
// used as a fallback when the direct timedtext URL (from --dump-json metadata)
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

// Download page's transcript box: try the platform's own auto-sub track
// first (YouTube almost always has one; most other platforms don't expose
// any via yt-dlp at all) before falling back to whisper below.
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

// Search page's fallback when a video has no real caption data of its own:
// transcribe its audio ourselves with a local whisper.cpp model (called via
// fetchWhisperCaptionsForSearch below, on just the audio track — no full
// video download needed for a preview). Also reused directly by the
// download page's transcript box (below) on the already-downloaded file,
// for every platform that doesn't expose caption data of its own — the
// multilingual tiny model transcribes in whatever language is actually
// spoken, it doesn't translate. `-ml 2` caps each VTT cue at ~2 words so
// the existing parseVttWords() (built for YouTube's word-timed cues) gets
// near-word-level timing here too, instead of one giant per-sentence cue.
// Bounded and best-effort: any failure (missing binary, ffmpeg hiccup,
// timeout) just resolves to [] rather than breaking the preview/download.
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

    exec(extractCmd, { timeout: 40 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (ffErr, ffStdout, ffStderr) => {
      if (ffErr || !fs.existsSync(wavPath)) {
        console.error(`[whisper] ffmpeg audio-extract failed for ${mediaPath}:`, (ffStderr || ffErr?.message || '').slice(0, 300));
        return resolve([]);
      }

      // 180s (not the earlier 90s) — a free-tier/shared-CPU server can take
      // a while on even a few minutes of audio once you add the tiny
      // model's own load time on every invocation (no persistent process),
      // and real videos are routinely 3-5+ minutes long. The download
      // page's client-side abort (fetchAndRenderDownloadCaptions) waits up
      // to 240s specifically to stay comfortably above this + ffmpeg's own
      // 40s, so a real timeout here still reaches the client as a normal
      // "not available" response instead of racing the client's own abort.
      const whisperCmd = `${quoteIfPath(WHISPER_BIN)} -m "${WHISPER_MODEL}" -f "${wavPath}" -ml 2 -ovtt -of "${base}" -np`;
      exec(whisperCmd, { timeout: 180 * 1000, shell: true, cwd: DOWNLOADS_DIR, env: { ...process.env, LD_LIBRARY_PATH: path.dirname(WHISPER_BIN) } }, (wErr, wStdout, wStderr) => {
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

// uid -> { id, url, type, quality, percent, status, error, result, startedAt }
// Tracked separately from the main /api/download request/response cycle so
// the bell-notification UI can poll /api/download/progress for live percent
// even if the ORIGINAL long-lived request's connection drops (a real network
// hiccup on the client doesn't abort the server-side yt-dlp process) — the
// frontend keeps polling this independently and still sees it through to
// completion (or lets the user cancel it) instead of the notification being
// stuck showing "in progress" forever with no way to resolve it.
const downloadProgress = new Map();

// yt-dlp prints progress like "[download]  45.2% of 10.00MiB at 1.2MiB/s"
// with \r (not \n) between updates in a real terminal — a raw stdout chunk
// can contain several of these squashed together, so take the LAST match.
function parseDownloadPercent(chunk) {
  const matches = [...chunk.matchAll(/\[download\]\s+(\d+(?:\.\d+)?)%/g)];
  return matches.length ? parseFloat(matches[matches.length - 1][1]) : null;
}

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
  downloadProgress.set(uid, { id: safeId, url, type, quality, percent: 0, status: 'downloading', startedAt: timestamp });

  const ytdlpCmd = await resolveYtdlpCommand();
  if (!ytdlpCmd) {
    try { fs.unlinkSync(activeFile); } catch {}
    activeDownloadProcesses.delete(uid);
    downloadProgress.set(uid, { id: safeId, url, type, quality, percent: 0, status: 'error', error: 'yt-dlp install nahi hai. Install karo ya YTDLP_PATH set karo', startedAt: timestamp });
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
      downloadProgress.set(uid, { id: safeId, url, type, quality, percent: 0, status: 'error', error: msg, startedAt: timestamp });
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
    // Captions/transcription yahan se hata di gayi — frontend ab download
    // page par khud ek live audio waveform dikhata hai (Web Audio API se,
    // koi fetch/server round-trip nahi), isliye file ready hote hi turant
    // respond kar dete hain, kisi bhi extra processing ka wait kiye bina.
    const resultPayload = { success: true, fileUrl, filename, title, type, quality, format: ext };
    // Bell-notification polling (/api/download/progress) ke liye — agar
    // original request ka connection beech mein hi toot gaya ho (network
    // hiccup), tab bhi polling se yahi final result mil jaata hai.
    downloadProgress.set(uid, { id: safeId, url, type, quality, percent: 100, status: 'done', result: resultPayload, startedAt: timestamp });
    res.json(resultPayload);
  };

  // detached:true makes yt-dlp (and anything it spawns, like ffmpeg for
  // merging) the leader of its own process group on POSIX — that's what
  // lets /api/download/cancel kill the whole group (negative pid), not
  // just the top-level shell exec() actually launches. Harmless no-op-ish
  // on Windows local dev, where cancel falls back to killing just the
  // direct child instead.
  // exec() ka returned ChildProcess bhi ek real stream hai (spawn() jaisa
  // hi) — is par apna khud ka 'data' listener laga sakte hain live percent
  // parse karne ke liye, exec() ke apne internal buffering (jo callback ko
  // poora stdout deta hai) se bilkul alag/independent.
  function trackPercent(child) {
    child.stdout?.on('data', (chunk) => {
      const percent = parseDownloadPercent(chunk.toString());
      if (percent !== null) {
        const entry = downloadProgress.get(uid);
        if (entry && entry.status === 'downloading') entry.percent = percent;
      }
    });
  }

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
      trackPercent(retryChild);
      procEntry.child = retryChild;
      return;
    }
    activeDownloadProcesses.delete(uid);
    handleResult(error, stdout, stderr);
  });
  trackPercent(primaryChild);
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
  const prevProgress = downloadProgress.get(uid);
  if (prevProgress) downloadProgress.set(uid, { ...prevProgress, status: 'cancelled' });
  logActivity(req, 'download_cancelled', {});
  res.json({ success: true });
});

// Bell-notification dropdown polls this — independent of the main
// /api/download request, so it still reflects the true status even if that
// original connection broke (client-side network hiccup) partway through.
app.get('/api/download/progress', requireAuth, (req, res) => {
  const entry = downloadProgress.get(req.user.id);
  res.json(entry || { status: 'none' });
});

// Download page's transcript box: file ready hote hi /api/download turant
// respond kar deta hai (waveform bhi turant dikhta hai — koi fetch nahi),
// aur ye route alag se, background mein, asli word-by-word text laata hai
// — best-effort, fail ho jaaye ya na milein toh bhi [] hi resolve hota hai.
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

// ─── VIDEO EDITOR — upload apni video, trim/noise-reduce/volume/speed ────
const MAX_EDIT_UPLOAD_BYTES = 300 * 1024 * 1024; // 300MB

const editUploadStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, DOWNLOADS_DIR),
  // Download flow ke jaisa hi naming (uid prefix) — taaki /api/edit/process
  // aur /files/ dono jagah wahi ownership check (startsWith uid) kaam kare.
  filename: (req, file, cb) => {
    const uid = req.user.id;
    const ext = (path.extname(file.originalname) || '.mp4').toLowerCase();
    cb(null, `${uid.slice(0,8)}_${Date.now()}_upload${ext}`);
  },
});
const editUpload = multer({
  storage: editUploadStorage,
  limits: { fileSize: MAX_EDIT_UPLOAD_BYTES },
  fileFilter: (req, file, cb) => cb(null, file.mimetype.startsWith('video/')),
});

app.post('/api/edit/upload', requireAuth, (req, res) => {
  editUpload.single('video')(req, res, (err) => {
    if (err instanceof multer.MulterError && err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File bahut badi hai (max 300MB)' });
    }
    if (err) return res.status(400).json({ error: 'Upload fail ho gaya. Sirf video files chalti hain' });
    if (!req.file) return res.status(400).json({ error: 'Koi video file nahi mili' });

    const uid = req.user.id;
    logActivity(req, 'edit_upload', { filename: req.file.filename, sizeMB: Math.round(req.file.size / 1024 / 1024) });
    res.json({ success: true, filename: req.file.filename, fileUrl: `/files/${req.file.filename}` });
  });
});

// ─── MEDIA CONVERTER — koi bhi video ya audio upload karke seedha
//     /api/convert (upar wala, download-flow wala hi) se kisi bhi format
//     mein badal sakte hain — same uid-prefixed naming isliye rakha hai
//     taaki /api/convert ka ownership check bina kisi badlaav ke chal jaaye. ──
const mediaConvertUploadStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, DOWNLOADS_DIR),
  filename: (req, file, cb) => {
    const uid = req.user.id;
    const ext = (path.extname(file.originalname) || '.mp4').toLowerCase();
    cb(null, `${uid.slice(0,8)}_${Date.now()}_mc${ext}`);
  },
});
const mediaConvertUpload = multer({
  storage: mediaConvertUploadStorage,
  limits: { fileSize: MAX_EDIT_UPLOAD_BYTES },
  fileFilter: (req, file, cb) => cb(null, file.mimetype.startsWith('video/') || file.mimetype.startsWith('audio/')),
});

app.post('/api/media-convert/upload', requireAuth, (req, res) => {
  mediaConvertUpload.single('media')(req, res, (err) => {
    if (err instanceof multer.MulterError && err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File bahut badi hai (max 300MB)' });
    }
    if (err) return res.status(400).json({ error: 'Upload fail ho gaya. Sirf video/audio files chalti hain' });
    if (!req.file) return res.status(400).json({ error: 'Koi file nahi mili' });

    const type = req.file.mimetype.startsWith('audio/') ? 'audio' : 'video';
    logActivity(req, 'media_convert_upload', { filename: req.file.filename, sizeMB: Math.round(req.file.size / 1024 / 1024) });
    res.json({ success: true, filename: req.file.filename, fileUrl: `/files/${req.file.filename}`, type });
  });
});

// atempo sirf 0.5x-2x range hi accept karta hai (isse zyada/kam ke liye
// chain karna padta), aur speed dropdown isi range tak limited hai, isliye
// yahan seedha ek hi atempo instance kaafi hai.
const EDIT_ALLOWED_SPEEDS = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2];
const EDIT_ALLOWED_EFFECTS = ['none', 'grayscale', 'sepia', 'vignette', 'blur', 'sharpen', 'invert', 'warm', 'cool', 'highcontrast', 'fade', 'oldfilm'];
const EDIT_EFFECT_FILTERS = {
  grayscale: 'hue=s=0',
  sepia: 'colorchannelmixer=.393:.769:.189:0:.349:.686:.168:0:.272:.534:.131',
  vignette: 'vignette',
  blur: 'boxblur=4:1',
  sharpen: 'unsharp=5:5:1.0:5:5:0.0',
  invert: 'negate',
  warm: 'colorbalance=rs=0.2:gs=0.05:bs=-0.15',
  cool: 'colorbalance=rs=-0.15:gs=0:bs=0.2',
  highcontrast: 'eq=contrast=1.6:saturation=1.15',
  fade: 'eq=contrast=0.85:brightness=0.08:saturation=0.85',
  oldfilm: 'hue=s=0,eq=contrast=1.3,vignette',
};
const EDIT_ALLOWED_FONTS = { sans: 'sans-serif', serif: 'serif', mono: 'monospace' };
const EDIT_MAX_TEXT_LEN = 300;

// Filter-graph string mein path daalne se pehle: backslash -> forward slash
// (Windows dev par C:\Users\... jaisa path bhi chal jaaye) aur ':' escape
// karo (filter syntax mein ':' key=value separator hai, isliye drive-letter
// wale colon ko '\:' banana zaroori hai — Linux/prod paths mein colon hota
// hi nahi to wahan ye no-op rehta hai).
function escapeFfmpegFilterPath(p) {
  return p.replace(/\\/g, '/').replace(/:/g, '\\:');
}

app.post('/api/edit/process', requireAuth, async (req, res) => {
  const {
    filename, trimStart, trimEnd, noiseReduction, volume, speed,
    effect, brightness, contrast, saturation, fadeIn, fadeOut,
    textEnabled, text, textFont, textSize, textColor, textX, textY, textBox,
  } = req.body;
  if (!filename) return res.status(400).json({ error: 'Filename zaroori hai' });

  const uid = req.user.id;
  const safeName = path.basename(String(filename)); // path traversal guard
  if (!safeName.startsWith(`${uid.slice(0,8)}_`)) {
    return res.status(403).json({ error: 'Yeh file aapki nahi hai' });
  }
  const inputPath = path.join(DOWNLOADS_DIR, safeName);
  if (!fs.existsSync(inputPath)) return res.status(404).json({ error: 'File nahi mili. Pehle dobara upload karo' });

  const start = trimStart !== undefined && trimStart !== null && trimStart !== '' ? Number(trimStart) : null;
  const end   = trimEnd   !== undefined && trimEnd   !== null && trimEnd   !== '' ? Number(trimEnd)   : null;
  if (start !== null && (!Number.isFinite(start) || start < 0)) return res.status(400).json({ error: 'Trim start invalid hai' });
  if (end !== null && (!Number.isFinite(end) || end <= (start || 0))) return res.status(400).json({ error: 'Trim end invalid hai (start se bada hona chahiye)' });

  const vol = volume !== undefined && volume !== null && volume !== '' ? Number(volume) : 1;
  if (!Number.isFinite(vol) || vol <= 0 || vol > 5) return res.status(400).json({ error: 'Volume invalid hai (0-5 ke beech)' });

  const spd = speed !== undefined && speed !== null && speed !== '' ? Number(speed) : 1;
  if (!EDIT_ALLOWED_SPEEDS.includes(spd)) return res.status(400).json({ error: 'Speed invalid hai' });

  const denoise = noiseReduction === true || noiseReduction === 'true';
  const doFadeIn = fadeIn === true || fadeIn === 'true';
  const doFadeOut = fadeOut === true || fadeOut === 'true';

  const fx = effect && EDIT_ALLOWED_EFFECTS.includes(effect) ? effect : 'none';

  const bright = brightness !== undefined && brightness !== null && brightness !== '' ? Number(brightness) : 0;
  const contr  = contrast   !== undefined && contrast   !== null && contrast   !== '' ? Number(contrast)   : 1;
  const satur  = saturation !== undefined && saturation !== null && saturation !== '' ? Number(saturation) : 1;
  if (!Number.isFinite(bright) || bright < -1 || bright > 1) return res.status(400).json({ error: 'Brightness invalid hai (-1 se 1 ke beech)' });
  if (!Number.isFinite(contr) || contr < 0 || contr > 3) return res.status(400).json({ error: 'Contrast invalid hai (0-3 ke beech)' });
  if (!Number.isFinite(satur) || satur < 0 || satur > 3) return res.status(400).json({ error: 'Saturation invalid hai (0-3 ke beech)' });

  const doText = textEnabled === true || textEnabled === 'true';
  const textStr = typeof text === 'string' ? text : '';
  if (doText && !textStr.trim()) return res.status(400).json({ error: 'Text daalo ya Text overlay disable karo' });
  if (doText && textStr.length > EDIT_MAX_TEXT_LEN) return res.status(400).json({ error: `Text bahut lamba hai (max ${EDIT_MAX_TEXT_LEN} characters)` });
  const fontKey = EDIT_ALLOWED_FONTS[textFont] ? textFont : 'sans';
  const fSize = textSize !== undefined && textSize !== null && textSize !== '' ? Number(textSize) : 36;
  if (!Number.isFinite(fSize) || fSize < 12 || fSize > 160) return res.status(400).json({ error: 'Font size invalid hai (12-160 ke beech)' });
  const colorHex = /^#[0-9a-fA-F]{6}$/.test(textColor || '') ? textColor : '#ffffff';
  // Text ab kisi bhi (x%, y%) par drag karke rakha ja sakta hai — preset
  // top/center/bottom ki jagah seedha percentage position leke ffmpeg ki
  // drawtext expression banate hain, jo preview mein dikhi drag position se
  // hu-ba-hu match kare.
  const tx = textX !== undefined && textX !== null && textX !== '' ? Number(textX) : 50;
  const ty = textY !== undefined && textY !== null && textY !== '' ? Number(textY) : 90;
  if (!Number.isFinite(tx) || tx < 0 || tx > 100) return res.status(400).json({ error: 'Text X position invalid hai' });
  if (!Number.isFinite(ty) || ty < 0 || ty > 100) return res.status(400).json({ error: 'Text Y position invalid hai' });
  const withBox = textBox === true || textBox === 'true';

  const ffmpegCmd = await resolveFfmpegCommand();
  if (!ffmpegCmd) {
    return res.status(500).json({
      error: 'ffmpeg install nahi hai. Edit karne ke liye ffmpeg chahiye',
      detail: 'https://ffmpeg.org/download.html se install karo ya FFMPEG_PATH set karo.'
    });
  }

  // Kam se kam ek edit zaroor select ho — warna ye sirf ek expensive no-op
  // re-encode hi ban jaata.
  const hasAnyEdit = start !== null || end !== null || denoise || vol !== 1 || spd !== 1 ||
    fx !== 'none' || bright !== 0 || contr !== 1 || satur !== 1 || doFadeIn || doFadeOut || doText;
  if (!hasAnyEdit) {
    return res.status(400).json({ error: 'Kam se kam ek edit option choose karo' });
  }

  // Fade in/out ke liye final (trim + speed apply hone ke baad wali) duration
  // chahiye — isliye agar trim end nahi diya, to poore source ki duration
  // ffprobe se nikaalte hain. Trim (-ss/-t) hamesha SOURCE timeline par hota
  // hai (upar wale comment jaisa), par setpts/atempo pehle hi apply ho chuke
  // filter-chain mein fade se pehle, isliye fade ka st= final/output timeline
  // (jo already sped-up hai) ke against hi sahi baithta hai.
  let fadeInDur = 0, fadeOutDur = 0, fadeOutStart = 0;
  if (doFadeIn || doFadeOut) {
    let sourceDur = null;
    if (end !== null) {
      sourceDur = end - (start || 0);
    } else {
      const ffprobeCmd = await resolveFfprobeCommand();
      const totalDur = ffprobeCmd ? await getMediaDuration(ffprobeCmd, inputPath) : null;
      if (totalDur !== null) sourceDur = totalDur - (start || 0);
    }
    const finalDur = sourceDur !== null && sourceDur > 0 ? sourceDur / spd : null;
    if (finalDur !== null) {
      const fd = Math.min(1, finalDur / 2.5);
      if (doFadeIn) fadeInDur = fd;
      if (doFadeOut) { fadeOutDur = fd; fadeOutStart = Math.max(0, finalDur - fd); }
    }
  }

  // Text overlay: user ka text ek temp .txt file mein likh dete hain aur
  // drawtext ke textfile= param se padhte hain (text= inline karne par
  // quotes/colons/backslash sab manually escape karne padte — bahut fragile
  // aur injection-prone). expansion=none zaroori hai warna text mein '%'
  // (jaise "100% off") ko drawtext apna khud ka %{...} expression samajh kar
  // "Stray %" error de deta (live test se confirm hua).
  let textFilePath = null;
  if (doText) {
    textFilePath = path.join(DOWNLOADS_DIR, `${uid.slice(0,8)}_${Date.now()}_overlay.txt`);
    fs.writeFileSync(textFilePath, textStr, 'utf8');
  }

  const audioFilters = [];
  if (denoise) audioFilters.push('afftdn');
  if (vol !== 1) audioFilters.push(`volume=${vol}`);
  if (spd !== 1) audioFilters.push(`atempo=${spd}`);
  if (fadeInDur > 0) audioFilters.push(`afade=t=in:st=0:d=${fadeInDur.toFixed(2)}`);
  if (fadeOutDur > 0) audioFilters.push(`afade=t=out:st=${fadeOutStart.toFixed(2)}:d=${fadeOutDur.toFixed(2)}`);

  const videoFilters = [];
  if (fx !== 'none') videoFilters.push(EDIT_EFFECT_FILTERS[fx]);
  if (bright !== 0 || contr !== 1 || satur !== 1) videoFilters.push(`eq=brightness=${bright}:contrast=${contr}:saturation=${satur}`);
  if (doText) {
    const escapedPath = escapeFfmpegFilterPath(textFilePath);
    // Preview mein text ka center (drag point) x%/y% par hota hai — wahi
    // center yahan bhi maintain karte hain (text_w/2, text_h/2 minus karke)
    // taaki final export exactly wahi jagah dikhaye jahan preview mein tha.
    const xExpr = `(w*${(tx / 100).toFixed(4)})-text_w/2`;
    const yExpr = `(h*${(ty / 100).toFixed(4)})-text_h/2`;
    const boxParts = withBox ? ':box=1:boxcolor=0x000000@0.45:boxborderw=12' : '';
    videoFilters.push(`drawtext=textfile='${escapedPath}':reload=0:expansion=none:font=${EDIT_ALLOWED_FONTS[fontKey]}:fontsize=${fSize}:fontcolor=0x${colorHex.slice(1)}${boxParts}:x=${xExpr}:y=${yExpr}`);
  }
  if (spd !== 1) videoFilters.push(`setpts=PTS/${spd}`);
  if (fadeInDur > 0) videoFilters.push(`fade=t=in:st=0:d=${fadeInDur.toFixed(2)}`);
  if (fadeOutDur > 0) videoFilters.push(`fade=t=out:st=${fadeOutStart.toFixed(2)}:d=${fadeOutDur.toFixed(2)}`);

  const outName    = `${path.basename(safeName, path.extname(safeName))}_edited_${Date.now()}.mp4`;
  const outputPath = path.join(DOWNLOADS_DIR, outName);

  // -ss aur -t (duration) DONO input options hain (before -i) — isliye trim
  // hamesha ORIGINAL/source timeline par hi hota hai, chahe -vf/-af mein
  // speed filters (setpts/atempo) kitne bhi ho. Agar -to ko output option
  // ke roop mein use karte (after -i), to wo POST-FILTER timeline par apply
  // hota — speed change ke saath combine karne par galat trim length deta
  // (live test se confirm hua: trim(1-8) + speed 1.25x se expected 5.6s ki
  // jagah 8s output aaya tha, kyunki -to speed-adjusted output pts par cut
  // kar raha tha, source par nahi).
  const argParts = [quoteIfPath(ffmpegCmd), '-y'];
  if (start !== null) argParts.push('-ss', start);
  if (end !== null) argParts.push('-t', end - (start || 0));
  argParts.push('-i', `"${inputPath}"`);
  if (videoFilters.length) argParts.push('-vf', `"${videoFilters.join(',')}"`);
  if (audioFilters.length) argParts.push('-af', `"${audioFilters.join(',')}"`);
  argParts.push('-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-c:a', 'aac', `"${outputPath}"`);
  const command = argParts.join(' ');

  console.log(`[${uid.slice(0,8)}] Editing video: ${safeName} -> ${outName}`, { start, end, denoise, vol, spd, fx, doText });

  exec(command, { timeout: 5 * 60 * 1000, shell: true, cwd: DOWNLOADS_DIR }, (error, stdout, stderr) => {
    if (textFilePath) fs.unlink(textFilePath, () => {}); // best-effort cleanup, temp file hai
    if (error || !fs.existsSync(outputPath)) {
      console.error('ffmpeg edit error:', stderr || (error && error.message));
      return res.status(500).json({ error: 'Edit fail ho gaya. Dobara try karo' });
    }
    console.log(`[${uid.slice(0,8)}] Edited: ${outName}`);
    logActivity(req, 'video_edit', { sourceFilename: safeName, outName, start, end, denoise, vol, spd, fx, doText });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  });
});

// ─── PLAYGROUND FILE GENERATION — AI ke jawab mein ```pdf/docx/xlsx/pptx
//     fenced block ko real downloadable file mein badalta hai (jaise
//     ChatGPT ka Code Interpreter). AI sirf structured plain text/CSV deta
//     hai — asli PDF/Word/Excel/PowerPoint yahan bante hain. ──────────────

// "**bold**" ko segments mein todta hai taaki PDF/DOCX dono mein bold text
// sahi se render ho (dono libraries ko run-by-run bold/normal chahiye).
function splitBoldSegments(line) {
  const segments = [];
  const re = /\*\*(.+?)\*\*/g;
  let lastIndex = 0, m;
  while ((m = re.exec(line))) {
    if (m.index > lastIndex) segments.push({ text: line.slice(lastIndex, m.index), bold: false });
    segments.push({ text: m[1], bold: true });
    lastIndex = re.lastIndex;
  }
  if (lastIndex < line.length) segments.push({ text: line.slice(lastIndex), bold: false });
  return segments.length ? segments : [{ text: line, bold: false }];
}

// Simple markdown-jaisa text ko line-by-line block type mein todta hai —
// PDF aur DOCX dono generator isi common structure ko use karte hain.
function parseDocLines(text) {
  return text.split('\n').map(raw => {
    const line = raw.replace(/\r$/, '');
    if (/^#\s+/.test(line))   return { type: 'h1', text: line.replace(/^#\s+/, '') };
    if (/^##\s+/.test(line))  return { type: 'h2', text: line.replace(/^##\s+/, '') };
    if (/^###\s+/.test(line)) return { type: 'h3', text: line.replace(/^###\s+/, '') };
    if (/^\s*[-*]\s+/.test(line)) return { type: 'bullet', text: line.replace(/^\s*[-*]\s+/, '') };
    const numMatch = line.match(/^\s*(\d+)\.\s+(.*)$/);
    if (numMatch) return { type: 'numbered', num: numMatch[1], text: numMatch[2] };
    if (line.trim() === '') return { type: 'blank', text: '' };
    return { type: 'text', text: line };
  });
}

// pdfkit ke standard 14 base fonts (Helvetica/Helvetica-Bold) sirf WinAnsi
// encoding support karte hain — koi bhi emoji ya doosra exotic Unicode
// character (jaise 📅) daalne par silently garbled bytes render hote hain
// (jaise "📅" ban jaata hai "Ø=ÜÄ"). Yahan sirf wahi characters rakhte hain
// jo WinAnsi mein safely map hote hain, baaki hata dete hain.
const PDF_WINANSI_SAFE_EXTRAS = new Set([0x2013, 0x2014, 0x2018, 0x2019, 0x201A, 0x201C, 0x201D, 0x201E, 0x2020, 0x2021, 0x2022, 0x2026, 0x2030, 0x2039, 0x203A, 0x2122, 0x20AC, 0x0192]);
function sanitizePdfText(text) {
  const filtered = Array.from(text).filter(ch => {
    const code = ch.codePointAt(0);
    return code === 0x0A || (code >= 0x20 && code <= 0x7E) || (code >= 0xA0 && code <= 0xFF) || PDF_WINANSI_SAFE_EXTRAS.has(code);
  }).join('');
  // Emoji hatane ke baad line mein aksar orphan double-space reh jaata hai
  // (jaise "📅 Date:" → " Date:") — per-line trim se wo clean ho jaata hai.
  return filtered.split('\n').map(line => line.replace(/[ \t]+/g, ' ').trim()).join('\n');
}

function generatePdfBuffer(content) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 50, size: 'A4' });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    for (const line of parseDocLines(sanitizePdfText(content))) {
      if (line.type === 'blank') { doc.moveDown(0.6); continue; }
      const prefix = line.type === 'bullet' ? '•  ' : line.type === 'numbered' ? `${line.num}.  ` : '';
      const fontSize = line.type === 'h1' ? 20 : line.type === 'h2' ? 16 : line.type === 'h3' ? 13 : 11;
      const baseBold = line.type === 'h1' || line.type === 'h2' || line.type === 'h3';
      doc.fontSize(fontSize);
      const segments = splitBoldSegments(prefix + line.text);
      segments.forEach((seg, i) => {
        doc.font(baseBold || seg.bold ? 'Helvetica-Bold' : 'Helvetica');
        doc.text(seg.text, { continued: i < segments.length - 1 });
      });
      doc.moveDown(baseBold ? 0.5 : 0.25);
    }
    doc.end();
  });
}

async function generateDocxBuffer(content) {
  const paragraphs = parseDocLines(content).map(line => {
    if (line.type === 'blank') return new DocxParagraph({ text: '' });
    const prefix = line.type === 'numbered' ? `${line.num}. ` : '';
    const segments = splitBoldSegments(prefix + line.text);
    const runs = segments.map(seg => new DocxTextRun({ text: seg.text, bold: seg.bold }));
    const opts = { children: runs };
    if (line.type === 'h1') opts.heading = DocxHeadingLevel.HEADING_1;
    else if (line.type === 'h2') opts.heading = DocxHeadingLevel.HEADING_2;
    else if (line.type === 'h3') opts.heading = DocxHeadingLevel.HEADING_3;
    else if (line.type === 'bullet') opts.bullet = { level: 0 };
    return new DocxParagraph(opts);
  });
  const doc = new DocxDocument({ sections: [{ children: paragraphs }] });
  return DocxPacker.toBuffer(doc);
}

// Chhota hand-rolled CSV parser (quoted commas/newlines handle karta hai)
// — XLSX.utils.aoa_to_sheet ko seedha rows-of-arrays chahiye.
function parseCsv(text) {
  const rows = [];
  let row = [], field = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; }
      else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      if (row.length > 1 || row[0] !== '') rows.push(row);
      row = [];
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function generateXlsxBuffer(content) {
  const rows = parseCsv(content.trim());
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

async function generatePptxBuffer(content) {
  const pptx = new PptxGenJS();
  const slides = content.split(/\n\s*---\s*\n/);
  for (const raw of slides) {
    const lines = raw.split('\n').map(l => l.trim()).filter(l => l !== '');
    if (!lines.length) continue;
    const slide = pptx.addSlide();
    const title = lines[0].replace(/^#+\s*/, '');
    slide.addText(title, { x: 0.5, y: 0.4, w: '90%', fontSize: 28, bold: true, color: '363636' });
    const bodyLines = lines.slice(1).map(l => l.replace(/^[-*]\s+/, ''));
    if (bodyLines.length) {
      slide.addText(
        bodyLines.map(t => ({ text: t, options: { bullet: true, breakLine: true } })),
        { x: 0.5, y: 1.3, w: '90%', h: 4.5, fontSize: 18, color: '444444' }
      );
    }
  }
  return pptx.write({ outputType: 'nodebuffer' });
}

// TXT/MD ke liye koi "conversion" nahi karna padta — jo text diya hai wahi
// content hai, bas extension/mime badalta hai. RTF ek simple hand-rolled
// format hai (koi library nahi chahiye) — same markdown-jaisi structure
// (parseDocLines/splitBoldSegments) PDF/DOCX ke saath consistent rakhne
// ke liye reuse karta hai.
function generateTxtBuffer(content) {
  return Buffer.from(content, 'utf8');
}

// RTF spec mein backslash/braces escape karna zaroori hai, aur non-ASCII
// characters ko \uN? escape sequence mein likhna padta hai (N signed
// 16-bit) — isliye astral (surrogate-pair) characters jaise emoji ko
// dobara UTF-16 code units mein todhte hain taaki har unit apna \uN? paaye.
function escapeRtfText(text) {
  let out = '';
  for (const ch of text) {
    const code = ch.codePointAt(0);
    if (ch === '\\') out += '\\\\';
    else if (ch === '{') out += '\\{';
    else if (ch === '}') out += '\\}';
    else if (code < 128) out += ch;
    else if (code > 0xFFFF) {
      const c = code - 0x10000;
      const hi = 0xD800 + (c >> 10);
      const lo = 0xDC00 + (c & 0x3FF);
      out += `\\u${hi}?\\u${lo}?`;
    } else {
      out += `\\u${code}?`;
    }
  }
  return out;
}

function generateRtfBuffer(content) {
  const parts = ['{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\n'];
  for (const line of parseDocLines(content)) {
    if (line.type === 'blank') { parts.push('\\par\n'); continue; }
    const prefix = line.type === 'bullet' ? '\u2022  ' : line.type === 'numbered' ? `${line.num}.  ` : '';
    const fs = line.type === 'h1' ? 40 : line.type === 'h2' ? 32 : line.type === 'h3' ? 26 : 22;
    const baseBold = line.type === 'h1' || line.type === 'h2' || line.type === 'h3';
    parts.push(`\\fs${fs} `);
    for (const seg of splitBoldSegments(prefix + line.text)) {
      const bold = baseBold || seg.bold;
      if (bold) parts.push('\\b ');
      parts.push(escapeRtfText(seg.text));
      if (bold) parts.push('\\b0 ');
    }
    parts.push('\\par\n');
  }
  parts.push('}');
  return Buffer.from(parts.join(''), 'utf8');
}

// Playground ke chat-driven file cards + naya standalone "File Converter"
// home-page tool — dono isi ek generator map aur endpoint ko reuse karte
// hain (format+content -> real file, bas itna hi common contract hai).
const DOC_FILE_GENERATORS = {
  pdf: generatePdfBuffer, docx: generateDocxBuffer, xlsx: generateXlsxBuffer, pptx: generatePptxBuffer,
  txt: generateTxtBuffer, md: generateTxtBuffer, rtf: generateRtfBuffer,
};
const MAX_PG_FILE_CONTENT_CHARS = 200000;

app.post('/api/playground/generate-file', requireAuth, async (req, res) => {
  const { format, content } = req.body;
  const generator = DOC_FILE_GENERATORS[format];
  if (!generator) return res.status(400).json({ error: 'Ye file format support nahi hai' });
  if (typeof content !== 'string' || !content.trim()) return res.status(400).json({ error: 'Content khaali hai' });
  if (content.length > MAX_PG_FILE_CONTENT_CHARS) return res.status(400).json({ error: 'Content bahut bada hai' });

  try {
    const buffer = await generator(content);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pg.${format}`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'playground_file_generate', { format });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('generate-file error:', err);
    res.status(500).json({ error: 'File generate nahi ho payi. Dobara try karo' });
  }
});

// ─── FILE CONVERTER — images ko seedha PDF/DOCX/PPTX mein embed karna ─────
// Koi OCR/text-extraction nahi — jaisa hai waisa hi visual image naye
// document format mein daal dete hain (iLovePDF ke "image to file" tools
// jaisa). Sirf PNG/JPEG allowed hain kyunki pdfkit aur docx dono sirf
// inhi do formats ko natively embed kar sakte hain.
const MAX_FC_IMAGE_BYTES = 15 * 1024 * 1024;
const MAX_FC_IMAGES_PER_REQUEST = 20;

function generateImagePdfBuffer(images) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 36, size: 'A4' });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
    images.forEach((img, i) => {
      if (i > 0) doc.addPage();
      const areaW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      const areaH = doc.page.height - doc.page.margins.top - doc.page.margins.bottom;
      doc.image(img.buffer, doc.page.margins.left, doc.page.margins.top, { fit: [areaW, areaH], align: 'center', valign: 'center' });
    });
    doc.end();
  });
}

async function generateImageDocxBuffer(images) {
  const MAX_W_PX = 600; // ~6.25in usable page width @ 96dpi
  const MAX_H_PX = 750; // ~7.8in usable page height @ 96dpi
  const paragraphs = images.map((img, i) => {
    let w = MAX_W_PX, h = Math.round(MAX_W_PX * 0.75);
    try {
      const dim = imageSize(img.buffer);
      if (dim.width && dim.height) {
        const scale = Math.min(MAX_W_PX / dim.width, MAX_H_PX / dim.height, 1);
        w = Math.round(dim.width * scale);
        h = Math.round(dim.height * scale);
      }
    } catch { /* dimensions na mile to default box size use ho jaata hai */ }
    return new DocxParagraph({
      pageBreakBefore: i > 0,
      children: [new DocxImageRun({
        data: img.buffer,
        type: img.mime === 'image/png' ? 'png' : 'jpg',
        transformation: { width: w, height: h },
      })],
    });
  });
  const doc = new DocxDocument({ sections: [{ children: paragraphs }] });
  return DocxPacker.toBuffer(doc);
}

async function generateImagePptxBuffer(images) {
  const pptx = new PptxGenJS();
  const slideW = pptx.presLayout.width / 914400, slideH = pptx.presLayout.height / 914400;
  const margin = 0.3;
  const w = slideW - margin * 2, h = slideH - margin * 2;
  for (const img of images) {
    const slide = pptx.addSlide();
    const dataUrl = `data:${img.mime};base64,${img.buffer.toString('base64')}`;
    slide.addImage({ data: dataUrl, x: margin, y: margin, w, h, sizing: { type: 'contain', w, h } });
  }
  return pptx.write({ outputType: 'nodebuffer' });
}

const IMAGE_DOC_GENERATORS = { pdf: generateImagePdfBuffer, docx: generateImageDocxBuffer, pptx: generateImagePptxBuffer };

app.post('/api/doc-tool/images-to-file', requireAuth, async (req, res) => {
  const { format, images } = req.body;
  const generator = IMAGE_DOC_GENERATORS[format];
  if (!generator) return res.status(400).json({ error: 'Image se ye file format nahi ban sakta' });
  if (!Array.isArray(images) || !images.length) return res.status(400).json({ error: 'Koi image nahi mili' });
  if (images.length > MAX_FC_IMAGES_PER_REQUEST) return res.status(400).json({ error: `Ek baar mein max ${MAX_FC_IMAGES_PER_REQUEST} images allowed hain` });

  const parsed = [];
  for (const dataUrl of images) {
    const m = /^data:(image\/(?:png|jpeg));base64,(.+)$/.exec(dataUrl || '');
    if (!m) return res.status(400).json({ error: 'Sirf PNG ya JPEG images allowed hain' });
    const buffer = Buffer.from(m[2], 'base64');
    if (buffer.length > MAX_FC_IMAGE_BYTES) return res.status(400).json({ error: 'Ek image 15MB se badi hai' });
    parsed.push({ buffer, mime: m[1] });
  }

  try {
    const buffer = await generator(parsed);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_fc.${format}`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'file_converter_image_generate', { format, count: parsed.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('images-to-file error:', err);
    res.status(500).json({ error: 'File generate nahi ho payi. Dobara try karo' });
  }
});

// ─── IMAGE CONVERTER — koi bhi format se koi bhi format, quality compression,
//     aur target file-size (KB/MB) tak size ghatao/badhao ──────────────────
const MAX_IMG_TOOL_BYTES = 25 * 1024 * 1024; // 25MB per image
const imgToolUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_IMG_TOOL_BYTES } });

function withImgToolUpload(multerMiddleware, handler) {
  return (req, res) => {
    multerMiddleware(req, res, (err) => {
      if (err instanceof multer.MulterError && err.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ error: `Image bahut badi hai (max ${MAX_IMG_TOOL_BYTES / 1024 / 1024}MB)` });
      }
      if (err) return res.status(400).json({ error: 'Upload fail ho gaya' });
      handler(req, res);
    });
  };
}

// sharp.format se verified: jpeg/png/webp/tiff/gif/heif — sabhi is prebuilt
// binary mein INPUT aur OUTPUT dono support karte hain. BMP ke liye koi
// encoder available nahi hai is build mein, aur HEIC/HEIF output patent-
// encumbered HEVC codec maangta hai — isliye output list mein sirf wahi
// formats hain jo directly test karke verify kiye gaye hain. HEIC/HEIF
// (iPhone photos) input ke taur par already accept ho jaate hain (sharp
// khud decode kar leta hai) — bas naye output ke taur par offer nahi kiya.
const IMAGE_CONVERT_FORMATS = {
  jpeg: { ext: 'jpg',  mime: 'image/jpeg', label: 'JPG' },
  png:  { ext: 'png',  mime: 'image/png',  label: 'PNG' },
  webp: { ext: 'webp', mime: 'image/webp', label: 'WebP' },
  avif: { ext: 'avif', mime: 'image/avif', label: 'AVIF' },
  tiff: { ext: 'tiff', mime: 'image/tiff', label: 'TIFF' },
  gif:  { ext: 'gif',  mime: 'image/gif',  label: 'GIF' },
};
// In formats mein hi numeric quality knob hota hai jiska file-size par
// predictable/monotonic asar hota hai — target file-size search (binary
// search + up/downscale) sirf inhi ke liye chalate hain. PNG lossless hai
// (bina palette ke quality ka size par kaam asar) aur GIF palette-limited
// hai, dono mein reliable target-size search possible nahi hai.
const IMAGE_QUALITY_FORMATS = new Set(['jpeg', 'webp', 'avif', 'tiff']);

function encodeImageBuffer(buffer, targetFormat, quality, scale, origWidth) {
  let pipeline = sharp(buffer, { failOn: 'none' }).rotate(); // EXIF orientation ke hisaab se auto-seedha karo
  if (scale && Math.abs(scale - 1) > 0.001 && origWidth) {
    const w = Math.min(8000, Math.max(16, Math.round(origWidth * scale))); // 8000px cap — DoS-safe upscale limit
    pipeline = pipeline.resize({ width: w });
  }
  switch (targetFormat) {
    case 'jpeg': pipeline = pipeline.jpeg({ quality, mozjpeg: true }); break;
    case 'png':  pipeline = pipeline.png({ compressionLevel: 9, quality, ...(quality < 100 ? { palette: true } : {}) }); break;
    case 'webp': pipeline = pipeline.webp({ quality }); break;
    case 'avif': pipeline = pipeline.avif({ quality }); break;
    case 'tiff': pipeline = pipeline.tiff({ quality, compression: 'jpeg' }); break;
    case 'gif':  pipeline = pipeline.gif(); break;
    default: throw new Error('Ye format support nahi hai');
  }
  return pipeline.toBuffer();
}

// Target file-size (KB/MB) mode — user ek exact size maangta hai (jaise
// government-form photo/signature ke liye "50KB se kam" ya "kam se kam
// 20KB"). Pehle quality=1 aur quality=100 (native resolution) par size
// naapte hain: agar target dono ke beech mein hai to quality binary-search
// karte hain; agar target quality=100 se bhi bada hai to upscale karte hain
// (real detail add nahi hoti, par size genuinely badhta hai — yही ek tarika
// hai file ko "bada" banane ka); agar target quality=1 se bhi chhota hai to
// downscale karte hain.
async function convertImageToTargetSize(buffer, targetFormat, targetSizeBytes, origWidth) {
  if (!IMAGE_QUALITY_FORMATS.has(targetFormat)) {
    return encodeImageBuffer(buffer, targetFormat, 90, 1, origWidth);
  }

  const atMin = await encodeImageBuffer(buffer, targetFormat, 1, 1, origWidth);
  const atMax = await encodeImageBuffer(buffer, targetFormat, 100, 1, origWidth);

  if (targetSizeBytes >= atMax.length) {
    // Pehle multiplicatively upscale karte hain jab tak target cross na ho
    // jaaye (bracket dhoondte hain), phir usi bracket ke andar SCALE ko
    // binary-search karte hain — sirf ek jump se rukne par size target se
    // kaafi zyada (overshoot) ho sakta tha, ye refinement usko exact ke
    // kaafi kareeb le aata hai.
    let loScale = 1, hiScale = 1, hiOut = atMax, guard = 0;
    while (hiOut.length < targetSizeBytes && hiScale < 8 && guard < 8) {
      loScale = hiScale;
      hiScale = Math.min(8, hiScale * 1.35);
      hiOut = await encodeImageBuffer(buffer, targetFormat, 100, hiScale, origWidth);
      guard++;
    }
    if (hiOut.length < targetSizeBytes) return hiOut; // 8x cap tak bhi nahi pahoncha — yahi max possible hai

    let best = hiOut, bestDiff = Math.abs(hiOut.length - targetSizeBytes);
    let lo = loScale, hi = hiScale;
    for (let i = 0; i < 6; i++) {
      const mid = (lo + hi) / 2;
      const out = await encodeImageBuffer(buffer, targetFormat, 100, mid, origWidth);
      const diff = Math.abs(out.length - targetSizeBytes);
      if (diff < bestDiff) { best = out; bestDiff = diff; }
      if (out.length > targetSizeBytes) hi = mid; else lo = mid;
    }
    return best;
  }

  if (targetSizeBytes <= atMin.length) {
    // Same refinement, downscale direction (bracket + scale binary-search).
    let loScale = 1, hiScale = 1, loOut = atMin, guard = 0;
    while (loOut.length > targetSizeBytes && loScale > 0.02 && guard < 8) {
      hiScale = loScale;
      loScale = Math.max(0.02, loScale * 0.7);
      loOut = await encodeImageBuffer(buffer, targetFormat, 1, loScale, origWidth);
      guard++;
    }
    if (loOut.length > targetSizeBytes) return loOut; // 0.02x floor tak bhi target se bada hai — yahi min possible hai

    let best = loOut, bestDiff = Math.abs(loOut.length - targetSizeBytes);
    let lo = loScale, hi = hiScale;
    for (let i = 0; i < 6; i++) {
      const mid = (lo + hi) / 2;
      const out = await encodeImageBuffer(buffer, targetFormat, 1, mid, origWidth);
      const diff = Math.abs(out.length - targetSizeBytes);
      if (diff < bestDiff) { best = out; bestDiff = diff; }
      if (out.length > targetSizeBytes) lo = mid; else hi = mid;
    }
    return best;
  }

  let lo = 1, hi = 100, best = atMax, bestDiff = Math.abs(atMax.length - targetSizeBytes);
  for (let i = 0; i < 7; i++) {
    const mid = Math.round((lo + hi) / 2);
    const out = await encodeImageBuffer(buffer, targetFormat, mid, 1, origWidth);
    const diff = Math.abs(out.length - targetSizeBytes);
    if (diff < bestDiff) { best = out; bestDiff = diff; }
    if (out.length > targetSizeBytes) hi = mid - 1; else lo = mid + 1;
    if (lo > hi) break;
  }
  return best;
}

app.post('/api/image-tools/convert', requireAuth, withImgToolUpload(imgToolUpload.single('image'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi image nahi mili' });

  const targetFormat = String(req.body.targetFormat || '').toLowerCase();
  if (!IMAGE_CONVERT_FORMATS[targetFormat]) return res.status(400).json({ error: 'Ye output format support nahi hai' });

  const targetSizeKB = req.body.targetSizeKB ? Number(req.body.targetSizeKB) : null;
  const quality = Math.min(100, Math.max(1, Number(req.body.quality) || 85));

  try {
    const srcMeta = await sharp(req.file.buffer, { failOn: 'none' }).metadata();
    if (!srcMeta.width || !srcMeta.height) return res.status(400).json({ error: 'Ye file valid image nahi lag rahi' });

    let outBuffer;
    if (targetSizeKB && targetSizeKB > 0) {
      outBuffer = await convertImageToTargetSize(req.file.buffer, targetFormat, Math.round(targetSizeKB * 1024), srcMeta.width);
    } else {
      outBuffer = await encodeImageBuffer(req.file.buffer, targetFormat, quality, 1, srcMeta.width);
    }

    const fmtMeta = IMAGE_CONVERT_FORMATS[targetFormat];
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_img.${fmtMeta.ext}`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), outBuffer);

    logActivity(req, 'image_convert', { targetFormat, originalSizeBytes: req.file.size, outputSizeBytes: outBuffer.length });
    res.json({
      success: true,
      fileUrl: `/files/${outName}`,
      filename: outName,
      format: targetFormat,
      originalSizeBytes: req.file.size,
      outputSizeBytes: outBuffer.length,
      width: srcMeta.width,
      height: srcMeta.height,
    });
  } catch (err) {
    console.error('Image convert error:', err);
    res.status(500).json({ error: 'Image convert nahi ho payi. Dobara try karo' });
  }
}));

const MAX_IMG_BATCH_FILES = 20;

// Multiple images ek saath — sab par same format/quality/target-size apply
// hoke ek hi ZIP mein wapas aate hain (jaisa Split PDF/PDF-to-JPG ka pattern hai).
app.post('/api/image-tools/convert-batch', requireAuth, withImgToolUpload(imgToolUpload.array('images', MAX_IMG_BATCH_FILES), async (req, res) => {
  if (!req.files || !req.files.length) return res.status(400).json({ error: 'Koi images nahi mili' });

  const targetFormat = String(req.body.targetFormat || '').toLowerCase();
  if (!IMAGE_CONVERT_FORMATS[targetFormat]) return res.status(400).json({ error: 'Ye output format support nahi hai' });

  const targetSizeKB = req.body.targetSizeKB ? Number(req.body.targetSizeKB) : null;
  const quality = Math.min(100, Math.max(1, Number(req.body.quality) || 85));
  const fmtMeta = IMAGE_CONVERT_FORMATS[targetFormat];

  try {
    const zip = new AdmZip();
    let totalOriginal = 0, totalOutput = 0, converted = 0;

    for (const file of req.files) {
      try {
        const srcMeta = await sharp(file.buffer, { failOn: 'none' }).metadata();
        if (!srcMeta.width || !srcMeta.height) continue; // corrupt/non-image file — skip, don't fail the whole batch

        const outBuffer = (targetSizeKB && targetSizeKB > 0)
          ? await convertImageToTargetSize(file.buffer, targetFormat, Math.round(targetSizeKB * 1024), srcMeta.width)
          : await encodeImageBuffer(file.buffer, targetFormat, quality, 1, srcMeta.width);

        const baseName = (file.originalname || `image-${converted + 1}`).replace(/\.[^.]+$/, '').replace(/[^a-zA-Z0-9_-]/g, '_') || `image-${converted + 1}`;
        zip.addFile(`${baseName}.${fmtMeta.ext}`, outBuffer);
        totalOriginal += file.size;
        totalOutput += outBuffer.length;
        converted++;
      } catch (perFileErr) {
        console.error('Batch image convert (single file) error:', perFileErr);
      }
    }

    if (!converted) return res.status(400).json({ error: 'Koi bhi image convert nahi ho payi — valid image files upload karo' });

    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_imgbatch.zip`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), zip.toBuffer());

    logActivity(req, 'image_convert_batch', { targetFormat, count: converted, totalOriginal, totalOutput });
    res.json({
      success: true,
      fileUrl: `/files/${outName}`,
      filename: outName,
      format: targetFormat,
      count: converted,
      totalOriginalSizeBytes: totalOriginal,
      totalOutputSizeBytes: totalOutput,
    });
  } catch (err) {
    console.error('Image batch convert error:', err);
    res.status(500).json({ error: 'Batch convert nahi ho paya. Dobara try karo' });
  }
}));

// ─── PDF TOOLS — Merge/Split/Compress/Convert/Edit, iLovePDF-jaisa ────────
// Uploads seedha memory mein aate hain (koi disk-write nahi, ek-hi request
// mein process ho ke turant DOWNLOADS_DIR mein result likh dete hain) —
// baaki upload flows (Video Editor) ki tarah "upload karo, phir baad mein
// process karo" wala do-step process yahan zaroori nahi kyunki har tool
// ek hi immediate action hai.
const MAX_PDF_TOOL_BYTES = 30 * 1024 * 1024; // 30MB per file
const pdfToolUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_PDF_TOOL_BYTES } });

function withPdfToolUpload(multerMiddleware, handler) {
  return (req, res) => {
    multerMiddleware(req, res, (err) => {
      if (err instanceof multer.MulterError && err.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ error: `File bahut badi hai (max ${MAX_PDF_TOOL_BYTES / 1024 / 1024}MB)` });
      }
      if (err) return res.status(400).json({ error: 'Upload fail ho gaya' });
      handler(req, res);
    });
  };
}

const pdfMimeCheck  = f => (f.mimetype === 'application/pdf' || /\.pdf$/i.test(f.originalname)) ? null : 'Sirf PDF files allowed hain';
const docxMimeCheck = f => /\.docx$/i.test(f.originalname) ? null : 'Sirf DOCX (Word) files allowed hain';
const pptxMimeCheck = f => /\.pptx$/i.test(f.originalname) ? null : 'Sirf PPTX (PowerPoint) files allowed hain';
const xlsxMimeCheck = f => /\.xlsx$/i.test(f.originalname) ? null : 'Sirf XLSX (Excel) files allowed hain';

// pdf-parse ek bahut purani bundled pdfjs (v1.10.100) use karta hai jo
// pdfkit/pdf-lib jaise modern PDF writers ke output (aur kabhi-kabhi kuch
// real-world PDFs) par "bad XRef entry" jaisi errors deti hai — verified
// isi session mein: pdfkit/pdf-lib se banaye 2-page test PDFs pdf-parse se
// bilkul parse nahi hue. Isliye PDF Tools ke liye alag se modern
// pdfjs-dist (dynamic import — sirf ESM build available hai) use karte
// hain, jo teeno cases (pdfkit, pdf-lib, real Chrome-printed PDF) mein
// sahi se per-page text nikaal ke deta hai.
let _pdfjsPromise = null;
function getPdfjs() { return _pdfjsPromise || (_pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs')); }

async function extractPdfPages(buffer) {
  const pdfjs = await getPdfjs();
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buffer), useSystemFonts: true }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    // Items ke beech y-position badalne par naya line maante hain (2pt
    // tolerance) — isse paragraph/line breaks kaafi had tak asli jaisi
    // structure mein wapas aa jaate hain, sirf ek-line-mein-sab-kuch nahi.
    let text = '', lastY = null;
    for (const item of content.items) {
      const y = item.transform[5];
      if (lastY !== null && Math.abs(y - lastY) > 2) text += '\n';
      else if (text && !/[\n ]$/.test(text)) text += ' ';
      text += item.str;
      lastY = y;
    }
    pages.push(text.trim());
  }
  return pages;
}

async function generatePdfToDocxBuffer(buffer) {
  const text = (await extractPdfPages(buffer)).join('\n\n').trim();
  if (!text) throw new Error('NO_TEXT');
  return generateDocxBuffer(text);
}
async function generatePdfToPptxBuffer(buffer) {
  const pages = await extractPdfPages(buffer);
  if (!pages.some(p => p.trim())) throw new Error('NO_TEXT');
  const content = pages.map((p, i) => `# Page ${i + 1}\n${p}`).join('\n\n---\n\n');
  return generatePptxBuffer(content);
}
async function generatePdfToXlsxBuffer(buffer) {
  const pages = await extractPdfPages(buffer);
  const rows = pages.flatMap((p, i) => [`--- Page ${i + 1} ---`, ...p.split('\n')]).filter(l => l.trim()).map(l => [l]);
  if (!rows.length) throw new Error('NO_TEXT');
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

async function generateDocxToPdfBuffer(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const text = (result.value || '').trim();
  if (!text) throw new Error('NO_TEXT');
  return generatePdfBuffer(text);
}
async function generatePptxToPdfBuffer(buffer) {
  const zip = new AdmZip(buffer);
  const slideEntries = zip.getEntries()
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
    .sort((a, b) => parseInt(a.entryName.match(/(\d+)/)[1], 10) - parseInt(b.entryName.match(/(\d+)/)[1], 10));
  const slides = slideEntries.map((e, i) => {
    const xml = e.getData().toString('utf8');
    const lines = [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
      .map(p => [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => decodeXmlEntities(m[1])).join(''))
      .filter(l => l.trim());
    return { number: i + 1, lines };
  });
  const content = slides.map(s => `# Slide ${s.number}\n${s.lines.map(l => `- ${l}`).join('\n')}`).join('\n\n');
  if (!content.trim()) throw new Error('NO_TEXT');
  return generatePdfBuffer(content);
}
async function generateXlsxToPdfBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const content = wb.SheetNames.map(name => `## ${name}\n${XLSX.utils.sheet_to_csv(wb.Sheets[name])}`).join('\n\n');
  if (!content.trim()) throw new Error('NO_TEXT');
  return generatePdfBuffer(content);
}

async function mergePdfs(buffers) {
  const merged = await PdfLibDocument.create();
  for (const buf of buffers) {
    const src = await PdfLibDocument.load(buf, { ignoreEncryption: true });
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(p => merged.addPage(p));
  }
  return Buffer.from(await merged.save());
}

async function splitPdfToZip(buffer) {
  const src = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const count = src.getPageCount();
  const zip = new AdmZip();
  for (let i = 0; i < count; i++) {
    const out = await PdfLibDocument.create();
    const [copied] = await out.copyPages(src, [i]);
    out.addPage(copied);
    zip.addFile(`page-${i + 1}.pdf`, Buffer.from(await out.save()));
  }
  return { buffer: zip.toBuffer(), count };
}

// pdf-lib khud image re-encoding/downsampling nahi karta (wo asli
// "compress" jo Ghostscript jaisa tool karta hai) — ye sirf structural
// overhead (metadata, duplicate objects, uncompressed xref table) hi
// kam karta hai. Image-heavy PDFs par savings modest ho sakti hain, jo
// UI mein bhi honestly dikhaya jaata hai (before/after size).
async function compressPdf(buffer) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true, updateMetadata: false });
  doc.setTitle(''); doc.setAuthor(''); doc.setSubject(''); doc.setKeywords([]); doc.setProducer(''); doc.setCreator('');
  return Buffer.from(await doc.save({ useObjectStreams: true }));
}

function makeSingleFileConvertRoute(mimeCheck, generator, outFormat, activityName) {
  return async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Koi file nahi mili' });
    const errMsg = mimeCheck(req.file);
    if (errMsg) return res.status(400).json({ error: errMsg });
    try {
      const buffer = await generator(req.file.buffer);
      const uid = req.user.id;
      const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.${outFormat}`;
      fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
      logActivity(req, activityName, {});
      res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
    } catch (err) {
      if (err.message === 'NO_TEXT') return res.status(400).json({ error: 'Is file mein koi readable text nahi mila' });
      console.error(`${activityName} error:`, err);
      res.status(500).json({ error: 'Convert nahi ho paaya. Kya ye ek valid file hai?' });
    }
  };
}

app.post('/api/pdf-tools/pdf-to-word', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(pdfMimeCheck, generatePdfToDocxBuffer, 'docx', 'pdf_to_word')));
app.post('/api/pdf-tools/pdf-to-pptx', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(pdfMimeCheck, generatePdfToPptxBuffer, 'pptx', 'pdf_to_pptx')));
app.post('/api/pdf-tools/pdf-to-excel', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(pdfMimeCheck, generatePdfToXlsxBuffer, 'xlsx', 'pdf_to_excel')));
app.post('/api/pdf-tools/word-to-pdf', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(docxMimeCheck, generateDocxToPdfBuffer, 'pdf', 'word_to_pdf')));
app.post('/api/pdf-tools/pptx-to-pdf', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(pptxMimeCheck, generatePptxToPdfBuffer, 'pdf', 'pptx_to_pdf')));
app.post('/api/pdf-tools/excel-to-pdf', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'),
  makeSingleFileConvertRoute(xlsxMimeCheck, generateXlsxToPdfBuffer, 'pdf', 'excel_to_pdf')));

app.post('/api/pdf-tools/merge', requireAuth, withPdfToolUpload(pdfToolUpload.array('files', 10), async (req, res) => {
  const files = req.files || [];
  if (files.length < 2) return res.status(400).json({ error: 'Kam se kam 2 PDF files upload karo' });
  const invalid = files.find(f => pdfMimeCheck(f));
  if (invalid) return res.status(400).json({ error: 'Sirf PDF files allowed hain' });
  try {
    const buffer = await mergePdfs(files.map(f => f.buffer));
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_merge', { count: files.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('pdf merge error:', err);
    res.status(500).json({ error: 'PDFs merge nahi ho paayi. Kya sabhi valid PDF files hain?' });
  }
}));

app.post('/api/pdf-tools/split', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const { buffer, count } = await splitPdfToZip(req.file.buffer);
    if (count < 2) return res.status(400).json({ error: 'Is PDF mein sirf 1 page hai, split karne ke liye kam se kam 2 pages chahiye' });
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.zip`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_split', { pageCount: count });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, pageCount: count });
  } catch (err) {
    console.error('pdf split error:', err);
    res.status(500).json({ error: 'PDF split nahi ho paayi. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/compress', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const originalSize = req.file.buffer.length;
    const buffer = await compressPdf(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_compress', { originalSize, compressedSize: buffer.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, originalSize, compressedSize: buffer.length });
  } catch (err) {
    console.error('pdf compress error:', err);
    res.status(500).json({ error: 'PDF compress nahi ho paayi. Kya ye ek valid PDF hai?' });
  }
}));

// ─── EDIT PDF — visual editor: text/image overlays ko asli PDF coordinates
// par draw karta hai. Client canvas-pixel positions ko already PDF-point
// coordinates mein convert karke bhejta hai (bottom-left origin, jaisa
// PDF spec mein hota hai) — server sirf seedha drawText/drawImage karta hai.
app.post('/api/pdf-tools/edit', requireAuth, async (req, res) => {
  const { pdfBase64, edits } = req.body;
  if (typeof pdfBase64 !== 'string' || !pdfBase64) return res.status(400).json({ error: 'PDF data missing' });
  if (pdfBase64.length * 0.75 > MAX_PDF_TOOL_BYTES) return res.status(400).json({ error: `PDF bahut badi hai (max ${MAX_PDF_TOOL_BYTES / 1024 / 1024}MB)` });
  if (!Array.isArray(edits) || !edits.length) return res.status(400).json({ error: 'Koi edit nahi mila' });
  if (edits.length > 200) return res.status(400).json({ error: 'Bahut zyada edits hain (max 200)' });

  try {
    const doc = await PdfLibDocument.load(Buffer.from(pdfBase64, 'base64'), { ignoreEncryption: true });
    const pageCount = doc.getPageCount();
    const font = await doc.embedFont(PdfLibStandardFonts.Helvetica);
    const boldFont = await doc.embedFont(PdfLibStandardFonts.HelveticaBold);
    const imageCache = new Map();

    for (const edit of edits) {
      const pageIndex = Number(edit.page);
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount) continue;
      const page = doc.getPage(pageIndex);
      const x = Number(edit.x), y = Number(edit.y);
      if (!Number.isFinite(x) || !Number.isFinite(y)) continue;

      if (edit.type === 'text') {
        const text = String(edit.text || '').slice(0, 500);
        if (!text) continue;
        const size = Math.min(200, Math.max(4, Number(edit.fontSize) || 16));
        const hex = /^#[0-9a-fA-F]{6}$/.test(edit.color || '') ? edit.color : '#000000';
        const r = parseInt(hex.slice(1, 3), 16) / 255, g = parseInt(hex.slice(3, 5), 16) / 255, b = parseInt(hex.slice(5, 7), 16) / 255;
        page.drawText(text, { x, y, size, font: edit.bold ? boldFont : font, color: pdfLibRgb(r, g, b) });
      } else if (edit.type === 'image') {
        const m = /^data:(image\/(?:png|jpeg));base64,(.+)$/.exec(edit.imageData || '');
        if (!m) continue;
        const w = Math.max(1, Number(edit.width) || 100), h = Math.max(1, Number(edit.height) || 100);
        let embedded = imageCache.get(edit.imageData);
        if (!embedded) {
          const imgBytes = Buffer.from(m[2], 'base64');
          embedded = m[1] === 'image/png' ? await doc.embedPng(imgBytes) : await doc.embedJpg(imgBytes);
          imageCache.set(edit.imageData, embedded);
        }
        page.drawImage(embedded, { x, y, width: w, height: h });
      }
    }

    const outBytes = Buffer.from(await doc.save());
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), outBytes);
    logActivity(req, 'pdf_edit', { editCount: edits.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('pdf edit error:', err);
    res.status(500).json({ error: 'PDF edit nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
});

// ─── PDF TOOLS (extended set) — Watermark, Rotate, Page Numbers, Repair,
//     PDF/A, Unlock/Protect, Organize, Compare, PDF↔JPG, HTML to PDF, OCR ──
// @cantoo/pdf-lib ek maintained fork hai jo original pdf-lib se drop-in
// compatible hai, bas encrypt()/password-load support extra hai (verified
// isi session mein: encrypt karke password ke saath dobara load hota hai;
// "unlock" ke liye pages ko fresh document mein copy karna padta hai kyunki
// sirf re-save karne se bhi encryption dictionary reh jaata hai).

function hexToRgb01(hex) {
  const m = /^#([0-9a-fA-F]{6})$/.exec(hex || '');
  const h = m ? m[1] : '888888';
  return { r: parseInt(h.slice(0, 2), 16) / 255, g: parseInt(h.slice(2, 4), 16) / 255, b: parseInt(h.slice(4, 6), 16) / 255 };
}

async function watermarkPdf(buffer, { text, opacity, rotation, fontSize, color }) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const font = await doc.embedFont(PdfLibStandardFonts.HelveticaBold);
  const { r, g, b } = hexToRgb01(color);
  for (const page of doc.getPages()) {
    const { width, height } = page.getSize();
    const textWidth = font.widthOfTextAtSize(text, fontSize);
    page.drawText(text, {
      x: width / 2 - textWidth / 2, y: height / 2, size: fontSize, font,
      color: pdfLibRgb(r, g, b), opacity, rotate: pdfLibDegrees(rotation),
    });
  }
  return Buffer.from(await doc.save());
}

async function rotatePdfBuffer(buffer, angle) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  for (const page of doc.getPages()) {
    page.setRotation(pdfLibDegrees((page.getRotation().angle + angle + 360) % 360));
  }
  return Buffer.from(await doc.save());
}

async function addPageNumbersToPdf(buffer, { position, startNumber }) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const font = await doc.embedFont(PdfLibStandardFonts.Helvetica);
  doc.getPages().forEach((page, i) => {
    const { width, height } = page.getSize();
    const label = String(startNumber + i);
    const size = 11;
    const textWidth = font.widthOfTextAtSize(label, size);
    const x = position === 'bottom-left' ? 30 : position === 'bottom-right' ? width - 30 - textWidth : width / 2 - textWidth / 2;
    const y = position === 'top-center' ? height - 30 : 20;
    page.drawText(label, { x, y, size, font, color: pdfLibRgb(0.35, 0.35, 0.35) });
  });
  return Buffer.from(await doc.save());
}

// pdf-lib se corruption "fix" karna best-effort hi hai — bas lenient parsing
// ke saath load karke ek clean re-save kar dete hain (kai chhoti-mooti xref/
// structural issues isi se theek ho jaati hain). Har corruption fix nahi
// hoga — agar file bilkul hi load na ho paaye to error hi sahi jawab hai.
async function repairPdfBuffer(buffer) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true, throwOnInvalidObject: false });
  return Buffer.from(await doc.save());
}

// IMPORTANT: ye ek asli, certified PDF/A conversion NAHI hai (wo ICC color
// profiles, XMP metadata, font-embedding guarantees jaisi cheezein maangta
// hai jo pdf-lib provide nahi karta) — sirf structural cleanup (object
// streams off, metadata clear) karta hai. UI mein bhi yahi disclose hota hai.
async function pdfToPdfABestEffort(buffer) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  doc.setProducer('PastePro'); doc.setCreator('PastePro PDF Tools');
  return Buffer.from(await doc.save({ useObjectStreams: false }));
}

async function unlockPdfBuffer(buffer, password) {
  const doc = await PdfLibDocument.load(buffer, { password });
  // Sirf re-save karne se bhi encryption dictionary reh jaata hai — isliye
  // pages ko ek fresh, kabhi-encrypt-na-hue document mein copy karte hain.
  const fresh = await PdfLibDocument.create();
  const copied = await fresh.copyPages(doc, doc.getPageIndices());
  copied.forEach(p => fresh.addPage(p));
  return Buffer.from(await fresh.save());
}

async function protectPdfBuffer(buffer, password) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  doc.encrypt({ userPassword: password, ownerPassword: password, permissions: { printing: 'highResolution' } });
  return Buffer.from(await doc.save());
}

async function organizePdfBuffer(buffer, pageOrder) {
  const src = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const fresh = await PdfLibDocument.create();
  const copied = await fresh.copyPages(src, pageOrder);
  copied.forEach(p => fresh.addPage(p));
  return Buffer.from(await fresh.save());
}

async function comparePdfTexts(bufferA, bufferB) {
  const pagesA = await extractPdfPages(bufferA);
  const pagesB = await extractPdfPages(bufferB);
  const maxPages = Math.max(pagesA.length, pagesB.length);
  const pageDiffs = [];
  let totalAdded = 0, totalRemoved = 0;
  for (let i = 0; i < maxPages; i++) {
    const parts = diffLines(pagesA[i] || '', pagesB[i] || '');
    let changed = false;
    parts.forEach(p => { if (p.added) { totalAdded++; changed = true; } if (p.removed) { totalRemoved++; changed = true; } });
    pageDiffs.push({ page: i + 1, changed, parts: parts.map(p => ({ value: p.value, added: !!p.added, removed: !!p.removed })) });
  }
  return { pageDiffs, totalAdded, totalRemoved, pagesA: pagesA.length, pagesB: pagesB.length };
}

// pdf.js sirf browser/ESM-friendly hai — server-side rasterize (real page
// image chahiye) ke liye Playwright ka headless Chromium isi client-side
// pdf.js build ko load karta hai jo Edit PDF mein bhi use hota hai, taaki
// dono jagah rendering behavior consistent rahe.
async function renderPdfPagesToImages(buffer, { scale = 1.8, format = 'jpeg', quality = 0.85 } = {}) {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.setContent('<!doctype html><html><body></body></html>');
    await page.addScriptTag({ url: `${PDFJS_CDN_BASE}/pdf.min.js` });
    const base64 = buffer.toString('base64');
    const dataUrls = await page.evaluate(async ({ base64, scale, format, quality, workerSrc }) => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const doc = await window.pdfjsLib.getDocument({ data: bytes }).promise;
      const out = [];
      for (let i = 1; i <= doc.numPages; i++) {
        const p = await doc.getPage(i);
        const viewport = p.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width; canvas.height = viewport.height;
        await p.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        out.push(canvas.toDataURL(`image/${format}`, quality));
      }
      return out;
    }, { base64, scale, format, quality, workerSrc: `${PDFJS_CDN_BASE}/pdf.worker.min.js` });
    return dataUrls.map(u => Buffer.from(u.split(',')[1], 'base64'));
  } finally {
    await browser.close();
  }
}
const PDFJS_CDN_BASE = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174';

async function pdfToJpgZip(buffer) {
  const images = await renderPdfPagesToImages(buffer, { format: 'jpeg', quality: 0.85, scale: 1.8 });
  const zip = new AdmZip();
  images.forEach((img, i) => zip.addFile(`page-${i + 1}.jpg`, img));
  return { buffer: zip.toBuffer(), count: images.length };
}

// SSRF guard — HTML to PDF ek URL bhi accept karta hai jo server khud fetch
// karta hai, isliye private/internal network addresses (localhost, LAN
// ranges, cloud metadata endpoint) ko explicitly block karna zaroori hai.
// (DNS-rebinding se poori tarah immune nahi hai — resolve-then-connect ke
// beech theoretically DNS badal sakta hai — par casual SSRF attempts ke
// against ye proportionate protection hai.)
function isPrivateIp(ip) {
  if (net.isIP(ip) === 4) {
    const p = ip.split('.').map(Number);
    return p[0] === 10 || p[0] === 127 || p[0] === 0 || (p[0] === 169 && p[1] === 254) || (p[0] === 172 && p[1] >= 16 && p[1] <= 31) || (p[0] === 192 && p[1] === 168);
  }
  if (net.isIP(ip) === 6) {
    const l = ip.toLowerCase();
    return l === '::1' || l.startsWith('fc') || l.startsWith('fd') || l.startsWith('fe80');
  }
  return true;
}
async function assertUrlIsSafeToFetch(urlStr) {
  let u;
  try { u = new URL(urlStr); } catch { throw new Error('URL invalid hai'); }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') throw new Error('Sirf http/https URLs allowed hain');
  if (u.hostname === 'localhost') throw new Error('Ye URL allowed nahi hai');
  let addresses;
  try { addresses = await dns.lookup(u.hostname, { all: true }); } catch { throw new Error('URL resolve nahi ho paya'); }
  if (!addresses.length || addresses.some(a => isPrivateIp(a.address))) throw new Error('Ye URL allowed nahi hai');
  return u;
}

async function htmlToPdfBuffer(input) {
  const trimmed = input.trim();
  const isUrl = /^https?:\/\/\S+$/i.test(trimmed) && !/\s/.test(trimmed);
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    if (isUrl) {
      const safeUrl = await assertUrlIsSafeToFetch(trimmed);
      await page.goto(safeUrl.toString(), { waitUntil: 'networkidle', timeout: 20000 });
    } else {
      await page.setContent(trimmed, { waitUntil: 'networkidle', timeout: 20000 });
    }
    const pdfBytes = await page.pdf({ format: 'A4', printBackground: true, margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' } });
    return Buffer.from(pdfBytes);
  } finally {
    await browser.close();
  }
}

// OCR PDF — scanned/image-based PDFs mein koi embedded text nahi hota,
// isliye pehle Playwright se rasterize karte hain, phir har page image
// Playground jaisa hi vision model OCR se guzarte hain (ek-ek karke, rate
// limit se bachne ke liye). Result ek Word document hai (invisible-layer
// wali "searchable PDF" nahi bana sakte — uske liye word-level bounding
// boxes chahiye hote jo ek plain-text vision response nahi deta).
async function ocrImageBuffer(buffer, mime) {
  const dataUrl = `data:${mime};base64,${buffer.toString('base64')}`;
  const upstream = await fetch(`${OPENROUTER_BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${OPENROUTER_API_KEY}` },
    body: JSON.stringify({
      model: OPENROUTER_VISION_MODEL,
      stream: false,
      messages: [
        { role: 'system', content: 'Is image mein jo bhi text likha hai wo bilkul verbatim nikaal ke do — line breaks jitna ho sake wahi rakho. Sirf extracted text likho, koi extra comment nahi. Agar text na ho to sirf likho: (no text found)' },
        { role: 'user', content: [{ type: 'image_url', image_url: { url: dataUrl } }] },
      ],
    }),
  });
  if (!upstream.ok) throw new Error(`OpenRouter HTTP ${upstream.status}`);
  const data = await upstream.json();
  return data.choices?.[0]?.message?.content || '';
}

async function ocrPdfToDocxBuffer(buffer) {
  if (!OPENROUTER_API_KEY) throw new Error('NO_API_KEY');
  const images = await renderPdfPagesToImages(buffer, { format: 'jpeg', quality: 0.9, scale: 2 });
  const pageTexts = [];
  for (const img of images) {
    try { pageTexts.push((await ocrImageBuffer(img, 'image/jpeg')).trim()); }
    catch { pageTexts.push('(is page se text nahi nikal paaye)'); }
  }
  const fullText = pageTexts.map((t, i) => `--- Page ${i + 1} ---\n${t}`).join('\n\n');
  if (!fullText.trim()) throw new Error('NO_TEXT');
  return generateDocxBuffer(fullText);
}

app.post('/api/pdf-tools/jpg-to-pdf', requireAuth, withPdfToolUpload(pdfToolUpload.array('files', 20), async (req, res) => {
  const files = req.files || [];
  if (!files.length) return res.status(400).json({ error: 'Koi image nahi mili' });
  const invalid = files.find(f => f.mimetype !== 'image/png' && f.mimetype !== 'image/jpeg');
  if (invalid) return res.status(400).json({ error: 'Sirf PNG ya JPEG images allowed hain' });
  try {
    const buffer = await generateImagePdfBuffer(files.map(f => ({ buffer: f.buffer, mime: f.mimetype })));
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_jpg_to_pdf', { count: files.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('jpg-to-pdf error:', err);
    res.status(500).json({ error: 'PDF banane mein error aayi' });
  }
}));

app.post('/api/pdf-tools/watermark', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const text = String(req.body.text || '').slice(0, 60).trim();
  if (!text) return res.status(400).json({ error: 'Watermark text likho' });
  const opacity = Math.min(1, Math.max(0.05, Number(req.body.opacity) || 0.3));
  const rotation = Math.min(90, Math.max(-90, Number(req.body.rotation) || -45));
  const fontSize = Math.min(120, Math.max(10, Number(req.body.fontSize) || 48));
  const color = /^#[0-9a-fA-F]{6}$/.test(req.body.color || '') ? req.body.color : '#888888';
  try {
    const buffer = await watermarkPdf(req.file.buffer, { text, opacity, rotation, fontSize, color });
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_watermark', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('watermark error:', err);
    res.status(500).json({ error: 'Watermark lagane mein error aayi. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/rotate', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const angle = [90, 180, 270].includes(Number(req.body.angle)) ? Number(req.body.angle) : 90;
  try {
    const buffer = await rotatePdfBuffer(req.file.buffer, angle);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_rotate', { angle });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('rotate error:', err);
    res.status(500).json({ error: 'Rotate nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/page-numbers', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const position = ['bottom-left', 'bottom-center', 'bottom-right', 'top-center'].includes(req.body.position) ? req.body.position : 'bottom-center';
  const startNumber = Math.max(1, Math.min(9999, parseInt(req.body.startNumber, 10) || 1));
  try {
    const buffer = await addPageNumbersToPdf(req.file.buffer, { position, startNumber });
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_page_numbers', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('page-numbers error:', err);
    res.status(500).json({ error: 'Page numbers add nahi ho paaye. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/repair', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const buffer = await repairPdfBuffer(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_repair', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('repair error:', err);
    res.status(500).json({ error: 'Ye file repair nahi ho paayi — damage bahut zyada hai' });
  }
}));

app.post('/api/pdf-tools/pdf-to-pdfa', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const buffer = await pdfToPdfABestEffort(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_to_pdfa', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('pdf-to-pdfa error:', err);
    res.status(500).json({ error: 'Convert nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/unlock', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const password = String(req.body.password || '');
  if (!password) return res.status(400).json({ error: 'PDF ka password daalo' });
  try {
    const buffer = await unlockPdfBuffer(req.file.buffer, password);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_unlock', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    res.status(400).json({ error: 'Password galat hai ya file corrupt hai' });
  }
}));

app.post('/api/pdf-tools/protect', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const password = String(req.body.password || '');
  if (password.length < 4) return res.status(400).json({ error: 'Password kam se kam 4 characters ka hona chahiye' });
  try {
    const buffer = await protectPdfBuffer(req.file.buffer, password);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_protect', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('protect error:', err);
    res.status(500).json({ error: 'Password lagane mein error aayi. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/organize', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  let pageOrder;
  try { pageOrder = JSON.parse(req.body.pageOrder || '[]'); } catch { return res.status(400).json({ error: 'Page order invalid hai' }); }
  if (!Array.isArray(pageOrder) || !pageOrder.length || !pageOrder.every(n => Number.isInteger(n) && n >= 0)) {
    return res.status(400).json({ error: 'Kam se kam ek page rakho' });
  }
  try {
    const buffer = await organizePdfBuffer(req.file.buffer, pageOrder);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_organize', { pageCount: pageOrder.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('organize error:', err);
    res.status(500).json({ error: 'Organize nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/compare', requireAuth, withPdfToolUpload(pdfToolUpload.fields([{ name: 'fileA', maxCount: 1 }, { name: 'fileB', maxCount: 1 }]), async (req, res) => {
  const fileA = req.files?.fileA?.[0], fileB = req.files?.fileB?.[0];
  if (!fileA || !fileB) return res.status(400).json({ error: 'Dono PDF files upload karo' });
  const errA = pdfMimeCheck(fileA), errB = pdfMimeCheck(fileB);
  if (errA || errB) return res.status(400).json({ error: 'Sirf PDF files allowed hain' });
  try {
    const result = await comparePdfTexts(fileA.buffer, fileB.buffer);
    logActivity(req, 'pdf_compare', {});
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('compare error:', err);
    res.status(500).json({ error: 'Compare nahi ho paaya. Kya dono valid PDF hain?' });
  }
}));

app.post('/api/pdf-tools/pdf-to-jpg', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const { buffer, count } = await pdfToJpgZip(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.zip`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_to_jpg', { pageCount: count });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, pageCount: count });
  } catch (err) {
    console.error('pdf-to-jpg error:', err);
    res.status(500).json({ error: 'JPG banane mein error aayi' });
  }
}));

const MAX_HTML_TO_PDF_CHARS = 300000;
app.post('/api/pdf-tools/html-to-pdf', requireAuth, async (req, res) => {
  const html = typeof req.body.html === 'string' ? req.body.html : '';
  if (!html.trim()) return res.status(400).json({ error: 'HTML ya URL daalo' });
  if (html.length > MAX_HTML_TO_PDF_CHARS) return res.status(400).json({ error: 'Content bahut bada hai' });
  try {
    const buffer = await htmlToPdfBuffer(html);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'html_to_pdf', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('html-to-pdf error:', err);
    res.status(400).json({ error: err.message && /allowed nahi|invalid hai|resolve nahi/.test(err.message) ? err.message : 'PDF banane mein error aayi. Kya HTML/URL valid hai?' });
  }
});

app.post('/api/pdf-tools/ocr', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  if (!OPENROUTER_API_KEY) return res.status(500).json({ error: 'OpenRouter API key server par configure nahi hai' });
  try {
    const buffer = await ocrPdfToDocxBuffer(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.docx`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_ocr', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    if (err.message === 'NO_TEXT') return res.status(400).json({ error: 'Is PDF ke images se koi text nahi nikal paaye' });
    console.error('ocr error:', err);
    res.status(500).json({ error: 'OCR nahi ho paaya. Dobara try karo' });
  }
}));

// ─── PDF TOOLS (round 3) — Crop, PDF Forms, Redact, AI Summarizer,
//     Translate PDF, PDF to Markdown ───────────────────────────────

async function queryOpenRouterText(systemPrompt, userContent) {
  const callModel = async (model) => fetch(`${OPENROUTER_BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${OPENROUTER_API_KEY}` },
    body: JSON.stringify({ model, stream: false, messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userContent }] }),
  });
  let upstream = await callModel(OPENROUTER_MODEL);
  if ((upstream.status === 429 || upstream.status === 404) && OPENROUTER_FALLBACK_MODEL !== OPENROUTER_MODEL) {
    upstream = await callModel(OPENROUTER_FALLBACK_MODEL);
  }
  if (!upstream.ok) throw new Error(`OpenRouter HTTP ${upstream.status}`);
  const data = await upstream.json();
  return data.choices?.[0]?.message?.content || '';
}

const MAX_AI_TEXT_TOOL_CHARS = 30000;

async function summarizePdfText(text) {
  return queryOpenRouterText(
    'Tum ek summarization assistant ho. Diye gaye document ka concise, clear summary do — pehle ek short paragraph, phir key points bullet mein. Hinglish/Hindi document ho to usi tone mein jawab do, warna English mein. Koi preamble mat likho ("Here is a summary:" jaisa) — seedha summary se shuru karo.',
    text.slice(0, MAX_AI_TEXT_TOOL_CHARS)
  );
}

async function translatePdfText(text, targetLang) {
  return queryOpenRouterText(
    `Tum ek translation assistant ho. Diya gaya text ko ${targetLang} mein translate karo — meaning bilkul accurate rakho, paragraph structure jitna ho sake preserve karo. Sirf translated text likho, koi extra comment ya preamble nahi.`,
    text.slice(0, MAX_AI_TEXT_TOOL_CHARS)
  );
}

async function pdfToMarkdownBuffer(buffer) {
  const pages = await extractPdfPages(buffer);
  const md = pages.map((p, i) => `## Page ${i + 1}\n\n${p}`).join('\n\n---\n\n').trim();
  if (!md) throw new Error('NO_TEXT');
  return Buffer.from(md, 'utf8');
}

async function cropPdfBuffer(buffer, { top, bottom, left, right }) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  for (const page of doc.getPages()) {
    const { width, height } = page.getSize();
    const mTop = height * top / 100, mBottom = height * bottom / 100;
    const mLeft = width * left / 100, mRight = width * right / 100;
    page.setCropBox(mLeft, mBottom, Math.max(1, width - mLeft - mRight), Math.max(1, height - mTop - mBottom));
  }
  return Buffer.from(await doc.save());
}

// PDF Forms — pdf-lib se existing AcroForm fields padh/fill kar sakte hain,
// ya (agar koi field hi nahi hai) naye clickable text fields add kar sakte
// hain (real fillable AcroForm fields banate hain, koi image/overlay hack nahi).
async function detectPdfFormFields(buffer) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const form = doc.getForm();
  return form.getFields().map(f => ({ name: f.getName(), type: f.constructor.name }));
}

async function fillPdfForm(buffer, values) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const form = doc.getForm();
  for (const [name, value] of Object.entries(values)) {
    try {
      const field = form.getField(name);
      const ctorName = field.constructor.name;
      if (ctorName === 'PDFTextField') field.setText(String(value ?? ''));
      else if (ctorName === 'PDFCheckBox') { if (value) field.check(); else field.uncheck(); }
      else if (ctorName === 'PDFDropdown' || ctorName === 'PDFRadioGroup' || ctorName === 'PDFOptionList') field.select(String(value));
    } catch { /* field na mile ya value type mismatch ho to skip karo, baaki fields fill hote rahein */ }
  }
  form.flatten();
  return Buffer.from(await doc.save());
}

async function createPdfFormFields(buffer, placements) {
  const doc = await PdfLibDocument.load(buffer, { ignoreEncryption: true });
  const form = doc.getForm();
  placements.forEach((p, i) => {
    const page = doc.getPage(p.page);
    const field = form.createTextField(p.name || `field_${i + 1}`);
    field.addToPage(page, { x: p.x, y: p.y, width: p.width, height: p.height, borderWidth: 1 });
  });
  return Buffer.from(await doc.save());
}

// Redact PDF — sirf ek black box upar draw karna security ke liye kaafi
// nahi hai (text neeche se copy-paste ho sakta hai) — isliye poora page
// hi rasterize (flatten to image) karte hain redaction boxes ke saath,
// taaki result mein koi selectable/extractable text hi na bache.
async function redactPdfToImages(buffer, redactionsByPage) {
  const scale = 1.8;
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.setContent('<!doctype html><html><body></body></html>');
    await page.addScriptTag({ url: `${PDFJS_CDN_BASE}/pdf.min.js` });
    const base64 = buffer.toString('base64');
    const dataUrls = await page.evaluate(async ({ base64, scale, workerSrc, redactionsByPage }) => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const doc = await window.pdfjsLib.getDocument({ data: bytes }).promise;
      const out = [];
      for (let i = 1; i <= doc.numPages; i++) {
        const p = await doc.getPage(i);
        const viewport = p.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width; canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        await p.render({ canvasContext: ctx, viewport }).promise;
        ctx.fillStyle = '#000';
        for (const r of (redactionsByPage[i - 1] || [])) {
          const xPx = r.x * scale;
          const yPxTop = canvas.height - (r.y + r.height) * scale;
          ctx.fillRect(xPx, yPxTop, r.width * scale, r.height * scale);
        }
        out.push(canvas.toDataURL('image/png'));
      }
      return out;
    }, { base64, scale, workerSrc: `${PDFJS_CDN_BASE}/pdf.worker.min.js`, redactionsByPage });
    return dataUrls.map(u => Buffer.from(u.split(',')[1], 'base64'));
  } finally {
    await browser.close();
  }
}

app.post('/api/pdf-tools/crop', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  const clamp = v => Math.min(45, Math.max(0, Number(v) || 0));
  const top = clamp(req.body.top), bottom = clamp(req.body.bottom), left = clamp(req.body.left), right = clamp(req.body.right);
  try {
    const buffer = await cropPdfBuffer(req.file.buffer, { top, bottom, left, right });
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_crop', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('crop error:', err);
    res.status(500).json({ error: 'Crop nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/summarize', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  if (!OPENROUTER_API_KEY) return res.status(500).json({ error: 'OpenRouter API key server par configure nahi hai' });
  try {
    const pages = await extractPdfPages(req.file.buffer);
    const text = pages.join('\n\n').trim();
    if (!text) return res.status(400).json({ error: 'Is PDF mein koi text nahi mila' });
    const summary = await summarizePdfText(text);
    logActivity(req, 'pdf_summarize', {});
    res.json({ success: true, summary });
  } catch (err) {
    console.error('summarize error:', err);
    res.status(500).json({ error: 'Summarize nahi ho paaya. Dobara try karo' });
  }
}));

const PDF_TRANSLATE_LANGS = ['Hindi', 'English', 'Spanish', 'French', 'German', 'Arabic', 'Chinese (Simplified)', 'Japanese', 'Portuguese', 'Russian'];
app.post('/api/pdf-tools/translate', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  if (!OPENROUTER_API_KEY) return res.status(500).json({ error: 'OpenRouter API key server par configure nahi hai' });
  const targetLang = PDF_TRANSLATE_LANGS.includes(req.body.targetLang) ? req.body.targetLang : 'English';
  try {
    const pages = await extractPdfPages(req.file.buffer);
    const text = pages.join('\n\n').trim();
    if (!text) return res.status(400).json({ error: 'Is PDF mein koi text nahi mila' });
    const translated = await translatePdfText(text, targetLang);
    if (!translated.trim()) throw new Error('NO_TEXT');
    // DOCX output — pdfkit ke standard fonts sirf WinAnsi (Latin-script)
    // characters render kar sakte hain, isliye Hindi/Arabic/Chinese/Japanese/
    // Russian jaisi target languages ke liye PDF output silently khaali ban
    // jaata (verified: sanitizePdfText Devanagari poori tarah strip kar deta
    // hai). DOCX/OOXML mein aisi koi font-encoding limitation nahi hai.
    const buffer = await generateDocxBuffer(translated);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.docx`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_translate', { targetLang });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('translate error:', err);
    res.status(500).json({ error: 'Translate nahi ho paaya. Dobara try karo' });
  }
}));

app.post('/api/pdf-tools/to-markdown', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const buffer = await pdfToMarkdownBuffer(req.file.buffer);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.md`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_to_markdown', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    if (err.message === 'NO_TEXT') return res.status(400).json({ error: 'Is PDF mein koi text nahi mila' });
    console.error('to-markdown error:', err);
    res.status(500).json({ error: 'Markdown banane mein error aayi' });
  }
}));

app.post('/api/pdf-tools/forms/detect', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  try {
    const fields = await detectPdfFormFields(req.file.buffer);
    res.json({ success: true, fields });
  } catch (err) {
    console.error('forms/detect error:', err);
    res.status(500).json({ error: 'Form fields detect nahi ho paaye. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/forms/fill', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  let values;
  try { values = JSON.parse(req.body.values || '{}'); } catch { return res.status(400).json({ error: 'Values invalid hain' }); }
  try {
    const buffer = await fillPdfForm(req.file.buffer, values);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_forms_fill', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('forms/fill error:', err);
    res.status(500).json({ error: 'Form fill nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/forms/create', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  let placements;
  try { placements = JSON.parse(req.body.placements || '[]'); } catch { return res.status(400).json({ error: 'Fields invalid hain' }); }
  if (!Array.isArray(placements) || !placements.length) return res.status(400).json({ error: 'Kam se kam ek field place karo' });
  try {
    const buffer = await createPdfFormFields(req.file.buffer, placements);
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_forms_create', { count: placements.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('forms/create error:', err);
    res.status(500).json({ error: 'Form fields add nahi ho paaye. Kya ye ek valid PDF hai?' });
  }
}));

app.post('/api/pdf-tools/redact', requireAuth, withPdfToolUpload(pdfToolUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi PDF file nahi mili' });
  const errMsg = pdfMimeCheck(req.file);
  if (errMsg) return res.status(400).json({ error: errMsg });
  let redactionsByPage;
  try { redactionsByPage = JSON.parse(req.body.redactionsByPage || '[]'); } catch { return res.status(400).json({ error: 'Redaction areas invalid hain' }); }
  if (!Array.isArray(redactionsByPage) || !redactionsByPage.some(r => Array.isArray(r) && r.length)) {
    return res.status(400).json({ error: 'Kam se kam ek area select karo' });
  }
  try {
    const images = await redactPdfToImages(req.file.buffer, redactionsByPage);
    const buffer = await generateImagePdfBuffer(images.map(b => ({ buffer: b, mime: 'image/png' })));
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_pt.pdf`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), buffer);
    logActivity(req, 'pdf_redact', {});
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName });
  } catch (err) {
    console.error('redact error:', err);
    res.status(500).json({ error: 'Redact nahi ho paaya. Kya ye ek valid PDF hai?' });
  }
}));

// ─── PDF TOOLS WORKFLOWS — kai tools ko ek saved chain mein jod ke ek hi
//     click mein sabko sequence mein run karo (iLovePDF ke "Workflow"
//     feature jaisa). Har entry ka `run` upar-defined generator functions
//     ko hi seedha reuse karta hai — koi logic duplicate nahi hua.
// Sirf wahi tools shamil hain jo "ek buffer/images in -> ek buffer out"
// hain, bina kisi per-page interactive click ke (Edit/Sign/Organize/Redact/
// Forms/Compare in sabko interactive canvas chahiye, isliye workflow steps
// mein fit nahi hote — v1 scope se bahar rakha hai).
const PDFT_MARGIN_FIELDS = [
  { id: 'top', label: 'Top Margin', type: 'range', default: 0, min: 0, max: 40, step: 1, unit: '%' },
  { id: 'bottom', label: 'Bottom Margin', type: 'range', default: 0, min: 0, max: 40, step: 1, unit: '%' },
  { id: 'left', label: 'Left Margin', type: 'range', default: 0, min: 0, max: 40, step: 1, unit: '%' },
  { id: 'right', label: 'Right Margin', type: 'range', default: 0, min: 0, max: 40, step: 1, unit: '%' },
];
const WORKFLOW_TOOLS = {
  'jpg-to-pdf': { label: 'JPG to PDF', category: 'Organize', inputType: 'image', outputType: 'pdf', extraFields: [],
    run: async (images) => generateImagePdfBuffer(images) },
  split: { label: 'Split PDF', category: 'Organize', inputType: 'pdf', outputType: 'zip', extraFields: [],
    run: async (buffer) => (await splitPdfToZip(buffer)).buffer },
  compress: { label: 'Compress PDF', category: 'Optimize', inputType: 'pdf', outputType: 'pdf', extraFields: [],
    run: async (buffer) => compressPdf(buffer) },
  repair: { label: 'Repair PDF', category: 'Optimize', inputType: 'pdf', outputType: 'pdf', extraFields: [],
    run: async (buffer) => repairPdfBuffer(buffer) },
  ocr: { label: 'OCR PDF', category: 'Optimize', inputType: 'pdf', outputType: 'docx', extraFields: [],
    run: async (buffer) => { if (!OPENROUTER_API_KEY) throw new Error('OpenRouter API key configure nahi hai'); return ocrPdfToDocxBuffer(buffer); } },
  rotate: { label: 'Rotate PDF', category: 'Edit', inputType: 'pdf', outputType: 'pdf',
    extraFields: [{ id: 'angle', label: 'Rotate Karo', type: 'select', default: '90', options: [{ value: '90', label: '90° (Clockwise)' }, { value: '180', label: '180°' }, { value: '270', label: '270° (Anti-clockwise)' }] }],
    run: async (buffer, opts) => rotatePdfBuffer(buffer, [90, 180, 270].includes(Number(opts.angle)) ? Number(opts.angle) : 90) },
  watermark: { label: 'Watermark', category: 'Edit', inputType: 'pdf', outputType: 'pdf',
    extraFields: [
      { id: 'text', label: 'Watermark Text', type: 'text', default: 'CONFIDENTIAL', maxlength: 60 },
      { id: 'opacity', label: 'Opacity', type: 'range', default: 30, min: 5, max: 100, step: 5, unit: '%' },
      { id: 'rotation', label: 'Rotation', type: 'range', default: -45, min: -90, max: 90, step: 5, unit: '°' },
      { id: 'fontSize', label: 'Font Size', type: 'range', default: 48, min: 10, max: 120, step: 2, unit: 'px' },
      { id: 'color', label: 'Color', type: 'color', default: '#888888' },
    ],
    run: async (buffer, opts) => watermarkPdf(buffer, {
      text: String(opts.text || 'CONFIDENTIAL').slice(0, 60).trim() || 'CONFIDENTIAL',
      opacity: Math.min(1, Math.max(0.05, Number(opts.opacity) / 100 || 0.3)),
      rotation: Math.min(90, Math.max(-90, Number(opts.rotation) || -45)),
      fontSize: Math.min(120, Math.max(10, Number(opts.fontSize) || 48)),
      color: /^#[0-9a-fA-F]{6}$/.test(opts.color || '') ? opts.color : '#888888',
    }) },
  'page-numbers': { label: 'Page Numbers', category: 'Edit', inputType: 'pdf', outputType: 'pdf',
    extraFields: [
      { id: 'position', label: 'Position', type: 'select', default: 'bottom-center', options: [
        { value: 'bottom-center', label: 'Bottom Center' }, { value: 'bottom-left', label: 'Bottom Left' },
        { value: 'bottom-right', label: 'Bottom Right' }, { value: 'top-center', label: 'Top Center' },
      ] },
      { id: 'startNumber', label: 'Start Number', type: 'number', default: 1, min: 1, max: 9999 },
    ],
    run: async (buffer, opts) => addPageNumbersToPdf(buffer, {
      position: ['bottom-left', 'bottom-center', 'bottom-right', 'top-center'].includes(opts.position) ? opts.position : 'bottom-center',
      startNumber: Math.max(1, Math.min(9999, parseInt(opts.startNumber, 10) || 1)),
    }) },
  crop: { label: 'Crop PDF', category: 'Organize', inputType: 'pdf', outputType: 'pdf', extraFields: PDFT_MARGIN_FIELDS,
    run: async (buffer, opts) => {
      const clamp = v => Math.min(45, Math.max(0, Number(v) || 0));
      return cropPdfBuffer(buffer, { top: clamp(opts.top), bottom: clamp(opts.bottom), left: clamp(opts.left), right: clamp(opts.right) });
    } },
  protect: { label: 'Protect PDF', category: 'Security', inputType: 'pdf', outputType: 'pdf',
    extraFields: [{ id: 'password', label: 'Naya Password Set Karo', type: 'password', default: '' }],
    run: async (buffer, opts) => {
      const password = String(opts.password || '');
      if (password.length < 4) throw new Error('Password kam se kam 4 characters ka hona chahiye');
      return protectPdfBuffer(buffer, password);
    } },
  unlock: { label: 'Unlock PDF', category: 'Security', inputType: 'pdf', outputType: 'pdf',
    extraFields: [{ id: 'password', label: 'PDF Ka Current Password', type: 'password', default: '' }],
    run: async (buffer, opts) => {
      const password = String(opts.password || '');
      if (!password) throw new Error('PDF ka password daalo');
      return unlockPdfBuffer(buffer, password);
    } },
  'pdf-to-pdfa': { label: 'PDF to PDF/A', category: 'Convert', inputType: 'pdf', outputType: 'pdf', extraFields: [],
    run: async (buffer) => pdfToPdfABestEffort(buffer) },
  'pdf-to-word': { label: 'PDF to Word', category: 'Convert', inputType: 'pdf', outputType: 'docx', extraFields: [],
    run: async (buffer) => generatePdfToDocxBuffer(buffer) },
  'pdf-to-pptx': { label: 'PDF to PowerPoint', category: 'Convert', inputType: 'pdf', outputType: 'pptx', extraFields: [],
    run: async (buffer) => generatePdfToPptxBuffer(buffer) },
  'pdf-to-excel': { label: 'PDF to Excel', category: 'Convert', inputType: 'pdf', outputType: 'xlsx', extraFields: [],
    run: async (buffer) => generatePdfToXlsxBuffer(buffer) },
  'word-to-pdf': { label: 'Word to PDF', category: 'Convert', inputType: 'docx', outputType: 'pdf', extraFields: [],
    run: async (buffer) => generateDocxToPdfBuffer(buffer) },
  'pptx-to-pdf': { label: 'PowerPoint to PDF', category: 'Convert', inputType: 'pptx', outputType: 'pdf', extraFields: [],
    run: async (buffer) => generatePptxToPdfBuffer(buffer) },
  'excel-to-pdf': { label: 'Excel to PDF', category: 'Convert', inputType: 'xlsx', outputType: 'pdf', extraFields: [],
    run: async (buffer) => generateXlsxToPdfBuffer(buffer) },
  'pdf-to-jpg': { label: 'PDF to JPG', category: 'Convert', inputType: 'pdf', outputType: 'zip', extraFields: [],
    run: async (buffer) => (await pdfToJpgZip(buffer)).buffer },
  'to-markdown': { label: 'PDF to Markdown', category: 'Convert', inputType: 'pdf', outputType: 'md', extraFields: [],
    run: async (buffer) => pdfToMarkdownBuffer(buffer) },
  translate: { label: 'Translate PDF', category: 'Intelligence', inputType: 'pdf', outputType: 'docx',
    extraFields: [{ id: 'targetLang', label: 'Target Language', type: 'select', default: 'Hindi', options: PDF_TRANSLATE_LANGS.map(l => ({ value: l, label: l })) }],
    run: async (buffer, opts) => {
      if (!OPENROUTER_API_KEY) throw new Error('OpenRouter API key configure nahi hai');
      const targetLang = PDF_TRANSLATE_LANGS.includes(opts.targetLang) ? opts.targetLang : 'English';
      const pages = await extractPdfPages(buffer);
      const text = pages.join('\n\n').trim();
      if (!text) throw new Error('Is PDF mein koi text nahi mila');
      const translated = await translatePdfText(text, targetLang);
      if (!translated.trim()) throw new Error('Translate fail ho gaya');
      return generateDocxBuffer(translated);
    } },
};
const WORKFLOW_OUTPUT_EXT = { pdf: 'pdf', docx: 'docx', pptx: 'pptx', xlsx: 'xlsx', md: 'md', zip: 'zip' };

app.get('/api/workflows/tools', requireAuth, (req, res) => {
  const tools = Object.entries(WORKFLOW_TOOLS).map(([key, t]) => ({
    key, label: t.label, category: t.category, inputType: t.inputType, outputType: t.outputType, extraFields: t.extraFields,
  }));
  res.json({ success: true, tools });
});

app.get('/api/workflows', requireAuth, async (req, res) => {
  try {
    const workflows = await Workflow.find({ user: req.user.id }).sort({ createdAt: -1 }).lean();
    res.json({ success: true, workflows: workflows.map(w => ({ id: w._id, name: w.name, steps: w.steps, createdAt: w.createdAt })) });
  } catch (err) {
    console.error('workflows list error:', err);
    res.status(500).json({ error: 'Workflows load nahi ho paaye' });
  }
});

app.post('/api/workflows', requireAuth, async (req, res) => {
  const name = String(req.body.name || '').trim().slice(0, 80);
  const steps = Array.isArray(req.body.steps) ? req.body.steps : [];
  if (!name) return res.status(400).json({ error: 'Workflow ka naam likho' });
  if (!steps.length) return res.status(400).json({ error: 'Kam se kam ek step add karo' });
  if (steps.length > 10) return res.status(400).json({ error: 'Ek workflow mein max 10 steps allowed hain' });

  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    if (!step || typeof step.tool !== 'string' || !WORKFLOW_TOOLS[step.tool]) {
      return res.status(400).json({ error: `Step ${i + 1} mein invalid tool hai` });
    }
    if (i > 0) {
      const prevOut = WORKFLOW_TOOLS[steps[i - 1].tool].outputType;
      const curIn = WORKFLOW_TOOLS[step.tool].inputType;
      if (prevOut !== curIn) {
        return res.status(400).json({ error: `Step ${i + 1} (${WORKFLOW_TOOLS[step.tool].label}) step ${i} ke output (${prevOut}) ke saath compatible nahi hai` });
      }
    }
  }

  try {
    const cleanSteps = steps.map(s => ({ tool: s.tool, options: s.options && typeof s.options === 'object' ? s.options : {} }));
    const workflow = await Workflow.create({ user: req.user.id, name, steps: cleanSteps });
    logActivity(req, 'workflow_create', { steps: cleanSteps.length });
    res.json({ success: true, workflow: { id: workflow._id, name: workflow.name, steps: workflow.steps, createdAt: workflow.createdAt } });
  } catch (err) {
    console.error('workflow create error:', err);
    res.status(500).json({ error: 'Workflow save nahi ho paya' });
  }
});

app.delete('/api/workflows/:id', requireAuth, async (req, res) => {
  try {
    const result = await Workflow.deleteOne({ _id: req.params.id, user: req.user.id });
    if (!result.deletedCount) return res.status(404).json({ error: 'Workflow nahi mila' });
    res.json({ success: true });
  } catch (err) {
    console.error('workflow delete error:', err);
    res.status(500).json({ error: 'Workflow delete nahi ho paya' });
  }
});

app.post('/api/workflows/:id/run', requireAuth, withPdfToolUpload(pdfToolUpload.array('files', 20), async (req, res) => {
  const files = req.files || [];
  if (!files.length) return res.status(400).json({ error: 'File(s) upload karo' });

  let workflow;
  try {
    workflow = await Workflow.findOne({ _id: req.params.id, user: req.user.id });
  } catch {
    return res.status(404).json({ error: 'Workflow nahi mila' });
  }
  if (!workflow) return res.status(404).json({ error: 'Workflow nahi mila' });

  const firstTool = WORKFLOW_TOOLS[workflow.steps[0]?.tool];
  if (!firstTool) return res.status(400).json({ error: 'Workflow mein invalid tool hai' });

  try {
    let current = firstTool.inputType === 'image'
      ? files.map(f => ({ buffer: f.buffer, mime: f.mimetype }))
      : files[0].buffer;
    let outputType = firstTool.inputType;

    for (const step of workflow.steps) {
      const toolDef = WORKFLOW_TOOLS[step.tool];
      if (!toolDef) throw new Error(`Tool "${step.tool}" available nahi hai`);
      current = await toolDef.run(current, step.options || {});
      outputType = toolDef.outputType;
    }

    const ext = WORKFLOW_OUTPUT_EXT[outputType] || 'pdf';
    const uid = req.user.id;
    const outName = `${uid.slice(0, 8)}_${Date.now()}_wf.${ext}`;
    fs.writeFileSync(path.join(DOWNLOADS_DIR, outName), current);
    logActivity(req, 'workflow_run', { workflowId: String(workflow._id), steps: workflow.steps.length });
    res.json({ success: true, fileUrl: `/files/${outName}`, filename: outName, outputType });
  } catch (err) {
    console.error('workflow run error:', err);
    res.status(400).json({ error: err.message || 'Workflow run nahi ho paya' });
  }
}));

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
        { role: 'system', content: `Tum PastePro Assistant ho. Seedha, sahi aur helpful jawab do. Jab Hindi/Hinglish mein poocha jaaye, usi mein jawab do.

Agar user koi real downloadable file maange, use ek fenced code block mein is tarah do (sirf tabhi jab user specifically wo file-type maange, normal jawabon mein nahi):
- PDF document: \`\`\`pdf fenced block, content plain text/markdown jaisa (# heading, ## subheading, - bullet, **bold**).
- Word document: \`\`\`docx fenced block, wahi markdown-jaisa formatting.
- Excel/spreadsheet: \`\`\`xlsx fenced block mein CSV format do — comma-separated values, pehli row column headers honi chahiye.
- PowerPoint: \`\`\`pptx fenced block — har slide ko apni line \`---\` se alag karo, har slide ki pehli line uska title ho, baaki lines bullet points.
\`\`\`pdf/docx/xlsx/pptx blocks mein kabhi emoji ya koi bhi Unicode symbol mat daalo (jaise 📅, ✓, ★) — PDF generator sirf plain text/ASCII safely render karta hai, emoji daalne par wo garbled/corrupt dikhta hai.
Code files (Python/JS/HTML/CSS/etc.) ke liye normal fenced block hi use karo (\`\`\`python, \`\`\`js, wagera) — wo already downloadable hai, koi special format nahi chahiye.` },
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
      const text = (result.value || '').trim();
      if (text) return { text };
      // Sirf-image DOCX (jaise File Converter ka "image to Word" tool
      // banata hai) mammoth se koi text nahi deta — seedha zip ke andar se
      // embedded images nikaal ke preview mein dikha dete hain taaki
      // confusing "no text" message ki jagah asli content dikhe. Media
      // filenames content-hash based hote hain (image1/image2 jaisi
      // sequence nahi), isliye sirf sorting se sahi order guarantee nahi
      // hota — document.xml ke <a:blip r:embed="rIdN"/> references ko
      // unke relationships se resolve karke asli paragraph order milta hai.
      try {
        const zip = new AdmZip(buffer);
        const docXmlEntry = zip.getEntry('word/document.xml');
        const relsEntry = zip.getEntry('word/_rels/document.xml.rels');
        const docXml = docXmlEntry ? docXmlEntry.getData().toString('utf8') : '';
        const relsXml = relsEntry ? relsEntry.getData().toString('utf8') : '';
        const relMap = {};
        for (const m of relsXml.matchAll(/<Relationship[^>]*\sId="([^"]+)"[^>]*\sTarget="([^"]+)"/g)) relMap[m[1]] = m[2];
        const images = [...docXml.matchAll(/r:embed="([^"]+)"/g)]
          .map(m => relMap[m[1]])
          .filter(Boolean)
          .map(target => zip.getEntry(`word/${target}`))
          .filter(Boolean)
          .map(entry => {
            const ext = entry.entryName.split('.').pop().toLowerCase();
            const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'bmp' ? 'image/bmp' : 'image/jpeg';
            return `data:${mime};base64,${entry.getData().toString('base64')}`;
          });
        if (images.length) return { text: '', images };
      } catch { /* fall through */ }
      return { text: '' };
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
        // Is slide ka embedded image dhoondo (jaise File Converter ka "image
        // to PowerPoint" tool banata hai — har slide mein ek full-size image).
        let image = null;
        try {
          const slideFileName = e.entryName.split('/').pop();
          const relEntry = zip.getEntry(`ppt/slides/_rels/${slideFileName}.rels`);
          const relXml = relEntry ? relEntry.getData().toString('utf8') : '';
          const m = /Target="\.\.\/media\/([^"]+)"/.exec(relXml);
          const mediaEntry = m ? zip.getEntry(`ppt/media/${m[1]}`) : null;
          if (mediaEntry) {
            const ext = m[1].split('.').pop().toLowerCase();
            const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'bmp' ? 'image/bmp' : 'image/jpeg';
            image = `data:${mime};base64,${mediaEntry.getData().toString('base64')}`;
          }
        } catch { /* image na mile to text-only slide jaisa treat hota hai */ }
        return { number: i + 1, lines, image };
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
    // 404 bhi retry karte hain kyunki OpenRouter free-tier model slugs
    // kabhi-kabhi discontinue ho jaate hain (verified directly: is session
    // mein hi default OPENROUTER_MODEL retire ho chuka tha, aur upstream
    // isko rate-limit ki tarah 429 nahi, seedha 404 "model unavailable"
    // deta hai) — bina is check ke primary model retire hote hi poora
    // Playground text-chat broken ho jaata, sirf fallback model set karne
    // tak. No verified-working second vision model yet, so this retry only
    // applies to plain text queries.
    if ((upstream.status === 429 || upstream.status === 404) && !imageDataUrls.length && OPENROUTER_FALLBACK_MODEL !== OPENROUTER_MODEL) {
      console.log(`OpenRouter primary model unavailable (HTTP ${upstream.status}), retrying with fallback model...`);
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
      slides: result.slides, // pptx only — per-slide paragraph lines (+ embedded image) for a readable preview
      images: result.images, // docx only — sirf-image DOCX ke embedded images (jab koi text na ho)
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

// Music Player ka "Share" button ek link banata hai (?music=<videoId>) —
// jo bhi wo link khole, humein us EK specific video ka info chahiye hota hai
// (poori keyword search nahi) taaki seedha wahi track load karke play kiya
// ja sake.
app.get('/api/music/track-info', requireAuth, async (req, res) => {
  const videoId = String(req.query.videoId || '').trim();
  if (!/^[\w-]{5,20}$/.test(videoId)) return res.status(400).json({ error: 'Invalid video id' });

  const ytdlpCmd = await resolveYtdlpCommand();
  if (!ytdlpCmd) return res.status(500).json({ error: 'yt-dlp install nahi hai' });

  const { cmd, extraArgs } = splitYtdlpCommand(ytdlpCmd);
  const url = `https://www.youtube.com/watch?v=${videoId}`;
  const baseArgs    = ['--skip-download', '--dump-json', ...cookiesArgs(), ...proxyArgs(), url];
  const primaryArgs = [...extraArgs, ...baseArgs];
  const retryArgs   = [...extraArgs, '--extractor-args', 'youtube:player_client=visionos,tv_simply,ios,android,mweb,web_creator,web,tv;formats=missing_pot', ...baseArgs];

  const handle = (error, stdout, stderr) => {
    if (error) {
      console.error('Track-info error:', stderr || error.message);
      return res.status(500).json({ error: 'Ye video load nahi ho saka' });
    }
    try {
      const info = JSON.parse(stdout);
      const thumbs = info.thumbnails || [];
      const thumb  = thumbs.length ? thumbs[thumbs.length - 1].url : null;
      const track = {
        id: info.id || videoId,
        title: info.title || 'Untitled',
        channel: info.channel || info.uploader || '',
        duration: info.duration_string || '',
        views: info.view_count || 0,
        thumbnail: thumb,
        url: info.webpage_url || url,
      };
      logActivity(req, 'music_shared_track_open', { videoId });
      res.json({ success: true, track });
    } catch (e) {
      res.status(500).json({ error: 'Ye video load nahi ho saka' });
    }
  };

  execFile(cmd, primaryArgs, { timeout: 20 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 20 * 1024 * 1024 }, (error, stdout, stderr) => {
    const hitBotCheck = error && (stderr || '').toLowerCase().includes('sign in to confirm');
    if (hitBotCheck) {
      return execFile(cmd, retryArgs, { timeout: 20 * 1000, cwd: DOWNLOADS_DIR, maxBuffer: 20 * 1024 * 1024 }, (error2, stdout2, stderr2) => handle(error2, stdout2, stderr2));
    }
    handle(error, stdout, stderr);
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
