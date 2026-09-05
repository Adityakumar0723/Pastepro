# PastePro — Setup Guide (Hindi)

## Files jo tumhe mili hain:
- `index.html`  → Frontend (browser mein open karo)
- `server.js`   → Backend (Node.js server)
- `package.json` → Node dependencies

---

## STEP 1 — Local MongoDB setup

1. Install and start MongoDB Community Edition.
2. The app connects by default to:
   ```
   mongodb://127.0.0.1:27017/pastepro
   ```
3. Optional: copy `.env.example` to `.env` and set a strong `JWT_SECRET` or a different `MONGODB_URI`.

Email/password accounts are saved in MongoDB. Passwords are stored as bcrypt hashes, never plain text.

---

## STEP 2 — yt-dlp install karna (Video downloader)

### Windows:
```
winget install yt-dlp
```
ya https://github.com/yt-dlp/yt-dlp/releases se .exe download karo

Agar aapne `yt-dlp` global install nahi kiya hai, toh `yt-dlp.exe` ko is project root mein daal do.

Aap Windows mein alternate path bhi use kar sakte ho:
```powershell
setx YTDLP_PATH "C:\path\to\yt-dlp.exe"
```
Phir terminal band karke dobara open karo.

### Mac:
```
brew install yt-dlp
```

### Linux/Ubuntu:
```
pip install yt-dlp
```
ya
```
sudo wget https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp -O /usr/local/bin/yt-dlp
sudo chmod a+rx /usr/local/bin/yt-dlp
```

### Check karo:
```
yt-dlp --version
```

---

## STEP 3 — Backend chalu karna

```bash
# Backend folder mein jaao
cd pastepro-backend

# Dependencies install karo
npm install

# Server start karo
node server.js
```

Server `http://localhost:3002` pe start ho jayega.

---

## STEP 4 — Frontend open karna

`index.html` ko browser mein open karo (VS Code Live Server use karo ya directly double-click)

---

## STEP 5 — Test karo

1. Sign up with an email and password
2. YouTube link paste karo
3. "Video Download" click karo
4. Server process karega aur download link milega

---

## Production Deployment ke liye

Backend ko deploy karna hai toh:
- **Railway.app** (free tier available)
- **Render.com** (free tier)
- **VPS (DigitalOcean/Vultr)** — recommended for heavy use

Frontend ke liye:
- **Netlify** ya **Vercel** pe drag and drop karo

Backend deploy karne ke baad `index.html` mein yeh line update karo:
```js
const BACKEND = 'https://your-backend-url.railway.app';
```

---

## Important Notes

⚠️ **YouTube Terms of Service**: Personal use ke liye theek hai, commercial use mein copyright issues aa sakte hain.
⚠️ **File size limit**: 200MB set hai server.js mein, badha sakte ho.
⚠️ **Downloads 30 min baad delete ho jaate hain** automatically.

---

## Folder Structure

```
pastepro/
├── index.html                    ← Frontend
├── server.js                     ← Backend
├── package.json                  ← npm config
├── .env.example                  ← MongoDB URI and JWT settings
└── downloads/                    ← (auto-create hoga)
```
