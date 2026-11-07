# Pak QR Print — Online (GitHub Pages + Firebase + Supabase)

Online version of Pak QR Print. Customer kahi se bhi (internet se) file bhejta hai,
dukaan ke dashboard par order khud khul jata hai (Auto-Open) aur dukaan ke PC par
print agent us ko auto print karta hai.

Files Supabase Storage (muft, 1GB, koi card nahi) par jati hain; queue aur login
Firebase Realtime Database / Auth par hain (Firebase Storage istemal nahi hota).

## Structure

- `index.html` — Customer page (Standard Documents/Photos + ID Card 2-in-1)
- `dashboard.html` — Admin dashboard (login, Live Queue, Auto-Open, toggles)
- `qr-poster.html` — A4 QR poster (counter ke liye)
- `styles.css`, `customer.js`, `dashboard.js`, `qr.js` — frontend
- `firebase-config.js` — **TEMPLATE: apni Firebase config yahan bharna lazmi hai**
- `supabase-config.js` — **Supabase config (file storage, muft 1GB, koi card nahi) — pehle se bhari hui hai**
- `firebase-rules.txt` — Realtime Database rules (Console mein paste karein; is file ka Storage wala hissa ab istemal nahi hota — files Supabase par jati hain)
- `pdfjs/` — PDF preview ke liye (local copy, koi CDN nahi)
- `print-agent/` — Dukaan ke Windows PC par chalne wala print agent
  - `agent.js` — Firebase queue dekhta hai, file download karta hai, print karta hai
  - `printEngine.js` — Offline v1 ka azmaya hua print engine (SumatraPDF / Word COM / Shell / PowerPoint COM / N-up)
  - `config.json` — **TEMPLATE: apni Firebase + dukaan ki details yahan bharna lazmi hai**
  - `Start-Print-Agent.bat` — Agent chalane ke liye double-click
  - `SumatraPDF.exe` — PDF print ke liye (isi folder mein hona lazmi hai)

## Flow

Customer (kahin bhi, internet) → GitHub Pages (ye files) → Supabase (files) + Firebase (queue)
→ Shop PC (print agent, ON + internet + printer) → Printer → Kaagaz

Shop PC band ho ya internet na ho to kuch print nahi hoga — order Firebase mein
mehfooz rehta hai, PC on hote hi print ho jata hai.

NOTE: Supabase bucket ki policy mein anon ke liye sirf upload + parhna hai,
delete nahi — is liye purani files bucket se khud delete nahi hoti. Supabase
dashboard (Storage > pakqrprint-files) se waqt par saaf karein.

## Setup

Poori tafseel Roman Urdu mein: `Pak-QR-Print-GitHub-Hidayat.txt` dekhein.
Mukhtasar: 1) GitHub repo + Pages on karein, 2) Firebase project (free tier) banayein
+ config paste karein, 3) Shop PC par print agent chalayein, 4) QR poster print karein.

## File support

PDF, JPG/PNG, Word (.doc/.docx), Excel (.xls/.xlsx), PowerPoint (.ppt/.pptx) —
jaise offline version karta hai. Print dukaan ke PC par hota hai, is liye
PC par MS Office hona chahiye. PowerPoint: agar PC par PowerPoint na ho to file
auto print nahi hogi — order mehfooz rahega aur dashboard par wajah likhi aayegi
(manual print kar sakte hain). Is ki guarantee nahi di jati.
