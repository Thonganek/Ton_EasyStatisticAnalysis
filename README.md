# Easy Statistic Analysis Tools v2.0.0

Professional Statistical Analysis Suite for Researchers  
Developed by นายจรัญญู ทองเอนก (Jarunyoo Thonganek)

---

## สารบัญ

- [วิธีติดตั้ง (Local)](#วิธีติดตั้ง-local)
- [วิธีติดตั้งจาก GitHub](#วิธีติดตั้งจาก-github)
- [วิธีเผยแพร่ออนไลน์ (Deploy)](#วิธีเผยแพร่ออนไลน์-deploy)
- [วิธีแก้ไขและอัปเดต](#วิธีแก้ไขและอัปเดต)
- [วิธีตั้งค่า AI](#วิธีตั้งค่า-ai)
- [คุณสมบัติทั้งหมด](#คุณสมบัติทั้งหมด)
- [โครงสร้างโปรเจค](#โครงสร้างโปรเจค)
- [การใช้งานเบื้องต้น](#การใช้งานเบื้องต้น)
- [แก้ปัญหาที่พบบ่อย](#แก้ปัญหาที่พบบ่อย)
- [License](#license)

---

## วิธีติดตั้ง (Local)

### สิ่งที่ต้องมี

- **Node.js** v16 ขึ้นไป — ดาวน์โหลดที่ [nodejs.org](https://nodejs.org)
- **Browser** อะไรก็ได้ (Chrome, Firefox, Edge)

### ขั้นตอน

```bash
# 1. เปิด Terminal / Command Prompt

# 2. เข้าโฟลเดอร์โปรเจค
cd d:/python/EesyStatisticAnalysis/webapp

# 3. ติดตั้ง dependencies (ครั้งแรกครั้งเดียว)
npm install

# 4. รันแอป
npm start
```

เปิดเบราว์เซอร์ที่ **http://localhost:3000**

Login: `thankyou` / `1234`

### สำหรับ Windows

Double-click ไฟล์ `run.bat` ได้เลย

---

## วิธีติดตั้งจาก GitHub

### วิธีที่ 1: Git Clone

```bash
# 1. Clone โปรเจค
git clone https://github.com/Thonganek/EesyStatisticAnalysis.git

# 2. เข้าโฟลเดอร์
cd EesyStatisticAnalysis

# 3. ติดตั้ง dependencies
npm install

# 4. รัน
npm start
```

### วิธีที่ 2: Download ZIP (ไม่ต้องใช้ Git)

1. ไปที่ https://github.com/Thonganek/EesyStatisticAnalysis
2. กดปุ่มเขียว **Code** → **Download ZIP**
3. แตกไฟล์ ZIP
4. เปิด Terminal ในโฟลเดอร์ที่แตกออกมา
5. รัน:

```bash
npm install
npm start
```

---

## วิธีเผยแพร่ออนไลน์ (Deploy)

ให้คนอื่นเข้าใช้ผ่านเว็บได้เลย ไม่ต้องติดตั้งอะไร

### วิธีที่ 1: Render.com (แนะนำ, ฟรี)

#### ขั้นตอนที่ 1: สมัคร

1. เปิด https://render.com
2. กด **Get Started for Free**
3. เลือก **Sign up with GitHub**
4. อนุญาตให้ Render เข้าถึง GitHub

#### ขั้นตอนที่ 2: สร้าง Web Service

1. กดปุ่ม **New +** (มุมขวาบน) → **Web Service**
2. เลือก **Build and deploy from a Git repository** → Next
3. เลือก repo **Thonganek/EesyStatisticAnalysis** → **Connect**

#### ขั้นตอนที่ 3: ตั้งค่า

| ช่อง | ใส่ |
|------|-----|
| Name | `eesy-stat` |
| Region | Singapore (Southeast Asia) |
| Branch | `main` |
| Runtime | `Node` |
| Build Command | `npm install` |
| Start Command | `npm start` |
| Instance Type | **Free** |

#### ขั้นตอนที่ 4: Deploy

1. กด **Create Web Service**
2. รอ 2-3 นาที จนขึ้น **"Your service is live"**
3. ได้ลิงก์ เช่น `https://eesy-stat.onrender.com`
4. แชร์ลิงก์นี้ให้ใครก็ได้!

### วิธีที่ 2: Vercel (ฟรี, เร็วมาก)

1. ไปที่ https://vercel.com → Sign up with GitHub
2. กด **Import Project** → เลือก **EesyStatisticAnalysis**
3. กด **Deploy**
4. ได้ลิงก์ เช่น `https://eesy-stat.vercel.app`

### วิธีที่ 3: Railway.app (ฟรี 500 ชม./เดือน)

1. ไปที่ https://railway.app → Sign up with GitHub
2. กด **New Project** → **Deploy from GitHub repo**
3. เลือก **EesyStatisticAnalysis**
4. Railway deploy อัตโนมัติ → ได้ลิงก์

### เปรียบเทียบ

| | Render | Vercel | Railway |
|---|---|---|---|
| ราคา | ฟรี | ฟรี | ฟรี 500 ชม. |
| ความเร็ว | ปานกลาง | เร็วมาก | เร็ว |
| หลับ (sleep) | หลังไม่ใช้ 15 นาที | ไม่หลับ | ไม่หลับ |
| ง่าย | ง่ายมาก | ง่ายมาก | ง่าย |

---

## วิธีแก้ไขและอัปเดต

### ขั้นตอน

```bash
# 1. แก้ไขโค้ดตามต้องการ

# 2. เข้าโฟลเดอร์โปรเจค
cd d:/python/EesyStatisticAnalysis/webapp

# 3. Add + Commit + Push
git add .
git commit -m "อธิบายสิ่งที่แก้ไข"
git push
```

Render / Vercel / Railway จะ **auto deploy** ภายใน 2-3 นาที

### ถ้า git push แล้ว error เรื่อง token

1. สร้าง Personal Access Token ใหม่:
   - ไปที่ https://github.com/settings/tokens/new
   - Note: `EesyStat`
   - ติ๊ก **repo**
   - กด Generate token
   - คัดลอก token (ghp_...)

2. อัปเดต remote URL:

```bash
git remote set-url origin https://Thonganek:TOKEN_ใหม่@github.com/Thonganek/EesyStatisticAnalysis.git
git push
```

### วิธีอัปขึ้น GitHub ครั้งแรก (กรณีเริ่มใหม่)

```bash
cd d:/python/EesyStatisticAnalysis/webapp
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/Thonganek/EesyStatisticAnalysis.git
git branch -M main
git push -u origin main
```

---

## วิธีตั้งค่า AI

แอปใช้ **Google Gemini API** สำหรับ AI วิเคราะห์ผลและ AI Chat

### ขั้นตอน

1. ไปที่ https://aistudio.google.com/apikey
2. Login ด้วย Google Account
3. กด **Create API Key** (ฟรี)
4. คัดลอก API Key

### ตั้งค่าในแอป

1. เข้าแอป → เมนู **AI Settings**
2. วาง API Key ในช่อง
3. เลือก Model:
   - `gemini-2.5-flash-lite` — เร็ว ประหยัด
   - `gemini-2.5-flash` — สมดุล (แนะนำ)
   - `gemini-2.5-pro` — แม่นยำสูง
4. กด **Save Settings**
5. กด **Test Connection** เพื่อทดสอบ

### การใช้งาน AI

- **AI วิเคราะห์ผล** — กดปุ่ม "AI วิเคราะห์ผล" ในทุกหน้าวิเคราะห์ เลือกรูปแบบ:
  - บทที่ 4 เชิงวิชาการ (APA)
  - บทคัดย่อ สรุปสั้น
  - หมายเหตุใต้ตาราง
  - สไลด์นำเสนอ
  - อธิบายให้เข้าใจง่าย
- **AI Chat** — คลิก "AI Statistical Chat" ในเมนูซ้าย ถาม-ตอบเรื่องสถิติได้

---

## คุณสมบัติทั้งหมด

### Data Analysis (วิเคราะห์ข้อมูลเบื้องต้น)

| เครื่องมือ | รายละเอียด |
|-----------|-----------|
| Descriptive | Mean, S.D., S.E., 95% CI, Skewness, Kurtosis |
| Numeric | วิเคราะห์ตัวแปรเชิงปริมาณ |
| Nominal | ความถี่ ร้อยละ Cumulative% |
| Likert Scale | 5/3 ระดับ ตั้งค่าเกณฑ์แปลผลได้ จัดอันดับ |
| Interval | จัดกลุ่มช่วงต่อเนื่อง หลายตัวแปร ตั้งค่า popup |
| Outlier | ตรวจค่าผิดปกติด้วย IQR |
| Normality | Shapiro-Wilk, Kolmogorov-Smirnov |

### Parametric Tests (สถิติ Parametric)

| สถิติ | ใช้เมื่อ | ผลลัพธ์ |
|-------|---------|---------|
| Independent t-test | เปรียบเทียบ 2 กลุ่มอิสระ | t, df, Mean Diff, p, 95% CI, Cohen's d |
| Paired t-test | เปรียบเทียบก่อน-หลัง | t, df, Mean Diff, p, 95% CI, Cohen's d |
| One-way ANOVA | เปรียบเทียบ 3+ กลุ่ม | F, p, eta-squared, Post-hoc |
| Two-way ANOVA | 2 ปัจจัย + ปฏิสัมพันธ์ | F, p, eta-squared |
| RM-ANOVA | วัดซ้ำ 3+ ครั้ง | F, p, Pairwise |
| ANCOVA | ควบคุมตัวแปรร่วม | F, p, Adjusted Means |

### Non-Parametric Tests (สถิติ Non-Parametric)

| สถิติ | ใช้แทน | ผลลัพธ์ |
|-------|--------|---------|
| Mann-Whitney U | Independent t-test | U, Z, p, r, Mean Rank |
| Wilcoxon | Paired t-test | W, Z, p, r, Ranks |
| Kruskal-Wallis | One-way ANOVA | H, p, eta-squared, Post-hoc |
| Friedman | RM-ANOVA | Chi-square, p, Kendall's W |
| Chi-Square | ความสัมพันธ์ตัวแปรกลุ่ม | Chi-square, p, Cramer's V |

### Advanced Analysis

| เครื่องมือ | รายละเอียด |
|-----------|-----------|
| Correlation | Pearson, Spearman, Kendall + Heatmap |
| Linear Regression | R, R-squared, Durbin-Watson, VIF, Beta |
| Logistic Regression | OR, Wald, AIC, Classification Table |
| Assumption Tests | Normality, Levene, VIF |
| Reliability | Cronbach's Alpha, Item-Total, Alpha if Deleted |
| Effect Size | Cohen's d, Hedges' g, OR, RR + 95% CI |

### AI Features

| ฟีเจอร์ | รายละเอียด |
|---------|-----------|
| AI วิเคราะห์ผล | สรุปผลวิจัย 5 รูปแบบ (บทที่ 4, บทคัดย่อ, ฯลฯ) |
| AI Chat | ถาม-ตอบเรื่องสถิติ แนะนำการใช้สถิติ |
| Assumption Check | ตรวจข้อตกลงอัตโนมัติทุกสถิติ พร้อมคำแนะนำ |
| คำแนะนำ | ทุกหน้ามี "เมื่อไหร่ควรใช้" พร้อมตัวอย่าง |

### Export & Templates

| ฟีเจอร์ | รายละเอียด |
|---------|-----------|
| Export Excel | ดาวน์โหลดผลวิเคราะห์เป็น .xlsx |
| Export Word | ดาวน์โหลดเป็น .doc พร้อมผล AI |
| Sample Templates | ข้อมูลตัวอย่าง + กรณีศึกษา ทุกสถิติ |

---

## โครงสร้างโปรเจค

```
EesyStatisticAnalysis/
├── package.json          # npm dependencies
├── server.js             # Express server + AI proxy + Template API
├── run.bat               # Windows launcher (double-click)
├── README.md             # ไฟล์นี้
├── .gitignore            # ไฟล์ที่ไม่อัปขึ้น GitHub
├── public/
│   ├── index.html        # หน้าเว็บหลัก (Single Page App)
│   ├── qrcode.jpg        # QR Code LINE
│   ├── developer.png     # รูปผู้พัฒนา
│   ├── css/
│   │   └── style.css     # Stylesheet (Pastel Blue Theme)
│   └── js/
│       ├── app.js        # Logic หลัก (navigation, AI, export)
│       └── stats.js      # คำนวณสถิติทั้งหมด (client-side)
└── resources/
    ├── qrcode.jpg
    └── developer.png
```

### Tech Stack

| ส่วน | เทคโนโลยี |
|------|----------|
| Backend | Node.js + Express |
| Frontend | HTML + CSS + JavaScript (SPA) |
| Statistics | jStat + Custom implementations |
| Excel I/O | SheetJS (xlsx) |
| AI | Google Gemini API |

---

## การใช้งานเบื้องต้น

### 1. Login
- Username: `thankyou`
- Password: `1234`

### 2. Upload ข้อมูล
- คลิก "Upload Data (.xlsx)" ในเมนูซ้าย
- เลือกไฟล์ Excel (.xlsx)
- ระบบจะอ่านข้อมูลและแสดง preview

### 3. เลือกการวิเคราะห์
- เลือกเมนูสถิติที่ต้องการจาก sidebar
- กดปุ่ม "เลือกตัวแปร" → popup แสดงตัวแปร → กดเลือก → ตกลง
- กด "Run Analysis"

### 4. ดูผลลัพธ์
- ตาราง Assumption Check แสดงก่อน (ตรวจข้อตกลง)
- ตารางผลวิเคราะห์หลัก
- กดปุ่ม "AI วิเคราะห์ผล" เพื่อสรุปผลอัตโนมัติ
- กด Download Excel / Word เพื่อส่งออก

### 5. เลือกตัวแปร (Popup)
- กดปุ่ม "เลือกตัวแปร" → Popup เปิดขึ้น
- กดที่ชื่อตัวแปร → เปลี่ยนเป็นสีน้ำเงิน (เลือกแล้ว)
- กดอีกครั้ง → ยกเลิกการเลือก
- ปุ่ม "เลือกทั้งหมด" / "ยกเลิกทั้งหมด"
- กด "ตกลง" → ตัวแปรที่เลือกแสดงเป็น tag

---

## แก้ปัญหาที่พบบ่อย

### npm start แล้วไม่ขึ้น

```bash
# ตรวจสอบว่าติดตั้ง Node.js แล้ว
node --version

# ถ้ายังไม่ติดตั้ง ไปดาวน์โหลดที่ nodejs.org

# ติดตั้ง dependencies ใหม่
npm install
npm start
```

### Port 3000 ถูกใช้งานอยู่

```bash
# Windows: หา process ที่ใช้ port 3000
netstat -ano | findstr :3000

# Kill process (แทน PID ด้วยเลขที่เจอ)
taskkill /PID เลข /F

# รันใหม่
npm start
```

### AI ใช้งานไม่ได้

1. ตรวจสอบว่าใส่ API Key แล้ว (AI Settings)
2. กด Test Connection
3. ถ้า error: สร้าง API Key ใหม่ที่ https://aistudio.google.com/apikey
4. ตรวจสอบ internet connection

### git push error

```bash
# สร้าง token ใหม่ที่ github.com/settings/tokens/new (ติ๊ก repo)
git remote set-url origin https://Thonganek:TOKEN_ใหม่@github.com/Thonganek/EesyStatisticAnalysis.git
git push
```

---

## License

Copyright (c) 2023-2026 นายจรัญญู ทองเอนก (Jarunyoo Thonganek)  
All Rights Reserved
