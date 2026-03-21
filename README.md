# PondPhuwinPermpoon Schedule

เว็บไซต์ตารางงานและกิจกรรมของ Pond · Phuwin · Permpoon  
รองรับผู้ใช้จำนวนมากโดยไม่โดน Google Apps Script quota

---

## Architecture

```
Google Sheet (ฐานข้อมูล)
    ↓
Apps Script Trigger (ทุก 5 นาที)
    ↓
Static JSON (data/)
    ↓
เว็บไซต์อ่านจาก Cloudflare CDN <- ผู้ใช้ทุกคน
```

ผู้ใช้ไม่ได้ยิง Apps Script โดยตรง — อ่านจาก Static JSON แทน  
ทำให้รองรับผู้ใช้ได้ไม่จำกัด โดยไม่โดน quota

---

## ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` | เว็บไซต์หลัก (deploy บน Cloudflare Workers) |
| `_headers` | HTTP headers สำหรับ Cloudflare |
| `data/events.json` | ข้อมูล events (static) |
| `data/works.json` | ข้อมูล works (static) |
| `data/birthdays.json` | ข้อมูล birthdays (static) |
| `data/anniversaries.json` | ข้อมูล anniversaries (static) |
| `README.md` | เอกสารนี้ |

---

## การตั้งค่า

### Google Sheet
- **Sheet ID:** `1hf42HWcjwKuLx25O5CNbJm6X2-XOHwmHkaEKINRt5Uo`
- **Tab ที่ใช้:** Events, Works, Birthdays, Anniversaries

### Apps Script
- **URL:** `https://script.google.com/macros/s/AKfycbwOqV8FAzHRl4jlNTebwOqzm5Si34J8IEW8UijV2grusYeF7TcilC9Lp5E40d-1cwf5/exec`
- **Trigger:** ทุก 5 นาที → `regenerateAllCaches()`

### เว็บไซต์
- **Host:** Cloudflare Workers
- **URL:** `https://schedule.trendforpondphuwin.workers.dev`
- **Repo:** `ppw1257/schedule` (GitHub)

---

## วิธี Deploy

### ครั้งแรก (ตั้งค่า)

1. เปิด Google Sheet → Extensions → Apps Script
2. วาง `Code.gs` ทับโค้ดเดิม
3. กด Run → `setupTriggers` (ทำครั้งเดียว)
4. Deploy → New Deployment → Web App
   - Execute as: Me
   - Who has access: Anyone

### อัปเดตโค้ด

**Code.gs (Apps Script):**
1. เปิด Apps Script → วางโค้ดใหม่ทับ
2. Deploy → Manage Deployments → Edit → New Version → Deploy

**index.html (เว็บ):**
1. Push ไฟล์ขึ้น GitHub repo `ppw1257/schedule`
2. Cloudflare Workers จะ deploy อัตโนมัติ

---

## Admin Panel

**URL:** `https://schedule.trendforpondphuwin.workers.dev/?panel=pppsecret`

เข้าสู่ระบบด้วย password → เพิ่ม/แก้ไข/ลบ Events, Works, Birthdays, Anniversaries

การทำงาน:
- Admin กด Save → บันทึกลง Google Sheet
- Cache refresh ทันที → ผู้ใช้เห็นข้อมูลใหม่ทันที (reload หน้า)

---

## Google Sheet Structure

### Events
| คอลัมน์ | ความหมาย |
|---|---|
| id | รหัสเฉพาะ (auto) |
| title | ชื่องาน |
| artists | ศิลปิน (Pond, Phuwin, Permpoon) |
| dateStart | วันเริ่ม (YYYY-MM-DD) |
| dateEnd | วันสิ้นสุด |
| startTime | เวลาเริ่ม |
| image | URL รูปภาพ |
| location | สถานที่ |
| city | เมือง |
| country | ประเทศ |
| type | ประเภท (Event, Concert, FM, etc.) |
| ticket1-3 | ลิงก์ตั๋ว |

### Works
| คอลัมน์ | ความหมาย |
|---|---|
| id | รหัสเฉพาะ |
| title | ชื่อผลงาน |
| artists | ศิลปิน |
| category | หมวด (Drama, Music, etc.) |
| year | ปี |
| image | URL รูป |

### Birthdays
| คอลัมน์ | ความหมาย |
|---|---|
| id | รหัสเฉพาะ |
| artist_name | ชื่อศิลปิน |
| birthday_month | เดือนเกิด |
| birthday_day | วันเกิด |
| color | สี hex |

### Anniversaries
| คอลัมน์ | ความหมาย |
|---|---|
| id | รหัสเฉพาะ |
| anniversary_name | ชื่อวันครบรอบ |
| anniversary_month | เดือน |
| anniversary_day | วัน |
| anniversary_year | ปีเริ่มต้น (optional) |

---

## Troubleshooting

### งานไม่ขึ้นบนเว็บ
1. เปิด Apps Script → รัน `checkCacheStatus()` → ดู log
2. ถ้า cache ว่าง → รัน `regenerateAllCaches()`
3. ถ้า error → ตรวจสอบ Sheet ID และ permission

### Admin บันทึกไม่ได้
1. ตรวจสอบ Apps Script URL ใน `index.html`
2. ตรวจสอบว่า Deploy เป็น Anyone ไม่ใช่ Anyone with Google account
3. ลอง deploy ใหม่

### เว็บเปิดไม่ได้ / SSL error
1. เปิด Cloudflare Dashboard
2. Workers & Pages → schedule
3. ตรวจสอบ deployment ล่าสุด

---

## Cache System

| เหตุการณ์ | Cache refresh |
|---|---|
| Admin บันทึกข้อมูล | ทันที |
| Trigger อัตโนมัติ | ทุก 5 นาที |
| ผู้ใช้เปิดเว็บ | อ่าน Static JSON (ไม่ยิง Sheet) |
| Cache หมดอายุ | โหลดใหม่ background |
| ออฟไลน์ | ใช้ Cache เก่า (7 วัน) |

---

## Responsive Breakpoints

| หน้าจอ | คอลัมน์ |
|---|---|
| < 400px (มือถือเล็ก) | 1 |
| 401-600px (มือถือ) | 2 |
| 601-800px (มือถือใหญ่) | 3 |
| 801-1100px (แท็บเล็ต) | 4 |
| > 1100px (เดสก์ท็อป) | 5 |

---

*Last updated: March 2026*
