# Thai Household Income Dashboard
## แดชบอร์ดวิเคราะห์ข้อมูลรายได้ครัวเรือนไทย

[![Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

แดชบอร์ดวิเคราะห์ข้อมูลรายได้ครัวเรือนไทยแยกตามจังหวัดและกลุ่มเศรษฐกิจและสังคม สร้างด้วย Python, Streamlit และ Plotly

---

## 📋 รายละเอียดชุดข้อมูล

| รายการ | ข้อมูล |
|--------|--------|
| ชื่อไฟล์ | `data.xlsx` (Sheet: `data`) |
| แหล่งข้อมูล | สำนักงานสถิติแห่งชาติ (NSO) |
| จำนวนแถว | 15,400 แถว |
| จำนวนคอลัมน์ | 11 คอลัมน์ |
| ปีที่ครอบคลุม | พ.ศ. 2566, 2568 |
| จำนวนจังหวัด | 77 จังหวัด |
| หน่วยข้อมูล | บาท/เดือน/ครัวเรือน |

### คอลัมน์ในชุดข้อมูล

| คอลัมน์ | คำอธิบาย |
|---------|----------|
| `year` | ปีพุทธศักราช |
| `province` | ชื่อจังหวัด |
| `source_income1` | หมวดรายได้ระดับที่ 1 (รายได้ประจำ / รายได้ไม่ประจำ / รายได้ทั้งสิ้น) |
| `source_income2` | หมวดรายได้ระดับที่ 2 |
| `source_income3` | หมวดรายได้ระดับที่ 3 (ละเอียดที่สุด) |
| `soc_eco_class1` | กลุ่มเศรษฐกิจ-สังคมระดับหลัก |
| `soc_eco_class2` | กลุ่มเศรษฐกิจ-สังคมระดับย่อย |
| `value` | มูลค่ารายได้ (บาท) |
| `unit` | หน่วย (บาท) |
| `attribute` | หมายเหตุเกี่ยวกับค่าพิเศษ (เช่น ค่าเป็น 0 หรือปัดเศษ) |
| `source` | แหล่งข้อมูล |

---

## 🛠️ เทคโนโลยีที่ใช้

- **Python** 3.9+
- **Streamlit** ≥ 1.28 — Web Application Framework
- **Pandas** ≥ 2.0 — Data Processing
- **Plotly** ≥ 5.18 — Interactive Charts
- **openpyxl** ≥ 3.1 — Excel File Reader

---

## 📁 โครงสร้างไฟล์

```
ThaiIncomeDashboard/
├── app.py              # แอปหลัก Streamlit
├── data.xlsx           # ชุดข้อมูลหลัก (จาก NSO)
├── requirements.txt    # Dependencies
├── README.md           # ไฟล์นี้
└── .gitignore          # ไฟล์ที่ไม่รวมใน Git
```

---

## 🚀 วิธีติดตั้งและรันในเครื่อง

### 1. Clone Repository

```bash
git clone https://github.com/YOUR_USERNAME/thai-household-income-dashboard.git
cd thai-household-income-dashboard
```

### 2. สร้าง Virtual Environment (แนะนำ)

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3. ติดตั้ง Dependencies

```bash
pip install -r requirements.txt
```

### 4. รันแอป

```bash
streamlit run app.py
```

แอปจะเปิดใน Browser อัตโนมัติที่ `http://localhost:8501`

---

## 🎛️ วิธีใช้งานตัวกรอง

| ตัวกรอง | ตัวเลือก | คำอธิบาย |
|---------|---------|----------|
| 📅 ปี (พ.ศ.) | 2566, 2568 | เลือกปีที่ต้องการวิเคราะห์ |
| 📍 จังหวัด | ทุกจังหวัด / จังหวัดเฉพาะ | กรองข้อมูลตามจังหวัด |
| 💰 หมวดรายได้ | ยอดรวม / แยกประเภท | เลือกระดับหมวดรายได้ |
| 👥 กลุ่มเศรษฐกิจ-สังคม | ทุกกลุ่ม / กลุ่มเฉพาะ | กรองตามกลุ่มอาชีพ |

---

## 📊 คำอธิบายกราฟและ KPI

### KPI Cards
1. **ปีที่วิเคราะห์** — ปีพุทธศักราชที่เลือก
2. **จังหวัดในผลลัพธ์** — ชื่อจังหวัดหรือจำนวนจังหวัดที่ผ่านตัวกรอง
3. **จำนวนรายการ** — จำนวนแถวข้อมูลที่ผ่านตัวกรองทั้งหมด
4. **จังหวัดค่าสูงสุด** — จังหวัดที่มีค่ารายได้ในหมวดที่เลือกสูงที่สุด

### กราฟ
1. **รายได้ตามจังหวัด** — แท่งเปรียบเทียบมูลค่าของหมวดรายได้ที่เลือกตามจังหวัด เรียงจากมากไปน้อย
2. **กลุ่มเศรษฐกิจ-สังคม** — แท่งแนวนอนเปรียบเทียบมูลค่าตามกลุ่มเศรษฐกิจ-สังคมระดับย่อย
3. **ประเภทรายได้ย่อย** — แท่งแนวนอนแสดงหมวดย่อยของรายได้ทั้งหมด (ไม่รวมยอดรวม)
4. **เปรียบเทียบระหว่างปี** — กราฟแท่งแบบ Grouped แสดงการเปลี่ยนแปลงระหว่างปี (เมื่อเลือกจังหวัดเฉพาะ)

---

## ⚠️ ข้อจำกัดในการตีความและรวมข้อมูลรายได้

> **สำคัญมาก:** ชุดข้อมูลนี้มีทั้งหมวดรายได้รวมและหมวดย่อยอยู่ในไฟล์เดียวกัน

- **ห้ามนำทุกแถวมารวมกัน** — จะเกิดการนับรายได้ซ้ำ
- หมวด `รายได้ทั้งสิ้นต่อเดือน` คือยอดรวมที่คำนวณไว้แล้วในชุดข้อมูล
- กราฟ "ประเภทรายได้ย่อย" แสดงเฉพาะหมวดย่อย (leaf rows) เพื่อป้องกันการนับซ้ำ
- ค่าติดลบในข้อมูลคือค่าจริงตามคำอธิบายของ NSO
- ค่า 0 ที่มี attribute `ศูนย์ยกกำลังดับเบิลยู` หมายถึงข้อมูลจากการสำรวจมีค่าเป็น 0 จริง

---

## 📤 วิธีอัปโหลดขึ้น GitHub

```bash
# ครั้งแรก (ยังไม่มี Repository)
git init
git add app.py requirements.txt README.md .gitignore data.xlsx
git commit -m "Add Thai Household Income Dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/thai-household-income-dashboard.git
git push -u origin main

# ครั้งต่อไป
git add -A
git commit -m "Update dashboard"
git push
```

---

## ☁️ วิธี Deploy บน Streamlit Community Cloud

1. ไปที่ [share.streamlit.io](https://share.streamlit.io)
2. ล็อกอินด้วยบัญชี GitHub
3. คลิก **"New app"**
4. เลือก Repository: `thai-household-income-dashboard`
5. เลือก Branch: `main`
6. ตั้ง Main file path: `app.py`
7. คลิก **"Deploy!"**
8. รอประมาณ 1–3 นาที จนกว่าแอปจะพร้อมใช้งาน

> **หมายเหตุ:** ตรวจสอบให้แน่ใจว่าไฟล์ `data.xlsx` อยู่ใน Repository ด้วย มิฉะนั้นแอปบน Cloud จะไม่สามารถโหลดข้อมูลได้

---

## 📞 ติดต่อ

โปรเจกต์นี้จัดทำสำหรับวิชา Data Visualization
