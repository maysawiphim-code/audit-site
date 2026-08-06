# ตรวจไฟล์รายงานจับตัวเลข (Site Count Report Checker)

เว็บแอปตรวจความถูกต้องของไฟล์รายงานจับตัวเลข (.xls / .xlsx) ก่อนส่งงาน
เน้นความถูกต้องของ **ยอดรวมคนและรถ** ซึ่งแต่ละไฟล์มีจำนวนจุดและจำนวนวันไม่เท่ากัน

## กติกาการตรวจ

- ตรวจเฉพาะ **ชีตที่มองเห็น** — ชีตที่ซ่อน (RPT05\_\*) และชีต **สรุป (2)** จะถูกข้าม
- ตรวจเฉพาะ **ช่องที่ใส่สีไว้ในไฟล์** — ช่องพื้นขาวหรือไม่มีสีจะไม่ถูกตรวจ
- จำนวนนับ **คลาดเคลื่อนได้ ±0.5**

## สิ่งที่ตรวจ

| เกณฑ์ | รายละเอียด |
|---|---|
| ยอดรวมคน/รถทั้งหมด | แถวยอดรวมย่อยที่ใช้สีเดียวกันบวกกันต้องเท่าแถวยอดรวมใหญ่ที่ใช้อีกสี (ไม่ผูกกับจำนวนจุด) |
| ผลรวมกลุ่มย่อย | เพศ / อายุ / อาชีพ / ทิศ / ประเภทรถ รวมกันต้องเท่ายอดรวมของบล็อก |
| ผลัด | ผลัดเช้า + ผลัดบ่าย + ผลัดดึก = รวมทั้งวัน |
| ค่าเฉลี่ย | คอลัมน์เฉลี่ยต้องเท่าค่าเฉลี่ยของทุกวันที่จับ |
| %สัดส่วน | ต้องตรงกับค่า ÷ ฐานของกลุ่มตัวเอง |
| ตารางรายชั่วโมง | รวมทุกช่วงเวลา = แถวรวม และยอดรวมต้องพบในชีต data |
| ชีตสรุป | ตัวเลขทุกค่าต้องมีที่มาจากชีต data |
| หัวเรื่องทุกชีต | ชื่อทำเล วันที่จับตัวเลข และปี พ.ศ. ต้องตรงกับชีต data |
| เซลล์สูตร | จับ #REF! #DIV/0! #VALUE! |
| ชีตแผนที่/กราฟ | อ่านตัวเลขในรูปด้วย OCR แล้วเทียบกับยอดรวม (เป็นตัวช่วย ต้องยืนยันด้วยตา) |

## รันในเครื่อง

```bash
git clone https://github.com/<user>/<repo>.git
cd <repo>
pip install -r requirements.txt
streamlit run app.py
```

ถ้าต้องการให้ฟีเจอร์อ่านตัวเลขในรูปทำงาน ให้ติดตั้งเพิ่ม (Ubuntu/Debian):

```bash
sudo apt install tesseract-ocr tesseract-ocr-tha libreoffice-calc
```

- `tesseract-ocr` ใช้อ่านตัวเลขในรูป
- `libreoffice-calc` ใช้แปลง `.xls` เป็น `.xlsx` เพื่อดึงรูปออกมา (ไฟล์ `.xlsx` ไม่ต้องใช้)

หากไม่ติดตั้งสองตัวนี้ ส่วนอื่นของแอปยังทำงานได้ตามปกติ

## ขึ้น GitHub

```bash
git init
git add .
git commit -m "ตรวจรายงานจับตัวเลข"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

## ขึ้น Streamlit Community Cloud

1. ไปที่ https://share.streamlit.io แล้วกด **New app**
2. เลือก repo · branch `main` · main file `app.py`
3. กด Deploy — `requirements.txt` และ `packages.txt` (tesseract, libreoffice) จะถูกติดตั้งอัตโนมัติ

> `packages.txt` ทำให้ deploy ครั้งแรกช้าขึ้นหลายนาทีเพราะติดตั้ง LibreOffice
> ถ้าไม่ต้องการฟีเจอร์อ่านรูป ให้ลบไฟล์ `packages.txt` ทิ้งได้

## โครงไฟล์

```
app.py            หน้าเว็บ Streamlit
checker.py        แกนตรวจสอบ (ใช้แยกเป็นไลบรารีหรือรันใน CI ได้)
requirements.txt  แพ็กเกจ Python
packages.txt      แพ็กเกจระบบสำหรับ Streamlit Cloud
```

ใช้ `checker.py` เดี่ยว ๆ ได้:

```python
import checker
rep = checker.audit(open("report.xls", "rb").read(), "report.xls")
print(len(rep.bads), "จุดที่ต้องแก้")
for i in rep.bads:
    print(i.sheet, i.cell, i.title, i.detail)
```
