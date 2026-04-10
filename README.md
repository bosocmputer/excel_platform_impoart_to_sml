# SML Platform Order Importer

นำเข้าคำสั่งซื้อจาก Shopee เข้าระบบ SML ผ่าน REST API

## ความต้องการของระบบ

- Python 3.9 ขึ้นไป
- Windows 10/11 หรือ macOS (แนะนำทดสอบบน Windows)

## การติดตั้ง

```bash
pip install -r requirements.txt
```

## วิธีใช้งาน

```bash
python main.py
```

## ขั้นตอนการใช้งาน

1. **ตั้งค่า** (Tab ตั้งค่า)
   - กรอก Server URL, GUID, Provider, configFileName, databaseName
   - กรอก doc_format_code, sale_code, vat_type, vat_rate
   - กรอก wh_code, shelf_code, unit_code (fallback กรณีสินค้าไม่มี default)
   - เลือก cust_code (พิมพ์เองหรือกด "โหลดจาก SML")
   - กด **บันทึกการตั้งค่า** (จะบันทึกลง `config.json` อัตโนมัติ)

2. **Import** (Tab Import)
   - กด **เลือกไฟล์** → เลือกไฟล์ Excel ที่ export จาก Shopee
   - ตรวจสอบ Preview Orders (จะแสดง orders ที่พร้อม import)
   - กด **Import** เพื่อเริ่มส่งเข้า SML

## ข้อกำหนดไฟล์ Excel Shopee

- ต้องเป็นไฟล์ที่ **export จาก Shopee Seller Center** → คำสั่งซื้อ → Export
- **สินค้าทุกชิ้นต้องมีรหัส SKU** ที่ตรงกับ item_code ใน SML
  - หากไม่มี SKU ระบบจะแจ้ง warning และข้าม order นั้น
  - วิธีกำหนด SKU: Shopee Seller Center → สินค้า → แก้ไขสินค้า → ตัวเลือก → รหัสอ้างอิง SKU

## สถานะที่ระบบกรองออก (ไม่ import)

- ที่ต้องจัดส่ง
- ยกเลิกแล้ว

## ระบบ Log

- ทุกครั้งที่ Import จะสร้างไฟล์ log ใหม่ชื่อ `import_log_YYYYMMDD_HHMMSS.csv` ในโฟลเดอร์เดียวกับไฟล์ Excel
- หาก import ค้างกลางทาง → โหลด log เก่าด้วยปุ่ม **โหลด Log (Retry)** แล้วกด **Retry Failed**

## SML API ที่ใช้

| API | Endpoint |
|---|---|
| Sale Invoice | `POST /SMLJavaRESTService/saleinvoice/` |
| Customer List | `GET /SMLJavaRESTService/v3/api/customer` |
| Product Info | `GET /SMLJavaRESTService/v3/api/product/{code}` |

อ้างอิง: https://docs.smlaccount.com/service/v2/sale_invoice.html

## Build เป็น .exe สำหรับ Windows

```bash
pip install pyinstaller
pyinstaller build.spec
# ได้ไฟล์: dist/SML_Shopee_Importer.exe
```

## Platform ที่รองรับ

| Platform | สถานะ |
|---|---|
| Shopee | ✅ พร้อมใช้งาน |
| Lazada | 🚧 อยู่ระหว่างพัฒนา |
| TikTok Shop | 🚧 อยู่ระหว่างพัฒนา |
