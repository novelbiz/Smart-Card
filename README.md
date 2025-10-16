# โปรแกรมอ่านบัตรประชาชนไทย (Thai National ID Card Reader)

โปรแกรม Command-line สำหรับอ่านข้อมูลจากบัตรประชาชนไทยผ่านเครื่องอ่านบัตรอัจฉริยะ (Smart Card Reader)

**พัฒนาโดย:** Novelbiz (https://novelbiz.co.th)  
**License:** MIT License 2025

---

## 📋 คุณสมบัติ (Features)

โปรแกรมสามารถอ่านข้อมูลจากบัตรประชาชนไทยได้ดังนี้:
- เลขบัตรประชาชน 13 หนัก
- ชื่อ-นามสกุล (ภาษาไทยและอังกฤษ)
- วันเกิด (แสดงทั้งรูปแบบไทยและอังกฤษ)
- เพศ
- วันที่ออกบัตร
- วันที่หมดอายุ
- สถานที่ออกบัตร
- ที่อยู่ตามทะเบียนบ้าน
- หมายเลขคำขอ

---

## 💻 ความต้องการของระบบ (System Requirements)

### ฮาร์ดแวร์
- เครื่องอ่านบัตรอัจฉริยะ (Smart Card Reader) ที่รองรับมาตรฐาน ISO 7816
- พอร์ต USB สำหรับเชื่อมต่อเครื่องอ่านบัตร

### ซอฟต์แวร์
- **ระบบปฏิบัติการ:** Windows 10/11 (แนะนำ) หรือ Linux/macOS
- **Python:** เวอร์ชัน 3.7 ขึ้นไป
- **Smart Card Service:** ต้องเปิดใช้งาน Smart Card Service บน Windows

---

## 🔧 การติดตั้ง (Installation)

### ขั้นตอนที่ 1: ติดตั้ง Python

ดาวน์โหลดและติดตั้ง Python จาก [python.org](https://www.python.org/downloads/)

ตรวจสอบการติดตั้ง:
```bash
python --version
```

### ขั้นตอนที่ 2: ติดตั้ง Driver สำหรับ Smart Card Reader

ติดตั้ง driver ของเครื่องอ่านบัตรตามคู่มือของผู้ผลิต

### ขั้นตอนที่ 3: เปิดใช้งาน Smart Card Service (สำหรับ Windows)

1. กดปุ่ม `Win + R` พิมพ์ `services.msc` แล้วกด Enter
2. หาบริการชื่อ **Smart Card**
3. คลิกขวาเลือก **Properties**
4. ตั้งค่า Startup type เป็น **Automatic**
5. คลิก **Start** เพื่อเริ่มบริการ
6. คลิก **OK**

### ขั้นตอนที่ 4: ติดตั้ง Library ที่จำเป็น

#### สำหรับ Windows:
```bash
pip install pyscard
```

#### สำหรับ Linux (Ubuntu/Debian):
```bash
# ติดตั้ง dependencies
sudo apt-get update
sudo apt-get install pcscd pcsc-tools libpcsclite-dev swig

# ติดตั้ง Python library
pip install pyscard
```

#### สำหรับ macOS:
```bash
# ติดตั้ง dependencies ผ่าน Homebrew
brew install swig

# ติดตั้ง Python library
pip install pyscard
```

### ขั้นตอนที่ 5: ดาวน์โหลดโปรแกรม

```bash
# ดาวน์โหลดไฟล์ หรือ clone repository
git clone [repository-url]
cd Thai_id_card
```

หรือคัดลอกไฟล์ `Thai_id_card_cli.py` ไปยังโฟลเดอร์ที่ต้องการ

---

## 🚀 วิธีใช้งาน (Usage)

### ขั้นตอนการใช้งาน

1. เชื่อมต่อเครื่องอ่านบัตรเข้ากับคอมพิวเตอร์
2. เปิด Command Prompt หรือ Terminal
3. ไปยังโฟลเดอร์ที่มีไฟล์โปรแกรม
4. รันคำสั่ง:

```bash
python Thai_id_card_cli.py
```

5. ใส่บัตรประชาชนลงในเครื่องอ่านบัตร
6. รอให้โปรแกรมอ่านข้อมูล
7. ข้อมูลจะแสดงบนหน้าจอ

### ตัวอย่างผลลัพธ์

```
======================================================================
               โปรแกรมอ่านบัตรประชาชนไทย - Novelbiz
                         MIT License 2025
======================================================================
[เริ่มต้น] โปรแกรมพร้อมทำงาน
======================================================================

[สถานะ] พบเครื่องอ่านบัตร: ACS ACR122U PICC Interface 0

[พร้อม] ใส่บัตรประชาชนเพื่ออ่านข้อมูล...

======================================================================
[เริ่มอ่าน] กำลังอ่านข้อมูลบัตรประชาชน...
======================================================================

[เชื่อมต่อ] เชื่อมต่อกับ: ACS ACR122U PICC Interface 0
[ATR] 3B 67 00 00 ...
[สำเร็จ] เลือก Thai ID Applet สำเร็จ (SW: 90 00)

----------------------------------------------------------------------
                         ข้อมูลบัตรประชาชน
----------------------------------------------------------------------

📌 เลขบัตรประชาชน: 1234567890123
👤 ชื่อ-นามสกุล (ไทย): นาย สมชาย ใจดี
👤 ชื่อ-นามสกุล (อังกฤษ): Mr. Somchai Jaidee
🎂 วันเกิด (ไทย): 15 มกราคม 2533
🎂 วันเกิด (อังกฤษ): 15 January 1990
⚧ เพศ: ชาย

📅 วันที่ออกบัตร (ไทย): 1 มกราคม 2563
📅 วันที่ออกบัตร (อังกฤษ): 1 January 2020
📅 วันที่หมดอายุ (ไทย): 31 ธันวาคม 2573
📅 วันที่หมดอายุ (อังกฤษ): 31 December 2030
🏢 สถานที่ออกบัตร: สำนักทะเบียนอำเภอเมือง จังหวัดกรุงเทพมหานคร

🏠 ที่อยู่: 123 ถนนสุขุมวิท แขวงคลองเตย เขตคลองเตย กรุงเทพมหานคร 10110
🔢 หมายเลขคำขอ: AA1234567890

======================================================================
[สำเร็จ] ✅ อ่านข้อมูลบัตรประชาชนเรียบร้อยแล้ว
======================================================================

[เสร็จสิ้น] อ่านข้อมูลเสร็จสมบูรณ์
```

---

## ❗ การแก้ปัญหา (Troubleshooting)

### ปัญหา: ไม่พบเครื่องอ่านบัตร

**แก้ไข:**
- ตรวจสอบว่าเครื่องอ่านบัตรเชื่อมต่อกับคอมพิวเตอร์แล้ว
- ติดตั้ง driver ของเครื่องอ่านบัตร
- ตรวจสอบว่า Smart Card Service ทำงานอยู่ (Windows)

### ปัญหา: ไม่สามารถติดตั้ง pyscard ได้

**แก้ไข:**
```bash
# ลองอัปเดต pip ก่อน
python -m pip install --upgrade pip

# สำหรับ Windows อาจต้องติดตั้ง Visual C++ Build Tools
# ดาวน์โหลดจาก: https://visualstudio.microsoft.com/visual-cpp-build-tools/

# ลองติดตั้งอีกครั้ง
pip install pyscard
```

### ปัญหา: Smart Card Service ไม่ทำงาน (Windows)

**แก้ไข:**
```bash
# เปิด Command Prompt แบบ Administrator แล้วรันคำสั่ง:
sc start SCardSvr
```

### ปัญหา: ไม่สามารถอ่านข้อมูลจากบัตรได้

**แก้ไข:**
- ทำความสะอาดบัตรและเครื่องอ่านบัตร
- ตรวจสอบว่าบัตรใส่ถูกทิศทาง
- ลองถอดบัตรแล้วใส่ใหม่
- รีสตาร์ทโปรแกรม

### ปัญหา: ข้อความแสดงผลผิดเพี้ยน

**แก้ไข:**
- ตรวจสอบว่า Terminal รองรับ UTF-8 และ TIS-620
- สำหรับ Windows ให้ตั้งค่า Code Page:
```bash
chcp 65001
```

---

## 🏗️ โครงสร้างของโปรแกรม (Code Structure)

### คลาสหลัก: `IDCardReader`

โปรแกรมประกอบด้วยคลาสหลัก 1 คลาส ที่มี methods สำคัญดังนี้:

#### 1. **Helper Functions** - ฟังก์ชันช่วยเหลือ

```python
decode_text(data)
```
- แปลงข้อมูล bytes จากบัตรเป็นข้อความที่อ่านได้
- รองรับ encoding แบบ TIS-620 (มาตรฐานภาษาไทย)

```python
send_apdu_with_get_response(connection, apdu)
```
- ส่งคำสั่ง APDU (Application Protocol Data Unit) ไปยังบัตร
- จัดการ GET_RESPONSE อัตโนมัติเมื่อบัตรตอบกลับด้วย SW1=0x61

```python
parse_thai_date(date_str)
```
- แปลงวันที่จากรูปแบบ YYYYMMDD เป็นรูปแบบที่อ่านง่าย
- คำนวณปีพุทธศักราช (Thai) และคริสต์ศักราช (English)
- รองรับกรณีพิเศษ "99999999" = ตลอดชีพ

```python
disconnect_card()
```
- ตัดการเชื่อมต่อกับบัตรอย่างปลอดภัย

#### 2. **Card Reader Functions** - ฟังก์ชันเครื่องอ่านบัตร

```python
check_service_status()
```
- ตรวจสอบว่า Smart Card Service ทำงานอยู่หรือไม่ (Windows)
- ใช้คำสั่ง `sc query SCardSvr`

```python
check_reader_status()
```
- ตรวจสอบว่ามีเครื่องอ่านบัตรเชื่อมต่ออยู่หรือไม่
- แสดงชื่อเครื่องอ่านบัตรที่พบ

#### 3. **Main Function** - ฟังก์ชันหลัก

```python
read_id_card()
```
- ฟังก์ชันหลักในการอ่านข้อมูลจากบัตรประชาชน
- ทำงานตามลำดับ:
  1. เชื่อมต่อกับบัตร
  2. เลือก Thai ID Card Applet
  3. อ่านข้อมูลแต่ละฟิลด์
  4. แสดงผลและตัดการเชื่อมต่อ

```python
run()
```
- ฟังก์ชันรันโปรแกรมหลัก
- ตรวจสอบระบบและรอให้ผู้ใช้ใส่บัตร

---

## 🔐 APDU Commands ที่ใช้

APDU (Application Protocol Data Unit) คือคำสั่งมาตรฐานสำหรับสื่อสารกับ Smart Card

### คำสั่งเลือก Applet

```python
SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
THAI_ID_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
```
- `0x00, 0xA4` = SELECT command
- `0x04, 0x00` = เลือกโดยใช้ AID (Application Identifier)
- `0x08` = ความยาวของ AID
- AID ของบัตรประชาชนไทย = `A0 00 00 00 54 48 00 01`

### คำสั่งอ่านข้อมูล

รูปแบบคำสั่ง READ BINARY:
```python
[0x80, 0xB0, offset_high, offset_low, 0x02, 0x00, length]
```

**ตัวอย่างคำสั่งอ่านข้อมูลแต่ละฟิลด์:**

| ข้อมูล | คำสั่ง APDU | Offset | Length |
|--------|-------------|--------|---------|
| เลขบัตรประชาชน | `80 B0 00 04 02 00 0D` | 0x0004 | 13 bytes |
| ชื่อ-สกุล (ไทย) | `80 B0 00 11 02 00 64` | 0x0011 | 100 bytes |
| ชื่อ-สกุล (อังกฤษ) | `80 B0 00 75 02 00 64` | 0x0075 | 100 bytes |
| วันเกิด | `80 B0 00 D9 02 00 08` | 0x00D9 | 8 bytes |
| เพศ | `80 B0 00 E1 02 00 01` | 0x00E1 | 1 byte |
| วันออกบัตร | `80 B0 01 67 02 00 08` | 0x0167 | 8 bytes |
| วันหมดอายุ | `80 B0 01 6F 02 00 08` | 0x016F | 8 bytes |
| สถานที่ออกบัตร | `80 B0 00 F6 02 00 64` | 0x00F6 | 100 bytes |
| ที่อยู่ | `80 B0 15 79 02 00 64` | 0x1579 | 100 bytes |
| หมายเลขคำขอ | `80 B0 16 19 02 00 0E` | 0x1619 | 14 bytes |

---

## 📊 การทำงานของโปรแกรม (Flow)

### ขั้นตอนการทำงานหลัก

```
1. เริ่มโปรแกรม
   ↓
2. ตรวจสอบ Smart Card Service (Windows)
   ↓
3. ตรวจสอบเครื่องอ่านบัตร
   ↓
4. รอให้ผู้ใช้ใส่บัตร (Timeout 5 วินาที)
   ↓
5. เชื่อมต่อกับบัตร (T0 หรือ T1 protocol)
   ↓
6. อ่าน ATR (Answer To Reset)
   ↓
7. เลือก Thai ID Card Applet
   ├─ ส่งคำสั่ง SELECT
   └─ ตรวจสอบ SW (Status Word) = 90 00
   ↓
8. อ่านข้อมูลแต่ละฟิลด์
   ├─ ส่งคำสั่ง READ BINARY
   ├─ รับ Response และ SW
   ├─ จัดการ GET_RESPONSE (ถ้า SW1=61)
   └─ Decode ข้อมูล
   ↓
9. แสดงผลข้อมูล
   ↓
10. ตัดการเชื่อมต่อ
```

### การจัดการ Status Word (SW)

โปรแกรมตรวจสอบ Status Word ที่ได้รับจากบัตร:

- **90 00** = Success (สำเร็จ)
- **61 XX** = มีข้อมูลเพิ่มเติม XX bytes (ต้องส่ง GET_RESPONSE)
- **อื่นๆ** = Error (ผิดพลาด)

```python
response, sw1, sw2 = connection.transmit(apdu)

if sw1 == 0x61:  # มีข้อมูลเพิ่มเติม
    get_response = [0x00, 0xC0, 0x00, 0x00, sw2]
    response, sw1, sw2 = connection.transmit(get_response)

if sw1 == 0x90 and sw2 == 0x00:  # สำเร็จ
    # ประมวลผลข้อมูล
```

---

## 🎯 จุดเด่นของโค้ด

### 1. **Error Handling ที่ครอบคลุม**
- จัดการ Exception ในทุกขั้นตอน
- Timeout สำหรับการรอบัตร
- ตัดการเชื่อมต่ออัตโนมัติเมื่อเกิดข้อผิดพลาด

### 2. **การแสดงผลที่เป็นมิตร**
- แสดง emoji เพื่อความชัดเจน
- แบ่งหมวดหมู่ข้อมูลด้วยเส้นกั้น
- แสดงข้อมูลทั้งภาษาไทยและอังกฤษ

### 3. **รองรับ Encoding หลายแบบ**
- TIS-620 สำหรับภาษาไทย
- ASCII fallback เมื่อ decode ไม่สำเร็จ

### 4. **ตรวจสอบระบบอัตโนมัติ**
- ตรวจสอบ Smart Card Service (Windows)
- ตรวจสอบเครื่องอ่านบัตรก่อนเริ่มทำงาน

---

## 📝 หมายเหตุ

- โปรแกรมนี้ออกแบบมาสำหรับบัตรประชาชนไทยเท่านั้น
- ข้อมูลที่อ่านได้จะแสดงบนหน้าจอเท่านั้น ไม่มีการบันทึกลงไฟล์
- การอ่านข้อมูลต้องปฏิบัติตามกฎหมายและระเบียบที่เกี่ยวข้อง
- ควรใช้โปรแกรมนี้เพื่อวัตถุประสงค์ที่ถูกต้องตามกฎหมายเท่านั้น
- โปรแกรมใช้ Protocol T0 หรือ T1 ตามที่บัตรรองรับ
- Offset และ Length ของแต่ละฟิลด์เป็นไปตามมาตรฐานบัตรประชาชนไทย

---

## 📄 License

โปรแกรมนี้เผยแพร่ภายใต้ MIT License

```
Copyright 2025 NOVELBIZ CO., LTD.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

---

## 👨‍💻 ผู้พัฒนา

**Novelbiz Co., Ltd.**
- Website: https://novelbiz.co.th
- Email: support@novelbiz.co.th

---

## 🤝 การสนับสนุน

หากพบปัญหาหรือต้องการสอบถามข้อมูลเพิ่มเติม กรุณาติดต่อผ่าน:
- Website: https://novelbiz.co.th
- หรือสร้าง Issue ใน GitHub Repository

---

## 📚 เอกสารอ้างอิง

- [pyscard Documentation](https://pyscard.sourceforge.io/)
- [PC/SC Workgroup](https://www.pcscworkgroup.com/)
- [ISO/IEC 7816 Standard](https://www.iso.org/standard/54089.html)

---

**เวอร์ชัน:** 1.0.0  
**อัปเดตล่าสุด:** 2025
