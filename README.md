# 🚀 Rust Print API with PDF Resizing By Gemini
## !!AI-Generated. Human-Validated. Use with Discretion!!


โปรเจกต์นี้คือ RESTful API ที่พัฒนาด้วยภาษา Rust โดยใช้เฟรมเวิร์ค Actix-web เพื่อรับคำสั่งในการประมวลผลไฟล์ PDF (ปรับขนาดเป็น A6) และส่งไฟล์ที่ปรับขนาดแล้วไปยังเครื่องพิมพ์ที่กำหนด นอกจากนี้ยังสามารถติดตั้งเป็น Windows Service เพื่อให้ทำงานอยู่เบื้องหลังได้ตลอดเวลา

## ✨ คุณสมบัติหลัก

*   **RESTful API:** ใช้ Actix-web ในการสร้าง Endpoint สำหรับรับคำสั่งพิมพ์
*   **PDF Transformation:** ใช้ Crate `lopdf` ในการปรับขนาดหน้าและย่อ/ขยาย (Scale) เนื้อหาของไฟล์ PDF ให้เป็นขนาด **A6**
*   **Persistent Conversion:** เก็บไฟล์ A6 ที่แปลงแล้วไว้ในโฟลเดอร์ `printable_files/` โดยไม่ลบ
*   **Printing Integration:** ใช้ Crate `printers` ในการสั่งพิมพ์ไฟล์ผ่านระบบปฏิบัติการ (รองรับ Windows, Linux/CUPS, macOS)
*   **API Documentation:** มาพร้อมกับ Swagger UI (OpenAPI Specification) สำหรับการทดสอบและดูเอกสารผ่านเว็บเบราว์เซอร์
*   **Windows Service Support:** สามารถติดตั้งและรันเป็น Windows Service ได้ ทำให้แอปพลิเคชันทำงานอยู่เบื้องหลังได้อย่างต่อเนื่อง
*   **Health Check Endpoint:** มี `GET /` endpoint สำหรับตรวจสอบสถานะการทำงานของ Service โดยจะแสดงหน้า HTML อย่างง่าย
*   **Input Validation:** ตรวจสอบคำขอพิมพ์ให้ปลอดภัยขึ้น โดยบังคับให้ `filename` เป็นไฟล์ `.pdf` แบบ relative path เท่านั้น (ป้องกัน path traversal)
*   **Configurable Host/Port:** กำหนดค่า API ผ่าน Environment Variables ได้ (`RUST_PRINT_API_HOST`, `RUST_PRINT_API_PORT`)

---

## 🛠️ Prerequisites (สิ่งที่ต้องมี)

1.  **Rust Toolchain:** ติดตั้ง Rust เวอร์ชัน Stable ([rustup.rs](https://rustup.rs/))
2.  **เครื่องพิมพ์:** ต้องมีเครื่องพิมพ์ที่ติดตั้งและพร้อมใช้งานในระบบปฏิบัติการที่คุณรัน API
3.  **Dependence เฉพาะ OS (สำหรับ Linux/macOS):**
    *   ถ้าคุณใช้ **Linux** หรือ **macOS**, คุณอาจต้องติดตั้งไลบรารี **CUPS** (Common Unix Printing System) เพื่อให้ `printers` ทำงานได้ (ส่วนใหญ่มาพร้อมกับ OS อยู่แล้ว)

---

## ⚙️ Installation and Setup (การติดตั้งและการตั้งค่า)

### 1. Clone Repository

เริ่มต้นด้วยการคัดลอกโปรเจกต์นี้มายังเครื่องของคุณ:

```bash
git clone [YOUR_REPOSITORY_URL]
cd [YOUR_PROJECT_FOLDER]

```

### 2. Build the Project

คอมไพล์โปรเจกต์ในโหมด Release เพื่อประสิทธิภาพสูงสุด:

```bash
cargo build --release
```
ไฟล์ executable จะอยู่ที่ `target/release/rust-print-api.exe` (สำหรับ Windows)

---

## 🚀 การใช้งาน (Usage)

### 1. รันในโหมด Console (สำหรับการพัฒนาและทดสอบ)

คุณสามารถรันแอปพลิเคชันในโหมด Console เพื่อดู Log Output ได้โดยตรง:

```bash
cargo run -- --console
```
หรือรันไฟล์ executable โดยตรง:
```bash
.\target\release\rust-print-api.exe --console
```
เมื่อรันแล้ว API จะพร้อมใช้งานที่ `http://127.0.0.1:8080` และ Swagger UI ที่ `http://127.0.0.1:8080/swagger-ui/`

> หมายเหตุ: ค่าเริ่มต้นปัจจุบันของโปรเจกต์คือ `127.0.0.1:3000`  
> สามารถเปลี่ยนได้ด้วย:
> - `RUST_PRINT_API_HOST` (default: `127.0.0.1`)
> - `RUST_PRINT_API_PORT` (default: `3000`)

### 2. API Endpoints

*   **GET /**
    *   **Description:** ตรวจสอบสถานะการทำงานของ Service
    *   **Response:** หน้า HTML แสดงสถานะ "Service is running!"
*   **POST /api/print**
    *   **Description:** รับไฟล์ PDF และชื่อเครื่องพิมพ์ เพื่อปรับขนาดเป็น A6 และส่งไปยังเครื่องพิมพ์
    *   **Request Body (JSON):**
        ```json
        {
            "filename": "your_document.pdf",
            "printer_name": "Your_Printer_Name"
        }
        ```
    *   **Validation Rules:**
        * `filename` ต้องไม่ว่าง
        * `filename` ต้องเป็นไฟล์นามสกุล `.pdf`
        * `filename` ต้องเป็น relative filename (ห้าม absolute path และห้ามมี `..`)
        * `printer_name` ต้องไม่ว่าง
    *   **Response (JSON):**
        ```json
        {
            "status": "success",
            "message": "Resized to A6, saved as your_document_a6.pdf, and sent to printer Your_Printer_Name"
        }
        ```
        หรือข้อความ Error หากเกิดปัญหา

---

## 🖥️ การติดตั้งเป็น Windows Service

เพื่อให้แอปพลิเคชันทำงานอยู่เบื้องหลังได้ตลอดเวลาและเริ่มต้นอัตโนมัติเมื่อ Windows เริ่มทำงาน คุณสามารถติดตั้งเป็น Windows Service ได้

### 1. สร้าง Service

เปิด **Command Prompt (Admin)** หรือ **PowerShell (Admin)** และรันคำสั่ง:

```bash
sc create "Rust Print API" binPath= "E:\Rust-API-Autoprint\target\release\rust-print-api.exe"
```
*   **"Rust Print API"**: คือชื่อ Service ที่คุณต้องการตั้ง
*   **binPath**: คือที่อยู่เต็มของไฟล์ executable ที่คุณคอมไพล์ไว้ในขั้นตอน `cargo build --release` (โปรดปรับเปลี่ยนตามพาธจริงของโปรเจกต์คุณ)

### 2. เริ่มต้น Service

หลังจากสร้าง Service แล้ว คุณสามารถเริ่มต้นได้ด้วยคำสั่ง:

```bash
sc start "Rust Print API"
```

### 3. ตรวจสอบสถานะ Service

คุณสามารถตรวจสอบสถานะของ Service ได้จาก Services Manager ของ Windows หรือใช้คำสั่ง:

```bash
sc query "Rust Print API"
```

### 4. หยุด Service

หากต้องการหยุด Service:

```bash
sc stop "Rust Print API"
```

### 5. ลบ Service

หากต้องการลบ Service ออกจากระบบ:

```bash
sc delete "Rust Print API"
```
**ข้อควรระวัง:** การลบ Service จะต้องหยุด Service นั้นก่อน

---
