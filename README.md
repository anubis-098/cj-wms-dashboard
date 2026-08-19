# CJ WMS Dashboard

CJ WMS Dashboard คือเว็บแสดงข้อมูลสถานะงานในคลังสินค้า CJ Logistics สำหรับติดตามงานหลักในคลังสินค้า ได้แก่

1. Inbound
2. Pick
3. Outbound

ระบบถูกออกแบบให้ User สามารถอัปโหลดไฟล์ Excel `.xlsx` เพื่อให้ระบบอ่านข้อมูลและนำไปแสดงผลบน Dashboard ได้ โดยมี 2 รูปแบบการใช้งานหลัก

- TV Display Mode: ใช้แสดงผลบน Samsung Smart TV / Kiosk Mode
- PC Admin Mode: ใช้สำหรับอัปโหลดไฟล์, ตั้งค่า Dashboard, และ Custom Layout โดยต้อง Login ก่อนใช้งาน

---

## Tech Stack

### Frontend

- React + Vite
- ApexCharts
- TailwindCSS

### Backend

- FastAPI
- Python
- Pandas / OpenPyXL สำหรับอ่านไฟล์ Excel

### Database

- MySQL

### Container

- Docker Compose
- ยังไม่ใช้ในช่วงแรก แต่เตรียมโครงสร้างไว้สำหรับอนาคต

### Target Display

- Samsung Smart TV
- Kiosk Mode
- Browser Full Screen

---

## Main Features

### 1. Excel Upload

User สามารถอัปโหลดไฟล์ Excel `.xlsx` ผ่านหน้าเว็บ

ระบบจะอ่านข้อมูลจากไฟล์ Excel และแยกข้อมูลตามหมวดงาน เช่น

- Inbound
- Pick
- Outbound

ข้อมูลที่อ่านได้จะถูกบันทึกลง MySQL เพื่อใช้แสดงผลบน Dashboard

---

### 2. Dashboard Display

Dashboard แสดงภาพรวมสถานะงานในคลังสินค้า เช่น

- จำนวนงานทั้งหมด
- จำนวนงานที่เสร็จแล้ว
- จำนวนงานคงค้าง
- เปอร์เซ็นต์ความคืบหน้า
- สถานะตามช่วงเวลา
- กราฟเปรียบเทียบ Inbound / Pick / Outbound

---

### 3. TV Display Mode

หน้าจอสำหรับ Samsung Smart TV จะเน้น

- ตัวอักษรใหญ่
- สีชัดเจน
- อ่านง่ายจากระยะไกล
- Auto Refresh
- Full Screen
- ไม่ต้อง Login
- ใช้ URL เฉพาะสำหรับเปิดบน TV

ตัวอย่าง URL:

```txt
http://server-ip:8000/tv