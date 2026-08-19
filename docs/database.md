# Database Design Draft

เอกสารนี้เป็นร่าง schema สำหรับ MySQL ของ CJ WMS Dashboard

## Tables

### users

- id
- username
- password_hash
- role
- is_active
- created_at
- updated_at

### upload_batches

- id
- source_file_name
- stored_file_path
- uploaded_by
- uploaded_at
- status
- error_message

### warehouse_tasks

- id
- batch_id
- work_group: inbound, pick, outbound
- reference_no
- customer
- destination
- status: pending, in_progress, completed, cancelled
- plan_qty
- completed_qty
- pending_qty
- work_date
- created_at
- updated_at

### dashboard_settings

- id
- refresh_seconds
- theme
- show_inbound
- show_pick
- show_outbound
- updated_by
- updated_at

### app_settings

ใช้เก็บ configuration แบบ JSON key-value ในช่วง scaffold ก่อนแยก schema รายตารางละเอียด

- setting_key: primary key เช่น `dashboard_settings`, `dashboard_widgets`, `workspace_layout`
- setting_value: JSON
- updated_at

ข้อมูลที่บันทึกตอนนี้:

- `dashboard_settings`: refresh interval, theme, section visibility
- `dashboard_widgets`: layout ของ dashboard widget เดิม
- `workspace_layout`: Box/Widget layout ของหน้า Workspace Edit Mode

## Indexes

- warehouse_tasks(batch_id)
- warehouse_tasks(work_group)
- warehouse_tasks(status)
- warehouse_tasks(work_date)
- upload_batches(uploaded_at)
