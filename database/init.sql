CREATE DATABASE IF NOT EXISTS cj_wms_dashboard
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

CREATE USER IF NOT EXISTS 'cj_wms_user'@'localhost'
  IDENTIFIED BY 'change_me';

CREATE USER IF NOT EXISTS 'cj_wms_user'@'%'
  IDENTIFIED BY 'change_me';

GRANT ALL PRIVILEGES ON cj_wms_dashboard.* TO 'cj_wms_user'@'localhost';
GRANT ALL PRIVILEGES ON cj_wms_dashboard.* TO 'cj_wms_user'@'%';

FLUSH PRIVILEGES;

USE cj_wms_dashboard;

CREATE TABLE IF NOT EXISTS app_settings (
  setting_key VARCHAR(80) NOT NULL,
  setting_value JSON NOT NULL,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (setting_key)
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS excel_uploads (
  id VARCHAR(36) NOT NULL,
  category VARCHAR(20) NOT NULL DEFAULT 'inbound',
  original_filename VARCHAR(255) NOT NULL,
  stored_filename VARCHAR(255) NOT NULL,
  content_type VARCHAR(120) NOT NULL,
  file_size BIGINT NOT NULL,
  file_data LONGBLOB NOT NULL,
  parsed_data JSON NOT NULL,
  uploaded_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_excel_uploads_stored_filename (stored_filename),
  KEY idx_excel_uploads_category (category),
  KEY idx_excel_uploads_uploaded_at (uploaded_at)
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS users (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  username VARCHAR(80) NOT NULL,
  password_hash VARCHAR(255) NOT NULL,
  role VARCHAR(40) NOT NULL DEFAULT 'admin',
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_users_username (username)
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS upload_batches (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  source_file_name VARCHAR(255) NOT NULL,
  stored_file_path VARCHAR(500) NOT NULL,
  uploaded_by BIGINT UNSIGNED NULL,
  uploaded_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  status VARCHAR(40) NOT NULL DEFAULT 'uploaded',
  error_message TEXT NULL,
  PRIMARY KEY (id),
  KEY idx_upload_batches_uploaded_at (uploaded_at),
  KEY idx_upload_batches_status (status),
  CONSTRAINT fk_upload_batches_uploaded_by
    FOREIGN KEY (uploaded_by) REFERENCES users(id)
    ON DELETE SET NULL
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS warehouse_tasks (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  batch_id BIGINT UNSIGNED NULL,
  work_group ENUM('inbound', 'pick', 'outbound') NOT NULL,
  reference_no VARCHAR(120) NULL,
  customer VARCHAR(255) NULL,
  destination VARCHAR(255) NULL,
  status ENUM('pending', 'in_progress', 'completed', 'cancelled') NOT NULL DEFAULT 'pending',
  plan_qty DECIMAL(18, 2) NOT NULL DEFAULT 0,
  completed_qty DECIMAL(18, 2) NOT NULL DEFAULT 0,
  pending_qty DECIMAL(18, 2) NOT NULL DEFAULT 0,
  work_date DATE NULL,
  raw_data JSON NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_warehouse_tasks_batch_id (batch_id),
  KEY idx_warehouse_tasks_work_group (work_group),
  KEY idx_warehouse_tasks_status (status),
  KEY idx_warehouse_tasks_work_date (work_date),
  CONSTRAINT fk_warehouse_tasks_batch_id
    FOREIGN KEY (batch_id) REFERENCES upload_batches(id)
    ON DELETE SET NULL
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COLLATE=utf8mb4_unicode_ci;

INSERT INTO app_settings (setting_key, setting_value)
VALUES
(
  'dashboard_settings',
  JSON_OBJECT(
    'refresh_seconds', 60,
    'theme', 'light',
    'show_inbound', true,
    'show_pick', true,
    'show_outbound', true
  )
)
ON DUPLICATE KEY UPDATE setting_value = setting_value;

INSERT INTO app_settings (setting_key, setting_value)
VALUES
(
  'workspace_layout',
  JSON_OBJECT(
    'boxes',
    JSON_ARRAY(
      JSON_OBJECT(
        'id', 'box-1',
        'size', '1x1',
        'title', 'Box 1x1',
        'cell', 0,
        'columns', 1,
        'rows', 1,
        'widgets', JSON_ARRAY(
          JSON_OBJECT(
            'id', 'widget-1',
            'type', 'title',
            'label', 'Title',
            'slot', 0
          )
        )
      )
    )
  )
)
ON DUPLICATE KEY UPDATE setting_value = setting_value;

INSERT INTO app_settings (setting_key, setting_value)
VALUES
(
  'dashboard_widgets',
  JSON_OBJECT(
    'widgets',
    JSON_ARRAY()
  )
)
ON DUPLICATE KEY UPDATE setting_value = setting_value;
