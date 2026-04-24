-- ============================================================
-- HR DATA QUALITY AUDIT — MySQL Schema & Queries
-- Project: Risk Control & Data Quality Audit System
-- Author:  [Your Name]
-- Tools:   MySQL, Python, Excel
-- ============================================================


-- ── SCHEMA SETUP ────────────────────────────────────────────

CREATE DATABASE IF NOT EXISTS hr_audit_db;
USE hr_audit_db;

DROP TABLE IF EXISTS hr_audit_results;
DROP TABLE IF EXISTS hr_employees;

CREATE TABLE hr_employees (
    employee_id       VARCHAR(10)    PRIMARY KEY,
    full_name         VARCHAR(100)   NOT NULL,
    age               INT,
    date_of_birth     DATE,
    gender            VARCHAR(10),
    email             VARCHAR(100),
    phone             VARCHAR(20),
    department        VARCHAR(50),
    job_role          VARCHAR(80),
    join_date         DATE,
    salary            DECIMAL(12,2),
    employment_status VARCHAR(20),
    manager_id        VARCHAR(10),
    pan_number        VARCHAR(20),
    contract_expiry   DATE
);

CREATE TABLE hr_audit_results (
    employee_id       VARCHAR(10)  PRIMARY KEY,
    risk_score        INT          DEFAULT 0,
    risk_level        VARCHAR(10),
    missing_email     TINYINT(1)   DEFAULT 0,
    invalid_email     TINYINT(1)   DEFAULT 0,
    missing_salary    TINYINT(1)   DEFAULT 0,
    negative_salary   TINYINT(1)   DEFAULT 0,
    missing_dept      TINYINT(1)   DEFAULT 0,
    underage_risk     TINYINT(1)   DEFAULT 0,
    missing_dob       TINYINT(1)   DEFAULT 0,
    missing_phone     TINYINT(1)   DEFAULT 0,
    missing_pan       TINYINT(1)   DEFAULT 0,
    duplicate_email   TINYINT(1)   DEFAULT 0,
    salary_outlier    TINYINT(1)   DEFAULT 0,
    total_issues      INT          DEFAULT 0,
    audit_date        TIMESTAMP    DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (employee_id) REFERENCES hr_employees(employee_id)
);


-- ── DATA IMPORT (run after loading CSV via Python/LOAD DATA) ─

-- Load CSV into MySQL:
-- LOAD DATA INFILE '/path/to/hr_raw_data.csv'
-- INTO TABLE hr_employees
-- FIELDS TERMINATED BY ',' ENCLOSED BY '"'
-- LINES TERMINATED BY '\n'
-- IGNORE 1 ROWS;


-- ── AUDIT QUERIES ────────────────────────────────────────────

-- 1. Total employees per department
SELECT
    COALESCE(NULLIF(department,''), '⚠ MISSING') AS department,
    COUNT(*) AS total_employees,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM hr_employees), 1) AS pct_share
FROM hr_employees
GROUP BY department
ORDER BY total_employees DESC;


-- 2. Missing critical fields summary
SELECT
    SUM(CASE WHEN email IS NULL OR email = '' THEN 1 ELSE 0 END)       AS missing_email,
    SUM(CASE WHEN phone IS NULL OR phone = '' THEN 1 ELSE 0 END)       AS missing_phone,
    SUM(CASE WHEN salary IS NULL THEN 1 ELSE 0 END)                    AS missing_salary,
    SUM(CASE WHEN department IS NULL OR department = '' THEN 1 ELSE 0 END) AS missing_dept,
    SUM(CASE WHEN date_of_birth IS NULL THEN 1 ELSE 0 END)             AS missing_dob,
    SUM(CASE WHEN pan_number IS NULL OR pan_number = '' THEN 1 ELSE 0 END) AS missing_pan
FROM hr_employees;


-- 3. Detect invalid email formats
SELECT employee_id, full_name, email, department
FROM hr_employees
WHERE email IS NOT NULL
  AND email != ''
  AND email NOT REGEXP '^[a-zA-Z0-9._%+\\-]+@[a-zA-Z0-9.\\-]+\\.[a-zA-Z]{2,}$';


-- 4. Detect negative or zero salaries (data integrity risk)
SELECT employee_id, full_name, department, salary,
       CASE WHEN salary < 0 THEN '❌ Negative' ELSE '⚠ Zero' END AS issue
FROM hr_employees
WHERE salary IS NOT NULL AND salary <= 0
ORDER BY salary ASC;


-- 5. Underage employees — COMPLIANCE RISK
SELECT employee_id, full_name, age, department, join_date,
       'UNDERAGE - LEGAL RISK' AS risk_flag
FROM hr_employees
WHERE age < 18
ORDER BY age ASC;


-- 6. Salary outlier detection (beyond 2 standard deviations)
SELECT employee_id, full_name, department, salary,
    ROUND((SELECT AVG(salary) FROM hr_employees WHERE salary > 0), 2) AS avg_salary,
    ROUND(ABS(salary - (SELECT AVG(salary) FROM hr_employees WHERE salary > 0)) /
          (SELECT STDDEV(salary) FROM hr_employees WHERE salary > 0), 2) AS z_score
FROM hr_employees
WHERE salary > 0
HAVING z_score > 2
ORDER BY z_score DESC;


-- 7. Duplicate email detection
SELECT email, COUNT(*) AS duplicate_count,
       GROUP_CONCAT(employee_id ORDER BY employee_id SEPARATOR ', ') AS affected_employees
FROM hr_employees
WHERE email IS NOT NULL AND email != ''
GROUP BY email
HAVING COUNT(*) > 1
ORDER BY duplicate_count DESC;


-- 8. Risk level distribution
SELECT
    risk_level,
    COUNT(*) AS employee_count,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM hr_audit_results), 1) AS percentage,
    ROUND(AVG(risk_score), 2) AS avg_risk_score,
    MAX(risk_score) AS max_risk_score
FROM hr_audit_results
GROUP BY risk_level
ORDER BY FIELD(risk_level, 'HIGH', 'MEDIUM', 'LOW', 'CLEAN');


-- 9. Department-level risk analysis
SELECT
    e.department,
    COUNT(*) AS total_employees,
    SUM(CASE WHEN a.risk_level = 'HIGH' THEN 1 ELSE 0 END) AS high_risk,
    SUM(CASE WHEN a.risk_level = 'MEDIUM' THEN 1 ELSE 0 END) AS medium_risk,
    ROUND(AVG(a.risk_score), 2) AS avg_risk_score,
    ROUND(AVG(CASE WHEN e.salary > 0 THEN e.salary END), 2) AS avg_salary
FROM hr_employees e
JOIN hr_audit_results a ON e.employee_id = a.employee_id
WHERE e.department != '' AND e.department IS NOT NULL
GROUP BY e.department
ORDER BY high_risk DESC, avg_risk_score DESC;


-- 10. Top 10 highest risk employees with issue details
SELECT
    e.employee_id, e.full_name, e.department, e.age, e.salary,
    a.risk_score, a.risk_level, a.total_issues,
    CONCAT_WS(', ',
        IF(a.underage_risk, 'UNDERAGE', NULL),
        IF(a.negative_salary, 'NEG SALARY', NULL),
        IF(a.missing_dept, 'NO DEPT', NULL),
        IF(a.invalid_email, 'BAD EMAIL', NULL),
        IF(a.missing_pan, 'NO PAN', NULL)
    ) AS issues_detail
FROM hr_employees e
JOIN hr_audit_results a ON e.employee_id = a.employee_id
WHERE a.risk_level = 'HIGH'
ORDER BY a.risk_score DESC
LIMIT 10;


-- 11. Month-wise hiring trend
SELECT
    YEAR(join_date) AS year,
    MONTH(join_date) AS month,
    COUNT(*) AS new_hires,
    ROUND(AVG(salary), 0) AS avg_salary
FROM hr_employees
WHERE join_date IS NOT NULL
GROUP BY YEAR(join_date), MONTH(join_date)
ORDER BY year, month;


-- 12. Inactive employees with expired contracts (action needed)
SELECT employee_id, full_name, department, employment_status,
       contract_expiry,
       DATEDIFF(CURDATE(), contract_expiry) AS days_since_expiry
FROM hr_employees
WHERE employment_status = 'Inactive'
  AND contract_expiry IS NOT NULL
  AND contract_expiry < CURDATE()
ORDER BY days_since_expiry DESC;
