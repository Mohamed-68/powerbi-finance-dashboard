/* ============================================================
   Finance Reporting Layer (Portfolio SQL)
   Author: <Your Name>
   Purpose:
     - Provide a clean reporting view for P&L analytics
     - Example KPI pack (Actual vs Budget + Variance)
     - Data quality checks (duplicates / missing months)
   ============================================================ */

---------------------------------------------------------------
-- 1) Base Fact View (clean semantic layer shape)
---------------------------------------------------------------
-- Replace schema/table names as needed (e.g., dbo.fact_pnl).
-- This view standardizes scenario naming and ensures consistent types.

-- CREATE OR REPLACE VIEW vw_fact_pnl AS   -- (Postgres)
-- CREATE VIEW dbo.vw_fact_pnl AS          -- (SQL Server)
WITH base AS (
    SELECT
        CAST(month_end_date AS date)         AS month_end_date,
        EXTRACT(YEAR  FROM month_end_date)   AS year,
        EXTRACT(MONTH FROM month_end_date)   AS month_number,
        CASE
            WHEN EXTRACT(MONTH FROM month_end_date) IN (1,2,3)   THEN 'Q1'
            WHEN EXTRACT(MONTH FROM month_end_date) IN (4,5,6)   THEN 'Q2'
            WHEN EXTRACT(MONTH FROM month_end_date) IN (7,8,9)   THEN 'Q3'
            ELSE 'Q4'
        END                                   AS quarter,
        UPPER(TRIM(scenario))                 AS scenario,
        CAST(account_code AS varchar(10))     AS account_code,
        CAST(amount AS numeric(18,2))         AS amount
    FROM fact_pnl
)
SELECT
    month_end_date,
    year,
    month_number,
    quarter,
    scenario,
    account_code,
    amount
FROM base;


---------------------------------------------------------------
-- 2) KPI Pack View (Monthly Actual vs Budget + Variance)
---------------------------------------------------------------
-- Finance logic:
-- Revenue accounts: 4001, 4002, 4010
-- COGS accounts:    5001, 5002, 5003  (usually negative)
-- OPEX accounts:    6001-6006         (usually negative)
-- EBITDA = Revenue + COGS + OPEX

WITH f AS (
    SELECT
        CAST(month_end_date AS date) AS month_end_date,
        UPPER(TRIM(scenario))        AS scenario,
        CAST(account_code AS varchar(10)) AS account_code,
        CAST(amount AS numeric(18,2))     AS amount
    FROM fact_pnl
),
monthly AS (
    SELECT
        month_end_date,
        scenario,

        SUM(CASE WHEN account_code IN ('4001','4002','4010') THEN amount ELSE 0 END) AS revenue,
        SUM(CASE WHEN account_code IN ('5001','5002','5003') THEN amount ELSE 0 END) AS cogs,
        SUM(CASE WHEN account_code IN ('6001','6002','6003','6004','6005','6006') THEN amount ELSE 0 END) AS opex

    FROM f
    GROUP BY month_end_date, scenario
),
kpi AS (
    SELECT
        month_end_date,
        scenario,
        revenue,
        cogs,
        opex,
        (revenue + cogs + opex) AS ebitda
    FROM monthly
),
pivoted AS (
    SELECT
        month_end_date,

        MAX(CASE WHEN scenario = 'ACTUAL' THEN revenue END) AS revenue_actual,
        MAX(CASE WHEN scenario = 'BUDGET' THEN revenue END) AS revenue_budget,

        MAX(CASE WHEN scenario = 'ACTUAL' THEN cogs END)    AS cogs_actual,
        MAX(CASE WHEN scenario = 'BUDGET' THEN cogs END)    AS cogs_budget,

        MAX(CASE WHEN scenario = 'ACTUAL' THEN opex END)    AS opex_actual,
        MAX(CASE WHEN scenario = 'BUDGET' THEN opex END)    AS opex_budget,

        MAX(CASE WHEN scenario = 'ACTUAL' THEN ebitda END)  AS ebitda_actual,
        MAX(CASE WHEN scenario = 'BUDGET' THEN ebitda END)  AS ebitda_budget

    FROM kpi
    GROUP BY month_end_date
)
SELECT
    month_end_date,

    revenue_actual,
    revenue_budget,
    (revenue_actual - revenue_budget) AS revenue_variance,
    CASE WHEN revenue_budget = 0 THEN NULL
         ELSE (revenue_actual - revenue_budget) / revenue_budget END AS revenue_variance_pct,

    cogs_actual,
    cogs_budget,
    (cogs_actual - cogs_budget) AS cogs_variance,
    CASE WHEN cogs_budget = 0 THEN NULL
         ELSE (cogs_actual - cogs_budget) / cogs_budget END AS cogs_variance_pct,

    opex_actual,
    opex_budget,
    (opex_actual - opex_budget) AS opex_variance,
    CASE WHEN opex_budget = 0 THEN NULL
         ELSE (opex_actual - opex_budget) / opex_budget END AS opex_variance_pct,

    ebitda_actual,
    ebitda_budget,
    (ebitda_actual - ebitda_budget) AS ebitda_variance,
    CASE WHEN ebitda_budget = 0 THEN NULL
         ELSE (ebitda_actual - ebitda_budget) / ebitda_budget END AS ebitda_variance_pct

FROM pivoted
ORDER BY month_end_date;


---------------------------------------------------------------
-- 3) Data Quality Checks
---------------------------------------------------------------

-- 3A) Duplicate key check: month_end_date + scenario + account_code
SELECT
    CAST(month_end_date AS date) AS month_end_date,
    UPPER(TRIM(scenario))        AS scenario,
    CAST(account_code AS varchar(10)) AS account_code,
    COUNT(*) AS row_count
FROM fact_pnl
GROUP BY
    CAST(month_end_date AS date),
    UPPER(TRIM(scenario)),
    CAST(account_code AS varchar(10))
HAVING COUNT(*) > 1
ORDER BY row_count DESC;


-- 3B) Missing months per scenario (requires a date dimension table dim_date)
-- Assumes dim_date has month_end_date values.
SELECT
    d.month_end_date,
    s.scenario
FROM dim_date d
CROSS JOIN (SELECT 'ACTUAL' AS scenario UNION ALL SELECT 'BUDGET') s
LEFT JOIN (
    SELECT DISTINCT CAST(month_end_date AS date) AS month_end_date,
                    UPPER(TRIM(scenario)) AS scenario
    FROM fact_pnl
) f
  ON f.month_end_date = d.month_end_date
 AND f.scenario = s.scenario
WHERE f.month_end_date IS NULL
ORDER BY d.month_end_date, s.scenario;
