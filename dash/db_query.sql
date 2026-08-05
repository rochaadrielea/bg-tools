-- ================================================================
-- queries.sql
-- ----------------------------------------------------------------
-- All SQL queries that build the dashboard charts.
-- Open this file in SQLite Viewer. To run a single query:
--   1. Click anywhere inside the query you want
--   2. Highlight it (from -- Query N: down to the semicolon ;)
--   3. Right-click -> "Run selected" (or press F5)
-- ================================================================
-- database: C:/Users/BL3914/OneDrive - Beyond Gravity Services AG/Documents/Development/dash/quality.db

-- ================================================================
-- Query 1: The 5 KPI metrics at top of Slide 2
-- Chart: NCs Open, Closed, Major, Production, Supplier
-- ================================================================
SELECT
    SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS ncs_wip,
    SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS ncs_closed,
    SUM(CASE WHEN classification LIKE 'Major%' AND is_open = 1 THEN 1 ELSE 0 END) AS major_open,
    SUM(CASE WHEN is_supplier_nc = 1 AND is_open = 1 THEN 1 ELSE 0 END) AS supplier_open,
    SUM(CASE WHEN is_supplier_nc = 0 AND is_open = 1 THEN 1 ELSE 0 END) AS production_open,
    SUM(CASE WHEN project IS NULL THEN 1 ELSE 0 END) AS blank_project,
    SUM(CASE WHEN detection_area IS NULL THEN 1 ELSE 0 END) AS blank_detection
FROM nc;

-- PATH: C:/Users/BL3914/OneDrive - Beyond Gravity Services AG\Documents\Development\dash\quality.db
-- ================================================================
-- Query 2: Top 6 Projects - NCs WIP (Slide 3 top-left chart)
-- ================================================================
SELECT
    COALESCE(project, '(no project)') AS project,
    COUNT(*) AS open_ncs
FROM nc
WHERE is_open = 1
GROUP BY project
ORDER BY open_ncs DESC
LIMIT 6;


-- ================================================================
-- Query 3: NCs by Detection Area YTD 2026 (Slide 3 top-right chart)
-- ================================================================
SELECT
    COALESCE(detection_area, 'BLANK - to clean') AS area,
    COUNT(*) AS n
FROM nc
WHERE substr(created_on, 1, 4) = '2026'
GROUP BY area
ORDER BY n DESC
LIMIT 12;


-- ================================================================
-- Query 4: Top 6 Projects Opened YTD 2026 (Slide 3 bottom-left chart)
-- ================================================================
SELECT
    COALESCE(project, '(no project)') AS project,
    COUNT(*) AS n
FROM nc
WHERE substr(created_on, 1, 4) = '2026'
GROUP BY project
ORDER BY n DESC
LIMIT 6;


-- ================================================================
-- Query 5: NC Total per Year (Slide 3 bottom-right chart)
-- ================================================================
SELECT
    substr(created_on, 1, 4) AS yr,
    COUNT(*) AS n
FROM nc
WHERE created_on IS NOT NULL
GROUP BY yr
ORDER BY yr;


-- ================================================================
-- Query 6: Monthly Trend Opens vs Closes (Slide 4 line chart)
-- ================================================================
WITH opens AS (
    SELECT substr(created_on, 1, 7) AS month, COUNT(*) AS opened
    FROM nc
    WHERE created_on IS NOT NULL
    GROUP BY month
),
closes AS (
    SELECT substr(closure_date, 1, 7) AS month, COUNT(*) AS closed
    FROM nc
    WHERE closure_date IS NOT NULL
    GROUP BY month
)
SELECT
    o.month,
    o.opened,
    COALESCE(c.closed, 0) AS closed
FROM opens o
LEFT JOIN closes c USING (month)
WHERE o.month >= '2026-01'
ORDER BY o.month;


-- ================================================================
-- Query 7: Production vs Supplier (Slide 4 bottom-left chart)
-- ================================================================
SELECT
    CASE WHEN is_supplier_nc = 1 THEN 'Supplier' ELSE 'Production' END AS source,
    SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
    SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_ncs
FROM nc
GROUP BY source;


-- ================================================================
-- Query 8: Open NCs by Owner (Slide 4 bottom-right table)
-- ================================================================
SELECT
    COALESCE(owner, '(no owner)') AS owner,
    SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
    MAX(days_open) AS oldest_days
FROM nc
GROUP BY owner
HAVING open_ncs > 0
ORDER BY open_ncs DESC
LIMIT 10;


-- ================================================================
-- Query 9: Data Quality - rows needing cleaning
-- ================================================================
SELECT
    nc_id,
    owner,
    project,
    detection_area,
    classification,
    status
FROM nc
WHERE project IS NULL
   OR detection_area IS NULL
   OR classification IS NULL
   OR owner IS NULL
ORDER BY owner, nc_id;


-- ================================================================
-- BONUS Query 10: NCs by classification level
-- ================================================================
SELECT
    COALESCE(classification, '(no class)') AS class,
    COUNT(*) AS n
FROM nc
GROUP BY class
ORDER BY n DESC;


-- ================================================================
-- BONUS Query 11: 10 oldest open NCs (urgent focus)
-- ================================================================
SELECT
    nc_id,
    owner,
    project,
    days_open,
    created_on
FROM nc
WHERE is_open = 1
ORDER BY days_open DESC
LIMIT 10;


-- ================================================================
-- BONUS Query 12: Closure rate by owner
-- ================================================================
SELECT
    owner,
    COUNT(*) AS total,
    SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed,
    ROUND(100.0 * SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_closed
FROM nc
GROUP BY owner
ORDER BY total DESC;


-- ================================================================
-- BONUS Query 13: NCs opened per month (all time)
-- ================================================================
SELECT
    substr(created_on, 1, 7) AS month,
    COUNT(*) AS n
FROM nc
WHERE created_on IS NOT NULL
GROUP BY month
ORDER BY month;