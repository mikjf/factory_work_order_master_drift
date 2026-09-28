-- Master drift - probe queries, 2026-09-28. READ-ONLY (SELECT only).
-- Purpose: settle two open questions before the first real query.
-- Scope mirrors FTS: parts LIKE 'M.%', production work orders BASE_ID LIKE '%OP%'.
-- To widen, delete the two scope lines marked SCOPE.


-- ============================================================================
-- 1a. WORK_ORDER.STATUS - which values exist, for work orders AND masters.
--     One row per TYPE x STATUS. Masters are shown so their status values are
--     not mistaken for work-order ones later.
-- ============================================================================
SELECT
    WO.TYPE,
    WO.STATUS,
    COUNT(*) AS N_ROWS
FROM dbo.WORK_ORDER AS WO
WHERE WO.PART_ID LIKE 'M.%'                          -- SCOPE
GROUP BY WO.TYPE, WO.STATUS
ORDER BY WO.TYPE, WO.STATUS;


-- ============================================================================
-- 1b. Operation status x work-order status, production work orders only,
--     and how many of those operations already have ANY labour ticket.
--     This is the evidence for "not started": an operation status that has
--     tickets on it is not a "not started" status.
-- ============================================================================
-- The ticket check sits in the inner query because SQL Server refuses a
-- subquery inside SUM() (Msg 130). The inner query is one row per operation,
-- so the outer counts are counts of operations.
SELECT
    WO_STATUS,
    OP_STATUS,
    COUNT(*)                 AS N_OPERATIONS,
    COUNT(DISTINCT WO_KEY)   AS N_WORK_ORDERS,
    SUM(HAS_TICKET)          AS N_OPS_WITH_TICKETS
FROM (
    SELECT
        WO.STATUS AS WO_STATUS,
        O.STATUS  AS OP_STATUS,
        CONCAT(WO.BASE_ID, '|', WO.LOT_ID, '|', WO.SPLIT_ID, '|', WO.SUB_ID) AS WO_KEY,
        CASE WHEN EXISTS (
            SELECT 1
            FROM dbo.LABOR_TICKET AS LT
            WHERE LT.WORKORDER_BASE_ID  = O.WORKORDER_BASE_ID
              AND LT.WORKORDER_LOT_ID   = O.WORKORDER_LOT_ID
              AND LT.WORKORDER_SPLIT_ID = O.WORKORDER_SPLIT_ID
              AND LT.WORKORDER_SUB_ID   = O.WORKORDER_SUB_ID
              AND LT.WORKORDER_TYPE     = O.WORKORDER_TYPE
              AND LT.OPERATION_SEQ_NO   = O.SEQUENCE_NO
        ) THEN 1 ELSE 0 END AS HAS_TICKET
    FROM dbo.WORK_ORDER AS WO
    INNER JOIN dbo.OPERATION AS O
        ON  O.WORKORDER_BASE_ID  = WO.BASE_ID
        AND O.WORKORDER_LOT_ID   = WO.LOT_ID
        AND O.WORKORDER_SPLIT_ID = WO.SPLIT_ID
        AND O.WORKORDER_SUB_ID   = WO.SUB_ID
        AND O.WORKORDER_TYPE     = WO.TYPE           -- explicit, unlike fts.sql
    WHERE WO.TYPE = 'W'
      AND WO.BASE_ID LIKE '%OP%'                     -- SCOPE
      AND WO.PART_ID LIKE 'M.%'                      -- SCOPE
) AS OPS
GROUP BY WO_STATUS, OP_STATUS
ORDER BY WO_STATUS, OP_STATUS;


-- ============================================================================
-- 2a. PART_SITE - is every part at exactly one site?
--     N_PARTS_MULTI_ROW > 0 means a join on PART_ID alone fans out.
--     N_PARTS_MULTI_REV > 0 is the dangerous case: sites disagree on the
--     current master revision (ENGINEERING_MSTR).
-- ============================================================================
SELECT
    COUNT(*)                                          AS N_PARTS,
    SUM(CASE WHEN N_ROWS > 1 THEN 1 ELSE 0 END)       AS N_PARTS_MULTI_ROW,
    SUM(CASE WHEN N_REVS > 1 THEN 1 ELSE 0 END)       AS N_PARTS_MULTI_REV,
    MAX(N_ROWS)                                       AS MAX_ROWS_PER_PART
FROM (
    SELECT
        PS.PART_ID,
        COUNT(*)                            AS N_ROWS,
        COUNT(DISTINCT PS.ENGINEERING_MSTR) AS N_REVS
    FROM dbo.PART_SITE AS PS
    WHERE PS.PART_ID LIKE 'M.%'                      -- SCOPE
    GROUP BY PS.PART_ID
) AS X;


-- ============================================================================
-- 2b. Only if 2a shows N_PARTS_MULTI_ROW > 0: the rows themselves, to see
--     which column tells the sites apart.
-- ============================================================================
SELECT TOP 50 PS.*
FROM dbo.PART_SITE AS PS
WHERE PS.PART_ID IN (
    SELECT PART_ID
    FROM dbo.PART_SITE
    WHERE PART_ID LIKE 'M.%'                         -- SCOPE
    GROUP BY PART_ID
    HAVING COUNT(*) > 1
)
ORDER BY PS.PART_ID;
