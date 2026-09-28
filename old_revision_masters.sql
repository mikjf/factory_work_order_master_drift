-- Master drift - masters for levels 2 and 3, 2026-09-28. READ-ONLY.
-- Promoted from probe 4 of probe_queries.sql (same query, same scope).
-- probe_queries.sql - the probes that settled status codes, sites and
-- unreleased work orders - was deleted on 2026-09-28; it stays in git
-- history.
--
-- For every part with an open work order on an OLDER revision than today's
-- (WO.USER_8 <> PART_SITE.ENGINEERING_MSTR), every master operation of:
--   ROLE 'OWN' = the revision the work order was built from (WO.USER_8)
--   ROLE 'CUR' = the part's current revision (ENGINEERING_MSTR)
-- one row per master operation. master_drift.py reads this next to the
-- output of master_drift.sql. Run both on the same day - masters change.
--
-- MASTER_EXISTS = 0: no TYPE 'M' row for that revision (then no operations).
-- A work order with USER_8 NULL gives an 'OWN' row with REV NULL and
-- MASTER_EXISTS = 0.
--
-- The work-order scope MUST stay in step with master_drift.sql
-- (kinds, statuses, part filter), or level 2 will miss parts.

WITH OLD_WO AS (
    SELECT DISTINCT
        WO.PART_ID,
        WO.USER_8           AS OWN_REV,
        PS.ENGINEERING_MSTR AS CUR_REV
    FROM dbo.WORK_ORDER AS WO
    INNER JOIN dbo.PART_SITE AS PS
        ON PS.PART_ID = WO.PART_ID
    WHERE WO.TYPE = 'W'
      AND WO.STATUS IN ('F', 'R')                            -- as master_drift.sql
      AND (WO.BASE_ID LIKE '%OP' OR WO.BASE_ID LIKE '%OPP')  -- as master_drift.sql
      AND WO.PART_ID LIKE 'M.%'                              -- as master_drift.sql
      AND (WO.USER_8 <> PS.ENGINEERING_MSTR OR WO.USER_8 IS NULL)
),
REVS AS (
    SELECT PART_ID, OWN_REV AS REV, 'OWN' AS ROLE FROM OLD_WO
    UNION
    SELECT PART_ID, CUR_REV AS REV, 'CUR' AS ROLE FROM OLD_WO
)
SELECT
    R.PART_ID,
    R.ROLE,
    R.REV,
    CASE WHEN WO_M.BASE_ID IS NULL THEN 0 ELSE 1 END AS MASTER_EXISTS,
    WO_M.BASE_ID          AS MASTER_BASE_ID,
    WO_M.STATUS           AS MASTER_STATUS,
    O.SEQUENCE_NO,
    O.OPERATION_TYPE      AS OPERATION,
    O.RESOURCE_ID,
    O.RUN_TYPE,
    O.RUN,
    O.LOAD_SIZE_QTY,
    O.SETUP_HRS,
    O.USER_5
FROM REVS AS R
LEFT JOIN dbo.WORK_ORDER AS WO_M
    ON  WO_M.PART_ID = R.PART_ID
    AND WO_M.LOT_ID  = R.REV
    AND WO_M.TYPE    = 'M'
LEFT JOIN dbo.OPERATION AS O
    ON  O.WORKORDER_BASE_ID  = WO_M.BASE_ID
    AND O.WORKORDER_LOT_ID   = WO_M.LOT_ID
    AND O.WORKORDER_SPLIT_ID = WO_M.SPLIT_ID
    AND O.WORKORDER_SUB_ID   = WO_M.SUB_ID
    AND O.WORKORDER_TYPE     = WO_M.TYPE
ORDER BY R.PART_ID, R.ROLE DESC, R.REV, O.SEQUENCE_NO;
