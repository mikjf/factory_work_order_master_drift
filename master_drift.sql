-- Master drift - level 1, since 2026-09-28. READ-ONLY (SELECT only).
-- Run by extract.py (read-only connection) together with
-- old_revision_masters.sql; never edited for Excel's sake.
-- One row per open, not-started work-order operation, with the work order's
-- copy and its own revision's master side by side, and one flag per field.
-- This is LEVEL 1 of master drift; levels 2 and 3 are built by
-- master_drift.py from this output plus old_revision_masters.sql.
--
-- NAMING: WO_* = the work order's own copy (dbo.OPERATION of the work order):
--                what the master said when the WO was created, unless the
--                WO was edited since.
--         M_*  = the master of the WO's OWN revision (WO.USER_8), today.
--                WO <> M: master edited in place after the WO was created,
--                or the WO edited by hand.
-- fts.sql calls the work order's copy "*_MASTER" - not here, on purpose.
-- Matched like fts.sql: same SEQUENCE_NO AND same OPERATION_TYPE. Exact here,
-- because the WO copied its own revision line by line. Since 2026-10-21 a
-- line whose sequence is no longer in the master is paired by name when the
-- name is unique in the master and in the WO (MATCHED_BY = 'UNIQUE_NAME'):
-- the master was renumbered in place after the WO was created.
--
-- NOT HERE: comparison with the CURRENT revision's master. Tried on
-- 2026-09-28 with the same sequence + name matching: it paired only 508 of
-- 1,401 older-revision lines, because a new revision inserts a step and
-- shifts every later step down one slot. Cross-revision pairing needs rules
-- SQL expresses badly, so it lives in master_drift.py (level 3). Since
-- 2026-10-21 (PM and SC) the CURRENT revision is the reference for an order
-- on an older revision; this query still delivers the own-revision side.
--
-- SETTLED:
--   - "not started" = WO status F/R, op status F/R, no labour ticket.
--     U work orders are out (user, 2026-09-28): probe 3 showed them to be
--     old orders, none with a ticket in 2026.
--   - PART_SITE joined on PART_ID only; probe 2a (2026-09-28) found one row
--     per part, so no fan-out today.
-- ASSUMED, watched by master_drift.py:
--   - O.RESOURCE_ID is the PLANNED resource (VISUAL naming, unconfirmed).
--   - Master looked up by PART_ID + LOT_ID + TYPE, as fts.sql does; if a
--     master can have several SPLIT/SUB rows this fans out - the script
--     stops on a duplicate operation key.

SELECT
    WO.BASE_ID                  AS WORK_ORDER,
    WO.LOT_ID,
    WO.SPLIT_ID,
    WO.SUB_ID,
    CASE WHEN WO.BASE_ID LIKE '%OPP' THEN 'OPP'
         WHEN WO.BASE_ID LIKE '%OP'  THEN 'OP'
    END                         AS WO_KIND,
    WO.STATUS                   AS WO_STATUS,
    WO.USER_9                   AS PROJECT_CODE,
    (
        SELECT STRING_VAL
        FROM dbo.USER_DEF_FIELDS AS UDF
        WHERE DOCUMENT_ID = P.ID AND ID = 'UDF-0000026'
    )                           AS FAMILY,
    P.ID                        AS PART_ID,
    P.DESCRIPTION               AS PART_DESCRIPTION,
    WO.USER_8                   AS WO_REV,
    PS.ENGINEERING_MSTR         AS CURRENT_REV,
    -- Agreed: a WO on an older revision is its own category (routing changed),
    -- listed apart. 1 = built from the current master revision.
    CASE WHEN WO.USER_8 = PS.ENGINEERING_MSTR THEN 1 ELSE 0 END
                                AS ON_CURRENT_REV,
    WO.CREATE_DATE,
    WO.DESIRED_RLS_DATE,
    WO.DESIRED_WANT_DATE,
    O.SEQUENCE_NO               AS OPERATION_SEQUENCE,
    O.OPERATION_TYPE            AS OPERATION,
    O.STATUS                    AS OP_STATUS,
    -- ADDED 2026-10-12: the work order operation's own ROWID, so a line of
    -- the list can be found in the ERP directly. master_drift.py accepts
    -- exports without it (blank column).
    O.ROWID                     AS WORKORDER_OPS_ROWID,

    -- side by side: work order copy | master
    -- Read by extract.py straight from the database (2026-10-15), so resource
    -- IDs keep their leading zeros ('034') without the 'RES:' text marker
    -- they carried for the paste into Excel on 13-14 Oct.
    O.RESOURCE_ID               AS WO_RESOURCE_ID,
    O_M.RESOURCE_ID             AS M_RESOURCE_ID,
    O.RUN_TYPE                  AS WO_RUN_TYPE,
    O_M.RUN_TYPE                AS M_RUN_TYPE,
    O.RUN                       AS WO_RUN,
    O_M.RUN                     AS M_RUN,
    O.LOAD_SIZE_QTY             AS WO_LOAD_SIZE_QTY,
    O_M.LOAD_SIZE_QTY           AS M_LOAD_SIZE_QTY,
    O.SETUP_HRS                 AS WO_SETUP_HRS,
    O_M.SETUP_HRS               AS M_SETUP_HRS,
    O.USER_5                    AS WO_USER_5,
    O_M.USER_5                  AS M_USER_5,
    -- NOT compared: RUN_HRS and CALC_START_QTY depend on the order quantity
    -- (RUN_HRS is also derived and rounded), so they differ from the master
    -- by design.

    O_M.ROWID                   AS MASTER_OPS_ROWID,
    O_M.SEQUENCE_NO             AS M_SEQUENCE_NO,
    -- how the master line was found (2026-10-21, see the OUTER APPLY below):
    -- SEQUENCE    same sequence and same name - the WO copied this line
    -- UNIQUE_NAME the sequence differs (the master was renumbered after the
    --             WO was created: a step inserted, later ones shifted), but
    --             the name occurs exactly once in the master AND exactly once
    --             among all operations of the WO, started ones included
    O_M.MATCHED_BY,
    -- 1 = no master line found either way
    CASE WHEN O_M.ROWID IS NULL THEN 1 ELSE 0 END
                                AS NO_MASTER_LINE,
    -- why not: NO_MASTER        no master for the WO's revision at all
    --          NOT_IN_MASTER    the name is not in the master - a step added
    --                           to the WO on the go (rework S RIL ..., samples)
    --          NAME_NOT_UNIQUE  the name is in the master, but repeats in the
    --                           master or in the WO - not paired, no guessing
    CASE WHEN O_M.ROWID IS NOT NULL THEN NULL
         WHEN WO_M.BASE_ID IS NULL  THEN 'NO_MASTER'
         WHEN N_M.N = 0             THEN 'NOT_IN_MASTER'
         ELSE 'NAME_NOT_UNIQUE'
    END                         AS NO_MASTER_REASON,

    -- one flag per field; NULL = NULL counts as equal. Numbers are compared
    -- as numbers (a text compare gave false changes in FTS).
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN O.RESOURCE_ID   = O_M.RESOURCE_ID   OR (O.RESOURCE_ID   IS NULL AND O_M.RESOURCE_ID   IS NULL) THEN 0 ELSE 1 END AS D_RESOURCE,
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN O.RUN_TYPE      = O_M.RUN_TYPE      OR (O.RUN_TYPE      IS NULL AND O_M.RUN_TYPE      IS NULL) THEN 0 ELSE 1 END AS D_RUN_TYPE,
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN O.RUN           = O_M.RUN           OR (O.RUN           IS NULL AND O_M.RUN           IS NULL) THEN 0 ELSE 1 END AS D_RUN,
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN O.LOAD_SIZE_QTY = O_M.LOAD_SIZE_QTY OR (O.LOAD_SIZE_QTY IS NULL AND O_M.LOAD_SIZE_QTY IS NULL) THEN 0 ELSE 1 END AS D_LOAD_SIZE_QTY,
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN O.SETUP_HRS     = O_M.SETUP_HRS     OR (O.SETUP_HRS     IS NULL AND O_M.SETUP_HRS     IS NULL) THEN 0 ELSE 1 END AS D_SETUP_HRS,
    -- crew (USER_5, a text column) compared as a number. Since 2026-10-23 an
    -- EMPTY crew is a difference against a filled one (user: the crew must be
    -- filled in correctly); empty against empty is equal. Until then it was
    -- ISNULL(O.USER_5, 1) = ISNULL(O_M.USER_5, 1) - empty counted as 1.
    CASE WHEN O_M.ROWID IS NULL THEN NULL
         WHEN TRY_CONVERT(decimal(9, 2), O.USER_5) = TRY_CONVERT(decimal(9, 2), O_M.USER_5)
           OR (O.USER_5 IS NULL AND O_M.USER_5 IS NULL) THEN 0 ELSE 1 END AS D_USER_5

FROM dbo.WORK_ORDER AS WO
INNER JOIN dbo.PART AS P
    ON P.ID = WO.PART_ID
INNER JOIN dbo.OPERATION AS O
    ON  O.WORKORDER_BASE_ID  = WO.BASE_ID
    AND O.WORKORDER_LOT_ID   = WO.LOT_ID
    AND O.WORKORDER_SPLIT_ID = WO.SPLIT_ID
    AND O.WORKORDER_SUB_ID   = WO.SUB_ID
    AND O.WORKORDER_TYPE     = WO.TYPE
-- the current revision, next to the WO's own one
LEFT JOIN dbo.PART_SITE AS PS
    ON PS.PART_ID = P.ID
-- the master of the WO's OWN revision (agreed); the current one comes from
-- old_revision_masters.sql
LEFT JOIN dbo.WORK_ORDER AS WO_M
    ON  WO_M.PART_ID = WO.PART_ID
    AND WO_M.LOT_ID  = WO.USER_8
    AND WO_M.TYPE    = 'M'
-- how often this operation name occurs in the master, and among ALL
-- operations of the work order (any status: a rework copy of a step already
-- done still counts)
OUTER APPLY (
    SELECT COUNT(*) AS N
    FROM dbo.OPERATION AS X
    WHERE X.WORKORDER_BASE_ID  = WO_M.BASE_ID
      AND X.WORKORDER_LOT_ID   = WO_M.LOT_ID
      AND X.WORKORDER_SPLIT_ID = WO_M.SPLIT_ID
      AND X.WORKORDER_SUB_ID   = WO_M.SUB_ID
      AND X.WORKORDER_TYPE     = WO_M.TYPE
      AND X.OPERATION_TYPE     = O.OPERATION_TYPE
) AS N_M
OUTER APPLY (
    SELECT COUNT(*) AS N
    FROM dbo.OPERATION AS Y
    WHERE Y.WORKORDER_BASE_ID  = O.WORKORDER_BASE_ID
      AND Y.WORKORDER_LOT_ID   = O.WORKORDER_LOT_ID
      AND Y.WORKORDER_SPLIT_ID = O.WORKORDER_SPLIT_ID
      AND Y.WORKORDER_SUB_ID   = O.WORKORDER_SUB_ID
      AND Y.WORKORDER_TYPE     = O.WORKORDER_TYPE
      AND Y.OPERATION_TYPE     = O.OPERATION_TYPE
) AS N_W
-- the master line: same sequence AND same name first, as in fts.sql
-- (sequence alone can pair with a repurposed slot). Failing that, the line
-- with the same name - only when the name is unique on both sides (user,
-- 2026-10-21: a renumbered master hid real differences as "no master line").
OUTER APPLY (
    SELECT TOP 1
        X.ROWID, X.SEQUENCE_NO, X.RESOURCE_ID, X.RUN_TYPE, X.RUN,
        X.LOAD_SIZE_QTY, X.SETUP_HRS, X.USER_5,
        CASE WHEN X.SEQUENCE_NO = O.SEQUENCE_NO THEN 'SEQUENCE' ELSE 'UNIQUE_NAME' END AS MATCHED_BY
    FROM dbo.OPERATION AS X
    WHERE X.WORKORDER_BASE_ID  = WO_M.BASE_ID
      AND X.WORKORDER_LOT_ID   = WO_M.LOT_ID
      AND X.WORKORDER_SPLIT_ID = WO_M.SPLIT_ID
      AND X.WORKORDER_SUB_ID   = WO_M.SUB_ID
      AND X.WORKORDER_TYPE     = WO_M.TYPE
      AND X.OPERATION_TYPE     = O.OPERATION_TYPE
      AND (X.SEQUENCE_NO = O.SEQUENCE_NO OR (N_M.N = 1 AND N_W.N = 1))
    ORDER BY CASE WHEN X.SEQUENCE_NO = O.SEQUENCE_NO THEN 0 ELSE 1 END
) AS O_M
WHERE WO.TYPE = 'W'
  AND WO.STATUS IN ('F', 'R')
  -- not closed / cancelled / unreleased. Probe 1b: F/F 2,836 ops, R/R 7,252.
  AND O.STATUS IN ('F', 'R')
  -- not started: no labour ticket of any kind on this operation
  AND NOT EXISTS (
      SELECT 1
      FROM dbo.LABOR_TICKET AS LT
      WHERE LT.WORKORDER_BASE_ID  = O.WORKORDER_BASE_ID
        AND LT.WORKORDER_LOT_ID   = O.WORKORDER_LOT_ID
        AND LT.WORKORDER_SPLIT_ID = O.WORKORDER_SPLIT_ID
        AND LT.WORKORDER_SUB_ID   = O.WORKORDER_SUB_ID
        AND LT.WORKORDER_TYPE     = O.WORKORDER_TYPE
        AND LT.OPERATION_SEQ_NO   = O.SEQUENCE_NO
  )
  -- external operations are out (user, 2026-09-28). In the first run the
  -- 'E ' name prefix, an E* resource and a DAYS/* run type coincided on
  -- exactly the same 248 rows, so the name alone identifies them.
  AND O.OPERATION_TYPE NOT LIKE 'E %'
  -- 4-digit sequences only, as in FTS: masters were renumbered to 4 digits,
  -- so a 2/3-digit line belongs to an old routing. PARKED: an old work
  -- order released this year keeps its short numbers and falls out here -
  -- see OPEN_QUESTIONS.md. First run: 175 such rows on 16 work orders.
  AND LEN(CAST(O.SEQUENCE_NO AS VARCHAR(10))) = 4
  -- work-order kinds by BASE_ID suffix (user, 2026-09-28): OP production,
  -- OPP production on a new project - both in. OR (rework) is OUT: its
  -- routing is most likely an ad-hoc master, so a difference is not drift.
  -- OS is something else and stays out too.
  AND (   WO.BASE_ID LIKE '%OP'
       OR WO.BASE_ID LIKE '%OPP')
  AND P.ID LIKE 'M.%'                                -- SCOPE, as in FTS and the probes
ORDER BY
    WO.DESIRED_RLS_DATE,
    WO.BASE_ID,
    O.SEQUENCE_NO;
