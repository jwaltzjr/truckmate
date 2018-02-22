WITH ORDERS AS (
  SELECT
    BILL_NUMBER,
    DELIVER_BY,
    DATE(DELIVER_BY) - (DAYOFWEEK(DATE(DELIVER_BY))-1) DAYS "DELIVERY_WEEK",
    WEIGHT,
    PALLETS,
    AREA,
    CASE
      WHEN TMWIN.KRC_GET_COMPANY(DETAIL_LINE_ID) IS NOT NULL THEN
        TMWIN.KRC_GET_COMPANY(DETAIL_LINE_ID)
      WHEN TMWIN.KRC_GET_INTERLINER(DETAIL_LINE_ID) IS NOT NULL THEN
        13
    END "COMPANY"
  FROM TMWIN.TLORDER
  WHERE CURRENT_STATUS NOT IN ('CANCL','QUOTE')
  AND BILL_NUMBER NOT IN ('0','NA')
  AND END_ZONE NOT IN ('STAL-TERM','COMM-TERM','KELL-TERM')
  AND SITE_ID != 'SITEA'
  AND TMWIN.KRC_IS_COMPLETE(DETAIL_LINE_ID) = 1
  AND DELIVER_BY BETWEEN ((CURRENT DATE - (DAYOFWEEK(CURRENT DATE)-1) DAYS) - 13 MONTHS)
    AND (CURRENT DATE - (DAYOFWEEK(CURRENT DATE)-1) DAYS)
)

SELECT

  -- TOTALS
  "DELIVERY_WEEK",
  COUNT(*) "NUM_ORDERS",
  CAST(SUM(WEIGHT) AS INTEGER) "WEIGHT",
  CAST(AVG(WEIGHT) AS INTEGER) "AVG_WEIGHT",
  CAST(SUM(PALLETS) AS INTEGER) "PALLETS",
  CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER) "AVG_LBS_PLT",
  CAST(SUM(AREA) AS INTEGER) "POSITIONS",
  CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER) "AVG_LBS_POS",


  -- KELLER
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "NUM_ORDERS_10",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "WEIGHT_10",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "AVG_WEIGHT_10",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "PALLETS_10",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "AVG_LBS_PLT_10",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "POSITIONS_10",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 10
  ) "AVG_LBS_POS_10",


  -- COMMERCE LOCAL
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "NUM_ORDERS_11",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "WEIGHT_11",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "AVG_WEIGHT_11",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "PALLETS_11",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "AVG_LBS_PLT_11",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "POSITIONS_11",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 11
  ) "AVG_LBS_POS_11",


  -- STALEY
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "NUM_ORDERS_12",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "WEIGHT_12",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "AVG_WEIGHT_12",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "PALLETS_12",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "AVG_LBS_PLT_12",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "POSITIONS_12",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 12
  ) "AVG_LBS_POS_12",


  -- KRC LOGISTICS
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "NUM_ORDERS_13",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "WEIGHT_13",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "AVG_WEIGHT_13",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "PALLETS_13",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "AVG_LBS_PLT_13",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "POSITIONS_13",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 13
  ) "AVG_LBS_POS_13",


  -- RTI
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "NUM_ORDERS_14",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "WEIGHT_14",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "AVG_WEIGHT_14",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "PALLETS_14",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "AVG_LBS_PLT_14",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "POSITIONS_14",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 14
  ) "AVG_LBS_POS_14",


  -- COMMERCE REGIONAL
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "NUM_ORDERS_15",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "WEIGHT_15",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "AVG_WEIGHT_15",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "PALLETS_15",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "AVG_LBS_PLT_15",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "POSITIONS_15",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" = 15
  ) "AVG_LBS_POS_15",


  -- UNDEFINED
  (
    SELECT
    CASE
      WHEN COUNT(*) = 0 THEN
        NULL
      ELSE
        COUNT(*)
    END
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "NUM_ORDERS_UNDEF",
  (
    SELECT CAST(SUM(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "WEIGHT_UNDEF",
  (
    SELECT CAST(AVG(WEIGHT) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "AVG_WEIGHT_UNDEF",
  (
    SELECT CAST(SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "PALLETS_UNDEF",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(PALLETS) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "AVG_LBS_PLT_UNDEF",
  (
    SELECT CAST(SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "POSITIONS_UNDEF",
  (
    SELECT CAST(SUM(WEIGHT)/SUM(AREA) AS INTEGER)
    FROM ORDERS O2
    WHERE O2."DELIVERY_WEEK" = O."DELIVERY_WEEK"
    AND "COMPANY" IS NULL
  ) "AVG_LBS_POS_UNDEF"


FROM ORDERS O
GROUP BY "DELIVERY_WEEK"
ORDER BY "DELIVERY_WEEK" DESC
