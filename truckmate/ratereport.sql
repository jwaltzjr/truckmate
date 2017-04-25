SELECT
  CELL.RATING_ID "TARIFF",
  (
    SELECT LISTAGG(CRP.CLIENT_ID, ', ')
    FROM TMWIN.CLIENT_RATE_PROFILE CRP
    WHERE CRP.RATING_ID = CELL.RATING_ID
  ) "CUSTOMERS",
  HEAD.START_ZONE "ORIGIN",
  (
    SELECT LISTAGG(RSZ.ZONE_ID, ', ')
    FROM TMWIN.RATE_STARTZONES RSZ
    WHERE RSZ.RATING_ID = CELL.RATING_ID
  ) "ORIGIN_MS",
  ROW.TO_ZONE "DESTINATION",
  COL.BREAK_VALUE "BREAK",
  COL.IS_MINIMUM "IS_MIN",
  CELL.RATE_VALUE "RATE"
FROM TMWIN.RATE_CELL CELL
INNER JOIN TMWIN.RATE_HEADER HEAD
  ON HEAD.RATING_ID = CELL.RATING_ID
INNER JOIN TMWIN.RATE_ROW ROW
  ON CELL.RATE_ROW_ID = ROW.RATE_ROW_ID
  AND ROW.RATING_ID = CELL.RATING_ID
INNER JOIN TMWIN.RATE_COL COL
  ON CELL.RATE_COL_ID = COL.RATE_COL_ID
  AND COL.RATING_ID = CELL.RATING_ID
WHERE HEAD.UNIT_FACTOR = 100
AND HEAD.BREAK_UNIT = 'LB'
AND HEAD.SHEET_TYPE = 0
