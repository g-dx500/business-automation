CREATE TEMP TABLE sub AS (
    SELECT
        事業者ID
        , 活動日
        , 売上高

        -- PY (previous year; 前年度), CY (current year; 今年度)

        , CURRENT_DATE AS latest_day_cy
        , DATE_TRUNC('year', latest_day_cy) AS year_begin_cy
        , ADD_MONTHS(latest_day_cy, -12) AS latest_day_py
        , DATE_TRUNC('year', latest_day_py) AS year_begin_py

        -- flag

        , CASE WHEN 活動日 BETWEEN ADD_MONTHS(latest_day_cy + 1, -12) AND latest_day_cy THEN 1 ELSE 0 END AS is_cy_ttm
        , CASE WHEN 活動日 BETWEEN ADD_MONTHS(latest_day_py + 1, -12) AND latest_day_py THEN 1 ELSE 0 END AS is_py_ttm

        , CASE WHEN 活動日 BETWEEN year_begin_cy AND latest_day_cy THEN 1 ELSE 0 END AS is_cy_ytd
        , CASE WHEN 活動日 BETWEEN year_begin_py AND latest_day_py THEN 1 ELSE 0 END AS is_py_ytd

        , CASE WHEN 活動日 BETWEEN latest_day_cy - 30 + 1 AND latest_day_cy THEN 1 ELSE 0 END AS is_cy_t30d
        , CASE WHEN 活動日 BETWEEN latest_day_py - 30 + 1 AND latest_day_py THEN 1 ELSE 0 END AS is_py_t30d

        , CASE WHEN 活動日 BETWEEN latest_day_cy - 7 + 1 AND latest_day_cy THEN 1 ELSE 0 END AS is_cy_t7d
        , CASE WHEN 活動日 BETWEEN latest_day_py - 7 + 1 AND latest_day_py THEN 1 ELSE 0 END AS is_py_t7d

    FROM
        営業活動 act
    WHERE
        act.地域 = '関東'
        AND act.活動日 BETWEEN ADD_MONTHS(CURRENT_DATE, -25) AND CURRENT_DATE
);

SELECT
    sub.事業者ID
    , TO_CHAR(CURRENT_DATE, 'YYYY/MM/DD') AS run_date

    -- Sales (売上高) ※計算期間別の合計

    -- TTM (Trailing Twelve Months; 直近12か月)
    , SUM(CASE WHEN is_cy_ttm = 1 THEN sub.売上高 ELSE 0 END) AS ttm_cy_sales
    , SUM(CASE WHEN is_py_ttm = 1 THEN sub.売上高 ELSE 0 END) AS ttm_py_sales

    -- YTD (Year-to-Date; 年初来)
    , SUM(CASE WHEN is_cy_ytd = 1 THEN sub.売上高 ELSE 0 END) AS sales_cy_ytd
    , SUM(CASE WHEN is_py_ytd = 1 THEN sub.売上高 ELSE 0 END) AS sales_py_ytd

    -- T30D (直近30日)
    , SUM(CASE WHEN is_cy_t30d = 1 THEN sub.売上高 ELSE 0 END) AS sales_cy_t30d
    , SUM(CASE WHEN is_py_t30d = 1 THEN sub.売上高 ELSE 0 END) AS sales_py_t30d

    -- T7D (直近7日)
    , SUM(CASE WHEN is_cy_t7d = 1 THEN sub.売上高 ELSE 0 END) AS sales_cy_t7d
    , SUM(CASE WHEN is_py_t7d = 1 THEN sub.売上高 ELSE 0 END) AS sales_py_t7d

FROM
    sub
GROUP BY
    sub.事業者ID
;