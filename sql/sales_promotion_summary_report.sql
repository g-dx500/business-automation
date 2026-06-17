WITH base AS (
    SELECT
        事業者ID
        , 商品ID
        , 販売促進種別
    FROM
        販売促進
    WHERE
        地域 = '関東'
        AND 販売促進種別 IN('販促タイプA', '販促タイプB')
        AND テーブル更新日 BETWEEN CURRENT_DATE - 6 AND CURRENT_DATE
)


SELECT
    base.事業者ID
    , COUNT(DISTINCT CASE WHEN base.販売促進種別 = '販促タイプA' THEN base.商品ID ELSE NULL END) AS 販促タイプA適用商品数
    , COUNT(DISTINCT CASE WHEN base.販売促進種別 = '販促タイプB' THEN base.商品ID ELSE NULL END) AS 販促タイプB適用商品数
FROM
    base
WHERE
    base.deal_price > 1000
    AND base.forecasted_sales > 10000
GROUP BY
    1