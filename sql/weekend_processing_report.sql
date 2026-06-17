DROP TABLE IF EXISTS タイプA注文;
CREATE TEMP TABLE タイプA注文 AS (
    SELECT
        事業者ID
        , 商品ID
        , 注文ID
        , 数量
        , 売上高
    FROM
        企画タイプA注文
    WHERE
        地域 = '関東'
        AND テーブル更新日 BETWEEN CURRENT_DATE - 27 AND CURRENT_DATE
        AND 注文状況 != '返品'
);


DROP TABLE IF EXISTS 商品取扱_タイプA注文除外;
CREATE TEMP TABLE 商品取扱_タイプA注文除外 AS (
        WITH sub AS (
            SELECT
                事業者ID
                , 注文ID
                , 取扱日時
                , 取扱数量
            FROM
                商品取扱
            WHERE
                地域 = '関東'
                AND DATE(取扱日時) BETWEEN CURRENT_DATE - 27 AND CURRENT_DATE
        )
    
    SELECT
        事業者ID
        , SUM(CASE WHEN DATE_PART(dow, sub.取扱日時) = 6 THEN sub.取扱数量 ELSE 0 END) AS 土曜日合計取扱数量
        , SUM(CASE WHEN DATE_PART(dow, sub.取扱日時) = 0 THEN sub.取扱数量 ELSE 0 END) AS 日曜日合計取扱数量
        , SUM(sub.取扱数量) AS 合計取扱数量
    FROM
        sub
        LEFT JOIN タイプA注文 tao
            ON sub.注文ID = tao.注文ID
    WHERE
        tao.注文ID IS NULL
    GROUP BY
        1
);


DROP TABLE IF EXISTS 週末取扱;
CREATE TEMP TABLE 週末取扱 AS (
        WITH 有効事業者 AS (
            SELECT
                事業者ID
                , 事業者管理ID
                , 退会日時
            FROM
                事業者管理
            WHERE
               地域 = '関東'
               AND (DATE(退会日時) IS NULL OR
                    DATE(退会日時) >= CURRENT_DATE) -- NULLの場合は事業者の最新レコードであることを意味します。
        )
    
        , sub AS (
            SELECT
                eff.事業者ID
                , CASE WHEN opn.土曜日営業可否 = 'Y' THEN 'Y' ELSE 'N' END AS 土曜日営業可否
                , CASE WHEN opn.日曜日営業可否 = 'Y' THEN 'Y' ELSE 'N' END AS 日曜日営業可否
                , ROW_NUMBER() OVER (PARTITION BY eff.事業者ID ORDER BY eff.退会日時 DESC NULLS FIRST) AS 更新順位
            FROM
                有効事業者 eff
                INNER JOIN 事業者営業活動 opn
                    ON eff.事業者管理ID = opn.事業者管理ID
                    AND (opn.土曜日営業可否 = 'Y' OR opn.日曜日営業可否 = 'Y')
        )
    
    SELECT *
    FROM sub
    WHERE 更新順位 = 1 -- 重複削除
);


SELECT DISTINCT
    sp.事業者ID
    , we.土曜日営業可否
    , we.日曜日営業可否
    , sp.土曜日合計取扱数量 / NULLIF(sp.合計取扱数量, 0)::float AS 土曜日取扱割合 -- 直近4週間の土曜日に取扱した商品の比率
    , sp.日曜日合計取扱数量 / NULLIF(sp.合計取扱数量, 0)::float AS 日曜日取扱割合 -- 直近4週間の日曜日に取扱した商品の比率
FROM
    商品取扱_タイプA注文除外 sp
    LEFT JOIN 週末取扱 we
        ON sp.事業者ID = we.事業者ID
;