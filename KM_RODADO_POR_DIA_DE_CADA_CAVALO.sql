DECLARE @query AS NVARCHAR(MAX);

CREATE TABLE #OrderedDates (DATA DATE);

INSERT INTO #OrderedDates (DATA)
SELECT DISTINCT CONVERT(DATE, DATHOR) AS DATA
FROM RODVRA (NOLOCK)
ORDER BY DATA;

SET @query = '
WITH DailyKMT AS (
    SELECT 
        CODVEI,
        CONVERT(DATE, DATHOR) AS DATA,
        MAX(KMTVEI) AS ULTKMT,
        MIN(KMTVEI) AS PRIMEIRO_KMT
    FROM 
        RODVRA (NOLOCK)
    WHERE 
        KMTVEI <> 0
    GROUP BY 
        CODVEI,
        CONVERT(DATE, DATHOR)
),
AllDates AS (
    SELECT DISTINCT
        CODVEI,
        d.DATA
    FROM 
        RODVEI v (NOLOCK)
    CROSS JOIN 
        #OrderedDates d
    WHERE 
        v.PROPRI = ''S''
        AND v.RAS_ID IS NOT NULL
        AND v.SITUAC = ''1''
)
SELECT 
    ad.CODVEI,
    v.NUMVEI,
    CONVERT(VARCHAR(10), ad.DATA, 103) AS Datas,
    FORMAT(ISNULL((dk.ULTKMT - dk.PRIMEIRO_KMT) / 1000, 0), ''N2'') AS KM_RODADO
FROM 
    AllDates ad
LEFT JOIN 
    DailyKMT dk ON ad.CODVEI = dk.CODVEI AND ad.DATA = dk.DATA
LEFT JOIN 
    RODVEI v (NOLOCK) ON ad.CODVEI = v.CODVEI
WHERE 
    v.PROPRI = ''S''
    AND v.RAS_ID IS NOT NULL
    AND v.SITUAC = ''1''
ORDER BY 
    ad.CODVEI, ad.DATA;';

EXEC sp_executesql @query;

DROP TABLE #OrderedDates;
