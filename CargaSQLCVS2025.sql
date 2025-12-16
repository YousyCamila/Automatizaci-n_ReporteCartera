-----------------------------------------------------

-- 1. Crear tabla temporal y tabla para archivos

-----------------------------------------------------

IF OBJECT_ID('tempdb..#TempProcesos') IS NOT NULL DROP TABLE #TempProcesos;

CREATE TABLE #TempProcesos (

    ID_PROCESS BIGINT,

    RESPONSE NVARCHAR(MAX),

    DATE_LOG_TXT NVARCHAR (50),

    HORA_TXT NVARCHAR (50)

);

IF OBJECT_ID('tempdb..#files') IS NOT NULL DROP TABLE #files;

CREATE TABLE #files (archivo NVARCHAR(260));

-----------------------------------------------------

-- 2. Configuración de carpeta

-----------------------------------------------------

DECLARE @FolderPath NVARCHAR(500) = 'C:\test\2024_csv\';

DECLARE @cmd NVARCHAR(1000);

DECLARE @archivo NVARCHAR(260);

-----------------------------------------------------

-- 3. Obtener lista de archivos CSV

-----------------------------------------------------

SET @cmd = 'dir "' + @FolderPath + '*.csv" /b';

INSERT INTO #files (archivo)

EXEC xp_cmdshell @cmd;

-----------------------------------------------------

-- 4. Recorrer CSVs

-----------------------------------------------------

DECLARE cursor_csv CURSOR FOR 

SELECT archivo FROM #files WHERE archivo IS NOT NULL;

OPEN cursor_csv;

FETCH NEXT FROM cursor_csv INTO @archivo;

WHILE @@FETCH_STATUS = 0

BEGIN

    DECLARE @fullpath NVARCHAR(500) = @FolderPath + @archivo;

    PRINT '=== Cargando archivo: ' + @fullpath + ' ===';

    -----------------------------------------------------

    -- 5. Limpiar tabla temporal

    -----------------------------------------------------

    TRUNCATE TABLE #TempProcesos;

    -----------------------------------------------------

    -- 6. Cargar CSV actual en tabla temporal

    -----------------------------------------------------

    DECLARE @bulk NVARCHAR(MAX) = '

        BULK INSERT #TempProcesos

        FROM ''' + @fullpath + '''

        WITH (

            FIELDTERMINATOR = '';'',

            ROWTERMINATOR = ''\n'',

            FIRSTROW = 2

        );

    ';

    EXEC (@bulk);

    -----------------------------------------------------

    -- 7. Insertar solo los NO duplicados 

    --    (ni en el CSV ni en la tabla final)

    -----------------------------------------------------

INSERT INTO LogProcesos (ID_PROCESS, RESPONSE, DATE_LOG, HORA)
 
SELECT
 
    t.ID_PROCESS,
 
    t.RESPONSE,
 
    CAST(
 
        TRY_CONVERT(
 
            datetime,
 
            REPLACE(REPLACE(t.DATE_LOG_TXT, ' a. m.', ''), ' p. m.', ''),
 
            103
 
        ) AS date
 
    ) AS DATE_LOG,
 
    CAST(
 
        TRY_CONVERT(
 
            datetime,
 
            REPLACE(REPLACE(t.DATE_LOG_TXT, ' a. m.', ''), ' p. m.', ''),
 
            103
 
        ) AS time
 
    ) AS HORA
 
FROM (
 
    SELECT
 
        ID_PROCESS,
 
        RESPONSE,
 
        DATE_LOG_TXT,
 
        HORA_TXT,
 
        ROW_NUMBER() OVER (PARTITION BY ID_PROCESS ORDER BY ID_PROCESS) AS rn
 
    FROM #TempProcesos
 
) t
 
WHERE t.rn = 1
 
AND NOT EXISTS (
 
    SELECT 1
 
    FROM LogProcesos l
 
    WHERE l.ID_PROCESS = t.ID_PROCESS
 
);
 


    -----------------------------------------------------

    FETCH NEXT FROM cursor_csv INTO @archivo;

END

CLOSE cursor_csv;

DEALLOCATE cursor_csv;

PRINT '=== PROCESO FINALIZADO ===';
 