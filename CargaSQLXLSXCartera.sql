USE ReporteCartera;
GO

-----------------------------------------------------
-- 1. Tabla temporal
-----------------------------------------------------
IF OBJECT_ID('tempdb..#TempReporteCartera') IS NOT NULL
    DROP TABLE #TempReporteCartera;

CREATE TABLE #TempReporteCartera (
    CodIntermediario     VARCHAR(50),
    Intermediario        VARCHAR(200),
    Moneda               VARCHAR(50),
    Sector               VARCHAR(100),
    Sucursal             VARCHAR(100),
    Asegurado            VARCHAR(200),
    Ramo                 VARCHAR(150),
    NumeroFactura        VARCHAR(100),
    Poliza               VARCHAR(100),
    Endoso               VARCHAR(100),
    FechaEmision         VARCHAR(20),
    Antiguedad           VARCHAR(50),
    FechaVigenciaDesde   VARCHAR(20),
    FechaVigenciaHasta   VARCHAR(20),
    PrimaTotal           NVARCHAR(200)
);

-----------------------------------------------------
-- 2. Cargar Excel limpio
-----------------------------------------------------
INSERT INTO #TempReporteCartera
SELECT
    [COD INTERMEDIARIO],
    [INTERMEDIARIO],
    [MONEDA],
    [SECTOR],
    [SUCURSAL],
    [ASEGURADO],
    [RAMO],
    [NRO. FACTURA],
    [POLIZA],
    [ENDOSO],
    [FECHA EMISION],
    [ANTIGUEDAD],
    [FECHA VIGENCIA DESDE],
    [FECHA VIGENCIA HASTA],
    [PRIMA TOTAL]
FROM OPENROWSET(
    'Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;HDR=YES;Database=C:\test\2025_excel\ReporteCartera_Limpio.xlsx',
    'SELECT * FROM [Hoja1$]'
);

-----------------------------------------------------
-- 3. Insertar en tabla final (sin duplicar)
-----------------------------------------------------
INSERT INTO dbo.ReporteCartera (
    CodIntermediario,
    Intermediario,
    Moneda,
    Sector,
    Sucursal,
    Asegurado,
    Ramo,
    NumeroFactura,
    Poliza,
    Endoso,
    FechaEmision,
    Antiguedad,
    FechaVigenciaDesde,
    FechaVigenciaHasta,
    PrimaTotal
)
SELECT
    t.CodIntermediario,
    t.Intermediario,
    t.Moneda,
    t.Sector,
    t.Sucursal,
    t.Asegurado,
    t.Ramo,
    t.NumeroFactura,
    t.Poliza,
    t.Endoso,
    t.FechaEmision,
    t.Antiguedad,
    t.FechaVigenciaDesde,
    t.FechaVigenciaHasta,
    t.PrimaTotal
FROM #TempReporteCartera t
WHERE NOT EXISTS (
    SELECT 1
    FROM dbo.ReporteCartera r
    WHERE r.NumeroFactura = t.NumeroFactura
      AND r.Poliza = t.Poliza
);

-----------------------------------------------------
PRINT '=== CARGUE ÚNICO DE EXCEL FINALIZADO ===';
-----------------------------------------------------
