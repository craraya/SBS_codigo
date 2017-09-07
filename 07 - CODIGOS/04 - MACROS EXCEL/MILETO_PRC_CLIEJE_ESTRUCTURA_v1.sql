/*

	Tabla:
	[MILETO].[dbo].[TBL_CLIEJE_ESTRUCTURA]

*/

----------------------------------------------------------------------------------------------------------------------------------------------------
-- Ejecutivos Especialistas.

if object_id('tempdb..#Eje_Cash') is not null
drop table #Eje_Cash
go
select *
into #Eje_Cash
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\Ejecutivos_Especialistas.xlsx','SELECT * FROM [Eje_Cash$]')

if object_id('tempdb..#Eje_Btx') is not null
drop table #Eje_Btx
go
select *
into #Eje_Btx
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\Ejecutivos_Especialistas.xlsx','SELECT * FROM [Eje_Btx$]')

if object_id('tempdb..#Eje_Comex') is not null
drop table #Eje_Comex
go
select *
into #Eje_Comex
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\Ejecutivos_Especialistas.xlsx','SELECT * FROM [Eje_Comex$]')

if object_id('tempdb..#Eje_Factoring') is not null
drop table #Eje_Factoring
go
select *
into #Eje_Factoring
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\Ejecutivos_Especialistas.xlsx','SELECT * FROM [Eje_Factoring$]')

if object_id('tempdb..#Eje_Leasing') is not null
drop table #Eje_Leasing
go
select *
into #Eje_Leasing
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\Ejecutivos_Especialistas.xlsx','SELECT * FROM [Eje_Leasing$]')

----------------------------------------------------------------------------------------------------------------------------------------------------
-- Estructura Pyme
if object_id('tempdb..#Pyme') is not null
drop table #Pyme
go
select *
into #Pyme
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;HDR=YES;Database=U:\XLS\Estructuras\ESTRUCTURA FULL.xlsx','SELECT * FROM [Struct.$]')

----------------------------------------------------------------------------------------------------------------------------------------------------
-- Tabla Final

if object_id('tempdb..#Cliente_Eje') is not null
drop table #Cliente_Eje
go
select [Ano Mes] as Ano_Mes, [Rut Cli] as rut, [Vr Cli] as dv, [Nom Cli] as nombre, Bca, [Macro Banca] as Macro_Banca
	, Regional ,Plataforma, Ofi, Eje
into #Cliente_Eje
from MILETO.[dbo].[TBL_CLIENTE_EJECUTIVO]
where [Macro Banca] in ('Empresas','Grandes Empresas','Inmobiliaria','Empresarios','Emprendedores') 
and [ano mes] = (select max([ano mes]) from MILETO.[dbo].[TBL_CLIENTE_EJECUTIVO])

update A set A.Eje = B.Eje
from #Cliente_Eje as A, [BCI_MarketingEmpresas_ODS].[dbo].[Tab_StockCtaCte] as B
where A.rut = B.rut

TRUNCATE TABLE MILETO.[dbo].[TBL_CLIEJE_ESTRUCTURA]

INSERT INTO MILETO.[dbo].[TBL_CLIEJE_ESTRUCTURA]
select getdate() as fecha_proceso, CLIEJE.*
,CASH.Eje_Bel as Eje_Cash, COMEX.Eje_Comex, BTX.Eje_Bel as Eje_Btx, FAC.Eje_Fact, LEA.Eje_Lea
,TIERS.[Tier Final] as Segmento_Eje
,EP.Segmento as Segmento_Pyme, INST.Segmento_Institucional
from #Cliente_Eje as CLIEJE
LEFT JOIN (select distinct * from #Eje_Cash) as CASH on CLIEJE.Eje = CASH.Eje_Com
LEFT JOIN (select distinct * from #Eje_Comex) as COMEX on CLIEJE.Eje = COMEX.Eje_Com
LEFT JOIN (select distinct * from #Eje_Btx) as BTX on CLIEJE.Eje = BTX.Eje_Com
LEFT JOIN (select distinct * from #Eje_Factoring) as FAC on CLIEJE.Eje = FAC.Eje_Com
LEFT JOIN (select distinct * from #Eje_Leasing) as LEA on CLIEJE.Eje = LEA.Eje_Com
LEFT JOIN CAA.[dbo].[TBL_TIERS] as TIERS on CLIEJE.rut = TIERS.[Rut Cli] and CLIEJE.Ano_Mes = TIERS.[Ano Mes]
LEFT JOIN (select distinct [USER], Segmento from #Pyme) as EP on CLIEJE.Eje = EP.[USER]
LEFT JOIN CAA.dbo.TEMP_Segmento_Institucional_201505 as INST on CLIEJE.rut = INST.rut
order by CLIEJE.rut

/*

select *
from MILETO.[dbo].[TBL_CLIEJE_ESTRUCTURA]

*/