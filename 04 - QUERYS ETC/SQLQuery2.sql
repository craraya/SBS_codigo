
/*
	Base CLLOGDRR01p\ODS
	CA
*/

-- LTV
select *
from ods.dbo.Fact_LTV_20150331

-- Operaciones: 
--	TipoOperacion ['',A,P]
--	Estado ['',1,2,3,4,5,6,9]
--	Producto [muchos]
select *
from ods.dbo.Fact_Operaciones_20160531

select *
from ods.dbo.Fact_Mrie_20160531

select *
from ods.dbo.Fact_TarjetaCreditos_20160531

select t_operacion, count(*) as n
from ods.dbo.Fact_OperacionesContables_20160531
group by t_operacion

-- revision politica originacion item 2.14

select *
from RT_ORIGINACION.MIS.OriSolicitud

select *
,row_number() over(partition by soldependencia order by soldependencia, solnumero) as rk
from RT_ORIGINACION.MIS.OriSolicitud
where SolFechaCreacion >= '20160101'
--and ltrim(SolCategoriaSolicitud) in ('Solicitud')
order by solDependencia, solNumero, rk


select SolUltimaResolucion, count(*) as n
from RT_ORIGINACION.MIS.OriSolicitud
where SolFechaCreacion >= '20160101'
group by SolUltimaResolucion

select *, year(SolFechaCreacion)*100+month(SolFechaCreacion) as periodo
from RT_ORIGINACION.MIS.OriSolicitud
where SolFechaCreacion >= '2016-01-01'



