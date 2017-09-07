


select A.*
into [RT_ANALYTICS_ORI].dbo.ME_Fact_VectorOMDM_2016_2017
from (select row_number() over(partition by id_solicitud order by Fecha_Proceso + HoraProceso desc) as rk, *
	from MACA.[ODS].[dbo].[Fact_VectorOMDM]
	where EtapaEvaluacion = 'P2' and DecisionSolicitud <> 'VF') as A
where A.rk = 1 and Fecha_Proceso >= '2016-01-01' --and Fecha_Proceso <= '2016-12-31' 
--and TipoCliente_Titular = 1


select *
from [RT_ANALYTICS_ORI].dbo.ME_Fact_VectorOMDM_2016_2017
