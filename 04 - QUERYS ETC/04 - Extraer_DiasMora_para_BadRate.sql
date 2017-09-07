
/*
	Selecciona los dias mora de los segmento de CF, PB y MF de los últimos 18 meses
	Over 30 a 6 meses, es el monto en mora30 al mes 6 sobre el monto en mora al mes 0
	Ever 30 a 6 meses, es el numero de Operaciones con flag de mora30 desde el mes 0 al mes 6 
	sobre el numero de Operaciones al mes 0
*/


select cast(fecha as date) as fecha, cast(substring(cast(FechaOtorgamiento as nvarchar(10)),1,4)+'-'+substring(cast(FechaOtorgamiento as nvarchar(10)),5,2)+'-'+
substring(cast(FechaOtorgamiento as nvarchar(10)),7,2) as date) as fecha_ot
, rut, segmento, tipo_DeudaCartaMensual, n_operacion
, DiasMora, m_informado
from [RT_OPERACIONES].[ew].[Operaciones_Critical_20150331_CIE]
where isnull(DiasMora,0) > 0 and segmento in ('Banca Consumo','Banca Microempresa','Personas')
order by rut, tipo_DeudaCartaMensual

union
select cast(fecha as date) as fecha, cast(substring(cast(FechaOtorgamiento as nvarchar(10)),1,4)+'-'+substring(cast(FechaOtorgamiento as nvarchar(10)),5,2)+'-'+
substring(cast(FechaOtorgamiento as nvarchar(10)),7,2) as date) as fecha_ot
, rut, segmento, tipo_DeudaCartaMensual, n_operacion
, DiasMora, m_informado
from [RT_OPERACIONES].[ew].[Operaciones_Critical_20150430_CIE]
where isnull(DiasMora,0) > 0 and segmento in ('Banca Consumo','Banca Microempresa','Personas')
order by rut, tipo_DeudaCartaMensual

select *
from [RT_OPERACIONES].[ew].[Operaciones_Critical_20150430_CIE]
