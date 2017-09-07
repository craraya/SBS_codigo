/*
	Resumen Histórico Renta
*/
insert into #aux

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140131_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140228_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140331_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140430_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140530_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140630_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140731_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140829_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20140930_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20141030_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20141128_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20141230_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150130_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150227_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150331_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150430_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150529_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150630_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150731_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150831_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20150930_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20151030_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20151130_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20151230_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160129_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160229_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160331_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160429_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160531_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')

union

select cast(FechaOtorgamiento/100 as int) as fecha_ot
, year(fecha)*100+month(fecha) as periodo, segmento, tipo_DeudaCartaMensual, rut, rtrim(n_operacion) as n_operacion, m_InfTotal, m_Informado
from RT_OPERACIONES.ew.Operaciones_Critical_20160630_CIE
where tipo_DeudaCartaMensual in ('Línea de Crédito - Consumo','Renegociado - Consumo','Tarjeta de Crédito - Consumo','Crédito en Cuotas','Crédito en Cuotas (DFC)')
and m_InfTotal>0 and segmento in ('Banca Consumo','Personas')
