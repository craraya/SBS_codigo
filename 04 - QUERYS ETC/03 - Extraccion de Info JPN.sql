Información que les puede servir para bajar datos


-------------------------------------------------------------------
---Información del Cliente

-- datos generales del cliente, ejemplo:
--Renta actual, fecha nacimiento( para edad), genero, patrimonio


select   *
from ODS.dbo.Fact_Cliente_20160630  --última base vigente


-------------------------------------------------------------------
--información de las campañas

-- Estos datos son los que genera el área de campañas
--filtros de rechazo, tipo cliente, ofertas, renta real y estimada, tipos de oferta, etc

select   *
from BD_Campañas.dbo.Campaña_201608  --esta es la última campaña

-------------------------------------------------------------------

--información operaciones vigentes (stock)

--saldos, mmoras, provisiones, segmentos ,productos

select  *
from RT_OPERACIONES.ew.Operaciones_Critical_20160630_CIE --última base vigente


-------------------------------------------------------------------

--información para selección de camadas

--ejemplo para una camada del segmento Personas , crédito en cuotas

select  *
from RT_OPERACIONES.ew.Operaciones_Critical_20160630_CIE --última base vigente
where segmento='Personas' 
and tipo_DeudaCartaMensual='Crédito en Cuotas'
and FechaOtorgamiento>='20160601'



-------------------------------------------------------------------
--información bases de SBIF

select *
from SBI_DEU.dbo.Fact_LibroDeudores_201605 --última base vigente

--para Tasa Malo utilizada en Hipotecarios.
--si el cliente posee mora, cartera vencida o castigada se marca el flag

select *
from SBI_DEU.dbo.Fact_LibroDeudores_201605
where (m_deudadirectamorosa+m_deudadirectavencida+m_castigosdirectos)>0



Base Hipotecarios

Por otro lado para Hipotecarios, existe una base administrada por sic, la cual contiene datos de las solicitudes ingresadas y su estado en primera vuelta, segmentos, excepciones, etc…( David soto tambien maneja esta información)

\\Cls2521125796o\base_sic

Clave Sic2405



Base Vintage

El histotico del Vintage se encuentra en mi equipo:
C:\Respaldo\JPN\Vintage\01 – Vintage
Nombre base: “9 9 - Vintage Historico”
Acá se encuentra el vintage mora 30+ por producto y segmento 
( PB,CF;MF)  Cuotas, cuotas DFC, hipotecarios, hipotecarios FFGG, Educación,Educación, entre otros



Esto lo seguiré complementando
