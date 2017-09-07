/*
	Lee VectorStrategy
*/

select Canal, year(fecha)*100+month(fecha) as periodo , count(*)as n
from ods.dbo.fact_VectorStrategy
where sistema = 25
group by Canal, year(fecha)*100+month(fecha) 
order by Canal, year(fecha)*100+month(fecha)

select A.* 
--- Para el analisis del Champion
,ChmpChll = case 
		when substring(rut,len(rut),1) in (0,1,2) and substring(Path_Solicitud,1,1) = 'P'
			and periodo >= 201606
			then 'Personal Banking Governance Excepciones' 
		when substring(rut,len(rut),1) in (3,4,5) and substring(Path_Solicitud,1,3) in ('P02','P05')
			and periodo >= 201606
			then 'Personal Banking Nuevos Factores TDSR' 
		when substring(Path_Solicitud,1,1) = 'P'
			then 'Personal Banking Cartera Campeona'
		else 'Consumer Finance' end
from (
select year(fecha)*100+month(fecha) as periodo
	,sistema
	,solicitud
	,rut
	,c_sis
	,F_ProcesoSW
	,H_ProcesoSW
	,Decision_Group
	,n_LogicPath
	,ScoreSinacofiCliente
	,n_puntajesw
	,m_MontoSolicitado
	,m_DeudaConsumoTotalSSA
	,Canal
	,Fecha as FechaFVS
	,b_ClienteCompraCartera
	,n_risk
	,sbif_r04_dir_cli
	,m_LeverageNoHipotecarioProyectado
	,m_TDSR_ClienteProyectado
	,m_SolicitadoConsumoCuota
	,m_MontoTarjeta
	,m_MontoLinea
	,ProductoConsumo
	,ProductoTarjeta
	,ProductoLinea
	,m_RentaUFCalculada = m_RentaUFCalculada/100
	,space(3) as Res_Final
	,space(1) as CatExc
	,space(1) as cruce_x
	,space(1) as FProceso
	,space(1) as cascada
	,Pidio_CS = case when isnull(ltrim(rtrim(productoconsumo)),'')<>'' and ltrim(rtrim(productoconsumo))<>'99999' then 1 else 0 end
	,Pidio_TC = case when isnull(ltrim(rtrim(productotarjeta)),'')<>'' and ltrim(rtrim(productotarjeta))<>'99999'  then 1 else 0 end
	,Pidio_LC = case when isnull(ltrim(rtrim(productolinea)),'')<>'' and ltrim(rtrim(productolinea))<>'99999'   then 1 else 0 end
	,AnoMes = convert(varchar(6),F_ProcesoSW,112)
	,Path_Solicitud = case
			when c_sis = 'MGN' and n_LogicPath = 110 then 'P01'
			when c_sis = 'MGN' and n_LogicPath = 121 then 'P02'
			when c_sis = 'MGN' and n_LogicPath = 130 then 'P03'
			when c_sis = 'MGN' and n_LogicPath = 140 then 'P04'
			when c_sis = 'MGN' and n_LogicPath = 150 then 'P05'
			when c_sis = 'MGN' and n_LogicPath = 161 then 'P06'
			when c_sis = 'CMP' and n_LogicPath = 110 then 'C01'
			when c_sis = 'CMP' and n_LogicPath = 121 then 'C02'
			when c_sis = 'CMP' and n_LogicPath = 130 then 'C03'
			when c_sis = 'CMP' and n_LogicPath = 140 then 'C04'
			when c_sis = 'CMP' and n_LogicPath = 150 then 'C05'
			when c_sis = 'CMP' and n_LogicPath = 161 then 'C06'
			else 'Sin Data' end
	,Canal_Venta_SW = case when Canal = 'A' then 'Sucursal'
	   					   when Canal = 'B' then 'FFVV'
						   end
	,cTypes = case
			when n_LogicPath = 110 then 'Prospecto No Campaña'
			when n_LogicPath = 121 then 'Cliente Campaña'
			when n_LogicPath = 130 then 'Renegociado'
			when n_LogicPath = 140 then 'Moroso'
			when n_LogicPath = 150 then 'Cliente No Campaña'
			when n_LogicPath = 161 then 'Prospecto Campaña'
			else 'Sin Campaña' end
	,Est1raVlta = case 
		   when decision_group = 'AC' then 'Aprobado'
		   when decision_group = 'DL' then 'Rechazo'
		   when decision_group = 'RV' then 'Zona Gris'
		   when decision_group = 'VF' then 'Devueltas'
		  else 'Sin Dato 1ra Vuelta' end
	,Origen = 'SW'
	--into ##FcVc_SW
from ods.dbo.fact_VectorStrategy
where convert(varchar(6),F_ProcesoSW,112) = '201605' and Decision_Group <> 'VF'
) as A
where A.Path_Solicitud in ('P01','P02','P05','P06','P07','P09','P10','Sin Data','')
and cTypes in ('Cliente Campaña','Cliente No Campaña','Prospecto Campaña','Prospecto No Campaña','Sin Campaña')
