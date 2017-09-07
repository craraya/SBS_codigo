/*
	Lee OMDM
*/

select A.*,B.sistema
,ChmpChll = case 
		when substring(A.rut,len(A.rut),1) in (0,1,2) and substring(Path_Solicitud,1,1) = 'P'
			--and anomesOMDM >= 201606
			then 'Personal Banking Governance Excepciones' 
		when substring(A.rut,len(A.rut),1) in (3,4,5) and substring(Path_Solicitud,1,3) in ('P02','P05')
			--and anomesOMDM >= 201606
			then 'Personal Banking Nuevos Factores TDSR' 
		when substring(Path_Solicitud,1,1) = 'P'
			then 'Personal Banking Cartera Campeona'
		else 'Consumer Finance' end
from (
select cast(rut as int) as rut_i, * from ods.dbo.fact_vectoromdm
where year(fecha)*100+month(fecha) = 201605) as A -- 13.423
LEFT JOIN (
		select distinct rut, Solicitud, sistema, canal
		, Canal_Venta_SW = case when canal = 1 then 'Sucursal' when canal = 8 then 'FFVV' else '' end
		--,row_number() over(partition by rut, solicitud order by fecha desc)
		--into ##FcSo_OMDM
		from ods.dbo.Fact_Solicitud
		where convert(varchar(6),Fecha,112) = '201605'
		and c_Escenario = 'I'
		and EstadoMotor NOT IN ('', 'V')
		AND Solicitud NOT IN (100218,100102,20052,20079,10049,10014,10057)
		--and sistema = 25
		) as B
		on A.rut_i = B.rut and A.id_solicitud = B.Solicitud
where B.sistema = 25
and Path_Solicitud in ('P01','P02','P05','P06','P07','P09','P10','Sin Data','')











select year(fecha)*100+month(fecha) as periodo, canal , count(*)as n
from ods.dbo.fact_vectoromdm
where sistema = 25
group by year(fecha)*100+month(fecha), Canal
order by year(fecha)*100+month(fecha), Canal

select id_Solicitud
	,Rut
	,Path_Solicitud
	,Fecha
	,DecisionSolicitud
	,SegmentoCliente_Titular
	,RiskIndicator_Titular
	,m_SolicitadoCS
	,m_SolicitadoLC
	,m_SolicitadoTC
	,Fecha_Proceso
	,HoraProceso
	,AnoMesOMDM = convert(varchar(6),Fecha,112)
	,c_sis
	,cTypes = case
			when Path_Solicitud = ''P01'' then ''Prospecto No Campaña''
			when Path_Solicitud = ''P02'' then ''Cliente Campaña''
			when Path_Solicitud = ''P03'' then ''Renegociado''
			when Path_Solicitud = ''P04'' then ''Moroso''
			when Path_Solicitud = ''P05'' then ''Cliente No Campaña''
			when Path_Solicitud = ''P06'' then ''Prospecto Campaña''
			else ''Sin Campaña'' end
	,Pidio_CS = case when m_SolicitadoCS>0 then 1 else 0 end
	,Pidio_LC = case when m_SolicitadoLC>0 then 1 else 0 end
	,Pidio_TC = case when m_SolicitadoTC>0 then 1 else 0 end
	,Est1raVlta = case when decisionSolicitud = ''AC'' then ''Aprobado''
						when decisionSolicitud = ''DL'' then ''Rechazo''
						when decisionSolicitud = ''IN'' then ''Zona Gris''
						when decisionSolicitud = ''VF'' then ''Devueltas''
					else ''Sin Dato 1ra Vuelta'' end
	,RiskIndicator =
		case
			when RiskIndicator_Titular = 1 then ''A''
			when RiskIndicator_Titular = 2 then ''B''
			when RiskIndicator_Titular = 3 then ''C''
			when RiskIndicator_Titular = 4 then ''D''
			when RiskIndicator_Titular = 5 then ''E''
		else '''' end
	,origen = ''OMDM''
into ##FcVc_OMDM
from maca.ods.dbo.fact_vectoromdm
where convert(varchar(6),Fecha,112) = '+@FecProc+'
and EtapaEvaluacion = ''P2''
and DecisionSolicitud <> ''VF''
AND id_Solicitud NOT IN (100218,100102,20052,20079,10049,10014,10057)'