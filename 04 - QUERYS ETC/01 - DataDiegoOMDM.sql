----------------------------------------------------------------------------------------------------------
--***************************************************
print 'Genera Data Temporal OMDM del Fact_VectorOMDM'
--***************************************************

declare @ExeSql nvarchar(max)
declare @FecProc varchar(06)
declare @Iter int
declare @CntFil int
declare @FileSW varchar(60)
declare @FileOMDM varchar(60)
declare @FileFinal varchar(60)
declare @FecCM varchar(08)
declare @FecIni varchar(06)
declare @FecFin varchar(06)

set @FecIni = '201511'  /* Fecha Inicio Proceso (Inclusive) - aaaamm */
set @FecFin = '201511'  /* Fecha Fin Proceso (Inclusive) - aaaamm */

set @Iter = (select datediff(month,substring(@FecIni,1,4)+'-'+substring(@FecIni,5,2)+'-01',
                                   substring(@FecFin,1,4)+'-'+substring(@FecFin,5,2)+'-01')) + 1

set @CntFil = 1
while @CntFil <= @Iter
  begin 
  

set @FecProc = @FecIni  /* Fecha del Proceso - aaaamm */

IF object_id('tempdb..##FV_OMDM') is not null drop table ##FV_OMDM

set @ExeSql = ''
set @ExeSql = '
	select * ,OriDta = ''OMDM''
into ##FV_OMDM
from ods.dbo.fact_vectoromdm
where convert(varchar(6),Fecha,112) = '+@FecProc+'
and EtapaEvaluacion = ''P2''
and DecisionSolicitud <> ''VF''
AND id_Solicitud NOT IN (100218,100102,20052,20079,10049,10014,10057)'
Execute sp_executesql @ExeSql

----------------------------------------------------------------------------------------------------------

--***************************************************
print 'Ordena Data Temporal OMDM del Fact_VectorOMDM'
--***************************************************

if object_id('tempdb..#FV_OMDM1') is not null drop table #FV_OMDM1

select ROW_NUMBER() over(partition by Rut order by Rut,Fecha desc) as N_reg2,*
into #FV_OMDM1
from ##FV_OMDM

----------------------------------------------------------------------------------------------------------

--***********************************************************************
print 'Selecciona Registro Unico Tabla Temporal OMDM del Fact_VectorOMDM'
--***********************************************************************

if object_id('tempdb..#FV_OMDM2') is not null drop table #FV_OMDM2

select * 
into #FV_OMDM2
from #FV_OMDM1 where N_reg2 = 1

----------------------------------------------------------------------------------------------------------
--****************************************************
print 'Genera Data Temporal OMDM del Fact_Excepciones'
--****************************************************

IF object_id('tempdb..##FE_OMDM') is not null drop table ##FE_OMDM

set @ExeSql = ''
set @ExeSql = '
	select *
	into ##FE_OMDM
	FROM ods.dbo.fact_excepciones
	WHERE Sistema = 0 AND
	t_llamada = ''P2'' AND
	c_Observacion <> ''RSV'' and
	convert(varchar(6),Fecha,112) = '+@FecProc+' AND
	Solicitud NOT IN (100218,100102,20052,20079,10049,10014,10057)'
Execute sp_executesql @ExeSql

----------------------------------------------------------------------------------------------------------
--********************************************************************************************
print 'Join de Tablas Temporales Fact_VectorOMDM y Fact_Excepciones - Genera Excepciones OMDM'
--********************************************************************************************

if object_id('tempdb..#FVyFE_OMDM') is not null drop table #FVyFE_OMDM

select a.*,isnull(b.Fecha,'') as FechaExc,b.Sistema as SistemaExc,b.Solicitud as SolicitudExc,b.c_Observacion,b.TipoExcepcion,b.t_llamada
       ,space(30) as Clasificacion
into #FVyFE_OMDM
from #FV_OMDM2 a
Left join ##FE_OMDM b 
  on a.Id_Solicitud = b.solicitud and 
     DecisionSolicitud <> 'AC' and
     (convert(char(8),a.Fecha_Proceso, 112)+substring(a.HoraProceso,1,2)+substring(a.HoraProceso,4,2)+substring(a.HoraProceso,7,2)) 
      between 
        (convert(char(8), b.fecha,112)+
         substring(convert(char(8),dateadd(ss,-300, b.fecha),108),1,2)+
         substring(convert(char(8),dateadd(ss,-300, b.fecha),108),4,2)+
         substring(convert(char(8),dateadd(ss,-300, b.fecha),108),7,2))
      and
        (convert(char(8), b.fecha,112)+
         substring(convert(char(8),dateadd(ss,300, b.fecha),108),1,2)+
         substring(convert(char(8),dateadd(ss,300, b.fecha),108),4,2)+
         substring(convert(char(8),dateadd(ss,300, b.fecha),108),7,2))

update a
 set FechaExc = b.Fecha
    ,SistemaExc = b.Sistema
    ,SolicitudExc = b.Solicitud
    ,c_Observacion = b.c_Observacion
    ,TipoExcepcion = b.TipoExcepcion
    ,t_llamada = b.t_llamada
from #FVyFE_OMDM a
join ##FE_OMDM b 
  on a.Id_Solicitud = b.solicitud and 
     a.DecisionSolicitud <> 'AC' 
where a.SolicitudExc is null

----------------------------------------------------------------------------------------------------------
--*****************************************************
print 'Actualiza Clasificación de las Excepciones OMDM'
--*****************************************************


update a
 set Clasificacion = (select Clasificacion from BD_Triad.dbo.Tabla_Unica_Excepciones b where a.c_Observacion = b.RC and b.Motor = 'OMDM')
from #FVyFE_OMDM a


----------------------------------------------------------------------------------------------------------
--*****************************
print 'Genera Tabla Final OMDM'
--*****************************

     if object_id('tempdb..##TablaFinal_OMDM')is null
       begin 
         set @ExeSql = 'select * into ##TablaFinal_OMDM from #FVyFE_OMDM'
         EXECUTE sp_executesql @ExeSql
       end
     else
      begin  
     if object_id('tempdb..##TablaFinal_OMDM')is not null
       begin 
         set @ExeSql = 'insert into ##TablaFinal_OMDM select * from #FVyFE_OMDM'
         EXECUTE sp_executesql @ExeSql
       end
      end


  set @CntFil = (@CntFil + 1)
  set @FecIni = convert(char(8),dateadd(month,1,substring(@FecIni,1,4)+'-'+substring(@FecIni,5,2)+'-01'),112)
end

--select c_Observacion from ##TablaFinal_OMDM
--group by c_Observacion
--order by c_Observacion