----------------------------------------------------------------------------------------------------------
--***************************************************************
print 'Genera Data Temporal Strategyware del Fact_VectorStrategy'
--***************************************************************

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

set @FecIni = '201501'  /* Fecha Inicio Proceso (Inclusive) - aaaamm */
set @FecFin = '201502'  /* Fecha Fin Proceso (Inclusive) - aaaamm */

set @Iter = (select datediff(month,substring(@FecIni,1,4)+'-'+substring(@FecIni,5,2)+'-01',
                                   substring(@FecFin,1,4)+'-'+substring(@FecFin,5,2)+'-01')) + 1

set @CntFil = 1
while @CntFil <= @Iter
  begin 
  

set @FecProc = @FecIni  /* Fecha del Proceso - aaaamm */

if object_id ('tempdb..##FV_SW') is not null drop table ##FV_SW

set @ExeSql = ''
set @ExeSql = '	select *,OriDta = ''SW''
into ##FV_SW
from ods.dbo.fact_VectorStrategy
where convert(varchar(6),F_ProcesoSW,112) = '+@FecProc+' and Decision_Group <> ''VF'''
Execute sp_executesql @ExeSql

----------------------------------------------------------------------------------------------------------

--***************************************************************
print 'Ordena Data Temporal Strategyware del Fact_VectorStrategy'
--***************************************************************

if object_id('tempdb..#FV_SW1') is not null drop table #FV_SW1

select ROW_NUMBER() over(partition by Rut order by Rut,(F_ProcesoSW + H_ProcesoSW) desc) as N_reg1,*
into #FV_SW1
from ##FV_SW

----------------------------------------------------------------------------------------------------------

--***********************************************************************************
print 'Selecciona Registro Unico Tabla Temporal Strategyware del Fact_VectorStrategy'
--***********************************************************************************

if object_id('tempdb..#FV_SW2') is not null drop table #FV_SW2

select * 
into #FV_SW2
from #FV_SW1 where N_reg1 = 1

----------------------------------------------------------------------------------------------------------
print '**************************************'
print 'Parte Proceso para Generar Excepciones'
print '**************************************'

--************************************************************
print 'Genera Data Temporal Strategyware del Fact_Excepciones'
--************************************************************
----------------------------------------------------------------------------------------------------------

if object_id('tempdb..##FE_SW') is not null drop table ##FE_SW

set @ExeSql = ''
set @ExeSql = '
select *
into ##FE_SW
from ods.dbo.Fact_Excepciones a
Where convert(varchar(6),Fecha,112) = '+@FecProc+'
      and c_Observacion is not null 
      and TipoExcepcion = ''P'''
Execute sp_executesql @ExeSql

----------------------------------------------------------------------------------------------------------

--********************************************************************************************************
print 'Join de Tablas Temporales Fact_VectorStrategy y Fact_Excepciones - Genera Excepciones Strategyware'
--********************************************************************************************************

if object_id('tempdb..#FV_SW2yFE_SW') is not null drop table #FV_SW2yFE_SW

select a.*,isnull(b.Fecha,'') as FechaExc,b.Sistema as SistemaExc,b.Solicitud as SolicitudExc,b.c_Observacion,b.TipoExcepcion,b.t_llamada
       ,space(30) as Clasificacion
into #FV_SW2yFE_SW
from #FV_SW2 a
Left join ##FE_SW b 
  on a.solicitud = b.solicitud and 
     Decision_Group <> 'AC' and
     (convert(char(8),a.F_ProcesoSW, 112)+substring(a.H_ProcesoSW,1,2)+substring(a.H_ProcesoSW,4,2)+substring(a.H_ProcesoSW,7,2)) 
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

--segunda actualización de excepciones 

update a
 set FechaExc = b.Fecha
    ,SistemaExc = b.Sistema
    ,SolicitudExc = b.Solicitud
    ,c_Observacion = b.c_Observacion
    ,TipoExcepcion = b.TipoExcepcion
    ,t_llamada = b.t_llamada
from #FV_SW2yFE_SW a
join ##FE_SW b 
  on a.solicitud = b.solicitud and 
     a.Decision_Group <> 'AC' 
where a.SolicitudExc is null

--*************************************************************
print 'Actualiza Clasificación de las Excepciones Strategyware'
--*************************************************************

update a
 set Clasificacion = (select Clasificacion from BD_Triad.dbo.Tabla_Unica_Excepciones b where a.c_Observacion = b.RC and b.Motor = 'SW')
from #FV_SW2yFE_SW a

----------------------------------------------------------------------------------------------------------
--*************************************
print 'Genera Tabla Final Strategyware'
--*************************************

     if object_id('tempdb..##TablaFinal_SW')is null
       begin 
         set @ExeSql = 'select * into ##TablaFinal_SW from #FV_SW2yFE_SW'
         EXECUTE sp_executesql @ExeSql
       end
     else
      begin  
     if object_id('tempdb..##TablaFinal_SW')is not null
       begin 
         set @ExeSql = 'insert into ##TablaFinal_SW select * from #FV_SW2yFE_SW'
         EXECUTE sp_executesql @ExeSql
       end
      end


  set @CntFil = (@CntFil + 1)
  set @FecIni = convert(char(8),dateadd(month,1,substring(@FecIni,1,4)+'-'+substring(@FecIni,5,2)+'-01'),112)
end


--select * from #FV_SW2yFE_SW
--truncate table ##TablaFinal
--select count(*) from ##TablaFinal