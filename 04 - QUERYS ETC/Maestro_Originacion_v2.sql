
/*
Se encuentra disponible para vuestro uso el Datamart de Originación con información desde el 01/01/2013 hasta 31/05/2016. 
En el podrán encontrar toda la información disponible capturada desde la Fact_Solicitud, Fact_VectorStrategy
, Fact_VectorOMDM, Fact_Excepciones

La información ha sido reorganizada y agrupada bajo 6 entidades y se encuentra interconectada a través de un Surrogate Key 
(SolNumero, PerNumero, EvaNumero, IndNumero, ExcNumero, ProNumero) el cual es propio del Datamart y es utilizado para vincular 
la entidades indistintamente del origen.

Podrán encontrar en el modelo adjunto la información de todos los nombres de campos y tablas que la componen.

Se debe tener presente, que el Datamart se encuentra pensado para dar respuesta a diferentes tipos de necesidades, 
esto quiere decir, que  va capturando los distintos eventos en el ciclo de Vida de una Solicitud, por lo cual las 
Evaluaciones (SW y OMDM) consideran iteraciones en la línea del tiempo.  (Más adelante se muestra como recuperar una iteración)

Por otro lado, y pensando solamente en las solicitudes de Ventas, existen dos comportamientos que se deberán tener presente: 
Una “Solicitud de Cliente” (en el futuro Solicitud Padre) generara una  “Solicitud de Venta” (en el futuro Solicitud Hija), 
por lo cual se puede rastrear el origen a través de la Dependencia (SolDependencia, PerDependencia, EvaDependencia, 
IndDependencia, ExcDependencia, ProDependencia), o bien, se puede crear únicamente una Solicitud de Venta sin ninguna dependencia.

Como anexo a lo anterior, las Fuentes de Originación (PDN y EIAP) pueden generar diferentes tipos de solicitudes, las cuales 
también son capturadas y almacenadas.

Finalmente, les adjunto un Query Maestro para recuperar toda la información disponible.

cllogdrr01p\RT_ORIGINACION

*/


use RT_ORIGINACION

select year(solhija.SolFechaCreacion)*100+month(solhija.SolFechaCreacion) as periodo
-- Base de solicitudes
,solhija.*
-- Var de la solicitud Padre
,solPadre.SolCategoriaSolicitud as Categoria_Solicitud_Dependencia
,solPadre.SolSistemaNombre as Sistema_Nombre_Dependencia
-- Var del Perfil del Cliente // Todas
,perfil.*
-- Var del Producto
,producto.ProFamilia
,producto.ProProducto
,producto.ProNombreProducto
,producto.ProMontoSolicitado
-- Var Indicadores
,indicador.IndRiskIndicator
,indicador.IndTDSRActual
,indicador.IndTDSRProyectado
,indicador.IndLeverageNoHipotecario
,indicador.IndLeverageNoHipotecarioProyectado
,indicador.IndLeverageHipotecario
,indicador.IndLeverageHipotecarioProyectado
,indicador.IndSbifPeriodo
,indicador.IndSbifDeudaComercial
,indicador.IndSbifDeudaConsumo
,indicador.indSbifDeudaHipotecario
,indicador.indScoreOriginacion
,indicador.indScoreBehavior
,indicador.indScoreProvision
,indicador.indScoreSinacofi
,indicador.indTotalActivos
,indicador.indTotalPasivos
,indicador.indTotalPatrimonio
,indicador.IndFechaActualizacion as ind_fecha_act
--,evaluacion.*
from [RT_ORIGINACION].[MIS].[OriSolicitud] solhija -- select * from [RT_ORIGINACION].[MIS].[OriSolicitud]
join [RT_ORIGINACION].[MIS].[OriSolicitud] solpadre
	-- Una Solicitud de Venta no necesariamente debe tener un padre, ella a es su propio padre
	on ( solhija.soldependencia = solpadre.solnumero ) 
left join [RT_ORIGINACION].[MIS].[OriPerfil] perfil -- select * from [RT_ORIGINACION].[MIS].[OriPerfil]
	-- El perfil son atributos nominales que muy rara vez cambian en una interación, por eso unico (ultimo capturado)  
	on ( solpadre.SolNumero = perfil.PerNumero )
left join [RT_ORIGINACION].[MIS].[OriProducto] producto -- select * from [RT_ORIGINACION].[MIS].[OriProducto]
	-- El Producto siempre esta asociado a una Solicitud de Venta
	on ( solhija.SolNumero = producto.ProNumero )
left join [RT_ORIGINACION].[MIS].[OriIndicador] indicador -- select * from [RT_ORIGINACION].[MIS].[OriIndicador]
	on (-- and solpadre.SolNumero = indicador.IndNumero and indicador.IndIteracion=1  
		-- Corresponde a los Indicadores utilizados en la Primera Evaluación  
		solpadre.SolNumero = indicador.IndNumero and indicador.IndIteracion=solpadre.SolUltimaIteracion  --Corresponde a los Indicadores de la ultima Iteracion 
		)
left join [RT_ORIGINACION].[MIS].[OriEvaluacion] evaluacion -- select * from [RT_ORIGINACION].[MIS].[OriEvaluacion]
	on (-- and solpadre.SolNumero = evaluacion.EvaNumero and evaluacion.EvaIteracion = 1  
		-- Corresponde a los resultados de la Primera Evaluación  
		solpadre.SolNumero = evaluacion.EvaNumero and evaluacion.EvaIteracion = solpadre.SolUltimaIteracion  --Corresponde a los resultados de la Última Evaluación    
		)
--where solhija.SolFechaCreacion between '20160101' and '20160731' -- Filtro de Período
where solhija.SolFechaCreacion >= '2015-01-01' -- Filtro de Período
and solhija.SolCategoriaSolicitud = 'Venta' -- Pare efectos de analisis, tomar siempre la venta
order by solhija.solnumero, solhija.soldependencia

