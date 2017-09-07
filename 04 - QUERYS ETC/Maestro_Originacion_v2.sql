
/*
Se encuentra disponible para vuestro uso el Datamart de Originaci�n con informaci�n desde el 01/01/2013 hasta 31/05/2016. 
En el podr�n encontrar toda la informaci�n disponible capturada desde la Fact_Solicitud, Fact_VectorStrategy
, Fact_VectorOMDM, Fact_Excepciones

La informaci�n ha sido reorganizada y agrupada bajo 6 entidades y se encuentra interconectada a trav�s de un Surrogate Key 
(SolNumero, PerNumero, EvaNumero, IndNumero, ExcNumero, ProNumero) el cual es propio del Datamart y es utilizado para vincular 
la entidades indistintamente del origen.

Podr�n encontrar en el modelo adjunto la informaci�n de todos los nombres de campos y tablas que la componen.

Se debe tener presente, que el Datamart se encuentra pensado para dar respuesta a diferentes tipos de necesidades, 
esto quiere decir, que  va capturando los distintos eventos en el ciclo de Vida de una Solicitud, por lo cual las 
Evaluaciones (SW y OMDM) consideran iteraciones en la l�nea del tiempo.  (M�s adelante se muestra como recuperar una iteraci�n)

Por otro lado, y pensando solamente en las solicitudes de Ventas, existen dos comportamientos que se deber�n tener presente: 
Una �Solicitud de Cliente� (en el futuro Solicitud Padre) generara una  �Solicitud de Venta� (en el futuro Solicitud Hija), 
por lo cual se puede rastrear el origen a trav�s de la Dependencia (SolDependencia, PerDependencia, EvaDependencia, 
IndDependencia, ExcDependencia, ProDependencia), o bien, se puede crear �nicamente una Solicitud de Venta sin ninguna dependencia.

Como anexo a lo anterior, las Fuentes de Originaci�n (PDN y EIAP) pueden generar diferentes tipos de solicitudes, las cuales 
tambi�n son capturadas y almacenadas.

Finalmente, les adjunto un Query Maestro para recuperar toda la informaci�n disponible.

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
	-- El perfil son atributos nominales que muy rara vez cambian en una interaci�n, por eso unico (ultimo capturado)  
	on ( solpadre.SolNumero = perfil.PerNumero )
left join [RT_ORIGINACION].[MIS].[OriProducto] producto -- select * from [RT_ORIGINACION].[MIS].[OriProducto]
	-- El Producto siempre esta asociado a una Solicitud de Venta
	on ( solhija.SolNumero = producto.ProNumero )
left join [RT_ORIGINACION].[MIS].[OriIndicador] indicador -- select * from [RT_ORIGINACION].[MIS].[OriIndicador]
	on (-- and solpadre.SolNumero = indicador.IndNumero and indicador.IndIteracion=1  
		-- Corresponde a los Indicadores utilizados en la Primera Evaluaci�n  
		solpadre.SolNumero = indicador.IndNumero and indicador.IndIteracion=solpadre.SolUltimaIteracion  --Corresponde a los Indicadores de la ultima Iteracion 
		)
left join [RT_ORIGINACION].[MIS].[OriEvaluacion] evaluacion -- select * from [RT_ORIGINACION].[MIS].[OriEvaluacion]
	on (-- and solpadre.SolNumero = evaluacion.EvaNumero and evaluacion.EvaIteracion = 1  
		-- Corresponde a los resultados de la Primera Evaluaci�n  
		solpadre.SolNumero = evaluacion.EvaNumero and evaluacion.EvaIteracion = solpadre.SolUltimaIteracion  --Corresponde a los resultados de la �ltima Evaluaci�n    
		)
--where solhija.SolFechaCreacion between '20160101' and '20160731' -- Filtro de Per�odo
where solhija.SolFechaCreacion >= '2015-01-01' -- Filtro de Per�odo
and solhija.SolCategoriaSolicitud = 'Venta' -- Pare efectos de analisis, tomar siempre la venta
order by solhija.solnumero, solhija.soldependencia

