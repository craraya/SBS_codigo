USE [RT_SCORING]
GO
/****** Object:  StoredProcedure [dbo].[sp_update_solres_consumos_omdm]    Script Date: 07/20/2016 12:13:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<FERNANDO ALVAREZ SERRANO>
-- Create date: <31-03-2016>
-- Description:	< SE OBTIENEN LAS SOLICITUDES DE CREDITOS DE CONSUMO >
-- =============================================
 
ALTER PROCEDURE [dbo].[sp_update_solres_consumos_omdm]  @periodo_inicial VARCHAR(6)
/*

declare @error varchar(max),@return_value  int
exec @RETURN_VALUE= [sp_update_solres_consumos_omdm]		@periodo_inicial='201605'
SELECT	@ERROR as N'@ERROR'
SELECT @RETURN_VALUE

*/

AS
BEGIN
		
	drop table RT_FALVAREZ.DBO.paso1omdm
	SELECT *, 
       Segmento_label=CASE 
                        WHEN c_sis = 'CMP' THEN 'Consumer Finance' 
                        WHEN c_sis = 'MGN' THEN 'Personal Banking' 
                        WHEN c_sis = 'FST' THEN 'Personal Banking' 
                      END, 
       CONVERT(INT, [rut])       AS Rut_num, 
       Desc_Path_SW=CASE 
                        WHEN path_solicitud = 'P01' THEN 'Nuevo' 
                     WHEN path_solicitud = 'P02' THEN 'Antiguo Campana' 
                      WHEN path_solicitud = 'P03' THEN 'Renegociado' 
                       WHEN path_solicitud = 'P04' THEN 'Moroso' 
                        WHEN path_solicitud = 'P05' THEN 'Antiguo' 
                         WHEN path_solicitud = 'P06' THEN 'Nuevo Campana' else 'Otro' end 
       , 
       fecha_proceso AS FH_Proceso,
       score_titular as   Score_Sinacofi, 
       Score_Interno=scoreinterno_titular, 
       Score_BHV=puntajemodelo1_titular, 
       Risk_Indicator=riskindicator_titular
       , left(CONVERT(VARCHAR(6), fecha, 112),6) as camada 
	INTO   RT_FALVAREZ.DBO.paso1omdm -- select distinct LeverageActual_Titular
	FROM    LNKMACA.[ods].[dbo].[fact_vectoromdm] 
	WHERE c_sis IN ( 'FST', 'MGN', 'CMP' )
	--AND CONVERT(VARCHAR(6), fecha, 112)  between 201504 and 201512
	AND CONVERT(VARCHAR(6), fecha, 112)  =  @periodo_inicial
	and etapaEvaluacion='P2' and b_ultimocambio = 2
	
	-- select * from  RT_FALVAREZ.DBO.paso1omdm
	
	drop table  RT_FALVAREZ.DBO.maxomdm
	SELECT rut_num, 
       CONVERT(VARCHAR(6), fecha, 112) * 1 AS AnoMes, 
       Max(fh_proceso)                     FH_Proceso, 
       Count(1)                            Num_Solicitudes 
	INTO   RT_FALVAREZ.DBO.maxomdm
	FROM   RT_FALVAREZ.DBO.paso1omdm  
	GROUP  BY rut_num,    CONVERT(VARCHAR(6), fecha, 112) * 1 

	drop table RT_FALVAREZ.DBO.xxomdm 
	SELECT T1.*, 
       b.num_solicitudes 
	INTO   RT_FALVAREZ.DBO.xxomdm 
	FROM   RT_FALVAREZ.DBO.paso1omdm T1 
       JOIN (SELECT rut_num    AS rut, 
                    fh_proceso AS fechra, 
                    num_solicitudes 
             FROM   RT_FALVAREZ.DBO.maxomdm) b 
         ON T1.fh_proceso = b.fechra 
            AND T1.rut_num = b.rut 
	
	drop table RT_FALVAREZ.DBO.zieloz1omdm 
	SELECT fecha, 
    Ultimo_Estado_Solicitud=CASE WHEN estado = '1' THEN 'Ingresado' 
                                 WHEN estado = '2' THEN 'Pendiente' 
                                 WHEN estado = '3' THEN 'Aprobado' 
                                 WHEN estado = '4' THEN 'Rechazado' 
                                 WHEN estado = '5' THEN 'Pre-Aprobado' 
                                 WHEN estado = '6' THEN 'Pre-Rechazado' 
                                 WHEN estado = '7' THEN 'Anulado' 
                                 WHEN estado = '8' THEN 'Aprobado Cursado' 
                                 WHEN estado = '9' THEN 'Ratificada' 
                                 WHEN estado = '10' THEN 'Resuelta Pack' 
                                 ELSE 'Otro' 
                               END, 
       solicitud, 
       rut, 
       operacion, 
       codigopack, 
       Mto_Solic_FactSolicitud= CONVERT(FLOAT, Replace(c_montosolicitado, ',', 
                                               '.')) 
	INTO RT_FALVAREZ.DBO.zieloz1omdm -- select * 
	FROM   LNKMACA.[ods].[dbo].[fact_solicitud] 
	WHERE  b_ultimocambio = '2' 
    AND codigopack IN (SELECT solicitud FROM   RT_FALVAREZ.DBO.xxomdm ) 

    drop table RT_FALVAREZ.DBO.zieloz2omdm 
	SELECT fecha, 
    Ultimo_Estado_Solicitud=CASE 
                                 WHEN estado = '1' THEN 'Ingresado' 
                                 WHEN estado = '2' THEN 'Pendiente' 
                                 WHEN estado = '3' THEN 'Aprobado' 
                                 WHEN estado = '4' THEN 'Rechazado' 
                                 WHEN estado = '5' THEN 'Pre-Aprobado' 
                                 WHEN estado = '6' THEN 'Pre-Rechazado' 
                                 WHEN estado = '7' THEN 'Anulado' 
                                 WHEN estado = '8' THEN 'Aprobado Cursado' 
                                 WHEN estado = '9' THEN 'Ratificada' 
                                 WHEN estado = '10' THEN 'Resuelta Pack' 
                               ELSE 'Otro' 
                               END, 
       solicitud, 
       rut, 
       operacion, 
       codigopack, 
       Mto_Solic_FactSolicitud= CONVERT(FLOAT, Replace(c_montosolicitado, ',', 
                                               '.')) 
	INTO RT_FALVAREZ.DBO.zieloz2omdm --
	FROM   LNKMACA.[ods].[dbo].[fact_solicitud] 
	WHERE  b_ultimocambio = '2' 
    AND Solicitud IN (SELECT Solicitud FROM RT_FALVAREZ.DBO.xxomdm)


	-- q07
	ALTER TABLE RT_FALVAREZ.DBO.xxomdm 	ADD ultimo_estado_solicitud NVARCHAR(50) 

	drop table  RT_FALVAREZ.DBO.factsolicomdm
	SELECT * 
	INTO   RT_FALVAREZ.DBO.factsolicomdm 
	FROM   RT_FALVAREZ.DBO.zieloz2omdm 
	UNION 
	SELECT * 
	FROM   RT_FALVAREZ.DBO.zieloz1omdm 

	-- q09
	UPDATE tabla 
	SET    ultimo_estado_solicitud = a.ultimo_estado_solicitud 
	FROM   RT_FALVAREZ.DBO.xxomdm tabla, 
    RT_FALVAREZ.DBO.factsolicomdm a 
	WHERE  tabla.id_solicitud = a.solicitud 

	-- q10
	UPDATE tabla 
	SET    ultimo_estado_solicitud = a.ultimo_estado_solicitud 
	FROM   RT_FALVAREZ.DBO.xxomdm tabla, 
       RT_FALVAREZ.DBO.factsolicomdm a 
	WHERE  tabla.id_solicitud = a.codigopack 

	 drop table RT_FALVAREZ.DBO.exc1omdm
	SELECT solicitud, 
       c_observacion, 
       fecha, 
       antiguedad_lab= CASE 
                         WHEN c_observacion = 'KD' THEN 1 
                         ELSE 0 
                       END, 
       aval=CASE 
              WHEN c_observacion IN ( 'BC', 'BD', 'BE', 'DW', 
                                      'DX', 'EN', 'EO', 'EP', 
                                      'FG', 'FH', 'FI', 'FT', 
                                      'FU', 'GI', 'GJ', 'GK', 
                                      'J4', 'LJ', 'LK', 'LL', 
                                      'M6', 'W9', 'WL', 'Z1', 
                                      'EY', 'EZ', 'GT', 'GU', 
                                      'WZ', 'GT' ) THEN 1 
              ELSE 0 
            END, 
       boletin= CASE 
                  WHEN c_observacion = 'S1' THEN 1 
                  ELSE 0 
                END --boletin laboral del empleador >M$200                       
       , 
       bureau= CASE 
                 WHEN c_observacion IN ( 'B1', 'B3', 'CR', 'DM', 
                                         'DS', 'DT', 'EH', 'EI', 
                                         'EJ', 'EK', 'EL', 'EM', 
                                         'EX', 'FA', 'FB', 'FP', 
                                         'FQ', 'FZ', 'GB', 'GC', 
                                         'GD', 'GV', 'LD', 'LE', 
                                         'LF', 'PM', 'R8', 'U2', 
                                         'V1', 'WC', 'WO', 'X2', 
                                         'Z2', 'EW', 'W0', 'GS', 
                                         'GR', 'FC', 'FK' ) THEN 1 
                 ELSE 0 
               END, 
       conyuge= CASE 
                  WHEN c_observacion IN ( 'BF', 'BG', 'DU', 'DV', 
                                          'FD', 'FE', 'FF', 'FR', 
                                          'FS', 'FX', 'FY', 'GE', 
                                          'GF', 'GH', 'GO', 'GP', 
                                          'GQ', 'L2', 'LG', 'LH', 
                                          'LI', 'PN', 'PY', 'T3', 
                                          'V3', 'W3', 'WF', 'Z4', 
                                          'Z6', 'V6', 'V2', 'BB' ) THEN 1 
                  ELSE 0 
                END, 
       cta_cerrada=CASE 
                     WHEN c_observacion = 'V1'THEN 1 
                     ELSE 0 
                   END, 
       dda_declar_noActualiz=CASE 
                               WHEN c_observacion = 'U1' THEN 1 
                               ELSE 0 
                             END, 
       edad=CASE 
              WHEN c_observacion IN ( 'E1', 'EE', 'E2', 'CY', 
                                      'DH', 'B9', 'AC' ) THEN 1 
              ELSE 0 
            END, 
       empleador=CASE 
                   WHEN c_observacion IN ( 'AF', 'AM', 'DP', 'ET', 
                                           'FM', 'FO', 'V9', 'W6', 
                                           'WI', 'WW', 'DQ', 'DR', 
                                           'EF', 'EG', 'EU', 'FN', 
                                           'G4', 'S2' ) THEN 1 
                   ELSE 0 
                 END, 
       est_financ_noActualiz = CASE 
                                 WHEN c_observacion = 'K3' THEN 1 
                                 ELSE 0 
                               END, 
       Funcionario = CASE 
                       WHEN c_observacion = 'DB'THEN 1 
                       ELSE 0 
                     END, 
       LNH=CASE 
             WHEN c_observacion IN ( 'CL', 'CT', 'CO', 'CV', 
                                     'FL', 'FW', 'FJ' )THEN 1 
             ELSE 0 
           END, 
       excede_tope_tarjeta= CASE 
                              WHEN c_observacion = 'FL' THEN 1 
                              ELSE 0 
                            END, 
       MONTO =CASE 
                WHEN c_observacion IN ( 'CO', 'KF', 'KG', 'V5', 
                                        'DJ', 'DK', 'A2', 'GI', 
                                        'FT', 'FH', 'GJ' ) THEN 1 
                ELSE 0 
              END, 
       Patrimonio= CASE 
                     WHEN c_observacion IN ( 'AD', 'CM', 'AG', 'CP' ) THEN 1 
                     ELSE 0 
                   END, 
       Predictor= CASE 
                    WHEN c_observacion IN ( 'CQ', 'CW', 'KH', 'DA', 
                                            'KI', 'KK', 'KJ', 'CE', 
                                            'CF', 'KA', 'KB', 'KC', 
                                            'CG', 'CZ', 'DG' ) THEN 1 
                    ELSE 0 
                  END, 
       renta= CASE 
                WHEN c_observacion IN ( 'KE', 'DI', 'R3', 'DN', 
                                        'BS', 'BM', 'BH', 'BN', 
                                        'BI', 'BO', 'BP', 'BJ', 
                                        'BK', 'BQ', 'BL', 'BR', 
                                        'DN', 'R1', 'R2', 'R4', 'R5' ) THEN 1 
                ELSE 0 
              END, 
       TDSR= CASE 
               WHEN c_observacion IN ( 'CK', 'CS', 'CN', 'CU' ) THEN 1 
               ELSE 0 
             END, 
       Residencia=CASE 
                    WHEN c_observacion = 'Q2' THEN 1 
                    ELSE 0 
                  END, 
       Mora_comercio=CASE 
                       WHEN c_observacion = 'CR' THEN 1 
                       ELSE 0 
                     END, 
       Socio_Empresa = CASE 
                         WHEN c_observacion = 'EQ' THEN 1 
                         ELSE 0 
                       END, 
       Tipo_Cliente = CASE 
                        WHEN c_observacion = 'DL' THEN 1 
                        ELSE 0 
                      END, 
       Numero_Acreedores = CASE 
                             WHEN c_observacion = 'X1' THEN 1 
                             ELSE 0 
                           END, 
       Protesto = CASE 
                    WHEN c_observacion IN ( 'V8', 'V0', 'W9', 'W0' ) THEN 1 
                    ELSE 0 
                  END, 
       Rechazo_SIC = CASE 
                       WHEN c_observacion IN ( 'Z1', 'Z2' ) THEN 1 
                       ELSE 0 
                     END, 
       Margen = CASE 
                  WHEN c_observacion IN ( 'AN', 'DD', 'AP', 'AR', 'DE' )THEN 1 
                  ELSE 0 
                END, 
       Impedido = CASE 
                    WHEN c_observacion IN ( 'BC', 'BD', 'BE', 'B1' ) THEN 1 
                    ELSE 0 
                  END, 
       JovenProf = CASE 
                     WHEN c_observacion IN ( 'AZ', 'CA', 'B9', 'FT', 
                                             'FH', 'GJ', 'X6', 'Z3', 'B8' )THEN 
                     1 
                     ELSE 0 
                   END, 
       fallecido= CASE 
                    WHEN c_observacion = 'DY' THEN 1 
                    ELSE 0 
                  END, 
       renegociado= CASE 
                      WHEN c_observacion = 'B2' 
                            OR c_observacion = 'GK' THEN 1 
                      ELSE 0 
                    END, 
       Sin_coneccion= CASE 
                        WHEN c_observacion = 'D2' THEN 1 
                        ELSE 0 
                      END, 
       rut_no_registrado_sinacofi= CASE 
                                     WHEN c_observacion = 'LM' THEN 1 
                                     ELSE 0 
                                   END, 
       n_consultas_rut= CASE 
                          WHEN c_observacion = 'Z7' THEN 1 
                          ELSE 0 
                        END, 
       Sin_inf_score_Sinacofi= CASE 
                                 WHEN c_observacion = 'ES' THEN 1 
                                 ELSE 0 
                               END, 
       tb = CASE 
              WHEN c_observacion = 'GL' THEN 1 
              ELSE 0 
            END, 
       otro= CASE 
               WHEN c_observacion NOT IN ( 'KD', 'BC', 'BD', 'BE', 
                                           'DW', 'DX', 'EN', 'EO', 
                                           'EP', 'FG', 'FH', 'FI', 
                                           'FT', 'FU', 'GI', 'GJ', 
                                           'GK', 'J4', 'LJ', 'LK', 
                                           'LL', 'M6', 'W9', 'WL', 
                                           'Z1', 'EY', 'EZ', 'GT', 
                                           'GU', 'WZ', 'GT', 'S1', 
                                           'B1', 'B3', 'CR', 'DM', 
                                           'DS', 'DT', 'EH', 'EI', 
                                           'EJ', 'EK', 'EL', 'EM', 
                                           'EX', 'FA', 'FB', 'FP', 
                                           'FQ', 'FZ', 'GB', 'GC', 
                                           'GD', 'GV', 'LD', 'LE', 
                                           'LF', 'PM', 'R8', 'U2', 
                                           'V1', 'WC', 'WO', 'X2', 
                                           'Z2', 'EW', 'W0', 'GS', 
                                           'GR', 'FC', 'FK', 'BF', 
                                           'BG', 'DU', 'DV', 'FD', 
                                           'FE', 'FF', 'FR', 'FS', 
                                           'FX', 'FY', 'GE', 'GF', 
                                           'GH', 'GO', 'GP', 'GQ', 
                                           'L2', 'LG', 'LH', 'LI', 
                                           'PN', 'PY', 'T3', 'V3', 
                                           'W3', 'WF', 'Z4', 'Z6', 
                                           'V6', 'V2', 'BB', 'V1', 
                                           'U1', 'E1', 'EE', 'E2', 
                                           'CY', 'DH', 'B9', 'AC', 
                                           'AF', 'AM', 'DP', 'ET', 
                                           'FM', 'FO', 'V9', 'W6', 
                                           'WI', 'WW', 'DQ', 'DR', 
                                           'EF', 'EG', 'EU', 'FN', 
                                           'G4', 'S2', 'K3', 'DB', 
                                           'CL', 'CT', 'CO', 'CV', 
                                           'FL', 'FW', 'FJ', 'FL', 
                                           'CO', 'KF', 'KG', 'V5', 
                                           'DJ', 'DK', 'A2', 'GI', 
                                           'FT', 'FH', 'GJ', 'AD', 
                                           'CM', 'AG', 'CP', 'CQ', 
                                           'CW', 'KH', 'DA', 'KI', 
                                           'KK', 'KJ', 'CE', 'CF', 
                                           'KA', 'KB', 'KC', 'CG', 
                                           'CZ', 'DG', 'KE', 'DI', 
                                           'R3', 'DN', 'BS', 'BM', 
                                           'BH', 'BN', 'BI', 'BO', 
                                           'BP', 'BJ', 'BK', 'BQ', 
                                           'BL', 'BR', 'DN', 'R1', 
                                           'R2', 'R4', 'R5', 'CK', 
                                           'CS', 'CN', 'CU', 'Q2', 
                                           'CR', 'EQ', 'DL', 'X1', 
                                           'V8', 'V0', 'W9', 'W0', 
                                           'Z1', 'Z2', 'AN', 'DD', 
                                           'AP', 'AR', 'DE', 'BC', 
                                           'BD', 'BE', 'B1', 'AZ', 
                                           'CA', 'B9', 'FT', 'FH', 
                                           'GJ', 'X6', 'Z3', 'B8', 
                                           'DY', 'B2', 'GK', 'D2', 
                                           'LM', 'Z7', 'ES', 'GL' ) THEN 1 
               ELSE 0 
             END 
	INTO   RT_FALVAREZ.DBO.exc1omdm -- select * -- select distinct C_OBSERVACION
	FROM   [lnkmaca].[ods].[dbo].[fact_excepciones]  where tipoexcepcion='A' and t_llamada='P2' 
	and   solicitud IN (SELECT solicitud 
                     FROM   RT_FALVAREZ.DBO.xxomdm) 
                     
                     
/*
p1 negociador simulaciiones  
p2 solictud 
*/
	 drop table RT_FALVAREZ.DBO.exc2omdm
	SELECT solicitud, 
       Max(jovenprof)                  RC_JovenProf, 
       Max(antiguedad_lab)             RC_antiguedad_lab, 
       Max(aval)                       RC_aval, 
       Max(boletin)                    RC_boletin, 
       Max(bureau)                     RC_bureau, 
       Max(conyuge)                    RC_conyuge, 
       Max(cta_cerrada)                RC_cta_cerrada, 
       Max(dda_declar_noactualiz)      RC_dda_declar_noActualiz, 
       Max(edad)                       RC_edad, 
       Max(empleador)                  RC_empleador, 
       Max(est_financ_noactualiz)      RC_est_financ_noActualiz, 
       Max(funcionario)                RC_Funcionario, 
       Max(lnh)                        RC_LNH, 
       Max(excede_tope_tarjeta)        RC_excede_tope_tarjeta, 
       Max(monto)                      RC_MONTO, 
       Max(patrimonio)                 RC_Patrimonio, 
       Max(predictor)                  RC_Predictor, 
       Max(renta)                      RC_renta, 
       Max(tdsr)                       RC_TDSR, 
       Max(residencia)                 RC_Residencia, 
       Max(mora_comercio)              RC_Mora_comercio, 
       Max(socio_empresa)              RC_Socio_Empresa, 
       Max(tipo_cliente)               RC_Tipo_Cliente, 
       Max(numero_acreedores)          RC_Numero_Acreedores, 
       Max(protesto)                   RC_Protesto, 
       Max(rechazo_sic)                RC_Rechazo_SIC, 
       Max(margen)                     RC_Margen, 
       Max(impedido)                   RC_Impedido, 
       Max(fallecido)                  RC_fallecido, 
       Max(renegociado)                RC_renegociado, 
       Max(sin_coneccion)              RC_Sin_coneccion, 
       Max(rut_no_registrado_sinacofi) RC_rut_no_registrado_sinacofi, 
       Max(n_consultas_rut)            RC_n_consultas_rut, 
       Max(sin_inf_score_sinacofi)     RC_Sin_inf_score_Sinacofi, 
       Max(tb)                         RC_tb, 
       Max(otro)                       RC_otro, 
       rc_conteo = Count(*) 
	INTO   RT_FALVAREZ.DBO.exc2omdm 
	FROM   RT_FALVAREZ.DBO.exc1omdm
	GROUP  BY solicitud 


	-- q13
	ALTER TABLE RT_FALVAREZ.DBO.xxomdm 
	ADD rc_antiguedad_lab INT, rc_jovenprof INT, rc_aval INT, rc_boletin INT, 
	rc_bureau INT, rc_conyuge INT, rc_cta_cerrada INT, rc_dda_declar_noactualiz 
	INT, rc_edad INT, rc_empleador INT, rc_est_financ_noactualiz INT, 
	rc_funcionario INT, rc_lnh INT, rc_excede_tope_tarjeta INT, rc_monto INT, 
	rc_patrimonio INT, rc_predictor INT, rc_renta INT, rc_tdsr INT, rc_residencia 
	INT, rc_mora_comercio INT, rc_socio_empresa INT, rc_tipo_cliente INT, 
	rc_numero_acreedores INT, rc_protesto INT, rc_rechazo_sic INT, rc_margen INT, 
	rc_impedido INT, rc_tb INT, rc_fallecido INT, rc_renegociado INT, 
	rc_sin_coneccion INT, rc_rut_no_registrado_sinacofi INT, rc_n_consultas_rut 
	INT, rc_sin_inf_score_sinacofi INT, rc_otro INT, rc_conteo INT 

	-- q14
	UPDATE RT_FALVAREZ.DBO.xxomdm 
	SET    rc_antiguedad_lab = '0', 
       rc_aval = '0', 
       rc_boletin = '0', 
       rc_bureau = '0', 
       rc_conyuge = '0', 
       rc_cta_cerrada = '0', 
       rc_dda_declar_noactualiz = '0', 
       rc_edad = '0', 
       rc_empleador = '0', 
       rc_est_financ_noactualiz = '0', 
       rc_funcionario = '0', 
       rc_lnh = '0', 
       rc_monto = '0', 
       rc_patrimonio = '0', 
       rc_predictor = '0', 
       rc_renta = '0', 
       rc_tdsr = '0', 
       rc_residencia = '0', 
       rc_mora_comercio = '0', 
       rc_socio_empresa = '0', 
       rc_tipo_cliente = '0', 
       rc_numero_acreedores = '0', 
       rc_protesto = '0', 
       rc_rechazo_sic = '0', 
       rc_margen = '0', 
       rc_impedido = '0', 
       rc_fallecido = '0', 
       rc_renegociado = '0', 
       rc_sin_coneccion = '0', 
       rc_rut_no_registrado_sinacofi = '0', 
       rc_n_consultas_rut = '0', 
       rc_sin_inf_score_sinacofi = '0', 
       rc_otro = '0', 
       rc_tb = '0', 
       rc_jovenprof = '0', 
       rc_excede_tope_tarjeta = '0' 

	-- q15
	UPDATE tabla 
	SET    rc_antiguedad_lab = a.rc_antiguedad_lab, 
       rc_aval = a.rc_aval, 
       rc_boletin = a.rc_boletin, 
       rc_bureau = a.rc_bureau, 
       rc_conyuge = a.rc_conyuge, 
       rc_cta_cerrada = a.rc_cta_cerrada, 
       rc_dda_declar_noactualiz = a.rc_dda_declar_noactualiz, 
       rc_edad = a.rc_edad, 
       rc_empleador = a.rc_empleador, 
       rc_est_financ_noactualiz = a.rc_est_financ_noactualiz, 
       rc_funcionario = a.rc_funcionario, 
       rc_lnh = a.rc_lnh, 
       rc_monto = a.rc_monto, 
       rc_patrimonio = a.rc_patrimonio, 
       rc_predictor = a.rc_predictor, 
       rc_renta = a.rc_renta, 
       rc_tdsr = a.rc_tdsr, 
       rc_residencia = a.rc_residencia, 
       rc_mora_comercio = a.rc_mora_comercio, 
       rc_socio_empresa = a.rc_socio_empresa, 
       rc_tipo_cliente = a.rc_tipo_cliente, 
       rc_numero_acreedores = a.rc_numero_acreedores, 
       rc_protesto = a.rc_protesto, 
       rc_rechazo_sic = a.rc_rechazo_sic, 
       rc_margen = a.rc_margen, 
       rc_impedido = a.rc_impedido, 
       rc_fallecido = a.rc_fallecido, 
       rc_renegociado = a.rc_renegociado, 
       rc_sin_coneccion = a.rc_sin_coneccion, 
       rc_rut_no_registrado_sinacofi = a.rc_rut_no_registrado_sinacofi, 
       rc_n_consultas_rut = a.rc_n_consultas_rut, 
       rc_sin_inf_score_sinacofi = a.rc_sin_inf_score_sinacofi, 
       rc_tb = a.rc_tb, 
       rc_jovenprof = a.rc_jovenprof, 
       rc_excede_tope_tarjeta = a.rc_excede_tope_tarjeta, 
       rc_conteo = a.rc_conteo 
	FROM   RT_FALVAREZ.DBO.xxomdm tabla, 
       RT_FALVAREZ.DBO.exc2omdm a 
	WHERE  tabla.id_solicitud = a.solicitud 

	DROP TABLE RT_FALVAREZ.DBO.sw_solsomdm
	SELECT * 
	INTO   RT_FALVAREZ.DBO.sw_solsomdm
	FROM   RT_FALVAREZ.DBO.xxomdm 
	
	drop table RT_FALVAREZ.DBO.sw_sols_resomdm
	select * 
	INTO   RT_FALVAREZ.DBO.sw_sols_resomdm -- select  *
	FROM   RT_FALVAREZ.DBO.sw_solsomdm

	
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm add BHV_CLI_ANT int
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm add monto_castigos int
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm add RI CHAR(1)
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD CATEG_SINACOFI CHAR(10) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD CATEG_SCORE_INT CHAR(10) 
	
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_SEG_0 CHAR(2) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_SEG_0_label CHAR(15) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_CF_BHV_consumo CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_CF_BHV_hipotecario CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_PB_BHV_consumo CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_PB_BHV_hipotecario CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_PB_BHV_revolving CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_LABEL_BHV_MIN CHAR(50) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm ADD mod_tot_MAX_DIA_MORA FLOAT 
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_resomdm 
	SET BHV_CLI_ANT=B.f_ff_f_p , mod_tot_SEG_0=seg_0 ,mod_tot_SEG_0_label=SEG_0_label, mod_tot_CF_BHV_consumo=f_cc_f_p,
	 mod_tot_CF_BHV_hipotecario=f_ch_f_p ,mod_tot_PB_BHV_consumo=f_pc_f_p , mod_tot_PB_BHV_hipotecario=f_ph_f_p 
	 ,mod_tot_PB_BHV_revolving=f_pr_f_p ,mod_tot_LABEL_BHV_MIN =f_ff_f_s -- SELECT *  -- SELECT COUNT(1)
	FROM RT_FALVAREZ.DBO.sw_sols_resomdm A INNER JOIN [RT_SCORING].dbo.mod_tot_his_min B
	ON A.RUT_NUM=B.RUT AND A.CAMADA=B.PERIODO
	
	DROP TABLE RT_FALVAREZ.DBO.castigos_totales 
	SELECT CAS_SUM=SUM(CAS_MONTOCASTIGO+CAS_MONTO_ADJ_GAAP_CAN),
	CAS_COUNT=COUNT(1),
	LTRIM(RTRIM(CAS_RUT)) AS RUT, 
	CAS_PERIODO AS PERIODO 
	into RT_FALVAREZ.DBO.castigos_totales 
	FROM RT_OPERACIONES.DATARISK.CASTIGOS
	GROUP BY CAS_RUT,CAS_PERIODO
	ORDER BY CAS_RUT,PERIODO

	drop table RT_FALVAREZ.DBO.Castigos2
	select  b.*, camada, 
	cast((left(b.periodo,4)+'-'+right(b.periodo,2)+'-01 00:00:00.000') as datetime) as periodo_fch ,
    cast((left(a.camada,4)+'-'+right(a.camada,2)+'-01 00:00:00.000') as datetime) as camada_fch 
    into RT_FALVAREZ.DBO.Castigos2
    from RT_FALVAREZ.DBO.sw_sols_resomdm  a inner join RT_FALVAREZ.DBO.castigos_totales  b
	on A.RUT_NUM=B.RUT 
	
	drop table RT_FALVAREZ.DBO.Castigos3
	select distinct rut, cas_sum
	into RT_FALVAREZ.DBO.Castigos3 -- select * 
	from RT_FALVAREZ.DBO.Castigos2 
	where periodo_fch between dateadd(month,+1, camada_fch ) and dateadd(month,+12, camada_fch )

	--  select * from #castigos3
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_resomdm
	SET monto_castigos= case when b.cas_sum is null then 0 else b.cas_sum  end  -- SELECT *  -- SELECT COUNT(1)
	FROM RT_FALVAREZ.DBO.sw_sols_resomdm A INNER JOIN RT_FALVAREZ.DBO.Castigos3 B
	ON A.RUT_NUM=B.RUT 

	--alter table RT_FALVAREZ.DBO.sw_sols_res_paso drop column  ScoreSinacofiCliente 
	alter table  RT_FALVAREZ.DBO.sw_sols_resomdm add  ScoreSinacofiCliente  int 
	-- select * from #sw_sols_res

	update RT_FALVAREZ.DBO.sw_sols_resomdm
	set ScoreSinacofiCliente = convert(int,score_sinacofi)

	update RT_FALVAREZ.DBO.sw_sols_resomdm
	set score_sinacofi = convert(int,score_sinacofi)


	/* PEGAMOS MAXIMAS MORAS */

	drop table RT_FALVAREZ.DBO.sw_sols_resomdm2
	select a.*, moramax12m 
	into RT_FALVAREZ.DBO.sw_sols_resomdm2 -- select * 
	from RT_FALVAREZ.DBO.sw_sols_resomdm a left join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and a.camada=b.periodo 
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set moramax12m = b.moramax12m -- select convert(varchar(6),dateadd(month,+1, left(a.camada,6)+'01'),112),* 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a inner join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and convert(varchar(6),dateadd(month,+1, left(a.camada,6)+'01'),112)=b.periodo 
	where a.moramax12m is null 
	
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm2 add tipo_producto nvarchar(20)
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set tipo_producto='consumo'
	
	/****************************/
	/* Actualizo consumo monto **/
	/****************************/
	
	drop table RT_FALVAREZ.DBO.montos_consumo
	select rut,camada as periodo, max(convert(int,m_solicitadocs)) as monto
	into RT_FALVAREZ.DBO.montos_consumo -- select *
	from RT_FALVAREZ.DBO.sw_sols_resomdm where m_solicitadocs>0
	group by rut, camada
	order by rut, camada

	update A
	set a.m_solicitadoconsumo = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a 
	inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=b.periodo
	where tipo_producto='consumo' 
	
	update A
	set a.m_solicitadoconsumo = b.monto -- select m_solicitadocs,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , -1,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_solicitadocs <1
	
	update A
	set a.m_solicitadoconsumo = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , +1,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_solicitadocs <1
	
	update A
	set a.m_solicitadoconsumo = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , -2,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_solicitadocs <1
	
	update A
	set a.m_solicitadoconsumo = b.monto -- select * 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , +2,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_solicitadocs <1
	
	
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm2 add resultado_sw nvarchar(max) null

	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set resultado_sw = case when decisionsolicitud='AC' then 'Aceptado' 
			when decisionsolicitud='DL' then 'Rechazado'
			when decisionsolicitud='VF' then 'Devuelta'
			when decisionsolicitud='IN' then 'Zona de Analisis' else null end   
	from RT_FALVAREZ.DBO.sw_sols_resomdm2
	---- 
	-- alter table RT_FALVAREZ.DBO.sw_sols_resomdm2 add [ri] nvarchar(max) null
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set ri =CASE WHEN risk_indicator=1 THEN 'A' 
			WHEN risk_indicator=2 THEN 'B' 
			WHEN risk_indicator=3 THEN 'C' 
			WHEN risk_indicator=4 THEN 'D' 
			WHEN risk_indicator=5 THEN 'E' ELSE null END 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2	
		
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set ultimo_estado_solicitud =case when ultimo_estado_solicitud IS null  
									  then 'Otro' else ultimo_estado_solicitud  end 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2	
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set rut=convert(int, RUT)
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set m_solicitadoconsumo = case when m_solicitadoconsumo IS null then 0 else m_solicitadoconsumo/1000 end
	
		update RT_FALVAREZ.DBO.sw_sols_resomdm2
	set m_renta_titular = case when m_renta_titular IS null then 0 else m_renta_titular/1000 end
	
		delete  RT_FALVAREZ.DBO.sw_sols_resomdm2
	where m_renta_titular like '%*%'
	
	/* ORIGEN BASE */

	alter table RT_FALVAREZ.DBO.sw_sols_resomdm2 add baseorigen nvarchar(max) null
	
	update RT_FALVAREZ.DBO.sw_sols_resomdm2 
	set baseorigen=	'omdm' 
	
	/*COMPLEMENTA RENTA  */	

	drop table RT_FALVAREZ.DBO.complementarenta			
	select distinct Rut_num as rut ,id_solicitud as solicitud,case when round(rentacomplementada,0)<>m_renta_titular THEN 'SI' ELSE 'NO' END 
	AS complementa_renta,rentacomplementada,m_renta_titular
	into RT_FALVAREZ.DBO.complementarenta
	from RT_FALVAREZ.DBO.paso1omdm 

	alter table RT_FALVAREZ.DBO.sw_sols_resomdm2  add complementa_renta nvarchar(max) null

	update  RT_FALVAREZ.DBO.sw_sols_resomdm2  
	set complementa_renta='SI' -- SELECT * 
	from  RT_FALVAREZ.DBO.sw_sols_resomdm2  a inner join (SELECT RUT, SOLICITUD FROM RT_FALVAREZ.DBO.complementarenta	 WHERE complementa_renta='SI' ) B
	ON A.Rut_num=B.RUT AND A.id_solicitud=B.SOLICITUD
		
	update RT_FALVAREZ.DBO.sw_sols_resomdm2  
	set complementa_renta='NO'	
	WHERE complementa_renta IS NULL 
	
	---*--- 
	---*--- 
	
	/* aprobados por excepcion */
	
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_resomdm2  add APROBADO_EXCEPCION nvarchar(max) null
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_resomdm2 
	SET APROBADO_EXCEPCION=CASE WHEN (Resultado_SW='Rechazado' and ULTIMO_ESTADO_SOLICITUD='Aprobado Cursado') then 
	'SI' else 'NO' end  -- SELECT * 
	from RT_FALVAREZ.DBO.sw_sols_resomdm2 	
			
	alter table RT_FALVAREZ.DBO.sw_sols_resomdm2 add [antiguedad] int null
      ,[n_oper]  int null
      ,[n_oper_mif]  int null
      ,[con_covd_cotc]  int null
      ,[n_oper_consumo]  int null
      ,[n_oper_consumo_vivienda]  int null
      ,[n_oper_seg_personas] int null
      ,[n_oper_seg_consumo]  int null
      ,[n_oper_seg_micro] int null
      ,[n_mora30_12m] int null
      ,[n_mora60_12m] int null
      ,[n_mora90_12m] int null
      ,[n_covdtc12m] int null
      ,[n_enOp]  int null
      ,[moramax24m] int null
      ,[ncastigos12m] int null
      ,[nreneg12m] int null
									   
			
	update RT_FALVAREZ.DBO.sw_sols_resomdm2 
	set [antiguedad]=b.[antiguedad]
      ,[n_oper]=b.[n_oper]
      ,[n_oper_mif]=b.[n_oper_mif]
      ,[con_covd_cotc]=b.[con_covd_cotc]
      ,[n_oper_consumo]=b.[n_oper_consumo]
      ,[n_oper_consumo_vivienda]=b.[n_oper_consumo_vivienda]
      ,[n_oper_seg_personas]=b.[n_oper_seg_personas]
      ,[n_oper_seg_consumo]=b.[n_oper_seg_consumo]
      ,[n_oper_seg_micro]=b.[n_oper_seg_micro]
      ,[n_mora30_12m] =b.[n_mora30_12m]
      ,[n_mora60_12m] =b.[n_mora60_12m] 
      ,[n_mora90_12m] =b.[n_mora90_12m] 
      ,[n_covdtc12m] =b.[n_covdtc12m] 
      ,[n_enOp]  =b.[n_enOp]  
      ,[moramax24m] =b.[moramax24m] 
      ,[ncastigos12m]=b.[ncastigos12m]
      ,[nreneg12m]=b.[nreneg12m] -- select * 
    from RT_FALVAREZ.DBO.sw_sols_resomdm2  a inner join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and a.camada=b.periodo 
	
			delete  RT_FALVAREZ.DBO.sw_sols_resomdm2
	where m_renta_titular like '%*%'
	
	
	--- fin CODIGO  --
	
	delete RT_SCORING.dbo.sw_sols_res -- select *   from RT_SCORING.dbo.sw_sols_res
	where segmento_label in ('Personal Banking','Consumer Finance')
	and tipo_producto='consumo' and baseorigen='omdm'  and camada = @periodo_inicial 
	
	
	insert into RT_SCORING.dbo.sw_sols_res 
	SELECT  [segmento_label]
      ,[rut_num]
      ,[desc_path_sw]
      ,[resultado_sw]
      ,[fh_proceso]
      ,[score_sinacofi]
      ,[ValorUF] as [valoruf_sw]
      ,[score_interno]
      ,[score_bhv]
      ,CASE WHEN risk_indicator=1 THEN 'A' 
			WHEN risk_indicator=2 THEN 'B' 
			WHEN risk_indicator=3 THEN 'C' 
			WHEN risk_indicator=4 THEN 'D' 
			WHEN risk_indicator=5 THEN 'E' ELSE null END as [risk_indicator]
      ,[camada]
      ,[num_solicitudes]
      ,[ultimo_estado_solicitud]
      ,[fecha]
      ,0 as [sistema]
      ,id_solicitud as [solicitud]
      ,[c_sis]
      ,[TipoVivienda_Titular] as [tipoviviendacliente]
      ,[Edad_Titular] as [edad]
      ,[EstadoCivil_Titular] as [estadocivil]
      ,isnull([NivelEducacional_Titular],0) as [niveleducacion]
      ,'' as [n_antiguedadlaboral]
      ,[RiesgoActividad_Profesion_Titular] as [riesgoprofesion]
      ,'' as [tipocargoactual]
      ,[TipoDependencia_Titular] as [tipodependecia]
      ,[PoseeBienRaiz_Titular] as [b_bienraiz]
      ,case when [EnConvenio_Titular]='N' then 9999999 else 1111111  end as [conveniocliente]
      ,[rut]
      ,'' as [dicomcliente]
      ,[n_Acreedores_Titular] as [n_acreedores]
      ,tipovariabilidad_titular as [tiporentadeclarada]
      ,isnull(rentauf_solicitud,0) as [m_rentaufcalculada]
      ,isnull(tdsracreditado_titular,0) as [m_tdsracreditado]
      ,isnull([LeverageProyectado_Titular],0) as [m_leveragenohipotecarioproyectado]
      ,isnull(tdsrproyectado_solicitud,0) as [m_tdsr_clienteproyectado]
      ,isnull([m_Renta_Titular],0) as [m_rentaclientepesos]
      ,'' as [v_wrfprotestoscliente]
      ,isnull([antiguedad],0) as [antiguedadcliente]
      ,'' as [t_acc]
      ,[TipoCliente_Titular] as [tipocliente]
      ,[ClienteNuevo_Titular] as [b_clientenuevo]
      ,case when id_campana_titular='' then 'N' else 'S' end  as [b_campaña]
      ,[CumpleReglasCampana_Titular] as [b_cumplebases]
      ,[ValorUF] as [m_valoruf]
      ,[ProductoConsumo] as [producto]
      ,n_cuotasconsumo as [plazo]
      ,m_solicitadototal/1000  as [m_montosolicitado]
      ,'' as [canal]
      ,case when (m_solicitadotarjeta1 is null) then 0 
			when (m_solicitadotarjeta1 is not null and m_solicitadotarjeta2 IS null) then 1
			when (m_solicitadotarjeta1 is not null and m_solicitadotarjeta2 IS not null and m_solicitadotarjeta3 IS null) then 2
			when (m_solicitadotarjeta1 is not null and m_solicitadotarjeta2 IS not null and m_solicitadotarjeta3 IS not null) then 3
			else 0   end as [n_tarjeta]
      ,case when m_solicitadotarjeta1 IS null then 'N' else 'S' end  as [b_tarjeta]
      ,case when m_solicitadoLC >0 then 'S' else 'N' end  as [b_linea]
      ,case when m_solicitadoLC >0 then 'S' else 'N' end  as [b_cuenta]
      ,'' as [tarjeta1]
      ,'' as [tarjeta2]
      ,'' as [tarjeta3]
      ,(isnull([m_SolicitadoUSTarjeta1],0)+isnull([m_SolicitadoUSTarjeta2],0)+isnull([m_SolicitadoUSTarjeta3],0)) as [m_monedaextranjera]
      ,[m_SolicitadoLC]/1000 as [m_montolinea]
      ,(isnull([m_SolicitadoTarjeta1],0)+isnull([m_SolicitadoTarjeta2],0)+isnull([m_SolicitadoTarjeta3],0))/1000 as [m_montotarjeta]
      ,'' as [b_margenvigente]
      ,[n_Acreedores_Titular] as [n_acreedorescliente]
      ,'' as [productocuenta]
      ,[productolinea]
      ,[ProductoTarjeta1] as [productotarjeta]
      ,[productoconsumo]
      ,'' as [disponibilidadlineau6m]
      ,'' as [antiguedaddeudasbifm6m]
      ,'' as [b_lineadisponible]
      ,'' as [b_moracomercio]
      ,'' as [b_bureau]
      ,case when [Funcionario_Titular]='false' then 'N' else 'S'end   as [b_funcionario]
      ,[Jubilado_Titular] as [b_jubilado]
      ,'' as [b_consumoaldia]
      ,''  as [b_clientecompracartera]
      ,[tot_Activos_Titular]/1000  as [m_totalactivo]
      ,[tot_Pasivos_Titular]/1000 as [m_totalpasivo]
      ,'' as [m_deudaconsumonossa]
      ,'' as [m_deudaconsumototalssa]
      ,'' as [m_deudaconsumocuotassa]
      ,'' as [m_cupotajetassa]
      ,'' as [m_cupolineassa]
      ,'' as [m_prepagoconsumossa]
      ,'' as [m_prepagolineassa]
      ,'' as [m_prepagotarjetassa]
      ,'' as [m_prepagoconsumosbif]
      ,  CASE  WHEN Desc_Path_SW = 'Nuevo' then 110 
                     WHEN Desc_Path_SW = 'Antiguo Campana' then 120 
                      WHEN Desc_Path_SW ='Renegociado' then 130 
                       WHEN Desc_Path_SW = 'Moroso' then 140 
                        WHEN Desc_Path_SW =  'Antiguo' then 150 
                         WHEN Desc_Path_SW = 'Nuevo Campana' then 160 else 999 end as [n_logicpath]
      ,ri as [n_risk]
      ,'' as [n_puntajesw]
      ,[Fecha_Proceso] as [f_procesosw]
      ,'' as [h_procesosw]
      ,'' as [m_margenconsumo]
      ,'' as [m_margenconsumodisponible]
      ,'' as [m_margenlinea]
      ,'' as [m_margenlineadisponible]
      ,'' as [m_margentarjeta]
      ,'' as [m_margentarjetadisponible]
      ,'' as [m_margentotal]
      ,'' as [m_margentotaldisponible]
      ,decisionsolicitud as [decision_group]
      ,[b_ultimocambio]
      ,'' as [filler_3]
      ,'' as [v_predictorpublicequifax]
      ,'' as [segmento]
      ,[b_Sbif_U12M_Titular] as [b_sbif_u12m]
      ,'' as [promedioconsumou6m_u12m]
      ,'' as [scoresinacoficlientenobca]
      ,'' as [v_pje_mod_01]
      ,'' as [sc_bienraiz]
      ,'' as [sc_displinea]
      ,'' as [sc_promcons]
      ,'' as [sc_nacre]
      ,'' as [sc_edad]
      ,'' as [sc_ant_lab]
      ,'' as [sc_ddassa]
      ,'' as [sc_bbureau]
      ,'' as [sc_tdsr_acred]
      ,isnull([score_interno],0) as [score_cf_ori_calculado]
      ,'' as [sc2_disponibilidadlineau6m]
      ,'' as [sc2_m_deudaconsumonossa_renta]
      ,'' as [sc2_promedioconsumou6m_u12m]
      ,'' as [sc2_b_bienraiz]
      ,'' as [sc2_b_bureau]
      ,'' as [sc2_niveleducacion]
      ,'' as [sc2_tipodependecia]
      ,'' as [sc2_n_antiguedadlaboral]
      ,'' as [sc2_edad]
      ,isnull([score_interno],0) as [score_pb_ori_calculado]
      ,[bhv_cli_ant]
      ,[monto_castigos]
      ,[ri]
      ,[categ_sinacofi]
      ,[categ_score_int]
      ,[mod_tot_seg_0]
      ,[mod_tot_seg_0_label]
      ,[mod_tot_cf_bhv_consumo]
      ,[mod_tot_cf_bhv_hipotecario]
      ,[mod_tot_pb_bhv_consumo]
      ,[mod_tot_pb_bhv_hipotecario]
      ,[mod_tot_pb_bhv_revolving]
      ,[mod_tot_label_bhv_min]
      ,[mod_tot_max_dia_mora]
      ,[scoresinacoficliente]
      ,[tipo_producto]
      ,[baseorigen]
      ,complementa_renta
      ,APROBADO_EXCEPCION
      ,[antiguedad] 
      ,[n_oper]  
      ,[n_oper_mif]  
      ,[con_covd_cotc]  
      ,[n_oper_consumo]  
      ,[n_oper_consumo_vivienda]  
      ,[n_oper_seg_personas] 
      ,[n_oper_seg_consumo]  
      ,[n_oper_seg_micro] 
      ,[n_mora30_12m] 
      ,[n_mora60_12m] 
      ,[n_mora90_12m] 
      ,[n_covdtc12m] 
      ,[n_enOp]  
      ,[moramax24m] 
      ,[ncastigos12m] 
      ,[nreneg12m] 
       ,rc_antiguedad_lab 
       ,rc_jovenprof 
       ,rc_aval 
       ,rc_boletin 
       ,rc_bureau 
       ,rc_conyuge 
       ,rc_cta_cerrada 
       ,rc_dda_declar_noactualiz  
       ,rc_edad 
       ,rc_empleador 
       ,rc_est_financ_noactualiz 
       ,rc_funcionario 
       ,rc_lnh 
       ,rc_excede_tope_tarjeta 
       ,rc_monto 
       ,rc_patrimonio 
       ,rc_predictor 
       ,rc_renta 
       ,rc_tdsr 
       ,rc_residencia 
       ,rc_mora_comercio 
       ,rc_socio_empresa 
       ,rc_tipo_cliente 
       ,rc_numero_acreedores 
       ,rc_protesto 
       ,rc_rechazo_sic 
       ,rc_margen 
       ,rc_impedido 
       ,rc_tb 
       ,rc_fallecido 
       ,rc_renegociado 
       ,rc_sin_coneccion 
       ,rc_rut_no_registrado_sinacofi 
       ,rc_n_consultas_rut 
       ,rc_sin_inf_score_sinacofi 
       ,rc_otro 
       ,rc_conteo  -- select *
       ,LeverageActual_Solicitud
	   ,LeverageDisponible_Solicitud
	   ,LeverageEnTramite_Solicitud
	   ,LeverageExterno_Solicitud
	   ,LeverageLimite_Solicitud
	   ,LeverageProyectado_Solicitud
	   ,LeverageSolicitado_Solicitud  -- select *
	   ,'' as sw_m_LeverageActualLinea
		,'' as sw_m_LeverageSolicitadoLinea
		,'' as sw_m_LeverageProyectadoLinea
		,'' as sw_m_LeverageActualTarjeta
		,'' as sw_m_LeverageSolicitadoTarjeta
		,'' as sw_m_LeverageProyectadoTarjeta
		,'' as sw_m_LeverageActualConsumo
		,'' as sw_m_LeverageSolicitadoConsumo
		,'' as sw_m_LeverageProyectadoConsumo
		,'' as sw_m_LeverageActualComercial
  FROM RT_FALVAREZ.DBO.sw_sols_resomdm2
		
		-- select distinct m_renta_titular	FROM   RT_FALVAREZ.DBO.sw_sols_resomdm2 order by m_renta_titular where m_renta_titular>0
	
	/* FIN  ACTUALIZACION DE CONSUMOS  EN BASE PRINCIPAL */

----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------
end