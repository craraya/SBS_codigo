USE [RT_SCORING]
GO
/****** Object:  StoredProcedure [dbo].[sp_update_solres_consumos]    Script Date: 07/20/2016 12:14:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<FERNANDO ALVAREZ SERRANO>
-- Create date: <31-03-2016>
-- Description:	< SE OBTIENEN LAS SOLICITUDES DE CREDITOS DE CONSUMO >
-- =============================================
 
ALTER PROCEDURE [dbo].[sp_update_solres_consumos]  @periodo_inicial VARCHAR(6)
/*

declare @error varchar(max),@return_value  int
exec @RETURN_VALUE= [sp_update_solres_consumos]		@periodo_inicial='201605'
SELECT	@ERROR as N'@ERROR'
SELECT @RETURN_VALUE

*/


AS
BEGIN
		
	drop table RT_FALVAREZ.DBO.paso1
	SELECT *, 
       Segmento_label=CASE 
                        WHEN c_sis = 'CMP' THEN 'Consumer Finance' 
                        WHEN c_sis = 'MGN' THEN 'Personal Banking' 
                        WHEN c_sis = 'FST' THEN 'Personal Banking' 
                      END, 
       CONVERT(INT, [rut])       AS Rut_num, 
       Desc_Path_SW=CASE 
                      WHEN n_logicpath = '110' THEN 'Nuevo' 
                      WHEN n_logicpath = '210' THEN 'Nuevo' 
                      WHEN n_logicpath = '230' THEN 'Nuevo' 
                      WHEN n_logicpath = '120' THEN 'Antiguo Campana' 
                      WHEN n_logicpath = '121' THEN 'Antiguo Campana' 
                      WHEN n_logicpath = '130' THEN 'Renegociado' 
                      WHEN n_logicpath = '150' THEN 'Antiguo' 
                      WHEN n_logicpath = '160' THEN 'Nuevo Campana' 
                      WHEN n_logicpath = '170' THEN 'Campaña Renegociado' 
                      WHEN n_logicpath = '140' THEN 'Moroso' 
                      WHEN n_logicpath = '121' THEN 'Antiguo Campana' 
                      WHEN n_logicpath = '161' THEN 'Nuevo Campana' 
                      
                    END 
       , 
       Resultado_SW= CASE 
                       WHEN decision_group = 'AC' THEN 'Aceptado' 
                       WHEN decision_group = 'DL' THEN 'Rechazado' 
                       WHEN decision_group = 'RV' THEN 'Zona de Analisis' 
                       WHEN decision_group = 'VF' THEN 'Devuelta' 
                     END, 
       f_procesosw AS FH_Proceso,
       scoresinacoficliente      Score_Sinacofi, 
       ValorUF_SW=m_valoruf, 
       Score_Interno=n_puntajesw, 
       Score_BHV=v_pje_mod_01, 
       Risk_Indicator=n_risk
       , left(CONVERT(VARCHAR(6), fecha, 112),6) as camada 
	INTO   RT_FALVAREZ.DBO.paso1
	FROM    LNKMACA.[ods].[dbo].[fact_vectorstrategy] 
	WHERE c_sis IN ( 'FST', 'MGN', 'CMP' )
	AND CONVERT(VARCHAR(6), fecha, 112) = @periodo_inicial

	drop table  RT_FALVAREZ.DBO.maxima
	SELECT rut_num, 
       CONVERT(VARCHAR(6), fecha, 112) * 1 AS AnoMes, 
       Max(fh_proceso)                     FH_Proceso, 
       Count(1)                            Num_Solicitudes 
	INTO   RT_FALVAREZ.DBO.maxima
	FROM   RT_FALVAREZ.DBO.paso1  
	GROUP  BY rut_num,    CONVERT(VARCHAR(6), fecha, 112) * 1 

	drop table RT_FALVAREZ.DBO.xx 
	SELECT T1.*, 
       b.num_solicitudes 
	INTO   RT_FALVAREZ.DBO.xx 
	FROM   RT_FALVAREZ.DBO.paso1 T1 
       JOIN (SELECT rut_num    AS rut, 
                    fh_proceso AS fechra, 
                    num_solicitudes 
             FROM   RT_FALVAREZ.DBO.maxima) b 
         ON T1.fh_proceso = b.fechra 
            AND T1.rut_num = b.rut 
	
	drop table RT_FALVAREZ.DBO.zieloz1 
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
                               /* WHEN Estado='11' THEN 'Pre-Ratificada' 
                                WHEN Estado='12' THEN 'Expirado' 
                                WHEN Estado='13' THEN 'Ingresada Pack' 
                                WHEN Estado='14' THEN 'Retificada Pack' 
                                WHEN Estado='15' THEN 'Pre-Retificada Pack' 
                                */ 
                                 ELSE 'Otro' 
                               END, 
       solicitud, 
       rut, 
       operacion, 
       codigopack, 
       Mto_Solic_FactSolicitud= CONVERT(FLOAT, Replace(c_montosolicitado, ',', 
                                               '.')) 
	INTO RT_FALVAREZ.DBO.zieloz1 
	FROM   LNKMACA.[ods].[dbo].[fact_solicitud] 
	WHERE  b_ultimocambio = '2' 
    AND codigopack IN (SELECT solicitud FROM   RT_FALVAREZ.DBO.xx ) 

    drop table RT_FALVAREZ.DBO.zieloz2 
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
                               /* WHEN Estado='11' THEN 'Pre-Ratificada' 
                                WHEN Estado='12' THEN 'Expirado' 
                                WHEN Estado='13' THEN 'Ingresada Pack' 
                                WHEN Estado='14' THEN 'Retificada Pack' 
                                WHEN Estado='15' THEN 'Pre-Retificada Pack' 
                                */ 
                                 ELSE 'Otro' 
                               END, 
       solicitud, 
       rut, 
       operacion, 
       codigopack, 
       Mto_Solic_FactSolicitud= CONVERT(FLOAT, Replace(c_montosolicitado, ',', 
                                               '.')) 
	INTO RT_FALVAREZ.DBO.zieloz2 
	FROM   LNKMACA.[ods].[dbo].[fact_solicitud] 
	WHERE  b_ultimocambio = '2' 
    AND Solicitud IN (SELECT Solicitud FROM RT_FALVAREZ.DBO.xx)


	-- q07
	ALTER TABLE RT_FALVAREZ.DBO.xx 	ADD ultimo_estado_solicitud NVARCHAR(50) 

	drop table  RT_FALVAREZ.DBO.factsolic
	SELECT * 
	INTO   RT_FALVAREZ.DBO.factsolic 
	FROM   RT_FALVAREZ.DBO.zieloz2 
	UNION 
	SELECT * 
	FROM   RT_FALVAREZ.DBO.zieloz1 

	-- q09
	UPDATE tabla 
	SET    ultimo_estado_solicitud = a.ultimo_estado_solicitud 
	FROM   RT_FALVAREZ.DBO.xx tabla, 
    RT_FALVAREZ.DBO.factsolic a 
	WHERE  tabla.solicitud = a.solicitud 

	-- q10
	UPDATE tabla 
	SET    ultimo_estado_solicitud = a.ultimo_estado_solicitud 
	FROM   RT_FALVAREZ.DBO.xx tabla, 
       RT_FALVAREZ.DBO.factsolic a 
	WHERE  tabla.solicitud = a.codigopack 

	 drop table RT_FALVAREZ.DBO.exc1
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
	INTO   RT_FALVAREZ.DBO.exc1 -- select * -- select distinct C_OBSERVACION
	FROM   [lnkmaca].[ods].[dbo].[fact_excepciones] -- where tipoexcepcion='A' and t_llamada='P2' 
	WHERE  solicitud IN (SELECT solicitud 
                     FROM   RT_FALVAREZ.DBO.xx) 
                     
                     
/*
p1 negociador simulaciiones  
p2 solictud 
*/
	 drop table RT_FALVAREZ.DBO.exc2
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
	INTO   RT_FALVAREZ.DBO.exc2 
	FROM   RT_FALVAREZ.DBO.exc1
	GROUP  BY solicitud 


	-- q13
	ALTER TABLE RT_FALVAREZ.DBO.xx 
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
	UPDATE RT_FALVAREZ.DBO.xx 
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
	FROM   RT_FALVAREZ.DBO.xx tabla, 
       RT_FALVAREZ.DBO.exc2 a 
	WHERE  tabla.solicitud = a.solicitud 

	DROP TABLE RT_FALVAREZ.DBO.sw_sols_paso
	SELECT * 
	INTO   RT_FALVAREZ.DBO.sw_sols_paso
	FROM   RT_FALVAREZ.DBO.xx 
	
	drop table   RT_FALVAREZ.DBO.sw_sols_res_paso
	SELECT segmento_label, 
       rut_num, 
       desc_path_sw,
       resultado_sw, 
       fh_proceso, 
       score_sinacofi, 
       valoruf_sw, 
       score_interno, 
       score_bhv, 
       risk_indicator, 
       camada, 
       num_solicitudes, 
       ultimo_estado_solicitud, 

       rc_antiguedad_lab, 
       rc_jovenprof, 
       rc_aval, 
       rc_boletin, 
       rc_bureau, 
       rc_conyuge, 
       rc_cta_cerrada, 
       rc_dda_declar_noactualiz, 
       rc_edad, 
       rc_empleador, 
       rc_est_financ_noactualiz, 
       rc_funcionario, 
       rc_lnh, 
       rc_excede_tope_tarjeta, 
       rc_monto, 
       rc_patrimonio, 
       rc_predictor, 
       rc_renta, 
       rc_tdsr, 
       rc_residencia, 
       rc_mora_comercio, 
       rc_socio_empresa, 
       rc_tipo_cliente, 
       rc_numero_acreedores, 
       rc_protesto, 
       rc_rechazo_sic, 
       rc_margen, 
       rc_impedido, 
       rc_tb, 
       rc_fallecido, 
       rc_renegociado, 
       rc_sin_coneccion, 
       rc_rut_no_registrado_sinacofi, 
       rc_n_consultas_rut, 
       rc_sin_inf_score_sinacofi, 
       rc_otro, 
       rc_conteo, 

       fecha, 
       sistema, 
       solicitud, 
       c_sis, 
       tipoviviendacliente, 
       edad, 
       estadocivil, 
       niveleducacion, 
       n_antiguedadlaboral, 
       riesgoprofesion, 
       tipocargoactual, 
       tipodependecia, 
       b_bienraiz, 
       conveniocliente, 
       rut, 
       dicomcliente, 
       n_acreedores, 
       tiporentadeclarada, 
       m_rentaufcalculada, 
       m_tdsracreditado, 
       m_leveragenohipotecarioproyectado, 
       m_tdsr_clienteproyectado, 
       m_rentaclientepesos, 
       v_wrfprotestoscliente, 
       antiguedadcliente, 
       t_acc, 
       tipocliente, 
       b_clientenuevo, 
       b_campaña, 
       b_cumplebases, 
       m_valoruf, 
       producto, 
       plazo, 
       m_montosolicitado, 
       canal, 
       n_tarjeta, 
       b_tarjeta, 
       b_linea, 
       b_cuenta, 
       tarjeta1, 
       tarjeta2, 
       tarjeta3, 
       m_monedaextranjera, 
       m_montolinea, 
       m_montotarjeta, 
       b_margenvigente, 
       n_acreedorescliente, 
       productocuenta, 
       productolinea, 
       productotarjeta, 
       productoconsumo, 
       disponibilidadlineau6m, 
       antiguedaddeudasbifm6m, 
       b_lineadisponible, 
       b_moracomercio, 
       b_bureau, 
       b_funcionario, 
       b_jubilado, 
       b_consumoaldia, 
       b_clientecompracartera, 
       m_totalactivo, 
       m_totalpasivo, 
       m_deudaconsumonossa, 
       m_deudaconsumototalssa, 
       m_deudaconsumocuotassa, 
       m_cupotajetassa, 
       m_cupolineassa, 
       m_prepagoconsumossa, 
       m_prepagolineassa, 
       m_prepagotarjetassa, 
       m_prepagoconsumosbif, 
       n_logicpath, 
       n_risk, 
       n_puntajesw, 
       f_procesosw, 
       h_procesosw, 
       m_margenconsumo, 
       m_margenconsumodisponible, 
       m_margenlinea, 
       m_margenlineadisponible, 
       m_margentarjeta, 
       m_margentarjetadisponible, 
       m_margentotal, 
       m_margentotaldisponible,
       decision_group, 
       b_ultimocambio, 
       filler_3, 
       v_predictorpublicequifax, 
       segmento, 
       b_sbif_u12m, 
       promedioconsumou6m_u12m, 
       scoresinacoficliente, 
       scoresinacoficlientenobca, 
       v_pje_mod_01,
	   0 as sc_bienraiz,
	   0 as sc_displinea,
	   0 as sc_promcons,
   	   0 as sc_nacre,
	   0 as sc_edad,
	   0 as sc_ant_lab,
	   0 as sc_ddassa,
	   0 as sc_bbureau,
	   0 as sc_tdsr_acred,
	   0 as score_cf_ori_calculado,
	   0 as sc2_disponibilidadlineau6m,
       0 as sc2_m_deudaconsumonossa_renta,
       0 as sc2_promedioconsumou6m_u12m,
       0 as sc2_b_bienraiz,
       0 as sc2_b_bureau,
       0 as sc2_niveleducacion,
       0 as sc2_tipodependecia,
       0 as sc2_n_antiguedadlaboral,
       0 as sc2_edad,
	   0 as score_pb_ori_calculado	
	   , m_LeverageActualLinea,
m_LeverageSolicitadoLinea,
m_LeverageProyectadoLinea,
m_LeverageActualTarjeta,
m_LeverageSolicitadoTarjeta,
m_LeverageProyectadoTarjeta,
m_LeverageActualConsumo,
m_LeverageSolicitadoConsumo,
m_LeverageProyectadoConsumo,
m_LeverageActualComercial	
	INTO   RT_FALVAREZ.DBO.sw_sols_res_paso
	FROM   RT_FALVAREZ.DBO.sw_sols_paso 

	DROP TABLE RT_FALVAREZ.DBO.CODIGOS_RECHAZOS_ORIGEN
	select DISTINCT  Solicitud, c_observacion
	INTO RT_FALVAREZ.DBO.CODIGOS_RECHAZOS_ORIGEN
	From   RT_FALVAREZ.DBO.exc1

	-- q20
	update RT_FALVAREZ.DBO.sw_sols_res_paso set 
	sc_bienraiz=case
	when b_BienRaiz = 'N' then 47
	when b_BienRaiz = 'S' then 72
	else 55
	end,
	sc_displinea=case
	when DisponibilidadLineau6m = 'N' then 42
	when DisponibilidadLineau6m = 'S' then 72
	else 56
	end,
	sc_promcons=case
	when PromedioConsumoU6M_U12M = 0 then 104
	when PromedioConsumoU6M_U12M = 1 then 84
	when PromedioConsumoU6M_U12M >= 2 then 72
	else 97
	end,
	sc_nacre =case
	when n_AcreedoresCliente = 0 then 55
	when n_AcreedoresCliente = 1 then 75
	when n_AcreedoresCliente = 2 then 73
	when n_AcreedoresCliente >=3 then 72
	end,
	sc_edad = case
	when edad < 0 then 50
	when edad <= 359 then 35
	when edad <= 479 then 42
	when edad <= 600 then 51
	when edad <= 720 then 66
	when edad <= 960 then 72
	when edad > 960 then 50
	else 50
	end,
	sc_ant_lab = case
	when N_AntiguedadLaboral <= 18 then 53
	when N_AntiguedadLaboral <= 36 then 58
	when N_AntiguedadLaboral <= 84 then 63
	when N_AntiguedadLaboral <= 720 then 72
	when N_AntiguedadLaboral > 720 then 61
	end,
	sc_ddassa = case
	when m_DeudaConsumoTotalSSA = 0 then 46
	when m_DeudaConsumoTotalSSA > 0 then 72
	end,
	sc_bbureau = case
	when b_Bureau = 0 then 79
	when b_Bureau = 1 then 81
	when b_Bureau = 2 then 59
	when b_Bureau = 3 then 75
	when b_Bureau = 4 then 72
	end,
	sc_tdsr_acred = case
	when m_TDSRAcreditado < 20 then 78
	when m_TDSRAcreditado <= 100 then 72
	when m_TDSRAcreditado >100 then 74
	end

	-- q21
	update RT_FALVAREZ.DBO.sw_sols_res_paso
	set score_cf_ori_calculado = 
	sc_bienraiz + sc_displinea + sc_promcons +
	sc_nacre + sc_edad + sc_ant_lab + sc_ddassa +
	sc_bbureau + sc_tdsr_acred 


	update  RT_FALVAREZ.DBO.sw_sols_res_paso
	set m_rentaclientepesos=1 -- select * from sw_sols_res
	where m_rentaclientepesos=0

	-- q22
	update RT_FALVAREZ.DBO.sw_sols_res_paso set 
	sc2_disponibilidadlineau6m = case
	when disponibilidadlineau6m = 'N' then 33
	when disponibilidadlineau6m = 'S' then 69
	else 61
	end,
	sc2_m_deudaconsumonossa_renta = case
	when m_deudaconsumonossa/m_rentaclientepesos =  0 then 82
	when m_deudaconsumonossa/m_rentaclientepesos<= 3 then 91
	when m_deudaconsumonossa/m_rentaclientepesos > 3 then 69
	else 81
	end,
	sc2_promedioconsumou6m_u12m = case
	when promedioconsumou6m_u12m = 0 then 87
	when promedioconsumou6m_u12m = 1 then 77
	when promedioconsumou6m_u12m>= 2 then 69
	else 83
	end,
	sc2_b_bienraiz= case
	when b_bienraiz = 'N' then 53
	when b_bienraiz = 'S' then 69
	else 59
	end,
	sc2_b_bureau = case
	when b_bureau = 1 then 94
	when b_bureau = 2 then 53
	when b_bureau = 3 then 67
	when b_bureau = 4 then 69
	else 90
	end,
	sc2_niveleducacion = case
	when niveleducacion in (1,2,3,5,6,99) then 53
	when niveleducacion in (4, 7) then 69
	else 63
	end,
	sc2_tipodependecia = case
	when tipodependecia = 'D' then 89
    when tipodependecia = 'I' then 69
	else 87
	end,
	sc2_n_antiguedadlaboral = case
	when n_antiguedadlaboral > 0 and n_antiguedadlaboral >= 48 then 55
    when n_antiguedadlaboral <= 120 then 61
	when n_antiguedadlaboral <= 720 then 69
	else 59
	end,
	sc2_edad = case
	when edad > 0 and edad >= 324 then 46
    when edad <= 480 then 48
	when edad <= 600 then 51
	when edad <= 720 then 55
	when edad <= 960 then 69
	else 50
	end

	-- q23
	update RT_FALVAREZ.DBO.sw_sols_res_paso
	set score_pb_ori_calculado = 
	sc2_disponibilidadlineau6m + sc2_m_deudaconsumonossa_renta + 
	sc2_promedioconsumou6m_u12m + sc2_b_bienraiz + sc2_b_bureau +
	sc2_niveleducacion + sc2_tipodependecia + sc2_n_antiguedadlaboral + sc2_edad

	/* hasta  aqui es la original */
	
	-- SELECT * FROM #sw_sols_res
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso add BHV_CLI_ANT int
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso add monto_castigos int
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso add RI CHAR(1)
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD CATEG_SINACOFI CHAR(10) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD CATEG_SCORE_INT CHAR(10) 
	
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_SEG_0 CHAR(2) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_SEG_0_label CHAR(15) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_CF_BHV_consumo CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_CF_BHV_hipotecario CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_PB_BHV_consumo CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_PB_BHV_hipotecario CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_PB_BHV_revolving CHAR(5) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_LABEL_BHV_MIN CHAR(50) 
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso ADD mod_tot_MAX_DIA_MORA FLOAT 
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso 
	SET BHV_CLI_ANT=B.f_ff_f_p , mod_tot_SEG_0=seg_0 ,mod_tot_SEG_0_label=SEG_0_label, mod_tot_CF_BHV_consumo=f_cc_f_p,
	 mod_tot_CF_BHV_hipotecario=f_ch_f_p ,mod_tot_PB_BHV_consumo=f_pc_f_p , mod_tot_PB_BHV_hipotecario=f_ph_f_p 
	 ,mod_tot_PB_BHV_revolving=f_pr_f_p ,mod_tot_LABEL_BHV_MIN =f_ff_f_s -- SELECT *  -- SELECT COUNT(1)
	FROM RT_FALVAREZ.DBO.sw_sols_res_paso A INNER JOIN [RT_SCORING].dbo.mod_tot_his_min B
	ON A.RUT_NUM=B.RUT AND A.CAMADA=B.PERIODO
	
	DROP TABLE RT_FALVAREZ.DBO.castigos_totales 
	SELECT CAS_SUM=SUM(CAS_MONTOCASTIGO+CAS_MONTO_ADJ_GAAP_CAN),
	CAS_COUNT=COUNT(1),
	LTRIM(RTRIM(CAS_RUT)) AS RUT, 
	CAS_PERIODO AS PERIODO 
	into RT_FALVAREZ.DBO.castigos_totales -- select min(CAS_PERIODO)-- SELECT * 
	FROM RT_OPERACIONES.DATARISK.CASTIGOS
	GROUP BY CAS_RUT,CAS_PERIODO
	ORDER BY CAS_RUT,PERIODO
	-- SELECT * FROM #castigos_totales

	drop table RT_FALVAREZ.DBO.Castigos2
	select  b.*, camada, 
	cast((left(b.periodo,4)+'-'+right(b.periodo,2)+'-01 00:00:00.000') as datetime) as periodo_fch ,
    cast((left(a.camada,4)+'-'+right(a.camada,2)+'-01 00:00:00.000') as datetime) as camada_fch 
    into RT_FALVAREZ.DBO.Castigos2
    from RT_FALVAREZ.DBO.sw_sols_res_paso  a inner join RT_FALVAREZ.DBO.castigos_totales  b
	on A.RUT_NUM=B.RUT 
	
	drop table RT_FALVAREZ.DBO.Castigos3
	select distinct rut, cas_sum
	into RT_FALVAREZ.DBO.Castigos3 -- select * 
	from RT_FALVAREZ.DBO.Castigos2 
	where periodo_fch between dateadd(month,+1, camada_fch ) and dateadd(month,+12, camada_fch )

	--  select * from #castigos3
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
	SET monto_castigos= case when b.cas_sum is null then 0 else b.cas_sum  end  -- SELECT *  -- SELECT COUNT(1)
	FROM RT_FALVAREZ.DBO.sw_sols_res_paso A INNER JOIN RT_FALVAREZ.DBO.Castigos3 B
	ON A.RUT_NUM=B.RUT 

	alter table RT_FALVAREZ.DBO.sw_sols_res_paso drop column  ScoreSinacofiCliente 
	alter table  RT_FALVAREZ.DBO.sw_sols_res_paso add  ScoreSinacofiCliente  int 
	-- select * from #sw_sols_res

	update RT_FALVAREZ.DBO.sw_sols_res_paso
	set ScoreSinacofiCliente = convert(int,score_sinacofi)


	---------------------------
	/* 1  PB NUEVO CAMPAÑA */ 
	--------------------------- 
 
	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso 
	SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <516								THEN '000-515'
						  WHEN ScoreSinacofiCliente >=516 AND ScoreSinacofiCliente <685 THEN '516-684'
						  WHEN ScoreSinacofiCliente >=685 AND ScoreSinacofiCliente <771 THEN '685-770'
						  WHEN ScoreSinacofiCliente >=771 AND ScoreSinacofiCliente <804 THEN '771-803'
						  WHEN ScoreSinacofiCliente >=804								THEN '804-999' END 
	WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo Campana' 	  


 
	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
	SET CATEG_SCORE_INT= CASE WHEN SCORE_PB_ORI_CALCULADO <598									THEN '000-597'
						  WHEN SCORE_PB_ORI_CALCULADO >=598 AND SCORE_PB_ORI_CALCULADO <609 THEN '598-608'
						  WHEN SCORE_PB_ORI_CALCULADO >=609 AND SCORE_PB_ORI_CALCULADO <621 THEN '609-620'
						  WHEN SCORE_PB_ORI_CALCULADO >=621 AND SCORE_PB_ORI_CALCULADO <655 THEN '621-654'
						  WHEN SCORE_PB_ORI_CALCULADO >=655									THEN '655-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
	WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo Campana'
 

  
   -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI

	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
	SET RI= CASE			 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='621-654'  THEN 'B'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='000-597'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='598-608'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='609-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='000-597'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='598-608'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='609-620'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
						 WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo Campana'
 
 
 
 	 				 
--------------------------- 				  
 /* 2  PB NUEVO NO CAMPAÑA */ 	--	WHERE 	SEGMENTO_LABEL='PB' AND 	desc_path_sw='Nuevo'		  
--------------------------- 
		
		
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <516							    THEN '000-515'
						  WHEN ScoreSinacofiCliente >=516 AND ScoreSinacofiCliente <685 THEN '516-684'
						  WHEN ScoreSinacofiCliente >=685 AND ScoreSinacofiCliente <771 THEN '685-770'
						  WHEN ScoreSinacofiCliente >=771 AND ScoreSinacofiCliente <804 THEN '771-803'
						  WHEN ScoreSinacofiCliente >=804							    THEN '804-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo'	 -- AND CATEG_SINACOFI IS NULL ORDER BY 		ScoreSinacofiCliente	  




 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET CATEG_SCORE_INT= CASE WHEN SCORE_PB_ORI_CALCULADO <598									THEN '000-597'
						  WHEN SCORE_PB_ORI_CALCULADO >=598 AND SCORE_PB_ORI_CALCULADO <609 THEN '598-608'
						  WHEN SCORE_PB_ORI_CALCULADO >=609 AND SCORE_PB_ORI_CALCULADO <621 THEN '609-620'
						  WHEN SCORE_PB_ORI_CALCULADO >=621 AND SCORE_PB_ORI_CALCULADO <655 THEN '621-654'
						  WHEN SCORE_PB_ORI_CALCULADO >=655									THEN '655-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo'	


  

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET RI= CASE			 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
						 WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='621-654'  THEN 'B'
						 WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='000-597'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='598-608'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='609-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
						 WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='000-597'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='598-608'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='609-620'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
						 WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
						WHERE 	SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw='Nuevo'	

--------------------------- 
 /* 3  PB  ANTIGUO */ 	--	WHERE 	SEGMENTO_LABEL='PB' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')
--------------------------- 

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <381 THEN '000-380'
						  WHEN ScoreSinacofiCliente >=381 AND ScoreSinacofiCliente <785 THEN '381-784'
						  WHEN ScoreSinacofiCliente >=785 AND ScoreSinacofiCliente <836 THEN '785-835'
						  WHEN ScoreSinacofiCliente >=836 AND ScoreSinacofiCliente <862 THEN '836-861'
						  WHEN ScoreSinacofiCliente >=862 THEN '862-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana') -- AND CATEG_SINACOFI IS NULL ORDER BY 		ScoreSinacofiCliente	  

		
                                                                                                             
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso                                                                            
 SET CATEG_SCORE_INT= CASE WHEN BHV_CLI_ANT <568					THEN '000-567'                                   
						  WHEN BHV_CLI_ANT >=568 AND BHV_CLI_ANT <623 THEN '568-622'                                     		 
						  WHEN BHV_CLI_ANT >=623 AND BHV_CLI_ANT <655 THEN '623-654'                                     
						  WHEN BHV_CLI_ANT >=655 AND BHV_CLI_ANT <664 THEN '655-663'                                     
						  WHEN BHV_CLI_ANT >=664					  THEN '664-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')                               
                                                          
                                                                                                             
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso                                                                          
 SET RI= CASE            WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='000-567'  THEN 'E'                    
						 WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='568-622'  THEN 'D'                          
						 WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='623-654'  THEN 'C'                          
						 WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='655-663'  THEN 'B'                          
						 WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='664-999'  THEN 'B'                          
						                                                                                                 
						 WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='000-567'  THEN 'D'                          
						 WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='568-622'  THEN 'D'                          
						 WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='623-654'  THEN 'C'                          
						 WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='655-663'  THEN 'B'                          
						 WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='664-999'  THEN 'B'                          
						                                                                                                 
						 WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='000-567'  THEN 'D'                          
						 WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='568-622'  THEN 'B'                          
						 WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
						                                                                                                 
						 WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='000-567'  THEN 'A'                           
						 WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='568-622'  THEN 'A'                           
						 WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
						                                                                                                 
						 WHEN CATEG_SINACOFI='862-999' AND CATEG_SCORE_INT='000-567'  THEN 'A'                           
						 WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='568-622'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
						 WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                           
						 WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI                                       
						 WHERE SEGMENTO_LABEL='Personal Banking' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')                 
 		 				 

--------------------------- 
 /* 4  CF  NUEVO CAMPAÑA  */ 	--	WHERE 	SEGMENTO_LABEL='CF' AND 	desc_path_sw IN ('Nuevo Campana')
--------------------------- 

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente  <348								THEN '000-347'
						  WHEN ScoreSinacofiCliente >=348 AND ScoreSinacofiCliente <401 THEN '348-400'
						  WHEN ScoreSinacofiCliente >=401 AND ScoreSinacofiCliente <535 THEN '401-534'
						  WHEN ScoreSinacofiCliente >=535 AND ScoreSinacofiCliente <651 THEN '535-650'
						  WHEN ScoreSinacofiCliente >=651 AND ScoreSinacofiCliente <701 THEN '651-700'
						  WHEN ScoreSinacofiCliente >=701 AND ScoreSinacofiCliente <751 THEN '701-750'
						  WHEN ScoreSinacofiCliente >=751								THEN '751-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Nuevo Campana')-- AND CATEG_SINACOFI IS NULL ORDER BY 		ScoreSinacofiCliente	  

	

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET CATEG_SCORE_INT= CASE WHEN SCORE_CF_ORI_CALCULADO <541	    					    	THEN '000-540'
						  WHEN SCORE_CF_ORI_CALCULADO >=541 AND SCORE_CF_ORI_CALCULADO <551 THEN '541-550'
						  WHEN SCORE_CF_ORI_CALCULADO >=551 AND SCORE_CF_ORI_CALCULADO <567 THEN '551-566'
						  WHEN SCORE_CF_ORI_CALCULADO >=567 AND SCORE_CF_ORI_CALCULADO <581 THEN '567-580'
						  WHEN SCORE_CF_ORI_CALCULADO >=581 AND SCORE_CF_ORI_CALCULADO <591 THEN '581-590'
						  WHEN SCORE_CF_ORI_CALCULADO >=591 AND SCORE_CF_ORI_CALCULADO <611 THEN '591-610'
						  WHEN SCORE_CF_ORI_CALCULADO >=611 AND SCORE_CF_ORI_CALCULADO <621 THEN '611-620'
						  WHEN SCORE_CF_ORI_CALCULADO >=621 AND SCORE_CF_ORI_CALCULADO <631 THEN '621-630'
						  WHEN SCORE_CF_ORI_CALCULADO >=631						THEN '631-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Nuevo Campana')
  

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET RI= CASE			 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='567-580'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='581-590'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='591-610'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='611-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='621-630'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='000-540'  THEN 'E'
					 	 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='567-580'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='591-610'  THEN 'D'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='567-580'  THEN 'D'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='591-610'  THEN 'C'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='567-580'  THEN 'D'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='591-610'  THEN 'C'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='631-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='551-566'  THEN 'D'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='567-580'  THEN 'C'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='581-590'  THEN 'C'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='591-610'  THEN 'B'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='551-566'  THEN 'C'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='567-580'  THEN 'C'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='581-590'  THEN 'B'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='591-610'  THEN 'B'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='551-566'  THEN 'B'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='567-580'  THEN 'B'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='581-590'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='591-610'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
						WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Nuevo Campana')       
		 				 
--------------------------- 				  
 /* 5  CF NUEVO NO CAMPAÑA */ 	--	WHERE 	SEGMENTO_LABEL='CF' AND 	desc_path_sw='Nuevo'		  
--------------------------- 
				
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente  <348								THEN '000-347'
						  WHEN ScoreSinacofiCliente >=348 AND ScoreSinacofiCliente <401 THEN '348-400'
						  WHEN ScoreSinacofiCliente >=401 AND ScoreSinacofiCliente <535 THEN '401-534'
						  WHEN ScoreSinacofiCliente >=535 AND ScoreSinacofiCliente <651 THEN '535-650'
						  WHEN ScoreSinacofiCliente >=651 AND ScoreSinacofiCliente <701 THEN '651-700'
						  WHEN ScoreSinacofiCliente >=701 AND ScoreSinacofiCliente <751 THEN '701-750'
						  WHEN ScoreSinacofiCliente >=751								THEN '751-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw='Nuevo'		 -- AND CATEG_SINACOFI IS NULL ORDER BY 		ScoreSinacofiCliente	  
 
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET CATEG_SCORE_INT= CASE WHEN SCORE_CF_ORI_CALCULADO <541	    					    	THEN '000-540'
						  WHEN SCORE_CF_ORI_CALCULADO >=541 AND SCORE_CF_ORI_CALCULADO <551 THEN '541-550'
						  WHEN SCORE_CF_ORI_CALCULADO >=551 AND SCORE_CF_ORI_CALCULADO <567 THEN '551-566'
						  WHEN SCORE_CF_ORI_CALCULADO >=567 AND SCORE_CF_ORI_CALCULADO <581 THEN '567-580'
						  WHEN SCORE_CF_ORI_CALCULADO >=581 AND SCORE_CF_ORI_CALCULADO <591 THEN '581-590'
						  WHEN SCORE_CF_ORI_CALCULADO >=591 AND SCORE_CF_ORI_CALCULADO <611 THEN '591-610'
						  WHEN SCORE_CF_ORI_CALCULADO >=611 AND SCORE_CF_ORI_CALCULADO <621 THEN '611-620'
						  WHEN SCORE_CF_ORI_CALCULADO >=621 AND SCORE_CF_ORI_CALCULADO <631 THEN '621-630'
						  WHEN SCORE_CF_ORI_CALCULADO >=631						THEN '631-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw='Nuevo'		

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
 SET RI= CASE			 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='567-580'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='581-590'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='591-610'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='611-620'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='621-630'  THEN 'C'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='000-540'  THEN 'E'
					 	 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='567-580'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='591-610'  THEN 'D'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='567-580'  THEN 'D'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='591-610'  THEN 'C'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='631-999'  THEN 'B'
						                                                                       
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='567-580'  THEN 'D'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='581-590'  THEN 'D'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='591-610'  THEN 'C'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='611-620'  THEN 'B'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='621-630'  THEN 'B'
						 WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='631-999'  THEN 'A'
						                                                                       
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='551-566'  THEN 'D'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='567-580'  THEN 'C'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='581-590'  THEN 'C'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='591-610'  THEN 'B'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='551-566'  THEN 'C'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='567-580'  THEN 'C'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='581-590'  THEN 'B'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='591-610'  THEN 'B'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='000-540'  THEN 'E'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='541-550'  THEN 'D'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='551-566'  THEN 'B'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='567-580'  THEN 'B'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='581-590'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='591-610'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='611-620'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='621-630'  THEN 'A'
						 WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='631-999'  THEN 'A'
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
						WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw='Nuevo'	

--------------------------- 
 /* 6  CF  ANTIGUO */ 	--	WHERE 	SEGMENTO_LABEL='CF' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')
--------------------------- 

 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso
SET CATEG_SINACOFI=		CASE WHEN ScoreSinacofiCliente <348								THEN '000-347'
						  WHEN ScoreSinacofiCliente >=348 AND ScoreSinacofiCliente <526 THEN '348-525'
						  WHEN ScoreSinacofiCliente >=526 AND ScoreSinacofiCliente <591 THEN '526-590'
						  WHEN ScoreSinacofiCliente >=591 AND ScoreSinacofiCliente <663 THEN '591-662'
						  WHEN ScoreSinacofiCliente >=663 AND ScoreSinacofiCliente <710 THEN '663-709'
						  WHEN ScoreSinacofiCliente >=710 AND ScoreSinacofiCliente <733 THEN '710-732'
						  WHEN ScoreSinacofiCliente >=733 AND ScoreSinacofiCliente <762 THEN '733-761'
						  WHEN ScoreSinacofiCliente >=762 AND ScoreSinacofiCliente <779 THEN '762-778'
						  WHEN ScoreSinacofiCliente >=779								THEN '779-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana') -- AND CATEG_SINACOFI IS NULL ORDER BY 		ScoreSinacofiCliente	  


 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso                                                                           
 SET CATEG_SCORE_INT= CASE WHEN BHV_CLI_ANT <591 					THEN '000-590'                                   
						  WHEN BHV_CLI_ANT >=591 AND BHV_CLI_ANT <599 THEN '591-598'                                     		 
						  WHEN BHV_CLI_ANT >=599 AND BHV_CLI_ANT <605 THEN '599-604'                                     
						  WHEN BHV_CLI_ANT >=605 AND BHV_CLI_ANT <609 THEN '605-608'                                  
						  WHEN BHV_CLI_ANT >=609 AND BHV_CLI_ANT <616 THEN '609-615'   
						  WHEN BHV_CLI_ANT >=616					            THEN '616-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')                             
                                                              
                                                                                                             
 UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso                                                                          
 SET RI= CASE			 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='000-590'  THEN 'E' 
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='591-598'  THEN 'E'
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='599-604'  THEN 'E' 
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='605-608'  THEN 'E'                    
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='609-615'  THEN 'C' 
						 WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='616-999'  THEN 'C'                    
						                     
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='000-590'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='591-598'  THEN 'E'
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='599-604'  THEN 'D'
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='605-608'  THEN 'C'
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='609-615'  THEN 'C'
						 WHEN CATEG_SINACOFI='348-525' AND CATEG_SCORE_INT='616-999'  THEN 'C'
						                            
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='000-590'  THEN 'E'
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='591-598'  THEN 'E'
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='599-604'  THEN 'D'
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='605-608'  THEN 'C'
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='609-615'  THEN 'B'
						 WHEN CATEG_SINACOFI='526-590'  AND CATEG_SCORE_INT='616-999'  THEN 'B'
						   
						                         
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='000-590'  THEN 'D'
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='591-598'  THEN 'D'
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='599-604'  THEN 'D'
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='605-608'  THEN 'C'
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='609-615'  THEN 'B'
						 WHEN CATEG_SINACOFI= '591-662'  AND CATEG_SCORE_INT='616-999'  THEN 'B'
						 
						                           
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='000-590'  THEN 'D'
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='591-598'  THEN 'D'
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='599-604'  THEN 'D'
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='605-608'  THEN 'B'
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='609-615'  THEN 'A'
						 WHEN CATEG_SINACOFI= '663-709' AND CATEG_SCORE_INT='616-999'  THEN 'A'
						                            
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='000-590'  THEN 'C'
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='591-598'  THEN 'C'
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='599-604'  THEN 'C'
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='605-608'  THEN 'B'
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='609-615'  THEN 'A'
						 WHEN CATEG_SINACOFI='710-732'  AND CATEG_SCORE_INT='616-999'  THEN 'A'
						 
						    
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='000-590'  THEN 'C'
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='591-598'  THEN 'B' 
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='599-604'  THEN 'B' 
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='605-608'  THEN 'B' 
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='609-615'  THEN 'A' 
						 WHEN CATEG_SINACOFI='733-761' AND CATEG_SCORE_INT='616-999'  THEN 'A' 
						  
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='000-590'  THEN 'A'
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='591-598'  THEN 'A' 
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='599-604'  THEN 'A' 
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='605-608'  THEN 'A' 
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='609-615'  THEN 'A' 
						 WHEN CATEG_SINACOFI='762-778'  AND CATEG_SCORE_INT='616-999'  THEN 'A' 
						  
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='000-590'  THEN 'A'
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='591-598'  THEN 'A' 
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='599-604'  THEN 'A' 
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='605-608'  THEN 'A' 
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='609-615'  THEN 'A' 
						 WHEN CATEG_SINACOFI='779-999'  AND CATEG_SCORE_INT='616-999'  THEN 'A' 
						 						 END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI                                       
						 WHERE 	SEGMENTO_LABEL='Consumer Finance' AND 	desc_path_sw IN ('Antiguo','Antiguo Campana')              

	/* PEGAMOS MAXIMAS MORAS */

	drop table RT_FALVAREZ.DBO.sw_sols_res_paso2
	select a.*, moramax12m 
	into RT_FALVAREZ.DBO.sw_sols_res_paso2 -- select * 
	from RT_FALVAREZ.DBO.sw_sols_res_paso a left join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and a.camada=b.periodo 
	
	update RT_FALVAREZ.DBO.sw_sols_res_paso2
	set moramax12m = b.moramax12m -- select convert(varchar(6),dateadd(month,+1, left(a.camada,6)+'01'),112),* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and convert(varchar(6),dateadd(month,+1, left(a.camada,6)+'01'),112)=b.periodo 
	where a.moramax12m is null 
	
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso2 add tipo_producto nvarchar(20)
	
	update RT_FALVAREZ.DBO.sw_sols_res_paso2
	set tipo_producto='consumo'
	
	/****************************/
	/* Actualizo consumo monto **/
	/****************************/
	
	drop table RT_FALVAREZ.DBO.montos_consumo
	select rut,camada as periodo, max(convert(int,m_montosolicitado)) as monto
	into RT_FALVAREZ.DBO.montos_consumo -- select *
	from RT_FALVAREZ.DBO.sw_sols_res_paso where m_montosolicitado>0
	group by rut, camada
	order by rut, camada

	update A
	set a.m_montosolicitado = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a 
	inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=b.periodo
	where tipo_producto='consumo' 
	
	update A
	set a.m_montosolicitado = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , -1,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_montosolicitado <1
	
	update A
	set a.m_montosolicitado = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , +1,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_montosolicitado <1
	
	update A
	set a.m_montosolicitado = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , -2,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_montosolicitado <1
	
	update A
	set a.m_montosolicitado = b.monto -- select m_montosolicitado,monto,* 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join 	RT_FALVAREZ.DBO.montos_consumo b 
	on	a.rut=b.rut and a.camada=convert(varchar(6), dateadd(month , +2,left(b.periodo,6)+'01'),112)
	where tipo_producto='consumo'
	and  m_montosolicitado <1
	
	----------------------------------------------------------------------------------------------------------------------------
	
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso2 add baseorigen nvarchar(max) null
	
	update RT_FALVAREZ.DBO.sw_sols_res_paso2 
	set baseorigen=	'sw' 
	
	
	/*COMPLEMENTA RENTA  */	
	
	drop table RT_FALVAREZ.DBO.complementarentasw				
	select distinct rut, Solicitud,case when m_rentaconyugecliente >0 THEN 'SI' ELSE 'NO' END AS complementa_renta, m_rentaclientepesos,m_rentaconyugecliente
	into RT_FALVAREZ.DBO.complementarentasw
	from RT_FALVAREZ.DBO.paso1 

	alter table RT_FALVAREZ.DBO.sw_sols_res_paso2 add complementa_renta nvarchar(max) null

	update RT_FALVAREZ.DBO.sw_sols_res_paso2 
	set complementa_renta='SI' -- SELECT * 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join (SELECT RUT, SOLICITUD FROM RT_FALVAREZ.DBO.complementarentasw WHERE complementa_renta='SI' ) B
	ON A.Rut_num=B.RUT AND A.SOLICITUD=B.SOLICITUD
		
	update RT_FALVAREZ.DBO.sw_sols_res_paso2 
	set complementa_renta='NO'	
	WHERE complementa_renta IS NULL 
	---*--- 
	---*--- 
	
	/* aprobados por excepcion */
	
	ALTER TABLE RT_FALVAREZ.DBO.sw_sols_res_paso2 add APROBADO_EXCEPCION nvarchar(max) null
	
	UPDATE RT_FALVAREZ.DBO.sw_sols_res_paso2
	SET APROBADO_EXCEPCION=CASE WHEN (Resultado_SW='Rechazado' and ULTIMO_ESTADO_SOLICITUD='Aprobado Cursado') then 
	'SI' else 'NO' end  -- SELECT * 
	from RT_FALVAREZ.DBO.sw_sols_res_paso2	
				
	alter table RT_FALVAREZ.DBO.sw_sols_res_paso2 add [antiguedad] int null
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
									   
			
	update RT_FALVAREZ.DBO.sw_sols_res_paso2
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
      ,[nreneg12m]=b.[nreneg12m]
    from RT_FALVAREZ.DBO.sw_sols_res_paso2 a inner join [RT_SCORING].[dbo].[operaciones_critical_consol_min_perf] b
	on a.rut=b.rut and a.camada=b.periodo 
	
	
	
	
	--- fin CODIGO  --- 
	delete RT_SCORING.dbo.sw_sols_res -- select count(1) from RT_SCORING.dbo.sw_sols_res
	where segmento_label in ('Personal Banking','Consumer Finance','Banca Microempresa')
	and tipo_producto='consumo' and baseorigen='sw' and camada =  @periodo_inicial
	 
	--truncate table RT_SCORING.dbo.sw_sols_res
	insert into RT_SCORING.dbo.sw_sols_res
	SELECT  [segmento_label]
      ,[rut_num]
      ,[desc_path_sw]
      ,[resultado_sw]
      ,[fh_proceso]
      ,[score_sinacofi]
      ,[valoruf_sw]
      ,[score_interno]
      ,[score_bhv]
      ,[risk_indicator]
      ,[camada]
      ,[num_solicitudes]
      ,[ultimo_estado_solicitud]
      ,[fecha]
      ,[sistema]
      ,[solicitud]
      ,[c_sis]
      ,[tipoviviendacliente]
      ,[edad]
      ,[estadocivil]
      ,[niveleducacion]
      ,[n_antiguedadlaboral]
      ,[riesgoprofesion]
      ,[tipocargoactual]
      ,[tipodependecia]
      ,[b_bienraiz]
      ,[conveniocliente]
      ,[rut]
      ,[dicomcliente]
      ,[n_acreedores]
      ,[tiporentadeclarada]
      ,[m_rentaufcalculada]
      ,[m_tdsracreditado]
      ,[m_leveragenohipotecarioproyectado]
      ,[m_tdsr_clienteproyectado]
      ,[m_rentaclientepesos]
      ,[v_wrfprotestoscliente]
      ,[antiguedadcliente]
      ,[t_acc]
      ,[tipocliente]
      ,[b_clientenuevo]
      ,[b_campaña]
      ,[b_cumplebases]
      ,[m_valoruf]
      ,[producto]
      ,[plazo]
      ,[m_montosolicitado]
      ,[canal]
      ,[n_tarjeta]
      ,[b_tarjeta]
      ,[b_linea]
      ,[b_cuenta]
      ,[tarjeta1]
      ,[tarjeta2]
      ,[tarjeta3]
      ,[m_monedaextranjera]
      ,[m_montolinea]
      ,[m_montotarjeta]
      ,[b_margenvigente]
      ,[n_acreedorescliente]
      ,[productocuenta]
      ,[productolinea]
      ,[productotarjeta]
      ,[productoconsumo]
      ,[disponibilidadlineau6m]
      ,[antiguedaddeudasbifm6m]
      ,[b_lineadisponible]
      ,[b_moracomercio]
      ,[b_bureau]
      ,[b_funcionario]
      ,[b_jubilado]
      ,[b_consumoaldia]
      ,[b_clientecompracartera]
      ,[m_totalactivo]
      ,[m_totalpasivo]
      ,[m_deudaconsumonossa]
      ,[m_deudaconsumototalssa]
      ,[m_deudaconsumocuotassa]
      ,[m_cupotajetassa]
      ,[m_cupolineassa]
      ,[m_prepagoconsumossa]
      ,[m_prepagolineassa]
      ,[m_prepagotarjetassa]
      ,[m_prepagoconsumosbif]
      ,[n_logicpath]
      ,[n_risk]
      ,[n_puntajesw]
      ,[f_procesosw]
      ,[h_procesosw]
      ,[m_margenconsumo]
      ,[m_margenconsumodisponible]
      ,[m_margenlinea]
      ,[m_margenlineadisponible]
      ,[m_margentarjeta]
      ,[m_margentarjetadisponible]
      ,[m_margentotal]
      ,[m_margentotaldisponible]
      ,[decision_group]
      ,[b_ultimocambio]
      ,[filler_3]
      ,[v_predictorpublicequifax]
      ,[segmento]
      ,[b_sbif_u12m]
      ,[promedioconsumou6m_u12m]
      ,[scoresinacoficlientenobca]
      ,[v_pje_mod_01]
      ,[sc_bienraiz]
      ,[sc_displinea]
      ,[sc_promcons]
      ,[sc_nacre]
      ,[sc_edad]
      ,[sc_ant_lab]
      ,[sc_ddassa]
      ,[sc_bbureau]
      ,[sc_tdsr_acred]
      ,[score_cf_ori_calculado]
      ,[sc2_disponibilidadlineau6m]
      ,[sc2_m_deudaconsumonossa_renta]
      ,[sc2_promedioconsumou6m_u12m]
      ,[sc2_b_bienraiz]
      ,[sc2_b_bureau]
      ,[sc2_niveleducacion]
      ,[sc2_tipodependecia]
      ,[sc2_n_antiguedadlaboral]
      ,[sc2_edad]
      ,[score_pb_ori_calculado]
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
       ,'' as LeverageActual_Solicitud
	   ,'' as LeverageDisponible_Solicitud
	   ,'' as LeverageEnTramite_Solicitud
	   ,'' as LeverageExterno_Solicitud
	   ,'' as LeverageLimite_Solicitud
	   ,'' as LeverageProyectado_Solicitud
	   ,'' as LeverageSolicitado_Solicitud  -- select *
	   ,m_LeverageActualLinea as sw_m_LeverageActualLinea
		,m_LeverageSolicitadoLinea as sw_m_LeverageSolicitadoLinea
		,m_LeverageProyectadoLinea as sw_m_LeverageProyectadoLinea
		,m_LeverageActualTarjeta as sw_m_LeverageActualTarjeta
		,m_LeverageSolicitadoTarjeta as sw_m_LeverageSolicitadoTarjeta
		,m_LeverageProyectadoTarjeta as sw_m_LeverageProyectadoTarjeta
		,m_LeverageActualConsumo as sw_m_LeverageActualConsumo
		,m_LeverageSolicitadoConsumo as sw_m_LeverageSolicitadoConsumo
		,m_LeverageProyectadoConsumo as sw_m_LeverageProyectadoConsumo
		,m_LeverageActualComercial as sw_m_LeverageActualComercial
 		FROM  RT_FALVAREZ.DBO.sw_sols_res_paso2
		
	-- select distinct camada from RT_SCORING.DBO.sw_sols_res where [m_rentaclientepesos]>0
	 -- FROM rt_falvarez.dbo.sol_res_total 
	/* FIN  ACTUALIZACION DE CONSUMOS  EN BASE PRINCIPAL */

----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------
end