VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Visualizador_Rechazos_Sernac 
   Caption         =   ":::::::: Visualizador_Rechazos_Sernac"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14730
   OleObjectBlob   =   "Visualizador_Rechazos_Sernac.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Visualizador_Rechazos_Sernac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_carta_rechazo_Click()
Visualizador_Rechazos_Sernac.Hide


Carta_Cliente.txt_rut_cliente = txt_rut_cliente
Carta_Cliente.txt_n_solicitud = txt_n_solicitud
Carta_Cliente.txt_fecha_actual = txt_fecha_actual

Carta_Cliente.Show

End Sub

Private Sub cmd_resolucion_f_Click()

If txt_dv = txt_dv_compara Then


Call conectarBD


'''''''''' VERIFICACION CORRECTA DE EVALUACION  '''''''''

        'ssql = "Select top 1 a.n_solicitud,a.rut_cliente" _
        & " from tbl_micro_ficha_cliente a" _
        & " left join TBL_MICRO_PERFIL_RIESGO_CLIENTE       g   on a.n_solicitud = g.n_solicitud" _
        & " left join tbl_micro_metodologia_activo_circulante b on a.n_solicitud = b.n_solicitud" _
        & " left join tbl_micro_metodologia_maxima_produccion c on a.n_solicitud = c.n_solicitud" _
        & " left join tbl_micro_metodologia_iva d               on a.n_solicitud = d.n_solicitud" _
        & " left join TBL_MICRO_maestro_RECHAZOS_SERNAC_F e     on a.n_solicitud = e.n_solicitud" _
        & " where a.n_solicitud Is Not Null" _
        & " and g.n_solicitud is not null" _
        & " and e.n_solicitud is not null" _
        & " and(b.n_solicitud Is Not Null or c.n_solicitud Is Not Null or d.n_solicitud Is Not Null)" _
        & " and a.rut_cliente ='" & txt_rut_cliente & "'" _
        & " and a.n_solicitud>=35209" _
        & " group by  a.n_solicitud,a.rut_cliente" _
        & " order by a.n_solicitud desc" _


        ssql = "Select a.n_solicitud,a.rut_cliente from tbl_micro_ficha_cliente a left join tbl_micro_metodologia_activo_circulante b on a.n_solicitud = b.n_solicitud" _
        & " left join tbl_micro_metodologia_maxima_produccion c on a.n_solicitud = c.n_solicitud" _
        & " left join tbl_micro_metodologia_iva d on a.n_solicitud = d.n_solicitud" _
        & " where (b.n_solicitud Is Not Null Or c.n_solicitud Is Not Null Or d.n_solicitud Is Not Null)" _
        & " and a.rut_cliente ='" & txt_rut_cliente & "'" _
        & " group by  a.n_solicitud,a.rut_cliente" _
        & " order by a.n_solicitud desc" _

        Set rst = cnn.Execute(ssql, , adCmdText)
        
        If rst.EOF Then
         MsgBox ("Cliente Sin Ingreso De Evaluación o Fecha de evaluacion Consultada es anterior a la guardada "), vbCritical
         
         Else
         
'''''''''' TRAE EVALUACION  '''''''''
    ssql = "select top 1 [Rut_Cliente],[dv],[N_Solicitud],[r_f_mora_directa],[r_f_Vencido_directo],[r_f_castigo_directo],[r_f_protesto_interno],[r_f_renegociado],[r_f_file_negativo_tit],[r_f_castigo_historico],[r_f_morosidad_sinac],[r_f_protesto_sinac],[r_f_boletin_sinac],[r_f_n_acreedor],[r_f_cod_observacion_cliente],[r_f_ir_sinac]" _
        & ", [r_f_mora_directa_SBIF],[r_f_vdo_directo_SBIF],[r_f_cast_directo_SBIF],[r_f_vdo_indirecto_SBIF],[r_f_cast_indirecto_SBIF],[r_f_edad],[r_f_edad_maxima],[r_f_dir_comer_verif],[r_f_visita_ejecutivo],[r_f_telefono_verificado],[r_f_direc_part_verif],[r_f_plazo],[r_f_destinos],[r_f_antiguedad_veh],[r_f_leverage],[r_f_capacidad_pago]" _
        & ", [r_f_ir_tipo_cliente],[r_f_antiguedad_giro],[r_f_nivel_vta_inf_min],[r_f_nivel_vta_sup_max],[r_f_mora_directa_conyuge],[r_f_Vencido_directo_conyuge],[r_f_castigo_directo_conyuge],[r_f_protesto_interno_conyuge],[r_f_renegociado_conyuge],[r_f_file_negativo_conyuge],[r_f_castigo_historico_conyuge],[r_f_morosidad_sinac_conyuge],[r_f_protesto_sinac_conyuge]" _
        & ", [r_f_boletin_sinac_conyuge],[r_f_acreedores_conyuge],[r_f_cod_observacion_conyuge],[r_f_ir_sinac_conyuge],[r_f_mora_directa_SBIF_conyuge],[r_f_vdo_directo_SBIF_conyuge],[r_f_cast_directo_SBIF_conyuge],[r_f_vdo_indirecto_SBIF_conyuge],[r_f_cast_indirecto_SBIF_conyuge],[r_f_edad_conyuge],[r_f_edad_maxima_conyuge],[r_f_mora_directa_aval],[r_f_Vencido_directo_aval],[r_f_castigo_directo_aval]" _
        & ", [r_f_protesto_interno_aval],[r_f_renegociado_aval],[r_f_file_negativo_aval],[r_f_castigo_historico_aval],[r_f_morosidad_sinac_aval],[r_f_protesto_sinac_aval],[r_f_boletin_sinac_aval],[r_f_acreedores_aval],[r_f_cod_observacion_aval],[r_f_ir_sinac_aval],[r_f_mora_directa_SBIF_aval],[r_f_vdo_directo_SBIF_aval],[r_f_cast_directo_SBIF_aval],[r_f_vdo_indirecto_SBIF_aval],[r_f_cast_indirecto_SBIF_aval]" _
        & ", [r_f_edad_aval],[r_f_edad_maxima_aval],[r_f_costo_fijo_rub_trasp],[r_f_deuda_sbif_declarada],[r_f_costo_variable_ponde],[r_f_compra_tot_mensual],[r_f_factor_ajuste_compra_tot_iva],[Cod9],[Cod10],[Cod11],[Cod13],[Cod14],[Cod15],[Cod16],[Cod18],ISNULL([Resultado_APROBADO_final_cred],'') Resultado_APROBADO_final_cred ,ISNULL([Resultado_RECHAZADO_final_cred],'') Resultado_RECHAZADO_final_cred" _
        & ", ISNULL([Resultado_ZONAGRIS_final_cred],'') Resultado_ZONAGRIS_final_cred, [Resultado_final_Bancariza_Politica], [r_f_Mto_Maximo_Aut],[r_f_aviso_inconsis_cuota],isnull([r_f_factibilidad_consumo],'SD')r_f_factibilidad_consumo,isnull([r_f_Monto_Limite_consumo],'SD')r_f_Monto_Limite_consumo,ISNULL([r_f_capacidad_pago_consumo],'SD')r_f_capacidad_pago_consumo,ISNULL([r_f_plazo_consumo],'SD')r_f_plazo_consumo,ISNULL([r_f_mto_max_consumo],'SD')r_f_mto_max_consumo" _
        & ", ISNULL([r_f_min_prepago_consumo],'SD')r_f_min_prepago_consumo,ISNULL([r_f_min_prepago_comercial],'SD')r_f_min_prepago_comercial,ISNULL([resultado_final_aprobado_consumo],'SD')resultado_final_aprobado_consumo,ISNULL([resultado_final_rechazado_consumo],'SD')resultado_final_rechazado_consumo,ISNULL([resolucion_final],'SD')resolucion_final" _
        & "  FROM TBL_MICRO_maestro_RECHAZOS_SERNAC_F" _
        & "  WHERE rut_cliente = '" & txt_rut_cliente & "'" _
        & "  order by n_solicitud desc" _

        Set rst = cnn.Execute(ssql, , adCmdText)
        
        If Not rst.EOF Then
        
        txt_n_solicitud = rst!n_solicitud
        txt_r_f_mora_directa = rst!r_f_mora_directa
        txt_r_f_Vencido_directo = rst!r_f_Vencido_directo
        txt_r_f_castigo_directo = rst!r_f_castigo_directo
        txt_r_f_protesto_interno = rst!r_f_protesto_interno
        txt_r_f_renegociado = rst!r_f_renegociado
        txt_r_f_file_negativo_tit = rst!r_f_file_negativo_tit
        txt_r_f_castigo_historico = rst!r_f_castigo_historico
        txt_r_f_morosidad_sinac = rst!r_f_morosidad_sinac
        txt_r_f_protesto_sinac = rst!r_f_protesto_sinac
        txt_r_f_boletin_sinac = rst!r_f_boletin_sinac
        txt_r_f_n_acreedor = rst!r_f_n_acreedor
        txt_r_f_cod_observacion_cliente = rst!r_f_cod_observacion_cliente
        txt_r_f_ir_sinac = rst!r_f_ir_sinac
        txt_r_f_mora_directa_SBIF = rst!r_f_mora_directa_SBIF
        txt_r_f_vdo_directo_SBIF = rst!r_f_vdo_directo_SBIF
        txt_r_f_cast_directo_SBIF = rst!r_f_cast_directo_SBIF
        txt_r_f_vdo_indirecto_SBIF = rst!r_f_vdo_indirecto_SBIF
        txt_r_f_cast_indirecto_SBIF = rst!r_f_cast_indirecto_SBIF
        txt_r_f_edad = rst!r_f_edad
        txt_r_f_edad_maxima = rst!r_f_edad_maxima
        txt_r_f_dir_comer_verif = rst!r_f_dir_comer_verif
        txt_r_f_visita_ejecutivo = rst!r_f_visita_ejecutivo
        txt_r_f_telefono_verificado = rst!r_f_telefono_verificado
        txt_r_f_direc_part_verif = rst!r_f_direc_part_verif
        txt_r_f_plazo = rst!r_f_plazo
        txt_r_f_destinos = rst!r_f_destinos
        txt_r_f_antiguedad_veh = rst!r_f_antiguedad_veh
        txt_r_f_leverage = rst!r_f_leverage
        txt_r_f_capacidad_pago = rst!r_f_capacidad_pago
        txt_r_f_ir_tipo_cliente = rst!r_f_ir_tipo_cliente
        txt_r_f_antiguedad_giro = rst!r_f_antiguedad_giro
        txt_r_f_nivel_vta_inf_min = rst!r_f_nivel_vta_inf_min
        txt_r_f_nivel_vta_sup_max = rst!r_f_nivel_vta_sup_max
        txt_r_f_mora_directa_conyuge = rst!r_f_mora_directa_conyuge
        txt_r_f_Vencido_directo_conyuge = rst!r_f_Vencido_directo_conyuge
        txt_r_f_castigo_directo_conyuge = rst!r_f_castigo_directo_conyuge
        txt_r_f_protesto_interno_conyuge = rst!r_f_protesto_interno_conyuge
        txt_r_f_renegociado_conyuge = rst!r_f_renegociado_conyuge
        txt_r_f_file_negativo = rst!r_f_file_negativo_conyuge
        txt_r_f_castigo_historico_conyuge = rst!r_f_castigo_historico_conyuge
        txt_r_f_morosidad_sinac_conyuge = rst!r_f_morosidad_sinac_conyuge
        txt_r_f_protesto_sinac_conyuge = rst!r_f_protesto_sinac_conyuge
        txt_r_f_boletin_sinac_conyuge = rst!r_f_boletin_sinac_conyuge
        txt_r_f_acreedores_conyuge = rst!r_f_acreedores_conyuge
        txt_r_f_cod_observacion_conyuge = rst!r_f_cod_observacion_conyuge
        txt_r_f_ir_sinac_conyuge = rst!r_f_ir_sinac_conyuge
        txt_r_f_mora_directa_SBIF_conyuge = rst!r_f_mora_directa_SBIF_conyuge
        txt_r_f_vdo_directo_SBIF_conyuge = rst!r_f_vdo_directo_SBIF_conyuge
        txt_r_f_cast_directo_SBIF_conyuge = rst!r_f_cast_directo_SBIF_conyuge
        txt_r_f_vdo_indirecto_SBIF_conyuge = rst!r_f_vdo_indirecto_SBIF_conyuge
        txt_r_f_cast_indirecto_SBIF_conyuge = rst!r_f_cast_indirecto_SBIF_conyuge
        txt_r_f_edad_conyuge = rst!r_f_edad_conyuge
        txt_r_f_edad_maxima_conyuge = rst!r_f_edad_maxima_conyuge
        txt_r_f_mora_directa_aval = rst!r_f_mora_directa_aval
        txt_r_f_Vencido_directo_aval = rst!r_f_Vencido_directo_aval
        txt_r_f_castigo_directo_aval = rst!r_f_castigo_directo_aval
        txt_r_f_protesto_interno_aval = rst!r_f_protesto_interno_aval
        txt_r_f_renegociado_aval = rst!r_f_renegociado_aval
        txt_r_f_file_negativo_aval = rst!r_f_file_negativo_aval
        txt_r_f_castigo_historico_aval = rst!r_f_castigo_historico_aval
        txt_r_f_morosidad_sinac_aval = rst!r_f_morosidad_sinac_aval
        txt_r_f_protesto_sinac_aval = rst!r_f_protesto_sinac_aval
        txt_r_f_boletin_sinac_aval = rst!r_f_boletin_sinac_aval
        txt_r_f_acreedores_aval = rst!r_f_acreedores_aval
        txt_r_f_cod_observacion_aval = rst!r_f_cod_observacion_aval
        txt_r_f_ir_sinac_aval = rst!r_f_ir_sinac_aval
        txt_r_f_mora_directa_SBIF_aval = rst!r_f_mora_directa_SBIF_aval
        txt_r_f_vdo_directo_SBIF_aval = rst!r_f_vdo_directo_SBIF_aval
        txt_r_f_cast_directo_SBIF_aval = rst!r_f_cast_directo_SBIF_aval
        txt_r_f_vdo_indirecto_SBIF_aval = rst!r_f_vdo_indirecto_SBIF_aval
        txt_r_f_cast_indirecto_SBIF_aval = rst!r_f_cast_indirecto_SBIF_aval
        txt_r_f_edad_aval = rst!r_f_edad_aval
        txt_r_f_edad_maxima_aval = rst!r_f_edad_maxima_aval
        txt_r_f_costo_fijo_rub_trasp = rst!r_f_costo_fijo_rub_trasp
        txt_r_f_deuda_sbif_declarada = rst!r_f_deuda_sbif_declarada
        txt_r_f_costo_variable_ponde = rst!r_f_costo_variable_ponde
        txt_r_f_compra_tot_mensual = rst!r_f_compra_tot_mensual
        txt_r_f_factor_ajuste_compra_tot_iva = rst!r_f_factor_ajuste_compra_tot_iva
        txt_cod_9_sernac_final = rst!Cod9
        txt_cod_10_sernac_final = rst!Cod10
        txt_cod_11_sernac_final = rst!Cod11
        txt_cod_13_sernac_final = rst!Cod13
        txt_cod_14_sernac_final = rst!Cod14
        txt_cod_15_sernac_final = rst!Cod15
        txt_cod_16_sernac_final = rst!Cod16
        txt_cod_18_sernac_final = rst!Cod18
        txt_resultado_APROBADO_final_cred = rst!Resultado_APROBADO_final_cred
        txt_resultado_RECHAZADO_final_cred = rst!Resultado_RECHAZADO_final_cred
        txt_resultado_ZONAGRIS_final_cred = rst!resultado_ZONAGRIS_final_cred
        txt_r_f_bancarizado_politica = rst!Resultado_final_Bancariza_Politica
        txt_r_f_mto_maximo_aut = rst!r_f_Mto_Maximo_Aut
        txt_r_f_aviso_inconsis_cuota = rst!r_f_aviso_inconsis_cuota
        
        ''''parametros de consumo
        
        txt_r_f_factibilidad_consumo = rst!r_f_factibilidad_consumo
        txt_r_f_Monto_Limite_consumo = rst!r_f_Monto_Limite_consumo
        txt_r_f_capacidad_pago_consumo = rst!r_f_capacidad_pago_consumo
        txt_r_f_plazo_consumo = rst!r_f_plazo_consumo
        txt_r_f_mto_max_consumo = rst!r_f_mto_max_consumo
        txt_r_f_min_prepago_consumo = rst!r_f_min_prepago_consumo
        txt_r_f_min_prepago_comercial = rst!r_f_min_prepago_comercial
        txt_resultado_final_aprobado_consumo = rst!resultado_final_aprobado_consumo
        txt_resultado_final_rechazado_consumo = rst!resultado_final_rechazado_consumo
        txt_resolucion_final = rst!resolucion_final
        
                
        If rst!r_f_aviso_inconsis_cuota = "ZG" Then
        lbl_aviso_resolucion_final.Visible = True
        txt_cuerpo_aviso.Visible = True
        Else
        lbl_aviso_resolucion_final.Visible = False
        txt_cuerpo_aviso.Visible = False
        
        End If
        
        
        ''''TRAE FECHA Y HORA DE ULTIMA EVALUACION CORRECTA
        
        ssql = "select FECHA_INGRESO, HORA_INGRESO  FROM TBL_MICRO_FICHA_CLIENTE " _
        & " WHERE RUT_CLIENTE = '" & txt_rut_cliente & "'" _
        & " AND N_SOLiCITUD = '" & txt_n_solicitud & "'" _
              
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        txt_fecha_actual = rst!FECHA_INGRESO
        txt_hora_actual = rst!HORA_INGRESO
        
        
        ''''TRAE Metodologia Asignada a la ultima evaluacion correcta
        
        ssql = "select Metodologia_asignada  FROM tbl_micro_perfil_riesgo_cliente " _
        & " WHERE RUT_CLIENTE = '" & txt_rut_cliente & "'" _
        & " AND N_SOLiCITUD = '" & txt_n_solicitud & "'" _
              
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        TXT_ESTADO_METODOLOGIA_OCUPADA = rst!metodologia_asignada
        
        
        '''' trae aviso de inconsistencia en cuota
        

Else
         MsgBox ("Evaluación sin Datos en Rechazos ... Evaluación Incompleta Vuelva a Evaluar"), vbCritical
End If
End If
Else
   MsgBox "El Rut o Digito Verificador esta mal ingresado... Revise"

End If


End Sub

Private Sub cmd_volver_evaluacion_Click()
Unload Visualizador_Rechazos_Sernac
Menu_Principal_Micro.Show
End Sub

Private Sub Imprimir_resolucion_f_Click()
Visualizador_Rechazos_Sernac.PrintForm
End Sub



Private Sub txt_dv_Change()

   Dim I As Integer

    txt_dv = UCase(txt_dv)
    I = Len(txt_dv)
    txt_dv.SelStart = I
End Sub

Private Sub txt_rut_cliente_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
 

    Dim diga As Variant
       
    If Not IsNumeric(txt_rut_cliente) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_cliente = Empty
      End If
      
  
  
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_rut_cliente = Replace(txt_rut_cliente, "-", "")
    txt_rut_cliente = Replace(txt_rut_cliente, ".", "")
    txt_rut_cliente = Replace(txt_rut_cliente, ",", "")
    Vari3 = 2
    For I = 0 To Len(txt_rut_cliente) - 1
     If Left(Right(txt_rut_cliente, I + 1), 1) <> "." Then
      Vari1 = Vari1 + Left(Right(txt_rut_cliente, I + 1), 1) * Vari3
      Vari2 = Vari1 Mod 11
      Select Case Vari2
       Case 0
        txt_dv_compara.Text = "0"
       Case 1
        txt_dv_compara.Text = "K"
       Case Else
        txt_dv_compara.Text = 11 - Vari2
      End Select
      If Vari3 = 7 Then
       Vari3 = 2
      Else
       Vari3 = Vari3 + 1
      End If
     End If
    Next
    'fin digito verificador

End Sub

Private Sub UserForm_Click()

End Sub
