VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Estado_Resolucion_Final 
   Caption         =   "::::::::::::: Resumen Resolución Final"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   OleObjectBlob   =   "Estado_Resolucion_Final.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Estado_Resolucion_Final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_carta_rechazo_Click()
Unload Estado_Resolucion_Final
Carta_Cliente.Show
End Sub

Private Sub cmd_guardar_evaluacion_Click()

'Logica de Botones

Estado_Resolucion_Final.cmd_volver_pag_anterior.Enabled = False
Estado_Resolucion_Final.cmd_carta_rechazo.Enabled = True
Estado_Resolucion_Final.Imprimir_resolucion_f.Enabled = True
Estado_Resolucion_Final.cmd_volver_evaluacion.Enabled = True

''---------


Call conectarBD


   irespuesta = MsgBox("¿Esta Seguro Que Desea Guardar los Rechazos Finales?", vbYesNo)
    If irespuesta = vbYes Then

   
  
     ssql = "SELECT rut_cliente, max(n_solicitud) as n_solicitud FROM tbl_micro_ficha_cliente where rut_cliente = '" & txt_rut_cliente & "' group by rut_cliente"
        
               
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        
        If rst.EOF Then
           MsgBox ("Ejecutivo No Ingresado")
          Else
            If rst!rut_cliente = txt_rut_cliente Then
              txt_n_solicitud = rst!n_solicitud
            End If
          rst.MoveNext
        End If


If txt_r_f_mora_directa = "R" Or txt_r_f_Vencido_directo = "R" Or txt_r_f_castigo_directo = "R" Or txt_r_f_protesto_interno = "R" _
   Or txt_r_f_morosidad_sinac = "R" Or txt_r_f_protesto_sinac = "R" Or txt_r_f_boletin_sinac = "R" Or txt_r_f_mora_directa_SBIF = "R" _
    Or txt_r_f_vdo_directo_SBIF = "R" Or txt_r_f_cast_directo_SBIF = "R" Or txt_r_f_vdo_indirecto_SBIF = "R" Or txt_r_f_bancarizado_politica = "R" Then

txt_cod_9_sernac_final = "Codigo Rechazo Sernac 9 : Morosidad o Protestos Vigentes"

Else

txt_cod_9_sernac_final = 0

End If

'''

If txt_r_f_leverage = "R" Or txt_r_f_capacidad_pago = "R" Then
   
txt_cod_10_sernac_final = "Codigo Rechazo Sernac 10 : Excesiva Carga Financiera o de Endeudamiento"

Else
txt_cod_10_sernac_final = 0

End If

''''''''''

If txt_r_f_renegociado = "R" Or txt_r_f_file_negativo_tit = "R" Or txt_r_f_castigo_historico = "R" Or txt_r_f_cast_indirecto_SBIF = "R" Then
   
txt_cod_11_sernac_final = "Codigo Rechazo Sernac 11 : Incumplimiento Previo"

Else
txt_cod_11_sernac_final = 0

End If


''''''''''

If txt_r_f_n_acreedor = "R" Or txt_r_f_dir_comer_verif = "R" Or txt_r_f_visita_ejecutivo = "R" Or txt_r_f_telefono_verificado = "R" Or txt_r_f_direc_part_verif = "R" Or txt_r_f_plazo = "R" Or txt_r_f_destinos = "R" Or txt_r_f_antiguedad_veh = "R" Or txt_r_f_antiguedad_giro = "R" Or txt_r_f_bancarizado_politica = "R" Then
   
txt_cod_13_sernac_final = "Codigo Rechazo Sernac 13 : Incumplimiento en Parametros de politica de creditos"

Else
txt_cod_13_sernac_final = 0

End If

''''''''''

If txt_r_f_cod_observacion_cliente = "R" Or txt_r_f_ir_sinac = "R" Or txt_r_f_ir_tipo_cliente = "R" Then
   
txt_cod_14_sernac_final = "Codigo Rechazo Sernac 14 : Incumplimiento en Parametros de Score"

Else
txt_cod_14_sernac_final = 0

End If


''''''''''

If txt_r_f_edad = "R" Or txt_r_f_edad_maxima = "R" Then
   
txt_cod_15_sernac_final = "Codigo Rechazo Sernac 15 : Incumplimiento en Parametros de Edad"

Else
txt_cod_15_sernac_final = 0

End If


''''''''''

If txt_r_f_nivel_vta_inf_min = "R" Or txt_r_f_nivel_vta_sup_max = "R" Then
   
txt_cod_16_sernac_final = "Codigo Rechazo Sernac 16 : Incumplimiento en Parametros Renta"

Else
txt_cod_16_sernac_final = 0

End If


''''''''''

If txt_r_f_mora_directa_conyuge = "R" Or txt_r_f_Vencido_directo_conyuge = "R" Or txt_r_f_castigo_directo_conyuge = "R" Or txt_r_f_protesto_interno_conyuge = "R" Or txt_r_f_renegociado_conyuge = "R" Or txt_r_f_file_negativo_conyuge = "R" Or txt_r_f_castigo_historico_conyuge = "R" Or txt_r_f_morosidad_sinac_conyuge = "R" Or txt_r_f_protesto_sinac_conyuge = "R" Or txt_r_f_boletin_sinac_conyuge = "R" Or txt_r_f_acreedores_conyuge = "R" Or txt_r_f_cod_observacion_conyuge = "R" Or txt_r_f_ir_sinac_conyuge = "R" Or txt_r_f_mora_directa_SBIF_conyuge = "R" Or _
txt_r_f_vdo_directo_SBIF_conyuge = "R" Or txt_r_f_cast_directo_SBIF_conyuge = "R" Or txt_r_f_vdo_indirecto_SBIF_conyuge = "R" Or txt_r_f_cast_indirecto_SBIF_conyuge = "R" Or txt_r_f_edad_conyuge = "R" Or txt_r_f_edad_maxima_conyuge = "R" Or txt_r_f_mora_directa_aval = "R" Or txt_r_f_Vencido_directo_aval = "R" Or txt_r_f_castigo_directo_aval = "R" Or txt_r_f_protesto_interno_aval = "R" Or txt_r_f_renegociado_aval = "R" Or txt_r_f_file_negativo = "R" Or txt_r_f_castigo_historico_aval = "R" Or txt_r_f_morosidad_sinac_aval = "R" Or txt_r_f_protesto_sinac_aval = "R" Or _
txt_r_f_boletin_sinac_aval = "R" Or txt_r_f_acreedores_aval = "R" Or txt_r_f_cod_observacion_aval = "R" Or txt_r_f_ir_sinac_aval = "R" Or txt_r_f_mora_directa_SBIF_aval = "R" Or txt_r_f_vdo_indirecto_SBIF_aval = "R" Then

txt_cod_18_sernac_final = "Codigo Rechazo Sernac 18 : Incumplimiento en Parametros Renta"

Else
txt_cod_18_sernac_final = 0

End If
        


'----------------------------------------------------------
    
  
    ssql = "INSERT INTO TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F " _
    & "([Rut_Cliente],[dv],[n_solicitud],[r_f_mora_directa],[r_f_Vencido_directo],[r_f_castigo_directo],[r_f_protesto_interno],[r_f_renegociado],[r_f_file_negativo_tit],[r_f_castigo_historico],[r_f_morosidad_sinac],[r_f_protesto_sinac],[r_f_boletin_sinac],[r_f_n_acreedor],[r_f_cod_observacion_cliente],[r_f_ir_sinac],[r_f_mora_directa_SBIF],[r_f_vdo_directo_SBIF],[r_f_cast_directo_SBIF]," _
    & "[r_f_vdo_indirecto_SBIF],[r_f_cast_indirecto_SBIF],[r_f_edad],[r_f_edad_maxima],[r_f_dir_comer_verif],[r_f_visita_ejecutivo],[r_f_telefono_verificado],[r_f_direc_part_verif],[r_f_plazo],[r_f_destinos],[r_f_antiguedad_veh],[r_f_leverage],[r_f_capacidad_pago],[r_f_ir_tipo_cliente],[r_f_antiguedad_giro],[r_f_nivel_vta_inf_min],[r_f_nivel_vta_sup_max],[r_f_mora_directa_conyuge]," _
    & "[r_f_Vencido_directo_conyuge],[r_f_castigo_directo_conyuge],[r_f_protesto_interno_conyuge],[r_f_renegociado_conyuge],[r_f_file_negativo_conyuge],[r_f_castigo_historico_conyuge],[r_f_morosidad_sinac_conyuge],[r_f_protesto_sinac_conyuge],[r_f_boletin_sinac_conyuge],[r_f_acreedores_conyuge],[r_f_cod_observacion_conyuge],[r_f_ir_sinac_conyuge],[r_f_mora_directa_SBIF_conyuge],[r_f_vdo_directo_SBIF_conyuge]," _
    & "[r_f_cast_directo_SBIF_conyuge],[r_f_vdo_indirecto_SBIF_conyuge],[r_f_cast_indirecto_SBIF_conyuge],[r_f_edad_conyuge],[r_f_edad_maxima_conyuge],[r_f_mora_directa_aval],[r_f_Vencido_directo_aval],[r_f_castigo_directo_aval],[r_f_protesto_interno_aval],[r_f_renegociado_aval],[r_f_file_negativo_aval],[r_f_castigo_historico_aval],[r_f_morosidad_sinac_aval],[r_f_protesto_sinac_aval],[r_f_boletin_sinac_aval]," _
    & "[r_f_acreedores_aval],[r_f_cod_observacion_aval],[r_f_ir_sinac_aval],[r_f_mora_directa_SBIF_aval],[r_f_vdo_directo_SBIF_aval],[r_f_cast_directo_SBIF_aval],[r_f_vdo_indirecto_SBIF_aval],[r_f_cast_indirecto_SBIF_aval],[r_f_edad_aval],[r_f_edad_maxima_aval],[r_f_costo_fijo_rub_trasp],[r_f_deuda_sbif_declarada],[r_f_costo_variable_ponde],[r_f_compra_tot_mensual],[r_f_factor_ajuste_compra_tot_iva],[cod9],[cod10],[cod11],[cod13],[cod14],[cod15],[cod16],[cod18],[Resultado_APROBADO_final_cred],[Resultado_RECHAZADO_final_cred],[Resultado_ZONAGRIS_final_cred],[Resultado_final_Bancariza_Politica],[r_f_Mto_Maximo_Aut],[r_f_aviso_inconsis_cuota],[r_f_factibilidad_consumo],[r_f_Monto_Limite_consumo],[r_f_capacidad_pago_consumo],[r_f_plazo_consumo],[r_f_mto_max_consumo],[r_f_min_prepago_consumo],[r_f_min_prepago_comercial],[resultado_final_aprobado_consumo],[resultado_final_rechazado_consumo],[resolucion_final])" _
    & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_dv & "'),('" & txt_n_solicitud & "'),('" & txt_r_f_mora_directa & "'),('" & txt_r_f_Vencido_directo & "'),('" & txt_r_f_castigo_directo & "'),('" & txt_r_f_protesto_interno & "'),('" & txt_r_f_renegociado & "'),('" & txt_r_f_file_negativo_tit & "'),('" & txt_r_f_castigo_historico & "'),('" & txt_r_f_morosidad_sinac & "'),('" & txt_r_f_protesto_sinac & "'),('" & txt_r_f_boletin_sinac & "'),('" & txt_r_f_n_acreedor & "')" _
    & ",('" & txt_r_f_cod_observacion_cliente & "'),('" & txt_r_f_ir_sinac & "'),('" & txt_r_f_mora_directa_SBIF & "'),('" & txt_r_f_vdo_directo_SBIF & "'),('" & txt_r_f_cast_directo_SBIF & "'),('" & txt_r_f_vdo_indirecto_SBIF & "'),('" & txt_r_f_cast_indirecto_SBIF & "'),('" & txt_r_f_edad & "'),('" & txt_r_f_edad_maxima & "'),('" & txt_r_f_dir_comer_verif & "'),('" & txt_r_f_visita_ejecutivo & "'),('" & txt_r_f_telefono_verificado & "'),('" & txt_r_f_direc_part_verif & "'),('" & txt_r_f_plazo & "')" _
    & ",('" & txt_r_f_destinos & "'),('" & txt_r_f_antiguedad_veh & "'),('" & txt_r_f_leverage & "'),('" & txt_r_f_capacidad_pago & "'),('" & txt_r_f_ir_tipo_cliente & "'),('" & txt_r_f_antiguedad_giro & "'),('" & txt_r_f_nivel_vta_inf_min & "'),('" & txt_r_f_nivel_vta_sup_max & "'),('" & txt_r_f_mora_directa_conyuge & "'),('" & txt_r_f_Vencido_directo_conyuge & "'),('" & txt_r_f_castigo_directo_conyuge & "'),('" & txt_r_f_protesto_interno_conyuge & "'),('" & txt_r_f_renegociado_conyuge & "')" _
    & ",('" & txt_r_f_file_negativo & "'),('" & txt_r_f_castigo_historico_conyuge & "'),('" & txt_r_f_morosidad_sinac_conyuge & "'),('" & txt_r_f_protesto_sinac_conyuge & "'),('" & txt_r_f_boletin_sinac_conyuge & "'),('" & txt_r_f_acreedores_conyuge & "'),('" & txt_r_f_cod_observacion_conyuge & "'),('" & txt_r_f_ir_sinac_conyuge & "'),('" & txt_r_f_mora_directa_SBIF_conyuge & "'),('" & txt_r_f_vdo_directo_SBIF_conyuge & "'),('" & txt_r_f_cast_directo_SBIF_conyuge & "'),('" & txt_r_f_vdo_indirecto_SBIF_conyuge & "')" _
    & ",('" & txt_r_f_cast_indirecto_SBIF_conyuge & "'),('" & txt_r_f_edad_conyuge & "'),('" & txt_r_f_edad_maxima_conyuge & "'),('" & txt_r_f_mora_directa_aval & "'),('" & txt_r_f_Vencido_directo_aval & "'),('" & txt_r_f_castigo_directo_aval & "'),('" & txt_r_f_protesto_interno_aval & "'),('" & txt_r_f_renegociado_aval & "'),('" & txt_r_f_file_negativo_aval & "'),('" & txt_r_f_castigo_historico_aval & "'),('" & txt_r_f_morosidad_sinac_aval & "'),('" & txt_r_f_protesto_sinac_aval & "'),('" & txt_r_f_boletin_sinac_aval & "')" _
    & ",('" & txt_r_f_acreedores_aval & "'),('" & txt_r_f_cod_observacion_aval & "'),('" & txt_r_f_ir_sinac_aval & "'),('" & txt_r_f_mora_directa_SBIF_aval & "'),('" & txt_r_f_vdo_directo_SBIF_aval & "'),('" & txt_r_f_cast_directo_SBIF_aval & "'),('" & txt_r_f_vdo_indirecto_SBIF_aval & "'),('" & txt_r_f_cast_indirecto_SBIF_aval & "'),('" & txt_r_f_edad_aval & "'),('" & txt_r_f_edad_maxima_aval & "'),('" & txt_r_f_costo_fijo_rub_trasp & "'),('" & txt_r_f_deuda_sbif_declarada & "'),('" & txt_r_f_costo_variable_ponde & "'),('" & txt_r_f_compra_tot_mensual & "'),('" & txt_r_f_factor_ajuste_compra_tot_iva & "')" _
    & ",('" & txt_cod_9_sernac_final & "'),('" & txt_cod_10_sernac_final & "'),('" & txt_cod_11_sernac_final & "'),('" & txt_cod_13_sernac_final & "'),('" & txt_cod_14_sernac_final & "'),('" & txt_cod_15_sernac_final & "'),('" & txt_cod_16_sernac_final & "'),('" & txt_cod_18_sernac_final & "'),('" & txt_resultado_APROBADO_final_cred & "'),('" & txt_resultado_RECHAZADO_final_cred & "'),('" & txt_resultado_ZONAGRIS_final_cred & "'),('" & txt_r_f_bancarizado_politica & "'),('" & txt_r_f_mto_maximo_aut & "'),('" & txt_r_f_aviso_inconsis_cuota & "'),('" & txt_r_f_factibilidad_consumo & "'),('" & txt_r_f_Monto_Limite_consumo & "'),('" & txt_r_f_capacidad_pago_consumo & "'),('" & txt_r_f_plazo_consumo & "'),('" & txt_r_f_mto_max_consumo & "'),('" & txt_r_f_min_prepago_consumo & "'),('" & txt_r_f_min_prepago_comercial & "'),('" & txt_resultado_final_aprobado_consumo & "'),('" & txt_resultado_final_rechazado_consumo & "'),('" & txt_resolucion_final & "'))"
    
    cnn.Execute ssql
    
    
    
    '''trae la ultima fecha de evaluacion
        
    
        ssql = "select max(fecha_ingreso) fecha_ingreso" _
        & " from TBL_MICRO_ficha_cliente" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        Carta_Cliente.txt_fecha_actual = rst!FECHA_INGRESO
    
        Carta_Cliente.txt_rut_cliente = txt_rut_cliente
        Carta_Cliente.txt_n_solicitud = txt_n_solicitud

        
        
         cmd_carta_rechazo.Enabled = True
         Imprimir_resolucion_f.Enabled = True
         cmd_guardar_evaluacion.Enabled = False
        
    'ssql = "select rut_cliente,N_SOLICITUD,cod9" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
     '   Set rst = cnn.Execute(ssql, , adCmdText)
        


'If rst!cod9 <> 0 Then
'   ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
       & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
       & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_9_sernac_final & "'))"
        
'    cnn.Execute ssql'
'End If

    'ssql = "select rut_cliente,N_SOLICITUD,cod10" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
    '    Set rst = cnn.Execute(ssql, , adCmdText)

'
    'If rst!cod10 <> 0 Then
    
    'ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_10_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If
'

    'ssql = "select rut_cliente,N_SOLICITUD,cod11" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
     '   Set rst = cnn.Execute(ssql, , adCmdText)

'    If rst!cod11 <> 0 Then
    
 '   ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_11_sernac_final & "'))"
        
  '  cnn.Execute ssql
   ' End If
    
    'ssql = "select rut_cliente,N_SOLICITUD,cod13" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
     '   Set rst = cnn.Execute(ssql, , adCmdText)
    
'
    'If rst!cod13 <> 0 Then
    
    'ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_13_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If

'ssql = "select rut_cliente,N_SOLICITUD,cod14" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
 '       Set rst = cnn.Execute(ssql, , adCmdText)

'
  '  If rst!cod14 <> 0 Then
    
   ' ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_14_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If
'

'ssql = "select rut_cliente,N_SOLICITUD,cod15" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
 '       Set rst = cnn.Execute(ssql, , adCmdText)

  '  If rst!cod15 <> 0 Then
    
   ' ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_15_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If
'

'ssql = "select rut_cliente,N_SOLICITUD,cod16" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
 '       Set rst = cnn.Execute(ssql, , adCmdText)
        
  '  If rst!cod16 <> 0 Then
    
   ' ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_16_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If
'

'ssql = "select rut_cliente,N_SOLICITUD,cod18" _
        & " from TBL_MICRO_MAESTRO_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'"
        
 '       Set rst = cnn.Execute(ssql, , adCmdText)
        
  '  If rst!cod18 <> 0 Then
    
   ' ssql = "insert into TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & "([rut_cliente],[n_solicitud],[Tot_codigo_rechazo])" _
        & " VALUES (('" & txt_rut_cliente & "'), ('" & txt_n_solicitud & "'),('" & txt_cod_18_sernac_final & "'))"
        
    'cnn.Execute ssql
    'End If

        
    '    ssql = "select count(*) tot_codigo_rechazo" _
        & " from TBL_MICRO_TMP_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'" _
        & " and n_Solicitud = '" & txt_n_solicitud & "'"
        
     '   Set rst = cnn.Execute(ssql, , adCmdText)
        
      '  Carta_Cliente.txt_contador_rechazo = rst!tot_codigo_rechazo
        
        
End If

End Sub

Private Sub cmd_resolucion_f_Click()


''''' CALCULANDO RESULTADO FINAL DE EVALUACION CON VARIABLES RESUMIDAS

txt_resultado_APROBADO_final_cred = Empty
txt_resultado_ZONAGRIS_final_cred = Empty
txt_resultado_RECHAZADO_final_cred = Empty

txt_resultado_APROBADO_final_cred.BackColor = &H80000005
txt_resultado_ZONAGRIS_final_cred.BackColor = &H80000005
txt_resultado_RECHAZADO_final_cred.BackColor = &H80000005


If txt_r_f_mora_directa = "R" Or txt_r_f_Vencido_directo = "R" Or txt_r_f_castigo_directo = "R" Or txt_r_f_protesto_interno = "R" Or txt_r_f_renegociado = "R" Or txt_r_f_file_negativo_tit = "R" Or txt_r_f_castigo_historico = "R" Or txt_r_f_morosidad_sinac = "R" Or txt_r_f_protesto_sinac = "R" Or txt_r_f_boletin_sinac = "R" Or txt_r_f_n_acreedor = "R" Or txt_r_f_cod_observacion_cliente = "R" Or txt_r_f_ir_sinac = "R" Or txt_r_f_mora_directa_SBIF = "R" Or txt_r_f_vdo_directo_SBIF = "R" Or txt_r_f_cast_directo_SBIF = "R" Or _
    txt_r_f_vdo_indirecto_SBIF = "R" Or txt_r_f_cast_indirecto_SBIF = "R" Or txt_r_f_edad = "R" Or txt_r_f_edad_maxima = "R" Or txt_r_f_dir_comer_verif = "R" Or txt_r_f_visita_ejecutivo = "R" Or txt_r_f_telefono_verificado = "R" Or txt_r_f_direc_part_verif = "R" Or txt_r_f_plazo = "R" Or txt_r_f_destinos = "R" Or txt_r_f_antiguedad_veh = "R" Or txt_r_f_leverage = "R" Or txt_r_f_capacidad_pago = "R" Or txt_r_f_ir_tipo_cliente = "R" Or txt_r_f_antiguedad_giro = "R" Or txt_r_f_nivel_vta_inf_min = "R" Or txt_r_f_nivel_vta_sup_max = "R" Or _
    txt_r_f_mora_directa_conyuge = "R" Or txt_r_f_Vencido_directo_conyuge = "R" Or txt_r_f_castigo_directo_conyuge = "R" Or txt_r_f_protesto_interno_conyuge = "R" Or txt_r_f_renegociado_conyuge = "R" Or _
    txt_r_f_file_negativo_conyuge = "R" Or txt_r_f_castigo_historico_conyuge = "R" Or txt_r_f_morosidad_sinac_conyuge = "R" Or txt_r_f_n_acreedor_conyuge = "R" Or txt_r_f_cod_observacion_cliente_conyuge = "R" Or txt_r_f_ir_sinac_conyuge = "R" Or txt_r_f_mora_directa_SBIF_conyuge = "R" Or _
    txt_r_f_vdo_directo_SBIF_conyuge = "R" Or txt_r_f_cast_directo_SBIF_conyuge = "R" Or txt_r_f_vdo_indirecto_SBIF_conyuge = "R" Or txt_r_f_cast_indirecto_SBIF_conyuge = "R" Or txt_r_f_edad_conyuge = "R" Or txt_r_f_edad_maxima_conyuge = "R" Or txt_r_f_mora_directa_aval = "R" Or txt_r_f_Vencido_directo_aval = "R" Or txt_r_f_castigo_directo_aval = "R" Or txt_r_f_protesto_interno_aval = "R" Or txt_r_f_renegociado_aval = "R" Or txt_r_f_file_negativo_tit_aval = "R" Or txt_r_f_castigo_historico_aval = "R" Or txt_r_f_morosidad_sinac_aval = "R" Or txt_r_f_protesto_sinac_aval = "R" Or _
    txt_r_f_boletin_sinac_aval = "R" Or txt_r_f_n_acreedor_aval = "R" Or txt_r_f_cod_observacion_cliente_aval = "R" Or txt_r_f_ir_sinac_aval = "R" Or txt_r_f_mora_directa_SBIF_aval = "R" Or txt_r_f_vdo_directo_SBIF_aval = "R" Or txt_r_f_cast_directo_SBIF_aval = "R" Or txt_r_f_vdo_indirecto_SBIF_aval = "R" Or txt_r_f_cast_indirecto_SBIF_aval = "R" Or txt_r_f_edad_aval = "R" Or txt_r_f_edad_maxima_aval = "R" Or txt_r_f_costo_fijo_rub_trasp = "R" Or txt_r_f_factor_ajuste_compra_tot_iva = "R" Or txt_r_f_bancarizado_politica = "R" Or txt_r_f_mto_maximo_aut = "R" Then
   
    txt_resultado_RECHAZADO_final_cred.BackColor = &HFF&       ' ROJO
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = "R"
    Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
    Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
    cmd_guardar_evaluacion.Enabled = True
 
Else
 
'''   *************************************************************** INGRESO CONYUGE

If txt_marca_conyuge = 1 Then
    
        If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then

            If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
            txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
            txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" Or txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_compra_tot_mensual = "A" Then

            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
            
            
        End If


        ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then

            If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
                txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
                txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
                txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
                txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" And txt_r_f_mto_maximo_aut = "A" And _
                txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" Then

                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
                Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
                Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
                cmd_guardar_evaluacion.Enabled = True
                
            Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
                
                
            End If


        ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then

            If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
                txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
                txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
                txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
                txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" And txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_mto_maximo_aut = "A" And _
                txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" Then
            
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
                Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
                Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
                cmd_guardar_evaluacion.Enabled = True
                
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
                
            End If


       ' Else

       '     Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
        '    Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
        '    Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
        '    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
        '    Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
        '    cmd_guardar_evaluacion.Enabled = True


    End If
End If
'End If


'''    *************************************************************** INGRESO AVAL

If txt_marca_aval = 1 Then

    If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then

        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_compra_tot_mensual = "A" Then

            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
            
       End If

    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then

        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" Then

            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        End If
    
    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then

        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" Then

            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        End If
    End If
End If
    


''''''''''''''*********************** SOLO CLIENTE

If txt_marca_conyuge = 0 And txt_marca_aval = 0 Then

    If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then
    
        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And txt_r_f_bancarizado_politica = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_aviso_inconsis_cuota = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_compra_tot_mensual = "A" Then

            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        End If
        
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then
    
            If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
                txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
                txt_r_f_factor_ajuste_compra_tot_iva = "A" And txt_r_f_bancarizado_politica = "A" And txt_r_f_mto_maximo_aut = "A" And txt_r_f_aviso_inconsis_cuota = "A" And _
                txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" Then
            
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
                Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
                Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
                cmd_guardar_evaluacion.Enabled = True
                
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
                
            End If
            
    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
    
            If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
                txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
                txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_bancarizado_politica = "A" And txt_r_f_mto_maximo_aut = "A" And txt_r_f_aviso_inconsis_cuota = "A" And _
                txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" Then
            
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
                Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
                Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
                Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
                cmd_guardar_evaluacion.Enabled = True
                
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
                
            End If
    End If
End If


''''''''''''''*********************** CONYUGE Y AVAL

If txt_marca_conyuge = 2 And txt_marca_aval = 2 Then

    If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then

        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
            txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
            txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_compra_tot_mensual = "A" Then
    
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        End If
        
   ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then
    
        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
            txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
            txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" And txt_r_f_factor_ajuste_compra_tot_iva = "A" Then
        
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True
            
            
        End If
    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
    
        If txt_r_f_mora_directa = "A" And txt_r_f_Vencido_directo = "A" And txt_r_f_castigo_directo = "A" And txt_r_f_protesto_interno = "A" And txt_r_f_renegociado = "A" And txt_r_f_file_negativo_tit = "A" And txt_r_f_castigo_historico = "A" And txt_r_f_morosidad_sinac = "A" And txt_r_f_protesto_sinac = "A" And txt_r_f_boletin_sinac = "A" And txt_r_f_n_acreedor = "A" And txt_r_f_cod_observacion_cliente = "A" And txt_r_f_ir_sinac = "A" And txt_r_f_mora_directa_SBIF = "A" And txt_r_f_vdo_directo_SBIF = "A" And txt_r_f_cast_directo_SBIF = "A" And _
            txt_r_f_vdo_indirecto_SBIF = "A" And txt_r_f_cast_indirecto_SBIF = "A" And txt_r_f_edad = "A" And txt_r_f_edad_maxima = "A" And txt_r_f_dir_comer_verif = "A" And txt_r_f_visita_ejecutivo = "A" And txt_r_f_telefono_verificado = "A" And txt_r_f_direc_part_verif = "A" And txt_r_f_plazo = "A" And txt_r_f_destinos = "A" And (txt_r_f_antiguedad_veh = "A" Or txt_r_f_antiguedad_veh = "N/A") And txt_r_f_leverage = "A" And txt_r_f_capacidad_pago = "A" And txt_r_f_ir_tipo_cliente = "A" And txt_r_f_antiguedad_giro = "A" And txt_r_f_nivel_vta_inf_min = "A" And txt_r_f_nivel_vta_sup_max = "A" And _
            txt_r_f_mora_directa_conyuge = "A" And txt_r_f_Vencido_directo_conyuge = "A" And txt_r_f_castigo_directo_conyuge = "A" And txt_r_f_protesto_interno_conyuge = "A" And txt_r_f_renegociado_conyuge = "A" And _
            txt_r_f_file_negativo = "A" And txt_r_f_castigo_historico_conyuge = "A" And txt_r_f_morosidad_sinac_conyuge = "A" And txt_r_f_acreedores_conyuge = "A" And txt_r_f_cod_observacion_conyuge = "A" And txt_r_f_ir_sinac_conyuge = "A" And txt_r_f_mora_directa_SBIF_conyuge = "A" And _
            txt_r_f_vdo_directo_SBIF_conyuge = "A" And txt_r_f_cast_directo_SBIF_conyuge = "A" And txt_r_f_vdo_indirecto_SBIF_conyuge = "A" And txt_r_f_cast_indirecto_SBIF_conyuge = "A" And txt_r_f_edad_conyuge = "A" And txt_r_f_edad_maxima_conyuge = "A" And txt_r_f_costo_fijo_rub_trasp = "A" And _
            txt_r_f_mora_directa_aval = "A" And txt_r_f_Vencido_directo_aval = "A" And txt_r_f_castigo_directo_aval = "A" And txt_r_f_protesto_interno_aval = "A" And txt_r_f_renegociado_aval = "A" And txt_r_f_file_negativo_aval = "A" And txt_r_f_castigo_historico_aval = "A" And txt_r_f_morosidad_sinac_aval = "A" And txt_r_f_protesto_sinac_aval = "A" And _
            txt_r_f_boletin_sinac_aval = "A" And txt_r_f_acreedores_aval = "A" And txt_r_f_cod_observacion_aval = "A" And txt_r_f_ir_sinac_aval = "A" And txt_r_f_mora_directa_SBIF_aval = "A" And txt_r_f_vdo_directo_SBIF_aval = "A" And txt_r_f_cast_directo_SBIF_aval = "A" And txt_r_f_vdo_indirecto_SBIF_aval = "A" And txt_r_f_cast_indirecto_SBIF_aval = "A" And txt_r_f_edad_aval = "A" And txt_r_f_edad_maxima_aval = "A" And txt_r_f_mto_maximo_aut = "A" And _
            txt_r_f_costo_fijo_rub_trasp = "A" And txt_r_f_deuda_sbif_declarada = "A" And txt_r_f_costo_variable_ponde = "A" Then
        
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.BackColor = &H8000&     ' VERDE
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.ForeColor = &HFFFFFF
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"
    
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = ""
    
            cmd_guardar_evaluacion.Enabled = True

        Else

            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.BackColor = &H808080
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred.ForeColor = &HFFFFFF       ' PLOMO
            Estado_Resolucion_Final.txt_resultado_ZONAGRIS_final_cred = "ZG"
            Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = ""
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = ""
            cmd_guardar_evaluacion.Enabled = True


        End If
    End If
End If



End If  '''' PRIMER IF


'''*******************************************************************
''''SE CEPILLA AL CLIENTE CUANDO ESTA EN ESTADO FINAL APROBADO
'''*******************************************************************

If Evaluacion_Perfil.txt_tipo_cliente = "Antiguo No Prime" And Evaluacion_Perfil.cbx_bancarizado = "No" Then
        txt_resultado_APROBADO_final_cred = Empty
        txt_resultado_ZONAGRIS_final_cred = "ZG"
        txt_resultado_RECHAZADO_final_cred = Empty
       
End If


''''''' CALCULO RESULTADO FINAL CREDITO DE CONSUMO

If txt_r_f_factibilidad_consumo = "R" Or txt_r_f_Monto_Limite_consumo = "R" Or txt_r_f_capacidad_pago_consumo = "R" Or txt_r_f_plazo_consumo = "R" Or txt_r_f_mto_max_consumo = "R" Or txt_r_f_min_prepago_consumo = "R" Or txt_r_f_min_prepago_comercial = "R" Then
    Estado_Resolucion_Final.txt_resultado_final_rechazado_consumo = "R"
Else
    Estado_Resolucion_Final.txt_resultado_final_aprobado_consumo = "A"
End If


''''' calculo de RESOLUCION FINAL DE AMBOS CREDITOS ---COMERCIAL CONSUMO

If txt_resultado_RECHAZADO_final_cred = "R" Or txt_resultado_final_rechazado_consumo = "R" Then
txt_resolucion_final = "R"

ElseIf txt_resultado_APROBADO_final_cred = "A" And txt_resultado_final_aprobado_consumo = "A" Then
txt_resolucion_final = "A"

ElseIf txt_resultado_ZONAGRIS_final_cred = "ZG" Then
txt_resolucion_final = "ZG"

End If

End Sub

Private Sub cmd_volver_evaluacion_Click()

    Unload Estado_Resolucion_Final
    Unload Carta_Cliente
    Unload Ficha_Cliente_Micro
    Unload Evaluacion_Perfil
    Unload Metodologia_Activo_Circulante
    Unload Metodologia_IVA1
    Unload Metodologia_Maxima_Prod

    Menu_Principal_Micro.Show



'If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then
'    Estado_Resolucion_Final.Hide
'    Metodologia_Activo_Circulante.Show

'ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then
'    Estado_Resolucion_Final.Hide
'    Metodologia_IVA1.Show

'ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
'    Estado_Resolucion_Final.Hide
'    Metodologia_Maxima_Prod.Show

'End If

End Sub

Private Sub CommandButton1_Click()


    

End Sub

Private Sub cmd_volver_pag_anterior_Click()
If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then
         
     Estado_Resolucion_Final.Hide
     Metodologia_Activo_Circulante.Show
     Metodologia_Activo_Circulante.cmd_resumen_Estado_Rechazo.Enabled = False
    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then
    Estado_Resolucion_Final.Hide
    Metodologia_IVA1.Show
    Metodologia_IVA1.cmd_resumen_Estado_Rechazo.Enabled = False
    
    ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
    Estado_Resolucion_Final.Hide
    Metodologia_Maxima_Prod.Show
    Metodologia_Maxima_Prod.cmd_resumen_Estado_Rechazo.Enabled = False
    

End If
End Sub

Private Sub CommandButton2_Click()
Estado_Resolucion_Final.Hide
Ficha_Cliente_Micro.Show
End Sub

Private Sub Frame8_Click()

End Sub

Private Sub Imprimir_resolucion_f_Click()
Estado_Resolucion_Final.PrintForm
End Sub

Private Sub Label78_Click()

End Sub

Private Sub txt_r_f_acreedores_aval_Change()
If Estado_Resolucion_Final.txt_r_f_acreedores_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_acreedores_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_acreedores_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_acreedores_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_acreedores_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_acreedores_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_acreedores_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_antiguedad_giro_Change()
If Estado_Resolucion_Final.txt_r_f_antiguedad_giro = "R" Then

  Estado_Resolucion_Final.txt_r_f_antiguedad_giro_cod_cl.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_antiguedad_giro_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_antiguedad_veh_Change()
If Estado_Resolucion_Final.txt_r_f_antiguedad_veh = "R" Then

  Estado_Resolucion_Final.txt_r_f_antiguedad_veh_cod_cl.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_antiguedad_veh_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_bancarizado_politica_Change()

End Sub

Private Sub txt_r_f_boletin_sinac_aval_Change()
If Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_boletin_sinac_Change()
If Estado_Resolucion_Final.txt_r_f_boletin_sinac = "R" Then

  Estado_Resolucion_Final.txt_r_f_boletin_sinac_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_boletin_sinac_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_boletin_sinac_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_capacidad_pago_Change()
If Estado_Resolucion_Final.txt_r_f_capacidad_pago = "R" Then

  Estado_Resolucion_Final.txt_r_f_capacidad_pago_cod_cl.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_capacidad_pago_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_directo_SBIF_aval_Change()
If Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_directo_SBIF_Change()
If Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_directo_SBIF_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_indirecto_SBIF_aval_Change()
If Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_indirecto_SBIF_Change()
If Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cast_indirecto_SBIF_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_directo_aval_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_directo_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_directo_aval_cod.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_castigo_directo_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_directo_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_directo = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_directo_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_castigo_directo_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_directo_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_directo_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_directo_conyuge_cod.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_castigo_directo_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_historico_aval_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_historico_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_historico_aval_cod.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_castigo_historico_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_historico_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_historico = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_historico_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_castigo_historico_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_castigo_historico_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_castigo_historico_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_castigo_historico_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_castigo_historico_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cod_observacion_aval_Change()
If Estado_Resolucion_Final.txt_r_f_cod_observacion_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_cod_observacion_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cod_observacion_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cod_observacion_cliente_Change()
If Estado_Resolucion_Final.txt_r_f_cod_observacion_cliente = "R" Then

  Estado_Resolucion_Final.txt_r_f_cod_observacion_cliente_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_cod_observacion_cliente_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_cod_observacion_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_cod_observacion_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_cod_observacion_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_cod_observacion_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_destinos_Change()
If Estado_Resolucion_Final.txt_r_f_destinos = "R" Then

  Estado_Resolucion_Final.txt_r_f_destinos_cod_cl.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_destinos_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_dir_comer_verif_Change()
If Estado_Resolucion_Final.txt_r_f_dir_comer_verif = "R" Then

  Estado_Resolucion_Final.txt_r_f_dir_comer_verif_cod_cl.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_dir_comer_verif_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_direc_part_verif_Change()
If Estado_Resolucion_Final.txt_r_f_direc_part_verif = "R" Then

  Estado_Resolucion_Final.txt_r_f_direc_part_verif_cod_cl.Visible = True
     
Else
  txt_r_f_direc_part_verif_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_aval_Change()
If Estado_Resolucion_Final.txt_r_f_edad_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_edad_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_Change()
If Estado_Resolucion_Final.txt_r_f_edad = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_edad_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_edad_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_edad_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_maxima_aval_Change()
If Estado_Resolucion_Final.txt_r_f_edad_maxima_aval = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_maxima_aval_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_edad_maxima_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_maxima_Change()
If Estado_Resolucion_Final.txt_r_f_edad_maxima = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_maxima_cod_cl.Visible = True
     
Else
   Estado_Resolucion_Final.txt_r_f_edad_maxima_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_file_negativo_Change()
If Estado_Resolucion_Final.txt_r_f_file_negativo = "R" Then

  Estado_Resolucion_Final.txt_r_f_file_negativo_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_renegociado_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_edad_maxima_conyuge_Change()
If Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge = "R" Then

  Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge_cod.Visible = True
     
Else
  Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_file_negativo_aval_Change()
If Estado_Resolucion_Final.txt_r_f_file_negativo_aval = "R" Then

  txt_r_f_file_negativo_aval_cod.Visible = True
     
Else
   txt_r_f_file_negativo_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_file_negativo_conyuge_Change()
If txt_r_f_file_negativo_conyuge = "R" Then

  txt_r_f_file_negativo_conyuge_cod.Visible = True
     
Else
  txt_r_f_file_negativo_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_file_negativo_tit_Change()
If txt_r_f_file_negativo_tit = "R" Then

  txt_r_f_file_negativo_tit_cod_cl.Visible = True
     
Else
   txt_r_f_file_negativo_tit_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_ir_sinac_aval_Change()
If txt_r_f_ir_sinac_aval = "R" Then

  txt_r_f_ir_sinac_aval_cod.Visible = True
     
Else
  txt_r_f_ir_sinac_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_ir_sinac_Change()
If txt_r_f_ir_sinac = "R" Then

  txt_r_f_ir_sinac_cod_cl.Visible = False
     
Else
   txt_r_f_ir_sinac_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_ir_sinac_conyuge_Change()
If txt_r_f_ir_sinac_conyuge = "R" Then

  txt_r_f_ir_sinac_conyuge_cod.Visible = True
     
Else
  txt_r_f_ir_sinac_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_ir_tipo_cliente_Change()
If txt_r_f_ir_tipo_cliente = "R" Then

  txt_r_f_ir_tipo_cliente_cod_cl.Visible = True
     
Else
  txt_r_f_ir_tipo_cliente_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_leverage_Change()
If txt_r_f_leverage = "R" Then

  txt_r_f_leverage_cod_cl.Visible = True
     
Else
  txt_r_f_leverage_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_aval_Change()
If txt_r_f_mora_directa_aval = "R" Then

  txt_r_f_mora_directa_aval_cod.Visible = True
     
Else
   txt_r_f_mora_directa_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_Change()
If txt_r_f_mora_directa = "R" Then

  txt_r_f_mora_directa_cod_cl.Visible = True
     
Else
   txt_r_f_mora_directa_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_conyuge_Change()
If txt_r_f_mora_directa_conyuge = "R" Then

  txt_r_f_mora_directa_conyuge_cod.Visible = True
     
Else
   txt_r_f_mora_directa_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_SBIF_aval_Change()
If txt_r_f_mora_directa_SBIF_aval = "R" Then

  txt_r_f_mora_directa_SBIF_aval_cod.Visible = True
     
Else
  txt_r_f_mora_directa_SBIF_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_SBIF_Change()
If txt_r_f_mora_directa_SBIF = "R" Then

  txt_r_f_mora_directa_SBIF_cod_cl.Visible = True
     
Else
   txt_r_f_mora_directa_SBIF_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_mora_directa_SBIF_conyuge_Change()
If txt_r_f_mora_directa_SBIF_conyuge = "R" Then

  txt_r_f_mora_directa_SBIF_conyuge_cod.Visible = True
     
Else
  txt_r_f_mora_directa_SBIF_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_morosidad_sinac_aval_Change()
If txt_r_f_morosidad_sinac_aval = "R" Then

  txt_r_f_morosidad_sinac_aval_cod.Visible = True
     
Else
   txt_r_f_morosidad_sinac_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_morosidad_sinac_Change()
If txt_r_f_morosidad_sinac = "R" Then

  txt_r_f_morosidad_sinac_cod_cl.Visible = True
     
Else
   txt_r_f_morosidad_sinac_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_morosidad_sinac_conyuge_Change()
If txt_r_f_morosidad_sinac_conyuge = "R" Then

  txt_r_f_morosidad_sinac_conyuge_cod.Visible = True
     
Else
  txt_r_f_morosidad_sinac_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_n_acreedor_Change()
If txt_r_f_n_acreedor = "R" Then

 txt_r_f_n_acreedor_cod_cl.Visible = True
     
Else
  txt_r_f_n_acreedor_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_nivel_vta_inf_min_Change()
If txt_r_f_nivel_vta_inf_min = "R" Then

  txt_r_f_nivel_vta_inf_min_cod_cl.Visible = True
     
Else
  txt_r_f_nivel_vta_inf_min_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_nivel_vta_sup_max_Change()
If txt_r_f_nivel_vta_sup_max = "R" Then

  txt_r_f_nivel_vta_sup_max_cod_cl.Visible = True
     
Else
  txt_r_f_nivel_vta_sup_max_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_plazo_Change()
If txt_r_f_plazo = "R" Then

  txt_r_f_plazo_cod_cl.Visible = True
     
Else
  txt_r_f_plazo_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_interno_aval_Change()
If txt_r_f_protesto_interno_aval = "R" Then

  txt_r_f_protesto_interno_aval_cod.Visible = True
     
Else
   txt_r_f_protesto_interno_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_interno_Change()
If txt_r_f_protesto_interno = "R" Then

  txt_r_f_protesto_interno_cod_cl.Visible = True
     
Else
   txt_r_f_protesto_interno_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_interno_conyuge_Change()
If txt_r_f_protesto_interno_conyuge = "R" Then

  txt_r_f_protesto_interno_conyuge_cod.Visible = True
     
Else
   txt_r_f_protesto_interno_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_sinac_aval_Change()
If txt_r_f_protesto_sinac_aval = "R" Then

  txt_r_f_protesto_sinac_aval_cod.Visible = True
     
Else
   txt_r_f_protesto_sinac_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_sinac_Change()
If txt_r_f_protesto_sinac = "R" Then

  txt_r_f_protesto_sinac_cod_cl.Visible = True
     
Else
   txt_r_f_protesto_sinac_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_protesto_sinac_conyuge_Change()
If txt_r_f_protesto_sinac_conyuge = "R" Then

  txt_r_f_protesto_sinac_conyuge_cod.Visible = True
     
Else
  txt_r_f_protesto_sinac_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_renegociado_aval_Change()
If txt_r_f_renegociado_aval = "R" Then

  txt_r_f_renegociado_aval_cod.Visible = True
     
Else
   txt_r_f_renegociado_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_renegociado_Change()
If txt_r_f_renegociado = "R" Then

  txt_r_f_renegociado_cod_cl.Visible = True
     
Else
   txt_r_f_renegociado_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_renegociado_conyuge_Change()
If txt_r_f_renegociado_conyuge = "R" Then

  txt_r_f_renegociado_conyuge_cod.Visible = True
     
Else
  txt_r_f_renegociado_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_telefono_verificado_Change()
If txt_r_f_telefono_verificado = "R" Then

  txt_r_f_telefono_verificado_cod_cl.Visible = True
     
Else
  txt_r_f_telefono_verificado_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_directo_SBIF_aval_Change()
If txt_r_f_vdo_directo_SBIF_aval = "R" Then

  txt_r_f_vdo_directo_SBIF_aval_cod.Visible = True
     
Else
  txt_r_f_vdo_directo_SBIF_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_directo_SBIF_Change()
If txt_r_f_vdo_directo_SBIF = "R" Then

  txt_r_f_vdo_directo_SBIF_cod_cl.Visible = True
     
Else
   txt_r_f_vdo_directo_SBIF_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_directo_SBIF_conyuge_Change()
If txt_r_f_vdo_directo_SBIF_conyuge = "R" Then

  txt_r_f_vdo_directo_SBIF_conyuge_cod.Visible = True
     
Else
  txt_r_f_vdo_directo_SBIF_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_indirecto_SBIF_aval_Change()
If txt_r_f_vdo_indirecto_SBIF_aval = "R" Then

  txt_r_f_vdo_indirecto_SBIF_aval_cod.Visible = True
     
Else
  txt_r_f_vdo_indirecto_SBIF_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_indirecto_SBIF_Change()
If txt_r_f_vdo_indirecto_SBIF = "R" Then

  txt_r_f_vdo_indirecto_SBIF_cod_cl.Visible = True
     
Else
   txt_r_f_vdo_indirecto_SBIF_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_vdo_indirecto_SBIF_conyuge_Change()
If txt_r_f_vdo_indirecto_SBIF_conyuge = "R" Then

  txt_r_f_vdo_indirecto_SBIF_conyuge_cod.Visible = True
     
Else
  txt_r_f_vdo_indirecto_SBIF_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_Vencido_directo_aval_Change()
If txt_r_f_Vencido_directo_aval = "R" Then

  txt_r_f_Vencido_directo_aval_cod.Visible = True
     
Else
   txt_r_f_Vencido_directo_aval_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_Vencido_directo_Change()
If txt_r_f_Vencido_directo = "R" Then

  txt_r_f_Vencido_directo_cod_cl.Visible = True
     
Else
   txt_r_f_Vencido_directo_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_r_f_Vencido_directo_conyuge_Change()
If txt_r_f_Vencido_directo_conyuge = "R" Then

  txt_r_f_Vencido_directo_conyuge_cod.Visible = True
     
Else
   txt_r_f_Vencido_directo_conyuge_cod.Visible = False
   
End If
End Sub

Private Sub txt_r_f_visita_ejecutivo_Change()
If txt_r_f_visita_ejecutivo = "R" Then

  txt_r_f_visita_ejecutivo_cod_cl.Visible = True
     
Else
  txt_r_f_visita_ejecutivo_cod_cl.Visible = False
   
End If
End Sub

Private Sub txt_resultado_APROBADO_final_cred_Change()

End Sub

Private Sub txt_resultado_RECHAZADO_final_cred_Change()

End Sub

Private Sub txt_resultado_ZONAGRIS_final_cred_Change()

End Sub



Private Sub UserForm_Initialize()

txt_cod_9_sernac_final = 0
txt_cod_10_sernac_final = 0
txt_cod_11_sernac_final = 0
txt_cod_13_sernac_final = 0
txt_cod_14_sernac_final = 0
txt_cod_15_sernac_final = 0
txt_cod_16_sernac_final = 0
txt_cod_18_sernac_final = 0

cbx_estado_Sic.AddItem "Rechazado"
cbx_estado_Sic.AddItem "Zona Gris"
cbx_estado_Sic.AddItem "Aprob.Eje.Sin Facul."
cbx_estado_Sic.AddItem "Aprab.Eje.Con Facul.Créd.Con.Excep."
cbx_estado_Sic.AddItem "Aprob.Eje.Con Facul."
cbx_estado_Sic.AddItem "Aprob.Eje.Con Falcul.cred.sin.excep.-mto.> a.autorizado "



End Sub
