VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ficha_Cliente_Micro 
   Caption         =   "::::: Ficha Ingreso MicroEmpresa"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12705
   OleObjectBlob   =   "Ficha_Cliente_Micro.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Ficha_Cliente_Micro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbx_Accion_Change()

txt_estado_credito = Empty

cbx_Destino_Credito_Cap_Tra = Empty
cbx_Destino_Credito_Act_Fij = Empty
cbx_Destino_Credito_Nec_Per = Empty
cbx_Destino_Credito_Vivienda = Empty
txt_r_plazo_credito = Empty
txt_plazo_credito = Empty


If cbx_Accion = "Terreno" Or cbx_Accion = "Construccion" Or cbx_Accion = "Necesidades Personales" Then
  
  txt_r_accion = "R"
  lbl_accion.BackColor = &HFF&       'rojo
  lbl_accion.ForeColor = &H8000000E  'blanco
  txt_r_años_vehiculo = "N/A"
  txt_r_plazo_credito = "N/A"
  Estado_Resolucion_Final.txt_r_f_plazo = "N/A"
  
  Else
  
  txt_r_accion = "A"
  lbl_accion.BackColor = &HC000&
  lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes

End If


If cbx_Accion = "Capital De Trabajo" Then
lbl_tipo_credito_destino.Visible = True
cbx_Destino_Credito_Cap_Tra.Visible = True
cbx_Destino_Credito_Act_Fij.Visible = False
cbx_Destino_Credito_Nec_Per.Visible = False
cbx_Destino_Credito_Vivienda.Visible = False
cbx_Destino_Construccion.Visible = False
cbx_destino_terreno.Visible = False
txt_r_años_vehiculo.Visible = False
txt_r_años_vehiculo = "N/A"
txt_r_plazo_credito = "N/A"
Estado_Resolucion_Final.txt_r_f_plazo = "N/A"

cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
cbx_tipo_vehiculo.Visible = False
cbx_años_vehiculo_bus.Visible = False
cbx_años_vehiculo_auto.Visible = False
cbx_destino_terreno.Visible = False
txt_plazo_credito.Visible = True


ElseIf cbx_Accion = "Activo Fijo" Then
cbx_Destino_Credito_Act_Fij.Visible = True
lbl_tipo_credito_destino.Visible = True
cbx_Destino_Credito_Cap_Tra.Visible = False
cbx_Destino_Credito_Nec_Per.Visible = False
cbx_Destino_Credito_Vivienda.Visible = False
cbx_Destino_Construccion.Visible = False
cbx_destino_terreno.Visible = False
txt_r_años_vehiculo.Visible = True
txt_plazo_credito.Visible = True

'''''''''''''''agregado por C.M.A. 18-03-2011

ElseIf cbx_Accion = "Vivienda" Then
cbx_Destino_Credito_Act_Fij.Visible = False
lbl_tipo_credito_destino.Visible = True
cbx_Destino_Credito_Vivienda.Visible = True
cbx_Destino_Credito_Cap_Tra.Visible = False
cbx_Destino_Credito_Nec_Per.Visible = False
cbx_Destino_Construccion.Visible = False
cbx_destino_terreno.Visible = False
txt_r_años_vehiculo.Visible = False
txt_r_años_vehiculo = "N/A"
txt_r_plazo_credito = "N/A"
Estado_Resolucion_Final.txt_r_f_plazo = "N/A"

cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
cbx_tipo_vehiculo.Visible = False
cbx_años_vehiculo_bus.Visible = False
cbx_años_vehiculo_auto.Visible = False
txt_plazo_credito.Visible = True

ElseIf cbx_Accion = "Necesidades Personales" Then
cbx_Destino_Credito_Nec_Per.Visible = True
lbl_tipo_credito_destino.Visible = True
cbx_Destino_Credito_Act_Fij.Visible = False
cbx_Destino_Credito_Cap_Tra.Visible = False
cbx_Destino_Credito_Vivienda.Visible = False
cbx_Destino_Construccion.Visible = False
cbx_destino_terreno.Visible = False
txt_r_años_vehiculo.Visible = False
txt_r_años_vehiculo = "N/A"
txt_r_plazo_credito = "N/A"
Estado_Resolucion_Final.txt_r_f_plazo = "N/A"

cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
cbx_tipo_vehiculo.Visible = False
cbx_años_vehiculo_bus.Visible = False
cbx_años_vehiculo_auto.Visible = False
txt_plazo_credito.Visible = False
'txt_r_plazo_credito = Empty

ElseIf cbx_Accion = "Construccion" Then

cbx_Destino_Credito_Act_Fij.Visible = False
cbx_Destino_Credito_Vivienda.Visible = False
cbx_Destino_Credito_Cap_Tra.Visible = False
cbx_Destino_Credito_Nec_Per.Visible = False
cbx_Destino_Construccion.Visible = True
cbx_destino_terreno.Visible = False
txt_r_años_vehiculo.Visible = False
txt_r_años_vehiculo = "N/A"
txt_r_plazo_credito = "N/A"
Estado_Resolucion_Final.txt_r_f_plazo = "N/A"

cbx_Destino_Construccion = Empty
cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
cbx_tipo_vehiculo.Visible = False
cbx_años_vehiculo_bus.Visible = False
cbx_años_vehiculo_auto.Visible = False
txt_plazo_credito.Visible = False
'txt_r_plazo_credito = Empty


ElseIf cbx_Accion = "Terreno" Then

cbx_Destino_Credito_Act_Fij.Visible = False
cbx_Destino_Credito_Vivienda.Visible = False
cbx_Destino_Credito_Cap_Tra.Visible = False
cbx_Destino_Credito_Nec_Per.Visible = False
cbx_Destino_Construccion.Visible = False
cbx_destino_terreno.Visible = True
txt_r_años_vehiculo.Visible = False
txt_r_años_vehiculo = "N/A"
txt_r_plazo_credito = "N/A"
Estado_Resolucion_Final.txt_r_f_plazo = "N/A"

cbx_destino_terreno = Empty
cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
cbx_tipo_vehiculo.Visible = False
cbx_años_vehiculo_bus.Visible = False
cbx_años_vehiculo_auto.Visible = False
txt_plazo_credito.Visible = False
'txt_r_plazo_credito = Empty

End If

End Sub

Private Sub cbx_acred_bien_raiz_Change()
txt_estado_credito = Empty
r_txt_acreditacion_bien_raiz = cbx_acred_bien_raiz
End Sub

Private Sub cbx_acred_bien_raiz_no_Change()
txt_estado_credito = Empty


r_txt_acreditacion_bien_raiz = cbx_acred_bien_raiz_no
End Sub


Private Sub cbx_acreditacion_vehiculo_Change()
txt_estado_credito = Empty
txt_r_acreditacion_vehiculo = cbx_acreditacion_vehiculo
End Sub

Private Sub cbx_acreditacion_vehiculo_no_Change()

txt_estado_credito = Empty


txt_r_acreditacion_vehiculo = cbx_acreditacion_vehiculo_no
End Sub

Private Sub cbx_actividad_economica_formal_Change()
txt_estado_credito = Empty
txt_r_cbx_actividad_economica_informal_oficio = Empty

txt_r_actividad_economica = cbx_actividad_economica_formal
txt_r_cbx_actividad_economica_informal_oficio = "A"

If txt_bancarizado_politica = "No" And (cbx_actividad_economica_formal = "COMIDA RAPIDA") Then
    txt_r_cbx_actividad_economica_informal_oficio = "R"
End If

End Sub

Private Sub cbx_actividad_economica_formal_servicio_Change()
txt_estado_credito = Empty
txt_r_cbx_actividad_economica_informal_oficio = Empty
txt_r_cbx_actividad_economica_informal_oficio = "A"
txt_r_actividad_economica = cbx_actividad_economica_formal_servicio

If txt_bancarizado_politica = "No" And (cbx_actividad_economica_formal_servicio = "CONTRATISTA CONSTRUCCION" Or cbx_actividad_economica_formal_servicio = "CONFECCION" Or cbx_actividad_economica_formal_servicio = "COMIDA RAPIDA") Then
    txt_r_cbx_actividad_economica_informal_oficio = "R"
End If

End Sub

Private Sub cbx_actividad_economica_informal_oficio_Change()

txt_estado_credito = Empty
txt_r_actividad_economica = cbx_actividad_economica_informal_oficio

End Sub

Private Sub cbx_actividad_economica_semiformal_Change()

txt_estado_credito = Empty
txt_r_actividad_economica = cbx_actividad_economica_semiformal

End Sub

Private Sub cbx_antiguedad_rubro_Change()

txt_estado_credito = Empty

If txt_bancarizado_politica = "Si" Then
    txt_r_cbx_antiguedad_rubro = "A"
ElseIf txt_bancarizado_politica = "No" And (cbx_antiguedad_rubro = "DAI" Or cbx_antiguedad_rubro = "Iniciación De Actividades" Or cbx_antiguedad_rubro = "Carpeta Tributaria") Then
       txt_r_cbx_antiguedad_rubro = "A"
Else
       txt_r_cbx_antiguedad_rubro = "R"
End If

End Sub

Private Sub cbx_años_vehiculo_auto_Change()

txt_estado_credito = Empty

If cbx_años_vehiculo_auto = "Mayor a 5 Años" Then
  
  txt_r_años_vehiculo = "R"
  
  Else
  
  txt_r_años_vehiculo = "A"


End If
End Sub

Private Sub cbx_años_vehiculo_bus_Change()

txt_estado_credito = Empty

If cbx_años_vehiculo_bus = "Mayor a 10 Años" Then
  
  txt_r_años_vehiculo = "R"
  
  Else
  
  txt_r_años_vehiculo = "A"

End If
End Sub

Private Sub cbx_aval_inf_com_Change()

txt_estado_credito = Empty

If cbx_aval_inf_com = "Cumple" Then
  txt_r_aval_inf_com = "A"
        txt_r_aval_inf_com.ForeColor = &H8000000E  'blanco
        
        
  Else
    txt_r_aval_inf_com = "R"
    txt_r_aval_inf_com.ForeColor = &H8000000E  'blanco
  
End If

End Sub

Private Sub cbx_bien_Raiz_Change()

txt_estado_credito = Empty
txt_r_cbx_bien_Raiz = Empty
txt_r_cbx_bien_Raiz = "A"
    
    'If txt_bancarizado_politica = "Si" Then
        cbx_acred_bien_raiz.Visible = True
        cbx_valor_evaluo_bien_raiz.Visible = True
        cbx_acred_bien_raiz_no.Visible = False
        cbx_valor_evaluo_bien_raiz_no.Visible = False

    If cbx_bien_Raiz = "Arrendado" Then
        cbx_acred_bien_raiz.Visible = False
        cbx_valor_evaluo_bien_raiz.Visible = False
        cbx_acred_bien_raiz_no.Visible = True
        cbx_valor_evaluo_bien_raiz_no.Visible = True
    End If
    'End If


If txt_bancarizado_politica = "No" And cbx_bien_Raiz = "Propio" Then
   txt_r_cbx_bien_Raiz = "A"
ElseIf txt_bancarizado_politica = "No" And cbx_bien_Raiz <> "Propio" Then
    txt_r_cbx_bien_Raiz = "R"
End If

End Sub

Private Sub cbx_boletin_laboral_Change()

txt_estado_credito = Empty


If cbx_boletin_laboral = "Cumple" Or cbx_boletin_laboral = "No Cliente Banco" Or cbx_boletin_laboral = "No Bancarizado" Then

   txt_r_boletin_laboral = "A"
        lbl_boletin_laboral.BackColor = &HC000&
        lbl_boletin_laboral.ForeColor = &H8000000E  'blanco
        'lbl_boletin_laboral.BorderStyle = fmBorderStyleSingle 'con bordes
  
    Else: txt_r_boletin_laboral = "R"
        lbl_boletin_laboral.BackColor = &HFF&       'rojo
    lbl_boletin_laboral.ForeColor = &H8000000E  'blanco
    'lbl_boletin_laboral.BorderStyle = fmBorderStyleSingle 'con bordes
    
End If

End Sub

Private Sub cbx_cod_ejecutivo_Change()
txt_estado_credito = Empty
End Sub

Private Sub cbx_codigo_sucursal_Change()

If txt_dv = txt_dv_compara Then

txt_estado_credito = Empty

cbx_cod_ejecutivo.Clear
txt_estado_credito = Empty

    Call conectarBD
    

    '''''''''' TRAE COD.NOMBRE EJECUTIVOS '''''''''
    ssql = "select codigo_ejecutivo +'      '+ nombre_ejecutivo +' '+ apellido_ejecutivo as EJECUTIVO FROM TBL_ejecutivo " _
    & " WHERE (CODIGO_EJECUTIVO > 0 and CODIGO_EJECUTIVO <>9999 and CODIGO_EJECUTIVO <>999) " _
    & " and codigo_sucursal = '" & cbx_codigo_sucursal & "'" _
    & " and (cargo_ejecutivo ='EJECUTIVO MICROEMPRESA' or cargo_ejecutivo = 'AGENTE SUCURSAL')" _
    & " and estado_ejecutivo ='ACTIVO'" _
    & " ORDER BY codigo_sucursal, CODIGO_EJECUTIVO"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_cod_ejecutivo.AddItem rst!EJECUTIVO
        rst.MoveNext
    Loop

Else
  MsgBox ("Corrija Rut De Conyuge es Invalido Revise ...")

End If

    

    '''''''''' TRAE NOMBRE EVALUADORES '''''''''
    ssql = "select codigo_sucursal+'-'+nombre_ejecutivo+' '+apellido_ejecutivo as EVALUADOR FROM TBL_ejecutivo " _
    & " where cargo_ejecutivo like '%evaluador%' and estado_ejecutivo ='Activo' " _
    & " and estado_ejecutivo ='ACTIVO'" _
    & " ORDER BY codigo_sucursal"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_evaluador.AddItem rst!EVALUADOR
        rst.MoveNext
    Loop



End Sub

Private Sub cbx_conyuge_inf_com_Change()

txt_estado_credito = Empty

If cbx_conyuge_inf_com = "Cumple" Then
  txt_r_conyuge_inf_com = "A"
        txt_r_conyuge_inf_com.ForeColor = &H8000000E  'blanco
        
        
  Else
    txt_r_conyuge_inf_com = "R"
    txt_r_conyuge_inf_com.ForeColor = &H8000000E  'blanco
  
End If


End Sub

Private Sub cbx_Destino_Credito_Act_Fij_Change()

txt_plazo_credito = Empty

txt_estado_credito = Empty

r_tipo_credito_destino = cbx_Destino_Credito_Act_Fij

If cbx_Destino_Credito_Act_Fij = "Vehiculo" Then
   cbx_tipo_vehiculo.Visible = True
   txt_r_años_vehiculo = ""
    
Else
    cbx_tipo_vehiculo = Empty
    cbx_tipo_vehiculo.Visible = False
    cbx_años_vehiculo_bus.Visible = False
    cbx_años_vehiculo_auto.Visible = False
    txt_r_años_vehiculo = "N/A"
    txt_r_plazo_credito = "N/A"
    Estado_Resolucion_Final.txt_r_f_plazo = "N/A"
 
End If

End Sub

Private Sub cbx_Destino_Credito_Cap_Tra_Change()
txt_estado_credito = Empty
r_tipo_credito_destino = cbx_Destino_Credito_Cap_Tra

End Sub

Private Sub cbx_Destino_Credito_Nec_Per_Change()

txt_estado_credito = Empty

r_tipo_credito_destino = cbx_Destino_Credito_Nec_Per


End Sub



Private Sub cbx_Destino_Credito_Vivienda_Change()

txt_estado_credito = Empty
r_tipo_credito_destino = cbx_Destino_Credito_Vivienda

End Sub



Private Sub cbx_destino_terreno_Change()


End Sub

Private Sub cbx_Dir_Comercial_Verif_Change()
txt_estado_credito = Empty

If cbx_Dir_Comercial_Verif = "Si" Then
   txt_r_dir_comer_verif = "A"
   Estado_Resolucion_Final.txt_r_f_dir_comer_verif = "A"
   
   Else
      txt_r_dir_comer_verif = "ZG"
      Estado_Resolucion_Final.txt_r_f_dir_comer_verif = "ZG"
   
End If


End Sub

Private Sub cbx_direc_Part_Verif_Change()
txt_estado_credito = Empty

If cbx_direc_Part_Verif = "Si" Then
   
   txt_r_direc_part_verif = "A"
   Estado_Resolucion_Final.txt_r_f_direc_part_verif = "A"
   
   Else
      txt_r_direc_part_verif = "R"
      Estado_Resolucion_Final.txt_r_f_direc_part_verif = "R"
   
End If

End Sub

Private Sub cbx_ejecutivo_excepcion_Change()

End Sub

Private Sub cbx_envia_cic_Change()

End Sub

Private Sub cbx_estado_civil_Change()
txt_estado_credito = Empty
cbx_conyuge_inf_com = Empty

'''' CHEQUEA RUT DE CONYUGE QUE EXISTA COMO CONSULTA SINACOFI y en todas las bases segun politica

    If cbx_estado_civil = "Casado" Or cbx_estado_civil = "Separado De Hecho" Then
        
    If txt_rut_conyuge <> "" Then
       
           Call conectarBD
    
            ssql = "select rut, mora, protesto,infraccion_prev" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_conyuge & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
                If rst.EOF Then
                   MsgBox "Debe Consultar el Rut del Conyuge en Sinacofi"
                Else
                
                    If rst!mora = "No Cumple" Or rst!protesto = "No Cumple" Or rst!infraccion_prev = "No Cumple" Then
                       cbx_conyuge_inf_com = "No Cumple"
                    Else
                    
                        ssql = "select rut" _
                        & " from tbl_micro_resumen_prime_noprime" _
                        & " where rut = '" & txt_rut_conyuge & "'" _
                        & " and max_mora_total >0"
                        Set rst = cnn.Execute(ssql, , adCmdText)
                    
                        If Not rst.EOF Then
                           cbx_conyuge_inf_com = "No Cumple"
                           
                           Else
                                ssql = "select rut_numerico" _
                                & " from TBL_MICRO_RESUMEN_CRITICAL" _
                                & " where rut_numerico = '" & txt_rut_conyuge & "' and marca_regular =1"
                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                If Not rst.EOF Then
                                    cbx_conyuge_inf_com = "No Cumple"
                                                                        
                                    Else
                                        ssql = "select rut" _
                                        & " from TBL_MICRO_sbif" _
                                        & " where rut = '" & txt_rut_conyuge & "'"
                                        Set rst = cnn.Execute(ssql, , adCmdText)
                                                                                                            
                                          If Not rst.EOF Then
                                            cbx_conyuge_inf_com = "No Cumple"
                                                
                                            Else
                                                ssql = "select rut" _
                                                & " from TBL_MICRO_riesgo_renegociado" _
                                                & " where rut = '" & txt_rut_conyuge & "'"
                        
                                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                
                                                    If Not rst.EOF Then
                                                        cbx_conyuge_inf_com = "No Cumple"
                                                      
                                                      Else
                                                            ssql = "select rut10" _
                                                            & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                            & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_conyuge & "'" _
                                                            & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                                                If Not rst.EOF Then
                                                                cbx_conyuge_inf_com = "No Cumple"
                                                        
                                                                Else
                                                                    ssql = "select rut10" _
                                                                    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                                    & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_conyuge & "'" _
                                                                    & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                                
                                                                    If Not rst.EOF Then
                                                                    cbx_conyuge_inf_com = "No Cumple"
                                                                    
                                                                Else
                                                                    cbx_conyuge_inf_com = "Cumple"
                                                    
                                                            End If
                                                    
                                                        End If
                                                        
                                                    End If
                                            
                                            End If
                                    
                                    End If
                            
                            End If
                    
                    End If
                    
                End If
                
        Else
       
       MsgBox "Debe Ingresar Rut Conyuge"
    
    End If
       
    
'End If
Else
cbx_conyuge_inf_com = "Cumple"
End If
End Sub



Private Sub cbx_evaluador_Change()

End Sub

Private Sub cbx_forma_verif_dir_part_Change()
txt_estado_credito = Empty

End Sub

Private Sub cbx_grado_formalidad_Change()

txt_estado_credito = Empty


End Sub

Private Sub cbx_mora_sbif_Change()

txt_estado_credito = Empty


If cbx_mora_sbif = "Cumple" Or cbx_mora_sbif = "No Bancarizado" Then
  txt_r_mora_sbif = "A"
    lbl_mora_sbif.BackColor = &HC000&
       lbl_mora_sbif.ForeColor = &H8000000E  'blanco
     '  lbl_mora_sbif.BorderStyle = fmBorderStyleSingle 'con bordes
        
  Else: txt_r_mora_sbif = "R"
  lbl_mora_sbif.BackColor = &HFF&       'rojo
    lbl_mora_sbif.ForeColor = &H8000000E  'blanco
    'lbl_mora_sbif.BorderStyle = fmBorderStyleSingle 'con bordes
End If
End Sub

Private Sub cbx_Mora_Total_Sbif_Change()

txt_estado_credito = Empty


If cbx_Mora_Total_Sbif = "Cumple" Or cbx_Mora_Total_Sbif = "No Bancarizado" Then
  cbx_r_Mora_Total_Sbif = "A"
    
          lbl_Mora_Total_Sbif.BackColor = &HC000&   ' verde
        lbl_Mora_Total_Sbif.ForeColor = &H8000000E  'blanco
        'lbl_Mora_Total_Sbif.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else: cbx_r_Mora_Total_Sbif = "R"
      lbl_Mora_Total_Sbif.BackColor = &HFF&       'rojo
    lbl_Mora_Total_Sbif.ForeColor = &H8000000E  'blanco
    'lbl_Mora_Total_Sbif.BorderStyle = fmBorderStyleSingle 'con bordes
End If
End Sub

Private Sub cbx_Mora_Total_Sbif_indirecta_Change()

txt_estado_credito = Empty


If cbx_Mora_Total_Sbif_indirecta = "Cumple" Or cbx_Mora_Total_Sbif_indirecta = "No Bancarizado" Then
  cbx_r_Mora_Total_Sbif_indirecta = "A"
    '      lbl_venc_cast_SBIF.BackColor = &HC000&   ' verde
    '    lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
        'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
        
  Else
    cbx_r_Mora_Total_Sbif_indirecta = "R"
      'lbl_venc_cast_SBIF.BackColor = &HFF&       'rojo
     'lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
    'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
  
End If

End Sub

Private Sub cbx_morosidades_Change()

txt_estado_credito = Empty


If cbx_morosidades = "Cumple" Or cbx_morosidades = "No Cliente Banco" Or cbx_morosidades = "No Bancarizado" Then
  
  txt_r_morosidad = "A"
  lbl_morosidades.BackColor = &HC000&
  lbl_morosidades.ForeColor = &H8000000E  'blanco
  'lbl_morosidades.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else: txt_r_morosidad = "R"
  lbl_morosidades.BackColor = &HFF&       'rojo
  lbl_morosidades.ForeColor = &H8000000E  'blanco
  'lbl_morosidades.BorderStyle = fmBorderStyleSingle 'con bordes

End If

End Sub

Private Sub cbx_numero_acreedores_Change()

'------- ACREEDORES---------------------------------
  
txt_estado_credito = Empty
txt_r_n_acreedores = Empty

If cbx_numero_acreedores <= 3 Or cbx_numero_acreedores = 0 Then
  txt_r_n_acreedores = "A"
  Estado_Resolucion_Final.txt_r_f_n_acreedor = "A"
  
        lbl_numero_acreedores.BackColor = &HC000&
        lbl_numero_acreedores.ForeColor = &H8000000E  'blanco
  Else
  txt_r_n_acreedores = "R"
  Estado_Resolucion_Final.txt_r_f_n_acreedor = "R"
  lbl_numero_acreedores.BackColor = &HFF&       'rojo
  lbl_numero_acreedores.ForeColor = &H8000000E  'blanco
  
End If

End Sub

Private Sub cbx_pregunta_comercial_Change()

cbx_Destino_Construccion = Empty
cbx_años_vehiculo_auto = Empty
txt_r_accion = Empty
cbx_Accion = Empty
cbx_tipo_vehiculo = Empty
cbx_años_vehiculo_bus = Empty
txt_plazo_credito = Empty
txt_r_plazo_credito = Empty

txt_r_años_vehiculo = Empty
cbx_destino_terreno = Empty
cbx_Destino_Construccion = Empty
cbx_Destino_Credito_Vivienda = Empty
cbx_Destino_Credito_Nec_Per = Empty
cbx_Destino_Credito_Cap_Tra = Empty
cbx_Destino_Credito_Act_Fij = Empty

If cbx_pregunta_comercial = "No" Then
    
    cbx_años_vehiculo_auto.Locked = True
    txt_r_accion.Locked = True
    cbx_Accion.Locked = True
    cbx_tipo_vehiculo.Locked = True
    cbx_años_vehiculo_bus.Locked = True
    txt_plazo_credito.Locked = True
    txt_r_plazo_credito.Locked = True
        
    txt_r_años_vehiculo.Locked = True
    cbx_destino_terreno.Locked = True
    cbx_Destino_Construccion.Locked = True
    cbx_Destino_Credito_Vivienda.Locked = True
    cbx_Destino_Credito_Nec_Per.Locked = True
    cbx_Destino_Credito_Cap_Tra.Locked = True
    cbx_Destino_Credito_Act_Fij.Locked = True


ElseIf cbx_pregunta_comercial = "Si" Then
    

    cbx_años_vehiculo_auto.Locked = False
    txt_r_accion.Locked = False
    cbx_Accion.Locked = False
    cbx_tipo_vehiculo.Locked = False
    cbx_años_vehiculo_bus.Locked = False
    txt_plazo_credito.Locked = False
    txt_r_plazo_credito.Locked = False
        
    txt_r_años_vehiculo.Locked = False
    cbx_destino_terreno.Locked = False
    cbx_Destino_Construccion.Locked = False
    cbx_Destino_Credito_Vivienda.Locked = False
    cbx_Destino_Credito_Nec_Per.Locked = False
    cbx_Destino_Credito_Cap_Tra.Locked = False
    cbx_Destino_Credito_Act_Fij.Locked = False

End If
        

End Sub

Private Sub cbx_pregunta_consumo_Change()
        
        cbx_Accion_consumo = Empty
        txt_plazo_credito_consumo = Empty

If cbx_pregunta_consumo = "No" Then
    cbx_Accion_consumo.Locked = True
    txt_plazo_credito_consumo.Locked = True
    
    ElseIf cbx_pregunta_consumo = "Si" Then
        cbx_Accion_consumo.Locked = False
        txt_plazo_credito_consumo.Locked = False

    
    Else
        cbx_Accion_consumo = Empty
        txt_plazo_credito_consumo = Empty
        
    
End If
    
End Sub

Private Sub cbx_protestos_Change()

txt_estado_credito = Empty


If cbx_protestos = "Cumple" Or cbx_protestos = "No Cliente Banco" Or cbx_protestos = "No Bancarizado" Then
  txt_r_protestos = "A"
  lbl_protestos.BackColor = &HC000&
  lbl_protestos.ForeColor = &H8000000E  'blanco
  'lbl_protestos.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else: txt_r_protestos = "R"
  lbl_protestos.BackColor = &HFF&       'rojo
  lbl_protestos.ForeColor = &H8000000E  'blanco
  'lbl_protestos.BorderStyle = fmBorderStyleSingle 'con bordes
End If

End Sub

Private Sub cbx_r_venc_cast_SBIF_indirecta_Change()

End Sub





Private Sub cbx_telef_verif_Change()
txt_estado_credito = Empty

If cbx_telef_verif = "Si" Then
   
   txt_r_telefono_verificado = "A"
   Estado_Resolucion_Final.txt_r_f_telefono_verificado = "A"
   
   Else
      txt_r_telefono_verificado = "R"
      Estado_Resolucion_Final.txt_r_f_telefono_verificado = "R"
   
End If


End Sub

Private Sub cbx_tiene_score_Click()

End Sub

Private Sub cbx_tipo_cliente_Change()

txt_estado_credito = Empty
'txt_r_formalidad_negocio = Empty

txt_r_cbx_actividad_economica_informal_oficio = Empty



cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_formal_servicio.Visible = False
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = False

If cbx_tipo_cliente = "FORMALES" Then
txt_r_formalidad_negocio = "A"
txt_r_cbx_actividad_economica_informal_oficio = "A"

cbx_actividad_economica_formal.Visible = True
cbx_actividad_economica_formal_servicio.Visible = False
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = False

ElseIf cbx_tipo_cliente = "FORMAL SERVICIO O PRODUCCION" Then
txt_r_formalidad_negocio = "A"
txt_r_cbx_actividad_economica_informal_oficio = "A"

cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_formal_servicio.Visible = True
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = False

ElseIf cbx_tipo_cliente = "SEMIFORMALES" Then
txt_r_formalidad_negocio = "A"
txt_r_cbx_actividad_economica_informal_oficio = "A"

cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_formal_servicio.Visible = False
cbx_actividad_economica_semiformal.Visible = True
cbx_actividad_economica_informal_oficio.Visible = False

ElseIf cbx_tipo_cliente = "INFORMALES(Oficio)" Then
txt_r_formalidad_negocio = "A"
txt_r_cbx_actividad_economica_informal_oficio = "A"

cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_formal_servicio.Visible = False
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = True

End If

If txt_bancarizado_politica = "No" And (cbx_tipo_cliente <> "FORMALES" And cbx_tipo_cliente <> "FORMAL SERVICIO O PRODUCCION") Then
txt_r_formalidad_negocio = "R"
txt_r_cbx_actividad_economica_informal_oficio = "R"

End If

End Sub

Private Sub cbx_tipo_excepcion_Change()
    txt_marca_exc = 0
If cbx_tipo_excepcion = "No" Then
  cbx_ejecutivo_excepcion = Empty
  cbx_ejecutivo_excepcion.Visible = False
  lbl_ejecutivo_evaluador.Visible = False
Else
    txt_marca_exc = 1
    cbx_ejecutivo_excepcion.Visible = True
    lbl_ejecutivo_evaluador.Visible = True
End If
End Sub

Private Sub cbx_tipo_ME_Change()

txt_estado_credito = Empty


End Sub

Private Sub cbx_tipo_vehiculo_Change()

cbx_años_vehiculo_bus = Empty
cbx_años_vehiculo_auto = Empty
txt_r_años_vehiculo = Empty

If cbx_tipo_vehiculo = "Taxi" Or cbx_tipo_vehiculo = "RadioTaxi" Or cbx_tipo_vehiculo = "Colectivo" Then
   cbx_años_vehiculo_auto.Visible = True
   cbx_años_vehiculo_bus.Visible = False
Else
   cbx_años_vehiculo_auto.Visible = False
   cbx_años_vehiculo_bus.Visible = True
End If
End Sub

Private Sub cbx_valor_evaluo_bien_raiz_Change()
txt_estado_credito = Empty
txt_r_valor_bien_raiz = cbx_valor_evaluo_bien_raiz

End Sub

Private Sub cbx_valor_evaluo_bien_raiz_no_Change()

txt_estado_credito = Empty
txt_r_valor_bien_raiz = cbx_valor_evaluo_bien_raiz_no

End Sub

Private Sub cbx_vehiculos_propios_Change()

txt_estado_credito = Empty
cbx_r_vehiculos_propios = Empty

cbx_acreditacion_vehiculo.Visible = True
cbx_acreditacion_vehiculo_no.Visible = False

txt_r_cbx_vehiculos_propios = "A"

    If cbx_vehiculos_propios = "No" Then
        cbx_acreditacion_vehiculo.Visible = False
        cbx_acreditacion_vehiculo_no.Visible = True
    End If
    
'If txt_bancarizado_politica = "Si" Then
If txt_bancarizado_politica = "No" And cbx_vehiculos_propios = "Si" Then
    txt_r_cbx_vehiculos_propios = "A"
ElseIf txt_bancarizado_politica = "No" And cbx_vehiculos_propios = "No" Then
    txt_r_cbx_vehiculos_propios = "R"
End If

End Sub

Private Sub cbx_venc_cast_SBIF_Change()

txt_estado_credito = Empty


If cbx_venc_cast_SBIF = "Cumple" Or cbx_venc_cast_SBIF = "No Bancarizado" Then
  cbx_r_venc_cast_SBIF = "A"
          lbl_venc_cast_SBIF.BackColor = &HC000&   ' verde
        lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
        'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
        
  Else: cbx_r_venc_cast_SBIF = "R"
      lbl_venc_cast_SBIF.BackColor = &HFF&       'rojo
    lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
    'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
  
End If

End Sub



Private Sub cbx_venc_cast_SBIF_indirecta_Change()

txt_estado_credito = Empty


If cbx_venc_cast_SBIF_indirecta = "Cumple" Or cbx_venc_cast_SBIF_indirecta = "No Bancarizado" Then
  cbx_r_venc_cast_SBIF_indirecta = "A"
    '      lbl_venc_cast_SBIF.BackColor = &HC000&   ' verde
    '    lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
        'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
        
  Else
    cbx_r_venc_cast_SBIF_indirecta = "R"
      'lbl_venc_cast_SBIF.BackColor = &HFF&       'rojo
     'lbl_venc_cast_SBIF.ForeColor = &H8000000E  'blanco
    'lbl_venc_cast_SBIF.BorderStyle = fmBorderStyleSingle 'con bordes
  
End If


End Sub

Private Sub cbx_visita_eje_Change()
txt_estado_credito = Empty

If cbx_visita_eje = "Si" Then
   
   txt_r_visita_ejecutivo = "A"
   Estado_Resolucion_Final.txt_r_f_visita_ejecutivo = "A"
   
   Else
      txt_r_visita_ejecutivo = "ZG"
      Estado_Resolucion_Final.txt_r_f_visita_ejecutivo = "ZG"
   
End If


End Sub


Private Sub cmd_cerrar_caso_volver_menu_Click()

irespuesta = MsgBox("¿Comezará Una Nueva Evaluacion Esta Seguro?", vbYesNo)
        
        If irespuesta = vbYes Then

Unload Ficha_Cliente_Micro

Menu_Principal_Micro.Show
End If
End Sub

Private Sub cmd_grabar_Solicitud_Click()

            
'''''mail de viviana manriquez gerente de riesgo 25-08-2013
'''''---****** POLITICA NO BANCARIZADO
If txt_bancarizado_politica = "Si" Then
    txt_ESTADO_politica_bancarizado_new.BackColor = &HE0E0E0
    txt_ESTADO_politica_bancarizado_new = "N/A"

    ElseIf txt_bancarizado_politica = "No" And txt_r_edad = "A" And txt_r_formalidad_negocio = "A" And txt_r_meses_antiguedad = "A" And txt_r_cbx_antiguedad_rubro = "A" And txt_r_cbx_actividad_economica_informal_oficio = "A" And (txt_r_cbx_bien_Raiz = "A" Or txt_r_cbx_vehiculos_propios = "A") And txt_r_predictor_score_dicom = "A" Then
    txt_ESTADO_politica_bancarizado_new = "A"

ElseIf txt_bancarizado_politica = "No" And txt_r_edad = "A" And txt_r_formalidad_negocio = "A" And txt_r_meses_antiguedad = "A" And txt_r_cbx_antiguedad_rubro = "A" And txt_r_cbx_actividad_economica_informal_oficio = "A" And (txt_r_cbx_bien_Raiz = "A" Or txt_r_cbx_vehiculos_propios = "A") And txt_r_predictor_score_dicom = "ZG" Then
  txt_ESTADO_politica_bancarizado_new = "ZG"

    Else
    txt_ESTADO_politica_bancarizado_new = "R"
End If
            
'''''

If txt_marca_exc = 1 And cbx_ejecutivo_excepcion <> "" Or txt_marca_exc = 0 Then

cmd_Menu_Evaluacion.Enabled = False
cmd_grabar_Solicitud.Enabled = False

If txt_rut_cliente <> "" And txt_dv <> "" Or (cbx_Accion = "Construccion" And (cbx_tipo_vehiculo = "" And txt_plazo_credito = "" And cbx_años_vehiculo_bus = "")) And _
txt_n_carpeta_tributaria <> "" _
And cbx_codigo_sucursal <> "" _
And cbx_cod_ejecutivo <> "" _
And cbx_tipo_cliente <> "" _
And txt_antiguedad_meses <> "" _
And cbx_antiguedad_rubro <> "" _
And cbx_Dir_Comercial_Verif <> "" _
And cbx_visita_eje <> "" _
And cbx_telef_verif <> "" _
And cbx_direc_Part_Verif <> "" _
And cbx_forma_verif_dir_part <> "" _
And cbx_estado_civil <> "" _
And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
And cbx_antecedentes_int_bancos <> "" _
And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
And (cbx_Destino_Credito_Nec_Per <> "" Or cbx_Destino_Credito_Act_Fij <> "" Or cbx_Destino_Credito_Cap_Tra <> "" Or cbx_Destino_Credito_Vivienda <> "") And txt_estado_credito <> "" _
And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And cbx_conyuge_inf_com <> "" And txt_antiguedad_banco <> "" Then

cmd_Menu_Evaluacion.Enabled = True



'PASO DE VARIABLES GLOBALES FICHA DEL CLIENTE
'------------------------------
rut_cliente_ficha = txt_rut_cliente
dv_cliente_ficha = txt_dv
n_carpeta_tributaria_ficha = txt_n_carpeta_tributaria
formalidad_negocio_ficha = cbx_tipo_cliente
score_dicom_ficha = txt_score_dicom
antiguedad_meses_ficha = txt_antiguedad_meses
cbx_codigo_sucursal_ficha = cbx_codigo_sucursal
cbx_cod_ejecutivo_ficha = cbx_cod_ejecutivo
cbx_Accion_ficha = cbx_Accion
r_tipo_credito_destino_ficha = r_tipo_credito_destino
cbx_tipo_cliente_ficha = cbx_tipo_cliente

antiguedad_rubro_ficha = cbx_antiguedad_rubro
r_actividad_ficha = txt_r_actividad_economica
Dir_Comercial_Verif_ficha = cbx_Dir_Comercial_Verif
visita_eje_ficha = cbx_visita_eje
telef_verif_ficha = cbx_telef_verif
direc_Part_Verif_ficha = cbx_direc_Part_Verif
forma_verif_dir_part_ficha = cbx_forma_verif_dir_part
estado_civil_ficha = cbx_estado_civil
bien_Raiz_ficha = cbx_bien_Raiz
Acreditacion_Bien_Raiz_ficha = r_txt_acreditacion_bien_raiz
valor_bien_raiz_ficha = txt_r_valor_bien_raiz
vehiculos_propios_ficha = cbx_vehiculos_propios
acreditacion_vehiculo_ficha = txt_r_acreditacion_vehiculo
antecedentes_int_bancos_ficha = cbx_antecedentes_int_bancos
morosidades_ficha = cbx_morosidades
protestos_ficha = cbx_protestos
boletin_laboral_ficha = cbx_boletin_laboral
mora_sbif_ficha = cbx_mora_sbif
venc_cast_SBIF_ficha = cbx_venc_cast_SBIF
Mora_Total_Sbif_ficha = cbx_Mora_Total_Sbif
numero_acreedores_ficha = cbx_numero_acreedores
score_dicom_ficha = txt_score_dicom
grado_formalidad_ficha = cbx_grado_formalidad
tipo_ME_ficha = cbx_tipo_ME
n_trabajador_familia_ficha = txt_n_trabajador_familia
envia_cic_ficha = cbx_envia_cic
Credito_Fogape_ficha = cbx_credito_fogape
Campana_ficha = txt_campana
conyuge_inf_com = cbx_conyuge_inf_com


'ESTADOS DE FLITROS DE RECHAZADOS FICHA CLIENTE
Estado_AIB_ficha = txt_r_aib
Estado_Morosidad_ficha = txt_r_morosidad
Estado_Protestos_ficha = txt_r_protestos
Estado_Boletin_Laboral_ficha = txt_r_boletin_laboral
Estado_Meses_Antiguedad_ficha = txt_r_meses_antiguedad
estado_mora_sbif_ficha = txt_r_mora_sbif
estado_venc_cast_SBIF_ficha = cbx_r_venc_cast_SBIF
estado_Mora_Total_Sbif_ficha = cbx_r_Mora_Total_Sbif
estado_numero_acreedores_ficha = txt_r_n_acreedores
estado_score_dicom_ficha = txt_r_predictor_score_dicom
estado_credito_ficha = txt_estado_credito

'------------------------

' La conexión a la base de datos

    Call conectarBD


ssql = "INSERT INTO TBL_MICRO_FICHA_CLIENTE " _
    & "([Rut_Cliente], [Dv], [N_Carpeta_Tributaria],[Id_Sucursal],[Codigo_Ejecutivo],[Destino_Credito]," _
    & " [Formalidad_Negocio],[Tipo_Credito_Destino],[Antiguedad_Negocio],[estado_antiguedad_negocio],[Certificacion_Antiguedad_Rubro],[r_actividad]," _
    & " [Dir_Comercial_Verif],[visita_eje],[telef_verif],[direc_Part_Verif],[forma_verif_dir_part]," _
    & " [estado_civil],[bien_Raiz],[Acreditacion_Bien_Raiz],[valor_bien_raiz],[vehiculos_propios]," _
    & " [acreditacion_vehiculo],[antecedentes_int_bancos],[morosidades],[protestos],[boletin_laboral]," _
    & " [mora_sbif], [venc_cast_SBIF],[Mora_Total_Sbif],[numero_acreedores],[score_dicom]," _
    & " [grado_formalidad],[tipo_ME],[n_trabajador_familia],[envia_cic],[Estado_AIB],[Estado_Morosidad],[Estado_Protestos]," _
    & " [Estado_Boletin_Laboral],[estado_mora_sbif],[estado_venc_cast_SBIF],[estado_Mora_Total_Sbif],[estado_numero_acreedores]," _
    & " [estado_score_dicom],[estado_credito],[fecha_ingreso],[hora_ingreso],[Credito_Fogape],[Campana],[conyuge_inf_com],[Estado_conyuge_inf_com],[antiguedad_bco],[Tipo_Excepcion],[Ejecutivo_Excepcion],[Evaluador_Sucursal],[Credito_Fogain],[plazo_credito],[Credito_Comercial],[Credito_Consumo],[Accion_Consumo],[Plazo_Credito_Consumo],[bancarizado_FC_PNB],[r_formalidad_negocio_PNB],[r_cbx_antiguedad_rubro_PNB],[r_cbx_actividad_economica_PNB],[r_cbx_bien_Raiz_PNB],[r_cbx_vehiculos_propios_PNB],[ESTADO_politica_bancarizado_new_PNB])" _
    & " VALUES (('" & rut_cliente_ficha & "') " _
    & ",('" & dv_cliente_ficha & "') , ('" & n_carpeta_tributaria_ficha & "') ,('" & cbx_codigo_sucursal_ficha & "') " _
    & ",substring('" & cbx_cod_ejecutivo_ficha & "',1,4), ('" & cbx_Accion_ficha & "') , ('" & cbx_tipo_cliente_ficha & "') " _
    & ",('" & r_tipo_credito_destino_ficha & "') ,('" & antiguedad_meses_ficha & "'),('" & txt_r_meses_antiguedad & "')" _
    & ",('" & antiguedad_rubro_ficha & "') ,('" & r_actividad_ficha & "'), ('" & Dir_Comercial_Verif_ficha & "'), ('" & visita_eje_ficha & "') " _
    & ",('" & telef_verif_ficha & "'),('" & direc_Part_Verif_ficha & "'), ('" & forma_verif_dir_part_ficha & "')" _
    & ",('" & estado_civil_ficha & "'),('" & bien_Raiz_ficha & "'), ('" & Acreditacion_Bien_Raiz_ficha & "'), ('" & valor_bien_raiz_ficha & "')" _
    & ",('" & vehiculos_propios_ficha & "'), ('" & acreditacion_vehiculo_ficha & "'), ('" & antecedentes_int_bancos_ficha & "') " _
    & ",('" & morosidades_ficha & "'), ('" & protestos_ficha & "'), ('" & boletin_laboral_ficha & "'), ('" & mora_sbif_ficha & "')" _
    & ",('" & venc_cast_SBIF_ficha & "'), ('" & Mora_Total_Sbif_ficha & "'), ('" & numero_acreedores_ficha & "'), ('" & score_dicom_ficha & "')" _
    & ",('" & grado_formalidad_ficha & "'),('" & tipo_ME_ficha & "'), ('" & n_trabajador_familia_ficha & "'), ('" & envia_cic_ficha & "')" _
    & ",('" & Estado_AIB_ficha & "'),('" & Estado_Morosidad_ficha & "'),('" & Estado_Protestos_ficha & "')" _
    & ",('" & Estado_Boletin_Laboral_ficha & "'),('" & estado_mora_sbif_ficha & "'),('" & estado_venc_cast_SBIF_ficha & "')" _
    & ",('" & estado_Mora_Total_Sbif_ficha & "'),('" & estado_numero_acreedores_ficha & "'),('" & estado_score_dicom_ficha & "')" _
    & ",('" & estado_credito_ficha & "'), ('" & txt_fecha_ingreso_compara & "'), ('" & txt_hora_actual & "'),('" & cbx_credito_fogape & "'),('" & txt_campana & "'),('" & cbx_conyuge_inf_com & "'),('" & txt_r_conyuge_inf_com & "'),('" & txt_antiguedad_banco & "'),('" & cbx_tipo_excepcion & "'),('" & cbx_ejecutivo_excepcion & "'),('" & cbx_evaluador & "'),('" & cbx_credito_fogain & "'),('" & txt_plazo_credito & "'), ('" & cbx_pregunta_comercial & "'),('" & cbx_pregunta_consumo & "'),('" & cbx_Accion_consumo & "'),('" & txt_plazo_credito_consumo & "'),('" & txt_bancarizado_politica & "'),('" & txt_r_formalidad_negocio & "'),('" & txt_r_cbx_antiguedad_rubro & "'),('" & txt_r_cbx_actividad_economica_informal_oficio & "'),('" & txt_r_cbx_bien_Raiz & "'),('" & txt_r_cbx_vehiculos_propios & "'),('" & txt_ESTADO_politica_bancarizado_new & "'))"
    
    cnn.Execute ssql
    
  Else
MsgBox " Faltan Datos Que Ingresar En Ficha De Cliente"
    
End If

  Else
    MsgBox "Falta Ingresar Datos De Excepcion", vbCritical
  
  End If
     
If txt_ESTADO_politica_bancarizado_new = "No Cumple Politica Bancarizado" Then
    cmd_Menu_Evaluacion.Enabled = False
    cmd_grabar_Solicitud.Enabled = False
End If

End Sub

Private Sub cmd_imprimir1_meto_ac_Click()
Ficha_Cliente_Micro.PrintForm
End Sub


Private Sub cmd_Menu_Evaluacion_Click()

txt_rut_cliente.Enabled = False
txt_dv.Enabled = False


If txt_rut_cliente <> "" Or (cbx_Accion = "Construccion" And (cbx_tipo_vehiculo = "" And txt_plazo_credito = "" And cbx_años_vehiculo_bus = "")) And txt_dv <> "" And _
txt_n_carpeta_tributaria <> "" _
And cbx_codigo_sucursal <> "" _
And cbx_cod_ejecutivo <> "" _
And cbx_tipo_cliente <> "" _
And txt_antiguedad_meses <> "" _
And cbx_antiguedad_rubro <> "" _
And cbx_Dir_Comercial_Verif <> "" _
And cbx_visita_eje <> "" _
And cbx_telef_verif <> "" _
And cbx_direc_Part_Verif <> "" _
And cbx_forma_verif_dir_part <> "" _
And cbx_estado_civil <> "" _
And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
And (cbx_actividad_economica_informal_oficio <> "" Or cbx_actividad_economica_formal_servicio <> "" Or cbx_actividad_economica_semiformal <> "" Or cbx_actividad_economica_formal <> "") _
And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
And cbx_antecedentes_int_bancos <> "" _
And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
And (cbx_Destino_Credito_Nec_Per <> "" Or cbx_Destino_Credito_Act_Fij <> "" Or cbx_Destino_Credito_Cap_Tra <> "" Or cbx_Destino_Credito_Vivienda <> "") _
And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And cbx_conyuge_inf_com <> "" Then

Ficha_Cliente_Micro.Hide


Evaluacion_Perfil.Show

'And cbx_Accion <> "" _

Else
MsgBox " Faltan Datos Que Ingresar En Ficha De Cliente"
End If
End Sub

Private Sub cmd_salir_sistema_Click()
    Unload Ficha_Cliente_Micro
    Menu_Principal_Micro.Show
    
End Sub


Private Sub cmd_verificacion_ingreso_anteced_Click()

If cbx_pregunta_comercial = "Si" Then


''IFORMES COMERCIALES AVAL  CUANDO ES BLANCO
If txt_rut_aval = "" Then
cbx_aval_inf_com = "N/A"
txt_r_aval_inf_com = "A"
End If

''''''' Paso de marca si cliente tiene aval y/o conyuge

If txt_rut_cliente <> "" And txt_rut_conyuge = "" And txt_rut_aval = "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 0
   Estado_Resolucion_Final.txt_marca_aval = 0
End If


If txt_rut_conyuge <> "" And txt_rut_aval = "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 1
Else
    Estado_Resolucion_Final.txt_marca_aval = 0
End If


If txt_rut_aval <> "" And txt_rut_conyuge = "" Then
   Estado_Resolucion_Final.txt_marca_aval = 1
Else
    Estado_Resolucion_Final.txt_marca_conyuge = 0
End If

If txt_rut_conyuge <> "" And txt_rut_aval <> "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 2
   Estado_Resolucion_Final.txt_marca_aval = 2
End If



''''''''''''''''vERIFICIACION DE EDAD CLIENTE MINIMA'''''''''''''

If txt_bancarizado_politica = "Si" Then

    If txt_edad = "" Then
    txt_r_edad = ""
    End If

'''Edades minima

    If txt_score_dicom > 0 And txt_edad < 21 Then
        Estado_Resolucion_Final.txt_r_f_edad = "R"
        Ficha_Cliente_Micro.txt_r_edad = "R"
    Else
        Estado_Resolucion_Final.txt_r_f_edad = "A"
        Ficha_Cliente_Micro.txt_r_edad = "A"
    End If


If (txt_score_dicom > 0 Or txt_score_dicom = 0) And txt_edad > 72 Then
        Estado_Resolucion_Final.txt_r_f_edad_maxima = "R"
        Ficha_Cliente_Micro.txt_r_edad = "R"
Else
   Estado_Resolucion_Final.txt_r_f_edad_maxima = "A"
   Ficha_Cliente_Micro.txt_r_edad = "A"
End If

ElseIf txt_bancarizado_politica = "No" And txt_edad >= 35 Then
            Estado_Resolucion_Final.txt_r_f_edad = "A"
            Ficha_Cliente_Micro.txt_r_edad = "A"
Else
            Estado_Resolucion_Final.txt_r_f_edad_maxima = "R"
            Ficha_Cliente_Micro.txt_r_edad = "R"
End If

''''' indice de riesgo 0 debe tener a lo menos 40 años el cliente
'''''Edades maxima
''''cliente
'''conyuge
'-------------------------------------------

cmd_grabar_Solicitud.Enabled = False

MsgBox "Presione OK. Para Comenzar Proceso", vbInformation


cbx_envia_cic.Visible = False
cbx_envia_cic = Empty
cmd_Menu_Evaluacion.Enabled = False


''''**************************Empresa************
''***********************************************

If txt_rut_cliente > 48000000 Then
        If txt_rut_cliente <> "" And txt_rut_aval <> "" _
            And txt_dv <> "" _
            And txt_n_carpeta_tributaria <> "" _
            And cbx_codigo_sucursal <> "" _
            And cbx_cod_ejecutivo <> "" _
            And cbx_tipo_cliente <> "" _
            And txt_antiguedad_meses <> "" _
            And cbx_antiguedad_rubro <> "" _
            And cbx_Dir_Comercial_Verif <> "" _
            And cbx_visita_eje <> "" _
            And cbx_telef_verif <> "" _
            And cbx_direc_Part_Verif <> "" _
            And cbx_forma_verif_dir_part <> "" _
            And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
            And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
            And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
            And cbx_antecedentes_int_bancos <> "" _
            And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
            And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
            And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
            And (cbx_Destino_Credito_Nec_Per <> "" Or cbx_Destino_Credito_Act_Fij <> "" Or cbx_Destino_Credito_Cap_Tra <> "" Or cbx_Destino_Credito_Vivienda <> "") _
            And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And txt_antiguedad_banco <> "" And cbx_pregunta_consumo <> "" And cbx_pregunta_comercial <> "" Then
        Dim fec1
        Dim hora1
            fec1 = Format(Date, "yyyy/mm/dd")
            txt_fecha_ingreso_compara = fec1
        
            hora1 = hora
            txt_hora_actual = Time

           cbx_envia_cic.Visible = False
           lbl_enviar_cic.Visible = False
               If (txt_r_aib = "R" Or txt_r_morosidad = "R" Or txt_r_protestos = "R" Or txt_r_boletin_laboral = "R" Or txt_r_meses_antiguedad = "R" _
                    Or txt_r_n_acreedores = "R" Or txt_r_mora_sbif = "R" Or cbx_r_venc_cast_SBIF = "R" Or cbx_r_Mora_Total_Sbif = "R" Or _
                    txt_r_predictor_score_dicom = "R" Or txt_r_conyuge_inf_com = "R") Then

                    txt_estado_credito = "Rechazado"
                    Else
                    txt_estado_credito = "Aprobado"
                End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' TRAER CLIENTO NO CLIENTE DESDE BASE diaria operaciones
''''''''''''''''''''''''''''''''''''
Call conectarBD
    ssql = "select RUT" _
            & " FROM TBL_MICRO_MACA_CLIENTE_NO_CLIENTE_viG" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
 
          Else
          
            ssql = "select RUT" _
            & " FROM tbl_micro_cliente_antiguo_nuevo" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If Not rst.EOF Then
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
            Else
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "Si"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "Si"
            'End If
    End If
End If

    ''''' TRAE   MAX MORA Y MORA PROMEDIO
        
    ssql = "select rut,  flag_mm,  flag_mp from tbl_micro_resumen_prime_noprime" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
                Evaluacion_Perfil.cbx_historia_pago = "Sin Mora"
                Evaluacion_Perfil.cbx_mora_maxima = "Sin Mora"
            Else
                Evaluacion_Perfil.cbx_historia_pago = rst!flag_mp
                Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
            End If


            MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
            cmd_grabar_Solicitud.Enabled = True

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        
        cmd_grabar_Solicitud.Enabled = False

    End If
End If




'''******************************************************
'''*******************CLICLO PARA persona natural
'''******************************************************

If txt_rut_cliente < 48000000 Then
    If txt_rut_cliente <> "" _
        And txt_dv <> "" _
        And (((cbx_estado_civil = "Soltero" Or cbx_estado_civil = "Divorciado" Or cbx_estado_civil = "Viudo") And txt_rut_conyuge = "") _
        Or ((cbx_estado_civil = "Casado" Or cbx_estado_civil = "Separado De Hecho") And txt_rut_conyuge <> "") _
        Or (cbx_Accion = "Construccion" Or cbx_Accion = "Necesidades Personales" And (cbx_tipo_vehiculo = "" And txt_plazo_credito = "" And cbx_años_vehiculo_bus = ""))) _
        And txt_n_carpeta_tributaria <> "" _
        And cbx_codigo_sucursal <> "" _
        And cbx_cod_ejecutivo <> "" _
        And cbx_tipo_cliente <> "" _
        And txt_antiguedad_meses <> "" _
        And cbx_antiguedad_rubro <> "" _
        And cbx_Dir_Comercial_Verif <> "" _
        And cbx_visita_eje <> "" _
        And cbx_telef_verif <> "" _
        And cbx_direc_Part_Verif <> "" _
        And cbx_forma_verif_dir_part <> "" _
        And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
        And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
        And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
        And cbx_antecedentes_int_bancos <> "" _
        And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
        And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
        And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
        And (cbx_Destino_Credito_Nec_Per <> "" Or cbx_Destino_Credito_Act_Fij <> "" Or cbx_Destino_Credito_Cap_Tra <> "" Or cbx_Destino_Credito_Vivienda <> "") _
        And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And cbx_conyuge_inf_com <> "" And txt_antiguedad_banco <> "" And cbx_pregunta_consumo <> "" And cbx_pregunta_comercial <> "" Then

        fec1 = Format(Date, "yyyy/mm/dd")
        txt_fecha_ingreso_compara = fec1

        hora1 = hora
        txt_hora_actual = Time
        cbx_envia_cic.Visible = False
        lbl_enviar_cic.Visible = False

        If (txt_r_aib = "R" Or txt_r_morosidad = "R" Or txt_r_protestos = "R" Or txt_r_boletin_laboral = "R" Or txt_r_meses_antiguedad = "R" _
            Or txt_r_n_acreedores = "R" Or txt_r_mora_sbif = "R" Or cbx_r_venc_cast_SBIF = "R" Or cbx_r_Mora_Total_Sbif = "R" Or _
            txt_r_predictor_score_dicom = "R" Or txt_r_conyuge_inf_com = "R") Then

            txt_estado_credito = "Rechazado"
        Else
            txt_estado_credito = "Aprobado"
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''' TRAER CLIENTO NO CLIENTE DESDE BASE HISTORICA
''''''''''''''''''''''''''''''''''''

Call conectarBD

    ssql = "select RUT" _
            & " FROM TBL_MICRO_MACA_CLIENTE_NO_CLIENTE_viG" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
          Else

    ssql = "select RUT" _
            & " FROM tbl_micro_cliente_antiguo_nuevo" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
          Else
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "Si"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "Si"
        End If
    
    End If
'End If
    
        
    ''''' TRAE   MAX MORA Y MORA PROMEDIO
        
    ssql = "select rut,  flag_mm,  flag_mp from tbl_micro_resumen_prime_noprime" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
                Evaluacion_Perfil.cbx_historia_pago = "Sin Mora"
                Evaluacion_Perfil.cbx_mora_maxima = "Sin Mora"
            Else
                Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
                Evaluacion_Perfil.cbx_historia_pago = rst!flag_mp
                'Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
            End If
            

MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
cmd_grabar_Solicitud.Enabled = True

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        cmd_grabar_Solicitud.Enabled = False
    
    End If
'End If


'''primer IF
ElseIf (cbx_Accion = "Activo Fijo" And cbx_Destino_Credito_Act_Fij = "Vehiculo") Then
        If (cbx_tipo_vehiculo <> "" And (cbx_años_vehiculo_auto <> "" Or cbx_años_vehiculo_bus <> "") And txt_plazo_credito <> "") Then

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        cmd_grabar_Solicitud.Enabled = False
    End If
    'Else
    '    MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
    '    cmd_grabar_Solicitud.Enabled = False
'End If

ElseIf cbx_pregunta_consumo = "Si" And cbx_Accion_consumo <> "" And txt_plazo_credito_consumo <> "" Then
    MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
    cmd_grabar_Solicitud.Enabled = True

'Else
'        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
'        cmd_grabar_Solicitud.Enabled = False

End If


'*******************************
'comentado x Cristian Moreno A. segun mail de VIVIANA MANRIQUEZ 31-07-2013 Area de Riesgo.
'*******************************

'''********************************************************
''''******AVISO DEL TIPO DE CLIENTES PARA LA EVALUACION
'''********************************************************

'If txt_Cliente_Nuevo = "Si" And txt_bancarizado_politica = "No" Then

'    MsgBox "Este Cliente Quedara RECHAZADO por ser NO BANCARIZADO", vbCritical
    
'End If


'''''''''''''''''''''SOLO CONSUMO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Else


''IFORMES COMERCIALES AVAL  CUANDO ES BLANCO
If txt_rut_aval = "" Then
cbx_aval_inf_com = "N/A"
txt_r_aval_inf_com = "A"
End If

''''''' Paso de marca si cliente tiene aval y/o conyuge

If txt_rut_cliente <> "" And txt_rut_conyuge = "" And txt_rut_aval = "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 0
   Estado_Resolucion_Final.txt_marca_aval = 0
End If


If txt_rut_conyuge <> "" And txt_rut_aval = "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 1
Else
    Estado_Resolucion_Final.txt_marca_aval = 0
End If


If txt_rut_aval <> "" And txt_rut_conyuge = "" Then
   Estado_Resolucion_Final.txt_marca_aval = 1
Else
    Estado_Resolucion_Final.txt_marca_conyuge = 0
End If

If txt_rut_conyuge <> "" And txt_rut_aval <> "" Then
   Estado_Resolucion_Final.txt_marca_conyuge = 2
   Estado_Resolucion_Final.txt_marca_aval = 2
End If



''''''''''''''''vERIFICIACION DE EDAD CLIENTE MINIMA'''''''''''''


If txt_edad = "" Then
   txt_r_edad = ""
End If

'''Edades minima

If txt_score_dicom > 0 And txt_edad < 21 Then
   Estado_Resolucion_Final.txt_r_f_edad = "R"
   Ficha_Cliente_Micro.txt_r_edad = "R"
Else
   Estado_Resolucion_Final.txt_r_f_edad = "A"
   Ficha_Cliente_Micro.txt_r_edad = "A"
End If


If (txt_score_dicom > 0 Or txt_score_dicom = 0) And txt_edad > 72 Then
        Estado_Resolucion_Final.txt_r_f_edad_maxima = "R"
        Ficha_Cliente_Micro.txt_r_edad = "R"
Else
   Estado_Resolucion_Final.txt_r_f_edad_maxima = "A"
   Ficha_Cliente_Micro.txt_r_edad = "A"
End If

''''' indice de riesgo 0 debe tener a lo menos 40 años el cliente
'''''Edades maxima
''''cliente
'''conyuge
'-------------------------------------------

cmd_grabar_Solicitud.Enabled = False

MsgBox "Presione OK. Para Comenzar Proceso", vbInformation


cbx_envia_cic.Visible = False
cbx_envia_cic = Empty
cmd_Menu_Evaluacion.Enabled = False


''''**************************Empresa************
''***********************************************

If txt_rut_cliente > 48000000 Then
        If txt_rut_cliente <> "" And txt_rut_aval <> "" _
            And txt_dv <> "" _
            And txt_n_carpeta_tributaria <> "" _
            And cbx_codigo_sucursal <> "" _
            And cbx_cod_ejecutivo <> "" _
            And cbx_tipo_cliente <> "" _
            And txt_antiguedad_meses <> "" _
            And cbx_antiguedad_rubro <> "" _
            And cbx_Dir_Comercial_Verif <> "" _
            And cbx_visita_eje <> "" _
            And cbx_telef_verif <> "" _
            And cbx_direc_Part_Verif <> "" _
            And cbx_forma_verif_dir_part <> "" _
            And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
            And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
            And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
            And cbx_antecedentes_int_bancos <> "" _
            And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
            And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
            And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
            And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And txt_antiguedad_banco <> "" And cbx_pregunta_consumo <> "" And cbx_pregunta_comercial <> "" Then
        Dim fec11
        Dim hora11
            fec11 = Format(Date, "yyyy/mm/dd")
            txt_fecha_ingreso_compara = fec11
        
            hora11 = hora
            txt_hora_actual = Time

           cbx_envia_cic.Visible = False
           lbl_enviar_cic.Visible = False
               If (txt_r_aib = "R" Or txt_r_morosidad = "R" Or txt_r_protestos = "R" Or txt_r_boletin_laboral = "R" Or txt_r_meses_antiguedad = "R" _
                    Or txt_r_n_acreedores = "R" Or txt_r_mora_sbif = "R" Or cbx_r_venc_cast_SBIF = "R" Or cbx_r_Mora_Total_Sbif = "R" Or _
                    txt_r_predictor_score_dicom = "R" Or txt_r_conyuge_inf_com = "R") Then

                    txt_estado_credito = "Rechazado"
                    Else
                    txt_estado_credito = "Aprobado"
                End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' TRAER CLIENTO NO CLIENTE DESDE BASE diaria operaciones
''''''''''''''''''''''''''''''''''''
Call conectarBD
    ssql = "select RUT" _
            & " FROM TBL_MICRO_MACA_CLIENTE_NO_CLIENTE_viG" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
 
          Else
          
            ssql = "select RUT" _
            & " FROM tbl_micro_cliente_antiguo_nuevo" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If Not rst.EOF Then
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
            Else
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "Si"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "Si"
            'End If
    End If
End If

    ''''' TRAE   MAX MORA Y MORA PROMEDIO
        
    ssql = "select rut,  flag_mm,  flag_mp from tbl_micro_resumen_prime_noprime" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
                Evaluacion_Perfil.cbx_historia_pago = "Sin Mora"
                Evaluacion_Perfil.cbx_mora_maxima = "Sin Mora"
            Else
                Evaluacion_Perfil.cbx_historia_pago = rst!flag_mp
                Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
            End If


            MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
            cmd_grabar_Solicitud.Enabled = True

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        
        cmd_grabar_Solicitud.Enabled = False

    End If
End If




'''******************************************************
'''*******************CLICLO PARA persona natural
'''******************************************************

If txt_rut_cliente < 48000000 Then
    If txt_rut_cliente <> "" _
        And txt_dv <> "" _
        And (((cbx_estado_civil = "Soltero" Or cbx_estado_civil = "Divorciado" Or cbx_estado_civil = "Viudo") And txt_rut_conyuge = "") _
        Or ((cbx_estado_civil = "Casado" Or cbx_estado_civil = "Separado De Hecho") And txt_rut_conyuge <> "") _
        Or (cbx_Accion = "Construccion" Or cbx_Accion = "Necesidades Personales" And (cbx_tipo_vehiculo = "" And txt_plazo_credito = "" And cbx_años_vehiculo_bus = ""))) _
        And txt_n_carpeta_tributaria <> "" _
        And cbx_codigo_sucursal <> "" _
        And cbx_cod_ejecutivo <> "" _
        And cbx_tipo_cliente <> "" _
        And txt_antiguedad_meses <> "" _
        And cbx_antiguedad_rubro <> "" _
        And cbx_Dir_Comercial_Verif <> "" _
        And cbx_visita_eje <> "" _
        And cbx_telef_verif <> "" _
        And cbx_direc_Part_Verif <> "" _
        And cbx_forma_verif_dir_part <> "" _
        And cbx_bien_Raiz <> "" And (cbx_acred_bien_raiz <> "" Or cbx_acred_bien_raiz_no <> "") _
        And (cbx_valor_evaluo_bien_raiz <> "" Or cbx_valor_evaluo_bien_raiz_no <> "") And cbx_vehiculos_propios <> "" _
        And (cbx_acreditacion_vehiculo <> "" Or cbx_acreditacion_vehiculo_no <> "") _
        And cbx_antecedentes_int_bancos <> "" _
        And cbx_morosidades <> "" And cbx_protestos <> "" And cbx_boletin_laboral <> "" _
        And cbx_mora_sbif <> "" And cbx_venc_cast_SBIF <> "" And cbx_Mora_Total_Sbif <> "" And cbx_numero_acreedores <> "" _
        And txt_score_dicom <> "" And cbx_grado_formalidad <> "" And cbx_tipo_ME <> "" And txt_n_trabajador_familia <> "" _
        And cbx_credito_fogape <> "" And cbx_credito_fogain <> "" And cbx_conyuge_inf_com <> "" And txt_antiguedad_banco <> "" And cbx_pregunta_consumo <> "" And cbx_pregunta_comercial <> "" Then

        fec1 = Format(Date, "yyyy/mm/dd")
        txt_fecha_ingreso_compara = fec1

        hora1 = hora
        txt_hora_actual = Time
        cbx_envia_cic.Visible = False
        lbl_enviar_cic.Visible = False

        If (txt_r_aib = "R" Or txt_r_morosidad = "R" Or txt_r_protestos = "R" Or txt_r_boletin_laboral = "R" Or txt_r_meses_antiguedad = "R" _
            Or txt_r_n_acreedores = "R" Or txt_r_mora_sbif = "R" Or cbx_r_venc_cast_SBIF = "R" Or cbx_r_Mora_Total_Sbif = "R" Or _
            txt_r_predictor_score_dicom = "R" Or txt_r_conyuge_inf_com = "R") Then

            txt_estado_credito = "Rechazado"
        Else
            txt_estado_credito = "Aprobado"
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''' TRAER CLIENTO NO CLIENTE DESDE BASE HISTORICA
''''''''''''''''''''''''''''''''''''

Call conectarBD

    ssql = "select RUT" _
            & " FROM TBL_MICRO_MACA_CLIENTE_NO_CLIENTE_viG" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
          Else

    ssql = "select RUT" _
            & " FROM tbl_micro_cliente_antiguo_nuevo" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
          Else
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "Si"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "Si"
        End If
    
    End If
'End If
    
        
    ''''' TRAE   MAX MORA Y MORA PROMEDIO
        
    ssql = "select rut,  flag_mm,  flag_mp from tbl_micro_resumen_prime_noprime" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
                Evaluacion_Perfil.cbx_historia_pago = "Sin Mora"
                Evaluacion_Perfil.cbx_mora_maxima = "Sin Mora"
            Else
                Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
                Evaluacion_Perfil.cbx_historia_pago = rst!flag_mp
                'Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
            End If
            

MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
cmd_grabar_Solicitud.Enabled = True

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        cmd_grabar_Solicitud.Enabled = False
    
    End If
'End If


'''primer IF
ElseIf (cbx_Accion = "Activo Fijo" And cbx_Destino_Credito_Act_Fij = "Vehiculo") Then
        If (cbx_tipo_vehiculo <> "" And (cbx_años_vehiculo_auto <> "" Or cbx_años_vehiculo_bus <> "") And txt_plazo_credito <> "") Then

    Else
        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
        cmd_grabar_Solicitud.Enabled = False
    End If
    'Else
    '    MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
    '    cmd_grabar_Solicitud.Enabled = False
'End If

ElseIf cbx_pregunta_consumo = "Si" And cbx_Accion_consumo <> "" And txt_plazo_credito_consumo <> "" Then
    MsgBox "Presione OK. para Continuar con Evaluación ", vbInformation
    cmd_grabar_Solicitud.Enabled = True

'Else
'        MsgBox "Faltan Datos Que Ingresar ó Revise la Politica de Riesgo", vbCritical
'        cmd_grabar_Solicitud.Enabled = False

End If


'*******************************
'comentado x Cristian Moreno A. segun mail de VIVIANA MANRIQUEZ 31-07-2013 Area de Riesgo.
'*******************************

'''********************************************************
''''******AVISO DEL TIPO DE CLIENTES PARA LA EVALUACION
'''********************************************************

'If txt_Cliente_Nuevo = "Si" And txt_bancarizado_politica = "No" Then

'    MsgBox "Este Cliente Quedara RECHAZADO por ser NO BANCARIZADO", vbCritical
    
'End If

End If




End Sub

Private Sub cmd_volver_menu_princ_Click()
Unload Ficha_Cliente_Micro
Menu_Principal_Micro.Show

End Sub




Private Sub CommandButton1_Click()
Ficha_Cliente_Micro.Hide
Estado_Resolucion_Final.Show
End Sub

Private Sub fr_r_estado_gestion_Click()

End Sub



Private Sub txt_antiguedad_meses_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

txt_estado_credito = Empty

If txt_bancarizado_politica = "Si" Then

    If txt_score_dicom > 0 Then
        If (cbx_tipo_cliente = "FORMALES" Or cbx_tipo_cliente = "FORMAL SERVICIO O PRODUCCION") And txt_antiguedad_meses < 12 Then
        
            txt_r_meses_antiguedad = "R"
            lbl_antiguedad_meses.BackColor = &HFF&       'rojo
            lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco

            ElseIf (cbx_tipo_cliente = "SEMIFORMALES" Or cbx_tipo_cliente = "INFORMALES(Oficio)") And txt_antiguedad_meses < 24 Then
            
            txt_r_meses_antiguedad = "R"
            lbl_antiguedad_meses.BackColor = &HFF&       'rojo
            lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco
        
            Else
            txt_r_meses_antiguedad = "A"
            lbl_antiguedad_meses.BackColor = &HC000&   ' verde
            lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco
    
        End If
        
Else

    If (cbx_tipo_cliente = "FORMALES" Or cbx_tipo_cliente = "FORMAL SERVICIO O PRODUCCION") And (txt_antiguedad_meses < 24 Or txt_edad < 40) Then
        
        txt_r_meses_antiguedad = "R"
        lbl_antiguedad_meses.BackColor = &HFF&       'rojo
        lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco
        
    ElseIf (cbx_tipo_cliente = "SEMIFORMALES" Or cbx_tipo_cliente = "INFORMALES(Oficio)") And (txt_antiguedad_meses < 36 Or txt_edad < 40) Then
        
        txt_r_meses_antiguedad = "R"
        lbl_antiguedad_meses.BackColor = &HFF&       'rojo
        lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco
        
    Else
        txt_r_meses_antiguedad = "A"
        lbl_antiguedad_meses.BackColor = &HC000&   ' verde
        lbl_antiguedad_meses.ForeColor = &H8000000E  'blanco

    End If
End If

    ElseIf txt_bancarizado_politica = "No" And txt_antiguedad_meses > 48 Then
    txt_r_meses_antiguedad = "A"
    ElseIf txt_bancarizado_politica = "No" And txt_antiguedad_meses <= 48 Then
    txt_r_meses_antiguedad = "R"

End If

End Sub



Private Sub txt_campana_Change()


End Sub


Private Sub txt_credito_comercial_vigente_mora_Change()

If txt_credito_comercial_vigente_mora = "Si" Then
    txt_credito_comercial_vigente_mora.BackColor = &H8000&
    txt_credito_comercial_vigente_mora.ForeColor = &H8000000E  'blanco

  Else
    txt_credito_comercial_vigente_mora = "No"
    txt_credito_comercial_vigente_mora.BackColor = &HFF&
    txt_credito_comercial_vigente_mora.ForeColor = &H8000000E
End If
End Sub

Private Sub txt_cod_observacion_Change()


End Sub

Private Sub txt_cred_comer_cuota_Change()

End Sub

Private Sub txt_dv_aval_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    txt_estado_credito = Empty

    Dim I As Integer

    txt_dv_aval = UCase(txt_dv_aval)
    I = Len(txt_dv_aval)
    txt_dv_aval.SelStart = I
    
End Sub

Private Sub txt_dv_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    txt_estado_credito = Empty

    Dim I As Integer

    txt_dv = UCase(txt_dv)
    I = Len(txt_dv)
    txt_dv.SelStart = I

End Sub

Private Sub txt_dv_conyuge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    txt_estado_credito = Empty

    Dim I As Integer

    txt_dv_conyuge = UCase(txt_dv_conyuge)
    I = Len(txt_dv_conyuge)
    txt_dv_conyuge.SelStart = I
End Sub

Private Sub txt_edad_AfterUpdate()

txt_r_edad = "A"

If txt_bancarizado_politica = "No" And txt_edad >= "35" Then
    txt_r_edad = "A"
ElseIf txt_bancarizado_politica = "Si" Then
    txt_r_edad = "A"
    Else
    txt_r_edad = "R"
End If

End Sub

Private Sub txt_edad_aval_Change()

If txt_edad_aval = "" Then
   Estado_Resolucion_Final.txt_r_f_edad_aval = "N/A"
   Estado_Resolucion_Final.txt_r_f_edad_maxima_aval = "N/A"
End If

If txt_score_dicom_AVAL_aux > 0 And txt_edad_aval < 21 Then
    Estado_Resolucion_Final.txt_r_f_edad_aval = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_edad_aval = "A"
End If

If (txt_score_dicom_AVAL_aux > 0 Or txt_score_dicom_AVAL_aux = 0) And txt_edad_aval > 72 Then
    Estado_Resolucion_Final.txt_r_f_edad_maxima_aval = "R"
    Else
   Estado_Resolucion_Final.txt_r_f_edad_maxima_aval = "A"
End If
End Sub

Private Sub txt_edad_Change()
txt_r_edad = "A"

If txt_bancarizado_politica = "No" And txt_edad >= "35" Then
    txt_r_edad = "A"
ElseIf txt_bancarizado_politica = "Si" Then
    txt_r_edad = "A"
    Else
    txt_r_edad = "R"
End If
End Sub

Private Sub txt_edad_conyuge_AfterUpdate()

End Sub

Private Sub txt_edad_conyuge_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub txt_edad_conyuge_Change()


If txt_edad_conyuge = "" Then
   Estado_Resolucion_Final.txt_r_f_edad_conyuge = "N/A"
   Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge = "N/A"
End If

If txt_score_dicom_conyuge_aux > 0 And txt_edad_conyuge < 21 Then
    Estado_Resolucion_Final.txt_r_f_edad_conyuge = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_edad_conyuge = "A"
End If

        If (txt_score_dicom_conyuge_aux > 0 Or txt_score_dicom_conyuge_aux = 0) And txt_edad_conyuge > 72 Then
            Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge = "R"
        Else
            Estado_Resolucion_Final.txt_r_f_edad_maxima_conyuge = "A"
 
    End If



End Sub

Private Sub txt_edad_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub txt_estado_credito_Change()
   
If txt_estado_credito = "Rechazado" Then
        
        lbl_estado_credito.BackColor = &HFF&       'rojo
        lbl_estado_credito.ForeColor = &H8000000E  'blanco


        Else
        lbl_estado_credito.BackColor = &HC000&   ' verde
        lbl_estado_credito.ForeColor = &H8000000E  'blanco
   
    End If
End Sub

Private Sub txt_ESTADO_politica_bancarizado_new_Change()

End Sub


Private Sub txt_fechanacimiento_AfterUpdate()

    ssql = "select datediff(MONTH,'" & Format(txt_fechanacimiento, "yyyymmdd") & "', getdate())/12 edad"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
    txt_edad = rst!edad
    txt_edad.Locked = True
    
    'Controlar error y así validamos cuando el formato no sea el correcto

End Sub



Private Sub txt_n_carpeta_tributaria_Change()

txt_estado_credito = Empty



If txt_dv <> txt_dv_compara Then
  MsgBox ("Rut Invalido Revise ...")

End If
    
End Sub

Private Sub txt_n_trabajador_familia_AfterUpdate()

End Sub

Private Sub txt_n_trabajador_familia_Change()

txt_estado_credito = Empty

End Sub


Private Sub txt_plazo_credito_AfterUpdate()
If (cbx_Accion = "Capital De Trabajo" And txt_plazo_credito > 15) Or (cbx_Accion = "Activo Fijo" And txt_plazo_credito > 48) _
Or (cbx_Accion = "Vivienda" And txt_plazo_credito > 240) Then
                txt_r_plazo_credito = "R"
                'lbl_score_dicom.BackColor = &HFF&       'rojo
                'lbl_score_dicom.ForeColor = &H8000000E  'blanco
    
    Else
                txt_r_plazo_credito = "A"
                'lbl_score_dicom.BackColor = &HC000& 'VERDE
                'lbl_score_dicom.ForeColor = &H8000000E  'blanco

End If
End Sub

Private Sub txt_protesto_interno_Change()

If txt_protesto_interno = "Si" Then
  txt_r_protesto_interno = "R"
  
  Else
      txt_r_protesto_interno = "A"
 
End If

End Sub



Private Sub txt_r_accion_Change()

End Sub

Private Sub txt_r_aib_Change()

End Sub

Private Sub txt_r_años_vehiculo_Change()


Estado_Resolucion_Final.txt_r_f_antiguedad_veh = txt_r_años_vehiculo

End Sub

Private Sub txt_r_bancarizado_Change()

End Sub

Private Sub txt_r_cbx_actividad_economica_informal_oficio_Change()

End Sub



Private Sub txt_r_cod_observacion_Change()
If txt_cod_observacion = 4 Or txt_cod_observacion = 6 Or txt_cod_observacion = 7 Or txt_cod_observacion = 8 _
   Or txt_cod_observacion = 10 Or txt_cod_observacion = 12 Then
   
   txt_r_cod_observacion = "R"
   
   ElseIf txt_cod_observacion = 3 Or txt_cod_observacion = 5 Or txt_cod_observacion = 9 Or txt_cod_observacion = 11 _
   Or txt_cod_observacion = 13 Or txt_cod_observacion = 14 Or txt_cod_observacion = 15 Then
   
    txt_r_cod_observacion = "ZG"
    
    ElseIf txt_cod_observacion = 0 Then
          txt_r_cod_observacion = "A"
End If




End Sub


Private Sub txt_r_edad_Change()

End Sub

Private Sub txt_r_formalidad_negocio_Change()

End Sub

Private Sub txt_r_meses_antiguedad_Change()

End Sub

Private Sub txt_r_plazo_credito_Change()

End Sub

Private Sub txt_r_predictor_score_dicom_Change()

End Sub

Private Sub txt_renegociado_Change()

If txt_renegociado = "Si" Then
    
    txt_r_renegociado = "R"
    txt_renegociado.BackColor = &HFF&
    txt_renegociado.ForeColor = &H8000000E  'blanco

  Else
    txt_r_renegociado = "A"

    txt_renegociado.BackColor = &H8000000D
    txt_renegociado.ForeColor = &H8000000E

End If

End Sub

Private Sub txt_rut_aval_AfterUpdate()

If txt_rut_aval <> "" Then

txt_edad_aval.Locked = True

''''''''''''''''''  CALCULA EDAD DE CLIENTE
    
    ssql = "select cliente,datediff(MONTH,f_nacimiento, getdate())/12 edad,f_nacimiento" _
            & " from TBL_MICRO_FACT_MACA_CLIENTE" _
            & " where cliente = '" & txt_rut_aval & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    

If rst.EOF Then
        
    txt_edad_aval.Locked = False
    MsgBox "Ingrese Edad si el Cliente es Persona Natural", vbCritical
    
    ElseIf rst!f_nacimiento = "01/01/1900" And rst!cliente < 45000000 Then
    
    MsgBox "No Registramos Correctamente la EDAD Del Cliente, Ingreselo... ", vbCritical
    Else
        txt_edad_aval = rst!edad
    
End If

End If

'''' inhibe campos si es persona juridica

If txt_rut_aval >= 45000000 Then
    txt_edad_aval.Locked = True
    cbx_estado_civil.Locked = True
    txt_edad_aval = Empty
   
End If

End Sub

Private Sub txt_rut_aval_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   
   
If txt_rut_aval <> "" Then

   
''''''''''''''''''''''''''''''''''''''''''''''''
'''Revision De la datos consulta SINACOFI  CODIGO OBSERVACION aval
''''''''''''''''''''''''''''''''''''''''''''''''
 
ssql = "select top 1 rut, protesto, mora, infraccion_prev, score, convert(varchar,fecha_consulta,111) fecha_consulta, convert(varchar,getdate(),111) fecha_hoy, Cod_Observacion,convert(varchar,getdate()-1,111) fecha_menos1,convert(varchar,getdate()-1,111) fecha_hoy_menos1" _
        & " from tbl_micro_sinacofi" _
        & " where rut = '" & txt_rut_aval & "'" _
        & " order by fecha_consulta desc" _

    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            MsgBox "No existe Rut Ingresado en Sistema Local Sinacofi", vbCritical
        Else
       
        If (rst!cod_observacion = 4 Or rst!cod_observacion = 6 Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12) Then
        
            MsgBox "No se Puede Evaluar este cliente por su codigo de observación", vbCritical
            txt_rut_cliente = Empty
            txt_rut_conyuge = Empty
            txt_rut_aval = Empty
            txt_edad_conyuge = Empty
            txt_edad_aval = Empty
            txt_edad = Empty
            txt_r_edad = Empty
            
        Else
    
    
            If rst.EOF Then
                
                   MsgBox "No existe RUT en base de datos para comenzar Evaluación Solicite la emision de los datos", vbCritical
        
              Else
                    txt_fecha_actual_menos1_sinacofi = rst!fecha_menos1
                    If rst!fecha_consulta <> rst!fecha_hoy And rst!fecha_consulta <> rst!fecha_hoy_menos1 Then

                              MsgBox "Los Datos de Sinacofi han Vencido Solicite su Emision", vbCritical
                
               Else

                             If rst!cod_observacion = 4 Or rst!cod_observacion = 6 _
                                Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12 Then
                                txt_cod_observacion = rst!cod_observacion
    
                             MsgBox "No se Puede Evaluar este cliente por codigo de observación de aval ", vbCritical
                           
                            
                
                 Else

                                 If rst!cod_observacion = 2 Or rst!cod_observacion = 3 Or rst!cod_observacion = 5 Or rst!cod_observacion = 9 _
                                    Or rst!cod_observacion = 11 Or rst!cod_observacion = 13 Or rst!cod_observacion = 14 Or rst!cod_observacion = 15 Then
        
                                            Estado_Resolucion_Final.txt_r_f_cod_observacion_aval = "ZG"
                                            
    ''''''MORA - PROTESTOS - BOLETIN COMERCIAL AVAL   SEGUN CONSULTA SINACOFI
    
    
            ssql = "select rut, mora, protesto,infraccion_prev, score" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_aval & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
            If Not rst.EOF Then
            
                '''MORA aval AVAL
                If rst!mora = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "R"
                
                End If
                
                '''PROTESTO aval AVAL
                   
                If rst!protesto = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "R"
                
                End If
                   
                   
                '''BOLETIN aval AVAL
                   
                If rst!infraccion_prev = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "R"
                
                End If
        
        
        
                '''IR MINIMO AVAL

                If rst!score >= 513 Then
                
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_aval = "R"
                
                End If
        
        End If
                                            
                                            
                                            
                                            
                
                Else
                        
                            If rst!cod_observacion = 0 Then
                                        Estado_Resolucion_Final.txt_r_f_cod_observacion_aval = "A"
                                       
                                       If rst!protesto = "Cumple" Then
                                          Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "A"
                                                                                  
                                          Else
                                          Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "R"
                                       End If
                                    
                                    
                                       If rst!mora = "Cumple" Then
                                           Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "A"
                                            Else
                                             Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "R"
                                       End If
                                    
                                    
                                       If rst!infraccion_prev = "Cumple" Then
                                             Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "A"
                                            Else
                                             Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "R"
                                        End If
                                       
                                       
                                           ''''''MORA - PROTESTOS - BOLETIN COMERCIAL AVAL   SEGUN CONSULTA SINACOFI
    
    
            ssql = "select rut, mora, protesto,infraccion_prev, score" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_aval & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
            If Not rst.EOF Then
            
                txt_score_dicom_AVAL_aux = rst!score
            
                '''MORA aval AVAL
                If rst!mora = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_aval = "R"
                
                End If
                
                '''PROTESTO aval AVAL
                   
                If rst!protesto = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_aval = "R"
                
                End If
                   
                   
                '''BOLETIN aval AVAL
                   
                If rst!infraccion_prev = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_aval = "R"
                
                End If
        
        
        
                '''IR MINIMO AVAL

                If rst!score >= 513 Then
                
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_aval = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_aval = "R"
                
                End If
        
        End If

                                       
                                       
        
            End If
           End If
        End If
        End If
    End If
 End If
   
   
   
   
   
   
   
   
   
   
   
   
   
   
txt_estado_credito = Empty
   
txt_dv_aval = Empty
    
    Dim diga As Variant
       
    If Not IsNumeric(txt_rut_aval) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_aval = Empty
      End If
      
  
  
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_rut_aval = Replace(txt_rut_aval, "-", "")
    txt_rut_aval = Replace(txt_rut_aval, ".", "")
    txt_rut_aval = Replace(txt_rut_aval, ",", "")
    Vari3 = 2
    For I = 0 To Len(txt_rut_aval) - 1
     If Left(Right(txt_rut_aval, I + 1), 1) <> "." Then
      Vari1 = Vari1 + Left(Right(txt_rut_aval, I + 1), 1) * Vari3
      Vari2 = Vari1 Mod 11
      Select Case Vari2
       Case 0
        txt_dv_compara_aval.Text = "0"
       Case 1
        txt_dv_compara_aval.Text = "K"
       Case Else
        txt_dv_compara_aval.Text = 11 - Vari2
      End Select
      If Vari3 = 7 Then
       Vari3 = 2
      Else
       Vari3 = Vari3 + 1
      End If
     End If
    Next
    'fin digito verificador
    
    
    
    

    
    '''' CHEQUEA RUT DE aval QUE EXISTA COMO CONSULTA SINACOFI y en todas las bases segun politica

    If txt_rut_aval <> "" Then
       
           Call conectarBD
    
            ssql = "select rut, mora, protesto,infraccion_prev" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_aval & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
                If rst.EOF Then
                   MsgBox "Debe Consultar el Rut del Aval en Sinacofi"
                Else
                
                    If rst!mora = "No Cumple" Or rst!protesto = "No Cumple" Or rst!infraccion_prev = "No Cumple" Then
                       cbx_aval_inf_com = "No Cumple"
                    Else
                    
                        ssql = "select rut" _
                        & " from tbl_micro_resumen_prime_noprime" _
                        & " where rut = '" & txt_rut_aval & "'" _
                        & " and max_mora_total >0"
                        Set rst = cnn.Execute(ssql, , adCmdText)
                    
                        If Not rst.EOF Then
                           cbx_aval_inf_com = "No Cumple"
                           
                           Else
                                ssql = "select rut_numerico, marca_regular" _
                                & " from TBL_MICRO_RESUMEN_CRITICAL" _
                                & " where rut_numerico = '" & txt_rut_aval & "' and marca_regular =1"
                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                If Not rst.EOF Then
                                    cbx_aval_inf_com = "No Cumple"
                                                                        
                                    Else
                                        ssql = "select rut" _
                                        & " from TBL_MICRO_sbif" _
                                        & " where rut = '" & txt_rut_aval & "'"
                                        Set rst = cnn.Execute(ssql, , adCmdText)
                                                                                                            
                                          If Not rst.EOF Then
                                            cbx_aval_inf_com = "No Cumple"
                                                
                                            Else
                                                ssql = "select rut" _
                                                & " from TBL_MICRO_riesgo_renegociado" _
                                                & " where rut = '" & txt_rut_aval & "'"
                        
                                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                
                                                    If Not rst.EOF Then
                                                        cbx_aval_inf_com = "No Cumple"
                                                      
                                                      Else
                                                            ssql = "select rut10" _
                                                            & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                            & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_aval & "'" _
                                                            & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                                                If Not rst.EOF Then
                                                                cbx_aval_inf_com = "No Cumple"
                                                        
                                                                Else
                                                                    ssql = "select rut10" _
                                                                    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                                    & " where cast(SUBSTRING(rut10,1,9) as int)= '" & txt_rut_aval & "'" _
                                                                    & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                                
                                                                    If Not rst.EOF Then
                                                                    cbx_aval_inf_com = "No Cumple"
                                                                    
                                                                Else
                                                                    cbx_aval_inf_com = "Cumple"
                                                    
                                                            End If
                                                    
                                                        End If
                                                        
                                                    End If
                                            
                                            End If
                                    
                                    End If
                            
                            End If
                    
                    End If
                    
                End If
                
Else
cbx_aval_inf_com = "Cumple"
End If
    
        
        '''''' PROTESTO INTERNO - FILENEG - FRAUDE aval
    ssql = "select rut10" _
    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
    & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_aval & "'" _
    & " and fltro_cod ='R002'"
                       
    Set rst = cnn.Execute(ssql, , adCmdText)
                                                       
    If rst.EOF Then
    Estado_Resolucion_Final.txt_r_f_file_negativo_aval = "A"
    
    Else
    Estado_Resolucion_Final.txt_r_f_file_negativo_aval = "R"
    
    End If
    
    
    

    
    
'''''''''''''''''''''''''''''''''''''''''''''
'''''  NUMEROS DE ACRREEDORES   conyuge
'''''''''''''''''''''''''''''''''''''''''''''

    ssql = "select rut,n_institucionescondeuda+m_deudacreditocomerciales n_acreedores" _
            & " from tbl_micro_fact_ods_librodeudores" _
            & " where rut = '" & txt_rut_aval & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)
        
If rst.EOF Then
        txt_aux_n_acreedores_AVAL = 0
    Else
        txt_aux_n_acreedores_AVAL = rst!n_acreedores
    
End If

  If rst.EOF Then
  
        Estado_Resolucion_Final.txt_r_f_acreedores_aval = "A"

    ElseIf txt_aux_n_acreedores_AVAL = 0 Then
           Estado_Resolucion_Final.txt_r_f_acreedores_aval = "A"

    ElseIf rst!n_acreedores <= 3 Then
           Estado_Resolucion_Final.txt_r_f_acreedores_aval = "A"
           
        Else
        Estado_Resolucion_Final.txt_r_f_acreedores_aval = "R"
    
End If

    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---MORA ----
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ssql = "select * " _
            & " FROM TBL_MICRO_fact_RIESGO_MORA_DIA" _
            & " wHERE rut = '" & txt_rut_aval & "'" _
            & " AND diasmora >0 "
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_aval = "A"
            
        Else
        
            Estado_Resolucion_Final.txt_r_f_mora_directa_aval = "R"
        End If
        
        
    '''''' CASTIGOS TRAIDOS DESDE SERVIDOR RIESGO CLIENTE
          
          ssql = "select cas_rut" _
            & " FROM tbl_micro_RIESGO_CASTIGOS" _
            & " wHERE cas_rut = '" & txt_rut_aval & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
             Estado_Resolucion_Final.txt_r_f_castigo_directo_aval = "A"
            
        Else
        
             Estado_Resolucion_Final.txt_r_f_castigo_directo_aval = "R"
        End If
    
    
    '''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---VENCI-CASTIGO
    
      'ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_aval & "'"
    
      ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE rut_num = '" & txt_rut_cliente & "'"
    
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "A"
            
            ElseIf rst!dias_mora = 0 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "A"
            
            ElseIf rst!dias_mora <= 30 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "A"
            
            ElseIf rst!dias_mora >= 31 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "A"
            
            ElseIf rst!dias_mora >= 91 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "R"
            
            ElseIf rst!dias_mora >= 181 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_aval = "R"
End If
    
    
'------- Moras--sbif VIGENTE------------------------------------

    ssql = "select m_deudadirectamorosa" _
        & " FROM TBL_MICRO_FACT_ODS_LIBRODEUDORES " _
        & " WHERE rut = '" & txt_rut_aval & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
    
    If rst.EOF Then
        txt_aux_mora_directa_sbif_AVAL = 0
     Else
        txt_aux_mora_directa_sbif_AVAL = rst!m_deudadirectamorosa
        
    End If
    
    
    If rst.EOF Then
        Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_aval = "A"
     
        ElseIf txt_aux_mora_directa_sbif_AVAL = 0 Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_aval = "A"
            
        ElseIf rst!m_deudadirectamorosa = 0 Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_aval = "A"
            
        Else
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_aval = "R"
            
    End If

    
    
'------- Moras--sbif ultimos 12 mese------------------------------------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ssql = "select mora30, venc, cast, indcast, indvenc" _
        & " FROM tbl_micro_sbif " _
        & " WHERE rut = '" & txt_rut_aval & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)

        
If rst.EOF Then
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_aval = "A"
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval = "A"
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_aval = "A"
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval = "A"
    
    txt_credito_comercial_vigente_mora = "No"
    
Else
   
    If rst!venc = 1 Then
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_aval = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_aval = "A"
    End If
    '
    
    If rst!cast Then
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_aval = "A"
    End If
    
    '''''DEUDAS INDIRECTAS ULTIMOS 12 MESES
    
    If rst!indcast = 1 Then
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_aval = "A"
    End If
    '
    
    If rst!indvenc = 1 Then
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_aval = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_aval = "A"
    End If
    '
End If
    
  
 '''''''''''''''TRAE creditos varios protestos internos
    
            ssql = "select rut " _
                    & " from tbl_micro_maca_protestos" _
                    & " where tipo = 'P'" _
                    & " and rut = '" & txt_rut_aval & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                Estado_Resolucion_Final.txt_r_f_protesto_interno_aval = "A"
                Else
                Estado_Resolucion_Final.txt_r_f_protesto_interno_aval = "R"
            End If
  
  
    ''''RENEGOCIADOs
    ''''''''''''''''''''''''''''''
    
    ssql = "select rut" _
            & " FROM TBL_MICRO_RIESGO_RENEGOCIADO" _
            & " wHERE rut = '" & txt_rut_aval & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_renegociado_aval = "A"
    
          Else
            Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "R"
    
        End If
        
        
    '''''' PROTESTO INTERNO - FILENEG - FRAUDE TITULAR
    ssql = "select rut10" _
    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
    & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_aval & "'" _
    & " and fltro_cod ='R002'"
                       
    Set rst = cnn.Execute(ssql, , adCmdText)
                                                       
    If rst.EOF Then
    Estado_Resolucion_Final.txt_r_f_file_negativo_aval = "A"

    Else
    Estado_Resolucion_Final.txt_r_f_file_negativo_aval = "R"
    
    End If
        
        
        
                


'''''''''''''''''''''  CASTIGOS HISTORICOS aval
'''''''''''''''''''''''''''''''''''''''''''''''''''
            ssql = "select rut " _
                    & " from tbl_micro_maca_castigos_his" _
                    & " where rut = '" & txt_rut_aval & "'"
                        
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                Estado_Resolucion_Final.txt_r_f_castigo_historico_aval = "A"
                Else
                Estado_Resolucion_Final.txt_r_f_castigo_historico_aval = "R"
            End If






End If  '''''''''''''CIERRE DE LA PRIMERA CONDICION PARA INGRESO A LA RUTINA DE EVALUACION FICHA


End If

End Sub



Private Sub txt_rut_aval_Change()
txt_edad_aval = Empty
End Sub

Private Sub txt_rut_cliente_Change()
txt_edad = Empty
txt_r_edad = Empty

txt_r_formalidad_negocio = Empty
txt_r_cbx_antiguedad_rubro = Empty
txt_r_cbx_actividad_economica_informal_oficio = Empty
txt_r_cbx_bien_Raiz = Empty
txt_r_cbx_vehiculos_propios = Empty

cbx_estado_civil.Clear
cbx_estado_civil.AddItem "Soltero"
cbx_estado_civil.AddItem "Divorciado"
cbx_estado_civil.AddItem "Viudo"

End Sub

Private Sub txt_rut_conyuge_AfterUpdate()

If txt_rut_conyuge <> "" Then


''''''''''''''''''  CALCULA EDAD DE CLIENTE
    
    ssql = "select cliente,datediff(MONTH,f_nacimiento, getdate())/12 edad,f_nacimiento" _
            & " from TBL_MICRO_FACT_MACA_CLIENTE" _
            & " where cliente = '" & txt_rut_conyuge & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    

If rst.EOF Then
        
txt_edad_conyuge.Locked = False
    MsgBox "Ingrese Edad si el Cliente es Persona Natural", vbCritical
    
    ElseIf rst!f_nacimiento = "01/01/1900" And rst!cliente < 45000000 Then
    txt_edad_conyuge.Locked = False
    MsgBox "No Registramos Correctamente la EDAD del Conyuge, Ingreselo... ", vbCritical
    
    Else
        txt_edad_conyuge = rst!edad
    
End If

End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''

txt_estado_credito = Empty
cbx_conyuge_inf_com = Empty

'''' CHEQUEA RUT DE CONYUGE QUE EXISTA COMO CONSULTA SINACOFI y en todas las bases segun politica

    If txt_rut_cliente > 45000000 Then
        
    If txt_rut_conyuge <> "" Then
       
    txt_r_conyuge_inf_com = Empty
       
           Call conectarBD
    
            ssql = "select rut, mora, protesto,infraccion_prev" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_conyuge & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
                If rst.EOF Then
                   MsgBox "Debe Consultar el Rut del Conyuge en Sinacofi"
                Else
                
                    If rst!mora = "No Cumple" Or rst!protesto = "No Cumple" Or rst!infraccion_prev = "No Cumple" Then
                       cbx_conyuge_inf_com = "No Cumple"
                    Else
                    
                        ssql = "select rut" _
                        & " from tbl_micro_resumen_prime_noprime" _
                        & " where rut = '" & txt_rut_conyuge & "'" _
                        & " and max_mora_total >0"
                        Set rst = cnn.Execute(ssql, , adCmdText)
                    
                        If Not rst.EOF Then
                           cbx_conyuge_inf_com = "No Cumple"
                           
                           Else
                                ssql = "select rut_numerico" _
                                & " from TBL_MICRO_RESUMEN_CRITICAL" _
                                & " where rut_numerico = '" & txt_rut_conyuge & "' and marca_regular =1"
                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                If Not rst.EOF Then
                                    cbx_conyuge_inf_com = "No Cumple"
                                                                        
                                    Else
                                        ssql = "select rut" _
                                        & " from TBL_MICRO_sbif" _
                                        & " where rut = '" & txt_rut_conyuge & "'"
                                        Set rst = cnn.Execute(ssql, , adCmdText)
                                                                                                            
                                          If Not rst.EOF Then
                                            cbx_conyuge_inf_com = "No Cumple"
                                                
                                            Else
                                                ssql = "select rut" _
                                                & " from TBL_MICRO_riesgo_renegociado" _
                                                & " where rut = '" & txt_rut_conyuge & "'"
                        
                                                Set rst = cnn.Execute(ssql, , adCmdText)
                                                
                                                    If Not rst.EOF Then
                                                        cbx_conyuge_inf_com = "No Cumple"
                                                      
                                                      Else
                                                            ssql = "select rut10" _
                                                            & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                            & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_conyuge & "'" _
                                                            & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                        
                                                                If Not rst.EOF Then
                                                                cbx_conyuge_inf_com = "No Cumple"
                                                        
                                                                Else
                                                                    ssql = "select rut10" _
                                                                    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
                                                                    & " where cast(SUBSTRING(rut10,1,9) as int) = '" & txt_rut_conyuge & "'" _
                                                                    & " and FLTRO_COD ='R002'"
                        
                                                            Set rst = cnn.Execute(ssql, , adCmdText)
                                                                
                                                                    If Not rst.EOF Then
                                                                    cbx_conyuge_inf_com = "No Cumple"
                                                                    
                                                                Else
                                                                    cbx_conyuge_inf_com = "Cumple"
                                                    
                                                            End If
                                                    
                                                        End If
                                                        
                                                    End If
                                            
                                            End If
                                    
                                    End If
                            
                            End If
                    
                    End If
                    
                End If
                
        Else
       
       MsgBox "Debe Ingresar Rut Conyuge"
    
    End If
       
End If





End Sub

Private Sub txt_rut_conyuge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   
If txt_rut_conyuge <> "" Then

   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Revision De la datos consulta SINACOFI  CODIGO OBSERVACION CONYUGE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
        ssql = "select top 1 rut, protesto, mora, infraccion_prev, score, convert(varchar,fecha_consulta,111) fecha_consulta, convert(varchar,getdate(),111) fecha_hoy, Cod_Observacion,convert(varchar,getdate()-1,111) fecha_menos1,convert(varchar,getdate()-1,111) fecha_hoy_menos1" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_conyuge & "'" _
            & " order by fecha_consulta desc" _

            Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
       
        MsgBox "No existe Rut Ingresado en Sistema Local Sinacofi", vbCritical
        
        Else
    
       If (rst!cod_observacion = 4 Or rst!cod_observacion = 6 Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12) Then
        
            MsgBox "No se Puede Evaluar este cliente por su codigo de observación", vbCritical
            txt_rut_cliente = Empty
            txt_rut_conyuge = Empty
            txt_rut_aval = Empty
            txt_edad_conyuge = Empty
            txt_edad_aval = Empty
            txt_edad = Empty
            txt_r_edad = Empty
            
        Else
    
    
            If rst.EOF Then
                
                   MsgBox "No existe RUT en base de datos para comenzar Evaluación Solicite la emision de los datos", vbCritical
        
              Else
                       txt_fecha_actual_menos1_sinacofi = rst!fecha_menos1
                    If rst!fecha_consulta <> rst!fecha_hoy And rst!fecha_consulta <> rst!fecha_hoy_menos1 Then

                        MsgBox "Los Datos de Sinacofi han Vencido Solicite su Emision", vbCritical
                
               Else

                             If rst!cod_observacion = 4 Or rst!cod_observacion = 6 _
                                Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12 Then
                                txt_cod_observacion = rst!cod_observacion
    
                             MsgBox "No se Puede Evaluar este cliente por codigo de observación de Conyuge", vbCritical
                
                 Else

                                 If rst!cod_observacion = 2 Or rst!cod_observacion = 3 Or rst!cod_observacion = 5 Or rst!cod_observacion = 9 _
                                    Or rst!cod_observacion = 11 Or rst!cod_observacion = 13 Or rst!cod_observacion = 14 Or rst!cod_observacion = 15 Then
        
                                            Estado_Resolucion_Final.txt_r_f_cod_observacion_conyuge = "ZG"
                                            
    
    
    ''''''MORA - PROTESTOS - BOLETIN COMERCIAL CONYUGE   SEGUN CONSULTA SINACOFI
    
    
            ssql = "select rut, mora, protesto,infraccion_prev,score" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_conyuge & "'"
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
        If Not rst.EOF Then
         
                '''MORA CLIENTE CONYUGE
                If rst!mora = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_conyuge = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_morosidad_sinac_conyuge = "R"
                
                End If
                
                '''PROTESTO CLIENTE CONYUGE
                   
                If rst!protesto = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_conyuge = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_protesto_sinac_conyuge = "R"
                
                End If
                   
                   
                '''BOLETIN CLIENTE CONYUGE
                   
                If rst!infraccion_prev = "Cumple" Then
                
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge = "R"
                
                End If

                '''IR MINIMO CONYUGE

                If rst!score >= 513 Then
                
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_conyuge = "A"
                
                Else
                    
                   Estado_Resolucion_Final.txt_r_f_ir_sinac_conyuge = "R"
                
                End If

        End If
                                            
                                            
                                            
                                            
                
                Else
                            'paso de score CONYUGE
                            txt_score_dicom_conyuge_aux = rst!score
                        
                            If rst!cod_observacion = 0 Then
                                        Estado_Resolucion_Final.txt_r_f_cod_observacion_conyuge = "A"
                                       
                                       If rst!protesto = "Cumple" Then
                                          Estado_Resolucion_Final.txt_r_f_protesto_sinac_conyuge = "A"
                                                                                  
                                          Else
                                          Estado_Resolucion_Final.txt_r_f_protesto_sinac_conyuge = "R"
                                       End If
                                    
                                    
                                       If rst!mora = "Cumple" Then
                                           Estado_Resolucion_Final.txt_r_f_morosidad_sinac_conyuge = "A"
                                            Else
                                             Estado_Resolucion_Final.txt_r_f_morosidad_sinac_conyuge = "R"
                                       End If
                                    
                                    
                                       If rst!infraccion_prev = "Cumple" Then
                                             Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge = "A"
                                            Else
                                             Estado_Resolucion_Final.txt_r_f_boletin_sinac_conyuge = "R"
                                        End If
                                                                 
                            'End If
                                    
                                    
                                    
                                    If rst!score >= 513 Then
                                    Estado_Resolucion_Final.txt_r_f_ir_sinac_conyuge = "A"
                                    Else
                                    Estado_Resolucion_Final.txt_r_f_ir_sinac_conyuge = "R"
                                    End If
                                       
                                       

        End If
                                       
                                       
        
            End If
           End If
        End If
        End If
    End If

   
   

  
If txt_rut_cliente = txt_rut_conyuge Then
   
   MsgBox "El Rut del Conyuge y Cliente no pueden ser iguales", vbCritical
   txt_rut_conyuge = Empty
   
Else

    txt_estado_credito = Empty
   

    txt_dv_conyuge = Empty
    
    Dim diga As Variant
       
    If Not IsNumeric(txt_rut_conyuge) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_conyuge = Empty
      End If
  
    ' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_rut_conyuge = Replace(txt_rut_conyuge, "-", "")
    txt_rut_conyuge = Replace(txt_rut_conyuge, ".", "")
    txt_rut_conyuge = Replace(txt_rut_conyuge, ",", "")
    Vari3 = 2
    For I = 0 To Len(txt_rut_conyuge) - 1
     If Left(Right(txt_rut_conyuge, I + 1), 1) <> "." Then
      Vari1 = Vari1 + Left(Right(txt_rut_conyuge, I + 1), 1) * Vari3
      Vari2 = Vari1 Mod 11
      Select Case Vari2
       Case 0
        txt_dv_compara_conyuge.Text = "0"
       Case 1
        txt_dv_compara_conyuge.Text = "K"
       Case Else
        txt_dv_compara_conyuge.Text = 11 - Vari2
      End Select
      If Vari3 = 7 Then
       Vari3 = 2
      Else
       Vari3 = Vari3 + 1
      End If
     End If
    Next
    'fin digito verificador
    
End If


Call conectarBD

        ''''''  FILENEG - conyuge
    ssql = "select rut10" _
    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
    & " where cast(SUBSTRING(rut10,1,9) as int) =  '" & txt_rut_conyuge & "'" _
    & " and fltro_cod ='R002'"

    Set rst = cnn.Execute(ssql, , adCmdText)
                                                       
    If rst.EOF Then
    Estado_Resolucion_Final.txt_r_f_file_negativo = "A"
    
    Else
    Estado_Resolucion_Final.txt_r_f_file_negativo = "R"
    
    End If




'''''''''''''''''''''''''''''''''''''''''''''
'''''  NUMEROS DE ACRREEDORES   conyuge
'''''''''''''''''''''''''''''''''''''''''''''

    ssql = "select rut,n_institucionescondeuda+m_deudacreditocomerciales n_acreedores" _
            & " from tbl_micro_fact_ods_librodeudores" _
            & " where rut = '" & txt_rut_conyuge & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)
        
If rst.EOF Then
        txt_aux_n_acreedores_conyuge = 0
    Else
        txt_aux_n_acreedores_conyuge = rst!n_acreedores
    
End If

  If rst.EOF Then
  
        Estado_Resolucion_Final.txt_r_f_acreedores_conyuge = "A"

    ElseIf txt_aux_n_acreedores_conyuge = 0 Then
           Estado_Resolucion_Final.txt_r_f_acreedores_conyuge = "A"

    ElseIf rst!n_acreedores <= 3 Then
           Estado_Resolucion_Final.txt_r_f_acreedores_conyuge = "A"
           
        Else
        Estado_Resolucion_Final.txt_r_f_acreedores_conyuge = "R"
    
End If




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---MORA ----
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ssql = "select * " _
            & " FROM TBL_MICRO_fact_RIESGO_MORA_DIA" _
            & " wHERE rut = '" & txt_rut_conyuge & "'" _
            & " AND diasmora >0 "
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_conyuge = "A"
            
        Else
        
            Estado_Resolucion_Final.txt_r_f_mora_directa_conyuge = "R"
        End If
    
    
    '''''' CASTIGOS TRAIDOS DESDE SERVIDOR RIESGO CLIENTE
          
          ssql = "select cas_rut" _
            & " FROM tbl_micro_RIESGO_CASTIGOS" _
            & " wHERE cas_rut = '" & txt_rut_conyuge & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
             Estado_Resolucion_Final.txt_r_f_castigo_directo_conyuge = "A"
            
        Else
        
             Estado_Resolucion_Final.txt_r_f_castigo_directo_conyuge = "R"
        End If
        
    
    
    '''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---VENCI-CASTIGO
    
      'ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_conyuge & "'"
    
      ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE rut_num = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "A"
            
            ElseIf rst!dias_mora = 0 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "A"
            
            ElseIf rst!dias_mora <= 30 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "A"
            
            ElseIf rst!dias_mora >= 31 Then
            'And rst!max_mora_total <= 90
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "A"
            
            ElseIf rst!dias_mora >= 91 Then
            'And rst!max_mora_total <= 180
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "R"
            
            ElseIf rst!dias_mora >= 181 Then
            Estado_Resolucion_Final.txt_r_f_Vencido_directo_conyuge = "R"
   
        End If


'------- Moras--sbif VIGENTE------------------------------------

    ssql = "select m_deudadirectamorosa" _
        & " FROM TBL_MICRO_FACT_ODS_LIBRODEUDORES " _
        & " WHERE rut = '" & txt_rut_conyuge & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
    
    If rst.EOF Then
        txt_aux_mora_directa_sbif_conyugE = 0
     Else
        txt_aux_mora_directa_sbif_conyugE = rst!m_deudadirectamorosa
        
    End If
    
    
    If rst.EOF Then
        Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_conyuge = "A"
     
        ElseIf txt_aux_mora_directa_sbif_conyugE = 0 Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_conyuge = "A"
            
        ElseIf rst!m_deudadirectamorosa = 0 Then
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_conyuge = "A"
            
        Else
            Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF_conyuge = "R"
            
    End If
     

'------- Moras--sbif ultimos 12 mese------------------------------------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ssql = "select mora30, venc, cast, indcast, indvenc" _
        & " FROM tbl_micro_sbif " _
        & " WHERE rut = '" & txt_rut_conyuge & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)

        
If rst.EOF Then

    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_conyuge = "A"
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge = "A"
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_conyuge = "A"
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge = "A"

    txt_credito_comercial_vigente_mora = "No"
    
       
Else
    
    If rst!venc = 1 Then
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_conyuge = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF_conyuge = "A"
    End If
    '
    
    If rst!cast Then
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF_conyuge = "A"
    End If
    
    '''''DEUDAS INDIRECTAS ULTIMOS 12 MESES
    
    If rst!indcast = 1 Then
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF_conyuge = "A"
    End If
    '
    
    If rst!indvenc = 1 Then
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_conyuge = "R"
    Else
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF_conyuge = "A"
    End If
    '
End If



'''''''''''''''''''''  CASTIGOS HISTORICOS CONYUGE
'''''''''''''''''''''''''''''''''''''''''''''''''''
            ssql = "select rut " _
                    & " from tbl_micro_maca_castigos_his" _
                    & " where rut = '" & txt_rut_conyuge & "'"
                        
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                Estado_Resolucion_Final.txt_r_f_castigo_historico_conyuge = "A"
                Else
                Estado_Resolucion_Final.txt_r_f_castigo_historico_conyuge = "R"
            End If


    ''''RENEGOCIADOs
    ''''''''''''''''''''''''''''''
    
    ssql = "select rut" _
            & " FROM TBL_MICRO_RIESGO_RENEGOCIADO" _
            & " wHERE rut = '" & txt_rut_conyuge & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            Estado_Resolucion_Final.txt_r_f_renegociado_conyuge = "A"
    
          Else
            Estado_Resolucion_Final.txt_r_f_renegociado_conyuge = "R"
    
        End If


'''''''''''''''TRAE creditos varios protestos internos
    
            ssql = "select rut " _
                    & " from tbl_micro_maca_protestos" _
                    & " where tipo = 'P'" _
                    & " and rut = '" & txt_rut_conyuge & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                Estado_Resolucion_Final.txt_r_f_protesto_interno_conyuge = "A"
                Else
                Estado_Resolucion_Final.txt_r_f_protesto_interno_conyuge = "R"
            End If


End If  '''''' CIERRE DE LA PRIMERA CONDICION PARA COMENZAR EVALUACION

End If

End Sub

Private Sub txt_rut_conyuge_Change()
txt_edad_conyuge = Empty

If txt_rut_conyuge <> "" Then

txt_edad_conyuge.Locked = True
cbx_estado_civil.Clear
cbx_estado_civil.AddItem "Casado"
cbx_estado_civil.AddItem "Separado De Hecho"

Else

cbx_estado_civil.Clear
 
cbx_estado_civil.AddItem "Divorciado"
cbx_estado_civil.AddItem "Soltero"
cbx_estado_civil.AddItem "Viudo"

End If

End Sub

Private Sub txt_score_dicom_AfterUpdate()

End Sub

Private Sub txt_score_dicom_Change()
   
   
If txt_bancarizado_politica = "Si" Then

        If txt_rut_cliente < 45000000 Then
            If txt_score_dicom < 347 Then
                txt_r_predictor_score_dicom = "R"
                lbl_score_dicom.BackColor = &HFF&       'rojo
                lbl_score_dicom.ForeColor = &H8000000E  'blanco
    
        Else
                txt_r_predictor_score_dicom = "A"
                lbl_score_dicom.BackColor = &HC000& 'VERDE
                lbl_score_dicom.ForeColor = &H8000000E  'blanco
        End If

        Else
        txt_r_predictor_score_dicom = "ZG"
    
    End If

ElseIf txt_bancarizado_politica = "No" And txt_score_dicom > 100 Then
    txt_r_predictor_score_dicom = "A"
    lbl_score_dicom.BackColor = &HC000& 'VERDE
    lbl_score_dicom.ForeColor = &H8000000E  'blanco
    
ElseIf txt_bancarizado_politica = "No" And txt_score_dicom = 0 Then
    
    'Cambio solicitado por Viviana Manriquez mail 16-10-2013
    'txt_r_predictor_score_dicom = "ZG"
    txt_r_predictor_score_dicom = "A"
    lbl_score_dicom.BackColor = &HC000& 'VERDE
    lbl_score_dicom.ForeColor = &H8000000E  'blanco

Else
                txt_r_predictor_score_dicom = "R"
                lbl_score_dicom.BackColor = &HFF&       'rojo
                lbl_score_dicom.ForeColor = &H8000000E  'blanco

End If

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
MsgBox ("Boton Deshabilitado Ocupe Opciones De Menu")
Cancel = True
End If
End Sub



Private Sub UserForm_Initialize()

    Call conectarBD
    
    '''''''''' TRAE CODIGO_SUCURSALES '''''''''
    ssql = "select distinct(codigo_sucursal) FROM TBL_ejecutivo " _
    & " where cargo_ejecutivo ='EJECUTIVO MICROEMPRESA'" _
    & " ORDER BY codigo_sucursal"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_codigo_sucursal.AddItem rst!codigo_sucursal
        rst.MoveNext
    Loop
    
    
    '''''''''' TRAE AUTORIZADORES DE EXCEPCIONES MICROEMPRESA '''''''''
    ssql = "select codigo_sucursal+''+codigo_ejecutivo+'-'+Nombre_Ejecutivo codigo_ejecutivo_excepcion FROM TBL_ejecutivo " _
    & " where aut_excepcion_micro =1" _
    & " ORDER BY codigo_sucursal"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_ejecutivo_excepcion.AddItem rst!codigo_ejecutivo_excepcion
        rst.MoveNext
    Loop
    

cbx_Accion.AddItem "Activo Fijo"
cbx_Accion.AddItem "Construccion" '''Agregado Sistema Automatico
cbx_Accion.AddItem "Capital De Trabajo"
cbx_Accion.AddItem "Necesidades Personales"
cbx_Accion.AddItem "Terreno"  '''Agregado Sistema Automatico
cbx_Accion.AddItem "Vivienda"

cbx_años_vehiculo_auto.AddItem "Menor a 5 Años" '''Agregado Sistema Automatico
cbx_años_vehiculo_auto.AddItem "Mayor a 5 Años" '''Agregado Sistema Automatico

cbx_años_vehiculo_bus.AddItem "Menor a 10 Años" '''Agregado Sistema Automatico
cbx_años_vehiculo_bus.AddItem "Mayor a 10 Años" '''Agregado Sistema Automatico


cbx_Destino_Construccion.AddItem "Compra Terreno"
cbx_Destino_Construccion.AddItem "Vivienda-Galpon-Bodega"
cbx_Destino_Construccion.AddItem "Cred.Comercial Corto Plazo"

cbx_tipo_vehiculo.AddItem "Colectivo"
cbx_tipo_vehiculo.AddItem "Buses"
cbx_tipo_vehiculo.AddItem "Liebres"
cbx_tipo_vehiculo.AddItem "Microbuses"
cbx_tipo_vehiculo.AddItem "Taxi"
cbx_tipo_vehiculo.AddItem "Radiotaxi"

cbx_destino_terreno.AddItem "Compra Terreno"

cbx_Destino_Credito_Cap_Tra.AddItem "Compra De Cartera"
cbx_Destino_Credito_Cap_Tra.AddItem "Materia Prima"
cbx_Destino_Credito_Cap_Tra.AddItem "Insumos"
cbx_Destino_Credito_Cap_Tra.AddItem "Mercaderias"
cbx_Destino_Credito_Cap_Tra.AddItem "Contratación Mano Obra"
cbx_Destino_Credito_Cap_Tra.AddItem "Gasto Operativo"


cbx_Destino_Credito_Act_Fij.AddItem "Compra Cartera"
cbx_Destino_Credito_Act_Fij.AddItem "Vehiculo"
cbx_Destino_Credito_Act_Fij.AddItem "Maquinaria"
cbx_Destino_Credito_Act_Fij.AddItem "Mejora De Tecnología"
cbx_Destino_Credito_Act_Fij.AddItem "Vehic.No-TrasPasajero"


cbx_Destino_Credito_Nec_Per.AddItem "Necesidades Personales"

cbx_Destino_Credito_Vivienda.AddItem "Bien Raiz"
cbx_Destino_Credito_Vivienda.AddItem "Cred.Comer.Corto Plazo"

cbx_antiguedad_rubro.AddItem "Asof"
cbx_antiguedad_rubro.AddItem "Carpeta Tributaria"
cbx_antiguedad_rubro.AddItem "Centros de Emprendimiento"
cbx_antiguedad_rubro.AddItem "Certificado Linea De Colectivo"
cbx_antiguedad_rubro.AddItem "Certificado Ministerio Trans./Telecom."
cbx_antiguedad_rubro.AddItem "Conatacoch"
cbx_antiguedad_rubro.AddItem "Corfo "
cbx_antiguedad_rubro.AddItem "DAI"
cbx_antiguedad_rubro.AddItem "Dideco"
cbx_antiguedad_rubro.AddItem "Federaciones de Transportes"
cbx_antiguedad_rubro.AddItem "Fosis"
cbx_antiguedad_rubro.AddItem "Indap "
cbx_antiguedad_rubro.AddItem "Iniciación De Actividades"
cbx_antiguedad_rubro.AddItem "Patente"
cbx_antiguedad_rubro.AddItem "Omdel"
cbx_antiguedad_rubro.AddItem "Ong"
cbx_antiguedad_rubro.AddItem "Permiso Municipal"
cbx_antiguedad_rubro.AddItem "Prodesal"
cbx_antiguedad_rubro.AddItem "Sercotec"
cbx_antiguedad_rubro.AddItem "Antig. Acreditada Cliente Bco"
cbx_antiguedad_rubro.AddItem "Certif. Junta Vecinos c/timbre"




'Actividades FORMALES
'------------------------

cbx_actividad_economica_formal.AddItem "ARTESANO"
cbx_actividad_economica_formal.AddItem "ALMACEN"
cbx_actividad_economica_formal.AddItem "BAZAR Y PAQUETERIA"
cbx_actividad_economica_formal.AddItem "BOTILLERIA"
cbx_actividad_economica_formal.AddItem "CARNICERIA"
cbx_actividad_economica_formal.AddItem "CAFETERIA"
cbx_actividad_economica_formal.AddItem "COMIDA RAPIDA"
cbx_actividad_economica_formal.AddItem "COMPRA MADERA"
cbx_actividad_economica_formal.AddItem "COMPRA Y VENTA DE ANIMALES"
cbx_actividad_economica_formal.AddItem "COMPRA Y VENTA DE LEÑA"
cbx_actividad_economica_formal.AddItem "COMPRA Y VENTA DE METAL"
cbx_actividad_economica_formal.AddItem "COMPRA Y VTA REPUESTOS MOTO"
cbx_actividad_economica_formal.AddItem "DISTRIBUIDORA DE HUEVOS"
cbx_actividad_economica_formal.AddItem "FERRETERIA"
cbx_actividad_economica_formal.AddItem "FRUTOS SECOS DEL PAIS"
cbx_actividad_economica_formal.AddItem "LIBRERIA"
cbx_actividad_economica_formal.AddItem "LUBRICENTRO"
cbx_actividad_economica_formal.AddItem "MINIMARKET"
cbx_actividad_economica_formal.AddItem "OPTICA"
cbx_actividad_economica_formal.AddItem "PANADERIA"
cbx_actividad_economica_formal.AddItem "RESTAURANT"
cbx_actividad_economica_formal.AddItem "SEMANERO"
cbx_actividad_economica_formal.AddItem "TIENDA DE ACCESORIOS CELULARES"
cbx_actividad_economica_formal.AddItem "TIENDA DE MENAJE"
cbx_actividad_economica_formal.AddItem "TIENDA DE ROPA"
cbx_actividad_economica_formal.AddItem "VENTA ALIMENTO MASCOTA"
cbx_actividad_economica_formal.AddItem "VENTA AL X MENOR DE MUEBLES"
cbx_actividad_economica_formal.AddItem "VENTA DE RELOJES Y JOYAS"
cbx_actividad_economica_formal.AddItem "VENTA DE FRUTAS Y VERDURAS"
cbx_actividad_economica_formal.AddItem "VENTA PRODUCUCTOS FOTOGRAFICOS"
cbx_actividad_economica_formal.AddItem "VENTA REPUESTO AGRICOLAS"
cbx_actividad_economica_formal.AddItem "VENTA REPUESTO AUTOMOVILES"
cbx_actividad_economica_formal.AddItem "VENTA REPUESTO BICICLETAS"
cbx_actividad_economica_formal.AddItem "VENTA Y DISTRIBUCION DE GAS"


'Actividades FORMAL SERVICIO O PRODUCCION
'------------------------

cbx_actividad_economica_formal_servicio.AddItem "BUSES"
cbx_actividad_economica_formal_servicio.AddItem "CAFETERIA"
cbx_actividad_economica_formal_servicio.AddItem "CENTRO DE LLAMADOS"
cbx_actividad_economica_formal_servicio.AddItem "CIBER CAFÉ"
cbx_actividad_economica_formal_servicio.AddItem "COLECTIVO"
cbx_actividad_economica_formal_servicio.AddItem "COMIDA RAPIDA"
cbx_actividad_economica_formal_servicio.AddItem "CONFECCION"
cbx_actividad_economica_formal_servicio.AddItem "CONTRATISTA CONSTRUCCION"
cbx_actividad_economica_formal_servicio.AddItem "CORREO MENSAJERIA"
cbx_actividad_economica_formal_servicio.AddItem "ESTACIONAMIENTO DE VEHICULOS"
cbx_actividad_economica_formal_servicio.AddItem "FABRICA DE PASTELES"
cbx_actividad_economica_formal_servicio.AddItem "FABRICACION Y REPARACION DE JOYAS"
cbx_actividad_economica_formal_servicio.AddItem "HOJALATERIA"
cbx_actividad_economica_formal_servicio.AddItem "IMPRENTA"
cbx_actividad_economica_formal_servicio.AddItem "JARDIN INFANTIL"
cbx_actividad_economica_formal_servicio.AddItem "MICROBUSES"
cbx_actividad_economica_formal_servicio.AddItem "MUEBLERIA"
cbx_actividad_economica_formal_servicio.AddItem "LAVANDERIA"
cbx_actividad_economica_formal_servicio.AddItem "LIEBRES"
cbx_actividad_economica_formal_servicio.AddItem "PAISAJISMO"
cbx_actividad_economica_formal_servicio.AddItem "PANADERIA O AMASANDERIA"
cbx_actividad_economica_formal_servicio.AddItem "PELUQUERIA"
cbx_actividad_economica_formal_servicio.AddItem "POLARIZACION DE VIDRIOS"
cbx_actividad_economica_formal_servicio.AddItem "PRODUCTORA DE EVENTOS"
cbx_actividad_economica_formal_servicio.AddItem "RADIO TAXI"
cbx_actividad_economica_formal_servicio.AddItem "REPARACION DE MOTOCICLETAS"
cbx_actividad_economica_formal_servicio.AddItem "REPARADORA DE ZAPATOS"
cbx_actividad_economica_formal_servicio.AddItem "RESIDENCIAL"
cbx_actividad_economica_formal_servicio.AddItem "RESTAURANT"
cbx_actividad_economica_formal_servicio.AddItem "SALON DE POOL"
cbx_actividad_economica_formal_servicio.AddItem "SERVICIOS COMPUTACIONALES"
cbx_actividad_economica_formal_servicio.AddItem "SERVICIOS DE BANQUETE"
cbx_actividad_economica_formal_servicio.AddItem "SERVICIOS DE FOTOCOPIA"
cbx_actividad_economica_formal_servicio.AddItem "SERVICIOS DE JARDINERIA Y MANTENCION PISCINAS"
cbx_actividad_economica_formal_servicio.AddItem "TALLER MECANICO"
cbx_actividad_economica_formal_servicio.AddItem "TAXI"
cbx_actividad_economica_formal_servicio.AddItem "TRANSPORTE DE CARGA"
cbx_actividad_economica_formal_servicio.AddItem "TRANSPORTE ESCOLAR"
cbx_actividad_economica_formal_servicio.AddItem "TRANSPORTE TURISMO"


'Actividades Semiformales
'---------------------------
cbx_actividad_economica_semiformal.AddItem "CARROS CON PERMISO FRENTE A"
cbx_actividad_economica_semiformal.AddItem "FERIA ARTESANIA"
cbx_actividad_economica_semiformal.AddItem "FERIAS LIBRES"
cbx_actividad_economica_semiformal.AddItem "FERIAS PERSAS"
cbx_actividad_economica_semiformal.AddItem "FLORISTAS"
cbx_actividad_economica_semiformal.AddItem "KIOSKO DE REVISTA Y DIARIOS"
cbx_actividad_economica_semiformal.AddItem "SEMANERO"


'Actividades Informales (Oficios)
'-----------------------------------
cbx_actividad_economica_informal_oficio.AddItem "ALBAÑIL"
cbx_actividad_economica_informal_oficio.AddItem "ARTESANO"
cbx_actividad_economica_informal_oficio.AddItem "CARPINTERO"
cbx_actividad_economica_informal_oficio.AddItem "CERAMISTA"
cbx_actividad_economica_informal_oficio.AddItem "CERRAJERO"
cbx_actividad_economica_informal_oficio.AddItem "ELECTRICISTA"
cbx_actividad_economica_informal_oficio.AddItem "GASFITER"
cbx_actividad_economica_informal_oficio.AddItem "JARDINERO"
cbx_actividad_economica_informal_oficio.AddItem "MAESTRO PINTOR"
cbx_actividad_economica_informal_oficio.AddItem "MECANICO"
cbx_actividad_economica_informal_oficio.AddItem "MODISTAS"
cbx_actividad_economica_informal_oficio.AddItem "SUPLEMENTERO"
cbx_actividad_economica_informal_oficio.AddItem "TAPICERO"
cbx_actividad_economica_informal_oficio.AddItem "ZAPATERO"
cbx_actividad_economica_informal_oficio.AddItem "PRODUCCION ALIM.CASEROS"

cbx_Dir_Comercial_Verif.AddItem "Si"
cbx_Dir_Comercial_Verif.AddItem "No"

cbx_visita_eje.AddItem "Si"
cbx_visita_eje.AddItem "No"

cbx_telef_verif.AddItem "Si"
cbx_telef_verif.AddItem "No"

cbx_direc_Part_Verif.AddItem "Si"
cbx_direc_Part_Verif.AddItem "No"

cbx_forma_verif_dir_part.AddItem "Servicios Basicos"
cbx_forma_verif_dir_part.AddItem "Visita Ejecutivo"

cbx_bien_Raiz.AddItem "Propio"
cbx_bien_Raiz.AddItem "Arrendado"
cbx_bien_Raiz.AddItem "Vive Con Familiares"

cbx_acred_bien_raiz.AddItem "No posee bien raiz"
cbx_acred_bien_raiz.AddItem "Escritura"
cbx_acred_bien_raiz.AddItem "Certificado De Dominio Vigente"
cbx_acred_bien_raiz.AddItem "Contribuciones"
cbx_acred_bien_raiz.AddItem "Comprobante Dividendo Hipotecario"
cbx_acred_bien_raiz.AddItem "Escritura Compra/Vta a Nombre Clie"

cbx_valor_evaluo_bien_raiz.AddItem "No posee bien raiz"
cbx_valor_evaluo_bien_raiz.AddItem "Menos De 10 M$"
cbx_valor_evaluo_bien_raiz.AddItem "10 a 20 M$"
cbx_valor_evaluo_bien_raiz.AddItem "20 a 30 M$"
cbx_valor_evaluo_bien_raiz.AddItem "Más De 30 M$"

cbx_acred_bien_raiz_no.AddItem "No"
cbx_valor_evaluo_bien_raiz_no.AddItem "No"
cbx_acreditacion_vehiculo_no.AddItem "No"

cbx_vehiculos_propios.AddItem "Si"
cbx_vehiculos_propios.AddItem "No"

cbx_acreditacion_vehiculo.AddItem "Certificado Anotaciones Vigentes"

cbx_antecedentes_int_bancos.AddItem "Cumple"
cbx_antecedentes_int_bancos.AddItem "No Cumple"


cbx_morosidades.AddItem "Cumple"
cbx_morosidades.AddItem "No Cumple"

cbx_protestos.AddItem "Cumple"
cbx_protestos.AddItem "No Cumple"

cbx_boletin_laboral.AddItem "Cumple"
cbx_boletin_laboral.AddItem "No Cumple"

cbx_credito_fogape.AddItem "Si"
cbx_credito_fogape.AddItem "No"

cbx_credito_fogain.AddItem "Si"
cbx_credito_fogain.AddItem "No"

cbx_numero_acreedores.AddItem "0"
cbx_numero_acreedores.AddItem "1"
cbx_numero_acreedores.AddItem "2"
cbx_numero_acreedores.AddItem "3"
cbx_numero_acreedores.AddItem "4"
cbx_numero_acreedores.AddItem "5"

cbx_mora_sbif.AddItem "Cumple"
cbx_mora_sbif.AddItem "No Cumple"
cbx_mora_sbif.AddItem "No Bancarizado"

cbx_venc_cast_SBIF.AddItem "Cumple"
cbx_venc_cast_SBIF.AddItem "No Cumple"
cbx_venc_cast_SBIF.AddItem "No Bancarizado"


cbx_Mora_Total_Sbif.AddItem "Cumple"
cbx_Mora_Total_Sbif.AddItem "No Cumple"
cbx_Mora_Total_Sbif.AddItem "No Bancarizado"

cbx_grado_formalidad.AddItem "Patente Comercial"
cbx_grado_formalidad.AddItem "Permiso Municipal"
cbx_grado_formalidad.AddItem "Iniciación De Actividades"
cbx_grado_formalidad.AddItem "Declaración De Impuestos"
cbx_grado_formalidad.AddItem "Permiso Higiene Ambiental"
cbx_grado_formalidad.AddItem "Cuaderno De Registros"
cbx_grado_formalidad.AddItem "Carpeta Tributaria"

cbx_tipo_ME.AddItem "Propia"
cbx_tipo_ME.AddItem "Negocio Familiar"
cbx_tipo_ME.AddItem "Arrendada"
cbx_tipo_ME.AddItem "Sociedad"
cbx_tipo_ME.AddItem "Otra"

cbx_tipo_cliente.AddItem "FORMALES"
cbx_tipo_cliente.AddItem "FORMAL SERVICIO O PRODUCCION"
cbx_tipo_cliente.AddItem "SEMIFORMALES"
cbx_tipo_cliente.AddItem "INFORMALES(Oficio)"

cbx_envia_cic.AddItem "Si"
cbx_envia_cic.AddItem "No"

cbx_conyuge_inf_com.AddItem "Si"
cbx_conyuge_inf_com.AddItem "No"

cbx_tiene_score.AddItem "Si"
cbx_tiene_score.AddItem "No"

cbx_tipo_excepcion.AddItem "No"
cbx_tipo_excepcion.AddItem "Cap.Trabaj Max.24 M."
cbx_tipo_excepcion.AddItem "Inversion Max.60 M."
cbx_tipo_excepcion.AddItem "Ant.Veh.Vig.Max 15 A."
cbx_tipo_excepcion.AddItem "Primera Cuota Max.90 Dias"
cbx_tipo_excepcion.AddItem "Mor.Sbif <=M$50 Tit."
cbx_tipo_excepcion.AddItem "Mor.Sbif <=M$50 Cony."
cbx_tipo_excepcion.AddItem "Mor.Comr <=M$50 Tit."
cbx_tipo_excepcion.AddItem "Mor.Comr <=M$50 Cony."
cbx_tipo_excepcion.AddItem "Numero Acre.3 a 4"
cbx_tipo_excepcion.AddItem "Ivas Pag.C/Atraso"
cbx_tipo_excepcion.AddItem "Pat.Pag.C/Atraso"

cbx_pregunta_consumo.AddItem "No"
cbx_pregunta_consumo.AddItem "Si"

cbx_Accion_consumo.AddItem "Compra Cartera"
cbx_Accion_consumo.AddItem "Libre Disponibilidad"


cbx_pregunta_comercial.AddItem "No"
cbx_pregunta_comercial.AddItem "Si"


End Sub

Private Sub txt_rut_cliente_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)


''' AVISO A CLIENTE POR CAMBIO DE DIRECCION SUCURSAL'''
''' ***************************************************
     ssql = "select exclusion_gc" _
            & " FROM TBL_CONSUMO_SIMULA_CRED" _
            & " wHERE rut = '" & txt_rut_cliente & "'"

            Set rst = cnn.Execute(ssql, , adCmdText)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If rst.EOF Then

        ElseIf rst!EXCLUSION_GC = 1 Then
                    MsgBox "¡Recuerde! Informar al cliente sobre el traslado de la sucursal Talca o San Felipe (Según corresponda)", vbExclamation
        End If

cbx_antecedentes_int_bancos = Empty
cbx_morosidades = Empty
cbx_protestos = Empty
cbx_boletin_laboral = Empty


If txt_rut_cliente <> "" Then

 cbx_numero_acreedores = 0
 
 
 '''''''''''' trae bancarizado
 
     ssql = "select rut" _
            & " FROM TBL_MICRO_INTERFAZ_Fact_LibroDeudores" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
        
           Evaluacion_Perfil.cbx_bancarizado = "No"
           txt_bancarizado_politica = "No"
           
            Else
            Evaluacion_Perfil.cbx_bancarizado = "Si"
            txt_bancarizado_politica = "Si"
        
        End If
        
 
 ''''''''verificacion RUT EN CAMPAÑA ''''''
    '''''''''' TRAE tipo cliente score  '''''''''
    ssql = "select rut FROM tbl_micro_campana " _
    & " WHERE rut = '" & txt_rut_cliente & "' and marca_vv=2"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    If rst.EOF Then
    txt_campana = "No"
    Else
    txt_campana = "Si"
    End If
 
 ''''''''verificacion RUT EN CAMPAÑA 48 MESES ENTREGADA X RIESGO MENSUALEMENTE''''''
    
    ssql = "select rut FROM tbl_micro_campana " _
    & " WHERE rut = '" & txt_rut_cliente & "' and marca_vv=1"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    If rst.EOF Then
    txt_campana_48M = "No"
    Else
    txt_campana_48M = "Si"
    End If
 
 ''''''''verificacion RUT EN BASE DE DATOS EVALUADOS ENTREGADA X RIESGO MENSUALEMENTE''''''
    
    ssql = "select rut FROM tbl_micro_campana " _
    & " WHERE rut = '" & txt_rut_cliente & "' and marca_vv=0"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    If rst.EOF Then
    txt_campana_evaluados = "No"
    Else
    txt_campana_evaluados = "Si"
    End If
 
  
''''''''''''''''''''''''''''''''''''''''''''''''
'''Revision De la datos consulta SINACOFI  CODIGO OBSERVACION TITULAR
''''''''''''''''''''''''''''''''''''''''''''''''
 
        ssql = "select top 1 rut, protesto, mora, infraccion_prev, score, convert(varchar,fecha_consulta,111) fecha_consulta, convert(varchar,getdate(),111) fecha_hoy, Cod_Observacion,convert(varchar,getdate()-1,111) fecha_menos1,convert(varchar,getdate()-1,111) fecha_hoy_menos1" _
            & " from tbl_micro_sinacofi" _
            & " where rut = '" & txt_rut_cliente & "'" _
            & " order by fecha_consulta desc" _

        Set rst = cnn.Execute(ssql, , adCmdText)
     
          If Not rst.EOF Then
        
            txt_fecha_actual_menos1_sinacofi = rst!fecha_menos1
           
        If (rst!cod_observacion = 4 Or rst!cod_observacion = 6 Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12) Then
        
            MsgBox "No se Puede Evaluar este cliente por su codigo de observación", vbCritical
            txt_rut_cliente = Empty
            txt_rut_conyuge = Empty
            txt_rut_aval = Empty
            txt_edad_conyuge = Empty
            txt_edad_aval = Empty
            txt_edad = Empty
            txt_r_edad = Empty
            
        Else
           
            If rst.EOF Then
                
                   MsgBox "No existe RUT en base de datos para comenzar Evaluación Solicite la emision de los datos", vbCritical
        
              Else
                            If rst!fecha_consulta <> rst!fecha_hoy And rst!fecha_consulta <> rst!fecha_hoy_menos1 Then

                              MsgBox "Los Datos de Sinacofi han Vencido Solicite su Emision", vbCritical
                
               Else

                             If rst!cod_observacion = 4 Or rst!cod_observacion = 6 _
                                Or rst!cod_observacion = 7 Or rst!cod_observacion = 8 Or rst!cod_observacion = 10 Or rst!cod_observacion = 12 Then
                                txt_cod_observacion = rst!cod_observacion
    
                             MsgBox "No se Puede Evaluar este cliente por su codigo de observación", vbCritical
                           
                
                 Else

                        
                        'CODIGO ACEPTADO PARA EL SISTEMA SEGUN CODIGO OBSERVACION SINACOFI
                                 
                                 
                                 If rst!cod_observacion = 2 Or rst!cod_observacion = 3 Or rst!cod_observacion = 5 Or rst!cod_observacion = 9 _
                                    Or rst!cod_observacion = 11 Or rst!cod_observacion = 13 Or rst!cod_observacion = 14 Or rst!cod_observacion = 15 Then
        
                                            cbx_morosidades = rst!mora
                                            cbx_protestos = rst!protesto
                                            cbx_boletin_laboral = rst!infraccion_prev
                                            txt_score_dicom = rst!score
                                            txt_cod_observacion = rst!cod_observacion
                                            Estado_Resolucion_Final.txt_r_f_cod_observacion_cliente = "ZG"
                                            txt_r_cod_observacion = "ZG"
              
                
                                    '''MORA CLIENTE
                                        If rst!mora = "Cumple" Then
                                       Estado_Resolucion_Final.txt_r_f_morosidad_sinac = "A"
                
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_morosidad_sinac = "R"
                
                                        End If
                
                                        '''PROTESTO CLIENTE
                   
                                        If rst!protesto = "Cumple" Then
                                        Estado_Resolucion_Final.txt_r_f_protesto_sinac = "A"
                
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_protesto_sinac = "R"
                
                                        End If
                   
                                        '''BOLETIN CLIENTE
                   
                                        If rst!infraccion_prev = "Cumple" Then
                                        Estado_Resolucion_Final.txt_r_f_boletin_sinac = "A"
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_boletin_sinac = "R"
                
                                    End If
                                            
                
                Else
                        'codigo cero CODIGO SINACOFI
                                    If rst!cod_observacion = 0 Then
                                        cbx_morosidades = rst!mora
                                        cbx_protestos = rst!protesto
                                        cbx_boletin_laboral = rst!infraccion_prev
                                        txt_score_dicom = rst!score
                                        txt_cod_observacion = rst!cod_observacion
                                        Estado_Resolucion_Final.txt_r_f_cod_observacion_cliente = "A"
                                        txt_r_cod_observacion = "A"
              
                                        txt_score_dicom_cliente_aux = rst!score
                
                                    '''MORA CLIENTE
                                        If rst!mora = "Cumple" Then
                                       Estado_Resolucion_Final.txt_r_f_morosidad_sinac = "A"
                
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_morosidad_sinac = "R"
                
                                        End If
                
                                        '''PROTESTO CLIENTE
                   
                                        If rst!protesto = "Cumple" Then
                                        Estado_Resolucion_Final.txt_r_f_protesto_sinac = "A"
                
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_protesto_sinac = "R"
                
                                        End If
                   
                                        '''BOLETIN CLIENTE
                   
                                        If rst!infraccion_prev = "Cumple" Then
                                        Estado_Resolucion_Final.txt_r_f_boletin_sinac = "A"
                                        Else
                                        Estado_Resolucion_Final.txt_r_f_boletin_sinac = "R"
                
                                    End If
                                       
        
            End If
           End If
        End If
        End If
    End If
 
 
    txt_edad.Locked = False
    cbx_estado_civil.Locked = False
 

''''-------------------------------
   
  

    txt_credito_comercial_vigente_mora = Empty
   
    txt_estado_credito = Empty
   
    txt_n_carpeta_tributaria = Empty
    txt_dv = Empty
    
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
    

'''apertura de la BASE DE DATOS
''''----------------------------
    Call conectarBD
    
    
''''Verificacion de RUT en la base de las cuotas pagadas
'''-----------------------------------------------------
    
    ssql = "select rut_cliente FROM TBL_MICRO_CUOTA_PAG_MAYOR_240 " _
    & " WHERE CAST(SUBSTRING(rut_cliente,1,9) AS INT) = '" & txt_rut_cliente & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
    
    If rst.EOF Then
        Estado_Resolucion_Final.lbl_aviso_resolucion_final.Visible = False
        Estado_Resolucion_Final.txt_cuerpo_aviso.Visible = False
        Estado_Resolucion_Final.txt_r_f_aviso_inconsis_cuota = "A"
    Else
        Estado_Resolucion_Final.lbl_aviso_resolucion_final.Visible = True
        Estado_Resolucion_Final.txt_cuerpo_aviso.Visible = True
        Estado_Resolucion_Final.txt_r_f_aviso_inconsis_cuota = "ZG"
    End If
  

'''Actualizacion de BANCARIZADO POLITICA MAIL DE C.BARRIOS 07-03-2012
'''Segunda modificacion con campo TXT_CLIENTE_NUEVO por MARIO SAN CRISTOBAL


    ssql = "select RUT" _
            & " FROM TBL_MICRO_MACA_CLIENTE_NO_CLIENTE_viG" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
            Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
            Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
 
          Else
          
            ssql = "select RUT" _
            & " FROM tbl_micro_cliente_antiguo_nuevo" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)
    
            If Not rst.EOF Then
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "No"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "No"
            Else
                Evaluacion_Perfil.cbx_Cliente_Nuevo = "Si"
                Ficha_Cliente_Micro.txt_Cliente_Nuevo = "Si"
    End If
End If


'------- Moras--sbif VIGENTE------------------------------------

    ssql = "select m_deudadirectamorosa" _
        & " FROM TBL_MICRO_FACT_ODS_LIBRODEUDORES " _
        & " WHERE rut = '" & txt_rut_cliente & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)

    If (rst.EOF) Then
          
        cbx_mora_sbif = "Cumple"
        
        ElseIf rst!m_deudadirectamorosa > 0 Then
        cbx_mora_sbif = "No Cumple"
        
        ElseIf rst!m_deudadirectamorosa = 0 Then
        cbx_mora_sbif = "Cumple"
    End If
    
    
'------- Moras--sbif ultimos 12 mese------------------------------------


    ssql = "select mora30, venc, cast, indcast, indvenc" _
        & " FROM tbl_micro_sbif " _
        & " WHERE rut = '" & txt_rut_cliente & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)

        
If rst.EOF Then
    'cbx_mora_sbif = "Cumple"
    cbx_venc_cast_SBIF = "Cumple"
    cbx_Mora_Total_Sbif = "Cumple"
    cbx_venc_cast_SBIF_indirecta = "Cumple"
    cbx_Mora_Total_Sbif_indirecta = "Cumple"
    cbx_antecedentes_int_bancos = "Cumple"
    
    cbx_numero_acreedores = 0
    txt_credito_comercial_vigente_mora = "No"
    
Else
   
    If rst!venc = 1 Then
    cbx_venc_cast_SBIF = "No Cumple"
    Else
    cbx_venc_cast_SBIF = "Cumple"
    End If
    '
    
    If rst!cast = 1 Then
    cbx_Mora_Total_Sbif = "No Cumple"
    Else
    cbx_Mora_Total_Sbif = "Cumple"
    End If
    
    '''''DEUDAS INDIRECTAS ULTIMOS 12 MESES
    
    If rst!indcast = 1 Then
    cbx_venc_cast_SBIF_indirecta = "No Cumple"
    Else
    cbx_venc_cast_SBIF_indirecta = "Cumple"
    End If
    '
    
    If rst!indvenc = 1 Then
    cbx_Mora_Total_Sbif_indirecta = "No Cumple"
    Else
    cbx_Mora_Total_Sbif_indirecta = "Cumple"
    End If
    '
End If


    ''''RENEGOCIADOs
    ''''''''''''''''''''''''''''''
    
    ssql = "select rut" _
            & " FROM TBL_MICRO_RIESGO_RENEGOCIADO" _
            & " wHERE rut = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            txt_renegociado = "No"
    
          Else
            txt_renegociado = "Si"
    
        End If


'''''TRAE LAS LINEAS DE CREDITOS CONSUMO

    ssql = "select rut" _
            & " FROM TBL_MICRO_FACT_RIESGO_MORA_DIA" _
            & " wHERE rut = '" & txt_rut_cliente & "'" _
            & " AND tipo_deudacartamensual = 'Línea de Crédito - Consumo'"
            
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            txt_linea_consumo = "No"
    
          Else
            txt_linea_consumo = "Si"
    
        End If


'''''TRAE LAS LINEAS DE CREDITOS CONSUMO

        ssql = "select rut" _
            & " FROM TBL_MICRO_FACT_RIESGO_MORA_DIA" _
            & " wHERE rut = '" & txt_rut_cliente & "'" _
            & " AND tipo_deudacartamensual = 'Línea de Crédito - Comercial'"
            
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            txt_linea_comercial = "No"
    
          Else
            txt_linea_comercial = "Si"
    
        End If
        
        
    '''''' PROTESTO INTERNO - FILENEG - FRAUDE TITULAR
    ssql = "select rut10" _
    & " from tbl_micro_RIESGO_FILEN_PROT_FRAUDE" _
    & " where CAST(SUBSTRING(rut10,1,9) AS INT) = '" & txt_rut_cliente & "'" _
    & " and fltro_cod ='R002'"
                       
    Set rst = cnn.Execute(ssql, , adCmdText)
                                                       
    If rst.EOF Then
    txt_r_file_negativo_tit = "A"

    Else
    txt_r_file_negativo_tit = "R"
    
    End If
 
    
'''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---MORA ----
    
    ssql = "select * " _
            & " FROM TBL_MICRO_fact_RIESGO_MORA_DIA" _
            & " wHERE rut = '" & txt_rut_cliente & "'" _
            & " AND diasmora >0 "
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            txt_r_mora_directa_interna = "A"
            
        Else
        
            txt_r_mora_directa_interna = "R"
        End If
    
    
'''''' CASTIGOS TRAIDOS DESDE SERVIDOR RIESGO CLIENTE
          
          ssql = "select cas_rut" _
            & " FROM tbl_micro_RIESGO_CASTIGOS" _
            & " wHERE cas_rut = '" & txt_rut_cliente & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        If rst.EOF Then
            txt_r_castigo_directo_interna = "A"
            
        Else
        
            txt_r_castigo_directo_interna = "R"
        End If
              
' *************************************************************************************************mod
    '''' TRAE Y CALCULA DEUDAS INTERNAS BANCO ---VENCI-CASTIGO
    
      'ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE CAST(SUBSTRING(rut,1,9) AS INT) = '" & txt_rut_cliente & "'"
    
      ssql = "select dias_mora" _
            & " FROM TBL_MICRO_MORA_MAX_ULT_12M" _
            & " wHERE rut_num = '" & txt_rut_cliente & "'"
    
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    
            If rst.EOF Then
            txt_r_Vencido_directo_interna = "A"
            
            ElseIf rst!dias_mora = 0 Then
            txt_r_Vencido_directo_interna = "A"
            
            ElseIf rst!dias_mora <= 30 Then
            txt_r_Vencido_directo_interna = "A"
            
            ElseIf rst!dias_mora >= 31 Then
            txt_r_Vencido_directo_interna = "A"
            
            ElseIf rst!dias_mora >= 91 Then
            txt_r_Vencido_directo_interna = "R"
            
            ElseIf rst!dias_mora >= 181 Then
            txt_r_Vencido_directo_interna = "R"

   
        End If
    
    
    
''''CODIFICACION DE CODIGOS DE APROBACION RECHAZO DE DEUDAS INTERNAS
             
If txt_r_mora_directa_interna = "R" Or txt_r_Vencido_directo_interna = "R" Or txt_r_castigo_directo_interna = "R" Then
        
    cbx_antecedentes_int_bancos = "No Cumple"
    txt_r_aib = "R"
        
ElseIf txt_r_mora_directa_interna = "A" And txt_r_Vencido_directo_interna = "A" And txt_r_castigo_directo_interna = "A" Then

    cbx_antecedentes_int_bancos = "Cumple"
    txt_r_aib = "A"

End If
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''
'''''  NUMEROS DE ACRREEDORES
'''''''''''''''''''''''''''''''''''''''''''''
    
                ssql = "select rut,n_institucionescondeuda+m_deudacreditocomerciales n_acreedores" _
                    & " from tbl_micro_fact_ods_librodeudores" _
                    & " where rut = '" & txt_rut_cliente & "'"
    
                Set rst = cnn.Execute(ssql, , adCmdText)
    
                If rst.EOF Then
                    cbx_numero_acreedores = ""
                    cbx_numero_acreedores = 0
                Else
                    cbx_numero_acreedores = rst!n_acreedores
                End If
    
            ssql1a = "select rut" _
                    & " from TBL_MICRO_FACT_ODS_D10" _
                    & " where rut = '" & txt_rut_cliente & "'"
                
                Set rst1a = cnn.Execute(ssql1a, , adCmdText)
            
                If rst1a.EOF Or cbx_numero_acreedores = 0 Then
                      'If rst1a.EOF Or rst!n_acreedores = 0 Or cbx_numero_acreedores = "0" Then
                      cbx_numero_acreedores = cbx_numero_acreedores + 1

        
   End If
    
''''''''''''''''''  CALCULA EDAD DE CLIENTE
    
    
    
    ssql = "select cliente,datediff(MONTH,f_nacimiento, getdate())/12 edad,f_nacimiento" _
            & " from TBL_MICRO_FACT_MACA_CLIENTE" _
            & " where cliente = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
    

If rst.EOF Then
        
    txt_edad.Locked = False
    MsgBox "Ingrese la fecha de nacimiento si el Cliente es Persona Natural", vbCritical
    txt_edad.Locked = False
    txt_fechanacimiento = "01/01/1900"
    
    txt_fechanacimiento.Locked = False
    
    ElseIf rst!f_nacimiento = "01/01/1900" And rst!cliente < 45000000 Then
    MsgBox "No Registramos Correctamente la Fecha de Nacimiento Del Cliente, Ingreselo... ", vbCritical
    txt_edad.Locked = False
    txt_fechanacimiento = "01/01/1900"
    txt_fechanacimiento.Locked = False
    'txt_edad = rst!edad
    Else
        txt_edad = rst!edad
        txt_fechanacimiento = rst!f_nacimiento
        txt_edad.Locked = True
        txt_fechanacimiento.Locked = True
    
    
End If

'''''''''''''''TRAE CREDITO_COMECIAL_CUOTA

        ssql = "select rut " _
                & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                & " where (tipo_deudacartamensual = 'Crédito Comercial' or tipo_deudacartamensual = 'Crédito Comercial (DFC)')" _
                & " and rut = '" & txt_rut_cliente & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)

        If rst.EOF Then
            txt_cred_comer_cuota = "No"
            Else
            txt_cred_comer_cuota = "Si"
        End If


'''''''''''''''CREDITO VIGENTE EN MORA

            ssql = "select rut " _
                & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                & " where rut = '" & txt_rut_cliente & "' and diasmora >0 "
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_credito_comercial_vigente_mora = "No"
                Else
                txt_credito_comercial_vigente_mora = "Si"
            End If




'''''''''''''''TRAE CREDITO_EDUCACION

            ssql = "select rut " _
                & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                & " where (tipo_deudacartamensual = 'Crédito Educación L.20027 (Comision)' or tipo_deudacartamensual = 'Crédito Educación L.20027 (Vendido)'" _
                & " or tipo_deudacartamensual = 'Crédito Educación L.20027 (Vendido) (D)' or tipo_deudacartamensual = 'Crédito Educación Ley 20027'" _
                & " or tipo_deudacartamensual = 'Crédito Educación')" _
                & " and rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_cred_educacion = "No"
                Else
                txt_cred_educacion = "Si"
            End If
    

'''''''''''''''TRAE CREDITO_credito_en_cuotas
    
            ssql = "select rut " _
                    & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                    & " where (tipo_deudacartamensual = 'Crédito en Cuotas' or tipo_deudacartamensual = 'Crédito en Cuotas (DFC)')" _
                    & " and rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_consumo_cuota = "No"
                Else
                txt_consumo_cuota = "Si"
            End If
    
    
'''''''''''''''TRAE CREDITO_credito hipotecario
    
            ssql = "select rut " _
                    & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                    & " where (tipo_deudacartamensual = 'Hipotecario Fines Generales' or tipo_deudacartamensual = 'Hipotecario Vivienda')" _
                    & " and rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_cred_hipotecario = "No"
                Else
                txt_cred_hipotecario = "Si"
            End If
    
    
'''''''''''''''TRAE tarjeta credito
    
            ssql = "select rut " _
                    & " from TBL_MICRO_FACT_RIESGO_MORA_DIA" _
                    & " where (tipo_deudacartamensual = 'Tarjeta de Crédito - Comercial' or tipo_deudacartamensual = 'Tarjeta de Crédito - Consumo'" _
                    & " or tipo_deudacartamensual = 'Renegociado - Crédito Educación')" _
                    & " and rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_tarjeta_cred = "No"
                Else
                txt_tarjeta_cred = "Si"
            End If
    
    
'''''''''''''''TRAE creditos varios protestos internos
    
            ssql = "select rut " _
                    & " from tbl_micro_maca_protestos" _
                    & " where tipo = 'P'" _
                    & " and rut = '" & txt_rut_cliente & "'"
    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                txt_protesto_interno = "No"
                Else
                txt_protesto_interno = "Si"
            End If
      
      
''''''''''''''CASTIGOS HISTORICOS CLIENTE-TITULAR

            ssql = "select rut " _
                    & " from tbl_micro_maca_castigos_his" _
                    & " where rut = '" & txt_rut_cliente & "'"
                    
            Set rst = cnn.Execute(ssql, , adCmdText)

            If rst.EOF Then
                Estado_Resolucion_Final.txt_r_f_castigo_historico = "A"
                Else
                Estado_Resolucion_Final.txt_r_f_castigo_historico = "R"
            End If
      
      
'''' inhibe campos si es persona juridica

If txt_rut_cliente >= 45000000 Then
    txt_edad.Locked = True
    cbx_estado_civil.Locked = True
    txt_edad = Empty
    txt_r_edad = Empty
   
End If
    
    
    End If ''''''''''''''' CIERRE DE PRIMERA CONDICION DE INGRESO
        
        Else
        MsgBox "No existe Rut Ingresado en Sistema Local Sinacofi", vbCritical

End If ''''''''''''''' CIERRE DE PRIMERA CONDICION DE INGRESO
End If


End Sub
