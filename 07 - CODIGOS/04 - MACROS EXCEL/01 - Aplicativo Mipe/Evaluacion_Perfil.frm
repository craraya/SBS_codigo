VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Evaluacion_Perfil 
   Caption         =   ":::Evaluacion Perfil De Cliente"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   OleObjectBlob   =   "Evaluacion_Perfil.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Evaluacion_Perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbx_bancarizado_AfterUpdate()
txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False
End Sub

Private Sub cbx_bancarizado_Change()

txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False

End Sub

Private Sub cbx_Cliente_Nuevo_AfterUpdate()

txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
cmd_metodologia_iva.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False



  lbl_antiguedad.Visible = False
'  txt_Antiguedad.Visible = False

  lbl_historia_pago.Visible = False
  cbx_historia_pago.Visible = False
  lbl_mora_maxima.Visible = False
  cbx_mora_maxima.Visible = False
  txt_registro_ventas.Visible = False
  txt_antiguedad_negocio.Visible = False
  txt_predictor_Score.Visible = False
  txt_tipo_cliente_actividad.Visible = False
  cbx_actividad_economica_formal.Visible = False
   lbl_registro_ventas.Visible = False
   lbl_antiguedad_negocio.Visible = False
   lbl_predictor_Score.Visible = False
   lbl_tipo_cliente_actividad.Visible = False
   lbl_actividad_economica.Visible = False
    lbl_porcentaje_r_ventas.Visible = False
    lbl_meses_antiguedad_negocio.Visible = False
    lbl_mayor_a_prodictor.Visible = False

    
If cbx_Cliente_Nuevo = "Si" Then
   lbl_antiguedad.Visible = False
'   txt_Antiguedad.Visible = False

   lbl_historia_pago.Visible = False
   cbx_historia_pago.Visible = False
   lbl_mora_maxima.Visible = False
   cbx_mora_maxima.Visible = False

  
ElseIf cbx_Cliente_Nuevo = "No" Then
     
'   lbl_antiguedad.Visible = True
'   txt_Antiguedad.Visible = True

   lbl_historia_pago.Visible = True
   cbx_historia_pago.Visible = True
   lbl_mora_maxima.Visible = True
   cbx_mora_maxima.Visible = True


  
End If
End Sub

Private Sub cbx_Cliente_Nuevo_Change()

'cmd_calcular_tipo_cliente.Enabled = False
txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
cmd_metodologia_iva.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False

Evaluacion_Perfil.txt_rut_cliente = Ficha_Cliente_Micro.txt_rut_cliente

  lbl_antiguedad.Visible = False
'  txt_Antiguedad.Visible = False
'  lbl_meses_antiguedad_banco.Visible = False
  lbl_historia_pago.Visible = False
  cbx_historia_pago.Visible = False
  lbl_mora_maxima.Visible = False
  cbx_mora_maxima.Visible = False
  txt_registro_ventas.Visible = False
  txt_antiguedad_negocio.Visible = False
  txt_predictor_Score.Visible = False
  txt_tipo_cliente_actividad.Visible = False
  cbx_actividad_economica_formal.Visible = False
   lbl_registro_ventas.Visible = False
   lbl_antiguedad_negocio.Visible = False
   lbl_predictor_Score.Visible = False
   lbl_tipo_cliente_actividad.Visible = False
   lbl_actividad_economica.Visible = False
    lbl_porcentaje_r_ventas.Visible = False
    lbl_meses_antiguedad_negocio.Visible = False
    lbl_mayor_a_prodictor.Visible = False
'   lbl_Numero_Credito_Cli_Antiguo.Visible = False
'   txt_n_credito_cli_antiguo.Visible = False
    
If cbx_Cliente_Nuevo = "Si" Then
   lbl_antiguedad.Visible = False
'   txt_Antiguedad.Visible = False
'   lbl_meses_antiguedad_banco.Visible = False
   lbl_historia_pago.Visible = False
   cbx_historia_pago.Visible = False
   lbl_mora_maxima.Visible = False
   cbx_mora_maxima.Visible = False

  
ElseIf cbx_Cliente_Nuevo = "No" Then
     
'   lbl_antiguedad.Visible = True
'   txt_Antiguedad.Visible = True
'   lbl_meses_antiguedad_banco.Visible = True
   lbl_historia_pago.Visible = True
   cbx_historia_pago.Visible = True
   lbl_mora_maxima.Visible = True
   cbx_mora_maxima.Visible = True
'   lbl_Numero_Credito_Cli_Antiguo.Visible = True
'   txt_n_credito_cli_antiguo.Visible = True

  
End If
End Sub

Private Sub cbx_historia_pago_AfterUpdate()

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False


R_Final_Perfil = Empty

'''****************************************************
'''Cambio realizador por mail de Mario San Cristobal
'''****************************************************
''''MORA PROMEDIO
If (cbx_historia_pago = "Hasta 5 dias" Or cbx_historia_pago = "Sin Mora") And (cbx_mora_maxima = "Sin Mora" Or cbx_mora_maxima = "Hasta 5 dias" Or cbx_mora_maxima = "Hasta 7 dias") Then
   txt_r_historia_Pago = "Excelente"

ElseIf (cbx_historia_pago = "Hasta 5 dias" Or cbx_historia_pago = "Sin Mora" Or cbx_historia_pago = "Hasta 7 dias") And (cbx_mora_maxima = "Sin Mora" Or cbx_mora_maxima = "Hasta 5 dias" Or cbx_mora_maxima = "Hasta 7 dias" Or cbx_mora_maxima = "Hasta 15 dias") Then
   txt_r_historia_Pago = "Bueno"

''ElseIf cbx_historia_pago = "Hasta 15 dias" Then
  Else
    txt_r_historia_Pago = "Regular"

End If

'cmd_calcular_tipo_cliente.Enabled = False
txt_tipo_cliente = Empty
End Sub

Private Sub cbx_mora_maxima_AfterUpdate()

txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False

R_Final_Perfil = Empty

If cbx_mora_maxima = "Hasta 7 dias" Or cbx_mora_maxima = "Sin Mora" Then
    txt_r_mora = "Excelente"

ElseIf cbx_mora_maxima = "Hasta 15 dias" Then
    txt_r_mora = "Bueno"
    
ElseIf cbx_mora_maxima = "Hasta 30 dias" Then
     txt_r_mora = "Regular"

End If

'If cbx_mora_maxima <> "" Then
'cmd_calcular_tipo_cliente.Enabled = True
'End If
End Sub

Private Sub cmd_calcular_tipo_cliente_Click()

Dim MarcarZG As Integer


Call conectarBD


''''''''''  NUERO DE CUOTAS PAGADAS

'            ssql = "SELECT rut, ULTIMACUOTAPAGADA" _
'        & " FROM tbl_micro_maca_operaciones_his" _
'        & " where cast(substring(rut,1,9) as integer) = '" & Ficha_Cliente_Micro.txt_rut_cliente & "'"
              
'    Set rst = cnn.Execute(ssql, , adCmdText)
              
'        If rst.EOF Then
'            Evaluacion_Perfil.txt_Antiguedad = 0
'          Else
'            Evaluacion_Perfil.txt_Antiguedad = rst!ultimacuotapagada
'        End If
        


    ''''' TRAE CLIENTE prime NOPRIME
    
    ssql = "SELECT rut, flag_mp, flag_mm" _
        & " FROM TBL_MICRO_RESUMEN_PRIME_NOPRIME" _
        & " where rut = '" & Ficha_Cliente_Micro.txt_rut_cliente & "'"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
              
        If rst.EOF Then
            Evaluacion_Perfil.cbx_historia_pago = "Sin Mora"
            Evaluacion_Perfil.cbx_mora_maxima = "Sin Mora"
          Else
            Evaluacion_Perfil.cbx_mora_maxima = rst!flag_mm
            Evaluacion_Perfil.cbx_historia_pago = rst!flag_mp

        End If
    
    ''''TRAE CUOTAS PAGADAS BD
    
       ' ssql = "SELECT rut,UltimaCuotaPagada" _
        '& " FROM TBL_micro_MACA_OPERACIONES" _
        '& " where rut = '" & Ficha_Cliente_Micro.txt_rut_cliente & "'"
        
        'Set rst = cnn.Execute(ssql, , adCmdText)
        
        'If rst.EOF Then
        'txt_Antiguedad = 0
        'Else
        'txt_Antiguedad = rst!UltimaCuotaPagada
        'End If
        
        
        
'''' *****************Se llena nuevamente con datos refrescados desde FACT_OPERACIONES (UNICA VEZ)
''*********************************
''' BASE DE DATOS COMPLEMENTO FACT_MORA_DIARIA
''*********************************
        ssql = "SELECT MARCA_REGULAR" _
        & " FROM TBL_MICRO_RESUMEN_CRITICAL_NEW" _
        & " where rut_numerico = '" & Ficha_Cliente_Micro.txt_rut_cliente & "'"
        
        Set rst = cnn.Execute(ssql, , adCmdText)
            
        If Not rst.EOF Then
        
        txt_tipo_cliente_duro = rst!marca_regular
        Evaluacion_Perfil.txt_cliente_ex_critical = 1
        Else
        
        txt_tipo_cliente_duro = 0
        Evaluacion_Perfil.txt_cliente_ex_critical = 0
        End If
   
   
        If cbx_Cliente_Nuevo = "No" And (cbx_historia_pago = "" Or cbx_mora_maxima = "") Then
        'If cbx_Cliente_Nuevo = "No" And (txt_Antiguedad = "" Or cbx_historia_pago = "" Or cbx_mora_maxima = "") Then
        MsgBox ("Faltan Datos que ingresar para Cliente Antiguo... Favor Revisar")
        txt_tipo_cliente = Empty

        ElseIf txt_r_historia_Pago = "Excelente" And txt_r_mora = "Excelente" Then
        txt_r_pago_mora = "Excelente"

        ElseIf txt_r_historia_Pago = "Excelente" And txt_r_mora = "Bueno" Then
        txt_r_pago_mora = "Bueno"
    
        ElseIf txt_r_historia_Pago = "Excelente" And txt_r_mora = "Regular" Then
        txt_r_pago_mora = "Regular"
    
        ElseIf txt_r_historia_Pago = "Bueno" And txt_r_mora = "Bueno" Then
        txt_r_pago_mora = "Bueno"
    
        ElseIf txt_r_historia_Pago = "Bueno" And txt_r_mora = "Regular" Then
        txt_r_pago_mora = "Regular"
    
        ElseIf txt_r_historia_Pago = "Bueno" And txt_r_mora = "Excelente" Then
        txt_r_pago_mora = "Bueno"
    
        ElseIf txt_r_historia_Pago = "Regular" And txt_r_mora = "Regular" Then
        txt_r_pago_mora = "Regular"
    
        ElseIf txt_r_historia_Pago = "Regular" And txt_r_mora = "Bueno" Then
        txt_r_pago_mora = "Regular"
    
        ElseIf txt_r_historia_Pago = "Regular" And txt_r_mora = "Excelente" Then
        txt_r_pago_mora = "Regular"
    
''''''''''''
  
    End If
'End If

'PASO DE VARIBLES FORMULARIO CLIENTES
'-------------------------------------------
txt_rut_cliente = rut_cliente_ficha
txt_dv = dv_cliente_ficha
txt_n_carpeta_tributaria = n_carpeta_tributaria_ficha
txt_tipo_cliente_actividad = formalidad_negocio_ficha
txt_predictor_Score = score_dicom_ficha
txt_antiguedad_negocio = antiguedad_meses_ficha
cbx_actividad_economica_formal = r_actividad_ficha
cbx_actividad_economica_semiformal = r_actividad_ficha
cbx_actividad_economica_formal_servicio = r_actividad_ficha
cbx_actividad_economica_informal_oficio = r_actividad_ficha



'MOSTRAR COMBOBOX ACTIVIDADES
'------------------------------
txt_r_tipo_cliente_actividad = Empty

If txt_tipo_cliente_actividad = "FORMALES" Then
cbx_actividad_economica_formal.Visible = True
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = False
cbx_actividad_economica_formal_servicio.Visible = False

ElseIf txt_tipo_cliente_actividad = "FORMAL SERVICIO O PRODUCCION" Then
cbx_actividad_economica_formal_servicio.Visible = True
cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = False

ElseIf txt_tipo_cliente_actividad = "SEMIFORMALES" Then
cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_semiformal.Visible = True
cbx_actividad_economica_informal_oficio.Visible = False
cbx_actividad_economica_formal_servicio.Visible = False

ElseIf txt_tipo_cliente_actividad = "INFORMALES(Oficio)" Then
cbx_actividad_economica_formal.Visible = False
cbx_actividad_economica_semiformal.Visible = False
cbx_actividad_economica_informal_oficio.Visible = True
cbx_actividad_economica_formal_servicio.Visible = False
End If
txt_r_tipo_cliente_actividad = txt_tipo_cliente_actividad


'-------------------------------------------------------------


If cbx_Cliente_Nuevo = "Si" And cbx_bancarizado = "Si" Then
  lbl_antiguedad.Visible = False
'  txt_Antiguedad.Visible = False
  cbx_historia_pago.Visible = False
  cbx_mora_maxima.Visible = False
  lbl_historia_pago.Visible = False
  lbl_mora_maxima.Visible = False
'  lbl_Numero_Credito_Cli_Antiguo.Visible = False
'  lbl_meses_antiguedad_banco.Visible = False
  
  txt_tipo_cliente = "Nuevo Con Historia Sbif"
  'txt_tipo_cliente = "Nuevo Bancarizado"
  
ElseIf cbx_Cliente_Nuevo = "Si" And cbx_bancarizado = "No" Then
  lbl_antiguedad.Visible = False
'  txt_Antiguedad.Visible = False
  cbx_historia_pago.Visible = False
  cbx_mora_maxima.Visible = False
  lbl_historia_pago.Visible = False
  lbl_mora_maxima.Visible = False
'  lbl_Numero_Credito_Cli_Antiguo.Visible = False
'  lbl_meses_antiguedad_banco.Visible = False
  
  txt_tipo_cliente = "Nuevo Sin Historia Sbif"
  'txt_tipo_cliente = "Nuevo No Bancarizado"
  
  
  End If
  
  lbl_registro_ventas.Visible = True
  txt_registro_ventas.Visible = True
'  lbl_antiguedad_negocio.Visible = True
  txt_antiguedad_negocio.Visible = True
  lbl_predictor_Score.Visible = True
  txt_predictor_Score.Visible = True
  lbl_tipo_cliente_actividad.Visible = True
  txt_tipo_cliente_actividad.Visible = True
  lbl_actividad_economica.Visible = True
'  cbx_actividad_economica_formal.Visible = True
  lbl_porcentaje_r_ventas.Visible = True
  lbl_meses_antiguedad_negocio.Visible = True
'  lbl_mayor_a_prodictor.Visible = True
  
  
'******************************************************************************************************
'VARIABLE Y CAMPO EN SQL   <TXT_ANTIGUEDAD>  CORRESPONDE A NUMERO DE CUOTAS PAGADAS EN BD
'(CAMBIO SEGUN MAIL DE C.BARRIOS 01-08-2011)
'******************************************************************************************************
    
'***********************************************************************
'''SE COMENTA POR MAIL DE AJUSTES RIESGO CON MARIO SAN CRISTOBAL
'***********************************************************************


''' trae APARICIONES DE FACT_OPERACIONES PARA DETERMINAR EL CLIENTE PRIME

        ssql = "select rut,n_apariciones" _
       & " from tbl_micro_aparicion_prime" _
       & " where cast(substring(rut,1,9) as integer) = '" & txt_rut_cliente & "'"
        
        Set rst = cnn.Execute(ssql, , adCmdText)
            
        If Not rst.EOF Then
            Evaluacion_Perfil.txt_apariciones_fact_ope = rst!n_apariciones
        Else
            Evaluacion_Perfil.txt_apariciones_fact_ope = 0
        End If

''' trae MORA DIARIA DE CLIENTE

        ssql = "select rut,numero_de_apariciones,max_mora_total,promedio" _
       & " from TBL_MICRO_RESUMEN_PRIME_NOPRIME" _
       & " where rut = '" & txt_rut_cliente & "'"
        
        Set rst = cnn.Execute(ssql, , adCmdText)
            
        If Not rst.EOF Then
            Evaluacion_Perfil.txt_apariciones_mora_dia = rst!numero_de_apariciones
            Evaluacion_Perfil.txt_mora_maxima = rst!max_mora_total
            Evaluacion_Perfil.txt_mora_promedio = rst!promedio
        Else
            Evaluacion_Perfil.txt_apariciones_mora_dia = 0
            Evaluacion_Perfil.txt_mora_maxima = 0
            Evaluacion_Perfil.txt_mora_promedio = 0
        
        End If
       
''********************************************
''''' determina el TIPO DE CLIENTE <<<<ANTIGUO PRIME>>>
''********************************************

If cbx_Cliente_Nuevo = "No" Then

    If txt_apariciones_fact_ope >= 18 Then
    
        If txt_apariciones_mora_dia >= 12 And txt_mora_promedio <= 5 And txt_mora_maxima <= 7 Then
            txt_tipo_cliente = "Antiguo Prime"
        ElseIf txt_apariciones_mora_dia >= 12 And txt_mora_promedio <= 7 And txt_mora_maxima <= 15 Then
            txt_tipo_cliente = "Antiguo Prime"
        ElseIf txt_apariciones_mora_dia >= 1 And txt_apariciones_mora_dia <= 11 And txt_mora_promedio <= 5 And txt_mora_maxima <= 7 And txt_cliente_ex_critical = 0 Then
            txt_tipo_cliente = "Antiguo Prime"
        ElseIf txt_apariciones_mora_dia >= 1 And txt_apariciones_mora_dia <= 11 And txt_mora_promedio <= 7 And txt_mora_maxima <= 15 And txt_cliente_ex_critical = 0 Then
            txt_tipo_cliente = "Antiguo Prime"
        Else
        txt_tipo_cliente = "Antiguo No Prime"
        
        End If
        Else
        txt_tipo_cliente = "Antiguo No Prime"
    End If
End If
    
    
    lbl_registro_ventas.Visible = True
    txt_registro_ventas.Visible = True
    lbl_antiguedad_negocio.Visible = False
    txt_antiguedad_negocio.Visible = False
    lbl_predictor_Score.Visible = True
    txt_predictor_Score.Visible = True
    lbl_tipo_cliente_actividad.Visible = True
    txt_tipo_cliente_actividad.Visible = True
    lbl_actividad_economica.Visible = True
'    cbx_actividad_economica_formal.Visible = True
    lbl_meses_antiguedad_negocio.Visible = False
    
        
    If txt_tipo_cliente = "Antiguo Prime" Then
       
    lbl_registro_ventas.Visible = False
    lbl_antiguedad_negocio.Visible = False
    lbl_predictor_Score.Visible = True
    lbl_tipo_cliente_actividad.Visible = False
    lbl_actividad_economica.Visible = False
    
    End If
    
    cmd_calcular_perfil.Enabled = True
    
 'End If
  
'''cambio solicitado por VIVIANA MANRIQUEZ Mail 14-10-2013
 'txt_rut_cliente = 12663566
 ' txt_tipo_cliente = "Nuevo Sin Historia Sbif"
 ' txt_campana = "No"
 ' txt_score_dicom = 0
 ' txt_campana_48M = "No"
 ' txt_campana_evaluados = "No"

'---NO BANCARIZADOS MAIL DE VIVIANA MANRIQUEZ

'---filtro inicial
If Ficha_Cliente_Micro.txt_bancarizado_politica = "Si" Then
  
    If (txt_tipo_cliente = "Nuevo Sin Historia Sbif" Or txt_tipo_cliente = "Nuevo Con Historia Sbif") And Ficha_Cliente_Micro.txt_campana = "No" _
    And Ficha_Cliente_Micro.txt_score_dicom < 790 Then
    
    txt_r_dicom_tipo_cliente = "R"
    MsgBox "El cliente no esta en campaña y no cumple con score sinacofi", vbCritical

    Else
 
        If Ficha_Cliente_Micro.txt_rut_cliente < 45000000 Then
            
            If txt_tipo_cliente = "Nuevo Con Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom < 535 Then
                txt_r_dicom_tipo_cliente = "R"
        
                    'ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom <= 100 _
                     '       And Ficha_Cliente_Micro.txt_score_dicom <> 0 Then
                        
                        'txt_r_dicom_tipo_cliente = "R"
        
                    'ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom = 0 Then
                    '    txt_r_dicom_tipo_cliente = "ZG"
 
                    ElseIf txt_tipo_cliente = "Antiguo No Prime" And Ficha_Cliente_Micro.txt_score_dicom < 409 Then
                        txt_r_dicom_tipo_cliente = "R"
          
                    ElseIf txt_tipo_cliente = "Antiguo Prime" And Ficha_Cliente_Micro.txt_score_dicom < 347 Then
                        txt_r_dicom_tipo_cliente = "R"
    
                    Else
                        txt_r_dicom_tipo_cliente = "A"

              End If
    
         Else
             txt_r_dicom_tipo_cliente = "ZG"

      End If

    End If


'..................................
'Rutina Nuevo para NO BANCARIZADOS.
'..................................

Else
'' PREGUNTA POR CAMPAÑA B.I.
    If cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana = "Si" And (Ficha_Cliente_Micro.txt_score_dicom > 100 Or Ficha_Cliente_Micro.txt_score_dicom = 0) _
        And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
            
            txt_r_dicom_tipo_cliente = "A"
            
        ElseIf cbx_Cliente_Nuevo = "No" And Ficha_Cliente_Micro.txt_campana = "Si" And Ficha_Cliente_Micro.txt_score_dicom > 409 _
        And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
        
            txt_r_dicom_tipo_cliente = "A"
            
        ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 100 Then
           txt_r_dicom_tipo_cliente = "R"
            
        ElseIf cbx_Cliente_Nuevo = "No" And Ficha_Cliente_Micro.txt_campana = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 409 Then
           txt_r_dicom_tipo_cliente = "R"
            
            
''' PREGUNTA POR CAMPAÑA DE 48 MESES (CAMPAÑITA)
        ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_48M = "Si" And (Ficha_Cliente_Micro.txt_score_dicom > 100 Or Ficha_Cliente_Micro.txt_score_dicom = 0) _
        And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
            
                txt_r_dicom_tipo_cliente = "A"
                
        ElseIf cbx_Cliente_Nuevo = "No" And Ficha_Cliente_Micro.txt_campana_48M = "Si" And Ficha_Cliente_Micro.txt_score_dicom > 409 _
        And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
            
                txt_r_dicom_tipo_cliente = "A"
                
        ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_48M = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 100 Then
           txt_r_dicom_tipo_cliente = "R"
            
        ElseIf cbx_Cliente_Nuevo = "No" And Ficha_Cliente_Micro.txt_campana_48M = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 409 Then
           txt_r_dicom_tipo_cliente = "R"
                
''' PREGUNTA POR CAMPAÑA EVALUADOS
            ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_evaluados = "Si" And Ficha_Cliente_Micro.txt_score_dicom > 790 _
            And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
                    
                    txt_r_dicom_tipo_cliente = "A"
                    
            ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_evaluados = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 790 Then

                txt_r_dicom_tipo_cliente = "R"
               
            ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_evaluados = "No" And (Ficha_Cliente_Micro.txt_score_dicom > 0 And Ficha_Cliente_Micro.txt_score_dicom <= 100) Then

                txt_r_dicom_tipo_cliente = "R"
                   
''' PREGUNTA POR CAMPAÑA EVALUADOS
            ElseIf cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_evaluados = "No" And (Ficha_Cliente_Micro.txt_score_dicom = 0 Or Ficha_Cliente_Micro.txt_score_dicom > 100) _
            And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") Then
                    MarcarZG = 1
                    MsgBox "Solicitar SCORE MINIMO a Adminsion Retail Banking (S.I.C.)", vbCritical
                    txt_r_dicom_tipo_cliente = "ZG"
                   
            Else
            txt_r_dicom_tipo_cliente = "R"
            
    End If
End If




'XXXXXXXXXXXXXXXXXXXXX
'INICIO COMENTADO X ERROR
'XXXXXXXXXXXXXXXXXXXXXX

'If (txt_tipo_cliente = "Nuevo Sin Historia Sbif") _
    And (Ficha_Cliente_Micro.txt_campana = "No" And Ficha_Cliente_Micro.txt_campana_48M = "No" Or _
    (Ficha_Cliente_Micro.txt_campana_evaluados = "No")) And Ficha_Cliente_Micro.txt_score_dicom < 790 Then

'comentado Vmanriquez '    (Ficha_Cliente_Micro.txt_score_dicom <> 0 And Ficha_Cliente_Micro.txt_score_dicom < 100) Then
    
 '       txt_r_dicom_tipo_cliente = "R"
        
  '      Else
                'txt_r_dicom_tipo_cliente = "ZG"
                'MsgBox "Solicitar SCORE MINIMO a Adminsion Retail Banking (S.I.C.)", vbCritical
'End If
'-----------------------
        
'If (txt_tipo_cliente = "Nuevo Sin Historia Sbif") _
    And (Ficha_Cliente_Micro.txt_campana = "No" And Ficha_Cliente_Micro.txt_campana_48M = "No" Or _
    (Ficha_Cliente_Micro.txt_campana_evaluados = "Si")) And Ficha_Cliente_Micro.txt_score_dicom < 790 Then

 '       txt_r_dicom_tipo_cliente = "R"
        
  '      Else
                'txt_r_dicom_tipo_cliente = "A"
   '
'End If
        
        'If (txt_tipo_cliente = "Nuevo Sin Historia Sbif" Or txt_tipo_cliente = "Nuevo Con Historia Sbif") And _
        '        Ficha_Cliente_Micro.txt_campana_evaluados = "Si" And Ficha_Cliente_Micro.txt_score_dicom < 790 Then
                
        '        txt_r_dicom_tipo_cliente = "R"
        
        'ElseIf Ficha_Cliente_Micro.txt_campana = "No" And Ficha_Cliente_Micro.txt_campana_48M = "No" And _
                Ficha_Cliente_Micro.txt_campana_evaluados = "No" Then
        
          '          txt_r_dicom_tipo_cliente = "ZG"
         '           MsgBox "Solicitar SCORE MINIMO a Adminsion Retail Banking (S.I.C.)", vbCritical
        
'If (txt_tipo_cliente = "Nuevo Sin Historia Sbif" Or txt_tipo_cliente = "Nuevo Con Historia Sbif") And _
    (Ficha_Cliente_Micro.txt_campana = "No" Or Ficha_Cliente_Micro.txt_campana_48M = "No" Or txt_campana_evaluados = "Si") And _
    Ficha_Cliente_Micro.txt_score_dicom < 790 Then
    
 '   txt_r_dicom_tipo_cliente = "R"
  '  MsgBox "El cliente no esta en campaña y no cumple con score sinacofi", vbCritical

'Else
 
 'If txt_rut_cliente < 45000000 Then
    
  '  If txt_tipo_cliente = "Nuevo Con Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom < 535 Then
   '     txt_r_dicom_tipo_cliente = "R"
        
      'ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom <= 100 _
      '      And Ficha_Cliente_Micro.txt_score_dicom <> 0 Then
      'txt_r_dicom_tipo_cliente = "R"
        
      'ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And Ficha_Cliente_Micro.txt_score_dicom = 0 Then
      
      'Cambio solicitado por Viviana Manriquez mail 16-10-2013
      'txt_r_dicom_tipo_cliente = "ZG"
      'txt_r_dicom_tipo_cliente = "A"
      
    '  ElseIf txt_tipo_cliente = "Antiguo No Prime" And Ficha_Cliente_Micro.txt_score_dicom < 409 Then
     ' txt_r_dicom_tipo_cliente = "R"
          
      'ElseIf txt_tipo_cliente = "Antiguo Prime" And Ficha_Cliente_Micro.txt_score_dicom < 347 Then
      'txt_r_dicom_tipo_cliente = "R"
    
      'Else
      'txt_r_dicom_tipo_cliente = "A"

'    End If
    
 '   Else
  '  txt_r_dicom_tipo_cliente = "ZG"

  'End If
'End If '------
'End If
    
'MsgBox "Solicitar SCORE MINIMO a Adminsion Retail Banking (S.I.C.)", vbCritical


'XXXXXXXXXXXXXXXXXXXXX
'FIN COMENTADO X ERROR
'XXXXXXXXXXXXXXXXXXXXXX


''********************************
''''CALCULA POLICA BANCARIZADO
''********************************

If Ficha_Cliente_Micro.txt_bancarizado_politica = "Si" Then

If txt_tipo_cliente = "Nuevo Sin Historia Sbif" And Ficha_Cliente_Micro.txt_bancarizado_politica = "No" Then
   Estado_Resolucion_Final.txt_r_f_bancarizado_politica = "R"

    ElseIf txt_tipo_cliente = "Antiguo No Prime" And Ficha_Cliente_Micro.txt_bancarizado_politica = "No" Then   ''' Solicitud de MARIO SAN CRISTOBAL 20-06-2012
    Estado_Resolucion_Final.txt_r_f_bancarizado_politica = "ZG"
   
    Else
        Estado_Resolucion_Final.txt_r_f_bancarizado_politica = "A"
                                
End If

Else
    Estado_Resolucion_Final.txt_r_f_bancarizado_politica = Ficha_Cliente_Micro.txt_ESTADO_politica_bancarizado_new

End If


'##############################################
' Agregando

Dim tipoVivienda As String
Dim AntigRubro As String
Dim R_Tipo_Cliente As String
Dim Msg_RSGO As String
Dim UF As Double
Dim LVGE As Double
Dim TDSR As Double
Dim RI As Double
Dim RI_char As String
Dim Respuesta As Double
Dim Sql2 As String
Dim n_solicitud As Double




    If Ficha_Cliente_Micro.cbx_bien_Raiz = "Arrendado" Then
        tipoVivienda = "3"
    End If

    If Ficha_Cliente_Micro.cbx_bien_Raiz = "Propio" Then
        tipoVivienda = "1"
    End If
    
    If Ficha_Cliente_Micro.cbx_bien_Raiz = "Vive Con Familiares" Then
        tipoVivienda = "4"
    End If
    
If Ficha_Cliente_Micro.cbx_antiguedad_rubro.Text = "Antig. Acreditada Cliente Bco" Then
    AntigRubro = "1"
ElseIf Ficha_Cliente_Micro.cbx_antiguedad_rubro.Text = "Iniciación De Actividades" Then
    AntigRubro = "2"
ElseIf Ficha_Cliente_Micro.cbx_antiguedad_rubro.Text = "DAI" Then
    AntigRubro = "3"
ElseIf Ficha_Cliente_Micro.cbx_antiguedad_rubro.Text = "Carpeta Tributaria" Then
    AntigRubro = "4"
Else
    AntigRubro = "5"
End If

If txt_tipo_cliente = "Antiguo Prime" Then
    R_Tipo_Cliente = "1"
    txt_tipo_cliente_muestra = "Antiguo"
ElseIf txt_tipo_cliente = "Antiguo No Prime" Then
    R_Tipo_Cliente = "2"
    txt_tipo_cliente_muestra = "Antiguo"
ElseIf txt_tipo_cliente = "Nuevo Con Historia Sbif" Then
    R_Tipo_Cliente = "3"
    txt_tipo_cliente_muestra = "Nuevo Con Historia Sbif"
ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" Then
    R_Tipo_Cliente = "4"
    txt_tipo_cliente_muestra = "Nuevo Sin Historia Sbif"
End If
      
     
conectarBD
    '''''CARGA_UF
    ssql = "select cast(valor_uf as int) valor_uf from tbl_valor_uf" _
        & " WHERE fecha_dia = '" & Format(Date, "yyyymmdd") & "'"
        
    Set rst = cnn.Execute(ssql, , adCmdText)

    UF = rst("valor_uf")

   rst.Close
   
conectarBD
     ssql = "SELECT rut_cliente, max(n_solicitud) as n_solicitud FROM tbl_micro_ficha_cliente where rut_cliente = '" & txt_rut_cliente & "' group by rut_cliente"
                       
        Set rst = cnn.Execute(ssql, , adCmdText)
            
        If rst.EOF Then
           MsgBox ("Cliente No ingresado en la ficha")
          Else
              n_solicitud = rst!n_solicitud
              rst.MoveNext
              
        End If
rst.Close
   


ssql = " EXEC sp_Com_Score_interno  " & Ficha_Cliente_Micro.txt_rut_cliente & "," _
        & "8, '" & Format(Date, "yyyymmdd") & "'," _
        & "'" & Format(Ficha_Cliente_Micro.txt_fechanacimiento, "yyyymmdd") & "'," _
        & "NULL," _
        & tipoVivienda & "," _
        & "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," _
        & Ficha_Cliente_Micro.txt_score_dicom & "," _
        & Trim(UF) & "," _
        & "'" & Format(DateAdd("m", Ficha_Cliente_Micro.txt_antiguedad_meses * -1, Date), "yyyymmdd") & "'," _
        & "'" & Left(Ficha_Cliente_Micro.cbx_vehiculos_propios.Text, 1) & "'," _
        & AntigRubro & "," _
        & "'" & Trim(R_Tipo_Cliente) & "'"
        

conectarBDRIESGO2


Set rst = cnn.Execute(ssql, , adCmdText)
 
LVGE = GVerificar_NULL(rst(1), "-") '4.0
TDSR = GVerificar_NULL(rst(2), "-") '0.6
Respuesta = GVerificar_NULL(rst(3), "-") '0
RI = GVerificar_NULL(rst(4), "-") ' 5



 If RI = "1" Then
    RI_char = "A"
    txt_r_dicom_tipo_cliente = "A"
ElseIf RI = "2" Then
    RI_char = "B"
    txt_r_dicom_tipo_cliente = "A"
ElseIf RI = "3" Then
    RI_char = "C"
    txt_r_dicom_tipo_cliente = "A"
ElseIf RI = "4" Then
    RI_char = "D"
    txt_r_dicom_tipo_cliente = "A"
ElseIf RI = "5" Then
    RI_char = "E"
    txt_r_dicom_tipo_cliente = "R"
Else
    RI_char = ""
    txt_r_dicom_tipo_cliente = ""
    cmd_calcular_perfil.Enabled = False
    MsgBox ("Este Rut a excedido el número de evaluaciones")
End If





 If Respuesta = "1" Then
    Msg_RSGO = "Aprobado"
ElseIf Respuesta = "2" Then
    Msg_RSGO = "Excede el máximo de consultas del mes"
ElseIf Respuesta = "0" Then
    Msg_RSGO = "Rechazo"
ElseIf Respuesta = "4" Then
    Msg_RSGO = "Cliente NO Bancarizado, debe cumplir política de No Bancarizados"
ElseIf Respuesta = "5" Then
    Msg_RSGO = "Cliente NO Bancarizado, NOcumple política de Crédito"
Else
    Msg_RSGO = "Respuesta No identificada"
End If
 
 
 
If txt_r_dicom_tipo_cliente = "A" And cbx_Cliente_Nuevo = "Si" And Ficha_Cliente_Micro.txt_campana_evaluados = "No" And (Ficha_Cliente_Micro.txt_score_dicom = 0 Or Ficha_Cliente_Micro.txt_score_dicom > 100) _
            And Ficha_Cliente_Micro.txt_r_edad = "A" And Ficha_Cliente_Micro.txt_r_formalidad_negocio = "A" And Ficha_Cliente_Micro.txt_r_meses_antiguedad = "A" And Ficha_Cliente_Micro.txt_r_cbx_antiguedad_rubro = "A" And Ficha_Cliente_Micro.txt_r_cbx_actividad_economica_informal_oficio = "A" And (Ficha_Cliente_Micro.txt_r_cbx_bien_Raiz = "A" Or Ficha_Cliente_Micro.txt_r_cbx_vehiculos_propios = "A") And MarcarZG = 1 Then
   txt_r_dicom_tipo_cliente = "ZG"
   Msg_RSGO = "Zona Gris"
End If
 
 
 
 Evaluacion_Perfil.txt_leverage.Text = LVGE
 txt_tdsr = TDSR
 lbl_MensajeRiesgo.Caption = Msg_RSGO
 txt_RI = RI_char
 
 
 'txt_leverage.Visible = True
 txt_RI.Visible = True
 'txt_tdsr.Visible = True
 
 lbl_RI.Visible = True
 'lbl_TDSR.Visible = True
 lbl_MensajeRiesgo.Visible = True
 'lbl_Leverage.Visible = True
 
 txt_r_dicom_tipo_cliente.Visible = True
 
 

rst.Close

Sql2 = "Insert into TBL_MICRO_EVALUACION_RIESGO" _
& " (n_solicitud,Rut_Num,   IN_BancaIngreso,    IN_CuadranteCliente," _
& " IN_FechaNacimiento, IN_ScoreSinacofi,   IN_ValorUF," _
& " IN_FInicioNegocio,  IN_VehiculosPropios,    IN_CertifAntigRubro," _
& " IN_RTipoCliente,    SQL,    Out_LVRGE,  Out_TDSR,   Out_Respuesta," _
& " Out_IndRiesgo,  FechaIngreso,   Responsable,    Cod_Ambiente)" _
& " values" _
& " (" & n_solicitud & "," _
& " " & Ficha_Cliente_Micro.txt_rut_cliente & "," _
& " 'MB' , 8 ," _
& " '" & Format(Ficha_Cliente_Micro.txt_fechanacimiento, "yyyymmdd") & "'," _
& " " & Ficha_Cliente_Micro.txt_score_dicom & "," _
& " " & Trim(UF) & "," _
& " '" & Format(DateAdd("m", Ficha_Cliente_Micro.txt_antiguedad_meses * -1, Date), "yyyymmdd") & "'," _
& " '" & Left(Ficha_Cliente_Micro.cbx_vehiculos_propios.Text, 1) & "'," _
& " " & AntigRubro & "," _
& " '" & Trim(R_Tipo_Cliente) & "', " _
& " '" & Replace(ssql, "'", "#CHR39#") & "', " _
& " " & Str(Trim(LVGE)) & ", " _
& " " & Str(Trim(TDSR)) & ", " _
& " " & Trim(Respuesta) & ", " _
& " " & Trim(RI) & ", " _
& " getdate()," _
& "'" & Environ("USERNAME") & "', " _
& "'PROD')"

conectarBD

cnn.BeginTrans
    On Error Resume Next
    cnn.Execute (Sql2)
    If Err.Number <> 0 Then 'error
            'MsgBox ("Se ha producido un error al Registrar el paso por Riesgo, Podrás continuear de todas formas")
            cnn.RollbackTrans
        Else
            cnn.CommitTrans
    End If
    


End Sub

Private Sub cmd_Volver_Ingreso_dPerfil_Click()
Ficha_Cliente_Micro.Show
End Sub



Private Sub cmd_imprimir1_meto_ac_Click()
Evaluacion_Perfil.PrintForm
End Sub

Private Sub cmd_volver_ficha_Click()

Evaluacion_Perfil.txt_tipo_cliente = Empty
Evaluacion_Perfil.R_Final_Perfil = Empty
Evaluacion_Perfil.cmd_metodologia_iva.Enabled = False
Evaluacion_Perfil.cmd_metodologia_activo_circulante = False
Evaluacion_Perfil.cmd_metodologia_maxima_produccion = False
Evaluacion_Perfil.cmd_calcular_perfil = False
Evaluacion_Perfil.txt_r_dicom_tipo_cliente = False


MsgBox "Recuerda Que Al Volver y Cambiar Datos Debes Recalcular Los Campos"
Ficha_Cliente_Micro.cmd_Menu_Evaluacion.Enabled = False

Evaluacion_Perfil.Hide
Ficha_Cliente_Micro.Show

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label182_Click()

End Sub





Private Sub txt_Antiguedad_AfterUpdate()
'txt_registro_ventas = Empty
'cbx_actividad_economica_informal_oficio = Empty
'cbx_actividad_economica_semiformal = Empty
'cbx_actividad_economica_formal = Empty
'cbx_actividad_economica_formal_servicio = Empty
'cmd_calcular_perfil.Enabled = False
'R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
'cmd_metodologia_activo_circulante.Enabled = False
'cmd_metodologia_maxima_produccion.Enabled = False

'cmd_calcular_tipo_cliente.Enabled = False
'txt_tipo_cliente = Empty
End Sub

Private Sub txt_Antiguedad_Change()

'txt_registro_ventas = Empty
'cbx_actividad_economica_informal_oficio = Empty
'cbx_actividad_economica_semiformal = Empty
'cbx_actividad_economica_formal = Empty
'cbx_actividad_economica_formal_servicio = Empty
'cmd_calcular_perfil.Enabled = False
'R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
'cmd_metodologia_activo_circulante.Enabled = False
'cmd_metodologia_maxima_produccion.Enabled = False

'cmd_calcular_tipo_cliente.Enabled = False
'txt_tipo_cliente = Empty

End Sub




Private Sub txt_registro_ventas_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

R_Final_Perfil = Empty

txt_registro_ventas = txt_registro_ventas
txt_registro_ventas_var = txt_registro_ventas / 100

txt_registro_ventas_dif_var = 100 - txt_registro_ventas
txt_registro_ventas_dif1_var = txt_registro_ventas_dif_var / 100


If txt_registro_ventas >= 0 And txt_registro_ventas <= 39 Then
    txt_r_registro_ventas = "Regular"

  ElseIf txt_registro_ventas >= 40 And txt_registro_ventas <= 79 Then
    txt_r_registro_ventas = "Bueno"

  ElseIf txt_registro_ventas >= 80 Then
    txt_r_registro_ventas = "Excelente"

End If

End Sub

Private Sub cbx_historia_pago_Change()

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False


R_Final_Perfil = Empty

''''MORA PROMEDIO
'If cbx_historia_pago = "Hasta 5 dias" Or cbx_historia_pago = "Sin Mora" Then
'   txt_r_historia_Pago = "Excelente"

'ElseIf cbx_historia_pago = "Hasta 7 dias" Then
'   txt_r_historia_Pago = "Bueno"

'ElseIf cbx_historia_pago = "Hasta 15 dias" Then
'    txt_r_historia_Pago = "Regular"

'End If

'''****************************************************
'''Cambio realizador por mail de Mario San Cristobal
'''****************************************************
''''MORA PROMEDIO
If (cbx_historia_pago = "Hasta 5 dias" Or cbx_historia_pago = "Sin Mora") And (cbx_mora_maxima = "Sin Mora" Or cbx_mora_maxima = "Hasta 5 dias" Or cbx_mora_maxima = "Hasta 7 dias") Then
   txt_r_historia_Pago = "Excelente"

ElseIf (cbx_historia_pago = "Hasta 5 dias" Or cbx_historia_pago = "Sin Mora" Or cbx_historia_pago = "Hasta 7 dias") And (cbx_mora_maxima = "Sin Mora" Or cbx_mora_maxima = "Hasta 5 dias" Or cbx_mora_maxima = "Hasta 7 dias" Or cbx_mora_maxima = "Hasta 15 dias") Then
   txt_r_historia_Pago = "Bueno"

''ElseIf cbx_historia_pago = "Hasta 15 dias" Then
  Else
    txt_r_historia_Pago = "Regular"

End If

'cmd_calcular_tipo_cliente.Enabled = False
txt_tipo_cliente = Empty

End Sub


Private Sub cbx_mora_maxima_Change()

txt_tipo_cliente = Empty

txt_registro_ventas = Empty
cbx_actividad_economica_informal_oficio = Empty
cbx_actividad_economica_semiformal = Empty
cbx_actividad_economica_formal = Empty
cbx_actividad_economica_formal_servicio = Empty
cmd_calcular_perfil.Enabled = False
R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False

R_Final_Perfil = Empty

If cbx_mora_maxima = "Hasta 7 dias" Or cbx_mora_maxima = "Sin Mora" Then
    txt_r_mora = "Excelente"

ElseIf cbx_mora_maxima = "Hasta 15 dias" Then
    txt_r_mora = "Bueno"
    
ElseIf cbx_mora_maxima = "Hasta 30 dias" Then
     txt_r_mora = "Regular"

End If

'If cbx_mora_maxima <> "" Then
'cmd_calcular_tipo_cliente.Enabled = True
'End If

End Sub


Private Sub cmd_activo_circulante_Click()
Unload Evaluacion_Perfil
Metodologia_Activo_Circulante.Show
End Sub

Private Sub cmd_calcular_perfil_Click()

R_Final_Perfil = Empty
'cmd_Metodologia_IVA1.Enabled = False
cmd_metodologia_activo_circulante.Enabled = False
cmd_metodologia_maxima_produccion.Enabled = False

'''' PREGUNTA SI EL CLIENTE ES CRITICO

If txt_tipo_cliente_duro = 1 Then
    R_Final_Perfil = "Regular"

    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''calcualo de metodologia con TIPO_CLIENTE """"REGULAR""""" TABLA ::::: TBL_MICRO_RESUMEN_CRITICAL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        If txt_registro_ventas <> "" Then

'------------------------------

'CALCULA ANTIGUEDAD NEGOCIO
'------------------------------

        If txt_antiguedad_negocio < 24 Then
        txt_r_antiguedad_Negocio = "Regular"

        ElseIf txt_antiguedad_negocio >= 24 And txt_antiguedad_negocio <= 60 Then
        txt_r_antiguedad_Negocio = "Bueno"

        ElseIf txt_antiguedad_negocio > 60 Then
        txt_r_antiguedad_Negocio = "Excelente"
     
        End If


        'CALCULA SCORE DICOM
        '--------------------

        'R_Final_Perfil = Empty


'' cambiado por solicitud mail de Viviana Manrqiuez 30-08-2013

        If (txt_predictor_Score > 100 And txt_predictor_Score <= 534) Or txt_predictor_Score = 0 Then
        txt_r_predictor_score = "Regular"

        ElseIf txt_predictor_Score >= 535 And txt_predictor_Score <= 630 Then
        txt_r_predictor_score = "Bueno"

        ElseIf txt_predictor_Score > 631 Then
        txt_r_predictor_score = "Excelente"

        End If

'''''MODIFICADO CMA 06-0-2012

        ''''ANTIGUO PRIME
        'If txt_tipo_cliente = "Antiguo Prime" And Val(txt_predictor_Score) * 1 > 631 Then
        'R_Final_Perfil = "Regular"

        If txt_tipo_cliente = "Antiguo Prime" And (Val(txt_predictor_Score) * 1 >= 347 And Val(txt_predictor_Score) * 1 <= 534) Then
        R_Final_Perfil = "Regular"

        'ElseIf txt_tipo_cliente = "Antiguo Prime" And (Val(txt_predictor_Score) * 1 >= 743 And Val(txt_predictor_Score) * 1 <= 845) Then
        'R_Final_Perfil = "Regular"


        '''''ANTIGUO NO PRIME
        ElseIf txt_tipo_cliente = "Antiguo No Prime" And (txt_r_registro_ventas = "Regular" Or txt_r_predictor_score = "Regular" Or txt_r_pago_mora = "Regular") Then
        R_Final_Perfil = "Regular"
    
        ElseIf txt_tipo_cliente = "Antiguo No Prime" And (txt_r_registro_ventas = "Bueno" Or txt_r_predictor_score = "Bueno" Or txt_r_pago_mora = "Bueno") Then
        R_Final_Perfil = "Regular"

        ''''NUEVO BANCARIZADO
        ElseIf txt_tipo_cliente = "Nuevo Con Historia Sbif" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        'ElseIf txt_tipo_cliente = "Nuevo Bancarizado" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        R_Final_Perfil = "Regular"

        ElseIf txt_tipo_cliente = "Nuevo Con Historia Sbif" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        'ElseIf txt_tipo_cliente = "Nuevo Bancarizado" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        R_Final_Perfil = "Regular"

        ''''NUEVO NO BANCARIZADO

        ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        'ElseIf txt_tipo_cliente = "Nuevo No Bancarizado" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        R_Final_Perfil = "Regular"

        ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        'ElseIf txt_tipo_cliente = "Nuevo No Bancarizado" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        R_Final_Perfil = "Regular"

        Else
        R_Final_Perfil = "Regular"

        End If

 
  
  'CALCULA METODOLOGIA A UTILIZAR
  '-----------------------------------
    
     If txt_tipo_cliente_actividad = "FORMALES" And txt_registro_ventas >= 0 And txt_registro_ventas <= 39 Then
         
     'cmd_Metodologia_IVA1.Enabled = False
     cmd_metodologia_activo_circulante.Enabled = True
     cmd_metodologia_maxima_produccion.Enabled = False

        ElseIf txt_tipo_cliente_actividad = "FORMALES" And _
       (txt_registro_ventas >= 40 And txt_registro_ventas <= 79 _
        Or txt_r_registro_ventas >= 80) Then

        cmd_metodologia_iva.Enabled = True
        cmd_metodologia_activo_circulante.Enabled = False
        cmd_metodologia_maxima_produccion.Enabled = False


        ElseIf txt_tipo_cliente_actividad = "SEMIFORMALES" Then
    
        cmd_metodologia_iva.Enabled = False
        cmd_metodologia_activo_circulante.Enabled = True
        cmd_metodologia_maxima_produccion.Enabled = False
  
     
        ElseIf txt_tipo_cliente_actividad = "INFORMALES(Oficio)" Or txt_tipo_cliente_actividad = "FORMAL SERVICIO O PRODUCCION" Then
     
        cmd_metodologia_iva.Enabled = False
        cmd_metodologia_activo_circulante.Enabled = False
        cmd_metodologia_maxima_produccion.Enabled = True
    
    End If


'GUARDA VARIBALES PUBLICAS DE PERFIL CLIENTE
'----------------------------------------------
 txt_tipo_cliente_evaluacion = txt_tipo_cliente
 R_Final_Perfil_evaluacion = R_Final_Perfil
   
    
'Paso de Variables para el INSERT EN SQL
'------------------------------ ------------------
 
Cliente_Nuevo_evaluacion = cbx_Cliente_Nuevo
Bancarizado_evaluacion = cbx_bancarizado
Antiguedad_banco_evaluacion = txt_Antiguedad
mora_promedio_dias_BD_evaluacion = cbx_historia_pago
mora_maxima_dias_BD_evaluacion = cbx_mora_maxima
txt_tipo_cliente_evaluacion = txt_tipo_cliente
registros_ventas_evaluacion = txt_registro_ventas
R_Final_Perfil_evaluacion = R_Final_Perfil

Else
    MsgBox ("Debe Ingresar El Valor Porcentual Del Registro De Ventas")
End If

'''''''''''''''''''''''fin de calculo metodologia cliente critico


Else

        If txt_registro_ventas <> "" Then

'------------------------------

'CALCULA ANTIGUEDAD NEGOCIO
'------------------------------

        If txt_antiguedad_negocio < 24 Then
        txt_r_antiguedad_Negocio = "Regular"

        ElseIf txt_antiguedad_negocio >= 24 And txt_antiguedad_negocio <= 60 Then
        txt_r_antiguedad_Negocio = "Bueno"

        ElseIf txt_antiguedad_negocio > 60 Then
        txt_r_antiguedad_Negocio = "Excelente"
     
        End If


        'CALCULA SCORE DICOM
        '--------------------

        'R_Final_Perfil = Empty


        If (txt_predictor_Score > 100 And txt_predictor_Score <= 534) Or txt_predictor_Score = 0 Then
        txt_r_predictor_score = "Regular"

        ElseIf txt_predictor_Score >= 535 And txt_predictor_Score <= 630 Then
        txt_r_predictor_score = "Bueno"

        ElseIf txt_predictor_Score > 631 Then
        txt_r_predictor_score = "Excelente"

        End If



        ''''ANTIGUO PRIME
        If txt_tipo_cliente = "Antiguo Prime" And Val(txt_predictor_Score) * 1 > 631 Then
        R_Final_Perfil = "Excelente"

        ElseIf txt_tipo_cliente = "Antiguo Prime" And (Val(txt_predictor_Score) * 1 >= 347 And Val(txt_predictor_Score) * 1 <= 534) Then
        R_Final_Perfil = "Regular"

        ElseIf txt_tipo_cliente = "Antiguo Prime" And (Val(txt_predictor_Score) * 1 >= 535 And Val(txt_predictor_Score) * 1 <= 630) Then
        R_Final_Perfil = "Bueno"


        '''''ANTIGUO NO PRIME
        ElseIf txt_tipo_cliente = "Antiguo No Prime" And (txt_r_registro_ventas = "Regular" Or txt_r_predictor_score = "Regular" Or txt_r_pago_mora = "Regular") Then
        R_Final_Perfil = "Regular"
    
        ElseIf txt_tipo_cliente = "Antiguo No Prime" And (txt_r_registro_ventas = "Bueno" Or txt_r_predictor_score = "Bueno" Or txt_r_pago_mora = "Bueno") Then
        R_Final_Perfil = "Bueno"

        ''''NUEVO BANCARIZADO
        ElseIf txt_tipo_cliente = "Nuevo Con Historia Sbif" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        'ElseIf txt_tipo_cliente = "Nuevo Bancarizado" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        R_Final_Perfil = "Regular"

        ElseIf txt_tipo_cliente = "Nuevo Con Historia Sbif" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        'ElseIf txt_tipo_cliente = "Nuevo Bancarizado" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        R_Final_Perfil = "Bueno"

        ''''NUEVO NO BANCARIZADO

        ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        'ElseIf txt_tipo_cliente = "Nuevo No Bancarizado" And (txt_r_registro_ventas = "Regular" Or txt_r_antiguedad_Negocio = "Regular" Or txt_r_predictor_score = "Regular") Then
        R_Final_Perfil = "Regular"

        ElseIf txt_tipo_cliente = "Nuevo Sin Historia Sbif" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        'ElseIf txt_tipo_cliente = "Nuevo No Bancarizado" And (txt_r_registro_ventas = "Bueno" Or txt_r_antiguedad_Negocio = "Bueno" Or txt_r_predictor_score = "Bueno") Then
        R_Final_Perfil = "Bueno"

        Else
        R_Final_Perfil = "Excelente"

        End If

 
  
  'CALCULA METODOLOGIA A UTILIZAR
  '-----------------------------------
    
     If txt_tipo_cliente_actividad = "FORMALES" And txt_registro_ventas >= 0 And txt_registro_ventas <= 39 Then
         
     cmd_metodologia_iva.Enabled = False
     cmd_metodologia_activo_circulante.Enabled = True
     cmd_metodologia_maxima_produccion.Enabled = False

        ElseIf txt_tipo_cliente_actividad = "FORMALES" And _
       (txt_registro_ventas >= 40 And txt_registro_ventas <= 79 _
        Or txt_r_registro_ventas >= 80) Then

        cmd_metodologia_iva.Enabled = True
        cmd_metodologia_activo_circulante.Enabled = False
        cmd_metodologia_maxima_produccion.Enabled = False


        ElseIf txt_tipo_cliente_actividad = "SEMIFORMALES" Then
    
        cmd_metodologia_iva.Enabled = False
        cmd_metodologia_activo_circulante.Enabled = True
        cmd_metodologia_maxima_produccion.Enabled = False
  
     
        ElseIf txt_tipo_cliente_actividad = "INFORMALES(Oficio)" Or txt_tipo_cliente_actividad = "FORMAL SERVICIO O PRODUCCION" Then
     
        cmd_metodologia_iva.Enabled = False
        cmd_metodologia_activo_circulante.Enabled = False
        cmd_metodologia_maxima_produccion.Enabled = True
    
    End If


'GUARDA VARIBALES PUBLICAS DE PERFIL CLIENTE
'----------------------------------------------
 txt_tipo_cliente_evaluacion = txt_tipo_cliente
 R_Final_Perfil_evaluacion = R_Final_Perfil
   
    
'Paso de Variables para el INSERT EN SQL
'------------------------------ ------------------
 
Cliente_Nuevo_evaluacion = cbx_Cliente_Nuevo
Bancarizado_evaluacion = cbx_bancarizado
Antiguedad_banco_evaluacion = txt_Antiguedad
mora_promedio_dias_BD_evaluacion = cbx_historia_pago
mora_maxima_dias_BD_evaluacion = cbx_mora_maxima
txt_tipo_cliente_evaluacion = txt_tipo_cliente
registros_ventas_evaluacion = txt_registro_ventas
R_Final_Perfil_evaluacion = R_Final_Perfil

Else
    MsgBox ("Debe Ingresar El Valor Porcentual Del Registro De Ventas")
End If



End If

End Sub


Private Sub cmd_metodologia_activo_circulante_Click()


If cbx_actividad_economica_formal <> "" Or cbx_actividad_economica_formal_servicio <> "" Or cbx_actividad_economica_semiformal <> "" Or cbx_actividad_economica_informal_oficio <> "" Then


conectarBD

''''' ---- DEUDAS LIBRODEUDORES VIGENTES E INDIRECTAS

        ssql = "SELECT m_deudacreditoconsumo, m_deudacomercial,m_creditohipotecario, m_cupolineacredito, m_deudaindirectavigente, m_deudadirectavigente" _
        & " FROM tbl_micro_fact_ods_librodeudores " _
        & " where rut = '" & txt_rut_cliente & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        If Not rst.EOF Then
    
        Metodologia_Activo_Circulante.txt_deuda_consumo = rst!m_deudacreditoconsumo * 1000
        Metodologia_Activo_Circulante.txt_deuda_comercial = rst!m_deudacomercial * 1000
        Metodologia_Activo_Circulante.txt_credito_hipotecario = rst!m_creditohipotecario * 1000
        Metodologia_Activo_Circulante.txt_cupo_linea_credito = rst!m_cupolineacredito * 1000
        Metodologia_Activo_Circulante.txt_deuda_indirecta = rst!m_deudaindirectavigente * 1000
        Metodologia_Activo_Circulante.txt_deudas_directas_vig = rst!m_deudadirectavigente * 1000
        
        End If
'        Metodologia_Activo_Circulante.txt_total_deudas_sbif = Int((Metodologia_Activo_Circulante.txt_deudas_directas_vig) + Val(Metodologia_Activo_Circulante.txt_deuda_comercial) + Val(Metodologia_Activo_Circulante.txt_cupo_linea_credito) + Val(Metodologia_Activo_Circulante.txt_deuda_indirecta_vigente) + Val(Metodologia_Activo_Circulante.txt_deuda_consumo_indirecta) + Val(Metodologia_Activo_Circulante.txt_credito_hipotecario))

      
'------------- DEUDAS D10----  CONSUMO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 2 then sum(m_deuda) else 0 end deuda_consumo" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 2" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Activo_Circulante.txt_deuda_d10_consumo = 0

        Else
        Metodologia_Activo_Circulante.txt_deuda_d10_consumo = rst!deuda_consumo
        
        End If
        
        
'------------- DEUDAS D10----  COMERCIAL

        ssql = "SELECT rut," _
                & " case when t_operacion = 1 then sum(m_deuda) else 0 end deuda_comercial" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 1" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Activo_Circulante.txt_deuda_d10_comercial = 0

        Else
        Metodologia_Activo_Circulante.txt_deuda_d10_comercial = rst!deuda_comercial
        
        End If



'------------- DEUDAS D10----  HIPOTECARIO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 3 then sum(m_deuda) else 0 end deuda_hipotecario" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 3" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Activo_Circulante.txt_deuda_d10_hipotecario = 0

        Else
        Metodologia_Activo_Circulante.txt_deuda_d10_hipotecario = rst!deuda_hipotecario
        
        End If
        
        
        
'------------- DEUDAS D10----  LINEA DE CREDITO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 7 then sum(m_deuda) else 0 end deuda_cupo_linea" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 7" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Activo_Circulante.txt_deuda_d10_linea = 0

        Else
        Metodologia_Activo_Circulante.txt_deuda_d10_linea = rst!deuda_cupo_linea
        
        End If
        
'''''''''''CALCULA TOTAL DE DEUDAS D10
        
        Metodologia_Activo_Circulante.txt_total_deuda_d10 = Int(Val(Metodologia_Activo_Circulante.txt_deuda_d10_consumo) + Val(Metodologia_Activo_Circulante.txt_deuda_d10_comercial) + Val(Metodologia_Activo_Circulante.txt_deuda_d10_hipotecario) + Val(Metodologia_Activo_Circulante.txt_deuda_d10_linea))
        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '''sacar fecha-hora
   
        Dim fec1
        Dim hora1
    
        fec1 = Format(Date, "yyyy-mm-dd")
        txt_fecha_actual = fec1


        '''''CARGA_UF
        ssql = "select cast(valor_uf as int) valor_uf from tbl_valor_uf" _
            & " WHERE fecha_dia = '" & txt_fecha_actual & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        Metodologia_Activo_Circulante.txt_valor_uf = rst!valor_uf
        'Metodologia_IVA1.txt_valor_uf = rst!valor_uf
        'Metodologia_Maxima_Prod.txt_valor_uf = rst!valor_uf


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        metodologia_asignada = "Activo Circulante"
        Evaluacion_Perfil.Hide
        Metodologia_Activo_Circulante.Show
       
       
Else
    MsgBox "Debe Ingresar Actividad Economica para Continuar", vbInformation
       
End If
              
End Sub

Private Sub cmd_metodologia_maxima_produccion_Click()


'If cbx_actividad_economica_formal <> "" Or cbx_actividad_economica_formal_servicio <> "" Or cbx_actividad_economica_semiformal <> "" Or cbx_actividad_economica_informal_oficio <> "" Then

conectarBD

''''' ---- DEUDAS LIBRODEUDORES VIGENTES E INDIRECTAS

        ssql = "SELECT m_deudacreditoconsumo, m_deudacomercial,m_creditohipotecario, m_cupolineacredito, m_deudaindirectavigente, m_deudadirectavigente" _
        & " FROM tbl_micro_fact_ods_librodeudores " _
        & " where rut = '" & txt_rut_cliente & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        If Not rst.EOF Then
    
        Metodologia_Maxima_Prod.txt_deuda_consumo = rst!m_deudacreditoconsumo * 1000
        Metodologia_Maxima_Prod.txt_deuda_comercial = rst!m_deudacomercial * 1000
        Metodologia_Maxima_Prod.txt_credito_hipotecario = rst!m_creditohipotecario * 1000
        Metodologia_Maxima_Prod.txt_cupo_linea_credito = rst!m_cupolineacredito * 1000
        Metodologia_Maxima_Prod.txt_deuda_indirecta = rst!m_deudaindirectavigente * 1000
        Metodologia_Maxima_Prod.txt_deudas_directas_vig = rst!m_deudadirectavigente * 1000

        End If
        
'------------- DEUDAS D10----  CONSUMO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 2 then sum(m_deuda) else 0 end deuda_consumo" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 2" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Maxima_Prod.txt_deuda_d10_consumo = 0

        Else
        Metodologia_Maxima_Prod.txt_deuda_d10_consumo = rst!deuda_consumo
        
        End If
        
        
'------------- DEUDAS D10----  COMERCIAL

        ssql = "SELECT rut," _
                & " case when t_operacion = 1 then sum(m_deuda) else 0 end deuda_comercial" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 1" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Maxima_Prod.txt_deuda_d10_comercial = 0

        Else
        Metodologia_Maxima_Prod.txt_deuda_d10_comercial = rst!deuda_comercial
        
        End If



'------------- DEUDAS D10----  HIPOTECARIO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 3 then sum(m_deuda) else 0 end deuda_hipotecario" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 3" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Maxima_Prod.txt_deuda_d10_hipotecario = 0

        Else
        Metodologia_Maxima_Prod.txt_deuda_d10_hipotecario = rst!deuda_hipotecario
        
        End If
        
        
        
'------------- DEUDAS D10----  LINEA DE CREDITO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 7 then sum(m_deuda) else 0 end deuda_cupo_linea" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 7" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_Maxima_Prod.txt_deuda_d10_linea = 0

        Else
        Metodologia_Maxima_Prod.txt_deuda_d10_linea = rst!deuda_cupo_linea
        
        End If
        
'''''''''''CALCULA TOTAL DE DEUDAS D10
        
        Metodologia_Maxima_Prod.txt_total_deuda_d10 = Int(Val(Metodologia_Maxima_Prod.txt_deuda_d10_consumo) + Val(Metodologia_Maxima_Prod.txt_deuda_d10_comercial) + Val(Metodologia_Maxima_Prod.txt_deuda_d10_hipotecario) + Val(Metodologia_Maxima_Prod.txt_deuda_d10_linea))
        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''sacar fecha-hora
   
        Dim fec1
        Dim hora1
    
        fec1 = Format(Date, "yyyy-mm-dd")
        txt_fecha_actual = fec1


        '''''CARGA_UF
        ssql = "select cast(valor_uf as int) valor_uf from tbl_valor_uf" _
            & " WHERE fecha_dia = '" & txt_fecha_actual & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        'Metodologia_Activo_Circulante.txt_valor_uf = rst!valor_uf
        'Metodologia_IVA1.txt_valor_uf = rst!valor_uf
        Metodologia_Maxima_Prod.txt_valor_uf = rst!valor_uf


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        metodologia_asignada = "Maxima Produccion"
        
        Evaluacion_Perfil.Hide
        Metodologia_Maxima_Prod.Show

'End If

'Else
'MsgBox "Debe Ingresar Actividad Economica para Continuar", vbInformation

End Sub

Private Sub cmd_metodologia_iva_Click()

'If cbx_actividad_economica_formal <> "" Or cbx_actividad_economica_formal_servicio <> "" Or cbx_actividad_economica_semiformal <> "" Or cbx_actividad_economica_informal_oficio <> "" Then


conectarBD

''''' ---- DEUDAS LIBRODEUDORES VIGENTES E INDIRECTAS

        ssql = "SELECT m_deudacreditoconsumo, m_deudacomercial,m_creditohipotecario, m_cupolineacredito, m_deudaindirectavigente, m_deudadirectavigente" _
        & " FROM tbl_micro_fact_ods_librodeudores " _
        & " where rut = '" & txt_rut_cliente & "'"
    
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        If Not rst.EOF Then
    
        Metodologia_IVA1.txt_deuda_consumo = rst!m_deudacreditoconsumo * 1000
        Metodologia_IVA1.txt_deuda_comercial = rst!m_deudacomercial * 1000
        Metodologia_IVA1.txt_credito_hipotecario = rst!m_creditohipotecario * 1000
        Metodologia_IVA1.txt_cupo_linea_credito = rst!m_cupolineacredito * 1000
        Metodologia_IVA1.txt_deuda_indirecta = rst!m_deudaindirectavigente * 1000
        Metodologia_IVA1.txt_deudas_directas_vig = rst!m_deudadirectavigente * 1000
        
        End If

'------------- DEUDAS D10----  CONSUMO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 2 then sum(m_deuda) else 0 end deuda_consumo" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 2" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_IVA1.txt_deuda_d10_consumo = 0

        Else
        Metodologia_IVA1.txt_deuda_d10_consumo = rst!deuda_consumo
        
        End If
        
        
'------------- DEUDAS D10----  COMERCIAL

        ssql = "SELECT rut," _
                & " case when t_operacion = 1 then sum(m_deuda) else 0 end deuda_comercial" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 1" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_IVA1.txt_deuda_d10_comercial = 0

        Else
        Metodologia_IVA1.txt_deuda_d10_comercial = rst!deuda_comercial
        
        End If



'------------- DEUDAS D10----  HIPOTECARIO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 3 then sum(m_deuda) else 0 end deuda_hipotecario" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 3" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_IVA1.txt_deuda_d10_hipotecario = 0

        Else
        Metodologia_IVA1.txt_deuda_d10_hipotecario = rst!deuda_hipotecario
        
        End If
        
        
        
'------------- DEUDAS D10----  LINEA DE CREDITO
        
        ssql = "SELECT rut," _
                & " case when t_operacion = 7 then sum(m_deuda) else 0 end deuda_cupo_linea" _
                & " from TBL_MICRO_FACT_ODS_D10" _
                & " where t_operacion = 7" _
                & " and rut = '" & txt_rut_cliente & "'" _
                & " group by rut, t_operacion"
                
                Set rst = cnn.Execute(ssql, , adCmdText)
                
        If rst.EOF Then
        
        Metodologia_IVA1.txt_deuda_d10_linea = 0

        Else
        Metodologia_IVA1.txt_deuda_d10_linea = rst!deuda_cupo_linea
        
        End If
        
        
'''''''''''CALCULA TOTAL DE DEUDAS D10
        
        Metodologia_IVA1.txt_total_deuda_d10 = Int(Val(Metodologia_IVA1.txt_deuda_d10_consumo) + Val(Metodologia_IVA1.txt_deuda_d10_comercial) + Val(Metodologia_IVA1.txt_deuda_d10_hipotecario) + Val(Metodologia_IVA1.txt_deuda_d10_linea))
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '''sacar fecha-hora
   
        Dim fec1
        Dim hora1
    
        fec1 = Format(Date, "yyyy-mm-dd")
        txt_fecha_actual = fec1


        '''''CARGA_UF
        ssql = "select cast(valor_uf as int) valor_uf from tbl_valor_uf" _
            & " WHERE fecha_dia = '" & txt_fecha_actual & "'"
            
        Set rst = cnn.Execute(ssql, , adCmdText)
    
        'Metodologia_Activo_Circulante.txt_valor_uf = rst!valor_uf
        Metodologia_IVA1.txt_valor_uf = rst!valor_uf
        'Metodologia_Maxima_Prod.txt_valor_uf = rst!valor_uf


'''''''''''''''''''''''''''''''''''''''''''''

        metodologia_asignada = "Iva"
    
        Evaluacion_Perfil.Hide
        Metodologia_IVA1.Show

'Else
'    MsgBox "Debe Ingresar Actividad Economica para Continuar", vbInformation
       
'End If

End Sub

Private Sub cmd_salir_sistema_Click()
    
irespuesta = MsgBox("¿Esta Seguro Que Desea Salir Del Sistema?", vbYesNo)
        
    If irespuesta = vbYes Then
    
        ActiveWorkbook.Save
        Workbooks("Microempresas_1401.xls").Close
        Application.Quit
    
    End If
    
    
End Sub



Private Sub txt_r_pago_mora_Change()

'If txt_r_historia_Pago = "Excelente" And txt_r_mora = "Excelente" Then
'    txt_r_pago_mora = "Excelente"

'    ElseIf txt_r_historia_Pago = "Excelente" And txt_r_mora = "Bueno" Then
'    txt_r_pago_mora = "Bueno"
    
'    Else
'    txt_r_pago_mora = "Regular"
'End If
  
End Sub





Private Sub txt_registro_ventas_var_Change()

End Sub

Private Sub txt_rut_cliente_Change()

End Sub

Private Sub txt_tipo_cliente_Change()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
MsgBox ("Boton Deshabilitado Ocupe Opciones De Menu")
Cancel = True
End If
End Sub


Private Sub UserForm_Initialize()

cbx_Cliente_Nuevo.AddItem "Si"
cbx_Cliente_Nuevo.AddItem "No"

cbx_bancarizado.AddItem "Si"
cbx_bancarizado.AddItem "No"

cbx_historia_pago.AddItem "Sin Mora"
cbx_historia_pago.AddItem "Hasta 5 dias"
cbx_historia_pago.AddItem "Hasta 7 dias"
cbx_historia_pago.AddItem "Hasta 15 dias"

cbx_mora_maxima.AddItem "Sin Mora"
cbx_mora_maxima.AddItem "Hasta 7 dias"
cbx_mora_maxima.AddItem "Hasta 15 dias"
cbx_mora_maxima.AddItem "Hasta 30 dias"

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
'COMENTADO X MAIL DE CBARRIOS cbx_actividad_economica_formal_servicio.AddItem "CONTRATISTA AGRICOLA"
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
cbx_actividad_economica_formal_servicio.AddItem "SEMANERO"
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


End Sub
