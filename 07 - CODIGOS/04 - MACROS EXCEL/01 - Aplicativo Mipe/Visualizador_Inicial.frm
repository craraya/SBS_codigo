VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Visualizador_Inicial 
   Caption         =   ":::::::::Visualizacion De Evaluacion Microempresa"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   OleObjectBlob   =   "Visualizador_Inicial.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Visualizador_Inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_evaluacion_Click()


        'Call limpiar_variables
                    
            R_Final_Perfil_1 = Empty
            txt_Antiguedad.Visible = True
            txt_historia_pago.Visible = True
            txt_mora_maxima.Visible = True
            txt_tipo_cliente_riesgo.Visible = True
            txt_registro_ventas_v.Visible = True
            txt_antiguedad_negocio_riesgo.Visible = True
            

        
        '' Apaga Botones Metodologia
        cmd_metodologia_iva.Enabled = False
        cmd_metodologia_activo_circulante.Enabled = False
        cmd_metodologia_maxima_produccion.Enabled = False

        

        Call conectarBD
        
        ssql = "Select a.n_solicitud,a.rut_cliente from tbl_micro_ficha_cliente a left join tbl_micro_metodologia_activo_circulante b on a.n_solicitud = b.n_solicitud" _
        & " left join tbl_micro_metodologia_maxima_produccion c on a.n_solicitud = c.n_solicitud" _
        & " left join tbl_micro_metodologia_iva d on a.n_solicitud = d.n_solicitud" _
        & " where (b.n_solicitud Is Not Null Or c.n_solicitud Is Not Null Or d.n_solicitud Is Not Null)" _
        & " and a.rut_cliente ='" & txt_rut_cliente & "'" _
        & " group by  a.n_solicitud,a.rut_cliente" _
        & " order by a.n_solicitud desc"

        Set rst = cnn.Execute(ssql, , adCmdText)
        
        If rst.EOF Then
        
         MsgBox ("Cliente Sin Ingreso De Evaluacion ó No Enviado Al S.I.C."), vbCritical
         
         Else
       
        txt_rut_cliente_verif = rst!rut_cliente
        txt_sol_verif = rst!n_solicitud
        
        ssql = "SELECT tbl_micro_ficha_cliente.rut_cliente,tbl_micro_ficha_cliente.dv,tbl_micro_ficha_cliente.n_solicitud," _
        & " n_carpeta_tributaria,id_sucursal,codigo_ejecutivo,Destino_Credito," _
        & " Formalidad_Negocio, Tipo_Credito_Destino,Antiguedad_Negocio,Certificacion_Antiguedad_Rubro," _
        & " R_Actividad, Dir_Comercial_Verif, visita_eje, telef_verif, direc_Part_Verif, forma_verif_dir_part," _
        & " estado_civil, bien_Raiz, Acreditacion_Bien_Raiz, valor_bien_raiz, vehiculos_propios, acreditacion_vehiculo," _
        & " antecedentes_int_bancos, Estado_AIB, morosidades, Estado_Morosidad, protestos,Estado_Protestos," _
        & " boletin_laboral, Estado_Boletin_Laboral, mora_sbif, estado_mora_sbif, venc_cast_SBIF, estado_venc_cast_SBIF," _
        & " Mora_Total_Sbif, estado_Mora_Total_Sbif, numero_acreedores, estado_numero_acreedores, score_dicom," _
        & " estado_score_dicom, grado_formalidad, tipo_ME, n_trabajador_familia, estado_credito, envia_cic, tbl_micro_ficha_cliente.Fecha_Ingreso," _
        & " tbl_micro_ficha_cliente.Hora_Ingreso, isnull(Credito_Fogape,'SD') as Credito_Fogape, isnull(Campana,'SD') as Campana, isnull(estado_antiguedad_negocio,'SD') as estado_antiguedad_negocio," _
        & " cliente_nuevo, bancarizado, Antiguedad_Banco, Mora_Promedio_Dias_BD, Mora_Maxima_Dias_BD," _
        & " R_Tipo_Cliente,Registro_Ventas, R_Final_Perfil, Metodologia_Asignada, isnull(conyuge_inf_com,'SD')as conyuge_inf_com, isnull(Estado_conyuge_inf_com,'SD') as Estado_conyuge_inf_com, isnull(antiguedad_bco,'SD') as Antiguedad_Bco, isnull(tipo_excepcion,'') as tipo_excepcion, isnull(ejecutivo_excepcion,'') as ejecutivo_excepcion, isnull(Credito_Comercial,'SD')Credito_Comercial,isnull(Credito_Consumo,'SD')Credito_Consumo,ISNULL(Accion_consumo,'SD')Accion_consumo,ISNULL(plazo_credito_consumo,'SD')plazo_credito_consumo,ISNULL(bancarizado_FC_PNB,'SD') bancarizado_FC_PNB,ISNULL(r_formalidad_negocio_PNB,'SD') r_formalidad_negocio_PNB,ISNULL(r_cbx_antiguedad_rubro_PNB,'SD') r_cbx_antiguedad_rubro_PNB,ISNULL(r_cbx_actividad_economica_PNB,'SD') r_cbx_actividad_economica_PNB,ISNULL(r_cbx_bien_Raiz_PNB,'SD') r_cbx_bien_Raiz_PNB,ISNULL(r_cbx_vehiculos_propios_PNB,'SD') r_cbx_vehiculos_propios_PNB,ISNULL(ESTADO_politica_bancarizado_new_PNB,'SD') ESTADO_politica_bancarizado_new_PNB" _
        & " from tbl_micro_ficha_cliente , tbl_micro_perfil_riesgo_cliente " _
        & " where tbl_micro_ficha_cliente.n_solicitud = '" & txt_sol_verif & "'" _
        & " and   tbl_micro_ficha_cliente.rut_cliente = '" & txt_rut_cliente_verif & "'" _
        & " and tbl_micro_perfil_riesgo_cliente.n_solicitud = '" & txt_sol_verif & "'" _
        & " and   tbl_micro_perfil_riesgo_cliente.rut_cliente = '" & txt_rut_cliente_verif & "'"

        
        Set rst = cnn.Execute(ssql, , adCmdText)
        
            txt_rut_cliente_v = rst!rut_cliente
            txt_rut_cliente_v_1 = rst!rut_cliente
            txt_dv_v = rst!dv
            txt_dv_v_1 = rst!dv
            txt_nsolicitud = rst!n_solicitud
            txt_n_carpeta_tributaria = rst!n_carpeta_tributaria
            txt_n_carpeta_tributaria_1 = rst!n_carpeta_tributaria
            txt_codigo_sucursal = rst!id_sucursal
            txt_cod_ejecutivo = rst!codigo_ejecutivo
            txt_Accion = rst!Destino_Credito
            txt_tipo_cliente_ficha = rst!Formalidad_Negocio
            txt_actividad_economica_informal_oficio = rst!r_actividad
            txt_Destino_Credito = rst!tipo_credito_destino
            txt_antiguedad_meses = rst!antiguedad_negocio
            txt_Antiguedad_Rubro = rst!certificacion_antiguedad_rubro
            txt_Dir_Comercial_Verif = rst!Dir_Comercial_Verif
            txt_visita_eje = rst!visita_eje
            txt_telef_verif = rst!telef_verif
            txt_direc_Part_Verif = rst!direc_Part_Verif
            txt_forma_verif_dir_part = rst!forma_verif_dir_part
            txt_estado_civil = rst!estado_civil
            txt_bien_Raiz = rst!bien_Raiz
            txt_acred_bien_raiz = rst!Acreditacion_Bien_Raiz
            txt_acred_bien_raiz_no = rst!Acreditacion_Bien_Raiz
            txt_valor_evaluo_bien_raiz = rst!valor_bien_raiz
            txt_valor_evaluo_bien_raiz_no = rst!valor_bien_raiz
            txt_vehiculos_propios = rst!vehiculos_propios
            txt_acreditacion_vehiculo = rst!acreditacion_vehiculo
            txt_acreditacion_vehiculo_no = rst!acreditacion_vehiculo
            txt_antecedentes_int_bancos = rst!antecedentes_int_bancos
            txt_r_aib_v = rst!Estado_AIB
            txt_morosidades = rst!morosidades
            txt_r_morosidad_v = rst!Estado_Morosidad
            txt_protestos = rst!protestos
            txt_r_protestos_v = rst!Estado_Protestos
            txt_boletin_laboral = rst!boletin_laboral
            txt_r_boletin_laboral_v = rst!Estado_Boletin_Laboral
            txt_mora_sbif = rst!mora_sbif
            txt_r_mora_sbif_v = rst!estado_mora_sbif
            txt_venc_cast_SBIF = rst!venc_cast_SBIF
            txt_r_venc_cast_SBIF_v = rst!estado_venc_cast_SBIF
            txt_Mora_Total_Sbif = rst!Mora_Total_Sbif
            txt_r_Mora_Total_Sbif_v = rst!estado_Mora_Total_Sbif
            txt_r_meses_antiguedad_v = rst!estado_antiguedad_negocio
            txt_numero_acreedores = rst!numero_acreedores
            txt_r_n_acreedores_v = rst!estado_numero_acreedores
            txt_score_dicom = rst!score_dicom
            txt_r_predictor_score_dicom_v = rst!estado_score_dicom
            txt_grado_formalidad = rst!grado_formalidad
            txt_tipo_ME = rst!tipo_ME
            txt_n_trabajador_familia = rst!n_trabajador_familia
            txt_envia_cic = rst!envia_Cic
            txt_estado_credito = rst!estado_credito
            txt_fecha_ingreso = rst!FECHA_INGRESO
            txt_hora_ingreso = rst!HORA_INGRESO
            txt_campana = rst!campana
            txt_credito_fogape = rst!credito_fogape
            'CAMPOS TABLA : TBL_MICRO_PERFIL_RIESGO_CLIENTE
            txt_Cliente_Nuevo = rst!Cliente_Nuevo
            txt_bancarizado = rst!bancarizado
            txt_Antiguedad = rst!antiguedad_banco
            txt_historia_pago = rst!mora_promedio_dias_BD
            txt_mora_maxima = rst!mora_maxima_dias_BD
            
            If rst!R_Tipo_Cliente = "Antiguo Prime" Or rst!R_Tipo_Cliente = "Antiguo No Prime" Then
                txt_tipo_cliente_riesgo = "Antiguo"
            Else
                txt_tipo_cliente_riesgo = rst!R_Tipo_Cliente
            End If
           
            txt_registro_ventas_v = rst!Registro_Ventas
            txt_antiguedad_negocio_riesgo = rst!antiguedad_negocio
            txt_predictor_Score = rst!score_dicom
            txt_tipo_cliente_actividad = rst!r_actividad
            txt_actividad_economica_formal = rst!Formalidad_Negocio
            R_Final_Perfil = rst!R_Final_Perfil
            R_Final_Perfil_1 = rst!metodologia_asignada
            txt_conyuge_inf_com = rst!conyuge_inf_com
            txt_r_conyuge_inf_com = rst!Estado_conyuge_inf_com
            txt_antiguedad_banco = rst!antiguedad_bco
            txt_tipo_excepcion = rst!tipo_excepcion
            txt_ejecutivo_excepcion = rst!ejecutivo_excepcion
            cbx_pregunta_comercial = rst!Credito_Comercial
            cbx_pregunta_consumo = rst!Credito_Consumo
            cbx_Accion_consumo = rst!Accion_consumo
            txt_plazo_credito_consumo = rst!plazo_credito_consumo
            txt_r_formalidad_negocio_v = rst!r_formalidad_negocio_PNB
            txt_r_cbx_antiguedad_rubro_v = rst!r_cbx_antiguedad_rubro_PNB
            txt_r_cbx_actividad_economica_informal_oficio_v = rst!r_cbx_actividad_economica_PNB
            txt_r_cbx_bien_Raiz_v = rst!r_cbx_bien_Raiz_PNB
            txt_r_cbx_vehiculos_propios_v = rst!r_cbx_vehiculos_propios_PNB
            txt_r_ESTADO_politica_bancarizado_new_v = rst!ESTADO_politica_bancarizado_new_PNB
        
        
            ssql = " select Out_IndRiesgo from TBL_MICRO_EVALUACION_RIESGO" _
            & " where Rut_Num=" & txt_rut_cliente_verif _
            & " and " _
            & " N_Solicitud=" & txt_sol_verif
            
            Set rst = cnn.Execute(ssql, , adCmdText)
            
            If Not rst.EOF Then
            
                If rst!Out_IndRiesgo = "1" Then
                    txt_Risk_Indicator = "A"
                ElseIf rst!Out_IndRiesgo = "2" Then
                    txt_Risk_Indicator = "B"
                ElseIf rst!Out_IndRiesgo = "3" Then
                    txt_Risk_Indicator = "C"
                ElseIf rst!Out_IndRiesgo = "4" Then
                    txt_Risk_Indicator = "D"
                ElseIf rst!Out_IndRiesgo = "5" Then
                    txt_Risk_Indicator = "E"
                Else
                    txt_Risk_Indicator = ""
                End If
            
            Else
                txt_Risk_Indicator.Visible = False
                lbl_Risk_Indicator.Visible = False
            End If
            
               
           
                
            
            If txt_Antiguedad = "" Then
                txt_Antiguedad.Visible = False
                'lbl_meses_ant_bco.Visible = False
            End If
            
            If txt_historia_pago = "" Then
                txt_historia_pago.Visible = False
            End If
            
            If txt_mora_maxima = "" Then
                txt_mora_maxima.Visible = False
            End If
            
            If txt_tipo_cliente_riesgo = "" Then
                txt_tipo_cliente_riesgo.Visible = False
            End If
            
            If txt_registro_ventas_v = "" Then
                txt_registro_ventas_v.Visible = False
            End If
            
            If txt_antiguedad_negocio_riesgo = "" Then
                txt_antiguedad_negocio_riesgo.Visible = False
                lbl_mes_antig_neg.Visible = False
            End If
            
            
            '''''''' SELECCION DE EJECUTIVO''''''''''
            ssql = "SELECT ISNULL(nombre_ejecutivo+' '+apellido_ejecutivo,'NO REGISTRADO') as nombre_ejecutivo" _
            & " from tbl_ejecutivo a, tbl_micro_ficha_cliente b" _
            & " where a.codigo_sucursal+a.codigo_ejecutivo = cast(b.id_sucursal as varchar)+cast(b.codigo_ejecutivo as varchar)" _
            & " and '" & txt_nsolicitud & "' = b.n_solicitud"
            '& " and estado_ejecutivo = 'Activo'"
            
            Set rst = cnn.Execute(ssql, , adCmdText)
        
            If rst.EOF Then
                txt_cod_ejecutivo_nombre = "NO REGISTRADO"
            Else
                txt_cod_ejecutivo_nombre = rst!nombre_ejecutivo
                rst.MoveNext
            End If

            
        End If
        
End Sub

Private Sub cmd_imprimir1_Click()
Visualizador_Inicial.PrintForm
End Sub

Private Sub cmd_metodologia_activo_circulante_Click()

Visualizador_Met_AC.txt_n_solicitud = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_AC.txt_rut_cliente_ing = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_AC.txt_dv_ing = Visualizador_Inicial.txt_dv_v

Visualizador_Met_AC.txt_n_solicitud_1 = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_AC.txt_rut_cliente_ing_1 = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_AC.txt_dv_ing_1 = Visualizador_Inicial.txt_dv_v

Visualizador_Met_AC.txt_codigo_sucursal_1 = Visualizador_Inicial.txt_codigo_sucursal
Visualizador_Met_AC.txt_cod_ejecutivo_nombre_1 = Visualizador_Inicial.txt_cod_ejecutivo_nombre
Visualizador_Met_AC.txt_fecha_ingreso_1 = Visualizador_Inicial.txt_fecha_ingreso
Visualizador_Met_AC.txt_hora_ingreso_1 = Visualizador_Inicial.txt_hora_ingreso

Visualizador_Inicial.Hide
Visualizador_Met_AC.Show
End Sub

Private Sub cmd_metodologia_iva_Click()

Visualizador_Met_IVA.txt_n_solicitud = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_IVA.txt_rut_cliente_ing = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_IVA.txt_dv_ing = Visualizador_Inicial.txt_dv_v

Visualizador_Met_IVA.txt_n_solicitud_1 = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_IVA.txt_rut_cliente_ing_1 = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_IVA.txt_dv_ing_1 = Visualizador_Inicial.txt_dv_v

Visualizador_Met_IVA.txt_codigo_sucursal_1 = Visualizador_Inicial.txt_codigo_sucursal
Visualizador_Met_IVA.txt_cod_ejecutivo_nombre_1 = Visualizador_Inicial.txt_cod_ejecutivo_nombre
Visualizador_Met_IVA.txt_fecha_ingreso_1 = Visualizador_Inicial.txt_fecha_ingreso
Visualizador_Met_IVA.txt_hora_ingreso_1 = Visualizador_Inicial.txt_hora_ingreso

Visualizador_Inicial.Hide
Visualizador_Met_IVA.Show

End Sub

Private Sub cmd_metodologia_maxima_produccion_Click()

Visualizador_Met_MP.txt_n_solicitud = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_MP.txt_rut_cliente_ing = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_MP.txt_dv_ing = Visualizador_Inicial.txt_dv_v

Visualizador_Met_MP.txt_n_solicitud_1 = Visualizador_Inicial.txt_nsolicitud
Visualizador_Met_MP.txt_rut_cliente_ing_1 = Visualizador_Inicial.txt_rut_cliente_v
Visualizador_Met_MP.txt_dv_ing_1 = Visualizador_Inicial.txt_dv_v

Visualizador_Met_MP.txt_codigo_sucursal_1 = Visualizador_Inicial.txt_codigo_sucursal
Visualizador_Met_MP.txt_cod_ejecutivo_nombre_1 = Visualizador_Inicial.txt_cod_ejecutivo_nombre
Visualizador_Met_MP.txt_fecha_ingreso_1 = Visualizador_Inicial.txt_fecha_ingreso
Visualizador_Met_MP.txt_hora_ingreso_1 = Visualizador_Inicial.txt_hora_ingreso

Visualizador_Inicial.Hide
Visualizador_Met_MP.Show

End Sub

Private Sub cmd_volver_menu_principal_Click()
Unload Visualizador_Inicial
Menu_Principal_Micro.Show
End Sub

Private Sub CommandButton1_Click()
Visualizador_Inicial.PrintForm
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub R_Final_Perfil_1_Change()
If R_Final_Perfil_1 = "Maxima Produccion" Then
    cmd_metodologia_iva.Enabled = False
    cmd_metodologia_activo_circulante.Enabled = False
    cmd_metodologia_maxima_produccion.Enabled = True
   
ElseIf R_Final_Perfil_1 = "Activo Circulante" Then

    cmd_metodologia_iva.Enabled = False
    cmd_metodologia_activo_circulante.Enabled = True
    cmd_metodologia_maxima_produccion.Enabled = False

ElseIf R_Final_Perfil_1 = "Iva" Then

    cmd_metodologia_iva.Enabled = True
    cmd_metodologia_activo_circulante.Enabled = False
    cmd_metodologia_maxima_produccion.Enabled = False
    
     End If
End Sub




Private Sub txt_dv_Change()

Dim I As Integer
txt_dv = UCase(txt_dv)
I = Len(txt_dv)
txt_dv.SelStart = I

If txt_dv <> txt_dv_compara Then
  MsgBox "Rut ó Dv  Incorrecto, Reingrese", vbCritical
  txt_rut_cliente.SetFocus

End If

End Sub

Private Sub txt_rut_cliente_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
       
    If Not IsNumeric(txt_rut_cliente) Then
        
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_cliente = Empty
        txt_rut_cliente.SetFocus
      
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
End Sub


Private Sub limpiar_variables()

            txt_rut_cliente_v = Empty
            txt_rut_cliente_v_1 = Empty
            txt_dv_v = Empty
            txt_dv_v_1 = Empty
            txt_nsolicitud = Empty
            txt_n_carpeta_tributaria = Empty
            txt_n_carpeta_tributaria_1 = Empty
            txt_codigo_sucursal = Empty
            txt_cod_ejecutivo = Empty
            txt_Accion = Empty
            txt_tipo_cliente_ficha = Empty
            txt_actividad_economica_informal_oficio = Empty
            txt_Destino_Credito = Empty
            txt_antiguedad_meses = Empty
            txt_Antiguedad_Rubro = Empty
            txt_Dir_Comercial_Verif = Empty
            txt_visita_eje = Empty
            txt_telef_verif = Empty
            txt_direc_Part_Verif = Empty
            txt_forma_verif_dir_part = Empty
            txt_estado_civil = Empty
            txt_bien_Raiz = Empty
            txt_acred_bien_raiz = Empty
            txt_acred_bien_raiz_no = Empty
            txt_valor_evaluo_bien_raiz = Empty
            txt_valor_evaluo_bien_raiz_no = Empty
            txt_vehiculos_propios = Empty
            txt_acreditacion_vehiculo = Empty
            txt_acreditacion_vehiculo_no = Empty
            txt_antecedentes_int_bancos = Empty
            txt_r_aib_v = Empty
            txt_morosidades = Empty
            txt_r_morosidad_v = Empty
            txt_protestos = Empty
            txt_r_protestos_v = Empty
            txt_boletin_laboral = Empty
            txt_r_boletin_laboral_v = Empty
            txt_mora_sbif = Empty
            txt_r_mora_sbif_v = Empty
            txt_venc_cast_SBIF = Empty
            txt_r_venc_cast_SBIF_v = Empty
            txt_Mora_Total_Sbif = Empty
            txt_r_Mora_Total_Sbif_v = Empty
            txt_r_meses_antiguedad_v = Empty
            txt_numero_acreedores = Empty
            txt_r_n_acreedores_v = Empty
            txt_score_dicom = Empty
            txt_r_predictor_score_dicom_v = Empty
            txt_grado_formalidad = Empty
            txt_tipo_ME = Empty
            txt_n_trabajador_familia = Empty
            txt_envia_cic = Empty
            txt_estado_credito = Empty
            txt_fecha_ingreso = Empty
            txt_hora_ingreso = Empty
            txt_campana = Empty
            txt_credito_fogape = Empty
            txt_Cliente_Nuevo = Empty
            txt_bancarizado = Empty
            txt_Antiguedad = Empty
            txt_historia_pago = Empty
            txt_mora_maxima = Empty
            txt_tipo_cliente_riesgo = Empty
            txt_registro_ventas_v = Empty
            txt_antiguedad_negocio_riesgo = Empty
            txt_predictor_Score = Empty
            txt_tipo_cliente_actividad = Empty
            txt_actividad_economica_formal = Empty
            R_Final_Perfil = Empty
            R_Final_Perfil_1 = Empty
            txt_cod_ejecutivo_nombre = Empty
            txt_tipo_excepcion = Empty
            txt_ejecutivo_excepcion = Empty
End Sub

