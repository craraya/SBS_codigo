Attribute VB_Name = "Módulo1"
  
 ''' DECLARACION DE VARIABLES EN LA CNX BD
 Public cnn As ADODB.Connection
 Public rst As ADODB.Recordset
 Public rstd10 As ADODB.Recordset
 Public rstUF As ADODB.Recordset
 Public rstSCORE As ADODB.Recordset
 Public rstPROT As ADODB.Recordset
 Public ssql As String
 
 
 ''''variable para cambiar clave de ejecutivo
 Public rut_cliente_var  As Variant
 Public dv_cliente_var As Variant
 
 'VARIABLES GLOBALES DEL FORMULARIO FICHA CLIENTE
 Public rut_cliente_ficha As Variant
 Public dv_cliente_ficha  As Variant
 Public n_carpeta_tributaria_ficha As Variant
 Public formalidad_negocio_ficha As Variant
 Public score_dicom_ficha As Variant
 Public antiguedad_meses_ficha As Variant
 Public cbx_codigo_sucursal_ficha As Variant
 Public txt_cod_ejecutivo_ficha As Variant
 Public cbx_Accion_ficha As Variant
 Public r_tipo_credito_destino_ficha As Variant
 Public cbx_tipo_cliente_ficha As Variant
  Public antiguedad_rubro_ficha As Variant
 Public r_actividad_ficha As Variant
 Public Dir_Comercial_Verif_ficha As Variant
 Public visita_eje_ficha As Variant
 Public telef_verif_ficha As Variant
 Public direc_Part_Verif_ficha As Variant
 Public forma_verif_dir_part_ficha As Variant
 Public estado_civil_ficha As Variant
 Public bien_Raiz_ficha As Variant
 Public Acreditacion_Bien_Raiz_ficha As Variant
 Public valor_bien_raiz_ficha As Variant
 Public vehiculos_propios_ficha As Variant
 Public acreditacion_vehiculo_ficha As Variant
 Public antecedentes_int_bancos_ficha As Variant
 Public morosidades_ficha As Variant
 Public protestos_ficha As Variant
 Public boletin_laboral_ficha As Variant
 Public mora_sbif_ficha As Variant
 Public venc_cast_SBIF_ficha As Variant
 Public Mora_Total_Sbif_ficha As Variant
 Public numero_acreedores_ficha As Variant
 Public grado_formalidad_ficha As Variant
 Public tipo_ME_ficha As Variant
 Public n_trabajador_familia_ficha As Variant
 Public envia_cic_ficha As Variant
 Public Estado_AIB_ficha As Variant
 Public Estado_Morosidad_ficha As Variant
 Public Estado_Protestos_ficha As Variant
 Public Estado_Boletin_Laboral_ficha As Variant
 Public estado_mora_sbif_ficha As Variant
 Public estado_venc_cast_SBIF_ficha As Variant
 Public estado_Mora_Total_Sbif_ficha As Variant
 Public estado_numero_acreedores_ficha As Variant
 Public estado_score_dicom_ficha As Variant
 Public estado_credito_ficha As Variant
 Public Credito_Fogape_ficha As Variant
 Public Campana_ficha As Variant
 Public Estado_Meses_Antiguedad_ficha As Variant
 Public conyuge_inf_com As Variant
 
 'VARIABLES GLOBALES DE FORMULARIO DE EVALUACION
 
 Public Cliente_Nuevo_evaluacion As Variant
 Public Bancarizado_evaluacion As Variant
 Public Antiguedad_banco_evaluacion As Variant
 Public mora_promedio_dias_BD_evaluacion As Variant
 Public mora_maxima_dias_BD_evaluacion As Variant
 Public txt_tipo_cliente_evaluacion As Variant
 Public registros_ventas_evaluacion As Variant
 Public R_Final_Perfil_evaluacion As Variant
 Public metodologia_asignada As Variant
 
 'VARIABLES GLOBALES METODOLOGIA ACTIVO CIRCULANTE
 
Public compra_promedio_mensual_MAC As Variant
Public veces_compra_meS_MAC As Variant
Public compra_total_mensual_ctm_MAC As Variant
Public caja_banco_MAC As Variant
Public materia_primas_MAC As Variant
Public mercaderias_MAC As Variant
Public cuenta_cobrar_MAC As Variant
Public otros_activos_circulantes_MAC As Variant
Public total_activos_circulantes_MAC As Variant
Public producto1_MAC As Variant
Public producto2_MAC As Variant
Public producto3_MAC As Variant
Public producto4_MAC As Variant
Public producto5_MAC As Variant
Public precio_venta1_MAC As Variant
Public precio_venta2_MAC As Variant
Public precio_venta3_MAC As Variant
Public precio_venta4_MAC As Variant
Public precio_venta5_MAC As Variant
Public materia_prima1_MAC As Variant
Public materia_prima2_MAC As Variant
Public materia_prima3_MAC As Variant
Public materia_prima4_MAC As Variant
Public materia_prima5_MAC As Variant
Public mano_obra1_MAC As Variant
Public mano_obra2_MAC As Variant
Public mano_obra3_MAC As Variant
Public mano_obra4_MAC As Variant
Public mano_obra5_MAC As Variant
Public incidencia_ventas1_MAC As Variant
Public incidencia_ventas2_MAC As Variant
Public incidencia_ventas3_MAC As Variant
Public incidencia_ventas4_MAC As Variant
Public incidencia_ventas5_MAC As Variant
Public r_cvcmo1_MAC As Variant
Public r_cvcmo2_MAC As Variant
Public r_cvcmo3_MAC As Variant
Public r_cvcmo4_MAC As Variant
Public r_cvcmo5_MAC As Variant
Public r_cvsmo1_MAC As Variant
Public r_cvsmo2_MAC As Variant
Public r_cvsmo3_MAC As Variant
Public r_cvsmo4_MAC As Variant
Public r_cvsmo5_MAC As Variant
Public r_cvppcmo1_MAC As Variant
Public r_cvppcmo2_MAC As Variant
Public r_cvppcmo3_MAC As Variant
Public r_cvppcmo4_MAC As Variant
Public r_cvppcmo5_MAC As Variant
Public r_cvppsmo1_MAC As Variant
Public r_cvppsmo2_MAC As Variant
Public r_cvppsmo3_MAC As Variant
Public r_cvppsmo4_MAC As Variant
Public r_cvppsmo5_MAC As Variant
Public Sub_Total_costo_variable_MAC As Variant
Public Sub_Total_x1_costo_variable_MAC As Variant
Public compra_total_mensual_MAC As Variant
Public venta_total_alto_MAC As Variant
Public venta_total_medio_MAC As Variant
Public venta_total_bajo_MAC As Variant
Public compra_total_max_corregida_MAC As Variant
Public venta_total_mes_alto_corregida_MAC As Variant
Public venta_total_mes_medio_corregida_MAC As Variant
Public venta_total_mes_bajo_corregida_MAC As Variant
Public arriendo_micro_MAC As Variant
Public sueldos_MAC As Variant
Public movilizacion_MAC As Variant
Public servicios_basicos_MAC As Variant
Public contador_MAC As Variant
Public lubricantes_MAC As Variant
Public neumaticos_MAC As Variant
Public afinamientos_MAC As Variant
Public patentes_seguros_MAC As Variant
Public otros_costos_fijos_MAC As Variant
Public total_costos_fijos_MAC As Variant
Public valor_uf_MAC As Variant
Public n_grupo_familiar_MAC As Variant
Public arriendo_vivienda_MAC As Variant
Public gastos_indicado_cliente_MAC As Variant
Public total_gasto_familiar_MAC As Variant
Public liquidacion_sueldo_MAC As Variant
Public jubilacion_MAC As Variant
Public montepio_MAC As Variant
Public arriendo_vivienda1_MAC As Variant
Public ingreso_segunda_microempresa_MAC As Variant
Public boleta_honorario_MAC As Variant
Public acreedor1_deuda_MAC As Variant
Public acreedor2_deuda_MAC As Variant
Public acreedor3_deuda_MAC As Variant
Public acreedor4_deuda_MAC As Variant
Public acreedor5_deuda_MAC As Variant
Public acreedor6_deuda_MAC As Variant
Public tipo_producto1_deuda_MAC As Variant
Public tipo_producto2_deuda_MAC As Variant
Public tipo_producto3_deuda_MAC As Variant
Public tipo_producto4_deuda_MAC As Variant
Public tipo_producto5_deuda_MAC As Variant
Public tipo_producto6_deuda_MAC As Variant
Public saldo_pendiente1_deuda_MAC As Variant
Public saldo_pendiente2_deuda_MAC As Variant
Public saldo_pendiente3_deuda_MAC As Variant
Public saldo_pendiente4_deuda_MAC As Variant
Public saldo_pendiente5_deuda_MAC As Variant
Public saldo_pendiente6_deuda_MAC As Variant
Public monto_cuota1_deuda_MAC As Variant
Public monto_cuota2_deuda_MAC As Variant
Public monto_cuota3_deuda_MAC As Variant
Public monto_cuota4_deuda_MAC As Variant
Public monto_cuota5_deuda_MAC As Variant
Public monto_cuota6_deuda_MAC As Variant
Public cuotas_pactadas1_deuda_MAC As Variant
Public cuotas_pactadas2_deuda_MAC As Variant
Public cuotas_pactadas3_deuda_MAC As Variant
Public cuotas_pactadas4_deuda_MAC As Variant
Public cuotas_pactadas5_deuda_MAC As Variant
Public cuotas_pactadas6_deuda_MAC As Variant
Public cuotas_pendientes1_deuda_MAC As Variant
Public cuotas_pendientes2_deuda_MAC As Variant
Public cuotas_pendientes3_deuda_MAC As Variant
Public cuotas_pendientes4_deuda_MAC As Variant
Public cuotas_pendientes5_deuda_MAC As Variant
Public cuotas_pendientes6_deuda_MAC As Variant
Public prepaga_deuda1_deuda_MAC As Variant
Public prepaga_deuda2_deuda_MAC As Variant
Public prepaga_deuda3_deuda_MAC As Variant
Public prepaga_deuda4_deuda_MAC As Variant
Public prepaga_deuda5_deuda_MAC As Variant
Public prepaga_deuda6_deuda_MAC As Variant
Public total_saldo_pendiente_deuda_MAC As Variant
Public total_deudas_MAC As Variant
Public numero_meses_tipo_mes_alto_flujo_MAC As Variant
Public numero_meses_tipo_mes_medio_flujo_MAC As Variant
Public numero_meses_tipo_mes_bajo_flujo_MAC As Variant
Public vta_formal_promedio_mes_alto_MAC As Variant
Public vta_formal_promedio_mes_medio_MAC As Variant
Public vta_formal_promedio_mes_bajo_MAC As Variant
Public vta_informal_promedio_mes_alto_MAC As Variant
Public vta_informal_promedio_mes_medio_MAC As Variant
Public vta_informal_promedio_mes_bajo_MAC As Variant
Public Venta_Total_Promedio_Mes_Alto_MAC As Variant
Public Venta_Total_Promedio_Mes_medio_MAC As Variant
Public Venta_Total_Promedio_Mes_bajo_MAC As Variant
Public resultado_operacional_alto_MAC As Variant
Public resultado_operacional_medio_MAC As Variant
Public resultado_operacional_bajo_MAC As Variant
Public capacidad_pago_mes_alto_MAC As Variant
Public capacidad_pago_mes_medio_MAC As Variant
Public capacidad_pago_mes_bajo_MAC As Variant
Public capacidad_pago_corregida_ajustada_mes_alto_MAC As Variant
Public capacidad_pago_corregida_ajustada_mes_medio_MAC As Variant
Public capacidad_pago_corregida_ajustada_mes_bajo_MAC As Variant
Public capacidad_pago_promedio_corregida_ajustada_MAC As Variant
Public monto_maximo_credito_MAC As Variant
Public cuota_credito_MAC As Variant
Public mto_bruto_solicitado_cliente_MAC As Variant
Public resolucion_credito_cuota_MAC As Variant
Public resolucion_credito_monto_MAC As Variant

Public txt_score_dicom_conyuge_aux As Variant
Public txt_score_dicom_AVAL_aux As Variant





  
 Sub autoformu_excel_invisible()
' Activa el autoformulario de Excel y hace invisible el libro, _
al cerrar el autoformulario se hace visible otra vez Excel
   Application.Visible = False
  On Error Resume Next
  ActiveSheet.ShowDataForm
   Application.Visible = True
  
End Sub


Sub Auto_Open()
Application.Visible = False
Acceso_Principal.Show

End Sub


Public Sub conectarBD()
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

With cnn
    
    cnn.Open "Provider=SQLNCLI11;" & _
    "Data Source=CLMOBCAMASD01P;" & _
    "Initial Catalog=BD_GES_FAM;" & _
    "User Id=GEN_CF;Password=Gen_cF1p$"
        
End With

End Sub

Public Sub conectarBDSYBASE()
'Set cnn = New ADODB.Connection
'Set rst = New ADODB.Recordset

'With cnn
    'Data Source=CLDC101;Port=5000;Database=BD_PRI_38;Uid=myUsername;Pwd=myPassword;
    'cnn.Open "Provider=Sybase.ASEOLEDBProvider;" & _
    "Server Name=CLDC101.chl.bns,60000;" & _
    "Initial Catalog=BD_PRI_38;" & _
    "User Id=usrinfcomer;Password=infor123"
    
'    Provider=Sybase.ASEOLEDBProvider;Server Name=myASEserver,5000;Initial Catalog=myDataBase;User Id=myUsername;Password=myPassword;
    
'Set varconect = New ADODB.Connection
'Set varGAdoConexion1 = New ADODB.Connection
'varconect.ConnectionString = "driver={SQL Server};server=CLDC101;Uid=usrinfcomer;Pwd=infor123;datadase=BD_PRI_38"
'varconect.CommandTimeout = 30
'varconect.Open
'frmbusca.Show
'varconect.CursorLocation = adUseServer
        
'End With

End Sub

Public Sub conectarBDRIESGO()
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

With cnn
    
    cnn.Open "Provider=SQLNCLI;" & _
    "Data Source=CLLOGDR02P;" & _
    "Initial Catalog=RT_OPERACIONES;" & _
    "User Id=consubi;Password=consubi1$"
        
End With

End Sub

Public Sub conectarBDRIESGO2()
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

With cnn
    
    'cnn.Open "PROVIDER=SQLOLEDB.1;DATA SOURCE=CLMOCRCD01P.chl.bns; UID=sys_comercial; PWD=scom.4321; DATABASE=Miscobranza"
    cnn.Open "PROVIDER=SQLOLEDB.1;DATA SOURCE=CLLOGDR02P.chl.bns; UID=sys_comercial; PWD=scom.4321; DATABASE=Miscobranza"
End With

End Sub


Public Function GVerificar_NULL(Campo As Variant, ByVal Def As String) As String

If IsNull(Campo) Then
    GVerificar_NULL = Def
Else
    GVerificar_NULL = Campo
End If

End Function



'Public Function GCentrar_Form(Formulario As Form)

'Formulario.Top = (Screen.Height - Formulario.Height) / 2
'Formulario.Left = (Screen.Width - Formulario.Width) / 2

'End Function



Public Function GValNum(ValKey As Integer, Optional CarEspecial As String) As Integer
Dim X As Integer
Dim largo As Integer
    
    If CarEspecial <> "" Then
    largo = Len(CarEspecial)
        For X = 1 To largo
            If Asc(Mid(CarEspecial, X, 1)) = ValKey Then
                GValNum = ValKey
                Exit Function
            End If
        Next
    End If
    
    If ValKey <> 8 And ValKey <> 13 And (ValKey < 48 Or ValKey > 57) Then
        ValKey = 0
    End If
GValNum = ValKey
End Function
Public Function GValLetra(ValKey As Integer, _
                                Optional CarEspecial As String, _
                                Optional Mayus As Boolean = True) As Integer

If ValKey = 39 Or ValKey = 34 Then ValKey = 0
If Mayus Then ValKey = Asc(UCase(Chr(ValKey)))
   
GValLetra = ValKey

End Function






