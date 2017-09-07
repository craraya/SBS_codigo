VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Visualizador_Met_AC 
   Caption         =   ":::::: Visualizador Metodologia Activo Circulante"
   ClientHeight    =   11895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   OleObjectBlob   =   "Visualizador_Met_AC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Visualizador_Met_AC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_visualizador_AC_Click()

        Call conectarBD

        ssql = "SELECT RUT_CLIENTE,N_SOLICITUD,DV,compra_promedio_mensual,veces_compra_mes,compra_total_mensual_ctm,caja_banco,materia_prima,mercaderias,cuenta_cobrar,otros_activos_circulantes,total_activo_circulante,producto1,producto2,producto3,producto4,producto5,precio_venta1,precio_venta2,precio_venta3,precio_venta4," _
            & " precio_venta5,materia_prima1,materia_prima2,materia_prima3,materia_prima4,materia_prima5,mano_obra1,mano_obra2,mano_obra3,mano_obra4,mano_obra5,incidencia_ventas1,incidencia_ventas2,incidencia_ventas3,incidencia_ventas4,incidencia_ventas5,r_cvcmo1,r_cvcmo2,r_cvcmo3,r_cvcmo4,r_cvcmo5,r_cvsmo1,r_cvsmo2,r_cvsmo3,r_cvsmo4," _
            & " r_cvsmo5,r_cvppcmo1,r_cvppcmo2,r_cvppcmo3,r_cvppcmo4,r_cvppcmo5,r_cvppsmo1,r_cvppsmo2,r_cvppsmo3,r_cvppsmo4,r_cvppsmo5,r_Subtotal_costo_variable,r_Subtotal_x1_costo_variable,r_compra_total_mensual,r_venta_total_alto,r_venta_total_medio,r_venta_total_bajo,r_compra_total_max_corregida,r_venta_total_alto_corregida," _
            & " r_venta_total_medio_corregida,r_venta_total_bajo_corregida,arriendo_micro,sueldos,movilizacion,servicios_basicos,contador,lubricantes,neumaticos,afinamientos,patentes_seguros,otros_costos_fijos,total_costos_fijos,valor_uf,n_grupo_familiar,arriendo_vivienda_Gastos_Fam,gastos_indicado_cliente,total_gasto_familiar,liquidacion_sueldo," _
            & " jubilacion,montepio,arriendo_vivienda_Otro_Ing,ingreso_segunda_microempresa,boleta_honorario,total_otros_ingresos,acreedor1_deuda,acreedor2_deuda,acreedor3_deuda,acreedor4_deuda,acreedor5_deuda,acreedor6_deuda,tipo_producto1_deuda,tipo_producto2_deuda,tipo_producto3_deuda,tipo_producto4_deuda,tipo_producto5_deuda,tipo_producto6_deuda," _
            & " saldo_pendiente1_deuda,saldo_pendiente2_deuda,saldo_pendiente3_deuda,saldo_pendiente4_deuda,saldo_pendiente5_deuda,saldo_pendiente6_deuda,monto_cuota1_deuda,monto_cuota2_deuda,monto_cuota3_deuda,monto_cuota4_deuda,monto_cuota5_deuda,monto_cuota6_deuda,cuotas_pactadas1_deuda,cuotas_pactadas2_deuda,cuotas_pactadas3_deuda,cuotas_pactadas4_deuda," _
            & " cuotas_pactadas5_deuda,cuotas_pactadas6_deuda,cuotas_pendientes1_deuda,cuotas_pendientes2_deuda,cuotas_pendientes3_deuda,cuotas_pendientes4_deuda,cuotas_pendientes5_deuda,cuotas_pendientes6_deuda,prepaga_cuota1_deuda,prepaga_cuota2_deuda,prepaga_cuota3_deuda,prepaga_cuota4_deuda,prepaga_cuota5_deuda,prepaga_cuota6_deuda,total_saldo_pendiente_deuda," _
            & " total_deudas,numero_meses_alto_flujo,numero_meses_medio_flujo,numero_meses_bajo_flujo,vta_formal_promedio_mes_alto_flujo,vta_formal_promedio_mes_medio_flujo,vta_formal_promedio_mes_bajo_flujo,vta_informal_promedio_mes_alto_flujo,vta_informal_promedio_mes_medio_flujo,vta_informal_promedio_mes_bajo_flujo,Venta_Total_Promedio_Mes_Alto_flujo," _
            & " Venta_Total_Promedio_Mes_medio_flujo,Venta_Total_Promedio_Mes_bajo_flujo,resultado_operacional_alto_flujo,resultado_operacional_medio_flujo,resultado_operacional_bajo_flujo,capacidad_pago_mes_alto_flujo,capacidad_pago_mes_medio_flujo,capacidad_pago_mes_bajo_flujo,cap_pago_corregida_ajus_mes_alto_flujo,cap_pago_corregida_ajus_mes_medio_flujo," _
            & " cap_pago_corregida_ajus_mes_bajo_flujo,cap_pago_promedio_corregida_ajustada_flujo,monto_maximo_credito_flujo,cuota_credito_flujo,mto_bruto_solicitado_cliente_flujo,resolucion_credito_cuota_flujo,resolucion_credito_monto_flujo,venta_total_promedio_anual,fecha_ingreso,hora_ingreso, impuesto" _
            & " from tbl_micro_metodologia_activo_circulante" _
            & " where '" & txt_rut_cliente_ing & "' = rut_cliente" _
            & " order by n_solicitud desc"
        
            Set rst = cnn.Execute(ssql, , adCmdText)
        
            If rst.EOF Then
        
            MsgBox ("Cliente Sin Ingreso De Evaluacion"), vbCritical
        
            Else
            
            txt_rut_cliente = rst!rut_cliente
            txt_n_solicitud = rst!n_solicitud
            txt_dv = rst!dv
            txt_compra_promedio_mensual = rst!compra_promedio_mensual
            txt_veces_compra_mes = rst!veces_compra_mes
            txt_compra_total_mensual_ctm = rst!compra_total_mensual_ctm
            txt_caja_banco = rst!caja_banco
            txt_materia_primas = rst!materia_prima
            txt_mercaderias = rst!mercaderias
            txt_cuenta_cobrar = rst!cuenta_cobrar
            txt_otros_activos_circulantes = rst!otros_activos_circulantes
            txt_total_activos_circulantes = rst!total_activo_circulante
            txt_producto1 = rst!producto1
            txt_producto2 = rst!producto2
            txt_producto3 = rst!producto3
            txt_producto4 = rst!producto4
            txt_producto5 = rst!producto5
            txt_precio_venta1 = rst!precio_venta1
            txt_precio_venta2 = rst!precio_venta2
            txt_precio_venta3 = rst!precio_venta3
            txt_precio_venta4 = rst!precio_venta4
            txt_precio_venta5 = rst!precio_venta5
            txt_materia_prima1 = rst!materia_prima1
            txt_materia_prima2 = rst!materia_prima2
            txt_materia_prima3 = rst!materia_prima3
            txt_materia_prima4 = rst!materia_prima4
            txt_materia_prima5 = rst!materia_prima5
            txt_mano_obra1 = rst!mano_obra1
            txt_mano_obra2 = rst!mano_obra2
            txt_mano_obra3 = rst!mano_obra3
            txt_mano_obra4 = rst!mano_obra4
            txt_mano_obra5 = rst!mano_obra5
            txt_incidencia_ventas1 = rst!incidencia_ventas1
            txt_incidencia_ventas2 = rst!incidencia_ventas2
            txt_incidencia_ventas3 = rst!incidencia_ventas3
            txt_incidencia_ventas4 = rst!incidencia_ventas4
            txt_incidencia_ventas5 = rst!incidencia_ventas5
            txt_r_cvcmo1 = rst!r_cvcmo1
            txt_r_cvcmo2 = rst!r_cvcmo2
            txt_r_cvcmo3 = rst!r_cvcmo3
            txt_r_cvcmo4 = rst!r_cvcmo4
            txt_r_cvcmo5 = rst!r_cvcmo5
            txt_r_cvsmo1 = rst!r_cvsmo1
            txt_r_cvsmo2 = rst!r_cvsmo2
            txt_r_cvsmo3 = rst!r_cvsmo3
            txt_r_cvsmo4 = rst!r_cvsmo4
            txt_r_cvsmo5 = rst!r_cvsmo5
            txt_r_cvppcmo1 = rst!r_cvppcmo1
            txt_r_cvppcmo2 = rst!r_cvppcmo2
            txt_r_cvppcmo3 = rst!r_cvppcmo3
            txt_r_cvppcmo4 = rst!r_cvppcmo4
            txt_r_cvppcmo5 = rst!r_cvppcmo5
            txt_r_cvppsmo1 = rst!r_cvppsmo1
            txt_r_cvppsmo2 = rst!r_cvppsmo2
            txt_r_cvppsmo3 = rst!r_cvppsmo3
            txt_r_cvppsmo4 = rst!r_cvppsmo4
            txt_r_cvppsmo5 = rst!r_cvppsmo5
            txt_Sub_Total = rst!r_Subtotal_costo_variable
            txt_Sub_Total_x1 = rst!r_Subtotal_x1_costo_variable
            txt_compra_total_mensual = rst!r_compra_total_mensual
            txt_venta_total_alto = rst!r_venta_total_alto
            txt_venta_total_medio = rst!r_venta_total_medio
            txt_venta_total_bajo = rst!r_venta_total_bajo
            txt_compra_total_max_corregida = rst!r_compra_total_max_corregida
            txt_venta_total_mes_alto_corregida = rst!r_venta_total_alto_corregida
            txt_venta_total_mes_medio_corregida = rst!r_venta_total_medio_corregida
            txt_venta_total_mes_bajo_corregida = rst!r_venta_total_bajo_corregida
            txt_arriendo_micro = rst!arriendo_micro
            txt_sueldos = rst!sueldos
            txt_movilizacion = rst!movilizacion
            txt_servicios_basicos = rst!servicios_basicos
            txt_contador = rst!contador
            txt_lubricantes = rst!lubricantes
            txt_neumaticos = rst!neumaticos
            txt_afinamientos = rst!afinamientos
            txt_patentes_seguros = rst!patentes_seguros
            txt_otros_costos_fijos = rst!otros_costos_fijos
            txt_total_costos_fijos = rst!total_costos_fijos
            txt_valor_uf = rst!valor_uf
            txt_n_grupo_familiar = rst!n_grupo_familiar
            txt_arriendo_vivienda = rst!arriendo_vivienda_Gastos_Fam
            txt_gastos_indicado_cliente = rst!gastos_indicado_cliente
            txt_total_gasto_familiar = rst!total_gasto_familiar
            txt_liquidacion_sueldo = rst!liquidacion_sueldo
            txt_jubilacion = rst!jubilacion
            txt_montepio = rst!montepio
            txt_arriendo_vivienda1 = rst!arriendo_vivienda_Otro_Ing
            txt_ingreso_segunda_microempresa = rst!ingreso_segunda_microempresa
            txt_boleta_honorario = rst!boleta_honorario
            txt_total_otros_ingresos = rst!total_otros_ingresos
            txt_acreedor1 = rst!acreedor1_deuda
            txt_acreedor2 = rst!acreedor2_deuda
            txt_acreedor3 = rst!acreedor3_deuda
            txt_acreedor4 = rst!acreedor4_deuda
            txt_acreedor5 = rst!acreedor5_deuda
            txt_acreedor6 = rst!acreedor6_deuda
            txt_tipo_producto1 = rst!tipo_producto1_deuda
            txt_tipo_producto2 = rst!tipo_producto2_deuda
            txt_tipo_producto3 = rst!tipo_producto3_deuda
            txt_tipo_producto4 = rst!tipo_producto4_deuda
            txt_tipo_producto5 = rst!tipo_producto5_deuda
            txt_tipo_producto6 = rst!tipo_producto6_deuda
            txt_saldo_pendiente1 = rst!saldo_pendiente1_deuda
            txt_saldo_pendiente2 = rst!saldo_pendiente2_deuda
            txt_saldo_pendiente3 = rst!saldo_pendiente3_deuda
            txt_saldo_pendiente4 = rst!saldo_pendiente4_deuda
            txt_saldo_pendiente5 = rst!saldo_pendiente5_deuda
            txt_saldo_pendiente6 = rst!saldo_pendiente6_deuda
            txt_monto_cuota1 = rst!monto_cuota1_deuda
            txt_monto_cuota2 = rst!monto_cuota2_deuda
            txt_monto_cuota3 = rst!monto_cuota3_deuda
            txt_monto_cuota4 = rst!monto_cuota4_deuda
            txt_monto_cuota5 = rst!monto_cuota5_deuda
            txt_monto_cuota6 = rst!monto_cuota6_deuda
            txt_cuotas_pactadas1 = rst!cuotas_pactadas1_deuda
            txt_cuotas_pactadas2 = rst!cuotas_pactadas2_deuda
            txt_cuotas_pactadas3 = rst!cuotas_pactadas3_deuda
            txt_cuotas_pactadas4 = rst!cuotas_pactadas4_deuda
            txt_cuotas_pactadas5 = rst!cuotas_pactadas5_deuda
            txt_cuotas_pactadas6 = rst!cuotas_pactadas6_deuda
            txt_cuotas_pendientes1 = rst!cuotas_pendientes1_deuda
            txt_cuotas_pendientes2 = rst!cuotas_pendientes2_deuda
            txt_cuotas_pendientes3 = rst!cuotas_pendientes3_deuda
            txt_cuotas_pendientes4 = rst!cuotas_pendientes4_deuda
            txt_cuotas_pendientes5 = rst!cuotas_pendientes5_deuda
            txt_cuotas_pendientes6 = rst!cuotas_pendientes6_deuda
            cbx_prepaga_deuda1 = rst!prepaga_cuota1_deuda
            cbx_prepaga_deuda2 = rst!prepaga_cuota2_deuda
            cbx_prepaga_deuda3 = rst!prepaga_cuota3_deuda
            cbx_prepaga_deuda4 = rst!prepaga_cuota4_deuda
            cbx_prepaga_deuda5 = rst!prepaga_cuota5_deuda
            cbx_prepaga_deuda6 = rst!prepaga_cuota6_deuda
            txt_total_saldo_pendiente = rst!total_saldo_pendiente_deuda
            txt_total_deudas = rst!total_deudas
            numero_meses_tipo_mes_alto = rst!numero_meses_alto_flujo
            numero_meses_tipo_mes_medio = rst!numero_meses_medio_flujo
            numero_meses_tipo_mes_bajo = rst!numero_meses_bajo_flujo
            txt_vta_formal_promedio_mes_alto = rst!vta_formal_promedio_mes_alto_flujo
            txt_vta_formal_promedio_mes_medio = rst!vta_formal_promedio_mes_medio_flujo
            txt_vta_formal_promedio_mes_bajo = rst!vta_formal_promedio_mes_bajo_flujo
            txt_vta_informal_promedio_mes_alto = rst!vta_informal_promedio_mes_alto_flujo
            txt_vta_informal_promedio_mes_medio = rst!vta_informal_promedio_mes_medio_flujo
            txt_vta_informal_promedio_mes_bajo = rst!vta_informal_promedio_mes_bajo_flujo
            txt_Venta_Total_Promedio_Mes_Alto = rst!Venta_Total_Promedio_Mes_Alto_flujo
            txt_Venta_Total_Promedio_Mes_Medio = rst!Venta_Total_Promedio_Mes_medio_flujo
            txt_Venta_Total_Promedio_Mes_Bajo = rst!Venta_Total_Promedio_Mes_bajo_flujo
            txt_resultado_operacional_mes_alto = rst!resultado_operacional_alto_flujo
            txt_resultado_operacional_mes_medio = rst!resultado_operacional_medio_flujo
            txt_resultado_operacional_mes_bajo = rst!resultado_operacional_bajo_flujo
            txt_capacidad_pago_mes_alto = rst!capacidad_pago_mes_alto_flujo
            txt_capacidad_pago_mes_medio = rst!capacidad_pago_mes_medio_flujo
            txt_capacidad_pago_mes_bajo = rst!capacidad_pago_mes_bajo_flujo
            txt_capacidad_pago_corregida_ajustada_mes_alto = rst!cap_pago_corregida_ajus_mes_alto_flujo
            txt_capacidad_pago_corregida_ajustada_mes_medio = rst!cap_pago_corregida_ajus_mes_medio_flujo
            txt_capacidad_pago_corregida_ajustada_mes_bajo = rst!cap_pago_corregida_ajus_mes_bajo_flujo
            txt_capacidad_pago_promedio_corregida_ajustada = rst!cap_pago_promedio_corregida_ajustada_flujo
            txt_monto_maximo_credito = rst!monto_maximo_credito_flujo
            txt_cuota_credito = rst!cuota_credito_flujo
            txt_mto_bruto_sol_cliente = rst!mto_bruto_solicitado_cliente_flujo
            txt_resolucion_credito_por_cuota = rst!resolucion_credito_cuota_flujo
            txt_aprobacion = rst!resolucion_credito_monto_flujo
            txt_fecha_actual = rst!FECHA_INGRESO
            txt_hora_actual = rst!HORA_INGRESO
            txt_costo_fijo_mes_alto = rst!total_costos_fijos
            txt_costo_fijo_mes_medio = rst!total_costos_fijos
            txt_costo_fijo_mes_bajo = rst!total_costos_fijos
            txt_otros_ingresos_mes_alto = rst!total_otros_ingresos
            txt_otros_ingresos_mes_medio = rst!total_otros_ingresos
            txt_otros_ingresos_mes_bajo = rst!total_otros_ingresos
            txt_Deudas_flujo_caja_mes_alto = rst!total_deudas
            txt_Deudas_flujo_caja_mes_medio = rst!total_deudas
            txt_Deudas_flujo_caja_mes_bajo = rst!total_deudas
            txt_gastos_familiares_mes_alto = rst!total_gasto_familiar
            txt_gastos_familiares_mes_medio = rst!total_gasto_familiar
            txt_gastos_familiares_mes_bajo = rst!total_gasto_familiar
            txt_impuesto = rst!impuesto
            
            ''''CALCULO COSTO VARIABLE FLUJO FINAL
            
            txt_costo_variable_mes_alto = Int(txt_Venta_Total_Promedio_Mes_Alto * txt_Sub_Total_x1)
            txt_costo_variable_mes_medio = Int(txt_Venta_Total_Promedio_Mes_Medio * txt_Sub_Total_x1)
            txt_costo_variable_mes_bajo = Int(txt_Venta_Total_Promedio_Mes_Bajo * txt_Sub_Total_x1)
            
            txt_venta_total_promedio_anual = rst!venta_total_promedio_anual
            
        End If
End Sub

Private Sub CommandButton1_Click()
Visualizador_Met_AC.PrintForm
End Sub

Private Sub cmd_volver_menu_inicial_1_Click()
cmd_volver_menu_inicial_Click
End Sub

Private Sub cmd_volver_menu_inicial_Click()
Visualizador_Met_AC.Hide
Visualizador_Inicial.txt_rut_cliente = Empty
Visualizador_Inicial.Show
End Sub


Private Sub CommandButton2_Click()
Visualizador_Met_AC.PrintForm
End Sub

Private Sub CommandButton3_Click()
Visualizador_Met_AC.PrintForm
End Sub

Private Sub CommandButton4_Click()
Visualizador_Met_AC.PrintForm
End Sub

Private Sub TextBox73_Change()

End Sub

Private Sub TextBox76_Change()

End Sub

Private Sub txt_afinamientos_Change()
txt_afinamientos = Format(txt_afinamientos, "##,##")
End Sub

Private Sub txt_arriendo_micro_Change()
txt_arriendo_micro = Format(txt_arriendo_micro, "##,##")
End Sub

Private Sub txt_arriendo_vivienda_Change()
txt_arriendo_vivienda = Format(txt_arriendo_vivienda, "##,##")
End Sub

Private Sub txt_arriendo_vivienda1_Change()
txt_arriendo_vivienda1 = Format(txt_arriendo_vivienda1, "##,##")
End Sub

Private Sub txt_boleta_honorario_Change()
txt_boleta_honorario = Format(txt_boleta_honorario, "##,##")
End Sub

Private Sub txt_caja_banco_Change()
txt_caja_banco = Format(txt_caja_banco, "##,##")
End Sub

Private Sub txt_capacidad_pago_corregida_ajustada_mes_alto_Change()
txt_capacidad_pago_corregida_ajustada_mes_alto = Format(txt_capacidad_pago_corregida_ajustada_mes_alto, "##,##")
End Sub

Private Sub txt_capacidad_pago_corregida_ajustada_mes_bajo_Change()
txt_capacidad_pago_corregida_ajustada_mes_bajo = Format(txt_capacidad_pago_corregida_ajustada_mes_bajo, "##,##")
End Sub

Private Sub txt_capacidad_pago_corregida_ajustada_mes_medio_Change()
txt_capacidad_pago_corregida_ajustada_mes_medio = Format(txt_capacidad_pago_corregida_ajustada_mes_medio, "##,##")
End Sub

Private Sub txt_capacidad_pago_mes_alto_Change()
txt_capacidad_pago_mes_alto = Format(txt_capacidad_pago_mes_alto, "##,##")
End Sub

Private Sub txt_capacidad_pago_mes_bajo_Change()
txt_capacidad_pago_mes_bajo = Format(txt_capacidad_pago_mes_bajo, "##,##")
End Sub

Private Sub txt_capacidad_pago_mes_medio_Change()
txt_capacidad_pago_mes_medio = Format(txt_capacidad_pago_mes_medio, "##,##")
End Sub

Private Sub txt_capacidad_pago_promedio_corregida_ajustada_Change()
txt_capacidad_pago_promedio_corregida_ajustada = Format(txt_capacidad_pago_promedio_corregida_ajustada, "##,##")
End Sub

Private Sub txt_compra_promedio_mensual_Change()
txt_compra_promedio_mensual = Format(txt_compra_promedio_mensual, "##,##")
End Sub

Private Sub txt_compra_total_max_corregida_Change()
txt_compra_total_max_corregida = Format(txt_compra_total_max_corregida, "##,##")
End Sub

Private Sub txt_compra_total_mensual_Change()
txt_compra_total_mensual = Format(txt_compra_total_mensual, "##,##")
End Sub

Private Sub txt_compra_total_mensual_ctm_Change()
txt_compra_total_mensual_ctm = Format(txt_compra_total_mensual_ctm, "##,##")
End Sub

Private Sub txt_contador_Change()
txt_contador = Format(txt_contador, "##,##")
End Sub

Private Sub txt_costo_fijo_mes_alto_Change()
txt_costo_fijo_mes_alto = Format(txt_costo_fijo_mes_alto, "##,##")
End Sub

Private Sub txt_costo_fijo_mes_bajo_Change()
txt_costo_fijo_mes_bajo = Format(txt_costo_fijo_mes_bajo, "##,##")
End Sub

Private Sub txt_costo_fijo_mes_medio_Change()
txt_costo_fijo_mes_medio = Format(txt_costo_fijo_mes_medio, "##,##")
End Sub

Private Sub txt_costo_variable_mes_alto_Change()
txt_costo_variable_mes_alto = Format(txt_costo_variable_mes_alto, "##,##")
End Sub

Private Sub txt_costo_variable_mes_bajo_Change()
txt_costo_variable_mes_bajo = Format(txt_costo_variable_mes_bajo, "##,##")
End Sub

Private Sub txt_costo_variable_mes_medio_Change()
txt_costo_variable_mes_medio = Format(txt_costo_variable_mes_medio, "##,##")
End Sub

Private Sub txt_cuenta_cobrar_Change()
txt_cuenta_cobrar = Format(txt_cuenta_cobrar, "##,##")
End Sub

Private Sub txt_cuota_credito_Change()
txt_cuota_credito = Format(txt_cuota_credito, "##,##")
End Sub

Private Sub txt_Deudas_flujo_caja_mes_alto_Change()
txt_Deudas_flujo_caja_mes_alto = Format(txt_Deudas_flujo_caja_mes_alto, "##,##")
End Sub

Private Sub txt_Deudas_flujo_caja_mes_bajo_Change()
txt_Deudas_flujo_caja_mes_bajo = Format(txt_Deudas_flujo_caja_mes_bajo, "##,##")
End Sub

Private Sub txt_Deudas_flujo_caja_mes_medio_Change()
txt_Deudas_flujo_caja_mes_medio = Format(txt_Deudas_flujo_caja_mes_medio, "##,##")
End Sub



Private Sub txt_dv_ing_Change()
        
        Call conectarBD

        ssql = "SELECT RUT_CLIENTE,N_SOLICITUD,DV,compra_promedio_mensual,veces_compra_mes,compra_total_mensual_ctm,caja_banco,materia_prima,mercaderias,cuenta_cobrar,otros_activos_circulantes,total_activo_circulante,producto1,producto2,producto3,producto4,producto5,precio_venta1,precio_venta2,precio_venta3,precio_venta4," _
            & " precio_venta5,materia_prima1,materia_prima2,materia_prima3,materia_prima4,materia_prima5,mano_obra1,mano_obra2,mano_obra3,mano_obra4,mano_obra5,incidencia_ventas1,incidencia_ventas2,incidencia_ventas3,incidencia_ventas4,incidencia_ventas5,r_cvcmo1,r_cvcmo2,r_cvcmo3,r_cvcmo4,r_cvcmo5,r_cvsmo1,r_cvsmo2,r_cvsmo3,r_cvsmo4," _
            & " r_cvsmo5,r_cvppcmo1,r_cvppcmo2,r_cvppcmo3,r_cvppcmo4,r_cvppcmo5,r_cvppsmo1,r_cvppsmo2,r_cvppsmo3,r_cvppsmo4,r_cvppsmo5,r_Subtotal_costo_variable,r_Subtotal_x1_costo_variable,r_compra_total_mensual,r_venta_total_alto,r_venta_total_medio,r_venta_total_bajo,r_compra_total_max_corregida,r_venta_total_alto_corregida," _
            & " r_venta_total_medio_corregida,r_venta_total_bajo_corregida,arriendo_micro,sueldos,movilizacion,servicios_basicos,contador,lubricantes,neumaticos,afinamientos,patentes_seguros,otros_costos_fijos,total_costos_fijos,valor_uf,n_grupo_familiar,arriendo_vivienda_Gastos_Fam,gastos_indicado_cliente,total_gasto_familiar,liquidacion_sueldo," _
            & " jubilacion,montepio,arriendo_vivienda_Otro_Ing,ingreso_segunda_microempresa,boleta_honorario,total_otros_ingresos,acreedor1_deuda,acreedor2_deuda,acreedor3_deuda,acreedor4_deuda,acreedor5_deuda,acreedor6_deuda,tipo_producto1_deuda,tipo_producto2_deuda,tipo_producto3_deuda,tipo_producto4_deuda,tipo_producto5_deuda,tipo_producto6_deuda," _
            & " saldo_pendiente1_deuda,saldo_pendiente2_deuda,saldo_pendiente3_deuda,saldo_pendiente4_deuda,saldo_pendiente5_deuda,saldo_pendiente6_deuda,monto_cuota1_deuda,monto_cuota2_deuda,monto_cuota3_deuda,monto_cuota4_deuda,monto_cuota5_deuda,monto_cuota6_deuda,cuotas_pactadas1_deuda,cuotas_pactadas2_deuda,cuotas_pactadas3_deuda,cuotas_pactadas4_deuda," _
            & " cuotas_pactadas5_deuda,cuotas_pactadas6_deuda,cuotas_pendientes1_deuda,cuotas_pendientes2_deuda,cuotas_pendientes3_deuda,cuotas_pendientes4_deuda,cuotas_pendientes5_deuda,cuotas_pendientes6_deuda,prepaga_cuota1_deuda,prepaga_cuota2_deuda,prepaga_cuota3_deuda,prepaga_cuota4_deuda,prepaga_cuota5_deuda,prepaga_cuota6_deuda,total_saldo_pendiente_deuda," _
            & " total_deudas,numero_meses_alto_flujo,numero_meses_medio_flujo,numero_meses_bajo_flujo,vta_formal_promedio_mes_alto_flujo,vta_formal_promedio_mes_medio_flujo,vta_formal_promedio_mes_bajo_flujo,vta_informal_promedio_mes_alto_flujo,vta_informal_promedio_mes_medio_flujo,vta_informal_promedio_mes_bajo_flujo,Venta_Total_Promedio_Mes_Alto_flujo," _
            & " Venta_Total_Promedio_Mes_medio_flujo,Venta_Total_Promedio_Mes_bajo_flujo,resultado_operacional_alto_flujo,resultado_operacional_medio_flujo,resultado_operacional_bajo_flujo,capacidad_pago_mes_alto_flujo,capacidad_pago_mes_medio_flujo,capacidad_pago_mes_bajo_flujo,cap_pago_corregida_ajus_mes_alto_flujo,cap_pago_corregida_ajus_mes_medio_flujo," _
            & " cap_pago_corregida_ajus_mes_bajo_flujo,cap_pago_promedio_corregida_ajustada_flujo,monto_maximo_credito_flujo,cuota_credito_flujo,mto_bruto_solicitado_cliente_flujo,resolucion_credito_cuota_flujo,resolucion_credito_monto_flujo,isnull(venta_total_promedio_anual,'SD') as venta_total_promedio_anual, fecha_ingreso,hora_ingreso,impuesto,ISNULL(tipo_credito_deuda1,'SD')tipo_credito_deuda1,ISNULL(tipo_credito_deuda2,'SD')tipo_credito_deuda2,ISNULL(tipo_credito_deuda3,'SD')tipo_credito_deuda3,ISNULL(tipo_credito_deuda4,'SD')tipo_credito_deuda4,ISNULL(tipo_credito_deuda5,'SD')tipo_credito_deuda5,ISNULL(tipo_credito_deuda6,'SD')tipo_credito_deuda6," _
            & " ISNULL(total_saldo_pendiente_consumo,'0')total_saldo_pendiente_consumo,ISNULL(total_deudas_consumo,'0')total_deudas_consumo,ISNULL(total_saldo_pendiente_comercial,'0')total_saldo_pendiente_comercial,ISNULL(total_deudas_comercial,'0')total_deudas_comercial,ISNULL(saldo_deuda_con_prepago_consumo,'0')saldo_deuda_con_prepago_consumo,ISNULL(saldo_deuda_con_prepago_comercial,'0')saldo_deuda_con_prepago_comercial,ISNULL(mto_cuota_con_prepago_consumo,'0')mto_cuota_con_prepago_consumo,ISNULL(mto_cuota_con_prepago_comercial,'0')mto_cuota_con_prepago_comercial,ISNULL(saldo_deuda_sin_prepago_consumo,'0')saldo_deuda_sin_prepago_consumo,ISNULL(saldo_deuda_sin_prepago_comercial,'0')saldo_deuda_sin_prepago_comercial,ISNULL(mto_cuota_sin_prepago_comercial,'0')mto_cuota_sin_prepago_comercial,ISNULL(mto_cuota_sin_prepago_consumo,'0')mto_cuota_sin_prepago_consumo" _
            & " from tbl_micro_metodologia_activo_circulante" _
            & " where '" & txt_rut_cliente_ing & "' = rut_cliente" _
            & " and '" & txt_n_solicitud & "' = n_solicitud" _
            & " order by n_solicitud desc"
        
            Set rst = cnn.Execute(ssql, , adCmdText)
        
            If rst.EOF Then
        
            MsgBox ("Cliente Sin Ingreso De Evaluacion"), vbCritical
        
            Else
            
            txt_rut_cliente = rst!rut_cliente
            txt_n_solicitud = rst!n_solicitud
            txt_dv = rst!dv
            txt_compra_promedio_mensual = rst!compra_promedio_mensual
            txt_veces_compra_mes = rst!veces_compra_mes
            txt_compra_total_mensual_ctm = rst!compra_total_mensual_ctm
            txt_caja_banco = rst!caja_banco
            txt_materia_primas = rst!materia_prima
            txt_mercaderias = rst!mercaderias
            txt_cuenta_cobrar = rst!cuenta_cobrar
            txt_otros_activos_circulantes = rst!otros_activos_circulantes
            txt_total_activos_circulantes = rst!total_activo_circulante
            txt_producto1 = rst!producto1
            txt_producto2 = rst!producto2
            txt_producto3 = rst!producto3
            txt_producto4 = rst!producto4
            txt_producto5 = rst!producto5
            txt_precio_venta1 = rst!precio_venta1
            txt_precio_venta2 = rst!precio_venta2
            txt_precio_venta3 = rst!precio_venta3
            txt_precio_venta4 = rst!precio_venta4
            txt_precio_venta5 = rst!precio_venta5
            txt_materia_prima1 = rst!materia_prima1
            txt_materia_prima2 = rst!materia_prima2
            txt_materia_prima3 = rst!materia_prima3
            txt_materia_prima4 = rst!materia_prima4
            txt_materia_prima5 = rst!materia_prima5
            txt_mano_obra1 = rst!mano_obra1
            txt_mano_obra2 = rst!mano_obra2
            txt_mano_obra3 = rst!mano_obra3
            txt_mano_obra4 = rst!mano_obra4
            txt_mano_obra5 = rst!mano_obra5
            txt_incidencia_ventas1 = rst!incidencia_ventas1
            txt_incidencia_ventas2 = rst!incidencia_ventas2
            txt_incidencia_ventas3 = rst!incidencia_ventas3
            txt_incidencia_ventas4 = rst!incidencia_ventas4
            txt_incidencia_ventas5 = rst!incidencia_ventas5
            txt_r_cvcmo1 = rst!r_cvcmo1
            txt_r_cvcmo2 = rst!r_cvcmo2
            txt_r_cvcmo3 = rst!r_cvcmo3
            txt_r_cvcmo4 = rst!r_cvcmo4
            txt_r_cvcmo5 = rst!r_cvcmo5
            txt_r_cvsmo1 = rst!r_cvsmo1
            txt_r_cvsmo2 = rst!r_cvsmo2
            txt_r_cvsmo3 = rst!r_cvsmo3
            txt_r_cvsmo4 = rst!r_cvsmo4
            txt_r_cvsmo5 = rst!r_cvsmo5
            txt_r_cvppcmo1 = rst!r_cvppcmo1
            txt_r_cvppcmo2 = rst!r_cvppcmo2
            txt_r_cvppcmo3 = rst!r_cvppcmo3
            txt_r_cvppcmo4 = rst!r_cvppcmo4
            txt_r_cvppcmo5 = rst!r_cvppcmo5
            txt_r_cvppsmo1 = rst!r_cvppsmo1
            txt_r_cvppsmo2 = rst!r_cvppsmo2
            txt_r_cvppsmo3 = rst!r_cvppsmo3
            txt_r_cvppsmo4 = rst!r_cvppsmo4
            txt_r_cvppsmo5 = rst!r_cvppsmo5
            txt_Sub_Total = rst!r_Subtotal_costo_variable
            txt_Sub_Total_x1 = rst!r_Subtotal_x1_costo_variable
            txt_compra_total_mensual = rst!r_compra_total_mensual
            txt_venta_total_alto = rst!r_venta_total_alto
            txt_venta_total_medio = rst!r_venta_total_medio
            txt_venta_total_bajo = rst!r_venta_total_bajo
            txt_compra_total_max_corregida = rst!r_compra_total_max_corregida
            txt_venta_total_mes_alto_corregida = rst!r_venta_total_alto_corregida
            txt_venta_total_mes_medio_corregida = rst!r_venta_total_medio_corregida
            txt_venta_total_mes_bajo_corregida = rst!r_venta_total_bajo_corregida
            txt_arriendo_micro = rst!arriendo_micro
            txt_sueldos = rst!sueldos
            txt_movilizacion = rst!movilizacion
            txt_servicios_basicos = rst!servicios_basicos
            txt_contador = rst!contador
            txt_lubricantes = rst!lubricantes
            txt_neumaticos = rst!neumaticos
            txt_afinamientos = rst!afinamientos
            txt_patentes_seguros = rst!patentes_seguros
            txt_otros_costos_fijos = rst!otros_costos_fijos
            txt_total_costos_fijos = rst!total_costos_fijos
            txt_valor_uf = rst!valor_uf
            txt_n_grupo_familiar = rst!n_grupo_familiar
            txt_arriendo_vivienda = rst!arriendo_vivienda_Gastos_Fam
            txt_gastos_indicado_cliente = rst!gastos_indicado_cliente
            txt_total_gasto_familiar = rst!total_gasto_familiar
            txt_liquidacion_sueldo = rst!liquidacion_sueldo
            txt_jubilacion = rst!jubilacion
            txt_montepio = rst!montepio
            txt_arriendo_vivienda1 = rst!arriendo_vivienda_Otro_Ing
            txt_ingreso_segunda_microempresa = rst!ingreso_segunda_microempresa
            txt_boleta_honorario = rst!boleta_honorario
            txt_total_otros_ingresos = rst!total_otros_ingresos
            txt_acreedor1 = rst!acreedor1_deuda
            txt_acreedor2 = rst!acreedor2_deuda
            txt_acreedor3 = rst!acreedor3_deuda
            txt_acreedor4 = rst!acreedor4_deuda
            txt_acreedor5 = rst!acreedor5_deuda
            txt_acreedor6 = rst!acreedor6_deuda
            txt_tipo_producto1 = rst!tipo_producto1_deuda
            txt_tipo_producto2 = rst!tipo_producto2_deuda
            txt_tipo_producto3 = rst!tipo_producto3_deuda
            txt_tipo_producto4 = rst!tipo_producto4_deuda
            txt_tipo_producto5 = rst!tipo_producto5_deuda
            txt_tipo_producto6 = rst!tipo_producto6_deuda
            txt_saldo_pendiente1 = rst!saldo_pendiente1_deuda
            txt_saldo_pendiente2 = rst!saldo_pendiente2_deuda
            txt_saldo_pendiente3 = rst!saldo_pendiente3_deuda
            txt_saldo_pendiente4 = rst!saldo_pendiente4_deuda
            txt_saldo_pendiente5 = rst!saldo_pendiente5_deuda
            txt_saldo_pendiente6 = rst!saldo_pendiente6_deuda
            txt_monto_cuota1 = rst!monto_cuota1_deuda
            txt_monto_cuota2 = rst!monto_cuota2_deuda
            txt_monto_cuota3 = rst!monto_cuota3_deuda
            txt_monto_cuota4 = rst!monto_cuota4_deuda
            txt_monto_cuota5 = rst!monto_cuota5_deuda
            txt_monto_cuota6 = rst!monto_cuota6_deuda
            txt_cuotas_pactadas1 = rst!cuotas_pactadas1_deuda
            txt_cuotas_pactadas2 = rst!cuotas_pactadas2_deuda
            txt_cuotas_pactadas3 = rst!cuotas_pactadas3_deuda
            txt_cuotas_pactadas4 = rst!cuotas_pactadas4_deuda
            txt_cuotas_pactadas5 = rst!cuotas_pactadas5_deuda
            txt_cuotas_pactadas6 = rst!cuotas_pactadas6_deuda
            txt_cuotas_pendientes1 = rst!cuotas_pendientes1_deuda
            txt_cuotas_pendientes2 = rst!cuotas_pendientes2_deuda
            txt_cuotas_pendientes3 = rst!cuotas_pendientes3_deuda
            txt_cuotas_pendientes4 = rst!cuotas_pendientes4_deuda
            txt_cuotas_pendientes5 = rst!cuotas_pendientes5_deuda
            txt_cuotas_pendientes6 = rst!cuotas_pendientes6_deuda
            cbx_prepaga_deuda1 = rst!prepaga_cuota1_deuda
            cbx_prepaga_deuda2 = rst!prepaga_cuota2_deuda
            cbx_prepaga_deuda3 = rst!prepaga_cuota3_deuda
            cbx_prepaga_deuda4 = rst!prepaga_cuota4_deuda
            cbx_prepaga_deuda5 = rst!prepaga_cuota5_deuda
            cbx_prepaga_deuda6 = rst!prepaga_cuota6_deuda
            txt_total_saldo_pendiente = rst!total_saldo_pendiente_deuda
            txt_total_deudas = rst!total_deudas
            numero_meses_tipo_mes_alto = rst!numero_meses_alto_flujo
            numero_meses_tipo_mes_medio = rst!numero_meses_medio_flujo
            numero_meses_tipo_mes_bajo = rst!numero_meses_bajo_flujo
            txt_vta_formal_promedio_mes_alto = rst!vta_formal_promedio_mes_alto_flujo
            txt_vta_formal_promedio_mes_medio = rst!vta_formal_promedio_mes_medio_flujo
            txt_vta_formal_promedio_mes_bajo = rst!vta_formal_promedio_mes_bajo_flujo
            txt_vta_informal_promedio_mes_alto = rst!vta_informal_promedio_mes_alto_flujo
            txt_vta_informal_promedio_mes_medio = rst!vta_informal_promedio_mes_medio_flujo
            txt_vta_informal_promedio_mes_bajo = rst!vta_informal_promedio_mes_bajo_flujo
            txt_Venta_Total_Promedio_Mes_Alto = rst!Venta_Total_Promedio_Mes_Alto_flujo
            txt_Venta_Total_Promedio_Mes_Medio = rst!Venta_Total_Promedio_Mes_medio_flujo
            txt_Venta_Total_Promedio_Mes_Bajo = rst!Venta_Total_Promedio_Mes_bajo_flujo
            txt_resultado_operacional_mes_alto = rst!resultado_operacional_alto_flujo
            txt_resultado_operacional_mes_medio = rst!resultado_operacional_medio_flujo
            txt_resultado_operacional_mes_bajo = rst!resultado_operacional_bajo_flujo
            txt_capacidad_pago_mes_alto = rst!capacidad_pago_mes_alto_flujo
            txt_capacidad_pago_mes_medio = rst!capacidad_pago_mes_medio_flujo
            txt_capacidad_pago_mes_bajo = rst!capacidad_pago_mes_bajo_flujo
            txt_capacidad_pago_corregida_ajustada_mes_alto = rst!cap_pago_corregida_ajus_mes_alto_flujo
            txt_capacidad_pago_corregida_ajustada_mes_medio = rst!cap_pago_corregida_ajus_mes_medio_flujo
            txt_capacidad_pago_corregida_ajustada_mes_bajo = rst!cap_pago_corregida_ajus_mes_bajo_flujo
            txt_capacidad_pago_promedio_corregida_ajustada = rst!cap_pago_promedio_corregida_ajustada_flujo
            txt_monto_maximo_credito = rst!monto_maximo_credito_flujo
            txt_cuota_credito = rst!cuota_credito_flujo
            txt_mto_bruto_sol_cliente = rst!mto_bruto_solicitado_cliente_flujo
            txt_resolucion_credito_por_cuota = rst!resolucion_credito_cuota_flujo
            txt_aprobacion = rst!resolucion_credito_monto_flujo
            txt_fecha_actual = rst!FECHA_INGRESO
            txt_hora_actual = rst!HORA_INGRESO
            txt_costo_fijo_mes_alto = rst!total_costos_fijos
            txt_costo_fijo_mes_medio = rst!total_costos_fijos
            txt_costo_fijo_mes_bajo = rst!total_costos_fijos
            txt_otros_ingresos_mes_alto = rst!total_otros_ingresos
            txt_otros_ingresos_mes_medio = rst!total_otros_ingresos
            txt_otros_ingresos_mes_bajo = rst!total_otros_ingresos
            txt_Deudas_flujo_caja_mes_alto = rst!total_deudas
            txt_Deudas_flujo_caja_mes_medio = rst!total_deudas
            txt_Deudas_flujo_caja_mes_bajo = rst!total_deudas
            txt_gastos_familiares_mes_alto = rst!total_gasto_familiar
            txt_gastos_familiares_mes_medio = rst!total_gasto_familiar
            txt_gastos_familiares_mes_bajo = rst!total_gasto_familiar
            txt_impuesto = rst!impuesto
            ''''CALCULO COSTO VARIABLE FLUJO FINAL
            
            txt_costo_variable_mes_alto = Int(txt_Venta_Total_Promedio_Mes_Alto * txt_Sub_Total_x1)
            txt_costo_variable_mes_medio = Int(txt_Venta_Total_Promedio_Mes_Medio * txt_Sub_Total_x1)
            txt_costo_variable_mes_bajo = Int(txt_Venta_Total_Promedio_Mes_Bajo * txt_Sub_Total_x1)
            
            txt_venta_total_promedio_anual = rst!venta_total_promedio_anual
            
            '''valores de CONSUMO
            cbx_tipo_credito_deuda1 = rst!tipo_credito_deuda1
            cbx_tipo_credito_deuda2 = rst!tipo_credito_deuda2
            cbx_tipo_credito_deuda3 = rst!tipo_credito_deuda3
            cbx_tipo_credito_deuda4 = rst!tipo_credito_deuda4
            cbx_tipo_credito_deuda5 = rst!tipo_credito_deuda5
            cbx_tipo_credito_deuda6 = rst!tipo_credito_deuda6
            txt_total_saldo_pendiente_consumo = rst!total_saldo_pendiente_consumo
            txt_total_deudas_consumo = rst!total_deudas_consumo
            txt_total_saldo_pendiente_comercial = rst!total_saldo_pendiente_comercial
            txt_total_deudas_comercial = rst!total_deudas_comercial
            txt_saldo_deuda_con_prepago_consumo = rst!saldo_deuda_con_prepago_consumo
            txt_saldo_deuda_con_prepago_comercial = rst!saldo_deuda_con_prepago_comercial
            txt_mto_cuota_con_prepago_consumo = rst!mto_cuota_con_prepago_consumo
            txt_mto_cuota_con_prepago_comercial = rst!mto_cuota_con_prepago_comercial
            txt_saldo_deuda_sin_prepago_consumo = rst!saldo_deuda_sin_prepago_consumo
            txt_saldo_deuda_sin_prepago_comercial = rst!saldo_deuda_sin_prepago_comercial
            txt_mto_cuota_sin_prepago_comercial = rst!mto_cuota_sin_prepago_comercial
            txt_mto_cuota_sin_prepago_consumo = rst!mto_cuota_sin_prepago_consumo
            
        End If
End Sub


Private Sub txt_gastos_familiares_mes_alto_Change()
txt_gastos_familiares_mes_alto = Format(txt_gastos_familiares_mes_alto, "##,##")
End Sub

Private Sub txt_gastos_familiares_mes_bajo_Change()
txt_gastos_familiares_mes_bajo = Format(txt_gastos_familiares_mes_bajo, "##,##")
End Sub

Private Sub txt_gastos_familiares_mes_medio_Change()
txt_gastos_familiares_mes_medio = Format(txt_gastos_familiares_mes_medio, "##,##")
End Sub

Private Sub txt_gastos_indicado_cliente_Change()
txt_gastos_indicado_cliente = Format(txt_gastos_indicado_cliente, "##,##")
End Sub

Private Sub txt_impuesto_Change()
txt_impuesto = Format(txt_impuesto, "##,##")
End Sub

Private Sub txt_ingreso_segunda_microempresa_Change()
txt_ingreso_segunda_microempresa = Format(txt_ingreso_segunda_microempresa, "##,##")
End Sub

Private Sub txt_jubilacion_Change()
txt_jubilacion = Format(txt_jubilacion, "##,##")
End Sub

Private Sub txt_liquidacion_sueldo_Change()
txt_liquidacion_sueldo = Format(txt_liquidacion_sueldo, "##,##")
End Sub

Private Sub txt_lubricantes_Change()
txt_lubricantes = Format(txt_lubricantes, "##,##")
End Sub

Private Sub txt_mano_obra1_Change()
txt_mano_obra1 = Format(txt_mano_obra1, "##,##")
End Sub
Private Sub txt_mano_obra2_Change()
txt_mano_obra2 = Format(txt_mano_obra2, "##,##")
End Sub
Private Sub txt_mano_obra5_Change()
txt_mano_obra5 = Format(txt_mano_obra5, "##,##")
End Sub
Private Sub txt_mano_obra3_Change()
txt_mano_obra3 = Format(txt_mano_obra3, "##,##")
End Sub
Private Sub txt_mano_obra4Change()
txt_mano_obra4 = Format(txt_mano_obra4, "##,##")
End Sub

Private Sub txt_materia_prima1_Change()
txt_materia_prima1 = Format(txt_materia_prima1, "##,##")
End Sub
Private Sub txt_materia_prima2_Change()
txt_materia_prima2 = Format(txt_materia_prima2, "##,##")
End Sub
Private Sub txt_materia_prima3_Change()
txt_materia_prima3 = Format(txt_materia_prima3, "##,##")
End Sub
Private Sub txt_materia_prima4_Change()
txt_materia_prima4 = Format(txt_materia_prima4, "##,##")
End Sub
Private Sub txt_materia_prima5_Change()
txt_materia_prima5 = Format(txt_materia_prima5, "##,##")
End Sub

Private Sub txt_materia_primas_Change()
txt_materia_primas = Format(txt_materia_primas, "##,##")
End Sub

Private Sub txt_mercaderias_Change()
txt_mercaderias = Format(txt_mercaderias, "##,##")
End Sub

Private Sub txt_montepio_Change()
txt_montepio = Format(txt_montepio, "##,##")
End Sub

Private Sub txt_monto_cuota1_Change()
txt_monto_cuota1 = Format(txt_monto_cuota1, "##,##")
End Sub
Private Sub txt_monto_cuota2_Change()
txt_monto_cuota2 = Format(txt_monto_cuota2, "##,##")
End Sub
Private Sub txt_monto_cuota3_Change()
txt_monto_cuota3 = Format(txt_monto_cuota3, "##,##")
End Sub
Private Sub txt_monto_cuota4_Change()
txt_monto_cuota4 = Format(txt_monto_cuota4, "##,##")
End Sub
Private Sub txt_monto_cuota5_Change()
txt_monto_cuota5 = Format(txt_monto_cuota5, "##,##")
End Sub
Private Sub txt_monto_cuota6_Change()
txt_monto_cuota6 = Format(txt_monto_cuota6, "##,##")
End Sub

Private Sub txt_monto_maximo_credito_Change()
txt_monto_maximo_credito = Format(txt_monto_maximo_credito, "##,##")
End Sub

Private Sub txt_movilizacion_Change()
txt_movilizacion = Format(txt_movilizacion, "##,##")
End Sub



Private Sub txt_mto_bruto_sol_cliente_Change()
txt_mto_bruto_sol_cliente = Format(txt_mto_bruto_sol_cliente, "##,##")
End Sub

Private Sub txt_n_grupo_familiar_Change()

End Sub

Private Sub txt_neumaticos_Change()
txt_neumaticos = Format(txt_neumaticos, "##,##")
End Sub

Private Sub txt_otros_activos_circulantes_Change()
txt_otros_activos_circulantes = Format(txt_otros_activos_circulantes, "##,##")
End Sub

Private Sub txt_otros_costos_fijos_Change()
txt_otros_costos_fijos = Format(txt_otros_costos_fijos, "##,##")
End Sub

Private Sub txt_otros_ingresos_mes_alto_Change()
txt_otros_ingresos_mes_alto = Format(txt_otros_ingresos_mes_alto, "##,##")
End Sub

Private Sub txt_otros_ingresos_mes_bajo_Change()
txt_otros_ingresos_mes_bajo = Format(txt_otros_ingresos_mes_bajo, "##,##")
End Sub

Private Sub txt_otros_ingresos_mes_medio_Change()
txt_otros_ingresos_mes_medio = Format(txt_otros_ingresos_mes_medio, "##,##")
End Sub

Private Sub txt_patentes_seguros_Change()
txt_patentes_seguros = Format(txt_patentes_seguros, "##,##")
End Sub

Private Sub txt_precio_venta1_Change()
txt_precio_venta1 = Format(txt_precio_venta1, "##,##")
End Sub

Private Sub txt_precio_venta2_Change()
txt_precio_venta2 = Format(txt_precio_venta2, "##,##")
End Sub

Private Sub txt_precio_venta3_Change()
txt_precio_venta3 = Format(txt_precio_venta3, "##,##")
End Sub

Private Sub txt_precio_venta4_Change()
txt_precio_venta4 = Format(txt_precio_venta4, "##,##")
End Sub

Private Sub txt_precio_venta5_Change()
txt_precio_venta5 = Format(txt_precio_venta5, "##,##")
End Sub

Private Sub txt_resolucion_credito_por_cuota_Change()

End Sub

Private Sub txt_resultado_operacional_mes_alto_Change()
txt_resultado_operacional_mes_alto = Format(txt_resultado_operacional_mes_alto, "##,##")
End Sub

Private Sub txt_resultado_operacional_mes_bajo_Change()
txt_resultado_operacional_mes_bajo = Format(txt_resultado_operacional_mes_bajo, "##,##")
End Sub

Private Sub txt_resultado_operacional_mes_medio_Change()
txt_resultado_operacional_mes_medio = Format(txt_resultado_operacional_mes_medio, "##,##")
End Sub

Private Sub txt_rut_cliente_ing_Change()

End Sub

Private Sub txt_saldo_pendiente1_Change()
txt_saldo_pendiente1 = Format(txt_saldo_pendiente1, "##,##")
End Sub
Private Sub txt_saldo_pendiente2_Change()
txt_saldo_pendiente2 = Format(txt_saldo_pendiente2, "##,##")
End Sub
Private Sub txt_saldo_pendiente3_Change()
txt_saldo_pendiente3 = Format(txt_saldo_pendiente3, "##,##")
End Sub
Private Sub txt_saldo_pendiente4_Change()
txt_saldo_pendiente4 = Format(txt_saldo_pendiente4, "##,##")
End Sub
Private Sub txt_saldo_pendiente5_Change()
txt_saldo_pendiente5 = Format(txt_saldo_pendiente5, "##,##")
End Sub
Private Sub txt_saldo_pendiente6_Change()
txt_saldo_pendiente6 = Format(txt_saldo_pendiente6, "##,##")
End Sub
Private Sub txt_servicios_basicos_Change()
txt_servicios_basicos = Format(txt_servicios_basicos, "##,##")
End Sub

Private Sub txt_sueldos_Change()
txt_sueldos = Format(txt_sueldos, "##,##")
End Sub

Private Sub txt_total_activos_circulantes_Change()
txt_total_activos_circulantes = Format(txt_total_activos_circulantes, "##,##")
End Sub

Private Sub txt_total_costos_fijos_Change()
txt_total_costos_fijos = Format(txt_total_costos_fijos, "##,##")
End Sub

Private Sub txt_total_deudas_Change()
txt_total_deudas = Format(txt_total_deudas, "##,##")
End Sub

Private Sub txt_total_gasto_familiar_Change()
txt_total_gasto_familiar = Format(txt_total_gasto_familiar, "##,##")
End Sub

Private Sub txt_total_otros_ingresos_Change()
txt_total_otros_ingresos = Format(txt_total_otros_ingresos, "##,##")
End Sub

Private Sub txt_total_saldo_pendiente_Change()
txt_total_saldo_pendiente = Format(txt_total_saldo_pendiente, "##,##")
End Sub

Private Sub txt_valor_uf_Change()
txt_valor_uf = Format(txt_valor_uf, "##,##")
End Sub



Private Sub txt_venta_total_alto_Change()
txt_venta_total_alto = Format(txt_venta_total_alto, "##,##")
End Sub

Private Sub txt_venta_total_bajo_Change()
txt_venta_total_bajo = Format(txt_venta_total_bajo, "##,##")
End Sub

Private Sub txt_venta_total_medio_Change()
txt_venta_total_medio = Format(txt_venta_total_medio, "##,##")
End Sub

Private Sub txt_venta_total_mes_alto_corregida_Change()
txt_venta_total_mes_alto_corregida = Format(txt_venta_total_mes_alto_corregida, "##,##")
End Sub

Private Sub txt_venta_total_mes_bajo_corregida_Change()
txt_venta_total_mes_bajo_corregida = Format(txt_venta_total_mes_bajo_corregida, "##,##")
End Sub

Private Sub txt_venta_total_mes_medio_corregida_Change()
txt_venta_total_mes_medio_corregida = Format(txt_venta_total_mes_medio_corregida, "##,##")
End Sub

Private Sub txt_venta_total_promedio_anual_Change()
txt_venta_total_promedio_anual = Format(txt_venta_total_promedio_anual, "##,##")
End Sub

Private Sub txt_Venta_Total_Promedio_Mes_Alto_Change()
txt_Venta_Total_Promedio_Mes_Alto = Format(txt_Venta_Total_Promedio_Mes_Alto, "##,##")
End Sub

Private Sub txt_Venta_Total_Promedio_Mes_Bajo_Change()
txt_Venta_Total_Promedio_Mes_Bajo = Format(txt_Venta_Total_Promedio_Mes_Bajo, "##,##")
End Sub

Private Sub txt_Venta_Total_Promedio_Mes_Medio_Change()
txt_Venta_Total_Promedio_Mes_Medio = Format(txt_Venta_Total_Promedio_Mes_Medio, "##,##")
End Sub

Private Sub txt_vta_formal_promedio_mes_alto_Change()
txt_vta_formal_promedio_mes_alto = Format(txt_vta_formal_promedio_mes_alto, "##,##")
End Sub

Private Sub txt_vta_formal_promedio_mes_bajo_Change()
txt_vta_formal_promedio_mes_bajo = Format(txt_vta_formal_promedio_mes_bajo, "##,##")
End Sub

Private Sub txt_vta_formal_promedio_mes_medio_Change()
txt_vta_formal_promedio_mes_medio = Format(txt_vta_formal_promedio_mes_medio, "##,##")
End Sub

Private Sub txt_vta_informal_promedio_mes_alto_Change()
txt_vta_informal_promedio_mes_alto = Format(txt_vta_informal_promedio_mes_alto, "##,##")
End Sub

Private Sub txt_vta_informal_promedio_mes_bajo_Change()
txt_vta_informal_promedio_mes_bajo = Format(txt_vta_informal_promedio_mes_bajo, "##,##")
End Sub

Private Sub txt_vta_informal_promedio_mes_medio_Change()
txt_vta_informal_promedio_mes_medio = Format(txt_vta_informal_promedio_mes_medio, "##,##")
End Sub

Private Sub UserForm_Click()

End Sub
