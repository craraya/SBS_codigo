VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Visualizador_Met_IVA 
   Caption         =   "::::: Visualizador Metodologia I.V.A."
   ClientHeight    =   11190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   OleObjectBlob   =   "Visualizador_Met_IVA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Visualizador_Met_IVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_visualizador_iva_Click()
        
        Call conectarBD

        ssql = "SELECT a.RUT_CLIENTE,a.n_solicitud,a.DV,producto1,producto2,producto3,producto4,producto5,precio_venta1,precio_venta2,precio_venta3,precio_venta4,precio_venta5,materia_prima1,materia_prima2," _
            & " materia_prima3,materia_prima4,materia_prima5,mano_obra1,mano_obra2,mano_obra3,mano_obra4,mano_obra5,incidencia_ventas1,incidencia_ventas2,incidencia_ventas3,incidencia_ventas4,incidencia_ventas5,r_cvcmo1,r_cvcmo2,r_cvcmo3,r_cvcmo4,r_cvcmo5,r_cvsmo1,r_cvsmo2,r_cvsmo3,r_cvsmo4,r_cvsmo5,r_cvppcmo1,r_cvppcmo2,r_cvppcmo3,r_cvppcmo4,r_cvppcmo5,r_cvppsmo1,r_cvppsmo2,r_cvppsmo3,r_cvppsmo4,r_cvppsmo5," _
            & " r_Subtotal_costo_variable,r_Subtotal_x1_costo_variable,r_total_iva_credito,r_total_iva_debito,r_total_compra_neta,r_total_vta_netas_formales,r_total_vta_netas_informales,r_total_compra_total,r_total_vta_total,r_total_margen_total,r_promedio_iva_credito,r_promedio_iva_debito,r_promedio_compra_neta,r_promedio_vta_netas_formales,r_promedio_vta_netas_informales,r_promedio_compra_total,r_promedio_vta_total,compra_promedio_mensual," _
            & " veces_compra_mes,r_porcentaje_compra_formal,r_tot_promedio_ventas_mes_alto,r_tot_promedio_ventas_mes_medio,r_tot_promedio_ventas_mes_bajo,r_tot_promedio_ventas_formal_mes_alto,r_tot_promedio_ventas_formal_mes_medio,r_tot_promedio_ventas_formal_mes_bajo,r_tot_promedio_ventas_informal_mes_alto,r_tot_promedio_ventas_informal_mes_medio,r_tot_promedio_ventas_informal_mes_bajo,arriendo_micro,sueldos,movilizacion,servicios_basicos," _
            & " contador,lubricantes,neumaticos,afinamientos,patentes_seguros,otros_costos_fijos,total_costos_fijos,valor_uf,n_grupo_familiar,arriendo_vivienda_Gastos_Fam,gastos_indicado_cliente,total_gasto_familiar,liquidacion_sueldo,jubilacion,montepio,arriendo_vivienda_Otro_Ing,ingreso_segunda_microempresa,boleta_honorario,total_otros_ingresos,acreedor1_deuda,acreedor2_deuda,acreedor3_deuda,acreedor4_deuda,acreedor5_deuda,acreedor6_deuda," _
            & " tipo_producto1_deuda,tipo_producto2_deuda,tipo_producto3_deuda,tipo_producto4_deuda,tipo_producto5_deuda,tipo_producto6_deuda,saldo_pendiente1_deuda,saldo_pendiente2_deuda,saldo_pendiente3_deuda,saldo_pendiente4_deuda,saldo_pendiente5_deuda,saldo_pendiente6_deuda,monto_cuota1_deuda,monto_cuota2_deuda,monto_cuota3_deuda,monto_cuota4_deuda,monto_cuota5_deuda,monto_cuota6_deuda,cuotas_pactadas1_deuda,cuotas_pactadas2_deuda,cuotas_pactadas3_deuda," _
            & " cuotas_pactadas4_deuda,cuotas_pactadas5_deuda,cuotas_pactadas6_deuda,cuotas_pendientes1_deuda,cuotas_pendientes2_deuda,cuotas_pendientes3_deuda,cuotas_pendientes4_deuda,cuotas_pendientes5_deuda,cuotas_pendientes6_deuda, prepaga_cuota1_deuda,prepaga_cuota2_deuda,prepaga_cuota3_deuda,prepaga_cuota4_deuda,prepaga_cuota5_deuda,prepaga_cuota6_deuda,total_saldo_pendiente_deuda,total_deudas,numero_meses_alto_flujo,numero_meses_medio_flujo,numero_meses_bajo_flujo," _
            & " vta_formal_promedio_mes_alto_flujo,vta_formal_promedio_mes_medio_flujo,vta_formal_promedio_mes_bajo_flujo,vta_informal_promedio_mes_alto_flujo,vta_informal_promedio_mes_medio_flujo,vta_informal_promedio_mes_bajo_flujo,Venta_Total_Promedio_Mes_Alto_flujo,Venta_Total_Promedio_Mes_medio_flujo,Venta_Total_Promedio_Mes_bajo_flujo,resultado_operacional_alto_flujo,resultado_operacional_medio_flujo,resultado_operacional_bajo_flujo,capacidad_pago_mes_alto_flujo," _
            & " capacidad_pago_mes_medio_flujo,capacidad_pago_mes_bajo_flujo,cap_pago_corregida_ajus_mes_alto_flujo,cap_pago_corregida_ajus_mes_medio_flujo,cap_pago_corregida_ajus_mes_bajo_flujo,cap_pago_promedio_corregida_ajustada_flujo,monto_maximo_credito_flujo,cuota_credito_flujo,mto_bruto_solicitado_cliente_flujo,resolucion_credito_cuota_flujo,resolucion_credito_monto_flujo,a.fecha_ingreso,a.hora_ingreso,b.Rut_Cliente, b.n_solicitud,b.Dv,b.Ano_Declaracion_Iva_Ene,b.Iva_Credito_Ene,b.Iva_Debito_Ene,b.Compras_Netas_Ene,b.Ventas_Netas_Formales_Ene,b.Ventas_Netas_Informales_Ene,b.Compra_Total_Ene,b.Venta_Total_Ene,b.Tipo_Mes_Ene,b.Margen_Total_Ene," _
            & " b.Ano_Declaracion_Iva_feb,b.Iva_Credito_feb,b.Iva_Debito_feb,b.Compras_Netas_feb,b.Ventas_Netas_Formales_feb,b.Ventas_Netas_Informales_feb,b.Compra_Total_feb,b.Venta_Total_feb,b.Tipo_Mes_feb,b.Margen_Total_feb,b.Ano_Declaracion_Iva_mar,b.Iva_Credito_mar,b.Iva_Debito_mar,b.Compras_Netas_mar,b.Ventas_Netas_Formales_mar,b.Ventas_Netas_Informales_mar,b.Compra_Total_mar,b.Venta_Total_mar,b.Tipo_Mes_mar,b.Margen_Total_mar,b.Ano_Declaracion_Iva_abr,b.Iva_Credito_abr,b.Iva_Debito_abr,b.Compras_Netas_abr," _
            & " b.Ventas_Netas_Formales_abr,b.Ventas_Netas_Informales_abr,b.Compra_Total_abr,b.Venta_Total_abr,b.Tipo_Mes_abr,b.margen_Total_abr,b.Ano_Declaracion_Iva_may,b.Iva_Credito_may,b.Iva_Debito_may,b.Compras_Netas_may,b.Ventas_Netas_Formales_may,b.Ventas_Netas_Informales_may,b.Compra_Total_may,b.Venta_Total_may,b.Tipo_Mes_may,b.margen_Total_may,b.Ano_Declaracion_Iva_jun,b.Iva_Credito_jun,b.Iva_Debito_jun,b.Compras_Netas_jun,b.Ventas_Netas_Formales_jun,b.Ventas_Netas_Informales_jun,b.Compra_Total_jun," _
            & " b.Venta_Total_jun, b.Tipo_Mes_jun,b.margen_Total_jun,b.Ano_Declaracion_Iva_jul,b.Iva_Credito_jul,b.Iva_Debito_jul,b.Compras_Netas_jul,b.Ventas_Netas_Formales_jul,b.Ventas_Netas_Informales_jul,b.Compra_Total_jul,b.Venta_Total_jul,b.Tipo_Mes_jul,b.margen_Total_jul,b.Ano_Declaracion_Iva_ago,b.Iva_Credito_ago,b.Iva_Debito_ago,b.Compras_Netas_ago,b.Ventas_Netas_Formales_ago,b.Ventas_Netas_Informales_ago,b.Compra_Total_ago,b.Venta_Total_ago,b.Tipo_Mes_ago,b.margen_Total_ago,b.Ano_Declaracion_Iva_sep,b.Iva_Credito_sep,b.Iva_Debito_sep," _
            & " b.Compras_Netas_sep,b.Ventas_Netas_Formales_sep,b.Ventas_Netas_Informales_sep,b.Compra_Total_sep,b.Venta_Total_sep,b.Tipo_Mes_sep,b.margen_Total_sep,b.Ano_Declaracion_Iva_oct,b.Iva_Credito_oct,b.Iva_Debito_oct,b.Compras_Netas_oct,b.Ventas_Netas_Formales_oct,b.Ventas_Netas_Informales_oct, b.Compra_Total_oct,b.Venta_Total_oct,b.Tipo_Mes_oct,b.margen_Total_oct,b.Ano_Declaracion_Iva_nov,b.Iva_Credito_nov,b.Iva_Debito_nov,b.Compras_Netas_nov,b.Ventas_Netas_Formales_nov,b.Ventas_Netas_Informales_nov,b.Compra_Total_nov,b.Venta_Total_nov,b.Tipo_Mes_nov,b.margen_Total_nov," _
            & " b.Ano_Declaracion_Iva_dic,b.Iva_Credito_dic,b.Iva_Debito_dic,b.Compras_Netas_dic,b.Ventas_Netas_Formales_dic,b.Ventas_Netas_Informales_dic,b.Compra_Total_dic,b.Venta_Total_dic,b.Tipo_Mes_dic,b.margen_Total_dic,b.Fecha_Ingreso,b.Hora_Ingreso,b.mes_inicio_iva,impuesto" _
            & " FROM TBL_MICRO_METODOLOGIA_IVA a, tbl_micro_iva_mes b" _
            & " WHERE '" & txt_rut_cliente_ing & "' = a.rut_cliente" _
            & " and a.rut_cliente = b.rut_cliente" _
            & " order by a.n_solicitud desc"
        
            Set rst = cnn.Execute(ssql, , adCmdText)
        
            If rst.EOF Then
        
            MsgBox ("Cliente Sin Ingreso De Evaluacion"), vbCritical
        
            Else
            
            txt_rut_cliente = rst!rut_cliente
            txt_n_solicitud = rst!n_solicitud
            txt_dv = rst!dv
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
            txt_total_iva_credito = rst!r_total_iva_credito
            txt_total_iva_debito = rst!r_total_iva_debito
            txt_total_compra_neta = rst!r_total_compra_neta
            txt_total_vta_netas_formales = rst!r_total_vta_netas_formales
            txt_total_vta_netas_informales = rst!r_total_vta_netas_informales
            txt_total_compra_total = rst!r_total_compra_total
            txt_total_vta_total = rst!r_total_vta_total
            txt_total_margen_total = rst!r_total_margen_total
            txt_promedio_iva_credito = rst!r_promedio_iva_credito
            txt_promedio_iva_debito = rst!r_promedio_iva_debito
            txt_promedio_compra_neta = rst!r_promedio_compra_neta
            txt_promedio_vta_netas_formales = rst!r_promedio_vta_netas_formales
            txt_promedio_vta_netas_informales = rst!r_promedio_vta_netas_informales
            txt_promedio_compra_total = rst!r_promedio_compra_total
            txt_promedio_vta_total = rst!r_promedio_vta_total
            txt_compra_promedio_mensual = rst!compra_promedio_mensual
            txt_veces_compra_mes = rst!veces_compra_mes
            txt_porcentaje_compra_formal = rst!r_porcentaje_compra_formal
            txt_prom_vta_meses_altos = rst!r_tot_promedio_ventas_mes_alto
            txt_prom_vta_meses_medios = rst!r_tot_promedio_ventas_mes_medio
            txt_prom_vta_meses_bajos = rst!r_tot_promedio_ventas_mes_bajo
            txt_prom_vtas_meses_altos_formal = rst!r_tot_promedio_ventas_formal_mes_alto
            txt_prom_vtas_meses_medios_formal = rst!r_tot_promedio_ventas_formal_mes_medio
            txt_prom_vtas_meses_bajos_formal = rst!r_tot_promedio_ventas_formal_mes_bajo
            txt_prom_vtas_meses_altos_informal = rst!r_tot_promedio_ventas_informal_mes_alto
            txt_prom_vtas_meses_medios_informal = rst!r_tot_promedio_ventas_informal_mes_medio
            txt_prom_vtas_meses_bajos_informal = rst!r_tot_promedio_ventas_informal_mes_bajo
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
            
            ''''' VARIABLES DE IVA MES
            
            'txt_rut_cliente = rst!txt_rut_cliente
            'txt_n_solicitud = rst!txt_n_solicitud
            'txt_dv_ = rst!txt_dv_
            txt_ano_iva_ene = rst!Ano_Declaracion_Iva_Ene
            txt_iva_credito_ene = rst!Iva_Credito_Ene
            txt_iva_debito_ene = rst!Iva_Debito_Ene
            txt_compra_neta_ene = rst!Compras_Netas_Ene
            txt_vta_netas_formales_ene = rst!Ventas_Netas_Formales_Ene
            txt_vta_netas_informales_ene = rst!Ventas_Netas_Informales_Ene
            txt_compra_total_ene = rst!Compra_Total_Ene
            txt_vta_total_ene = rst!Venta_Total_Ene
            txt_tipo_mes_ene = rst!Tipo_Mes_Ene
            txt_margen_total_ene = rst!Margen_Total_Ene
            txt_ano_iva_feb = rst!Ano_Declaracion_Iva_feb
            txt_iva_credito_feb = rst!Iva_Credito_feb
            txt_iva_debito_feb = rst!Iva_Debito_feb
            txt_compra_neta_feb = rst!Compras_Netas_feb
            txt_vta_netas_formales_feb = rst!Ventas_Netas_Formales_feb
            txt_vta_netas_informales_feb = rst!Ventas_Netas_Informales_feb
            txt_compra_total_feb = rst!Compra_Total_feb
            txt_vta_total_feb = rst!Venta_Total_feb
            txt_tipo_mes_feb = rst!Tipo_Mes_feb
            txt_margen_total_feb = rst!Margen_Total_feb
            txt_ano_iva_mar = rst!Ano_Declaracion_Iva_mar
            txt_iva_credito_mar = rst!Iva_Credito_mar
            txt_iva_debito_mar = rst!Iva_Debito_mar
            txt_compra_neta_mar = rst!Compras_Netas_mar
            txt_vta_netas_formales_mar = rst!Ventas_Netas_Formales_mar
            txt_vta_netas_informales_mar = rst!Ventas_Netas_Informales_mar
            txt_compra_total_mar = rst!Compra_Total_mar
            txt_vta_total_mar = rst!Venta_Total_mar
            txt_tipo_mes_mar = rst!Tipo_Mes_mar
            txt_margen_total_mar = rst!Margen_Total_mar
            txt_ano_iva_abr = rst!Ano_Declaracion_Iva_abr
            txt_iva_credito_abr = rst!Iva_Credito_abr
            txt_iva_debito_abr = rst!Iva_Debito_abr
            txt_compra_neta_abr = rst!Compras_Netas_abr
            txt_vta_netas_formales_abr = rst!Ventas_Netas_Formales_abr
            txt_vta_netas_informales_abr = rst!Ventas_Netas_Informales_abr
            txt_compra_total_abr = rst!Compra_Total_abr
            txt_vta_total_abr = rst!Venta_Total_abr
            txt_tipo_mes_abr = rst!Tipo_Mes_abr
            txt_margen_total_abr = rst!margen_Total_abr
            txt_ano_iva_may = rst!Ano_Declaracion_Iva_may
            txt_iva_credito_may = rst!Iva_Credito_may
            txt_iva_debito_may = rst!Iva_Debito_may
            txt_compra_neta_may = rst!Compras_Netas_may
            txt_vta_netas_formales_may = rst!Ventas_Netas_Formales_may
            txt_vta_netas_informales_may = rst!Ventas_Netas_Informales_may
            txt_compra_total_may = rst!Compra_Total_may
            txt_vta_total_may = rst!Venta_Total_may
            txt_tipo_mes_may = rst!Tipo_Mes_may
            txt_margen_total_may = rst!margen_Total_may
            txt_ano_iva_jun = rst!Ano_Declaracion_Iva_jun
            txt_iva_credito_jun = rst!Iva_Credito_jun
            txt_iva_debito_jun = rst!Iva_Debito_jun
            txt_compra_neta_jun = rst!Compras_Netas_jun
            txt_vta_netas_formales_jun = rst!Ventas_Netas_Formales_jun
            txt_vta_netas_informales_jun = rst!Ventas_Netas_Informales_jun
            txt_compra_total_jun = rst!Compra_Total_jun
            txt_vta_total_jun = rst!Venta_Total_jun
            txt_tipo_mes_jun = rst!Tipo_Mes_jun
            txt_margen_total_jun = rst!margen_Total_jun
            txt_ano_iva_jul = rst!Ano_Declaracion_Iva_jul
            txt_iva_credito_jul = rst!Iva_Credito_jul
            txt_iva_debito_jul = rst!Iva_Debito_jul
            txt_compra_neta_jul = rst!Compras_Netas_jul
            txt_vta_netas_formales_jul = rst!Ventas_Netas_Formales_jul
            txt_vta_netas_informales_jul = rst!Ventas_Netas_Informales_jul
            txt_compra_total_jul = rst!Compra_Total_jul
            txt_vta_total_jul = rst!Venta_Total_jul
            txt_tipo_mes_jul = rst!Tipo_Mes_jul
            txt_margen_total_jul = rst!margen_Total_jul
            txt_ano_iva_ago = rst!Ano_Declaracion_Iva_ago
            txt_iva_credito_ago = rst!Iva_Credito_ago
            txt_iva_debito_ago = rst!Iva_Debito_ago
            txt_compra_neta_ago = rst!Compras_Netas_ago
            txt_vta_netas_formales_ago = rst!Ventas_Netas_Formales_ago
            txt_vta_netas_informales_ago = rst!Ventas_Netas_Informales_ago
            txt_compra_total_ago = rst!Compra_Total_ago
            txt_vta_total_ago = rst!Venta_Total_ago
            txt_tipo_mes_ago = rst!Tipo_Mes_ago
            txt_margen_total_ago = rst!margen_Total_ago
            txt_ano_iva_sep = rst!Ano_Declaracion_Iva_sep
            txt_iva_credito_sep = rst!Iva_Credito_sep
            txt_iva_debito_sep = rst!Iva_Debito_sep
            txt_compra_neta_sep = rst!Compras_Netas_sep
            txt_vta_netas_formales_sep = rst!Ventas_Netas_Formales_sep
            txt_vta_netas_informales_sep = rst!Ventas_Netas_Informales_sep
            txt_compra_total_sep = rst!Compra_Total_sep
            txt_vta_total_sep = rst!Venta_Total_sep
            txt_tipo_mes_sep = rst!Tipo_Mes_sep
            txt_margen_total_sep = rst!margen_Total_sep
            txt_ano_iva_oct = rst!Ano_Declaracion_Iva_oct
            txt_iva_credito_oct = rst!Iva_Credito_oct
            txt_iva_debito_oct = rst!Iva_Debito_oct
            txt_compra_neta_oct = rst!Compras_Netas_oct
            txt_vta_netas_formales_oct = rst!Ventas_Netas_Formales_oct
            txt_vta_netas_informales_oct = rst!Ventas_Netas_Informales_oct
            txt_compra_total_oct = rst!Compra_Total_oct
            txt_vta_total_oct = rst!Venta_Total_oct
            txt_tipo_mes_oct = rst!Tipo_Mes_oct
            txt_margen_total_oct = rst!margen_Total_oct
            txt_ano_iva_nov = rst!Ano_Declaracion_Iva_nov
            txt_iva_credito_nov = rst!Iva_Credito_nov
            txt_iva_debito_nov = rst!Iva_Debito_nov
            txt_compra_neta_nov = rst!Compras_Netas_nov
            txt_vta_netas_formales_nov = rst!Ventas_Netas_Formales_nov
            txt_vta_netas_informales_nov = rst!Ventas_Netas_Informales_nov
            txt_compra_total_nov = rst!Compra_Total_nov
            txt_vta_total_nov = rst!Venta_Total_nov
            txt_tipo_mes_nov = rst!Tipo_Mes_nov
            txt_margen_total_nov = rst!margen_Total_nov
            txt_ano_iva_dic = rst!Ano_Declaracion_Iva_dic
            txt_iva_credito_dic = rst!Iva_Credito_dic
            txt_iva_debito_dic = rst!Iva_Debito_dic
            txt_compra_neta_dic = rst!Compras_Netas_dic
            txt_vta_netas_formales_dic = rst!Ventas_Netas_Formales_dic
            txt_vta_netas_informales_dic = rst!Ventas_Netas_Informales_dic
            txt_compra_total_dic = rst!Compra_Total_dic
            txt_vta_total_dic = rst!Venta_Total_dic
            txt_tipo_mes_dic = rst!Tipo_Mes_dic
            txt_margen_total_dic = rst!margen_Total_dic
            txt_fecha_actual = rst!FECHA_INGRESO
            txt_hora_actual = rst!HORA_INGRESO
            cbx_mes_inicio_iva = rst!mes_inicio_iva
            
        End If
End Sub



Private Sub cmd_volver_menu_inicial_1_Click()
cmd_volver_menu_inicial_Click
End Sub

Private Sub cmd_volver_menu_inicial_Click()
Visualizador_Met_IVA.Hide
Visualizador_Inicial.Show
Visualizador_Inicial.txt_rut_cliente = Empty

End Sub

Private Sub CommandButton1_Click()
Visualizador_Met_IVA.PrintForm
End Sub

Private Sub CommandButton2_Click()
Visualizador_Met_IVA.PrintForm
End Sub

Private Sub CommandButton3_Click()
Visualizador_Met_IVA.PrintForm
End Sub

Private Sub CommandButton4_Click()
Visualizador_Met_IVA.PrintForm
End Sub

Private Sub CommandButton5_Click()
Visualizador_Met_IVA.PrintForm
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

Private Sub txt_compra_neta_ene_Change()
txt_compra_neta_ene = Format(txt_compra_neta_ene, "##,##")
End Sub
Private Sub txt_compra_neta_feb_Change()
txt_compra_neta_feb = Format(txt_compra_neta_feb, "##,##")
End Sub
Private Sub txt_compra_neta_mar_Change()
txt_compra_neta_mar = Format(txt_compra_neta_mar, "##,##")
End Sub
Private Sub txt_compra_neta_abr_Change()
txt_compra_neta_abr = Format(txt_compra_neta_abr, "##,##")
End Sub
Private Sub txt_compra_neta_may_Change()
txt_compra_neta_may = Format(txt_compra_neta_may, "##,##")
End Sub
Private Sub txt_compra_neta_jun_Change()
txt_compra_neta_jun = Format(txt_compra_neta_jun, "##,##")
End Sub
Private Sub txt_compra_neta_jul_Change()
txt_compra_neta_jul = Format(txt_compra_neta_jul, "##,##")
End Sub
Private Sub txt_compra_neta_ago_Change()
txt_compra_neta_ago = Format(txt_compra_neta_ago, "##,##")
End Sub
Private Sub txt_compra_neta_sep_Change()
txt_compra_neta_sep = Format(txt_compra_neta_sep, "##,##")
End Sub
Private Sub txt_compra_neta_oct_Change()
txt_compra_neta_oct = Format(txt_compra_neta_oct, "##,##")
End Sub
Private Sub txt_compra_neta_nov_Change()
txt_compra_neta_nov = Format(txt_compra_neta_nov, "##,##")
End Sub
Private Sub txt_compra_neta_dic_Change()
txt_compra_neta_dic = Format(txt_compra_neta_dic, "##,##")
End Sub

Private Sub txt_compra_promedio_mensual_Change()
txt_compra_promedio_mensual = Format(txt_compra_promedio_mensual, "##,##")
End Sub

Private Sub txt_compra_total_ene_Change()
txt_compra_total_ene = Format(txt_compra_total_ene, "##,##")
End Sub
Private Sub txt_compra_total_feb_Change()
txt_compra_total_feb = Format(txt_compra_total_feb, "##,##")
End Sub
Private Sub txt_compra_total_mar_Change()
txt_compra_total_mar = Format(txt_compra_total_mar, "##,##")
End Sub
Private Sub txt_compra_total_abr_Change()
txt_compra_total_abr = Format(txt_compra_total_abr, "##,##")
End Sub
Private Sub txt_compra_total_may_Change()
txt_compra_total_may = Format(txt_compra_total_may, "##,##")
End Sub
Private Sub txt_compra_total_jun_Change()
txt_compra_total_jun = Format(txt_compra_total_jun, "##,##")
End Sub
Private Sub txt_compra_total_jul_Change()
txt_compra_total_jul = Format(txt_compra_total_jul, "##,##")
End Sub
Private Sub txt_compra_total_ago_Change()
txt_compra_total_ago = Format(txt_compra_total_ago, "##,##")
End Sub
Private Sub txt_compra_total_sep_Change()
txt_compra_total_sep = Format(txt_compra_total_sep, "##,##")
End Sub
Private Sub txt_compra_total_oct_Change()
txt_compra_total_oct = Format(txt_compra_total_oct, "##,##")
End Sub
Private Sub txt_compra_total_nov_Change()
txt_compra_total_nov = Format(txt_compra_total_nov, "##,##")
End Sub
Private Sub txt_compra_total_dic_Change()
txt_compra_total_dic = Format(txt_compra_total_dic, "##,##")
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

        ssql = "SELECT a.RUT_CLIENTE,a.n_solicitud,a.DV,producto1,producto2,producto3,producto4,producto5,precio_venta1,precio_venta2,precio_venta3,precio_venta4,precio_venta5,materia_prima1,materia_prima2," _
            & " materia_prima3,materia_prima4,materia_prima5,mano_obra1,mano_obra2,mano_obra3,mano_obra4,mano_obra5,incidencia_ventas1,incidencia_ventas2,incidencia_ventas3,incidencia_ventas4,incidencia_ventas5,r_cvcmo1,r_cvcmo2,r_cvcmo3,r_cvcmo4,r_cvcmo5,r_cvsmo1,r_cvsmo2,r_cvsmo3,r_cvsmo4,r_cvsmo5,r_cvppcmo1,r_cvppcmo2,r_cvppcmo3,r_cvppcmo4,r_cvppcmo5,r_cvppsmo1,r_cvppsmo2,r_cvppsmo3,r_cvppsmo4,r_cvppsmo5," _
            & " r_Subtotal_costo_variable,r_Subtotal_x1_costo_variable,r_total_iva_credito,r_total_iva_debito,r_total_compra_neta,r_total_vta_netas_formales,r_total_vta_netas_informales,r_total_compra_total,r_total_vta_total,r_total_margen_total,r_promedio_iva_credito,r_promedio_iva_debito,r_promedio_compra_neta,r_promedio_vta_netas_formales,r_promedio_vta_netas_informales,r_promedio_compra_total,r_promedio_vta_total,compra_promedio_mensual," _
            & " veces_compra_mes,r_porcentaje_compra_formal,r_tot_promedio_ventas_mes_alto,r_tot_promedio_ventas_mes_medio,r_tot_promedio_ventas_mes_bajo,r_tot_promedio_ventas_formal_mes_alto,r_tot_promedio_ventas_formal_mes_medio,r_tot_promedio_ventas_formal_mes_bajo,r_tot_promedio_ventas_informal_mes_alto,r_tot_promedio_ventas_informal_mes_medio,r_tot_promedio_ventas_informal_mes_bajo,arriendo_micro,sueldos,movilizacion,servicios_basicos," _
            & " contador,lubricantes,neumaticos,afinamientos,patentes_seguros,otros_costos_fijos,total_costos_fijos,valor_uf,n_grupo_familiar,arriendo_vivienda_Gastos_Fam,gastos_indicado_cliente,total_gasto_familiar,liquidacion_sueldo,jubilacion,montepio,arriendo_vivienda_Otro_Ing,ingreso_segunda_microempresa,boleta_honorario,total_otros_ingresos,acreedor1_deuda,acreedor2_deuda,acreedor3_deuda,acreedor4_deuda,acreedor5_deuda,acreedor6_deuda," _
            & " tipo_producto1_deuda,tipo_producto2_deuda,tipo_producto3_deuda,tipo_producto4_deuda,tipo_producto5_deuda,tipo_producto6_deuda,saldo_pendiente1_deuda,saldo_pendiente2_deuda,saldo_pendiente3_deuda,saldo_pendiente4_deuda,saldo_pendiente5_deuda,saldo_pendiente6_deuda,monto_cuota1_deuda,monto_cuota2_deuda,monto_cuota3_deuda,monto_cuota4_deuda,monto_cuota5_deuda,monto_cuota6_deuda,cuotas_pactadas1_deuda,cuotas_pactadas2_deuda,cuotas_pactadas3_deuda," _
            & " cuotas_pactadas4_deuda,cuotas_pactadas5_deuda,cuotas_pactadas6_deuda,cuotas_pendientes1_deuda,cuotas_pendientes2_deuda,cuotas_pendientes3_deuda,cuotas_pendientes4_deuda,cuotas_pendientes5_deuda,cuotas_pendientes6_deuda, prepaga_cuota1_deuda,prepaga_cuota2_deuda,prepaga_cuota3_deuda,prepaga_cuota4_deuda,prepaga_cuota5_deuda,prepaga_cuota6_deuda,total_saldo_pendiente_deuda,total_deudas,numero_meses_alto_flujo,numero_meses_medio_flujo,numero_meses_bajo_flujo," _
            & " vta_formal_promedio_mes_alto_flujo,vta_formal_promedio_mes_medio_flujo,vta_formal_promedio_mes_bajo_flujo,vta_informal_promedio_mes_alto_flujo,vta_informal_promedio_mes_medio_flujo,vta_informal_promedio_mes_bajo_flujo,Venta_Total_Promedio_Mes_Alto_flujo,Venta_Total_Promedio_Mes_medio_flujo,Venta_Total_Promedio_Mes_bajo_flujo,resultado_operacional_alto_flujo,resultado_operacional_medio_flujo,resultado_operacional_bajo_flujo,capacidad_pago_mes_alto_flujo," _
            & " capacidad_pago_mes_medio_flujo,capacidad_pago_mes_bajo_flujo,cap_pago_corregida_ajus_mes_alto_flujo,cap_pago_corregida_ajus_mes_medio_flujo,cap_pago_corregida_ajus_mes_bajo_flujo,cap_pago_promedio_corregida_ajustada_flujo,monto_maximo_credito_flujo,cuota_credito_flujo,mto_bruto_solicitado_cliente_flujo,resolucion_credito_cuota_flujo,resolucion_credito_monto_flujo,a.fecha_ingreso,a.hora_ingreso,b.Rut_Cliente, b.n_solicitud,b.Dv,b.Ano_Declaracion_Iva_Ene,b.Iva_Credito_Ene,b.Iva_Debito_Ene,b.Compras_Netas_Ene,b.Ventas_Netas_Formales_Ene,b.Ventas_Netas_Informales_Ene,b.Compra_Total_Ene,b.Venta_Total_Ene,b.Tipo_Mes_Ene,b.Margen_Total_Ene," _
            & " b.Ano_Declaracion_Iva_feb,b.Iva_Credito_feb,b.Iva_Debito_feb,b.Compras_Netas_feb,b.Ventas_Netas_Formales_feb,b.Ventas_Netas_Informales_feb,b.Compra_Total_feb,b.Venta_Total_feb,b.Tipo_Mes_feb,b.Margen_Total_feb,b.Ano_Declaracion_Iva_mar,b.Iva_Credito_mar,b.Iva_Debito_mar,b.Compras_Netas_mar,b.Ventas_Netas_Formales_mar,b.Ventas_Netas_Informales_mar,b.Compra_Total_mar,b.Venta_Total_mar,b.Tipo_Mes_mar,b.Margen_Total_mar,b.Ano_Declaracion_Iva_abr,b.Iva_Credito_abr,b.Iva_Debito_abr,b.Compras_Netas_abr," _
            & " b.Ventas_Netas_Formales_abr,b.Ventas_Netas_Informales_abr,b.Compra_Total_abr,b.Venta_Total_abr,b.Tipo_Mes_abr,b.margen_Total_abr,b.Ano_Declaracion_Iva_may,b.Iva_Credito_may,b.Iva_Debito_may,b.Compras_Netas_may,b.Ventas_Netas_Formales_may,b.Ventas_Netas_Informales_may,b.Compra_Total_may,b.Venta_Total_may,b.Tipo_Mes_may,b.margen_Total_may,b.Ano_Declaracion_Iva_jun,b.Iva_Credito_jun,b.Iva_Debito_jun,b.Compras_Netas_jun,b.Ventas_Netas_Formales_jun,b.Ventas_Netas_Informales_jun,b.Compra_Total_jun," _
            & " b.Venta_Total_jun, b.Tipo_Mes_jun,b.margen_Total_jun,b.Ano_Declaracion_Iva_jul,b.Iva_Credito_jul,b.Iva_Debito_jul,b.Compras_Netas_jul,b.Ventas_Netas_Formales_jul,b.Ventas_Netas_Informales_jul,b.Compra_Total_jul,b.Venta_Total_jul,b.Tipo_Mes_jul,b.margen_Total_jul,b.Ano_Declaracion_Iva_ago,b.Iva_Credito_ago,b.Iva_Debito_ago,b.Compras_Netas_ago,b.Ventas_Netas_Formales_ago,b.Ventas_Netas_Informales_ago,b.Compra_Total_ago,b.Venta_Total_ago,b.Tipo_Mes_ago,b.margen_Total_ago,b.Ano_Declaracion_Iva_sep,b.Iva_Credito_sep,b.Iva_Debito_sep," _
            & " b.Compras_Netas_sep,b.Ventas_Netas_Formales_sep,b.Ventas_Netas_Informales_sep,b.Compra_Total_sep,b.Venta_Total_sep,b.Tipo_Mes_sep,b.margen_Total_sep,b.Ano_Declaracion_Iva_oct,b.Iva_Credito_oct,b.Iva_Debito_oct,b.Compras_Netas_oct,b.Ventas_Netas_Formales_oct,b.Ventas_Netas_Informales_oct, b.Compra_Total_oct,b.Venta_Total_oct,b.Tipo_Mes_oct,b.margen_Total_oct,b.Ano_Declaracion_Iva_nov,b.Iva_Credito_nov,b.Iva_Debito_nov,b.Compras_Netas_nov,b.Ventas_Netas_Formales_nov,b.Ventas_Netas_Informales_nov,b.Compra_Total_nov,b.Venta_Total_nov,b.Tipo_Mes_nov,b.margen_Total_nov," _
            & " b.Ano_Declaracion_Iva_dic,b.Iva_Credito_dic,b.Iva_Debito_dic,b.Compras_Netas_dic,b.Ventas_Netas_Formales_dic,b.Ventas_Netas_Informales_dic,b.Compra_Total_dic,b.Venta_Total_dic,b.Tipo_Mes_dic,b.margen_Total_dic,b.Fecha_Ingreso,b.Hora_Ingreso,isnull(b.mes_inicio_iva,'SD') as Mes_Inicio_Iva, impuesto,ISNULL(tipo_credito_deuda1,'SD')tipo_credito_deuda1,ISNULL(tipo_credito_deuda2,'SD')tipo_credito_deuda2,ISNULL(tipo_credito_deuda3,'SD')tipo_credito_deuda3,ISNULL(tipo_credito_deuda4,'SD')tipo_credito_deuda4,ISNULL(tipo_credito_deuda5,'SD')tipo_credito_deuda5,ISNULL(tipo_credito_deuda6,'SD')tipo_credito_deuda6," _
            & " ISNULL(total_saldo_pendiente_consumo,'0')total_saldo_pendiente_consumo,ISNULL(total_deudas_consumo,'0')total_deudas_consumo,ISNULL(total_saldo_pendiente_comercial,'0')total_saldo_pendiente_comercial,ISNULL(total_deudas_comercial,'0')total_deudas_comercial,ISNULL(saldo_deuda_con_prepago_consumo,'0')saldo_deuda_con_prepago_consumo,ISNULL(saldo_deuda_con_prepago_comercial,'0')saldo_deuda_con_prepago_comercial,ISNULL(mto_cuota_con_prepago_consumo,'0')mto_cuota_con_prepago_consumo,ISNULL(mto_cuota_con_prepago_comercial,'0')mto_cuota_con_prepago_comercial,ISNULL(saldo_deuda_sin_prepago_consumo,'0')saldo_deuda_sin_prepago_consumo,ISNULL(saldo_deuda_sin_prepago_comercial,'0')saldo_deuda_sin_prepago_comercial,ISNULL(mto_cuota_sin_prepago_comercial,'0')mto_cuota_sin_prepago_comercial,ISNULL(mto_cuota_sin_prepago_consumo,'0')mto_cuota_sin_prepago_consumo" _
            & " FROM TBL_MICRO_METODOLOGIA_IVA a, tbl_micro_iva_mes b" _
            & " WHERE '" & txt_rut_cliente_ing & "' = a.rut_cliente" _
            & " and a.rut_cliente = b.rut_cliente" _
            & " and '" & txt_n_solicitud & "' = a.n_solicitud" _
            & " order by a.n_solicitud desc, b.n_solicitud desc"
        
            Set rst = cnn.Execute(ssql, , adCmdText)
        
            If rst.EOF Then
        
            MsgBox ("Cliente Sin Ingreso De Evaluacion"), vbCritical
        
            Else
            'Dim txt_Venta_Total_Promedio_Mes_Alto As Integer
            'Dim txt_Venta_Total_Promedio_Mes_Medio As Integer
            'Dim txt_Venta_Total_Promedio_Mes_Bajo As Integer
            
            'txt_Venta_Total_Promedio_Mes_Alto = 0
            'txt_Venta_Total_Promedio_Mes_Medio = 0
            'txt_Venta_Total_Promedio_Mes_Bajo = 0
            
            
            txt_rut_cliente = rst!rut_cliente
            txt_n_solicitud = rst!n_solicitud
            txt_dv = rst!dv
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
            txt_total_iva_credito = rst!r_total_iva_credito
            txt_total_iva_debito = rst!r_total_iva_debito
            txt_total_compra_neta = rst!r_total_compra_neta
            txt_total_vta_netas_formales = rst!r_total_vta_netas_formales
            txt_total_vta_netas_informales = rst!r_total_vta_netas_informales
            txt_total_compra_total = rst!r_total_compra_total
            txt_total_vta_total = rst!r_total_vta_total
            txt_total_margen_total = rst!r_total_margen_total
            txt_promedio_iva_credito = rst!r_promedio_iva_credito
            txt_promedio_iva_debito = rst!r_promedio_iva_debito
            txt_promedio_compra_neta = rst!r_promedio_compra_neta
            txt_promedio_vta_netas_formales = rst!r_promedio_vta_netas_formales
            txt_promedio_vta_netas_informales = rst!r_promedio_vta_netas_informales
            txt_promedio_compra_total = rst!r_promedio_compra_total
            txt_promedio_vta_total = rst!r_promedio_vta_total
            txt_compra_promedio_mensual = rst!compra_promedio_mensual
            txt_veces_compra_mes = rst!veces_compra_mes
            txt_porcentaje_compra_formal = rst!r_porcentaje_compra_formal
            txt_prom_vta_meses_altos = rst!r_tot_promedio_ventas_mes_alto
            txt_prom_vta_meses_medios = rst!r_tot_promedio_ventas_mes_medio
            txt_prom_vta_meses_bajos = rst!r_tot_promedio_ventas_mes_bajo
            txt_prom_vtas_meses_altos_formal = rst!r_tot_promedio_ventas_formal_mes_alto
            txt_prom_vtas_meses_medios_formal = rst!r_tot_promedio_ventas_formal_mes_medio
            txt_prom_vtas_meses_bajos_formal = rst!r_tot_promedio_ventas_formal_mes_bajo
            txt_prom_vtas_meses_altos_informal = rst!r_tot_promedio_ventas_informal_mes_alto
            txt_prom_vtas_meses_medios_informal = rst!r_tot_promedio_ventas_informal_mes_medio
            txt_prom_vtas_meses_bajos_informal = rst!r_tot_promedio_ventas_informal_mes_bajo
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
            
          
            
            ''''' VARIABLES DE IVA MES
            
            'txt_rut_cliente = rst!txt_rut_cliente
            'txt_n_solicitud = rst!txt_n_solicitud
            'txt_dv_ = rst!txt_dv_
            txt_ano_iva_ene = rst!Ano_Declaracion_Iva_Ene
            txt_iva_credito_ene = rst!Iva_Credito_Ene
            txt_iva_debito_ene = rst!Iva_Debito_Ene
            txt_compra_neta_ene = rst!Compras_Netas_Ene
            txt_vta_netas_formales_ene = rst!Ventas_Netas_Formales_Ene
            txt_vta_netas_informales_ene = rst!Ventas_Netas_Informales_Ene
            txt_compra_total_ene = rst!Compra_Total_Ene
            txt_vta_total_ene = rst!Venta_Total_Ene
            txt_tipo_mes_ene = rst!Tipo_Mes_Ene
            txt_margen_total_ene = rst!Margen_Total_Ene
            txt_ano_iva_feb = rst!Ano_Declaracion_Iva_feb
            txt_iva_credito_feb = rst!Iva_Credito_feb
            txt_iva_debito_feb = rst!Iva_Debito_feb
            txt_compra_neta_feb = rst!Compras_Netas_feb
            txt_vta_netas_formales_feb = rst!Ventas_Netas_Formales_feb
            txt_vta_netas_informales_feb = rst!Ventas_Netas_Informales_feb
            txt_compra_total_feb = rst!Compra_Total_feb
            txt_vta_total_feb = rst!Venta_Total_feb
            txt_tipo_mes_feb = rst!Tipo_Mes_feb
            txt_margen_total_feb = rst!Margen_Total_feb
            txt_ano_iva_mar = rst!Ano_Declaracion_Iva_mar
            txt_iva_credito_mar = rst!Iva_Credito_mar
            txt_iva_debito_mar = rst!Iva_Debito_mar
            txt_compra_neta_mar = rst!Compras_Netas_mar
            txt_vta_netas_formales_mar = rst!Ventas_Netas_Formales_mar
            txt_vta_netas_informales_mar = rst!Ventas_Netas_Informales_mar
            txt_compra_total_mar = rst!Compra_Total_mar
            txt_vta_total_mar = rst!Venta_Total_mar
            txt_tipo_mes_mar = rst!Tipo_Mes_mar
            txt_margen_total_mar = rst!Margen_Total_mar
            txt_ano_iva_abr = rst!Ano_Declaracion_Iva_abr
            txt_iva_credito_abr = rst!Iva_Credito_abr
            txt_iva_debito_abr = rst!Iva_Debito_abr
            txt_compra_neta_abr = rst!Compras_Netas_abr
            txt_vta_netas_formales_abr = rst!Ventas_Netas_Formales_abr
            txt_vta_netas_informales_abr = rst!Ventas_Netas_Informales_abr
            txt_compra_total_abr = rst!Compra_Total_abr
            txt_vta_total_abr = rst!Venta_Total_abr
            txt_tipo_mes_abr = rst!Tipo_Mes_abr
            txt_margen_total_abr = rst!margen_Total_abr
            txt_ano_iva_may = rst!Ano_Declaracion_Iva_may
            txt_iva_credito_may = rst!Iva_Credito_may
            txt_iva_debito_may = rst!Iva_Debito_may
            txt_compra_neta_may = rst!Compras_Netas_may
            txt_vta_netas_formales_may = rst!Ventas_Netas_Formales_may
            txt_vta_netas_informales_may = rst!Ventas_Netas_Informales_may
            txt_compra_total_may = rst!Compra_Total_may
            txt_vta_total_may = rst!Venta_Total_may
            txt_tipo_mes_may = rst!Tipo_Mes_may
            txt_margen_total_may = rst!margen_Total_may
            txt_ano_iva_jun = rst!Ano_Declaracion_Iva_jun
            txt_iva_credito_jun = rst!Iva_Credito_jun
            txt_iva_debito_jun = rst!Iva_Debito_jun
            txt_compra_neta_jun = rst!Compras_Netas_jun
            txt_vta_netas_formales_jun = rst!Ventas_Netas_Formales_jun
            txt_vta_netas_informales_jun = rst!Ventas_Netas_Informales_jun
            txt_compra_total_jun = rst!Compra_Total_jun
            txt_vta_total_jun = rst!Venta_Total_jun
            txt_tipo_mes_jun = rst!Tipo_Mes_jun
            txt_margen_total_jun = rst!margen_Total_jun
            txt_ano_iva_jul = rst!Ano_Declaracion_Iva_jul
            txt_iva_credito_jul = rst!Iva_Credito_jul
            txt_iva_debito_jul = rst!Iva_Debito_jul
            txt_compra_neta_jul = rst!Compras_Netas_jul
            txt_vta_netas_formales_jul = rst!Ventas_Netas_Formales_jul
            txt_vta_netas_informales_jul = rst!Ventas_Netas_Informales_jul
            txt_compra_total_jul = rst!Compra_Total_jul
            txt_vta_total_jul = rst!Venta_Total_jul
            txt_tipo_mes_jul = rst!Tipo_Mes_jul
            txt_margen_total_jul = rst!margen_Total_jul
            txt_ano_iva_ago = rst!Ano_Declaracion_Iva_ago
            txt_iva_credito_ago = rst!Iva_Credito_ago
            txt_iva_debito_ago = rst!Iva_Debito_ago
            txt_compra_neta_ago = rst!Compras_Netas_ago
            txt_vta_netas_formales_ago = rst!Ventas_Netas_Formales_ago
            txt_vta_netas_informales_ago = rst!Ventas_Netas_Informales_ago
            txt_compra_total_ago = rst!Compra_Total_ago
            txt_vta_total_ago = rst!Venta_Total_ago
            txt_tipo_mes_ago = rst!Tipo_Mes_ago
            txt_margen_total_ago = rst!margen_Total_ago
            txt_ano_iva_sep = rst!Ano_Declaracion_Iva_sep
            txt_iva_credito_sep = rst!Iva_Credito_sep
            txt_iva_debito_sep = rst!Iva_Debito_sep
            txt_compra_neta_sep = rst!Compras_Netas_sep
            txt_vta_netas_formales_sep = rst!Ventas_Netas_Formales_sep
            txt_vta_netas_informales_sep = rst!Ventas_Netas_Informales_sep
            txt_compra_total_sep = rst!Compra_Total_sep
            txt_vta_total_sep = rst!Venta_Total_sep
            txt_tipo_mes_sep = rst!Tipo_Mes_sep
            txt_margen_total_sep = rst!margen_Total_sep
            txt_ano_iva_oct = rst!Ano_Declaracion_Iva_oct
            txt_iva_credito_oct = rst!Iva_Credito_oct
            txt_iva_debito_oct = rst!Iva_Debito_oct
            txt_compra_neta_oct = rst!Compras_Netas_oct
            txt_vta_netas_formales_oct = rst!Ventas_Netas_Formales_oct
            txt_vta_netas_informales_oct = rst!Ventas_Netas_Informales_oct
            txt_compra_total_oct = rst!Compra_Total_oct
            txt_vta_total_oct = rst!Venta_Total_oct
            txt_tipo_mes_oct = rst!Tipo_Mes_oct
            txt_margen_total_oct = rst!margen_Total_oct
            txt_ano_iva_nov = rst!Ano_Declaracion_Iva_nov
            txt_iva_credito_nov = rst!Iva_Credito_nov
            txt_iva_debito_nov = rst!Iva_Debito_nov
            txt_compra_neta_nov = rst!Compras_Netas_nov
            txt_vta_netas_formales_nov = rst!Ventas_Netas_Formales_nov
            txt_vta_netas_informales_nov = rst!Ventas_Netas_Informales_nov
            txt_compra_total_nov = rst!Compra_Total_nov
            txt_vta_total_nov = rst!Venta_Total_nov
            txt_tipo_mes_nov = rst!Tipo_Mes_nov
            txt_margen_total_nov = rst!margen_Total_nov
            txt_ano_iva_dic = rst!Ano_Declaracion_Iva_dic
            txt_iva_credito_dic = rst!Iva_Credito_dic
            txt_iva_debito_dic = rst!Iva_Debito_dic
            txt_compra_neta_dic = rst!Compras_Netas_dic
            txt_vta_netas_formales_dic = rst!Ventas_Netas_Formales_dic
            txt_vta_netas_informales_dic = rst!Ventas_Netas_Informales_dic
            txt_compra_total_dic = rst!Compra_Total_dic
            txt_vta_total_dic = rst!Venta_Total_dic
            txt_tipo_mes_dic = rst!Tipo_Mes_dic
            txt_margen_total_dic = rst!margen_Total_dic
            txt_fecha_actual = rst!FECHA_INGRESO
            txt_hora_actual = rst!HORA_INGRESO
            cbx_mes_inicio_iva = rst!mes_inicio_iva
            
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

Private Sub txt_iva_credito_ene_Change()
txt_iva_credito_ene = Format(txt_iva_credito_ene, "##,##")
End Sub

Private Sub txt_iva_credito_feb_Change()
txt_iva_credito_feb = Format(txt_iva_credito_feb, "##,##")
End Sub

Private Sub txt_iva_credito_mar_Change()
txt_iva_credito_mar = Format(txt_iva_credito_mar, "##,##")
End Sub
Private Sub txt_iva_credito_abr_Change()
txt_iva_credito_abr = Format(txt_iva_credito_abr, "##,##")
End Sub
Private Sub txt_iva_credito_may_Change()
txt_iva_credito_may = Format(txt_iva_credito_may, "##,##")
End Sub
Private Sub txt_iva_credito_jun_Change()
txt_iva_credito_jun = Format(txt_iva_credito_jun, "##,##")
End Sub
Private Sub txt_iva_credito_jul_Change()
txt_iva_credito_jul = Format(txt_iva_credito_jul, "##,##")
End Sub
Private Sub txt_iva_credito_ago_Change()
txt_iva_credito_ago = Format(txt_iva_credito_ago, "##,##")
End Sub
Private Sub txt_iva_credito_sep_Change()
txt_iva_credito_sep = Format(txt_iva_credito_sep, "##,##")
End Sub
Private Sub txt_iva_credito_oct_Change()
txt_iva_credito_oct = Format(txt_iva_credito_oct, "##,##")
End Sub
Private Sub txt_iva_credito_nov_Change()
txt_iva_credito_nov = Format(txt_iva_credito_nov, "##,##")
End Sub
Private Sub txt_iva_credito_dic_Change()
txt_iva_credito_dic = Format(txt_iva_credito_dic, "##,##")
End Sub

Private Sub txt_iva_debito_ene_Change()
txt_iva_debito_ene = Format(txt_iva_debito_ene, "##,##")
End Sub
Private Sub txt_iva_debito_feb_Change()
txt_iva_debito_feb = Format(txt_iva_debito_feb, "##,##")
End Sub
Private Sub txt_iva_debito_mar_Change()
txt_iva_debito_mar = Format(txt_iva_debito_mar, "##,##")
End Sub
Private Sub txt_iva_debito_abr_Change()
txt_iva_debito_abr = Format(txt_iva_debito_abr, "##,##")
End Sub
Private Sub txt_iva_debito_may_Change()
txt_iva_debito_may = Format(txt_iva_debito_may, "##,##")
End Sub
Private Sub txt_iva_debito_jun_Change()
txt_iva_debito_jun = Format(txt_iva_debito_jun, "##,##")
End Sub
Private Sub txt_iva_debito_jul_Change()
txt_iva_debito_jul = Format(txt_iva_debito_jul, "##,##")
End Sub
Private Sub txt_iva_debito_ago_Change()
txt_iva_debito_ago = Format(txt_iva_debito_ago, "##,##")
End Sub
Private Sub txt_iva_debito_sep_Change()
txt_iva_debito_sep = Format(txt_iva_debito_sep, "##,##")
End Sub
Private Sub txt_iva_debito_oct_Change()
txt_iva_debito_oct = Format(txt_iva_debito_oct, "##,##")
End Sub
Private Sub txt_iva_debito_nov_Change()
txt_iva_debito_nov = Format(txt_iva_debito_nov, "##,##")
End Sub
Private Sub txt_iva_debito_dic_Change()
txt_iva_debito_dic = Format(txt_iva_debito_dic, "##,##")
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

Private Sub txt_margen_total_ene_Change()
txt_margen_total_ene = Format(txt_margen_total_ene, "##,##")
End Sub
Private Sub txt_margen_total_feb_Change()
txt_margen_total_feb = Format(txt_margen_total_feb, "##,##")
End Sub
Private Sub txt_margen_total_mar_Change()
txt_margen_total_mar = Format(txt_margen_total_mar, "##,##")
End Sub
Private Sub txt_margen_total_abr_Change()
txt_margen_total_abr = Format(txt_margen_total_abr, "##,##")
End Sub
Private Sub txt_margen_total_may_Change()
txt_margen_total_may = Format(txt_margen_total_may, "##,##")
End Sub
Private Sub txt_margen_total_jun_Change()
txt_margen_total_jun = Format(txt_margen_total_jun, "##,##")
End Sub
Private Sub txt_margen_total_jul_Change()
txt_margen_total_jul = Format(txt_margen_total_jul, "##,##")
End Sub
Private Sub txt_margen_total_ago_Change()
txt_margen_total_ago = Format(txt_margen_total_ago, "##,##")
End Sub
Private Sub txt_margen_total_sep_Change()
txt_margen_total_sep = Format(txt_margen_total_sep, "##,##")
End Sub
Private Sub txt_margen_total_oct_Change()
txt_margen_total_oct = Format(txt_margen_total_oct, "##,##")
End Sub
Private Sub txt_margen_total_nov_Change()
txt_margen_total_nov = Format(txt_margen_total_nov, "##,##")
End Sub
Private Sub txt_margen_total_dic_Change()
txt_margen_total_dic = Format(txt_margen_total_dic, "##,##")
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
txt_n_grupo_familiar = Format(txt_n_grupo_familiar, "##,##")
End Sub

Private Sub txt_neumaticos_Change()
txt_neumaticos = Format(txt_neumaticos, "##,##")
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

Private Sub txt_prom_vta_meses_altos_Change()
txt_prom_vta_meses_altos = Format(txt_prom_vta_meses_altos, "##,##")
End Sub

Private Sub txt_prom_vta_meses_bajos_Change()
txt_prom_vta_meses_bajos = Format(txt_prom_vta_meses_bajos, "##,##")
End Sub

Private Sub txt_prom_vta_meses_medios_Change()
txt_prom_vta_meses_medios = Format(txt_prom_vta_meses_medios, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_altos_formal_Change()
txt_prom_vtas_meses_altos_formal = Format(txt_prom_vtas_meses_altos_formal, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_altos_informal_Change()
txt_prom_vtas_meses_altos_informal = Format(txt_prom_vtas_meses_altos_informal, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_bajos_formal_Change()
txt_prom_vtas_meses_bajos_formal = Format(txt_prom_vtas_meses_bajos_formal, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_bajos_informal_Change()
txt_prom_vtas_meses_bajos_informal = Format(txt_prom_vtas_meses_bajos_informal, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_medios_formal_Change()
txt_prom_vtas_meses_medios_formal = Format(txt_prom_vtas_meses_medios_formal, "##,##")
End Sub

Private Sub txt_prom_vtas_meses_medios_informal_Change()
txt_prom_vtas_meses_medios_informal = Format(txt_prom_vtas_meses_medios_informal, "##,##")
End Sub

Private Sub txt_promedio_compra_neta_Change()
txt_promedio_compra_neta = Format(txt_promedio_compra_neta, "##,##")
End Sub

Private Sub txt_promedio_compra_total_Change()
txt_promedio_compra_total = Format(txt_promedio_compra_total, "##,##")
End Sub

Private Sub txt_promedio_iva_credito_Change()
txt_promedio_iva_credito = Format(txt_promedio_iva_credito, "##,##")
End Sub

Private Sub txt_promedio_iva_debito_Change()
txt_promedio_iva_debito = Format(txt_promedio_iva_debito, "##,##")
End Sub

Private Sub txt_promedio_vta_netas_formales_Change()
txt_promedio_vta_netas_formales = Format(txt_promedio_vta_netas_formales, "##,##")
End Sub

Private Sub txt_promedio_vta_netas_informales_Change()
txt_promedio_vta_netas_informales = Format(txt_promedio_vta_netas_informales, "##,##")
End Sub

Private Sub txt_promedio_vta_total_Change()
txt_promedio_vta_total = Format(txt_promedio_vta_total, "##,##")
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

Private Sub txt_total_compra_neta_Change()
txt_total_compra_neta = Format(txt_total_compra_neta, "##,##")
End Sub

Private Sub txt_total_compra_total_Change()
txt_total_compra_total = Format(txt_total_compra_total, "##,##")
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

Private Sub txt_total_iva_credito_Change()
txt_total_iva_credito = Format(txt_total_iva_credito, "##,##")
End Sub

Private Sub txt_total_iva_debito_Change()
txt_total_iva_debito = Format(txt_total_iva_debito, "##,##")
End Sub

Private Sub txt_total_margen_total_Change()
txt_total_margen_total = Format(txt_total_margen_total, "##,##")
End Sub

Private Sub txt_total_otros_ingresos_Change()
txt_total_otros_ingresos = Format(txt_total_otros_ingresos, "##,##")
End Sub

Private Sub txt_total_saldo_pendiente_Change()
txt_total_saldo_pendiente = Format(txt_total_saldo_pendiente, "##,##")
End Sub

Private Sub txt_total_vta_netas_formales_Change()
txt_total_vta_netas_formales = Format(txt_total_vta_netas_formales, "##,##")
End Sub

Private Sub txt_total_vta_netas_informales_Change()
txt_total_vta_netas_informales = Format(txt_total_vta_netas_informales, "##,##")
End Sub

Private Sub txt_total_vta_total_Change()
txt_total_vta_total = Format(txt_total_vta_total, "##,##")
End Sub

Private Sub txt_valor_uf_Change()
txt_valor_uf = Format(txt_valor_uf, "##,##")
End Sub


Private Sub txt_venta_total_promedio_anual_Change()

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

Private Sub txt_vta_netas_formales_ene_Change()
txt_vta_netas_formales_ene = Format(txt_vta_netas_formales_ene, "##,##")
End Sub
Private Sub txt_vta_netas_formales_feb_Change()
txt_vta_netas_formales_feb = Format(txt_vta_netas_formales_feb, "##,##")
End Sub
Private Sub txt_vta_netas_formales_mar_Change()
txt_vta_netas_formales_mar = Format(txt_vta_netas_formales_mar, "##,##")
End Sub
Private Sub txt_vta_netas_formales_abr_Change()
txt_vta_netas_formales_abr = Format(txt_vta_netas_formales_abr, "##,##")
End Sub
Private Sub txt_vta_netas_formales_may_Change()
txt_vta_netas_formales_may = Format(txt_vta_netas_formales_may, "##,##")
End Sub
Private Sub txt_vta_netas_formales_jun_Change()
txt_vta_netas_formales_jun = Format(txt_vta_netas_formales_jun, "##,##")
End Sub
Private Sub txt_vta_netas_formales_jul_Change()
txt_vta_netas_formales_jul = Format(txt_vta_netas_formales_jul, "##,##")
End Sub
Private Sub txt_vta_netas_formales_ago_Change()
txt_vta_netas_formales_ago = Format(txt_vta_netas_formales_ago, "##,##")
End Sub
Private Sub txt_vta_netas_formales_sep_Change()
txt_vta_netas_formales_sep = Format(txt_vta_netas_formales_sep, "##,##")
End Sub
Private Sub txt_vta_netas_formales_oct_Change()
txt_vta_netas_formales_oct = Format(txt_vta_netas_formales_oct, "##,##")
End Sub
Private Sub txt_vta_netas_formales_nov_Change()
txt_vta_netas_formales_nov = Format(txt_vta_netas_formales_nov, "##,##")
End Sub
Private Sub txt_vta_netas_formales_dic_Change()
txt_vta_netas_formales_dic = Format(txt_vta_netas_formales_dic, "##,##")
End Sub

Private Sub txt_vta_netas_informales_ene_Change()
txt_vta_netas_informales_ene = Format(txt_vta_netas_informales_ene, "##,##")
End Sub
Private Sub txt_vta_netas_informales_feb_Change()
txt_vta_netas_informales_feb = Format(txt_vta_netas_informales_feb, "##,##")
End Sub
Private Sub txt_vta_netas_informales_mar_Change()
txt_vta_netas_informales_mar = Format(txt_vta_netas_informales_mar, "##,##")
End Sub
Private Sub txt_vta_netas_informales_abr_Change()
txt_vta_netas_informales_abr = Format(txt_vta_netas_informales_abr, "##,##")
End Sub
Private Sub txt_vta_netas_informales_may_Change()
txt_vta_netas_informales_may = Format(txt_vta_netas_informales_may, "##,##")
End Sub
Private Sub txt_vta_netas_informales_jun_Change()
txt_vta_netas_informales_jun = Format(txt_vta_netas_informales_jun, "##,##")
End Sub
Private Sub txt_vta_netas_informales_jul_Change()
txt_vta_netas_informales_jul = Format(txt_vta_netas_informales_jul, "##,##")
End Sub
Private Sub txt_vta_netas_informales_ago_Change()
txt_vta_netas_informales_ago = Format(txt_vta_netas_informales_ago, "##,##")
End Sub
Private Sub txt_vta_netas_informales_sep_Change()
txt_vta_netas_informales_sep = Format(txt_vta_netas_informales_sep, "##,##")
End Sub
Private Sub txt_vta_netas_informales_oct_Change()
txt_vta_netas_informales_oct = Format(txt_vta_netas_informales_oct, "##,##")
End Sub
Private Sub txt_vta_netas_informales_nov_Change()
txt_vta_netas_informales_nov = Format(txt_vta_netas_informales_nov, "##,##")
End Sub
Private Sub txt_vta_netas_informales_dic_Change()
txt_vta_netas_informales_dic = Format(txt_vta_netas_informales_dic, "##,##")
End Sub

Private Sub txt_vta_total_ene_Change()
txt_vta_total_ene = Format(txt_vta_total_ene, "##,##")
End Sub
Private Sub txt_vta_total_feb_Change()
txt_vta_total_feb = Format(txt_vta_total_feb, "##,##")
End Sub
Private Sub txt_vta_total_mar_Change()
txt_vta_total_mar = Format(txt_vta_total_mar, "##,##")
End Sub
Private Sub txt_vta_total_abr_Change()
txt_vta_total_abr = Format(txt_vta_total_abr, "##,##")
End Sub
Private Sub txt_vta_total_may_Change()
txt_vta_total_may = Format(txt_vta_total_may, "##,##")
End Sub
Private Sub txt_vta_total_jun_Change()
txt_vta_total_jun = Format(txt_vta_total_jun, "##,##")
End Sub
Private Sub txt_vta_total_jul_Change()
txt_vta_total_jul = Format(txt_vta_total_jul, "##,##")
End Sub
Private Sub txt_vta_total_ago_Change()
txt_vta_total_ago = Format(txt_vta_total_ago, "##,##")
End Sub
Private Sub txt_vta_total_sep_Change()
txt_vta_total_sep = Format(txt_vta_total_sep, "##,##")
End Sub
Private Sub txt_vta_total_oct_Change()
txt_vta_total_oct = Format(txt_vta_total_oct, "##,##")
End Sub
Private Sub txt_vta_total_nov_Change()
txt_vta_total_nov = Format(txt_vta_total_nov, "##,##")
End Sub
Private Sub txt_vta_total_dic_Change()
txt_vta_total_dic = Format(txt_vta_total_dic, "##,##")
End Sub


Private Sub UserForm_Click()

End Sub
