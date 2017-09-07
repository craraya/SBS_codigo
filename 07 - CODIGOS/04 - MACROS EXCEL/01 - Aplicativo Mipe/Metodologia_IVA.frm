VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Metodologia_IVA 
   Caption         =   ":::: Metodologia IVA"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   OleObjectBlob   =   "Metodologia_IVA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Metodologia_IVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbx_mes_inicio_iva_Change()

txt_iva_credito_ene.Enabled = True
txt_iva_credito_feb.Enabled = True
txt_iva_credito_mar.Enabled = True
txt_iva_credito_abr.Enabled = True
txt_iva_credito_may.Enabled = True
txt_iva_credito_jun.Enabled = True
txt_iva_credito_jul.Enabled = True
txt_iva_credito_ago.Enabled = True
txt_iva_credito_sep.Enabled = True
txt_iva_credito_oct.Enabled = True
txt_iva_credito_nov.Enabled = True
txt_iva_credito_dic.Enabled = True

txt_iva_debito_ene.Enabled = True
txt_iva_debito_feb.Enabled = True
txt_iva_debito_mar.Enabled = True
txt_iva_debito_abr.Enabled = True
txt_iva_debito_may.Enabled = True
txt_iva_debito_jun.Enabled = True
txt_iva_debito_jul.Enabled = True
txt_iva_debito_ago.Enabled = True
txt_iva_debito_sep.Enabled = True
txt_iva_debito_oct.Enabled = True
txt_iva_debito_nov.Enabled = True
txt_iva_debito_dic.Enabled = True


txt_ano_iva_ene = 2011
txt_ano_iva_feb = 2011
txt_ano_iva_mar = 2011
txt_ano_iva_abr = 2011
txt_ano_iva_may = 2011
txt_ano_iva_jun = 2011
txt_ano_iva_jul = 2011
txt_ano_iva_ago = 2011
txt_ano_iva_sep = 2011
txt_ano_iva_oct = 2011
txt_ano_iva_nov = 2011
txt_ano_iva_dic = 2011


Dim mes
mes = Month(Date)
txt_mes_actual = mes

If cbx_mes_inicio_iva = "Enero" Then
  txt_ano_iva_ene = txt_ano_iva_ene + 1

ElseIf cbx_mes_inicio_iva = "Febrero" Then

  txt_ano_iva_ene = txt_ano_iva_ene + 1
    txt_ano_iva_feb = txt_ano_iva_feb + 1

ElseIf cbx_mes_inicio_iva = "Marzo" Then

  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  
ElseIf cbx_mes_inicio_iva = "Abril" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  
ElseIf cbx_mes_inicio_iva = "Mayo" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  

ElseIf cbx_mes_inicio_iva = "Junio" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1

ElseIf cbx_mes_inicio_iva = "Julio" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  
  
  ElseIf cbx_mes_inicio_iva = "Agosto" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  txt_ano_iva_ago = txt_ano_iva_ago + 1
  
  
  ElseIf cbx_mes_inicio_iva = "Septie" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  txt_ano_iva_ago = txt_ano_iva_ago + 1
  txt_ano_iva_sep = txt_ano_iva_sep + 1
  
  
  ElseIf cbx_mes_inicio_iva = "Octubre" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  txt_ano_iva_ago = txt_ano_iva_ago + 1
  txt_ano_iva_sep = txt_ano_iva_sep + 1
  txt_ano_iva_oct = txt_ano_iva_oct + 1
  
  
  ElseIf cbx_mes_inicio_iva = "Noviem" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  txt_ano_iva_ago = txt_ano_iva_ago + 1
  txt_ano_iva_sep = txt_ano_iva_sep + 1
  txt_ano_iva_oct = txt_ano_iva_oct + 1
  txt_ano_iva_nov = txt_ano_iva_nov + 1
  
  
  ElseIf cbx_mes_inicio_iva = "Diciem" Then
   
  txt_ano_iva_ene = txt_ano_iva_ene + 1
  txt_ano_iva_feb = txt_ano_iva_feb + 1
  txt_ano_iva_mar = txt_ano_iva_mar + 1
  txt_ano_iva_abr = txt_ano_iva_abr + 1
  txt_ano_iva_may = txt_ano_iva_may + 1
  txt_ano_iva_jun = txt_ano_iva_jun + 1
  txt_ano_iva_jul = txt_ano_iva_jul + 1
  txt_ano_iva_ago = txt_ano_iva_ago + 1
  txt_ano_iva_sep = txt_ano_iva_sep + 1
  txt_ano_iva_oct = txt_ano_iva_oct + 1
  txt_ano_iva_nov = txt_ano_iva_nov + 1
  txt_ano_iva_dic = txt_ano_iva_dic + 1

End If
End Sub



Private Sub cmd_calcula_costos_fijos_Click()


If (Ficha_Cliente_Micro.cbx_actividad_economica_formal_servicio = "TRANSPORTE DE CARGA" Or cbx_actividad_economica_formal_servicio = "TRANSPORTE ESCOLAR" Or cbx_actividad_economica_formal_servicio = "TRANSPORTE TURISMO") _
And (txt_lubricantes = 0 Or txt_neumaticos = 0 Or txt_afinamientos = 0 Or txt_patentes_seguros = 0) Then
MsgBox "Falta ingresar datos para la actividad seleccionada", vbCritical

txt_total_costos_fijos = Int((Val(txt_arriendo_micro) + Val(txt_sueldos) + Val(txt_movilizacion) + _
Val(txt_servicios_basicos) + Val(txt_contador) + Val(txt_lubricantes) + _
Val(txt_neumaticos) + Val(txt_afinamientos) + Val(txt_patentes_seguros) + Val(txt_otros_costos_fijos) + Val(txt_impuesto)) * 1.15)

txt_valida_costos_fijos = "ZG"


Else

txt_total_costos_fijos = Int((Val(txt_arriendo_micro) + Val(txt_sueldos) + Val(txt_movilizacion) + _
Val(txt_servicios_basicos) + Val(txt_contador) + Val(txt_lubricantes) + _
Val(txt_neumaticos) + Val(txt_afinamientos) + Val(txt_patentes_seguros) + Val(txt_otros_costos_fijos) + Val(txt_impuesto)) * 1.15)

txt_valida_costos_fijos = "A"


'PRENDE BOTON
cmd_calcula_gastos_familiares.Enabled = True

End If

End Sub

Private Sub cmd_calcula_deudas_Click()

cmd_calcular_flujo_Caja.Enabled = False

If txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And txt_ingreso_cantidad_deudas = 1 And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then
    
    If Val(txt_cuotas_pendientes1) <= 3 Or cbx_prepaga_deuda1 = "Si" Then
 

   'txt_total_saldo_pendiente = Val(txt_saldo_pendiente1)
   txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
    txt_sumar_mto_cuota1 = 0
   
   Else
   txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
   'txt_total_saldo_pendiente = Val(txt_saldo_pendiente1)
       txt_no_sumar_mto_cuota1 = 0
   
End If
End If

'cmd_calcular_flujo_Caja.Enabled = True
'------
If txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And txt_ingreso_cantidad_deudas = 2 And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendienes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then

      If txt_cuotas_pendientes1 <= 3 Or cbx_prepaga_deuda1 = "Si" Then
       
          txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
          txt_sumar_mto_cuota1 = 0
          Else
          txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
          txt_no_sumar_mto_cuota1 = 0
       End If

       If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
         txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
         txt_sumar_mto_cuota2 = 0
           Else
            txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
            txt_no_sumar_mto_cuota2 = 0
         
         End If
    
    


'cmd_calcular_flujo_Caja.Enabled = True

End If


If txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And txt_ingreso_cantidad_deudas = 3 And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendienes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendienes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then

      If txt_cuotas_pendientes1 <= 3 Or cbx_prepaga_deuda1 = "Si" Then
       
              txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_sumar_mto_cuota1 = 0
              Else
              txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_no_sumar_mto_cuota1 = 0
       End If

       
       If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
             txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_sumar_mto_cuota2 = 0
             Else
             txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_no_sumar_mto_cuota2 = 0
          End If
             
        If txt_cuotas_pendientes3 <= 3 Or cbx_prepaga_deuda3 = "Si" Then
             txt_no_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_sumar_mto_cuota3 = 0
             Else
             txt_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_no_sumar_mto_cuota3 = 0
                      
       End If




End If
       
If txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And txt_ingreso_cantidad_deudas = 4 And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendienes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendienes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendienes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then


       If txt_cuotas_pendientes1 <= 3 Or cbx_prepaga_deuda1 = "Si" Then
       
              txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_sumar_mto_cuota1 = 0
              Else
              txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_no_sumar_mto_cuota1 = 0
       End If

       
       If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
             txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_sumar_mto_cuota2 = 0
             Else
             txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_no_sumar_mto_cuota2 = 0
         
       End If


        If txt_cuotas_pendientes3 <= 3 Or cbx_prepaga_deuda3 = "Si" Then
             txt_no_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_sumar_mto_cuota3 = 0
             Else
             txt_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_no_sumar_mto_cuota3 = 0
         
       End If
       
               If txt_cuotas_pendientes4 <= 3 Or cbx_prepaga_deuda4 = "Si" Then
             txt_no_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_sumar_mto_cuota4 = 0
             Else
             txt_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_no_sumar_mto_cuota4 = 0
         
       End If

End If

             
If txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And txt_ingreso_cantidad_deudas = 5 And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendienes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendienes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendienes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendienes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then

              
       If txt_cuotas_pendientes1 <= 3 Or cbx_prepaga_deuda1 = "Si" Then
       
              txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_sumar_mto_cuota1 = 0
              Else
              txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_no_sumar_mto_cuota1 = 0
       End If

       
       If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
             txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_sumar_mto_cuota2 = 0
             Else
             txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_no_sumar_mto_cuota2 = 0
         
       End If


        If txt_cuotas_pendientes3 <= 3 Or cbx_prepaga_deuda3 = "Si" Then
             txt_no_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_sumar_mto_cuota3 = 0
             Else
             txt_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_no_sumar_mto_cuota3 = 0
         
       End If
       
        If txt_cuotas_pendientes4 <= 3 Or cbx_prepaga_deuda4 = "Si" Then
             txt_no_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_sumar_mto_cuota4 = 0
             Else
             txt_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_no_sumar_mto_cuota4 = 0
         
        End If
       
         If txt_cuotas_pendientes5 <= 3 Or cbx_prepaga_deuda5 = "Si" Then
             txt_no_sumar_mto_cuota5 = Val(txt_monto_cuota5)
             txt_sumar_mto_cuota5 = 0
             Else
             txt_sumar_mto_cuota5 = Val(txt_monto_cuota5)
             txt_no_sumar_mto_cuota5 = 0
         
       End If
      

End If

If txt_monto_cuota6 <> 0 And txt_monto_cuota6 <> "" And txt_ingreso_cantidad_deudas = 6 And Val(txt_cuotas_pactadas6) >= Val(txt_cuotas_pendienes6) _
And txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendienes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendienes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendienes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendienes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendienes1) Then


       If txt_cuotas_pendientes1 <= 3 Or cbx_prepaga_deuda1 = "Si" Then
       
              txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_sumar_mto_cuota1 = 0
              Else
              txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
              txt_no_sumar_mto_cuota1 = 0
       End If

       
       If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
             txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_sumar_mto_cuota2 = 0
             Else
             txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
             txt_no_sumar_mto_cuota2 = 0
         
       End If


        If txt_cuotas_pendientes3 <= 3 Or cbx_prepaga_deuda3 = "Si" Then
             txt_no_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_sumar_mto_cuota3 = 0
             Else
             txt_sumar_mto_cuota3 = Val(txt_monto_cuota3)
             txt_no_sumar_mto_cuota3 = 0
         
       End If
       
       If txt_cuotas_pendientes4 <= 3 Or cbx_prepaga_deuda4 = "Si" Then
             txt_no_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_sumar_mto_cuota4 = 0
             Else
             txt_sumar_mto_cuota4 = Val(txt_monto_cuota4)
             txt_no_sumar_mto_cuota4 = 0
         
       End If
       
        If txt_cuotas_pendientes5 <= 3 Or cbx_prepaga_deuda5 = "Si" Then
             txt_no_sumar_mto_cuota5 = Val(txt_monto_cuota5)
             txt_sumar_mto_cuota5 = 0
             Else
             txt_sumar_mto_cuota5 = Val(txt_monto_cuota5)
             txt_no_sumar_mto_cuota5 = 0
         
       End If

       
        If txt_cuotas_pendientes6 <= 3 Or cbx_prepaga_deuda6 = "Si" Then
             txt_no_sumar_mto_cuota6 = Val(txt_monto_cuota6)
             txt_sumar_mto_cuota6 = 0
             Else
             txt_sumar_mto_cuota6 = Val(txt_monto_cuota6)
             txt_no_sumar_mto_cuota6 = 0
         
       End If
End If



txt_total_saldo_pendiente = Val(txt_saldo_pendiente1) + Val(txt_saldo_pendiente2) + Val(txt_saldo_pendiente3) + Val(txt_saldo_pendiente4) + _
Val(txt_saldo_pendiente5) + Val(txt_saldo_pendiente6)

txt_total_deudas = Val(txt_sumar_mto_cuota1) + Val(txt_sumar_mto_cuota2) + Val(txt_sumar_mto_cuota3) + Val(txt_sumar_mto_cuota4) + Val(txt_sumar_mto_cuota5) _
+ Val(txt_sumar_mto_cuota6)


'''''COMPARA DEUDA DECLARADA CONTRA DEUDA SBIF (VIGENTE+MOROSA+VENCIDA+CASTIGO)

If txt_total_saldo_pendiente < txt_total_deuda_d10 Then
    txt_r_sbif_declarada = "ZG"
    Else
    txt_r_sbif_declarada = "A"
End


cmd_calcular_flujo_Caja.Enabled = True

End Sub

Private Sub cmd_calcula_gastos_familiares_Click()

cmd_calcula_otros_ingresos.Enabled = False

If txt_gastos_indicado_cliente <> "" And txt_gastos_indicado_cliente <> 0 Then

txt_total_gasto_familiar = Int((((txt_valor_uf * 6.5) + (15000 * txt_n_grupo_familiar))) + Val(txt_arriendo_vivienda))

txt_gasto_calc_sistema = txt_total_gasto_familiar
txt_mayor_gasto_familiar = txt_gastos_indicado_cliente

'prende boton
cmd_calcula_otros_ingresos.Enabled = True

If Val(txt_total_gasto_familiar) > Val(txt_gastos_indicado_cliente) Then
txt_total_gasto_familiar = txt_total_gasto_familiar

'prende boton
cmd_calcula_otros_ingresos.Enabled = True

Else
txt_total_gasto_familiar = Val(txt_gastos_indicado_cliente)

'prende boton
cmd_calcula_otros_ingresos.Enabled = True

End If

Else
MsgBox "Debe Ingresar Los Datos Tanto Para Gastos Familiares Como para Los Indicados por El Cliente"
End If

End Sub

Private Sub cmd_calcula_otros_ingresos_Click()
txt_total_otros_ingresos = Val(txt_liquidacion_sueldo) + Val(txt_jubilacion) + Val(txt_montepio) + Val(txt_arriendo_vivienda1) + Val(txt_ingreso_segunda_microempresa) + _
Val(txt_boleta_honorario)

'PRENDE BOTON
cmd_calcula_deudas.Enabled = True

End Sub

Private Sub cmd_calcular_flujo_Caja_Click()

'Paso de Meses Evaluados
numero_meses_tipo_mes_alto = txt_tipo_mes_r_suma_alto
numero_meses_tipo_mes_medio = txt_tipo_mes_r_suma_medio
numero_meses_tipo_mes_bajo = txt_tipo_mes_r_suma_bajo

' Paso Vtas Formales Promedio
txt_vta_formal_promedio_mes_alto = txt_prom_vtas_meses_altos_formal
txt_vta_formal_promedio_mes_medio = txt_prom_vtas_meses_medios_formal
txt_vta_formal_promedio_mes_bajo = txt_prom_vtas_meses_bajos_formal
                                   
' Paso Vtas Informales Promedio
txt_vta_informal_promedio_mes_alto = txt_prom_vtas_meses_altos_informal
txt_vta_informal_promedio_mes_medio = txt_prom_vtas_meses_medios_informal
txt_vta_informal_promedio_mes_bajo = txt_prom_vtas_meses_bajos_informal
                                     

'If (txt_cuota_credito <> "" Or txt_cuota_consumo <> "") Then
   
txt_costo_fijo_mes_alto = Val(txt_total_costos_fijos)
txt_costo_fijo_mes_medio = Val(txt_total_costos_fijos)
txt_costo_fijo_mes_bajo = Val(txt_total_costos_fijos)


txt_gastos_familiares_mes_alto = Val(txt_total_gasto_familiar)
txt_gastos_familiares_mes_medio = Val(txt_total_gasto_familiar)
txt_gastos_familiares_mes_bajo = Val(txt_total_gasto_familiar)

txt_otros_ingresos_mes_alto = Val(txt_total_otros_ingresos)
txt_otros_ingresos_mes_medio = Val(txt_total_otros_ingresos)
txt_otros_ingresos_mes_bajo = Val(txt_total_otros_ingresos)

txt_Deudas_flujo_caja_mes_alto = Val(txt_total_deudas)
txt_Deudas_flujo_caja_mes_medio = Val(txt_total_deudas)
txt_Deudas_flujo_caja_mes_bajo = Val(txt_total_deudas)



If (numero_meses_tipo_mes_alto) * 1 + (numero_meses_tipo_mes_medio) * 1 + (numero_meses_tipo_mes_bajo) * 1 > 12 Then
  MsgBox "La Suma Del Ingreso de Meses NO puede ser mayor a 12 ... Revise"
Else

txt_Venta_Total_Promedio_Mes_Alto = Val(txt_vta_formal_promedio_mes_alto) * 1 + Val(txt_vta_informal_promedio_mes_alto) * 1
txt_Venta_Total_Promedio_Mes_Medio = Val(txt_vta_formal_promedio_mes_medio) * 1 + Val(txt_vta_informal_promedio_mes_medio) * 1
txt_Venta_Total_Promedio_Mes_Bajo = Val(txt_vta_formal_promedio_mes_bajo) * 1 + Val(txt_vta_informal_promedio_mes_bajo) * 1

'txt_venta_total_promedio_anual = Int(txt_venta_total_mes_alto_corregida + txt_venta_total_mes_medio_corregida + txt_venta_total_mes_bajo_corregida) / 12

txt_costo_variable_mes_alto = Int(txt_Venta_Total_Promedio_Mes_Alto * txt_Sub_Total_x1)
txt_costo_variable_mes_medio = Int(txt_Venta_Total_Promedio_Mes_Medio * txt_Sub_Total_x1)
txt_costo_variable_mes_bajo = Int(txt_Venta_Total_Promedio_Mes_Bajo * txt_Sub_Total_x1)


txt_resultado_operacional_mes_alto = (txt_Venta_Total_Promedio_Mes_Alto) - (txt_costo_variable_mes_alto) - (txt_costo_fijo_mes_alto)
txt_resultado_operacional_mes_medio = (txt_Venta_Total_Promedio_Mes_Medio) - (txt_costo_variable_mes_medio) - (txt_costo_fijo_mes_medio)
txt_resultado_operacional_mes_bajo = (txt_Venta_Total_Promedio_Mes_Bajo) - (txt_costo_variable_mes_bajo) - (txt_costo_fijo_mes_bajo)

txt_capacidad_pago_mes_alto = (txt_resultado_operacional_mes_alto) * 1 + (txt_otros_ingresos_mes_alto) * 1 + (txt_segunda_microempresa_mes_alto) * 1 - (txt_Deudas_flujo_caja_mes_alto) * 1 - (txt_gastos_familiares_mes_alto) * 1
txt_capacidad_pago_mes_medio = (txt_resultado_operacional_mes_medio) * 1 + (txt_otros_ingresos_mes_medio) * 1 + (txt_segunda_microempresa_mes_medio) * 1 - (txt_Deudas_flujo_caja_mes_medio) * 1 - (txt_gastos_familiares_mes_medio) * 1
txt_capacidad_pago_mes_bajo = (txt_resultado_operacional_mes_bajo) * 1 + (txt_otros_ingresos_mes_bajo) * 1 + (txt_segunda_microempresa_mes_bajo) * 1 - (txt_Deudas_flujo_caja_mes_bajo) * 1 - (txt_gastos_familiares_mes_bajo) * 1


'Calculo De Factor Correccion
'------------------------------
'txt_tipo_cliente_form_evaluacion //// txt_tipo_riesgo_form_evaluacion

If txt_tipo_cliente_form_evaluacion = "Antiguo Prime" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then
   'txt_factor = 1
   txt_factor_consumo = 0.75
   'txt_leverage = 9
   
ElseIf txt_tipo_cliente_form_evaluacion = "Antiguo Prime" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.9
   txt_factor_consumo = 0.55
   'txt_leverage = 8

ElseIf txt_tipo_cliente_form_evaluacion = "Antiguo Prime" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.8
   txt_factor_consumo = 0
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Antiguo No Prime" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.9
   txt_factor_consumo = 0.55
   'txt_leverage = 8
   
ElseIf txt_tipo_cliente_form_evaluacion = "Antiguo No Prime" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.8
   txt_factor_consumo = 0.35
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Antiguo No Prime" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.8
   txt_factor_consumo = 0
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6
   
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5


ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo No Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6

ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo No Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5
   
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then
'ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo No Bancarizado" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.5
   txt_factor_consumo = 0
   'txt_leverage = 4


'FIN DE CALCULO

End If

'####################################################################
'FACTOR Y LEVERAGE DESDE RIESGO
    txt_factor = Evaluacion_Perfil.txt_tdsr
    txt_leverage = Evaluacion_Perfil.txt_leverage
'####################################################################


txt_capacidad_pago_corregida_ajustada_mes_alto = Int(txt_capacidad_pago_mes_alto * txt_factor)
txt_capacidad_pago_corregida_ajustada_mes_medio = Int(txt_capacidad_pago_mes_medio * txt_factor)
txt_capacidad_pago_corregida_ajustada_mes_bajo = Int(txt_capacidad_pago_mes_bajo * txt_factor)

txt_capacidad_pago_corregida_consumo_mes_alto = Int(txt_capacidad_pago_mes_alto * txt_factor_consumo)
txt_capacidad_pago_corregida_consumo_mes_medio = Int(txt_capacidad_pago_mes_medio * txt_factor_consumo)
txt_capacidad_pago_corregida_consumo_mes_bajo = Int(txt_capacidad_pago_mes_bajo * txt_factor_consumo)

txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 3)
txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 3)

txt_monto_maximo_credito = Int(txt_capacidad_pago_promedio_corregida_ajustada * txt_leverage)
txt_monto_maximo_consumo = Int(txt_capacidad_pago_promedio_corregida_consumo * txt_leverage)

txt_costo_variable_mes_alto = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Alto)
txt_costo_variable_mes_medio = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Medio)
txt_costo_variable_mes_bajo = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Bajo)


 
End If

cmd_calcular_flujo_Caja.Enabled = True
cmd_calcular_resolucion_cred.Enabled = True

'Else
'MsgBox "Debe Ingresar Los Datos Obligatorios para comenzar Calculo"
'End If


End Sub

Private Sub cmd_calcular_resolucion_cred_Click()

If txt_cuota_credito > 0 And txt_cuota_credito <> "" And txt_mto_bruto_sol_cliente > 0 And txt_mto_bruto_sol_cliente <> "" Then

If Val(txt_mto_bruto_sol_cliente) * 1 > 0 And Val(txt_cuota_credito) * 1 > 0 _
   And Val(txt_cuota_credito) * 1 <= Val(txt_capacidad_pago_promedio_corregida_ajustada) * 1 Then
    
    txt_aprobacion = "APROBADO"
    
    Else
     txt_aprobacion = "RECHAZADO"
End If

   
If Val(txt_mto_bruto_sol_cliente) * 1 > 0 And Val(txt_cuota_credito) * 1 > 0 _
       And Val(txt_mto_bruto_sol_cliente) * 1 <= Val(txt_monto_maximo_credito) * 1 Then
   
    txt_resolucion_credito_por_cuota = "APROBADO"
   
   Else
    
       txt_resolucion_credito_por_cuota = "RECHAZADO"
       
   End If

Else
  MsgBox ("Antes de Calcular Debes Ingresar Valores En Campos Correspondientes")
  
   End If
End Sub

Private Sub cmd_calcular_ventas_iva_Click()

LBL_ALARMA_PORCENTAJE_COMPRA_FORMAL.Visible = False

If txt_iva_credito_mar <> "" And txt_iva_credito_abr <> "" And txt_iva_credito_may <> "" And txt_iva_credito_jun <> "" _
And txt_iva_credito_jul <> "" And txt_iva_credito_ago <> "" And txt_iva_credito_sep <> "" And txt_iva_credito_oct <> "" _
And txt_iva_credito_nov <> "" And txt_iva_credito_dic <> "" And txt_iva_credito_ene <> "" And txt_iva_credito_feb <> "" _
And txt_iva_credito_mar <> 0 And txt_iva_credito_abr <> 0 And txt_iva_credito_may <> 0 And txt_iva_credito_jun <> 0 _
And txt_iva_credito_jul <> 0 And txt_iva_credito_ago <> 0 And txt_iva_credito_sep <> 0 And txt_iva_credito_oct <> 0 _
And txt_iva_credito_nov <> 0 And txt_iva_credito_dic <> 0 And txt_iva_credito_ene <> 0 And txt_iva_credito_feb <> 0 _
And txt_iva_debito_mar <> "" And txt_iva_debito_abr <> "" And txt_iva_debito_may <> "" And txt_iva_debito_jun <> "" _
And txt_iva_debito_jul <> "" And txt_iva_debito_ago <> "" And txt_iva_debito_sep <> "" And txt_iva_debito_oct <> "" _
And txt_iva_debito_nov <> "" And txt_iva_debito_dic <> "" And txt_iva_debito_ene <> "" And txt_iva_debito_feb <> "" _
And txt_iva_debito_mar <> 0 And txt_iva_debito_abr <> 0 And txt_iva_debito_may <> 0 And txt_iva_debito_jun <> 0 _
And txt_iva_debito_jul <> 0 And txt_iva_debito_ago <> 0 And txt_iva_debito_sep <> 0 And txt_iva_debito_oct <> 0 _
And txt_iva_debito_nov <> 0 And txt_iva_debito_dic <> 0 And txt_iva_debito_ene <> 0 And txt_iva_debito_feb <> 0 _
And txt_compra_promedio_mensual <> 0 And txt_compra_promedio_mensual <> "" And txt_veces_compra_mes <> 0 _
And txt_veces_compra_mes <> "" Then


' CALCULO DE IVA CREDITO
txt_compra_neta_mar = Int(txt_iva_credito_mar / txt_019)
txt_compra_neta_abr = Int(txt_iva_credito_abr / txt_019)
txt_compra_neta_may = Int(txt_iva_credito_may / txt_019)
txt_compra_neta_jun = Int(txt_iva_credito_jun / txt_019)
txt_compra_neta_jul = Int(txt_iva_credito_jul / txt_019)
txt_compra_neta_ago = Int(txt_iva_credito_ago / txt_019)
txt_compra_neta_sep = Int(txt_iva_credito_sep / txt_019)
txt_compra_neta_oct = Int(txt_iva_credito_oct / txt_019)
txt_compra_neta_nov = Int(txt_iva_credito_nov / txt_019)
txt_compra_neta_dic = Int(txt_iva_credito_dic / txt_019)
txt_compra_neta_ene = Int(txt_iva_credito_ene / txt_019)
txt_compra_neta_feb = Int(txt_iva_credito_feb / txt_019)


'CALCULO DE IVA DEBITO
txt_vta_netas_formales_mar = Int(txt_iva_debito_mar / txt_019)
txt_vta_netas_formales_abr = Int(txt_iva_debito_abr / txt_019)
txt_vta_netas_formales_may = Int(txt_iva_debito_may / txt_019)
txt_vta_netas_formales_jun = Int(txt_iva_debito_jun / txt_019)
txt_vta_netas_formales_jul = Int(txt_iva_debito_jul / txt_019)
txt_vta_netas_formales_ago = Int(txt_iva_debito_ago / txt_019)
txt_vta_netas_formales_sep = Int(txt_iva_debito_sep / txt_019)
txt_vta_netas_formales_oct = Int(txt_iva_debito_oct / txt_019)
txt_vta_netas_formales_nov = Int(txt_iva_debito_nov / txt_019)
txt_vta_netas_formales_dic = Int(txt_iva_debito_dic / txt_019)
txt_vta_netas_formales_ene = Int(txt_iva_debito_ene / txt_019)
txt_vta_netas_formales_feb = Int(txt_iva_debito_feb / txt_019)


' SUMA DE VENTA TOTAL IVA CREDITO
txt_total_iva_credito = (txt_iva_credito_mar) * 1 + (txt_iva_credito_abr) * 1 + (txt_iva_credito_may) * 1 + _
(txt_iva_credito_jun) * 1 + (txt_iva_credito_jul) * 1 + (txt_iva_credito_ago) * 1 + (txt_iva_credito_sep) * 1 _
+ (txt_iva_credito_oct) * 1 + (txt_iva_credito_nov) * 1 + (txt_iva_credito_dic) * 1 + (txt_iva_credito_ene) * 1 _
+ (txt_iva_credito_feb) * 1

' SUMA DE VENTA TOTAL IVA DEBITO
txt_total_iva_debito = (txt_iva_debito_mar) * 1 + (txt_iva_debito_abr) * 1 + (txt_iva_debito_may) * 1 + (txt_iva_debito_jun) * 1 + (txt_iva_debito_jul) * 1 + _
(txt_iva_debito_ago) * 1 + (txt_iva_debito_sep) * 1 + (txt_iva_debito_oct) * 1 + (txt_iva_debito_nov) * 1 + (txt_iva_debito_dic) * 1 + _
(txt_iva_debito_ene) * 1 + (txt_iva_debito_feb) * 1

'PROMEDIO IVA CREDITO / IVA DEBITO
txt_promedio_iva_credito = Int(txt_total_iva_credito / 12)
txt_promedio_iva_debito = Int(txt_total_iva_debito / 12)

'SUMA DE COMPRA_NETA
txt_total_compra_neta = (txt_compra_neta_mar) * 1 + (txt_compra_neta_abr) * 1 + (txt_compra_neta_may) * 1 + (txt_compra_neta_jun) * 1 + (txt_compra_neta_jul) * 1 + _
(txt_compra_neta_ago) * 1 + (txt_compra_neta_sep) * 1 + (txt_compra_neta_oct) * 1 + (txt_compra_neta_nov) * 1 + (txt_compra_neta_dic) * 1 + (txt_compra_neta_ene) * 1 + _
(txt_compra_neta_feb) * 1

txt_promedio_compra_neta = Int(txt_total_compra_neta / 12)

'txt_compra_promedio_mensual = 1



'SUMA DE VENTAS_NETAS_FORMALES
txt_total_vta_netas_formales = (txt_vta_netas_formales_mar) * 1 + (txt_vta_netas_formales_abr) * 1 + (txt_vta_netas_formales_may) * 1 + _
(txt_vta_netas_formales_jun) * 1 + (txt_vta_netas_formales_jul) * 1 + (txt_vta_netas_formales_ago) * 1 + (txt_vta_netas_formales_sep) * 1 + _
(txt_vta_netas_formales_oct) * 1 + (txt_vta_netas_formales_nov) * 1 + (txt_vta_netas_formales_dic) * 1 + (txt_vta_netas_formales_ene) * 1 + _
(txt_vta_netas_formales_feb) * 1

'PROMEDIO VTAS NETAS FORMALES
txt_promedio_vta_netas_formales = Int(txt_total_vta_netas_formales / 12)

'PORCENTAJE COMPRA FORMAL
txt_porcentaje_compra_formal = Round((txt_promedio_compra_neta / txt_compra_promedio_mensual), 2)

'CALCULAR COMPRA_TOTAL
txt_compra_total_mar = Int(txt_compra_neta_mar / txt_porcentaje_compra_formal)
txt_compra_total_abr = Int(txt_compra_neta_abr / txt_porcentaje_compra_formal)
txt_compra_total_may = Int(txt_compra_neta_may / txt_porcentaje_compra_formal)
txt_compra_total_jun = Int(txt_compra_neta_jun / txt_porcentaje_compra_formal)
txt_compra_total_jul = Int(txt_compra_neta_jul / txt_porcentaje_compra_formal)
txt_compra_total_ago = Int(txt_compra_neta_ago / txt_porcentaje_compra_formal)
txt_compra_total_sep = Int(txt_compra_neta_sep / txt_porcentaje_compra_formal)
txt_compra_total_oct = Int(txt_compra_neta_oct / txt_porcentaje_compra_formal)
txt_compra_total_nov = Int(txt_compra_neta_nov / txt_porcentaje_compra_formal)
txt_compra_total_dic = Int(txt_compra_neta_dic / txt_porcentaje_compra_formal)
txt_compra_total_ene = Int(txt_compra_neta_ene / txt_porcentaje_compra_formal)
txt_compra_total_feb = Int(txt_compra_neta_feb / txt_porcentaje_compra_formal)

'suma de compra total

txt_total_compra_total = (txt_compra_total_mar) * 1 + (txt_compra_total_abr) * 1 + (txt_compra_total_may) * 1 + (txt_compra_total_jun) * 1 + _
(txt_compra_total_jul) * 1 + (txt_compra_total_ago) * 1 + (txt_compra_total_sep) * 1 + (txt_compra_total_oct) * 1 + _
(txt_compra_total_nov) * 1 + (txt_compra_total_dic) * 1 + (txt_compra_total_ene) * 1 + (txt_compra_total_feb) * 1

'promedio compra total

txt_promedio_compra_total = Int(txt_total_compra_total / 12)

'CALCULA VENTA TOTAL
'Evaluacion_Perfil.txt_registro_ventas_var
If (txt_porcentaje_compra_formal) > (Evaluacion_Perfil.txt_registro_ventas_var) Then
txt_vta_total_mar = Int(txt_compra_total_mar / txt_Sub_Total_x1)
txt_vta_total_abr = Int(txt_compra_total_abr / txt_Sub_Total_x1)
txt_vta_total_may = Int(txt_compra_total_may / txt_Sub_Total_x1)
txt_vta_total_jun = Int(txt_compra_total_jun / txt_Sub_Total_x1)
txt_vta_total_jul = Int(txt_compra_total_jul / txt_Sub_Total_x1)
txt_vta_total_ago = Int(txt_compra_total_ago / txt_Sub_Total_x1)
txt_vta_total_sep = Int(txt_compra_total_sep / txt_Sub_Total_x1)
txt_vta_total_oct = Int(txt_compra_total_oct / txt_Sub_Total_x1)
txt_vta_total_nov = Int(txt_compra_total_nov / txt_Sub_Total_x1)
txt_vta_total_dic = Int(txt_compra_total_dic / txt_Sub_Total_x1)
txt_vta_total_ene = Int(txt_compra_total_ene / txt_Sub_Total_x1)
txt_vta_total_feb = Int(txt_compra_total_feb / txt_Sub_Total_x1)
Else
txt_vta_total_mar = Int(txt_vta_netas_formales_mar / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_abr = Int(txt_vta_netas_formales_abr / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_may = Int(txt_vta_netas_formales_may / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_jun = Int(txt_vta_netas_formales_jun / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_jul = Int(txt_vta_netas_formales_jul / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_ago = Int(txt_vta_netas_formales_ago / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_sep = Int(txt_vta_netas_formales_sep / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_oct = Int(txt_vta_netas_formales_oct / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_nov = Int(txt_vta_netas_formales_nov / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_dic = Int(txt_vta_netas_formales_dic / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_ene = Int(txt_vta_netas_formales_ene / Evaluacion_Perfil.txt_registro_ventas_var)
txt_vta_total_feb = Int(txt_vta_netas_formales_feb / Evaluacion_Perfil.txt_registro_ventas_var)
End If

'SUMA DE VENTA TOTAL

txt_total_vta_total = (txt_vta_total_mar) * 1 + (txt_vta_total_abr) * 1 + (txt_vta_total_may) * 1 + (txt_vta_total_jun) * 1 + (txt_vta_total_jul) * 1 + _
(txt_vta_total_ago) * 1 + (txt_vta_total_sep) * 1 + (txt_vta_total_oct) * 1 + (txt_vta_total_nov) * 1 + (txt_vta_total_dic) * 1 + _
(txt_vta_total_ene) * 1 + (txt_vta_total_feb) * 1

' PROMEDIO DE VENTAS TOTAL

txt_promedio_vta_total = Int(txt_total_vta_total / 12)

' CALCULO DE VENTA INFORMAL
txt_vta_netas_informales_mar = (txt_vta_total_mar - txt_vta_netas_formales_mar)
txt_vta_netas_informales_abr = (txt_vta_total_abr - txt_vta_netas_formales_abr)
txt_vta_netas_informales_may = (txt_vta_total_may - txt_vta_netas_formales_may)
txt_vta_netas_informales_jun = (txt_vta_total_jun - txt_vta_netas_formales_jun)
txt_vta_netas_informales_jul = (txt_vta_total_jul - txt_vta_netas_formales_jul)
txt_vta_netas_informales_ago = (txt_vta_total_ago - txt_vta_netas_formales_ago)
txt_vta_netas_informales_sep = (txt_vta_total_sep - txt_vta_netas_formales_sep)
txt_vta_netas_informales_oct = (txt_vta_total_oct - txt_vta_netas_formales_oct)
txt_vta_netas_informales_nov = (txt_vta_total_nov - txt_vta_netas_formales_nov)
txt_vta_netas_informales_dic = (txt_vta_total_dic - txt_vta_netas_formales_dic)
txt_vta_netas_informales_ene = (txt_vta_total_ene - txt_vta_netas_formales_ene)
txt_vta_netas_informales_feb = (txt_vta_total_feb - txt_vta_netas_formales_feb)

'suma de ventas informales
txt_total_vta_netas_informales = (txt_vta_netas_informales_mar) * 1 + (txt_vta_netas_informales_abr) * 1 + (txt_vta_netas_informales_may) * 1 + _
(txt_vta_netas_informales_jun) * 1 + (txt_vta_netas_informales_jul) * 1 + (txt_vta_netas_informales_ago) * 1 + (txt_vta_netas_informales_sep) * 1 + _
(txt_vta_netas_informales_oct) * 1 + (txt_vta_netas_informales_nov) * 1 + (txt_vta_netas_informales_dic) * 1 + (txt_vta_netas_informales_ene) * 1 + _
(txt_vta_netas_informales_feb) * 1

'calcular promedio vtas netas informales
txt_promedio_vta_netas_informales = Int(txt_total_vta_netas_informales / 12)
 
'calcular margen total
txt_margen_total_mar = txt_vta_total_mar - txt_compra_total_mar
txt_margen_total_abr = txt_vta_total_abr - txt_compra_total_abr
txt_margen_total_may = txt_vta_total_may - txt_compra_total_may
txt_margen_total_jun = txt_vta_total_jun - txt_compra_total_jun
txt_margen_total_jul = txt_vta_total_jul - txt_compra_total_jul
txt_margen_total_ago = txt_vta_total_ago - txt_compra_total_ago
txt_margen_total_sep = txt_vta_total_sep - txt_compra_total_sep
txt_margen_total_oct = txt_vta_total_oct - txt_compra_total_oct
txt_margen_total_nov = txt_vta_total_nov - txt_compra_total_nov
txt_margen_total_dic = txt_vta_total_dic - txt_compra_total_dic
txt_margen_total_ene = txt_vta_total_ene - txt_compra_total_ene
txt_margen_total_feb = txt_vta_total_feb - txt_compra_total_feb

'suma margen total
txt_total_margen_total = (txt_margen_total_mar) * 1 + (txt_margen_total_abr) * 1 + (txt_margen_total_may) * 1 + (txt_margen_total_jun) * 1 + (txt_margen_total_jul) * 1 + _
(txt_margen_total_ago) * 1 + (txt_margen_total_sep) * 1 + (txt_margen_total_oct) * 1 + (txt_margen_total_nov) * 1 + (txt_margen_total_dic) * 1 + _
(txt_margen_total_ene) * 1 + (txt_margen_total_feb) * 1


'Calcular Tipo Mes
 
'txt_vta_total_mar = 10
 'txt_prmedio_vta_total = 550


'---- MARZO
 
 txt_tipo_mes_r_suma_bajo = 0
 txt_tipo_mes_r_monto_bajo = 0
 txt_monto_vta_formal_bajo = 0
 txt_monto_vta_informal_bajo = 0
 txt_tipo_mes_r_suma_alto = 0
 txt_tipo_mes_r_monto_alto = 0
 txt_monto_vta_formal_alto = 0
 txt_monto_vta_informal_alto = 0
 txt_tipo_mes_r_suma_medio = 0
 txt_tipo_mes_r_monto_medio = 0
 txt_monto_vta_formal_medio = 0
 txt_monto_vta_informal_medio = 0
 
 
If Val(txt_vta_total_mar) < Val(txt_promedio_vta_total) * 0.9 Then
        txt_tipo_mes_mar = "Bajo"
        txt_tipo_mes_r_suma_bajo = Val(txt_r_mes_mar)
        txt_tipo_mes_r_monto_bajo = Val(txt_vta_total_mar)
        txt_monto_vta_formal_bajo = Val(txt_vta_netas_formales_mar)
        txt_monto_vta_informal_bajo = Val(txt_vta_netas_informales_mar)
        
        
    ElseIf Val(txt_vta_total_mar) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_mar = "Medio"
        txt_tipo_mes_r_suma_medio = Val(txt_r_mes_mar)
        txt_tipo_mes_r_monto_medio = Val(txt_vta_total_mar)
        txt_monto_vta_formal_medio = Val(txt_vta_netas_formales_mar)
        txt_monto_vta_informal_medio = Val(txt_vta_netas_informales_mar)
        
        
    Else
        txt_tipo_mes_mar = "Alto"
        txt_tipo_mes_r_suma_alto = Val(txt_r_mes_mar)
        txt_tipo_mes_r_monto_alto = Val(txt_vta_total_mar)
        txt_monto_vta_formal_alto = Val(txt_vta_netas_formales_mar)
        txt_monto_vta_informal_alto = Val(txt_vta_netas_informales_mar)
        
End If

'---- Abr
    
    If Val(txt_vta_total_abr) < Val(txt_promedio_vta_total) * 0.9 Then
            txt_tipo_mes_abr = "Bajo"
            txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_abr)
            txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_abr)
            txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_abr)
            txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_abr)
            
        
        ElseIf Val(txt_vta_total_abr) < Val(txt_promedio_vta_total * 1.1) Then
            txt_tipo_mes_abr = "Medio"
            txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_abr)
            txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_abr)
            txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_abr)
            txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_abr)
            
        Else
            txt_tipo_mes_abr = "Alto"
            txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_abr)
            txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_abr)
            txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_abr)
            txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_abr)
    End If
    
'---- May
    

    If Val(txt_vta_total_may) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_may = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_may)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_may)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_may)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_may)

        ElseIf Val(txt_vta_total_may) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_may = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_may)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_may)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_may)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_may)
   
        Else
        txt_tipo_mes_may = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_may)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_may)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_may)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_may)
   
    End If
   
   '---- Jun
      
   If Val(txt_vta_total_jun) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_jun = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_jun)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_jun)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_jun)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_jun)

    ElseIf Val(txt_vta_total_jun) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_jun = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_jun)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_jun)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_jun)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_jun)
    
    Else
        txt_tipo_mes_jun = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_jun)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_jun)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_jun)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_jun)
        
    End If
      

'---- Jul

    If Val(txt_vta_total_jul) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_jul = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_jul)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_jul)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_jul)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_jul)

        ElseIf Val(txt_vta_total_jul) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_jul = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_jul)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_jul)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_jul)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_jul)
        
        Else
        txt_tipo_mes_jul = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_jul)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_jul)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_jul)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_jul)
        
    End If

   
   '---- Ago
   
    If Val(txt_vta_total_ago) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_ago = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_ago)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_ago)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_ago)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_ago)

        ElseIf Val(txt_vta_total_ago) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_ago = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_ago)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_ago)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_ago)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_ago)
    
        Else
       txt_tipo_mes_ago = "Alto"
       txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_ago)
       txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_ago)
       txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_ago)
       txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_ago)

    End If
   

'---- Sep

    If Val(txt_vta_total_sep) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_sep = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_sep)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_sep)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_sep)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_sep)
        

    ElseIf Val(txt_vta_total_sep) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_sep = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_sep)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_sep)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_sep)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_sep)
   
   Else
      txt_tipo_mes_sep = "Alto"
      txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_sep)
      txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_sep)
      txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_sep)
      txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_sep)

   End If
   
   
'---- Oct
   
    If Val(txt_vta_total_oct) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_oct = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_oct)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_oct)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_oct)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_oct)

    ElseIf Val(txt_vta_total_oct) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_oct = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_oct)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_oct)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_oct)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_oct)
   
   Else
         txt_tipo_mes_oct = "Alto"
         txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_oct)
         txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_oct)
         txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_oct)
         txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_oct)
         
   End If

'---- Nov
      
      
   If Val(txt_vta_total_nov) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_nov = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_nov)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_nov)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_nov)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_nov)

    ElseIf Val(txt_vta_total_nov) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_nov = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_nov)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_nov)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_nov)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_nov)
   
   Else
        txt_tipo_mes_nov = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_nov)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_nov)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_nov)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_nov)
         
        End If
   

'---- Dic
   
   
   If Val(txt_vta_total_dic) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_dic = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_dic)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_dic)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_dic)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_dic)

   ElseIf Val(txt_vta_total_dic) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_dic = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_dic)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_dic)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_dic)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_dic)
   
   Else
        txt_tipo_mes_dic = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_dic)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_dic)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_dic)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_dic)
   
   End If
   

'---- Enero
   
   If Val(txt_vta_total_ene) < Val(txt_promedio_vta_total * 0.9) Then
        txt_tipo_mes_ene = "Bajo"
        txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_ene)
        txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_ene)
        txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_ene)
        txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_ene)

    ElseIf Val(txt_vta_total_ene) < Val(txt_promedio_vta_total * 1.1) Then
        txt_tipo_mes_ene = "Medio"
        txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_ene)
        txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_ene)
        txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_ene)
        txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_ene)
   
      Else
        txt_tipo_mes_ene = "Alto"
        txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_ene)
        txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_ene)
        txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_ene)
        txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_ene)
   
    End If

'---------Feb

   If Val(txt_vta_total_feb) < Val(txt_promedio_vta_total * 0.9) Then
   txt_tipo_mes_feb = "Bajo"
   txt_tipo_mes_r_suma_bajo = txt_tipo_mes_r_suma_bajo + Val(txt_r_mes_feb)
   txt_tipo_mes_r_monto_bajo = txt_tipo_mes_r_monto_bajo + Val(txt_vta_total_feb)
   txt_monto_vta_formal_bajo = txt_monto_vta_formal_bajo + Val(txt_vta_netas_formales_feb)
   txt_monto_vta_informal_bajo = txt_monto_vta_informal_bajo + Val(txt_vta_netas_informales_feb)

ElseIf Val(txt_vta_total_feb) < Val(txt_promedio_vta_total * 1.1) Then
   txt_tipo_mes_feb = "Medio"
   txt_tipo_mes_r_suma_medio = txt_tipo_mes_r_suma_medio + Val(txt_r_mes_feb)
   txt_tipo_mes_r_monto_medio = txt_tipo_mes_r_monto_medio + Val(txt_vta_total_feb)
   txt_monto_vta_formal_medio = txt_monto_vta_formal_medio + Val(txt_vta_netas_formales_feb)
   txt_monto_vta_informal_medio = txt_monto_vta_informal_medio + Val(txt_vta_netas_informales_feb)
   
Else
    txt_tipo_mes_feb = "Alto"
    txt_tipo_mes_r_suma_alto = txt_tipo_mes_r_suma_alto + Val(txt_r_mes_feb)
    txt_tipo_mes_r_monto_alto = txt_tipo_mes_r_monto_alto + Val(txt_vta_total_feb)
    txt_monto_vta_formal_alto = txt_monto_vta_formal_alto + Val(txt_vta_netas_formales_feb)
    txt_monto_vta_informal_alto = txt_monto_vta_informal_alto + Val(txt_vta_netas_informales_feb)
    
    
  
End If



'CALCULAR CUADROS POSTERIOR AL CALCULO IVA,.... VENTAS PROMEDIOS Y ..
'----------------------------------------------------------------------

If txt_tipo_mes_r_suma_alto <> 0 Then

txt_prom_vta_meses_altos = Int(Val(txt_tipo_mes_r_monto_alto) / Val(txt_tipo_mes_r_suma_alto))
txt_prom_vtas_meses_altos_formal = Int(Val(txt_monto_vta_formal_alto / txt_tipo_mes_r_suma_alto))
txt_prom_vtas_meses_altos_informal = Int(Val(txt_monto_vta_informal_alto / txt_tipo_mes_r_suma_alto))
txt_prom_vtas_meses_altos_informal = Int(Val(txt_monto_vta_informal_alto / txt_tipo_mes_r_suma_alto))

End If

If txt_tipo_mes_r_suma_medio <> 0 Then

txt_prom_vta_meses_medios = Int(Val(txt_tipo_mes_r_monto_medio) / Val(txt_tipo_mes_r_suma_medio))
txt_prom_vtas_meses_medios_formal = Int(Val(txt_monto_vta_formal_medio / txt_tipo_mes_r_suma_medio))
txt_prom_vtas_meses_medios_informal = Int(Val(txt_monto_vta_informal_medio / txt_tipo_mes_r_suma_medio))
txt_prom_vtas_meses_medios_informal = Int(Val(txt_monto_vta_informal_medio / txt_tipo_mes_r_suma_medio))

End If

If txt_tipo_mes_r_suma_bajo <> 0 Then

txt_prom_vta_meses_bajos = Int(Val(txt_tipo_mes_r_monto_bajo) / Val(txt_tipo_mes_r_suma_bajo))
txt_prom_vtas_meses_bajos_formal = Int(Val(txt_monto_vta_formal_bajo / txt_tipo_mes_r_suma_bajo))
txt_prom_vtas_meses_bajos_informal = Int(Val(txt_monto_vta_informal_bajo / txt_tipo_mes_r_suma_bajo))
txt_prom_vtas_meses_bajos_informal = Int(Val(txt_monto_vta_informal_bajo / txt_tipo_mes_r_suma_bajo))

End If





Else

MsgBox "Debe Ingresar Ivas / Compra Promedio Mensual y Numero de Veces"

End If

If Val(txt_procentaje_compra_formal) > 1 Then
MsgBox "El Porcentaje De Compra Formal No Puede Ser Mayor a 1 ...Recalcular"

End If
 
 'prende boton siguiente
 cmd_calcula_costos_fijos.Enabled = True

End Sub




Private Sub cmd_cerrar_caso_volver_menu_Click()
Unload Ficha_Cliente_Micro
Unload Evaluacion_Perfil
Unload Metodologia_Activo_Circulante
Unload Metodologia_IVA
Unload Metodologia_Maxima_Prod

Menu_Principal_Micro.Show
End Sub

Private Sub cmd_costo_promedio_ponderado_Click()

'Pasa Varibles Desde Ficha
txt_rut_cliente = rut_cliente_ficha
txt_dv = dv_cliente_ficha

txt_rut_cliente_pag2 = rut_cliente_ficha
txt_dv_pag2 = dv_cliente_ficha

txt_rut_cliente_pag3 = rut_cliente_ficha
txt_dv_pag3 = dv_cliente_ficha

txt_rut_cliente_pag4 = rut_cliente_ficha
txt_dv_pag4 = dv_cliente_ficha

txt_rut_cliente_pag5 = rut_cliente_ficha
txt_dv_pag5 = dv_cliente_ficha

txt_rut_cliente_pag6 = rut_cliente_ficha
txt_dv_pag6 = dv_cliente_ficha


'Pasa Varibles Desde evaluacion
txt_tipo_cliente_form_evaluacion = txt_tipo_cliente_evaluacion
txt_tipo_riesgo_form_evaluacion = R_Final_Perfil_evaluacion



If txt_cantidad_producto = 1 And txt_precio_venta1 >= 1 And txt_precio_venta1 <> "" _
And Val(txt_precio_venta1) > Val(txt_materia_prima1) And txt_incidencia_ventas1 <> "" Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3)
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3)
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3)


ElseIf txt_cantidad_producto = 2 And txt_precio_venta2 >= 1 And txt_precio_venta2 <> "" And txt_precio_venta1 >= 1 And _
txt_precio_venta1 <> "" And Val(txt_precio_venta1) > Val(txt_materia_prima1) And Val(txt_precio_venta2) > Val(txt_materia_prima2) _
And (txt_incidencia_ventas1 <> "" Or txt_incidencia_ventas2 <> "") Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3)
txt_r_cvcmo2 = Round(((Val(txt_materia_prima2) + Val(txt_mano_obra2)) / Val(txt_precio_venta2)), 3)
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3)
txt_r_cvsmo2 = Round((Val(txt_materia_prima2) / Val(txt_precio_venta2)), 3)
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppcmo2 = Round(txt_r_cvcmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppsmo2 = Round(txt_r_cvsmo2 * txt_incidencia_ventas2 * 0.01, 3)


ElseIf txt_cantidad_producto = 3 And txt_precio_venta2 >= 1 And txt_precio_venta2 <> "" And txt_precio_venta1 >= 1 And _
txt_precio_venta1 <> "" And txt_precio_venta3 >= 1 And txt_precio_venta3 <> "" And Val(txt_precio_venta1) > Val(txt_materia_prima1) _
And Val(txt_precio_venta2) > Val(txt_materia_prima2) And Val(txt_precio_venta3) > Val(txt_materia_prima3) _
And (txt_incidencia_ventas1 <> "" Or txt_incidencia_ventas2 <> "" Or txt_incidencia_ventas3 <> "") Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3)
txt_r_cvcmo2 = Round(((Val(txt_materia_prima2) + Val(txt_mano_obra2)) / Val(txt_precio_venta2)), 3)
txt_r_cvcmo3 = Round(((Val(txt_materia_prima3) + Val(txt_mano_obra3)) / Val(txt_precio_venta3)), 3)
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3)
txt_r_cvsmo2 = Round((Val(txt_materia_prima2) / Val(txt_precio_venta2)), 3)
txt_r_cvsmo3 = Round((Val(txt_materia_prima3) / Val(txt_precio_venta3)), 3)
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppcmo2 = Round(txt_r_cvcmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppcmo3 = Round(txt_r_cvcmo3 * txt_incidencia_ventas3 * 0.01, 3)
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppsmo2 = Round(txt_r_cvsmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppsmo3 = Round(txt_r_cvsmo3 * txt_incidencia_ventas3 * 0.01, 3)


ElseIf txt_cantidad_producto = 4 And txt_precio_venta2 >= 1 And txt_precio_venta2 <> "" And txt_precio_venta1 >= 1 _
And txt_precio_venta1 <> "" And txt_precio_venta3 >= 1 And txt_precio_venta3 <> "" And txt_precio_venta4 >= 1 _
And txt_precio_venta4 <> "" And Val(txt_precio_venta1) > Val(txt_materia_prima1) And Val(txt_precio_venta2) > Val(txt_materia_prima2) _
And Val(txt_precio_venta3) > Val(txt_materia_prima3) And Val(txt_precio_venta4) > Val(txt_materia_prima4) _
And (txt_incidencia_ventas1 <> "" Or txt_incidencia_ventas2 <> "" Or txt_incidencia_ventas3 <> "" _
Or txt_incidencia_ventas4 <> "") Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3)
txt_r_cvcmo2 = Round(((Val(txt_materia_prima2) + Val(txt_mano_obra2)) / Val(txt_precio_venta2)), 3)
txt_r_cvcmo3 = Round(((Val(txt_materia_prima3) + Val(txt_mano_obra3)) / Val(txt_precio_venta3)), 3)
txt_r_cvcmo4 = Round(((Val(txt_materia_prima4) + Val(txt_mano_obra4)) / Val(txt_precio_venta4)), 3)
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3)
txt_r_cvsmo2 = Round((Val(txt_materia_prima2) / Val(txt_precio_venta2)), 3)
txt_r_cvsmo3 = Round((Val(txt_materia_prima3) / Val(txt_precio_venta3)), 3)
txt_r_cvsmo4 = Round((Val(txt_materia_prima4) / Val(txt_precio_venta4)), 3)
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppcmo2 = Round(txt_r_cvcmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppcmo3 = Round(txt_r_cvcmo3 * txt_incidencia_ventas3 * 0.01, 3)
txt_r_cvppcmo4 = Round(txt_r_cvcmo4 * txt_incidencia_ventas4 * 0.01, 3)
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppsmo2 = Round(txt_r_cvsmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppsmo3 = Round(txt_r_cvsmo3 * txt_incidencia_ventas3 * 0.01, 3)
txt_r_cvppsmo4 = Round(txt_r_cvsmo4 * txt_incidencia_ventas4 * 0.01, 3)



ElseIf txt_cantidad_producto = 5 And txt_precio_venta2 >= 1 And txt_precio_venta2 <> "" And txt_precio_venta1 >= 1 _
And txt_precio_venta1 <> "" And txt_precio_venta3 >= 1 And txt_precio_venta3 <> "" And txt_precio_venta4 >= 1 _
And txt_precio_venta4 <> "" And txt_precio_venta5 >= 1 And txt_precio_venta5 <> "" And Val(txt_precio_venta1) > Val(txt_materia_prima1) _
And Val(txt_precio_venta2) > Val(txt_materia_prima2) And Val(txt_precio_venta3) > Val(txt_materia_prima3) And _
Val(txt_precio_venta4) > Val(txt_materia_prima4) And Val(txt_precio_venta5) > Val(txt_materia_prima5) _
And (txt_incidencia_ventas1 <> "" Or txt_incidencia_ventas2 <> "" Or txt_incidencia_ventas3 <> "" _
Or txt_incidencia_ventas4 <> "" Or txt_incidencia_ventas5 <> "") Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3)
txt_r_cvcmo2 = Round(((Val(txt_materia_prima2) + Val(txt_mano_obra2)) / Val(txt_precio_venta2)), 3)
txt_r_cvcmo3 = Round(((Val(txt_materia_prima3) + Val(txt_mano_obra3)) / Val(txt_precio_venta3)), 3)
txt_r_cvcmo4 = Round(((Val(txt_materia_prima4) + Val(txt_mano_obra4)) / Val(txt_precio_venta4)), 3)
txt_r_cvcmo5 = Round(((Val(txt_materia_prima5) + Val(txt_mano_obra5)) / Val(txt_precio_venta5)), 3)
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppcmo2 = Round(txt_r_cvcmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppcmo3 = Round(txt_r_cvcmo3 * txt_incidencia_ventas3 * 0.01, 3)
txt_r_cvppcmo4 = Round(txt_r_cvcmo4 * txt_incidencia_ventas4 * 0.01, 3)
txt_r_cvppcmo5 = Round(txt_r_cvcmo5 * txt_incidencia_ventas5 * 0.01, 3)
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3)
txt_r_cvsmo2 = Round((Val(txt_materia_prima2) / Val(txt_precio_venta2)), 3)
txt_r_cvsmo3 = Round((Val(txt_materia_prima3) / Val(txt_precio_venta3)), 3)
txt_r_cvsmo4 = Round((Val(txt_materia_prima4) / Val(txt_precio_venta4)), 3)
txt_r_cvsmo5 = Round((Val(txt_materia_prima5) / Val(txt_precio_venta5)), 3)
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3)
txt_r_cvppsmo2 = Round(txt_r_cvsmo2 * txt_incidencia_ventas2 * 0.01, 3)
txt_r_cvppsmo3 = Round(txt_r_cvsmo3 * txt_incidencia_ventas3 * 0.01, 3)
txt_r_cvppsmo4 = Round(txt_r_cvsmo4 * txt_incidencia_ventas4 * 0.01, 3)
txt_r_cvppsmo5 = Round(txt_r_cvsmo5 * txt_incidencia_ventas5 * 0.01, 3)


Else

MsgBox "El Precio Venta Debe Ser Mayor a cero y Precio Venta Mayor a Materia Prima"
End If
'------

TextBox_sin = Round((txt_r_cvppsmo1) * 1 + (txt_r_cvppsmo2) * 1 + (txt_r_cvppsmo3) * 1 + (txt_r_cvppsmo4) * 1 + (txt_r_cvppsmo5) * 1, 3)
TextBox_con = Round((txt_r_cvppcmo1) * 1 + (txt_r_cvppcmo2) * 1 + (txt_r_cvppcmo3) * 1 + (txt_r_cvppcmo4) * 1 + (txt_r_cvppcmo5) * 1, 3)

If (TextBox_sin * 1) > (TextBox_con * 1) Then
txt_Sub_Total = TextBox_sin
Else
txt_Sub_Total = TextBox_con
End If
txt_Sub_Total_x1 = txt_Sub_Total * 1.1


If Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5) = 100 Then
    'Prender siguiente Boton Calculo
    cmd_calcular_vta_total_mes_al_me_ba.Enabled = True
    
'''suma incidencias al estar correctas
txt_total_incidencias = Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5)

Else
 MsgBox "Las Incidencias Deben sumar 100%"
  
  'Prender siguiente Boton Calculo
cmd_calcular_vta_total_mes_al_me_ba.Enabled = False
  
End If


''''''' CONDICIONES PARA RESULTADO DE ESTADO COSTO VARIABLE PONDERADO
   
   If txt_Sub_Total_x1 >= 10 And txt_Sub_Total_x1 <= 20 Then
        
        txt_r_promedio_ponderado = "ZG"
        ElseIf txt_Sub_Total_x1 > 20 Then
        txt_r_promedio_ponderado = "A"
        
        ElseIf txt_Sub_Total_x1 < 10 Then
        txt_r_promedio_ponderado = "ZG"
    
  End If


End Sub



Private Sub cmd_guardar_evaluacion_Click()

If txt_mto_bruto_sol_cliente > 0 And txt_mto_bruto_sol_cliente <> "" And txt_cuota_credito > 0 And txt_cuota_credito <> "" Then

Dim fec1
Dim hora1

fec1 = Format(Date, "yyyy/mm/dd")
txt_fecha_actual = fec1

hora1 = hora
txt_hora_actual = Time


' La conexión a la base de datos
    Call conectarBD
    'Dim cn As ADODB.Connection
    'Set cn = New ADODB.Connection
    
    'Dim cnn As ADODB.Connection
    'Dim rst As ADODB.Recordset
    'Dim ssql As String
    ' Creamos un nuevo objeto Connection
    'Set cnn = New ADODB.Connection
    
    
    'cn.Open "Provider=SQLNCLI; " & _
    "Initial Catalog=BD_GES_FAM; " & _
    "Data Source=CLMOBCAMASD01P; " & _
    "integrated security=SSPI; persist security info=True;"
      
   'INSERTA DATOS A TABLA SQL
   
      irespuesta = MsgBox("¿Esta Seguro Que Desea Guardar La Evaluacion Final?", vbYesNo)
        
        If irespuesta = vbYes Then

    'Dim STSQL2 As String
    'Dim STSQL3 As String
    'Dim STSQL4 As String
        
    
    '----------------------------------------------------------------aqui en adelante

' SELECCIONA NUMERO DE SOLICITUD ASIGNADA A CLIENTE
    ' Abrimos la conexión
    'cnn.Open "Provider=SQLNCLI; " & _
    "Initial Catalog=BD_GES_FAM; " & _
    "Data Source=CLMOBCAMASD01P; " & _
    "integrated security=SSPI; persist security info=True;"
    
    'Set rst = New ADODB.Recordset
    
  
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
    
    
    
    'INSERTAR DATOS EN TABLA DE EVALUACION DEL RIESGO DEL CLIENTE
    ssql = "INSERT INTO TBL_MICRO_PERFIL_RIESGO_CLIENTE " _
    & "([Rut_Cliente], [N_Solicitud],[Dv],[Cliente_Nuevo], [Bancarizado],[Antiguedad_Banco],[mora_promedio_dias_BD],[mora_maxima_dias_BD]," _
    & " [R_Tipo_Cliente],[Registro_Ventas],[R_FINAL_PERFIL],[fecha_ingreso],[hora_ingreso],[metodologia_asignada])" _
    & " VALUES (('" & rut_cliente_ficha & "'),('" & txt_n_solicitud & "') " _
    & ",('" & dv_cliente_ficha & "') , ('" & Cliente_Nuevo_evaluacion & "'), ('" & Bancarizado_evaluacion & "'),('" & Antiguedad_banco_evaluacion & "')" _
    & ",('" & mora_promedio_dias_BD_evaluacion & "'), ('" & mora_maxima_dias_BD_evaluacion & "'),('" & txt_tipo_cliente_evaluacion & "')" _
    & ",('" & registros_ventas_evaluacion & "'), ('" & R_Final_Perfil_evaluacion & "'),('" & txt_fecha_actual & "'),('" & txt_hora_actual & "'), ('" & metodologia_asignada & "'))"
    
    'MsgBox STSQL2
    cnn.Execute ssql
    
    
    ssql = "INSERT INTO TBL_MICRO_METODOLOGIA_IVA " _
& " ([RUT_CLIENTE], [n_solicitud],[DV]," _
& " [producto1],[producto2],[producto3],[producto4],[producto5],[precio_venta1],[precio_venta2],[precio_venta3],[precio_venta4],[precio_venta5],[materia_prima1],[materia_prima2],[materia_prima3],[materia_prima4],[materia_prima5],[mano_obra1],[mano_obra2],[mano_obra3],[mano_obra4],[mano_obra5],[incidencia_ventas1],[incidencia_ventas2],[incidencia_ventas3],[incidencia_ventas4],[incidencia_ventas5],[r_cvcmo1],[r_cvcmo2],[r_cvcmo3],[r_cvcmo4],[r_cvcmo5],[r_cvsmo1],[r_cvsmo2],[r_cvsmo3],[r_cvsmo4],[r_cvsmo5],[r_cvppcmo1],[r_cvppcmo2],[r_cvppcmo3],[r_cvppcmo4],[r_cvppcmo5],[r_cvppsmo1],[r_cvppsmo2],[r_cvppsmo3],[r_cvppsmo4],[r_cvppsmo5],[r_Subtotal_costo_variable],[r_Subtotal_x1_costo_variable],[r_total_iva_credito],[r_total_iva_debito],[r_total_compra_neta],[r_total_vta_netas_formales],[r_total_vta_netas_informales],[r_total_compra_total],[r_total_vta_total],[r_total_margen_total],[r_promedio_iva_credito],[r_promedio_iva_debito],[r_promedio_compra_neta],[r_promedio_vta_netas_formales]," _
& " [r_promedio_vta_netas_informales],[r_promedio_compra_total],[r_promedio_vta_total],[compra_promedio_mensual],[veces_compra_mes],[r_porcentaje_compra_formal]," _
& " [r_tot_promedio_ventas_mes_alto],[r_tot_promedio_ventas_mes_medio],[r_tot_promedio_ventas_mes_bajo],[r_tot_promedio_ventas_formal_mes_alto],[r_tot_promedio_ventas_formal_mes_medio],[r_tot_promedio_ventas_formal_mes_bajo],[r_tot_promedio_ventas_informal_mes_alto],[r_tot_promedio_ventas_informal_mes_medio],[r_tot_promedio_ventas_informal_mes_bajo]," _
& " [arriendo_micro],[sueldos],[movilizacion],[servicios_basicos],[contador],[lubricantes],[neumaticos],[afinamientos],[patentes_seguros],[otros_costos_fijos],[total_costos_fijos],[valor_uf],[n_grupo_familiar],[arriendo_vivienda_Gastos_Fam],[gastos_indicado_cliente],[total_gasto_familiar],[liquidacion_sueldo],[jubilacion],[montepio],[arriendo_vivienda_Otro_Ing],[ingreso_segunda_microempresa],[boleta_honorario],[total_otros_ingresos],[acreedor1_deuda],[acreedor2_deuda],[acreedor3_deuda],[acreedor4_deuda],[acreedor5_deuda],[acreedor6_deuda],[tipo_producto1_deuda],[tipo_producto2_deuda],[tipo_producto3_deuda],[tipo_producto4_deuda],[tipo_producto5_deuda],[tipo_producto6_deuda],[saldo_pendiente1_deuda],[saldo_pendiente2_deuda],[saldo_pendiente3_deuda],[saldo_pendiente4_deuda],[saldo_pendiente5_deuda],[saldo_pendiente6_deuda],[monto_cuota1_deuda]," _
& " [monto_cuota2_deuda],[monto_cuota3_deuda],[monto_cuota4_deuda],[monto_cuota5_deuda],[monto_cuota6_deuda],[cuotas_pactadas1_deuda],[cuotas_pactadas2_deuda],[cuotas_pactadas3_deuda],[cuotas_pactadas4_deuda],[cuotas_pactadas5_deuda],[cuotas_pactadas6_deuda],[cuotas_pendientes1_deuda],[cuotas_pendientes2_deuda],[cuotas_pendientes3_deuda],[cuotas_pendientes4_deuda],[cuotas_pendientes5_deuda],[cuotas_pendientes6_deuda],[prepaga_cuota1_deuda],[prepaga_cuota2_deuda],[prepaga_cuota3_deuda],[prepaga_cuota4_deuda],[prepaga_cuota5_deuda],[prepaga_cuota6_deuda],[total_saldo_pendiente_deuda],[total_deudas],[numero_meses_alto_flujo],[numero_meses_medio_flujo],[numero_meses_bajo_flujo],[vta_formal_promedio_mes_alto_flujo],[vta_formal_promedio_mes_medio_flujo],[vta_formal_promedio_mes_bajo_flujo],[vta_informal_promedio_mes_alto_flujo],[vta_informal_promedio_mes_medio_flujo],[vta_informal_promedio_mes_bajo_flujo],[Venta_Total_Promedio_Mes_Alto_flujo]," _
& " [Venta_Total_Promedio_Mes_medio_flujo],[Venta_Total_Promedio_Mes_bajo_flujo],[resultado_operacional_alto_flujo],[resultado_operacional_medio_flujo],[resultado_operacional_bajo_flujo],[capacidad_pago_mes_alto_flujo],[capacidad_pago_mes_medio_flujo],[capacidad_pago_mes_bajo_flujo],[cap_pago_corregida_ajus_mes_alto_flujo],[cap_pago_corregida_ajus_mes_medio_flujo],[cap_pago_corregida_ajus_mes_bajo_flujo],[cap_pago_promedio_corregida_ajustada_flujo],[monto_maximo_credito_flujo],[cuota_credito_flujo],[mto_bruto_solicitado_cliente_flujo],[resolucion_credito_cuota_flujo],[resolucion_credito_monto_flujo],[fecha_ingreso],[hora_ingreso],[impuesto])" _
& " VALUES (('" & txt_rut_cliente & "'),('" & txt_n_solicitud & "'),('" & txt_dv & "'),('" & txt_producto1 & "'),('" & txt_producto2 & "'),('" & txt_producto3 & "'),('" & txt_producto4 & "'),('" & txt_producto5 & "'),('" & txt_precio_venta1 & "'),('" & txt_precio_venta2 & "'),('" & txt_precio_venta3 & "'),('" & txt_precio_venta4 & "'),('" & txt_precio_venta5 & "'),('" & txt_materia_prima1 & "'),('" & txt_materia_prima2 & "'),('" & txt_materia_prima3 & "'),('" & txt_materia_prima4 & "'),('" & txt_materia_prima5 & "'),('" & txt_mano_obra1 & "'),('" & txt_mano_obra2 & "'),('" & txt_mano_obra3 & "')" _
& ",('" & txt_mano_obra4 & "'),('" & txt_mano_obra5 & "'),('" & txt_incidencia_ventas1 & "'),('" & txt_incidencia_ventas2 & "'),('" & txt_incidencia_ventas3 & "'),('" & txt_incidencia_ventas4 & "'),('" & txt_incidencia_ventas5 & "'),('" & txt_r_cvcmo1 & "'),('" & txt_r_cvcmo2 & "'),('" & txt_r_cvcmo3 & "'),('" & txt_r_cvcmo4 & "'),('" & txt_r_cvcmo5 & "'),('" & txt_r_cvsmo1 & "'),('" & txt_r_cvsmo2 & "'),('" & txt_r_cvsmo3 & "'),('" & txt_r_cvsmo4 & "'),('" & txt_r_cvsmo5 & "'),('" & txt_r_cvppcmo1 & "'),('" & txt_r_cvppcmo2 & "'),('" & txt_r_cvppcmo3 & "'),('" & txt_r_cvppcmo4 & "'),('" & txt_r_cvppcmo5 & "'),('" & txt_r_cvppsmo1 & "'),('" & txt_r_cvppsmo2 & "'),('" & txt_r_cvppsmo3 & "'),('" & txt_r_cvppsmo4 & "'),('" & txt_r_cvppsmo5 & "'),('" & txt_Sub_Total & "'),('" & txt_Sub_Total_x1 & "'),('" & txt_total_iva_credito & "'),('" & txt_total_iva_debito & "'),('" & txt_total_compra_neta & "'),('" & txt_total_vta_netas_formales & "'),('" & txt_total_vta_netas_informales & "')" _
& ",('" & txt_total_compra_total & "'),('" & txt_total_vta_total & "'),('" & txt_total_margen_total & "'),('" & txt_promedio_iva_credito & "'),('" & txt_promedio_iva_debito & "'),('" & txt_promedio_compra_neta & "'),('" & txt_promedio_vta_netas_formales & "'),('" & txt_promedio_vta_netas_informales & "'),('" & txt_promedio_compra_total & "'),('" & txt_promedio_vta_total & "'),('" & txt_compra_promedio_mensual & "'),('" & txt_veces_compra_mes & "'),('" & txt_porcentaje_compra_formal & "'),('" & txt_prom_vta_meses_altos & "'),('" & txt_prom_vta_meses_medios & "'),('" & txt_prom_vta_meses_bajos & "'),('" & txt_prom_vtas_meses_altos_formal & "'),('" & txt_prom_vtas_meses_medios_formal & "') " _
& ",('" & txt_prom_vtas_meses_bajos_formal & "'),('" & txt_prom_vtas_meses_altos_informal & "'),('" & txt_prom_vtas_meses_medios_informal & "'),('" & txt_prom_vtas_meses_bajos_informal & "') " _
& ",('" & txt_arriendo_micro & "'),('" & txt_sueldos & "'),('" & txt_movilizacion & "'),('" & txt_servicios_basicos & "'),('" & txt_contador & "'),('" & txt_lubricantes & "'),('" & txt_neumaticos & "'),('" & txt_afinamientos & "'),('" & txt_patentes_seguros & "'),('" & txt_otros_costos_fijos & "'),('" & txt_total_costos_fijos & "'),('" & txt_valor_uf & "'),('" & txt_n_grupo_familiar & "'),('" & txt_arriendo_vivienda & "'),('" & txt_gastos_indicado_cliente & "'),('" & txt_total_gasto_familiar & "'),('" & txt_liquidacion_sueldo & "'),('" & txt_jubilacion & "'),('" & txt_montepio & "'),('" & txt_arriendo_vivienda1 & "'),('" & txt_ingreso_segunda_microempresa & "'),('" & txt_boleta_honorario & "'),('" & txt_total_otros_ingresos & "'),('" & txt_acreedor1 & "'),('" & txt_acreedor2 & "'),('" & txt_acreedor3 & "')" _
& ",('" & txt_acreedor4 & "'),('" & txt_acreedor5 & "'),('" & txt_acreedor6 & "'),('" & txt_tipo_producto1 & "'),('" & txt_tipo_producto2 & "'),('" & txt_tipo_producto3 & "'),('" & txt_tipo_producto4 & "'),('" & txt_tipo_producto5 & "'),('" & txt_tipo_producto6 & "'),('" & txt_saldo_pendiente1 & "'),('" & txt_saldo_pendiente2 & "'),('" & txt_saldo_pendiente3 & "'),('" & txt_saldo_pendiente4 & "'),('" & txt_saldo_pendiente5 & "'),('" & txt_saldo_pendiente6 & "'),('" & txt_monto_cuota1 & "'),('" & txt_monto_cuota2 & "'),('" & txt_monto_cuota3 & "'),('" & txt_monto_cuota4 & "'),('" & txt_monto_cuota5 & "'),('" & txt_monto_cuota6 & "'),('" & txt_cuotas_pactadas1 & "'),('" & txt_cuotas_pactadas2 & "'),('" & txt_cuotas_pactadas3 & "'),('" & txt_cuotas_pactadas4 & "'),('" & txt_cuotas_pactadas5 & "'),('" & txt_cuotas_pactadas6 & "'),('" & txt_cuotas_pendientes1 & "'),('" & txt_cuotas_pendientes2 & "'),('" & txt_cuotas_pendientes3 & "')" _
& ",('" & txt_cuotas_pendientes4 & "'),('" & txt_cuotas_pendientes5 & "'),('" & txt_cuotas_pendientes6 & "'),('" & cbx_prepaga_deuda1 & "'),('" & cbx_prepaga_deuda2 & "'),('" & cbx_prepaga_deuda3 & "'),('" & cbx_prepaga_deuda4 & "'),('" & cbx_prepaga_deuda5 & "'),('" & cbx_prepaga_deuda6 & "'),('" & txt_total_saldo_pendiente & "'),('" & txt_total_deudas & "'),('" & numero_meses_tipo_mes_alto & "'),('" & numero_meses_tipo_mes_medio & "'),('" & numero_meses_tipo_mes_bajo & "'),('" & txt_vta_formal_promedio_mes_alto & "'),('" & txt_vta_formal_promedio_mes_medio & "'),('" & txt_vta_formal_promedio_mes_bajo & "'),('" & txt_vta_informal_promedio_mes_alto & "'),('" & txt_vta_informal_promedio_mes_medio & "'),('" & txt_vta_informal_promedio_mes_bajo & "'),('" & txt_Venta_Total_Promedio_Mes_Alto & "'),('" & txt_Venta_Total_Promedio_Mes_Medio & "'),('" & txt_Venta_Total_Promedio_Mes_Bajo & "'),('" & txt_resultado_operacional_mes_alto & "')" _
& ",('" & txt_resultado_operacional_mes_medio & "'),('" & txt_resultado_operacional_mes_bajo & "'),('" & txt_capacidad_pago_mes_alto & "'),('" & txt_capacidad_pago_mes_medio & "'),('" & txt_capacidad_pago_mes_bajo & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_alto & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_medio & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_bajo & "'),('" & txt_capacidad_pago_promedio_corregida_ajustada & "'),('" & txt_monto_maximo_credito & "'),('" & txt_cuota_credito & "'),('" & txt_mto_bruto_sol_cliente & "'),('" & txt_resolucion_credito_por_cuota & "'),('" & txt_aprobacion & "'),('" & txt_fecha_actual & "'),('" & txt_hora_actual & "'),('" & txt_impuesto & "'))"

    'MsgBox STSQL3
    cnn.Execute ssql
    

 ssql = "INSERT INTO TBL_MICRO_IVA_MES " _
& "([Rut_Cliente],[n_solicitud],[Dv],[Ano_Declaracion_Iva_Ene],[Iva_Credito_Ene],[Iva_Debito_Ene],[Compras_Netas_Ene],[Ventas_Netas_Formales_Ene],[Ventas_Netas_Informales_Ene],[Compra_Total_Ene],[Venta_Total_Ene],[Tipo_Mes_Ene],[Margen_Total_Ene],[Ano_Declaracion_Iva_feb],[Iva_Credito_feb],[Iva_Debito_feb],[Compras_Netas_feb],[Ventas_Netas_Formales_feb],[Ventas_Netas_Informales_feb],[Compra_Total_feb],[Venta_Total_feb],[Tipo_Mes_feb],[Margen_Total_feb],[Ano_Declaracion_Iva_mar],[Iva_Credito_mar],[Iva_Debito_mar],[Compras_Netas_mar],[Ventas_Netas_Formales_mar],[Ventas_Netas_Informales_mar],[Compra_Total_mar],[Venta_Total_mar],[Tipo_Mes_mar],[Margen_Total_mar],[Ano_Declaracion_Iva_abr],[Iva_Credito_abr],[Iva_Debito_abr],[Compras_Netas_abr],[Ventas_Netas_Formales_abr],[Ventas_Netas_Informales_abr],[Compra_Total_abr],[Venta_Total_abr],[Tipo_Mes_abr],[margen_Total_abr],[Ano_Declaracion_Iva_may],[Iva_Credito_may],[Iva_Debito_may],[Compras_Netas_may],[Ventas_Netas_Formales_may],[Ventas_Netas_Informales_may]," _
& " [Compra_Total_may],[Venta_Total_may],[Tipo_Mes_may],[margen_Total_may],[Ano_Declaracion_Iva_jun],[Iva_Credito_jun],[Iva_Debito_jun],[Compras_Netas_jun],[Ventas_Netas_Formales_jun],[Ventas_Netas_Informales_jun],[Compra_Total_jun],[Venta_Total_jun],[Tipo_Mes_jun],[margen_Total_jun],[Ano_Declaracion_Iva_jul],[Iva_Credito_jul],[Iva_Debito_jul],[Compras_Netas_jul],[Ventas_Netas_Formales_jul],[Ventas_Netas_Informales_jul],[Compra_Total_jul],[Venta_Total_jul],[Tipo_Mes_jul],[margen_Total_jul],[Ano_Declaracion_Iva_ago],[Iva_Credito_ago],[Iva_Debito_ago],[Compras_Netas_ago],[Ventas_Netas_Formales_ago],[Ventas_Netas_Informales_ago],[Compra_Total_ago],[Venta_Total_ago],[Tipo_Mes_ago],[margen_Total_ago],[Ano_Declaracion_Iva_sep],[Iva_Credito_sep],[Iva_Debito_sep],[Compras_Netas_sep],[Ventas_Netas_Formales_sep],[Ventas_Netas_Informales_sep],[Compra_Total_sep],[Venta_Total_sep],[Tipo_Mes_sep],[margen_Total_sep],[Ano_Declaracion_Iva_oct],[Iva_Credito_oct],[Iva_Debito_oct],[Compras_Netas_oct], " _
& " [Ventas_Netas_Formales_oct],[Ventas_Netas_Informales_oct],[Compra_Total_oct],[Venta_Total_oct],[Tipo_Mes_oct],[margen_Total_oct],[Ano_Declaracion_Iva_nov],[Iva_Credito_nov],[Iva_Debito_nov],[Compras_Netas_nov],[Ventas_Netas_Formales_nov],[Ventas_Netas_Informales_nov],[Compra_Total_nov],[Venta_Total_nov],[Tipo_Mes_nov],[margen_Total_nov],[Ano_Declaracion_Iva_dic],[Iva_Credito_dic],[Iva_Debito_dic],[Compras_Netas_dic],[Ventas_Netas_Formales_dic],[Ventas_Netas_Informales_dic],[Compra_Total_dic],[Venta_Total_dic],[Tipo_Mes_dic],[margen_Total_dic],[Fecha_Ingreso],[Hora_Ingreso],[mes_inicio_iva])" _
& " VALUES(('" & txt_rut_cliente & "'),('" & txt_n_solicitud & "'),('" & txt_dv & "')" _
& ",('" & txt_ano_iva_ene & "'),('" & txt_iva_credito_ene & "'),('" & txt_iva_debito_ene & "'),('" & txt_compra_neta_ene & "'),('" & txt_vta_netas_formales_ene & "'),('" & txt_vta_netas_informales_ene & "'),('" & txt_compra_total_ene & "'),('" & txt_vta_total_ene & "'),('" & txt_tipo_mes_ene & "'),('" & txt_margen_total_ene & "'),('" & txt_ano_iva_feb & "'),('" & txt_iva_credito_feb & "'),('" & txt_iva_debito_feb & "'),('" & txt_compra_neta_feb & "'),('" & txt_vta_netas_formales_feb & "'),('" & txt_vta_netas_informales_feb & "'),('" & txt_compra_total_feb & "'),('" & txt_vta_total_feb & "'),('" & txt_tipo_mes_feb & "'),('" & txt_margen_total_feb & "'),('" & txt_ano_iva_mar & "'),('" & txt_iva_credito_mar & "'),('" & txt_iva_debito_mar & "'),('" & txt_compra_neta_mar & "'),('" & txt_vta_netas_formales_mar & "'),('" & txt_vta_netas_informales_mar & "'),('" & txt_compra_total_mar & "')" _
& ",('" & txt_vta_total_mar & "'), ('" & txt_tipo_mes_mar & "'),('" & txt_margen_total_mar & "'),('" & txt_ano_iva_abr & "'),('" & txt_iva_credito_abr & "'),('" & txt_iva_debito_abr & "'),('" & txt_compra_neta_abr & "'),('" & txt_vta_netas_formales_abr & "'),('" & txt_vta_netas_informales_abr & "'),('" & txt_compra_total_abr & "'),('" & txt_vta_total_abr & "'),('" & txt_tipo_mes_abr & "'),('" & txt_margen_total_abr & "'),('" & txt_ano_iva_may & "'),('" & txt_iva_credito_may & "'),('" & txt_iva_debito_may & "'),('" & txt_compra_neta_may & "'),('" & txt_vta_netas_formales_may & "'),('" & txt_vta_netas_informales_may & "'),('" & txt_compra_total_may & "'),('" & txt_vta_total_may & "'),('" & txt_tipo_mes_may & "'),('" & txt_margen_total_may & "'),('" & txt_ano_iva_jun & "'),('" & txt_iva_credito_jun & "'),('" & txt_iva_debito_jun & "'),('" & txt_compra_neta_jun & "'),('" & txt_vta_netas_formales_jun & "'),('" & txt_vta_netas_informales_jun & "')" _
& ",('" & txt_compra_total_jun & "'),('" & txt_vta_total_jun & "'),('" & txt_tipo_mes_jun & "'),('" & txt_margen_total_jun & "'),('" & txt_ano_iva_jul & "'),('" & txt_iva_credito_jul & "'),('" & txt_iva_debito_jul & "'),('" & txt_compra_neta_jul & "'),('" & txt_vta_netas_formales_jul & "'),('" & txt_vta_netas_informales_jul & "'),('" & txt_compra_total_jul & "'),('" & txt_vta_total_jul & "'),('" & txt_tipo_mes_jul & "'),('" & txt_margen_total_jul & "'),('" & txt_ano_iva_ago & "'),('" & txt_iva_credito_ago & "'),('" & txt_iva_debito_ago & "'),('" & txt_compra_neta_ago & "'),('" & txt_vta_netas_formales_ago & "'),('" & txt_vta_netas_informales_ago & "'),('" & txt_compra_total_ago & "'),('" & txt_vta_total_ago & "'),('" & txt_tipo_mes_ago & "'),('" & txt_margen_total_ago & "'),('" & txt_ano_iva_sep & "'),('" & txt_iva_credito_sep & "'),('" & txt_iva_debito_sep & "'),('" & txt_compra_neta_sep & "') ,('" & txt_vta_netas_formales_sep & "')" _
& ",('" & txt_vta_netas_informales_sep & "'),('" & txt_compra_total_sep & "'),('" & txt_vta_total_sep & "'),('" & txt_tipo_mes_sep & "'),('" & txt_margen_total_sep & "'),('" & txt_ano_iva_oct & "'),('" & txt_iva_credito_oct & "'),('" & txt_iva_debito_oct & "'),('" & txt_compra_neta_oct & "'),('" & txt_vta_netas_formales_oct & "'),('" & txt_vta_netas_informales_oct & "'),('" & txt_compra_total_oct & "'),('" & txt_vta_total_oct & "'),('" & txt_tipo_mes_oct & "'),('" & txt_margen_total_oct & "'),('" & txt_ano_iva_nov & "'),('" & txt_iva_credito_nov & "'),('" & txt_iva_debito_nov & "'),('" & txt_compra_neta_nov & "'),('" & txt_vta_netas_formales_nov & "'),('" & txt_vta_netas_informales_nov & "'),('" & txt_compra_total_nov & "'),('" & txt_vta_total_nov & "'),('" & txt_tipo_mes_nov & "'),('" & txt_margen_total_nov & "'),('" & txt_ano_iva_dic & "'),('" & txt_iva_credito_dic & "'),('" & txt_iva_debito_dic & "'),('" & txt_compra_neta_dic & "')" _
& ",('" & txt_vta_netas_formales_dic & "'),('" & txt_vta_netas_informales_dic & "'),('" & txt_compra_total_dic & "'),('" & txt_vta_total_dic & "'),('" & txt_tipo_mes_dic & "'),('" & txt_margen_total_dic & "'),('" & txt_fecha_actual & "'),('" & txt_hora_actual & "'),('" & cbx_mes_inicio_iva & "'))"
 
 'MsgBox STSQL4
 
    cnn.Execute ssql



    Unload Ficha_Cliente_Micro
    Unload Evaluacion_Perfil
    Unload Metodologia_Activo_Circulante
    Unload Metodologia_IVA
    Unload Metodologia_Maxima_Prod

    Menu_Principal_Micro.Show
    
 End If
 Else
    MsgBox "Debe Ingresar Cuota Comercial y Monto Solicitado Por El Cliente"
 End If
    
End Sub

Private Sub cmd_imprimir1_meto_ac_Click()
Metodologia_IVA.PrintForm
End Sub

Private Sub cmd_imprimir2_meto_ac_Click()
Metodologia_IVA.PrintForm
End Sub

Private Sub cmd_imprimir3_meto_ac_Click()

Metodologia_IVA.PrintForm

End Sub

Private Sub cmd_imprimir4_meto_ac_Click()
Metodologia_IVA.PrintForm
End Sub

Private Sub cmd_imprimir5_meto_ac_Click()
Metodologia_IVA.PrintForm
End Sub

Private Sub cmd_imprimir6_meto_ac_Click()
Metodologia_IVA.PrintForm
End Sub

Private Sub cmd_salir_sistema_Click()
    ActiveWorkbook.Save
    Workbooks("Microempresas_1401.xls").Close
    Application.Quit
End Sub

Private Sub cmd_resumen_Estado_Rechazo_Click()

Metodologia_IVA.Hide

Estado_Resolucion_Final.TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA"

Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.Enabled = False
Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.Enabled = False

Estado_Resolucion_Final.txt_r_f_mora_directa = Ficha_Cliente_Micro.txt_r_mora_directa_interna
Estado_Resolucion_Final.txt_r_f_Vencido_directo = Ficha_Cliente_Micro.txt_r_Vencido_directo_interna
Estado_Resolucion_Final.txt_r_f_castigo_directo = Ficha_Cliente_Micro.txt_r_castigo_directo_interna

''''MORAS SBIF
Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF = Ficha_Cliente_Micro.txt_r_mora_sbif
Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF = Ficha_Cliente_Micro.cbx_r_venc_cast_SBIF
Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF = Ficha_Cliente_Micro.cbx_r_Mora_Total_Sbif
Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF = Ficha_Cliente_Micro.cbx_r_venc_cast_SBIF_indirecta
Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF = Ficha_Cliente_Micro.cbx_r_Mora_Total_Sbif_indirecta

Estado_Resolucion_Final.txt_r_f_file_negativo_tit = Ficha_Cliente_Micro.txt_r_file_negativo_tit
Estado_Resolucion_Final.txt_r_f_n_acreedor = Ficha_Cliente_Micro.txt_r_n_acreedores
Estado_Resolucion_Final.txt_r_f_renegociado = Ficha_Cliente_Micro.txt_r_renegociado
Estado_Resolucion_Final.txt_r_f_protesto_interno = Ficha_Cliente_Micro.txt_r_protesto_interno
Estado_Resolucion_Final.txt_r_f_morosidad_sinac = Ficha_Cliente_Micro.txt_r_morosidad
Estado_Resolucion_Final.txt_r_f_protesto_sinac = Ficha_Cliente_Micro.txt_r_protestos
Estado_Resolucion_Final.txt_r_f_boletin_sinac = Ficha_Cliente_Micro.txt_r_boletin_laboral
Estado_Resolucion_Final.txt_r_f_plazo = Ficha_Cliente_Micro.txt_r_plazo_credito
Estado_Resolucion_Final.txt_r_f_destinos = Ficha_Cliente_Micro.txt_r_accion
Estado_Resolucion_Final.txt_r_f_antiguedad_veh = Ficha_Cliente_Micro.txt_r_años_vehiculo
Estado_Resolucion_Final.txt_r_f_edad = Ficha_Cliente_Micro.txt_r_edad
Estado_Resolucion_Final.txt_r_f_antiguedad_giro = Ficha_Cliente_Micro.txt_r_meses_antiguedad
Estado_Resolucion_Final.txt_r_f_ir_sinac = Ficha_Cliente_Micro.txt_r_predictor_score_dicom

Estado_Resolucion_Final.txt_r_f_ir_tipo_cliente = Evaluacion_Perfil.txt_r_dicom_tipo_cliente

Estado_Resolucion_Final.txt_r_f_deuda_sbif_declarada = Metodologia_IVA.txt_r_sbif_declarada
Estado_Resolucion_Final.txt_r_f_nivel_vta_inf_min = Metodologia_IVA.txt_r_venta_total_min
Estado_Resolucion_Final.txt_r_f_nivel_vta_sup_max = Metodologia_IVA.txt_r_venta_total_max
Estado_Resolucion_Final.txt_r_f_capacidad_pago = Metodologia_IVA.txt_r_capacidad_pago
Estado_Resolucion_Final.txt_r_f_costo_fijo_rub_trasp = Metodologia_IVA.txt_valida_costos_fijos
Estado_Resolucion_Final.txt_r_f_costo_variable_ponde = Metodologia_IVA.txt_r_promedio_ponderado
Estado_Resolucion_Final.txt_r_f_leverage = Metodologia_IVA.txt_r_leverage
Estado_Resolucion_Final.txt_r_f_costo_variable_ponde = Metodologia_IVA.txt_r_promedio_ponderado

''''' CALCULANDO RESULTADO FINAL CON VARIABLES RESUMIDAS

If txt_r_f_mora_directa <> "A" And txt_r_f_Vencido_directo <> "A" And txt_r_f_castigo_directo <> "A" And txt_r_f_protesto_interno <> "A" And _
    txt_r_f_renegociado <> "A" And txt_r_f_n_acreedor <> "A" And txt_r_f_file_negativo <> "A" And txt_r_f_base_fraude <> "A" And txt_r_f_mora_directa_SBIF <> "A" And _
    txt_r_f_vdo_directo_SBIF <> "A" And txt_r_f_cast_directo_SBIF <> "A" And txt_r_f_vdo_indirecto_SBIF <> "A" And txt_r_f_cast_indirecto_SBIF <> "A" And _
    txt_r_f_morosidad_sinac <> "A" And txt_r_f_protesto_sinac <> "A" And txt_r_f_boletin_sinac <> "A" And txt_r_f_ir_sinac <> "A" And txt_r_f_ir_tipo_cliente <> "A" And _
    txt_r_f_plazo <> "A" And txt_r_f_destinos <> "A" And txt_r_f_antiguedad_veh <> "A" And txt_r_f_leverage <> "A" And txt_r_f_capacidad_pago <> "A" And txt_r_f_edad <> "A" And _
    txt_r_f_antiguedad_giro <> "A" And txt_r_f_nivel_vta_inf_min <> "A" And txt_r_f_nivel_vta_sup_max <> "A" And txt_r_f_costo_fijo_rub_trasp <> "A" And txt_r_f_deuda_sbif_declarada <> "A" And _
    txt_r_f_costo_variable_ponde <> "A" And txt_r_f_compra_tot_mensual <> "A" And txt_r_f_factor_ajuste_compra_tot_iva <> "A" Then

    Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.Enabled = False
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.Enabled = True
    
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.BackColor = &HFF&
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.ForeColor = &H8000000E
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred = "R"   ''' ROJO
    
    Else
    
    Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.Enabled = True
    Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.Enabled = False
    
    txt_resultado_APROBADO_final_cred.BackColor = &HC000&
    txt_resultado_APROBADO_final_cred.ForeColor = &H8000000E ' VERDE
    Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred = "A"


End If

Estado_Resolucion_Final.Show

End Sub

Private Sub cmd_volver_evaluacion_Click()
MsgBox "Recuerda Que Al Volver y Cambiar Datos Debes Recalcular Los Campos"

'Evaluacion_Perfil.cbx_Cliente_Nuevo = Empty
'Evaluacion_Perfil.cbx_bancarizado = Empty
'Evaluacion_Perfil.txt_Antiguedad = Empty
'Evaluacion_Perfil.cbx_historia_pago = Empty
'Evaluacion_Perfil.cbx_mora_maxima = Empty
'Evaluacion_Perfil.txt_registro_ventas = Empty
'Evaluacion_Perfil.cbx_actividad_economica_formal = Empty
'Evaluacion_Perfil.cbx_actividad_economica_semiformal = Empty
'Evaluacion_Perfil.cbx_actividad_economica_informal_oficio = Empty
'Evaluacion_Perfil.cbx_actividad_economica_formal_servicio = Empty
'Evaluacion_Perfil.txt_tipo_cliente = Empty
'Evaluacion_Perfil.R_Final_Perfil = Empty

Metodologia_IVA.Hide
Evaluacion_Perfil.Show
End Sub

Private Sub cmd_volver_ficha_Click()
MsgBox "Recuerda Que Al Volver y Cambiar Datos Debes Recalcular Los Campos"
Ficha_Cliente_Micro.cmd_Menu_Evaluacion.Enabled = False
Metodologia_IVA.Hide
Ficha_Cliente_Micro.Show
End Sub

Private Sub Label251_Click()

End Sub

Private Sub Label38_Click()

End Sub

Private Sub txt_acreedor1_Change()
txt_acreedor1 = UCase(txt_acreedor1)
I = Len(txt_acreedor1)
txt_acreedor1.SelStart = I
End Sub

Private Sub txt_acreedor2_Change()
txt_acreedor2 = UCase(txt_acreedor2)
I = Len(txt_acreedor2)
txt_acreedor2.SelStart = I
End Sub

Private Sub txt_acreedor3_Change()
txt_acreedor3 = UCase(txt_acreedor3)
I = Len(txt_acreedor3)
txt_acreedor3.SelStart = I
End Sub

Private Sub txt_acreedor4_Change()
txt_acreedor4 = UCase(txt_acreedor4)
I = Len(txt_acreedor4)
txt_acreedor4.SelStart = I
End Sub

Private Sub txt_acreedor5_Change()
txt_acreedor5 = UCase(txt_acreedor5)
I = Len(txt_acreedor5)
txt_acreedor5.SelStart = I
End Sub

Private Sub txt_acreedor6_Change()
txt_acreedor6 = UCase(txt_acreedor6)
I = Len(txt_acreedor6)
txt_acreedor6.SelStart = I
End Sub

Private Sub txt_cantidad_producto_Change()
If Not IsNumeric(txt_cantidad_producto) Or txt_cantidad_producto = 0 Or txt_cantidad_producto > 5 Then
  MsgBox "El Número de Producto esta entre 1 y 5 y Debe Ser Numerico... Reingrese"
 
Else

txt_r_cvcmo1.Visible = True
txt_r_cvsmo1.Visible = True
txt_r_cvppcmo1.Visible = True
txt_r_cvppsmo1.Visible = True
txt_producto1.Visible = True
txt_precio_venta1.Visible = True
txt_materia_prima1.Visible = True
txt_mano_obra1.Visible = True
txt_incidencia_ventas1.Visible = True

txt_producto2.Visible = False
txt_precio_venta2.Visible = False
txt_materia_prima2.Visible = False
txt_mano_obra2.Visible = False
txt_incidencia_ventas2.Visible = False

txt_producto3.Visible = False
txt_precio_venta3.Visible = False
txt_materia_prima3.Visible = False
txt_mano_obra3.Visible = False
txt_incidencia_ventas3.Visible = False

txt_producto4.Visible = False
txt_precio_venta4.Visible = False
txt_materia_prima4.Visible = False
txt_mano_obra4.Visible = False
txt_incidencia_ventas4.Visible = False

txt_producto5.Visible = False
txt_precio_venta5.Visible = False
txt_materia_prima5.Visible = False
txt_mano_obra5.Visible = False
txt_incidencia_ventas5.Visible = False

txt_r_cvcmo2.Visible = False
txt_r_cvsmo2.Visible = False
txt_r_cvppcmo2.Visible = False
txt_r_cvppsmo2.Visible = False

txt_r_cvcmo3.Visible = False
txt_r_cvsmo3.Visible = False
txt_r_cvppcmo3.Visible = False
txt_r_cvppsmo3.Visible = False

txt_r_cvcmo4.Visible = False
txt_r_cvsmo4.Visible = False
txt_r_cvppcmo4.Visible = False
txt_r_cvppsmo4.Visible = False

txt_r_cvcmo5.Visible = False
txt_r_cvsmo5.Visible = False
txt_r_cvppcmo5.Visible = False
txt_r_cvppsmo5.Visible = False

If txt_cantidad_producto = 2 Then

txt_producto2.Visible = True
txt_precio_venta2.Visible = True
txt_materia_prima2.Visible = True
txt_mano_obra2.Visible = True
txt_incidencia_ventas2.Visible = True
txt_r_cvcmo2.Visible = True
txt_r_cvsmo2.Visible = True
txt_r_cvppcmo2.Visible = True
txt_r_cvppsmo2.Visible = True


ElseIf txt_cantidad_producto = 3 Then

txt_producto2.Visible = True
txt_precio_venta2.Visible = True
txt_materia_prima2.Visible = True
txt_mano_obra2.Visible = True
txt_incidencia_ventas2.Visible = True
txt_producto3.Visible = True
txt_precio_venta3.Visible = True
txt_materia_prima3.Visible = True
txt_mano_obra3.Visible = True
txt_incidencia_ventas3.Visible = True
txt_r_cvcmo2.Visible = True
txt_r_cvsmo2.Visible = True
txt_r_cvppcmo2.Visible = True
txt_r_cvppsmo2.Visible = True
txt_r_cvcmo3.Visible = True
txt_r_cvsmo3.Visible = True
txt_r_cvppcmo3.Visible = True
txt_r_cvppsmo3.Visible = True


ElseIf txt_cantidad_producto = 4 Then

txt_producto2.Visible = True
txt_precio_venta2.Visible = True
txt_materia_prima2.Visible = True
txt_mano_obra2.Visible = True
txt_incidencia_ventas2.Visible = True
txt_producto3.Visible = True
txt_precio_venta3.Visible = True
txt_materia_prima3.Visible = True
txt_mano_obra3.Visible = True
txt_incidencia_ventas3.Visible = True
txt_producto4.Visible = True
txt_precio_venta4.Visible = True
txt_materia_prima4.Visible = True
txt_mano_obra4.Visible = True
txt_incidencia_ventas4.Visible = True
txt_r_cvcmo2.Visible = True
txt_r_cvsmo2.Visible = True
txt_r_cvppcmo2.Visible = True
txt_r_cvppsmo2.Visible = True
txt_r_cvcmo3.Visible = True
txt_r_cvsmo3.Visible = True
txt_r_cvppcmo3.Visible = True
txt_r_cvppsmo3.Visible = True
txt_r_cvcmo4.Visible = True
txt_r_cvsmo4.Visible = True
txt_r_cvppcmo4.Visible = True
txt_r_cvppsmo4.Visible = True

ElseIf txt_cantidad_producto = 5 Then

txt_producto2.Visible = True
txt_precio_venta2.Visible = True
txt_materia_prima2.Visible = True
txt_mano_obra2.Visible = True
txt_incidencia_ventas2.Visible = True
txt_producto3.Visible = True
txt_precio_venta3.Visible = True
txt_materia_prima3.Visible = True
txt_mano_obra3.Visible = True
txt_incidencia_ventas3.Visible = True
txt_producto4.Visible = True
txt_precio_venta4.Visible = True
txt_materia_prima4.Visible = True
txt_mano_obra4.Visible = True
txt_incidencia_ventas4.Visible = True
txt_producto5.Visible = True
txt_precio_venta5.Visible = True
txt_materia_prima5.Visible = True
txt_mano_obra5.Visible = True
txt_incidencia_ventas5.Visible = True
txt_r_cvcmo2.Visible = True
txt_r_cvsmo2.Visible = True
txt_r_cvppcmo2.Visible = True
txt_r_cvppsmo2.Visible = True
txt_r_cvcmo3.Visible = True
txt_r_cvsmo3.Visible = True
txt_r_cvppcmo3.Visible = True
txt_r_cvppsmo3.Visible = True
txt_r_cvcmo4.Visible = True
txt_r_cvsmo4.Visible = True
txt_r_cvppcmo4.Visible = True
txt_r_cvppsmo4.Visible = True
txt_r_cvcmo5.Visible = True
txt_r_cvsmo5.Visible = True
txt_r_cvppcmo5.Visible = True
txt_r_cvppsmo5.Visible = True

ElseIf txt_cantidad_producto > 5 Then
MsgBox "La cantidad De Producto son Hasta 5"

End If
 
End If
End Sub


Private Sub txt_capacidad_pago_promedio_corregida_ajustada_Change()

End Sub

Private Sub txt_credito_hipotecario_Change()
txt_credito_hipotecario = Format(txt_credito_hipotecario, "##,##")
End Sub

Private Sub txt_cupo_linea_credito_Change()
txt_cupo_linea_credito = Format(txt_cupo_linea_credito, "##,##")
End Sub

Private Sub txt_deuda_comercial_Change()
txt_deuda_comercial = Format(txt_deuda_comercial, "##,##")
End Sub

Private Sub txt_deuda_consumo_indirecta_Change()
txt_deuda_consumo_indirecta = Format(txt_deuda_consumo_indirecta, "##,##")
End Sub

Private Sub txt_deuda_d10_comercial_Change()
txt_deuda_d10_comercial = Format(txt_deuda_d10_comercial, "##,##")
End Sub

Private Sub txt_deuda_d10_consumo_Change()
txt_deuda_d10_consumo = Format(txt_deuda_d10_consumo, "##,##")
End Sub

Private Sub txt_deuda_d10_hipotecario_Change()
txt_deuda_d10_hipotecario = Format(txt_deuda_d10_hipotecario, "##,##")
End Sub

Private Sub txt_deuda_d10_linea_Change()
txt_deuda_d10_linea = Format(txt_deuda_d10_linea, "##,##")
End Sub

Private Sub txt_deuda_indirecta_vigente_Change()
txt_deuda_indirecta_vigente = Format(txt_deuda_indirecta_vigente, "##,##")
End Sub

Private Sub txt_deuda_vigente_directa_Change()
txt_deuda_vigente_directa = Format(txt_deuda_vigente_directa, "##,##")
End Sub

Private Sub txt_deudas_directas_vig_Change()
txt_deudas_directas_vig = Format(txt_deudas_directas_vig, "##,##")
End Sub

Private Sub txt_ingreso_cantidad_deudas_Change()

If Not IsNumeric(txt_ingreso_cantidad_deudas) Or txt_ingreso_cantidad_deudas = 0 Or txt_ingreso_cantidad_deudas > 6 Then
  MsgBox "El Número de Producto esta entre 1 y 6 y Debe Ser Numerico... Reingrese"
 
Else

txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True
txt_saldo_pendiente1.Visible = True
cbx_prepaga_deuda1.Visible = True



txt_acreedor2.Visible = False
txt_acreedor3.Visible = False
txt_acreedor4.Visible = False
txt_acreedor5.Visible = False
txt_acreedor6.Visible = False

txt_saldo_pendiente2.Visible = False
txt_saldo_pendiente3.Visible = False
txt_saldo_pendiente4.Visible = False
txt_saldo_pendiente5.Visible = False
txt_saldo_pendiente6.Visible = False

txt_tipo_producto2.Visible = False
txt_tipo_producto3.Visible = False
txt_tipo_producto4.Visible = False
txt_tipo_producto5.Visible = False
txt_tipo_producto6.Visible = False



txt_monto_cuota2.Visible = False
txt_monto_cuota3.Visible = False
txt_monto_cuota4.Visible = False
txt_monto_cuota5.Visible = False
txt_monto_cuota6.Visible = False


txt_cuotas_pactadas2.Visible = False
txt_cuotas_pactadas3.Visible = False
txt_cuotas_pactadas4.Visible = False
txt_cuotas_pactadas5.Visible = False
txt_cuotas_pactadas6.Visible = False


txt_cuotas_pendientes2.Visible = False
txt_cuotas_pendientes3.Visible = False
txt_cuotas_pendientes4.Visible = False
txt_cuotas_pendientes5.Visible = False
txt_cuotas_pendientes6.Visible = False


cbx_prepaga_deuda2.Visible = False
cbx_prepaga_deuda3.Visible = False
cbx_prepaga_deuda4.Visible = False
cbx_prepaga_deuda5.Visible = False
cbx_prepaga_deuda6.Visible = False


If txt_ingreso_cantidad_deudas = 2 Then

txt_acreedor2.Visible = True
txt_tipo_producto2.Visible = True
txt_monto_cuota2.Visible = True
txt_cuotas_pactadas2.Visible = True
txt_cuotas_pendientes2.Visible = True
txt_saldo_pendiente2.Visible = True
cbx_prepaga_deuda2.Visible = True


txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True
txt_saldo_pendiente1.Visible = True
cbx_prepaga_deuda1.Visible = True


ElseIf txt_ingreso_cantidad_deudas = 3 Then

txt_acreedor3.Visible = True
txt_tipo_producto3.Visible = True
txt_monto_cuota3.Visible = True
txt_cuotas_pactadas3.Visible = True
txt_cuotas_pendientes3.Visible = True
txt_saldo_pendiente3.Visible = True
cbx_prepaga_deuda3.Visible = True

txt_acreedor2.Visible = True
txt_tipo_producto2.Visible = True
txt_monto_cuota2.Visible = True
txt_cuotas_pactadas2.Visible = True
txt_cuotas_pendientes2.Visible = True
txt_saldo_pendiente2.Visible = True
cbx_prepaga_deuda2.Visible = True

txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True
txt_saldo_pendiente1.Visible = True
cbx_prepaga_deuda1.Visible = True


ElseIf txt_ingreso_cantidad_deudas = 4 Then

txt_acreedor4.Visible = True
txt_tipo_producto4.Visible = True
txt_monto_cuota4.Visible = True
txt_cuotas_pactadas4.Visible = True
txt_cuotas_pendientes4.Visible = True
txt_saldo_pendiente4.Visible = True
cbx_prepaga_deuda4.Visible = True

txt_acreedor2.Visible = True
txt_acreedor3.Visible = True
txt_tipo_producto3.Visible = True
txt_monto_cuota3.Visible = True
txt_cuotas_pactadas3.Visible = True
txt_cuotas_pendientes3.Visible = True
txt_saldo_pendiente3.Visible = True
cbx_prepaga_deuda3.Visible = True

txt_tipo_producto2.Visible = True
txt_monto_cuota2.Visible = True
txt_cuotas_pactadas2.Visible = True
txt_cuotas_pendientes2.Visible = True
txt_saldo_pendiente2.Visible = True
cbx_prepaga_deuda2.Visible = True

txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True
txt_saldo_pendiente1.Visible = True
cbx_prepaga_deuda1.Visible = True


ElseIf txt_ingreso_cantidad_deudas = 5 Then

txt_acreedor5.Visible = True
txt_tipo_producto5.Visible = True
txt_monto_cuota5.Visible = True
txt_cuotas_pactadas5.Visible = True
txt_cuotas_pendientes5.Visible = True
txt_saldo_pendiente5.Visible = True
cbx_prepaga_deuda5.Visible = True

cbx_prepaga_deuda4.Visible = True
cbx_prepaga_deuda3.Visible = True
cbx_prepaga_deuda2.Visible = True
cbx_prepaga_deuda1.Visible = True

txt_acreedor2.Visible = True
txt_saldo_pendiente2.Visible = True
txt_acreedor3.Visible = True
txt_saldo_pendiente3.Visible = True
txt_acreedor4.Visible = True
txt_saldo_pendiente4.Visible = True
txt_tipo_producto4.Visible = True
txt_monto_cuota4.Visible = True
txt_cuotas_pactadas4.Visible = True
txt_cuotas_pendientes4.Visible = True
txt_tipo_producto3.Visible = True
txt_monto_cuota3.Visible = True
txt_cuotas_pactadas3.Visible = True
txt_cuotas_pendientes3.Visible = True
txt_tipo_producto2.Visible = True
txt_monto_cuota2.Visible = True
txt_cuotas_pactadas2.Visible = True
txt_cuotas_pendientes2.Visible = True
txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True

ElseIf txt_ingreso_cantidad_deudas = 6 Then

txt_acreedor6.Visible = True
txt_tipo_producto6.Visible = True
txt_monto_cuota6.Visible = True
txt_cuotas_pactadas6.Visible = True
txt_cuotas_pendientes6.Visible = True
txt_saldo_pendiente6.Visible = True
cbx_prepaga_deuda6.Visible = True
cbx_prepaga_deuda5.Visible = True
cbx_prepaga_deuda4.Visible = True
cbx_prepaga_deuda3.Visible = True
cbx_prepaga_deuda2.Visible = True
cbx_prepaga_deuda1.Visible = True


txt_acreedor2.Visible = True
txt_saldo_pendiente2.Visible = True
txt_acreedor3.Visible = True
txt_saldo_pendiente3.Visible = True
txt_acreedor4.Visible = True
txt_saldo_pendiente4.Visible = True
txt_acreedor5.Visible = True
txt_tipo_producto5.Visible = True
txt_saldo_pendiente5.Visible = True
txt_monto_cuota5.Visible = True
txt_cuotas_pactadas5.Visible = True
txt_cuotas_pendientes5.Visible = True
txt_tipo_producto4.Visible = True
txt_monto_cuota4.Visible = True
txt_cuotas_pactadas4.Visible = True
txt_cuotas_pendientes4.Visible = True
txt_tipo_producto3.Visible = True
txt_monto_cuota3.Visible = True
txt_cuotas_pactadas3.Visible = True
txt_cuotas_pendientes3.Visible = True
txt_tipo_producto2.Visible = True
txt_monto_cuota2.Visible = True
txt_cuotas_pactadas2.Visible = True
txt_cuotas_pendientes2.Visible = True
txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True

ElseIf txt_cantidad_producto > 5 Then
MsgBox "La cantidad De Producto son Hasta 6"


End If
 
End If

End Sub

Private Sub txt_mto_bruto_sol_cliente_Change()
If txt_monto_maximo_credito > txt_mto_bruto_sol_cliente Then
    txt_r_leverage = "A"
  'lbl_accion.BackColor = &HFF&       'rojo
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else
  
  txt_r_leverage = "R"
  'lbl_accion.BackColor = &HC000&
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
End If
End Sub

Private Sub txt_porcentaje_compra_formal_Change()
If txt_porcentaje_compra_formal > 1 Then
LBL_ALARMA_PORCENTAJE_COMPRA_FORMAL.Visible = True
End If

If txt_porcentaje_compra_formal > 5 Then
   txt_factor_ajuste_compra_tot_iva = "ZG"
ElseIf txt_porcentaje_compra_formal <= 5 Then
    txt_factor_ajuste_compra_tot_iva = "A"
End If

End Sub

Private Sub txt_r_venta_total_met_iva_Change()

End Sub

Private Sub txt_total_deuda_d10_Change()
txt_total_deuda_d10 = Format(txt_total_deuda_d10, "##,##")
End Sub

Private Sub txt_total_deudas_sbif_Change()
txt_total_deudas_sbif = Format(txt_total_deudas_sbif, "##,##")
End Sub

Private Sub txt_total_saldo_pendiente_Change()

End Sub

Private Sub txt_total_vta_total_AfterUpdate()
If (txt_total_vta_total / Metodologia_IVA.txt_valor_uf) > 2400 Then
  
  txt_r_venta_total_met_iva = "R"
  'lbl_accion.BackColor = &HFF&       'rojo
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else
  
  txt_r_venta_total_met_iva = "A"
  'lbl_accion.BackColor = &HC000&
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
End Sub

Private Sub txt_venta_total_Change()

If (txt_venta_total / Metodologia_IVA.txt_valor_uf) > 2400 Then
        txt_r_venta_total_max = "R"
Else
        txt_r_venta_total_max = "A"
End If


If (txt_venta_total / Metodologia_IVA.txt_valor_uf) < 120 Then

        txt_r_venta_total_min = "R"
Else
        txt_r_venta_total_min = "A"
End If


End Sub

Private Sub txt_venta_total_promedio_anual_Change()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
MsgBox ("Boton Deshabilitado Ocupe Opciones De Menu")
Cancel = True
End If
End Sub

Private Sub UserForm_Initialize()

 
cbx_mes_inicio_iva.AddItem "Enero"
cbx_mes_inicio_iva.AddItem "Febrero"
cbx_mes_inicio_iva.AddItem "Marzo"
cbx_mes_inicio_iva.AddItem "Abril"
cbx_mes_inicio_iva.AddItem "Mayo"
cbx_mes_inicio_iva.AddItem "Junio"
cbx_mes_inicio_iva.AddItem "Julio"
cbx_mes_inicio_iva.AddItem "Agosto"
cbx_mes_inicio_iva.AddItem "Septie"
cbx_mes_inicio_iva.AddItem "Octubre"
cbx_mes_inicio_iva.AddItem "Noviem"
cbx_mes_inicio_iva.AddItem "Diciemb"


cbx_prepaga_deuda1.AddItem "Si"
cbx_prepaga_deuda1.AddItem "No"

cbx_prepaga_deuda2.AddItem "Si"
cbx_prepaga_deuda2.AddItem "No"

cbx_prepaga_deuda3.AddItem "Si"
cbx_prepaga_deuda3.AddItem "No"

cbx_prepaga_deuda4.AddItem "Si"
cbx_prepaga_deuda4.AddItem "No"

cbx_prepaga_deuda5.AddItem "Si"
cbx_prepaga_deuda5.AddItem "No"

cbx_prepaga_deuda6.AddItem "Si"
cbx_prepaga_deuda6.AddItem "No"



End Sub


Private Sub txt_producto1_Change()
Dim I As Integer

txt_producto1 = UCase(txt_producto1)
I = Len(txt_producto1)
txt_producto1.SelStart = I
End Sub

Private Sub txt_producto2_Change()
txt_producto2 = UCase(txt_producto2)
I = Len(txt_producto2)
txt_producto2.SelStart = I
End Sub

Private Sub txt_producto3_Change()
txt_producto3 = UCase(txt_producto3)
I = Len(txt_producto3)
txt_producto3.SelStart = I
End Sub

Private Sub txt_producto4_Change()
txt_producto4 = UCase(txt_producto4)
I = Len(txt_producto4)
txt_producto4.SelStart = I
End Sub

Private Sub txt_producto5_Change()
txt_producto5 = UCase(txt_producto5)
I = Len(txt_producto5)
txt_producto5.SelStart = I
End Sub

Private Sub txt_tipo_producto1_Change()
txt_tipo_producto1 = UCase(txt_tipo_producto1)
I = Len(txt_tipo_producto1)
txt_tipo_producto1.SelStart = I
End Sub

Private Sub txt_tipo_producto2_Change()
txt_tipo_producto2 = UCase(txt_tipo_producto2)
I = Len(txt_tipo_producto2)
txt_tipo_producto2.SelStart = I
End Sub

Private Sub txt_tipo_producto3_Change()
txt_tipo_producto3 = UCase(txt_tipo_producto3)
I = Len(txt_tipo_producto3)
txt_tipo_producto3.SelStart = I
End Sub

Private Sub txt_tipo_producto4_Change()
txt_tipo_producto4 = UCase(txt_tipo_producto4)
I = Len(txt_tipo_producto4)
txt_tipo_producto4.SelStart = I
End Sub

Private Sub txt_tipo_producto5_Change()
txt_tipo_producto5 = UCase(txt_tipo_producto5)
I = Len(txt_tipo_producto5)
txt_tipo_producto5.SelStart = I
End Sub

Private Sub txt_tipo_producto6_Change()
txt_tipo_producto6 = UCase(txt_tipo_producto6)
I = Len(txt_tipo_producto6)
txt_tipo_producto6.SelStart = I
End Sub
