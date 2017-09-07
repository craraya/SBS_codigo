VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Credito_Consumo 
   Caption         =   "::::: Menu Credito Consumo"
   ClientHeight    =   9075.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   OleObjectBlob   =   "Credito_Consumo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Credito_Consumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guarda_negociador_consumo_Click()

' La conexión a la base de datos
     
Call conectarBD

irespuesta = MsgBox("¿Esta Seguro que Desea Guardar la Evaluacion Consumo?", vbYesNo)
    
    
 If irespuesta = vbYes Then
   ssql = "SELECT rut_cliente, max(n_solicitud) as n_solicitud  " _
      & " FROM tbl_micro_ficha_cliente where rut_cliente = '" & txt_rut_cliente_negociador & "' group by rut_cliente"
           
        Set rst = cnn.Execute(ssql, , adCmdText)
      
        If rst.EOF Then
           MsgBox ("El Rut del Cliente NO corresponde a la Evaluacion Consumo Vigente")
          Else
            If rst!rut_cliente = txt_rut_cliente_negociador Then
              txt_n_solicitud = rst!n_solicitud
            End If
          rst.MoveNext
        End If
'----------------------------------------------------------
    ssql = "INSERT INTO TBL_MICRO_EVALUACION_CONSUMO " _
    & "([Rut_Cliente], [n_solicitud],[metodologia], [Monto_Comercial],[plazo_comercial],[cuota_comercial],[monto_consumo]," _
    & " [plazo_consumo],[cuota_consumo],[total_monto_comercial_consumo],[total_cuota_comercial_consumo],[monto_limite_cliente],[plazo_limite_consumo],[Resolucion_Monto],[Resolucion_Plazo],[Resolucion_Cuota],[fecha_evaluacion],[hora_evaluacion])" _
    & " VALUES (('" & txt_rut_cliente_negociador & "'), ('" & txt_n_solicitud & "')" _
    & ",('" & txt_metologia_negociador & "') , ('" & txt_monto_comercial & "'), ('" & txt_plazo_comercial & "'),('" & txt_cuota_comercial & "')" _
    & ",('" & txt_monto_consumo & "'), ('" & txt_plazo_consumo & "'),('" & txt_cuota_consumo & "')" _
    & ",('" & txt_total_monto_comercial_consumo & "'), ('" & txt_total_cuota_comercial_consumo & "'),('" & txt_monto_limite_cliente & "'), ('" & txt_limite_plazo_consumo & "'),('" & txt_resolucion_monto & "'),('" & txt_resolucion_plazo & "'),('" & txt_resolucion_cuota & "'),('" & txt_fecha_actual & "'),('" & txt_hora_actual & "'))"
    
    cnn.Execute ssql
   
    MsgBox "Evaluacion Consumo Guardada"
    
End If



End Sub

Private Sub cmd_imprimir_negociador_Click()

Credito_Consumo.PrintForm

End Sub

Private Sub cmd_simula_Click()

    Dim fec1
    Dim hora1

    fec1 = Format(Date, "yyyy/mm/dd")
    txt_fecha_actual = fec1

    hora1 = hora
    txt_hora_actual = Time


txt_moitvo_rechazo1_ELM.Visible = False
txt_moitvo_rechazo2_ELMC.Visible = False
txt_moitvo_rechazo3_ELC.Visible = False
txt_moitvo_rechazo4_ELPC.Visible = False
txt_moitvo_rechazo5_NACCCOM.Visible = False
txt_moitvo_rechazo6_NACCCON.Visible = False


txt_r_monto_limite_cliente = "R"
txt_r_cuota_limite_cliente = "R"
txt_r_resolucion_plazo = "R"
txt_r_monto_comercial = "R"
txt_r_monto_consumo = "R"


txt_monto_limite_cliente.BackColor = &HC0C0C0

txt_total_monto_comercial_consumo = Val(txt_monto_comercial) * 1 + Val(txt_monto_consumo) * 1
txt_total_cuota_comercial_consumo = Val(txt_cuota_comercial) * 1 + Val(txt_cuota_consumo) * 1
txt_monto_consumo.BackColor = &HC0C0C0

'''''''

'''''''
If Val(txt_monto_consumo) * 1 > Val(5000000) * 1 Then
   txt_r_monto_maximo_consumo.BackColor = &HFF& ' rojo
   txt_r_monto_maximo_consumo = "R"
   txt_moitvo_rechazo2_ELMC.Visible = True
   
 Else
        txt_monto_consumo.BackColor = &HC000&     ' verde
        txt_r_monto_maximo_consumo.BackColor = &HC000&     ' verde
        txt_r_monto_maximo_consumo = "A"
        txt_moitvo_rechazo2_ELMC.Visible = False
   
End If



If Val(txt_total_monto_comercial_consumo) * 1 > Val(txt_monto_limite_cliente) * 1 Then
    txt_monto_limite_cliente.BackColor = &HFF& ' rojo
    txt_r_monto_limite_cliente = "R"
    txt_moitvo_rechazo1_ELM.Visible = True
Else
    txt_monto_limite_cliente.BackColor = &HC000&     ' verde
    txt_r_monto_limite_cliente = "A"
    txt_moitvo_rechazo1_ELM.Visible = False
End If




'''''''
If Val(txt_cuota_limite_cliente) * 1 > Val(txt_total_cuota_comercial_consumo) * 1 Then
    txt_r_cuota_limite_cliente.BackColor = &HC000&     ' verde
    txt_r_cuota_limite_cliente = "A"
    txt_moitvo_rechazo3_ELC.Visible = False
Else
    txt_r_cuota_limite_cliente = "R"
    txt_cuota_limite_cliente.BackColor = &HFF& ' rojo
    txt_moitvo_rechazo3_ELC.Visible = True
    
End If

'''''''
If Val(txt_plazo_consumo) * 1 > Val(txt_limite_plazo_consumo) * 1 Then
    txt_r_resolucion_plazo = "R"
    txt_moitvo_rechazo4_ELPC.Visible = True

    Else
        txt_r_resolucion_plazo = "A"
        txt_r_resolucion_plazo.BackColor = &HC000&     ' verde
        txt_moitvo_rechazo4_ELPC.Visible = False
End If

'''''''
If Val(txt_monto_comercial) * 1 < Val(Metodologia_Activo_Circulante.txt_saldo_deuda_con_prepago_comercial) * 1 Then
       txt_monto_comercial.BackColor = &HFF& ' rojo
       txt_r_monto_comercial = "R"
       txt_moitvo_rechazo5_NACCCOM.Visible = True
Else
        txt_r_monto_consumo.BackColor = &HC000&     ' verde
        txt_moitvo_rechazo5_NACCCOM.Visible = False
        txt_r_monto_comercial = "A"
End If

'''''''
If Val(txt_monto_consumo) * 1 < Val(Metodologia_Activo_Circulante.txt_saldo_deuda_con_prepago_consumo) * 1 Then
            txt_monto_consumo.BackColor = &HFF& ' rojo
            txt_r_monto_consumo = "R"
            txt_moitvo_rechazo6_NACCCON.Visible = True
Else
        txt_moitvo_rechazo6_NACCCON.Visible = False
        txt_r_monto_consumo.BackColor = &HC000&     ' verde
        txt_r_monto_consumo = "A"
        
End If

'''''''
If txt_r_monto_limite_cliente <> "R" And txt_r_monto_comercial <> "R" And txt_r_monto_consumo <> "R" Then
txt_resolucion_monto = "Aprobado"
Else
txt_resolucion_monto = "Rechazado"
End If

'''''''
If txt_r_resolucion_plazo <> "R" Then
txt_resolucion_plazo = "Aprobado"
Else
txt_resolucion_plazo = "Rechazado"
End If

'''''''
If txt_r_cuota_limite_cliente <> "R" Then
txt_resolucion_cuota = "Aprobado"
Else
txt_resolucion_cuota = "Rechazado"
End If

'''Resolucion final de negociador

If txt_r_monto_limite_cliente = "R" Or txt_r_cuota_limite_cliente = "R" Or txt_r_resolucion_plazo = "R" Or txt_r_monto_comercial = "R" Or txt_r_monto_consumo = "R" Or txt_r_monto_maximo_consumo = "R" Then

txt_resolucion_final_negociador = "Rechazo"
txt_resolucion_final_negociador.BackColor = &HFF& ' rojo

Else

txt_resolucion_final_negociador.BackColor = &HC000&     ' verde
txt_resolucion_final_negociador = "Aprobado"

End If

'''PASA RESOLUCION FINAL A HOJA DE ESTADO_RESOLUCION_FINAL

Estado_Resolucion_Final.txt_r_f_Monto_Limite_consumo = Credito_Consumo.txt_r_monto_limite_cliente
Estado_Resolucion_Final.txt_r_f_capacidad_pago_consumo = Credito_Consumo.txt_r_cuota_limite_cliente
Estado_Resolucion_Final.txt_r_f_plazo_consumo = Credito_Consumo.txt_r_resolucion_plazo
Estado_Resolucion_Final.txt_r_f_mto_max_consumo = Credito_Consumo.txt_r_monto_maximo_consumo
Estado_Resolucion_Final.txt_r_f_min_prepago_consumo = Credito_Consumo.txt_r_monto_consumo
Estado_Resolucion_Final.txt_r_f_min_prepago_comercial = Credito_Consumo.txt_r_monto_comercial



End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label25_Click()

End Sub



Private Sub cmd_volver_evaluacion_Click()

If txt_metologia_negociador = "Activo Circulante" Then

Metodologia_Activo_Circulante.txt_cuota_credito.Locked = True
Metodologia_Activo_Circulante.txt_mto_bruto_sol_cliente.Locked = True
Metodologia_Activo_Circulante.cmd_guardar_evaluacion.Enabled = False
Credito_Consumo.Hide
Metodologia_Activo_Circulante.Show

ElseIf txt_metologia_negociador = "Iva" Then

Metodologia_IVA1.txt_cuota_credito.Locked = True
Metodologia_IVA1.txt_mto_bruto_sol_cliente.Locked = True
Metodologia_IVA1.cmd_guardar_evaluacion.Enabled = False
Credito_Consumo.Hide
Metodologia_IVA1.Show

ElseIf txt_metologia_negociador = "Maxima Produccion" Then
Metodologia_Maxima_Prod.txt_cuota_credito.Locked = True
Metodologia_Maxima_Prod.txt_mto_bruto_sol_cliente.Locked = True
Metodologia_Maxima_Prod.cmd_guardar_evaluacion.Enabled = False
Credito_Consumo.Hide
Metodologia_Maxima_Prod.Show

End If

End Sub

Private Sub Label35_Click()

End Sub

Private Sub txt_cuota_comercial_Change()
txt_total_cuota_comercial_consumo = 0
End Sub
Private Sub txt_cuota_consumo_Change()
txt_total_cuota_comercial_consumo = 0
End Sub

Private Sub txt_metologia_negociador_Change()

End Sub

Private Sub txt_monto_comercial_Change()
txt_total_monto_comercial_consumo = 0
End Sub

Private Sub txt_monto_consumo_AfterUpdate()



End Sub

Private Sub txt_monto_consumo_Change()
txt_total_monto_comercial_consumo = 0
End Sub

Private Sub txt_rut_cliente_negociador_Change()

End Sub
