VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Estado_Filtros 
   Caption         =   "::::::: Estado De Filtros"
   ClientHeight    =   9915.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   OleObjectBlob   =   "Estado_Filtros.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Estado_Filtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_volver_evaluacion_Click()


If Estado_Filtros.TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Then

Metodologia_Activo_Circulante.Show
Unload cmd_volver_evaluacion

ElseIf Estado_Filtros.TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Then

Metodologia_IVA1.Show
Unload cmd_volver_evaluacion

ElseIf Estado_Filtros.TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then

Metodologia_Maxima_Prod.Show
Unload cmd_volver_evaluacion

End If

End Sub

Private Sub cmd_volver_ficha_Click()

End Sub

Private Sub TXT_ESTADO_METODOLOGIA_OCUPADA_Change()

If TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante" Or TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
    lbl_r_f_solo_iva.Visible = False
    txt_r_f_factor_ajuste_compra_tot_iva.Visible = False
 
 ElseIf TXT_ESTADO_METODOLOGIA_OCUPADA = "IVA" Or TXT_ESTADO_METODOLOGIA_OCUPADA = "Máxima Producción" Then
    lbl_r_f_compra_tot_AC.Visible = False
    txt_r_f_compra_tot_mensual.Visible = False

End If

End Sub

Private Sub txt_r_f_compra_tot_mensual_Change()

End Sub

Private Sub txt_resultado_APROBADO_final_cred_Change()

If txt_campana = "SI" Then
    txt_campana.BackColor = &HC000&
    txt_campana.ForeColor = &H8000000E  'blanco

  Else
    txt_campana.BackColor = &HFF&
    txt_campana.ForeColor = &H8000000E

End If

End Sub
