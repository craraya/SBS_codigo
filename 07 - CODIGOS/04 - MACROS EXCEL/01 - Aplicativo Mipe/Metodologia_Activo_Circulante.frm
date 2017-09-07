VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Metodologia_Activo_Circulante 
   Caption         =   "::::: Metodologia Activo Circulante"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13020
   OleObjectBlob   =   "Metodologia_Activo_Circulante.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Metodologia_Activo_Circulante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbx_prepaga_deuda1_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_prepaga_deuda2_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_prepaga_deuda3_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_prepaga_deuda4_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_prepaga_deuda5_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_prepaga_deuda6_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_tipo_credito_deuda1_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0



txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub cbx_tipo_credito_deuda2_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0


txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub cbx_tipo_credito_deuda3_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0


txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub cbx_tipo_credito_deuda4_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0


txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub cbx_tipo_credito_deuda5_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub cbx_tipo_credito_deuda6_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub cmd_calcula_costos_fijos_Click()

txt_total_costos_fijos = Int((Val(txt_arriendo_micro) + Val(txt_sueldos) + Val(txt_movilizacion) + _
Val(txt_servicios_basicos) + Val(txt_contador) + Val(txt_lubricantes) + _
Val(txt_neumaticos) + Val(txt_afinamientos) + Val(txt_patentes_seguros) + Val(txt_otros_costos_fijos) + Val(txt_impuesto)) * 1.15)

'PRENDE BOTON
cmd_calcula_gastos_familiares.Enabled = True

'End If
End Sub

Private Sub cmd_calcula_deudas_Click()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0




cmd_calcular_flujo_Caja.Enabled = False


'----1
If cbx_tipo_credito_deuda1 = "Consumo" Or cbx_tipo_credito_deuda1 = "Comercial" Then

    If txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
        Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
        Or txt_ingreso_cantidad_deudas = 6) _
        And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then
    
    If Val(txt_cuotas_pendientes1) <= 3 Or cbx_prepaga_deuda1 = "Si" Then
 
        txt_no_sumar_mto_cuota1 = Val(txt_monto_cuota1)
        txt_sumar_mto_cuota1 = 0
   
   Else
        txt_sumar_mto_cuota1 = Val(txt_monto_cuota1)
        txt_no_sumar_mto_cuota1 = 0
   End If

  End If
End If


If cbx_tipo_credito_deuda1 = "Consumo" Then

If txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) _
   And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then
    
    If cbx_prepaga_deuda1 = "Si" And cbx_tipo_credito_deuda1 = "Consumo" Then
        txt_saldo_deuda_con_prepago_consumo1 = Val(txt_saldo_pendiente1)
        txt_mto_cuota_con_prepago_consumo1 = Val(txt_monto_cuota1)
        
    ElseIf cbx_prepaga_deuda1 = "No" And cbx_tipo_credito_deuda1 = "Consumo" Then
        txt_saldo_deuda_sin_prepago_consumo1 = Val(txt_saldo_pendiente1)
        txt_mto_cuota_sin_prepago_consumo1 = Val(txt_monto_cuota1)
       
    End If
        
        txt_sumar_mto_cuota1_consumo = Val(txt_monto_cuota1)
        txt_sumar_mto_deuda1_consumo = Val(txt_saldo_pendiente1)
        txt_no_sumar_mto_cuota1_consumo = 0
   

  End If
End If

If cbx_tipo_credito_deuda1 = "Comercial" Then

If txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) _
And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then
    
    
    If cbx_prepaga_deuda1 = "Si" And cbx_tipo_credito_deuda1 = "Comercial" Then
        txt_saldo_deuda_con_prepago_comercial1 = Val(txt_saldo_pendiente1)
        txt_mto_cuota_con_prepago_comercial1 = Val(txt_monto_cuota1)
        
    ElseIf cbx_prepaga_deuda1 = "No" And cbx_tipo_credito_deuda1 = "Comercial" Then
        txt_saldo_deuda_sin_prepago_comercial1 = Val(txt_saldo_pendiente1)
        txt_mto_cuota_sin_prepago_comercial1 = Val(txt_monto_cuota1)
       
    End If
   
        txt_sumar_mto_cuota1_comercial = Val(txt_monto_cuota1)
        txt_sumar_mto_deuda1_comecial = Val(txt_saldo_pendiente1)
       txt_no_sumar_mto_cuota1_comercial = 0
   

    End If
End If

'------2
If cbx_tipo_credito_deuda2 = "Consumo" Or cbx_tipo_credito_deuda2 = "Comercial" Then

If txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) _
And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

      If txt_cuotas_pendientes2 <= 3 Or cbx_prepaga_deuda2 = "Si" Then
       
          txt_no_sumar_mto_cuota2 = Val(txt_monto_cuota2)
          txt_sumar_mto_cuota2 = 0
          
          Else
          txt_sumar_mto_cuota2 = Val(txt_monto_cuota2)
          txt_no_sumar_mto_cuota2 = 0
       End If

End If
End If

If cbx_tipo_credito_deuda2 = "Consumo" Then

If txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) _
And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

          If cbx_prepaga_deuda2 = "Si" And cbx_tipo_credito_deuda2 = "Consumo" Then
        txt_saldo_deuda_con_prepago_consumo2 = Val(txt_saldo_pendiente2)
        txt_mto_cuota_con_prepago_consumo2 = Val(txt_monto_cuota2)
        
        ElseIf cbx_prepaga_deuda2 = "No" And cbx_tipo_credito_deuda2 = "Consumo" Then
        txt_saldo_deuda_sin_prepago_consumo2 = Val(txt_saldo_pendiente2)
        txt_mto_cuota_sin_prepago_consumo2 = Val(txt_monto_cuota2)

          txt_sumar_mto_cuota2_consumo = 0
          
        End If
          txt_sumar_mto_cuota2_consumo = Val(txt_monto_cuota2)
          txt_sumar_mto_deuda2_consumo = Val(txt_saldo_pendiente2)
          txt_no_sumar_mto_cuota2_consumo = 0
       End If

End If


If cbx_tipo_credito_deuda2 = "Comercial" Then

If txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) _
And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

        If cbx_prepaga_deuda2 = "Si" And cbx_tipo_credito_deuda2 = "Comercial" Then
        txt_saldo_deuda_con_prepago_comercial2 = Val(txt_saldo_pendiente2)
        txt_mto_cuota_con_prepago_comercial2 = Val(txt_monto_cuota2)
        
        ElseIf cbx_prepaga_deuda2 = "No" And cbx_tipo_credito_deuda2 = "Comercial" Then
        txt_saldo_deuda_sin_prepago_comercial2 = Val(txt_saldo_pendiente2)
        txt_mto_cuota_sin_prepago_comercial2 = Val(txt_monto_cuota2)
        
        End If
        
          txt_sumar_mto_deuda2_comecial = Val(txt_saldo_pendiente2)
          txt_sumar_mto_cuota2_comercial = Val(txt_monto_cuota2)
          txt_no_sumar_mto_cuota2_comercial = 0

    End If
End If



'-----3
If cbx_tipo_credito_deuda3 = "Consumo" Or cbx_tipo_credito_deuda3 = "Comercial" Then

If txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

      If txt_cuotas_pendientes3 <= 3 Or cbx_prepaga_deuda3 = "Si" Then
       
              txt_no_sumar_mto_cuota3 = Val(txt_monto_cuota1)
              txt_sumar_mto_cuota3 = 0
              Else
              txt_sumar_mto_cuota3 = Val(txt_monto_cuota3)
              txt_sumar_mto_deuda3 = Val(txt_saldo_pendiente3)
              txt_no_sumar_mto_cuota3 = 0
       End If

       
    End If
End If


If cbx_tipo_credito_deuda3 = "Consumo" Then

If txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

      
        If cbx_prepaga_deuda3 = "Si" And cbx_tipo_credito_deuda3 = "Consumo" Then
            txt_saldo_deuda_con_prepago_consumo3 = Val(txt_saldo_pendiente3)
            txt_mto_cuota_con_prepago_consumo3 = Val(txt_monto_cuota3)
        
        ElseIf cbx_prepaga_deuda3 = "No" And cbx_tipo_credito_deuda3 = "Consumo" Then
            txt_saldo_deuda_sin_prepago_consumo3 = Val(txt_saldo_pendiente3)
            txt_mto_cuota_sin_prepago_consumo3 = Val(txt_monto_cuota3)
            
            txt_sumar_mto_cuota3_consumo = 0
        End If
              txt_sumar_mto_cuota3_consumo = Val(txt_monto_cuota3)
              txt_sumar_mto_deuda3_consumo = Val(txt_saldo_pendiente3)
              txt_no_sumar_mto_cuota3_consumo = 0


End If
End If

If cbx_tipo_credito_deuda3 = "Comercial" Then

If txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

          If cbx_prepaga_deuda3 = "Si" And cbx_tipo_credito_deuda3 = "Comercial" Then
            txt_saldo_deuda_con_prepago_comercial3 = Val(txt_saldo_pendiente3)
            txt_mto_cuota_con_prepago_comercial3 = Val(txt_monto_cuota3)
        
            ElseIf cbx_prepaga_deuda3 = "No" And cbx_tipo_credito_deuda3 = "Comercial" Then
            txt_saldo_deuda_sin_prepago_comercial3 = Val(txt_saldo_pendiente3)
            txt_mto_cuota_sin_prepago_comercial3 = Val(txt_monto_cuota3)
        
            txt_sumar_mto_cuota3_comercial = 0
            
          End If
            
              txt_sumar_mto_cuota3_comercial = Val(txt_monto_cuota3)
              txt_sumar_mto_deuda3_comecial = Val(txt_saldo_pendiente3)
              txt_no_sumar_mto_cuota3_comercial = 0

End If
End If


'---4

If cbx_tipo_credito_deuda4 = "Consumo" Or cbx_tipo_credito_deuda4 = "Comercial" Then

If txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then


       If txt_cuotas_pendientes4 <= 3 Or cbx_prepaga_deuda4 = "Si" Then
       
              txt_no_sumar_mto_cuota4 = Val(txt_monto_cuota4)
              txt_sumar_mto_cuota4 = 0
              Else
              txt_sumar_mto_cuota4 = Val(txt_monto_cuota4)
              txt_no_sumar_mto_cuota4 = 0
       End If

       

End If
End If


If cbx_tipo_credito_deuda4 = "Consumo" Then

If txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then


           If cbx_prepaga_deuda4 = "Si" And cbx_tipo_credito_deuda4 = "Consumo" Then
                txt_saldo_deuda_con_prepago_consumo4 = Val(txt_saldo_pendiente4)
                txt_mto_cuota_con_prepago_consumo4 = Val(txt_monto_cuota4)
        
            ElseIf cbx_prepaga_deuda4 = "No" And cbx_tipo_credito_deuda4 = "Consumo" Then
                txt_saldo_deuda_sin_prepago_consumo4 = Val(txt_saldo_pendiente4)
                txt_mto_cuota_sin_prepago_consumo4 = Val(txt_monto_cuota4)
                
                txt_sumar_mto_cuota4_consumo = 0
            End If
              
              
              txt_sumar_mto_cuota4_consumo = Val(txt_monto_cuota4)
              txt_sumar_mto_deuda4_consumo = Val(txt_saldo_pendiente4)
              txt_no_sumar_mto_cuota4_consumo = 0
       
End If
End If


If cbx_tipo_credito_deuda4 = "Comercial" Then

If txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then


            If cbx_prepaga_deuda4 = "Si" And cbx_tipo_credito_deuda4 = "Comercial" Then
                txt_saldo_deuda_con_prepago_comercial4 = Val(txt_saldo_pendiente4)
                txt_mto_cuota_con_prepago_comercial4 = Val(txt_monto_cuota4)
        
            ElseIf cbx_prepaga_deuda4 = "No" And cbx_tipo_credito_deuda4 = "Comercial" Then
                txt_saldo_deuda_sin_prepago_comercial4 = Val(txt_saldo_pendiente4)
                txt_mto_cuota_sin_prepago_comercial4 = Val(txt_monto_cuota4)
                
                txt_sumar_mto_cuota4_comercial = 0
            
            End If
              
              txt_sumar_mto_cuota4_comercial = Val(txt_monto_cuota4)
              txt_sumar_mto_deuda4_comecial = Val(txt_saldo_pendiente4)
              txt_no_sumar_mto_cuota4_comercial = 0
End If
End If



'-----5
If cbx_tipo_credito_deuda5 = "Consumo" Or cbx_tipo_credito_deuda5 = "Comercial" Then

If txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

              
       If txt_cuotas_pendientes5 <= 3 Or cbx_prepaga_deuda5 = "Si" Then
       
              txt_no_sumar_mto_cuota5 = Val(txt_monto_cuota5)
              txt_sumar_mto_cuota5 = 0
              Else
              txt_sumar_mto_cuota5 = Val(txt_monto_cuota5)
              txt_no_sumar_mto_cuota5 = 0
       End If

       
End If
End If


If cbx_tipo_credito_deuda5 = "Consumo" Then

If txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

              
        If cbx_prepaga_deuda5 = "Si" And cbx_tipo_credito_deuda5 = "Consumo" Then
            txt_saldo_deuda_con_prepago_consumo5 = Val(txt_saldo_pendiente5)
            txt_mto_cuota_con_prepago_consumo5 = Val(txt_monto_cuota5)
        
        ElseIf cbx_prepaga_deuda5 = "No" And cbx_tipo_credito_deuda5 = "Consumo" Then
            txt_saldo_deuda_sin_prepago_consumo5 = Val(txt_saldo_pendiente5)
            txt_mto_cuota_sin_prepago_consumo5 = Val(txt_monto_cuota5)
              txt_sumar_mto_cuota5_consumo = 0
              
        End If
              txt_sumar_mto_cuota5_consumo = Val(txt_monto_cuota5)
              txt_sumar_mto_deuda5_consumo = Val(txt_saldo_pendiente5)
              txt_no_sumar_mto_cuota5_consumo = 0
              
 End If
End If


If cbx_tipo_credito_deuda5 = "Comercial" Then

If txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

              
            If cbx_prepaga_deuda5 = "Si" And cbx_tipo_credito_deuda5 = "Comercial" Then
                txt_saldo_deuda_con_prepago_comercial5 = Val(txt_saldo_pendiente5)
                txt_mto_cuota_con_prepago_comercial5 = Val(txt_monto_cuota5)
        
            ElseIf cbx_prepaga_deuda5 = "No" And cbx_tipo_credito_deuda5 = "Comercial" Then
                txt_saldo_deuda_sin_prepago_comercial5 = Val(txt_saldo_pendiente5)
                txt_mto_cuota_sin_prepago_comercial5 = Val(txt_monto_cuota5)
                
                txt_sumar_mto_cuota5_comercial = 0
              
              End If
              
              txt_sumar_mto_cuota5_comercial = Val(txt_monto_cuota5)
              txt_sumar_mto_deuda5_comecial = Val(txt_saldo_pendiente5)
              txt_no_sumar_mto_cuota5_comercial = 0
       
End If
End If


'------6
If cbx_tipo_credito_deuda6 = "Consumo" Or cbx_tipo_credito_deuda6 = "Comercial" Then

If txt_monto_cuota6 <> 0 And txt_monto_cuota6 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas6) >= Val(txt_cuotas_pendientes6) _
And txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then


       If txt_cuotas_pendientes6 <= 3 Or cbx_prepaga_deuda6 = "Si" Then
       
              txt_no_sumar_mto_cuota6 = Val(txt_monto_cuota6)
              txt_sumar_mto_cuota6 = 0
              Else
              txt_sumar_mto_cuota6 = Val(txt_monto_cuota6)
              txt_no_sumar_mto_cuota6 = 0
       End If
              
       
End If
End If



If cbx_tipo_credito_deuda6 = "Consumo" Then

If txt_monto_cuota6 <> 0 And txt_monto_cuota6 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
  Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
  Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas6) >= Val(txt_cuotas_pendientes6) _
And txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then


           If cbx_prepaga_deuda6 = "Si" And cbx_tipo_credito_deuda6 = "Consumo" Then
                txt_saldo_deuda_con_prepago_consumo6 = Val(txt_saldo_pendiente6)
                txt_mto_cuota_con_prepago_consumo6 = Val(txt_monto_cuota6)
        
            ElseIf cbx_prepaga_deuda6 = "No" And cbx_tipo_credito_deuda6 = "Consumo" Then
                txt_saldo_deuda_sin_prepago_consumo6 = Val(txt_saldo_pendiente6)
                txt_mto_cuota_sin_prepago_consumo6 = Val(txt_monto_cuota6)
                
                txt_sumar_mto_cuota6_consumo = 0
            End If
              
              txt_sumar_mto_cuota6_consumo = Val(txt_monto_cuota6)
              txt_sumar_mto_deuda6_consumo = Val(txt_saldo_pendiente6)
              txt_no_sumar_mto_cuota6_consumo = 0

End If
End If


If cbx_tipo_credito_deuda6 = "Comercial" Then

If txt_monto_cuota6 <> 0 And txt_monto_cuota6 <> "" And (txt_ingreso_cantidad_deudas = 1 Or txt_ingreso_cantidad_deudas = 2 _
    Or txt_ingreso_cantidad_deudas = 3 Or txt_ingreso_cantidad_deudas = 4 Or txt_ingreso_cantidad_deudas = 5 _
    Or txt_ingreso_cantidad_deudas = 6) And Val(txt_cuotas_pactadas6) >= Val(txt_cuotas_pendientes6) _
    And txt_monto_cuota5 <> 0 And txt_monto_cuota5 <> "" And Val(txt_cuotas_pactadas5) >= Val(txt_cuotas_pendientes5) _
    And txt_monto_cuota4 <> 0 And txt_monto_cuota4 <> "" And Val(txt_cuotas_pactadas4) >= Val(txt_cuotas_pendientes4) _
    And txt_monto_cuota3 <> 0 And txt_monto_cuota3 <> "" And Val(txt_cuotas_pactadas3) >= Val(txt_cuotas_pendientes3) _
    And txt_monto_cuota2 <> 0 And txt_monto_cuota2 <> "" And Val(txt_cuotas_pactadas2) >= Val(txt_cuotas_pendientes2) _
    And txt_monto_cuota1 <> 0 And txt_monto_cuota1 <> "" And Val(txt_cuotas_pactadas1) >= Val(txt_cuotas_pendientes1) Then

           If cbx_prepaga_deuda6 = "Si" And cbx_tipo_credito_deuda6 = "Comercial" Then
                txt_saldo_deuda_con_prepago_comercial6 = Val(txt_saldo_pendiente6)
                txt_mto_cuota_con_prepago_comercial6 = Val(txt_monto_cuota6)
        
            ElseIf cbx_prepaga_deuda6 = "No" And cbx_tipo_credito_deuda6 = "Comercial" Then
                txt_saldo_deuda_sin_prepago_comercial6 = Val(txt_saldo_pendiente6)
                txt_mto_cuota_sin_prepago_comercial6 = Val(txt_monto_cuota6)
              
              txt_sumar_mto_cuota6_comercial = 0
            End If
            
              txt_sumar_mto_cuota6_comercial = Val(txt_monto_cuota6)
              txt_sumar_mto_deuda6_comecial = Val(txt_saldo_pendiente6)
              txt_no_sumar_mto_cuota6_comercial = 0

End If
End If

'calculos campo total $$saldo pendiente CONSUMO + COMERCIAL
txt_total_saldo_pendiente = Val(txt_saldo_pendiente1) + Val(txt_saldo_pendiente2) + Val(txt_saldo_pendiente3) + Val(txt_saldo_pendiente4) + _
Val(txt_saldo_pendiente5) + Val(txt_saldo_pendiente6)
'calculos campo total $$cuotas pendiente CONSUMO + COMERCIAL
txt_total_deudas = Val(txt_sumar_mto_cuota1) * 1 + Val(txt_sumar_mto_cuota2) * 1 + Val(txt_sumar_mto_cuota3) * 1 + Val(txt_sumar_mto_cuota4) * 1 + Val(txt_sumar_mto_cuota5) * 1 _
+ Val(txt_sumar_mto_cuota6) * 1

'calculos campo total $$saldo pendiente CONSUMO
txt_total_saldo_pendiente_consumo = Val(txt_sumar_mto_deuda1_consumo) + Val(txt_sumar_mto_deuda2_consumo) + Val(txt_sumar_mto_deuda3_consumo) + Val(txt_sumar_mto_deuda4_consumo) + _
Val(txt_sumar_mto_deuda5_consumo) + Val(txt_sumar_mto_deuda6_consumo)
'calculos campo total $$cuotas pendiente CONSUMO
txt_total_deudas_consumo = Val(txt_sumar_mto_cuota1_consumo) * 1 + Val(txt_sumar_mto_cuota2_consumo) * 1 + Val(txt_sumar_mto_cuota3_consumo) * 1 + Val(txt_sumar_mto_cuota4_consumo) * 1 + Val(txt_sumar_mto_cuota5_consumo) * 1 _
+ Val(txt_sumar_mto_cuota6_consumo) * 1

'calculos campo total $$saldo pendiente COMERCIAL
 txt_total_deudas_comercial = Val(txt_sumar_mto_cuota1_comercial) + Val(txt_sumar_mto_cuota2_comercial) + Val(txt_sumar_mto_cuota3_comercial) + Val(txt_sumar_mto_cuota4_comercial) + _
Val(txt_sumar_mto_cuota5_comercial) + Val(txt_sumar_mto_cuota6_comercial)
'calculos campo total $$cuotas pendiente COMERCIAL
txt_total_saldo_pendiente_comercial = Val(txt_sumar_mto_deuda1_comecial) * 1 + Val(txt_sumar_mto_deuda2_comecial) * 1 + Val(txt_sumar_mto_deuda3_comecial) * 1 + Val(txt_sumar_mto_deuda4_comecial) * 1 + Val(txt_sumar_mto_deuda5_comecial) * 1 _
+ Val(txt_sumar_mto_deuda6_comecial) * 1

'calculo de prepagos con y sin
txt_saldo_deuda_con_prepago_consumo = Val(txt_saldo_deuda_con_prepago_consumo1) * 1 + Val(txt_saldo_deuda_con_prepago_consumo2) * 1 + Val(txt_saldo_deuda_con_prepago_consumo3) * 1 + Val(txt_saldo_deuda_con_prepago_consumo4) * 1 + Val(txt_saldo_deuda_con_prepago_consumo5) * 1 + Val(txt_saldo_deuda_con_prepago_consumo6) * 1
txt_mto_cuota_con_prepago_consumo = Val(txt_mto_cuota_con_prepago_consumo1) * 1 + Val(txt_mto_cuota_con_prepago_consumo2) * 1 + Val(txt_mto_cuota_con_prepago_consumo3) * 1 + Val(txt_mto_cuota_con_prepago_consumo4) * 1 + Val(txt_mto_cuota_con_prepago_consumo5) * 1 + Val(txt_mto_cuota_con_prepago_consumo6) * 1
txt_saldo_deuda_con_prepago_comercial = Val(txt_saldo_deuda_con_prepago_comercial1) * 1 + Val(txt_saldo_deuda_con_prepago_comercial2) * 1 + Val(txt_saldo_deuda_con_prepago_comercial3) * 1 + Val(txt_saldo_deuda_con_prepago_comercial4) * 1 + Val(txt_saldo_deuda_con_prepago_comercial5) * 1 + Val(txt_saldo_deuda_con_prepago_comercial6) * 1
txt_mto_cuota_con_prepago_comercial = Val(txt_mto_cuota_con_prepago_comercial1) * 1 + Val(txt_mto_cuota_con_prepago_comercial2) * 1 + Val(txt_mto_cuota_con_prepago_comercial3) * 1 + Val(txt_mto_cuota_con_prepago_comercial4) * 1 + Val(txt_mto_cuota_con_prepago_comercial5) * 1 + Val(txt_mto_cuota_con_prepago_comercial6) * 1

txt_saldo_deuda_sin_prepago_consumo = Val(txt_saldo_deuda_sin_prepago_consumo1) * 1 + Val(txt_saldo_deuda_sin_prepago_consumo2) * 1 + Val(txt_saldo_deuda_sin_prepago_consumo3) * 1 + Val(txt_saldo_deuda_sin_prepago_consumo4) * 1 + Val(txt_saldo_deuda_sin_prepago_consumo5) * 1 + Val(txt_saldo_deuda_sin_prepago_consumo6) * 1
txt_mto_cuota_sin_prepago_consumo = Val(txt_mto_cuota_sin_prepago_consumo1) * 1 + Val(txt_mto_cuota_sin_prepago_consumo2) * 1 + Val(txt_mto_cuota_sin_prepago_consumo3) * 1 + Val(txt_mto_cuota_sin_prepago_consumo4) * 1 + Val(txt_mto_cuota_sin_prepago_consumo5) * 1 + Val(txt_mto_cuota_sin_prepago_consumo6) * 1
txt_saldo_deuda_sin_prepago_comercial = Val(txt_saldo_deuda_sin_prepago_comercial1) * 1 + Val(txt_saldo_deuda_sin_prepago_comercial2) * 1 + Val(txt_saldo_deuda_sin_prepago_comercial3) * 1 + Val(txt_saldo_deuda_sin_prepago_comercial4) * 1 + Val(txt_saldo_deuda_sin_prepago_comercial5) * 1 + Val(txt_saldo_deuda_sin_prepago_comercial6) * 1
txt_mto_cuota_sin_prepago_comercial = Val(txt_mto_cuota_sin_prepago_comercial1) * 1 + Val(txt_mto_cuota_sin_prepago_comercial2) * 1 + Val(txt_mto_cuota_sin_prepago_comercial3) * 1 + Val(txt_mto_cuota_sin_prepago_comercial4) * 1 + Val(txt_mto_cuota_sin_prepago_comercial5) * 1 + Val(txt_mto_cuota_sin_prepago_comercial6) * 1


'''SUMA PARA EL TOTAL DE DEUDAS
Total_Deuda_SBIF = txt_deuda_consumo * 1 + txt_deuda_comercial * 1 + txt_credito_hipotecario * 1 + txt_cupo_linea_credito * 1 + txt_deuda_indirecta * 1
txt_total_deuda_d10 = txt_deuda_d10_consumo * 1 + txt_deuda_d10_comercial * 1 + txt_deuda_d10_linea * 1 + txt_deuda_d10_hipotecario * 1


'''''COMPARA DEUDA DECLARADA CONTRA DEUDA SBIF (VIGENTE+MOROSA+VENCIDA+CASTIGO)

If txt_total_saldo_pendiente * 1 >= Total_Deuda_SBIF * 1 Then
    txt_r_sbif_declarada = "A"
    Else
    txt_r_sbif_declarada = "ZG"
End If

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

Private Sub cmd_calcula_ingreso_compara_mensual_Click()

txt_r_compra_activo_AC = Empty

cmd_costo_promedio_ponderado.Enabled = False

'Paso De Rut Cliente En Proceso
txt_rut_cliente = rut_cliente_ficha
txt_dv = dv_cliente_ficha

txt_rut_cliente_pag2 = rut_cliente_ficha
txt_dv_pag2 = dv_cliente_ficha

txt_rut_cliente_pag3 = rut_cliente_ficha
txt_dv_pag3 = dv_cliente_ficha

txt_rut_cliente_pag4 = rut_cliente_ficha
txt_dv_pag4 = dv_cliente_ficha

'FIN paso Variable

If Val(txt_compra_promedio_mensual) > 0 And Val(txt_veces_compra_mes) > 0 _
And (txt_caja_banco > 0 Or txt_materia_primas > 0 Or txt_mercaderias > 0 Or txt_cuenta_cobrar > 0 Or txt_otros_activos_circulantes > 0) Then

txt_compra_total_mensual_ctm = Val(txt_compra_promedio_mensual) * Val(txt_veces_compra_mes)
txt_total_activos_circulantes = Val(txt_caja_banco) + Val(txt_materia_primas) + Val(txt_mercaderias) + Val(txt_cuenta_cobrar) + Val(txt_otros_activos_circulantes)

txt_nota_superar60 = (Val(txt_compra_promedio_mensual) / Val(txt_total_activos_circulantes)) * 100

If txt_nota_superar60 > 60 Then
    
        MsgBox "La Compra Promedio No Puede Superar El 60% Del Activo Circulante"
        txt_total_activos_circulantes = Empty
        
    Else
    
    
    txt_40_total_Activos = txt_total_activos_circulantes * 0.4
    txt_60_total_Activos = txt_total_activos_circulantes * 0.6

        If txt_compra_promedio_mensual * 1 >= txt_40_total_Activos * 1 And txt_compra_promedio_mensual * 1 <= txt_60_total_Activos * 1 Then

            txt_r_compra_activo_AC = "ZG"
    
        ElseIf txt_compra_promedio_mensual * 1 < txt_40_total_Activos * 1 Then

            txt_r_compra_activo_AC = "A"

        End If
    
'PRENDE BOTON
cmd_costo_promedio_ponderado.Enabled = True

End If

Else
MsgBox "Los Campos Compra Promedio / Veces De Compras Y A Lo Menos Un Campos Activo Circulante Son De Ingreso Obligatorio"

End If

'''CALCULOS DE ALARMA TOTAL_ACTIVOS




End Sub

Private Sub cmd_calcula_otros_ingresos_Click()
txt_total_otros_ingresos = Val(txt_liquidacion_sueldo) + Val(txt_jubilacion) + Val(txt_montepio) + Val(txt_arriendo_vivienda1) + Val(txt_ingreso_segunda_microempresa) + _
Val(txt_boleta_honorario)

'PRENDE BOTON
cmd_calcula_deudas.Enabled = True

End Sub

Private Sub cmd_calcular_resolucion_cred_Click()

If txt_cuota_credito >= 0 And txt_cuota_credito <> "" And txt_mto_bruto_sol_cliente >= 0 And txt_mto_bruto_sol_cliente <> "" Then


txt_venta_formal_maxima = Val(txt_vta_formal_promedio_mes_alto * numero_meses_tipo_mes_alto) + Val(txt_vta_formal_promedio_mes_medio * numero_meses_tipo_mes_medio) + Val(txt_vta_formal_promedio_mes_bajo * numero_meses_tipo_mes_bajo)


If Val(txt_mto_bruto_sol_cliente) * 1 >= 0 And Val(txt_cuota_credito) * 1 >= 0 _
   And Val(txt_mto_bruto_sol_cliente) * 1 <= Val(txt_monto_maximo_credito) Then
    
    txt_aprobacion = "OK"
    
    Else
     txt_aprobacion = "RECHAZADO"
End If

   
If Val(txt_mto_bruto_sol_cliente) * 1 >= 0 And Val(txt_cuota_credito) * 1 >= 0 _
       And Val(txt_cuota_credito) <= Val(txt_capacidad_pago_promedio_corregida_ajustada) * 1 Then
   
    txt_resolucion_credito_por_cuota = "OK"
   
   Else
    
       txt_resolucion_credito_por_cuota = "RECHAZADO"
       
   End If

cmd_volver_evaluacion.Enabled = True
cmd_Volver_Ficha.Enabled = True
cmd_guardar_evaluacion.Enabled = True

Else
   MsgBox ("Antes de Calcular Debes Ingresar Valores En Campos Correspondientes")
   End If
   
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


If txt_cantidad_producto = 1 And txt_precio_venta1 >= 1 And txt_precio_venta1 <> "" _
And Val(txt_precio_venta1) > Val(txt_materia_prima1) And txt_incidencia_ventas1 <> "" Then

txt_r_cvcmo1 = Round(((Val(txt_materia_prima1) + Val(txt_mano_obra1)) / Val(txt_precio_venta1)), 3) * 1
txt_r_cvsmo1 = Round((Val(txt_materia_prima1) / Val(txt_precio_venta1)), 3) * 1
txt_r_cvppcmo1 = Round(txt_r_cvcmo1 * txt_incidencia_ventas1 * 0.01, 3) * 1
txt_r_cvppsmo1 = Round(txt_r_cvsmo1 * txt_incidencia_ventas1 * 0.01, 3) * 1


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


If Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5) >= 65 And Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5) <= 100 Then
    'Prender siguiente Boton Calculo
    cmd_calcular_vta_total_mes_al_me_ba.Enabled = True
    
'''suma incidencias al estar correctas
txt_total_incidencias = Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas1)

Else
 MsgBox "Las Incidencias Deben estar entre 65 y 100%"
  
  'Prender siguiente Boton Calculo
cmd_calcular_vta_total_mes_al_me_ba.Enabled = False
  
End If


''''''' CONDICIONES PARA RESULTADO DE ESTADO COSTO VARIABLE PONDERADO
   
   If txt_Sub_Total_x1 * 1 <= 0.2 Then
        
        txt_r_promedio_ponderado = "ZG"
        
        ElseIf txt_Sub_Total_x1 * 1 > 0.2 Then
        txt_r_promedio_ponderado = "A"
    
  End If


End Sub

Private Sub cmd_calcular_vta_total_mes_al_me_ba_Click()

'If txt_065 <> "" And txt_065 > 0 Then

txt_compra_total_mensual = txt_compra_promedio_mensual * txt_veces_compra_mes
txt_venta_total_alto = Int(txt_compra_total_mensual / txt_Sub_Total_x1) 'Primer cambio EBARRIA
txt_venta_total_medio = Int(txt_075 * txt_venta_total_alto)
txt_venta_total_bajo = Int(txt_060 * txt_venta_total_alto)

txt_compra_total_max_corregida = Int(txt_compra_total_mensual / ((Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5)) / 100)) 'Segundo Cambio cambio EBARRIA
txt_venta_total_mes_alto_corregida = Int(txt_venta_total_alto / ((Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5)) / 100))
txt_venta_total_mes_medio_corregida = Int(txt_venta_total_medio / ((Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5)) / 100))
txt_venta_total_mes_bajo_corregida = Int(txt_venta_total_bajo / ((Val(txt_incidencia_ventas1) + Val(txt_incidencia_ventas2) + Val(txt_incidencia_ventas3) + Val(txt_incidencia_ventas4) + Val(txt_incidencia_ventas5)) / 100))

'Prender siguiente Boton Calculo
cmd_calcula_costos_fijos.Enabled = True

txt_tipo_cliente_form_evaluacion = txt_tipo_cliente_evaluacion
txt_tipo_riesgo_form_evaluacion = R_Final_Perfil_evaluacion

'Else
'   MsgBox ("Ingresar Incidencia Sobre Compra Total")
'   End If
End Sub


Private Sub cmd_calcular_flujo_Caja_Click()

txt_r_capacidad_pago = Empty
txt_r_leverage = Empty
txt_r_mto_maximo_aut = Empty
txt_r_venta_total_min = Empty
txt_r_venta_total_max = Empty


'Incializa CAMPOS CALCULADOS
txt_vta_formal_promedio_mes_alto = 0
txt_vta_informal_promedio_mes_alto = 0
txt_Venta_Total_Promedio_Mes_Alto = 0
txt_costo_variable_mes_alto = 0
txt_costo_fijo_mes_alto = 0
txt_resultado_operacional_mes_alto = 0
txt_otros_ingresos_mes_alto = 0
txt_Deudas_flujo_caja_mes_alto = 0
txt_gastos_familiares_mes_alto = 0
txt_capacidad_pago_mes_alto = 0
txt_capacidad_pago_corregida_ajustada_mes_alto = 0

txt_vta_formal_promedio_mes_medio = 0
txt_vta_informal_promedio_mes_medio = 0
txt_Venta_Total_Promedio_Mes_Medio = 0
txt_costo_variable_mes_medio = 0
txt_costo_fijo_mes_medio = 0
txt_resultado_operacional_mes_medio = 0
txt_otros_ingresos_mes_medio = 0
txt_Deudas_flujo_caja_mes_medio = 0
txt_gastos_familiares_mes_medio = 0
txt_capacidad_pago_mes_medio = 0
txt_capacidad_pago_corregida_ajustada_mes_medio = 0

txt_vta_formal_promedio_mes_bajo = 0
txt_vta_informal_promedio_mes_bajo = 0
txt_Venta_Total_Promedio_Mes_Bajo = 0
txt_costo_variable_mes_bajo = 0
txt_costo_fijo_mes_bajo = 0
txt_resultado_operacional_mes_bajo = 0
txt_otros_ingresos_mes_bajo = 0
txt_Deudas_flujo_caja_mes_bajo = 0
txt_gastos_familiares_mes_bajo = 0
txt_capacidad_pago_mes_bajo = 0
txt_capacidad_pago_corregida_ajustada_mes_bajo = 0

txt_capacidad_pago_promedio_corregida_ajustada = Empty
'txt_cuota_credito = Empty
'txt_mto_bruto_sol_cliente = Empty
txt_resolucion_credito_por_cuota = Empty
txt_aprobacion = Empty
txt_venta_total_promedio_anual = Empty
txt_venta_total = Empty
txt_venta_formal_maxima = Empty
txt_r_venta_total_min = Empty


'''''condiciones de calculo para entrar a SUBRUTINA


If numero_meses_tipo_mes_alto * 1 <> 0 Then
   
txt_vta_formal_promedio_mes_alto = 0
txt_vta_informal_promedio_mes_alto = 0
txt_Venta_Total_Promedio_Mes_Alto = 0
txt_costo_variable_mes_alto = 0
txt_costo_fijo_mes_alto = 0
txt_resultado_operacional_mes_alto = 0
txt_otros_ingresos_mes_alto = 0
txt_Deudas_flujo_caja_mes_alto = 0
txt_gastos_familiares_mes_alto = 0
txt_capacidad_pago_mes_alto = 0
txt_capacidad_pago_corregida_ajustada_mes_alto = 0
   
Call CALCULO_FLUJO_CAJA_ALTO

End If

'''''''''''''''

If numero_meses_tipo_mes_medio * 1 <> 0 Then
txt_vta_formal_promedio_mes_medio = 0
txt_vta_informal_promedio_mes_medio = 0
txt_Venta_Total_Promedio_Mes_Medio = 0
txt_costo_variable_mes_medio = 0
txt_costo_fijo_mes_medio = 0
txt_resultado_operacional_mes_medio = 0
txt_otros_ingresos_mes_medio = 0
txt_Deudas_flujo_caja_mes_medio = 0
txt_gastos_familiares_mes_medio = 0
txt_capacidad_pago_mes_medio = 0
txt_capacidad_pago_corregida_ajustada_mes_medio = 0

Call CALCULO_FLUJO_CAJA_MEDIO

End If

'''''''''''''

If numero_meses_tipo_mes_bajo * 1 <> 0 Then

txt_vta_formal_promedio_mes_bajo = 0
txt_vta_informal_promedio_mes_bajo = 0
txt_Venta_Total_Promedio_Mes_Bajo = 0
txt_costo_variable_mes_bajo = 0
txt_costo_fijo_mes_bajo = 0
txt_resultado_operacional_mes_bajo = 0
txt_otros_ingresos_mes_bajo = 0
txt_Deudas_flujo_caja_mes_bajo = 0
txt_gastos_familiares_mes_bajo = 0
txt_capacidad_pago_mes_bajo = 0
txt_capacidad_pago_corregida_ajustada_mes_bajo = 0


Call CALCULO_FLUJO_CAJA_BAJO

End If

'''nuevo 23-03-2012
'**********************
txt_capacidad_pago_promedio_corregida_ajustada = Int(Val((txt_capacidad_pago_corregida_ajustada_mes_alto * numero_meses_tipo_mes_alto) + Val(txt_capacidad_pago_corregida_ajustada_mes_medio * numero_meses_tipo_mes_medio) + Val(txt_capacidad_pago_corregida_ajustada_mes_bajo * numero_meses_tipo_mes_bajo)) / 12)
txt_monto_maximo_credito = Int(txt_capacidad_pago_promedio_corregida_ajustada * txt_leverage)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub


Private Sub cmd_credito_consumo_Click()

'paso de parametros a negociador
'cuota y mto_credito
Credito_Consumo.txt_cuota_comercial = txt_cuota_credito
Credito_Consumo.txt_monto_comercial = txt_mto_bruto_sol_cliente
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Metodologia_Activo_Circulante.Hide
Credito_Consumo.Show

End Sub

Private Sub cmd_guardar_evaluacion_Click()

txt_metodologia_utilizada = "Activo Circulante"

Credito_Consumo.txt_metologia_negociador = txt_metodologia_utilizada

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''SE CHEQUEA SI USUARIO CUMPLE CON POLITICA PARA CREDITO DE CONSUMO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Ficha_Cliente_Micro.cbx_pregunta_consumo = "Si" Then

 If Evaluacion_Perfil.txt_tipo_cliente = "Antiguo Prime" And Evaluacion_Perfil.txt_predictor_Score > 446 Then
   
        Credito_Consumo.txt_monto_comercial = 0
        Credito_Consumo.txt_cuota_comercial = 0
        Credito_Consumo.txt_monto_consumo = 0
        Credito_Consumo.txt_cuota_consumo = 0

        Credito_Consumo.txt_cuota_limite_cliente = txt_capacidad_pago_promedio_corregida_ajustada
        Credito_Consumo.txt_monto_limite_cliente = txt_monto_maximo_credito
        Credito_Consumo.txt_plazo_consumo = Ficha_Cliente_Micro.txt_plazo_credito_consumo
        Credito_Consumo.txt_plazo_comercial = Ficha_Cliente_Micro.txt_plazo_credito
        Credito_Consumo.txt_rut_cliente_negociador = Ficha_Cliente_Micro.txt_rut_cliente
        Estado_Resolucion_Final.txt_r_f_factibilidad_consumo = "A"

        'Credito_Consumo.Show
        MsgBox "Cliente CUMPLE con condiciones para oferta de credito de consumo"
        cmd_credito_consumo.Enabled = True

    
    ElseIf Evaluacion_Perfil.txt_tipo_cliente = "Antiguo No Prime" And Evaluacion_Perfil.txt_predictor_Score > 615 Then
     
       
        Credito_Consumo.txt_monto_comercial = 0
        Credito_Consumo.txt_cuota_comercial = 0
        Credito_Consumo.txt_monto_consumo = 0
        Credito_Consumo.txt_cuota_consumo = 0

        Credito_Consumo.txt_cuota_limite_cliente = txt_capacidad_pago_promedio_corregida_ajustada
        Credito_Consumo.txt_monto_limite_cliente = txt_monto_maximo_credito
        Credito_Consumo.txt_plazo_consumo = Ficha_Cliente_Micro.txt_plazo_credito_consumo
        Credito_Consumo.txt_plazo_comercial = Ficha_Cliente_Micro.txt_plazo_credito
        Credito_Consumo.txt_rut_cliente_negociador = Ficha_Cliente_Micro.txt_rut_cliente
        Estado_Resolucion_Final.txt_r_f_factibilidad_consumo = "A"

        'Credito_Consumo.Show
        MsgBox "Cliente CUMPLE con condiciones para oferta de credito de consumo"
        cmd_credito_consumo.Enabled = True

        ElseIf Evaluacion_Perfil.txt_tipo_cliente = "Nuevo Con Historia Sbif" And (Evaluacion_Perfil.txt_predictor_Score > 622 _
            And Evaluacion_Perfil.cbx_actividad_economica_formal <> "ARTESANO" And Evaluacion_Perfil.cbx_actividad_economica_informal_oficio <> "COMIDA RAPIDA" And _
            Evaluacion_Perfil.cbx_actividad_economica_formal_servicio <> "MODISTAS" And Evaluacion_Perfil.cbx_actividad_economica_semiformal <> "FERIAS LIBRES") And Ficha_Cliente_Micro.cbx_pregunta_comercial = "Si" Then
       
    
        Credito_Consumo.txt_monto_comercial = 0
        Credito_Consumo.txt_cuota_comercial = 0
        Credito_Consumo.txt_monto_consumo = 0
        Credito_Consumo.txt_cuota_consumo = 0

        Credito_Consumo.txt_cuota_limite_cliente = txt_capacidad_pago_promedio_corregida_ajustada
        Credito_Consumo.txt_monto_limite_cliente = txt_monto_maximo_credito
        Credito_Consumo.txt_plazo_consumo = Ficha_Cliente_Micro.txt_plazo_credito_consumo
        Credito_Consumo.txt_plazo_comercial = Ficha_Cliente_Micro.txt_plazo_credito
        Credito_Consumo.txt_rut_cliente_negociador = Ficha_Cliente_Micro.txt_rut_cliente
        Estado_Resolucion_Final.txt_r_f_factibilidad_consumo = "A"

        'Credito_Consumo.Show
        MsgBox "Cliente CUMPLE con condiciones para oferta de credito de consumo"
        cmd_credito_consumo.Enabled = True
    

       Else
            MsgBox "Cliente NO cumple con politica para otorgar credito de consumo", vbCritical
            Estado_Resolucion_Final.txt_r_f_factibilidad_consumo = "R"
    
    End If

End If


'''''''''''''''''''FIN DE CHEQUEO ''''''''''''''''''''''''



cmd_resumen_Estado_Rechazo.Enabled = False


If txt_mto_bruto_sol_cliente >= 0 And txt_mto_bruto_sol_cliente <> "" And txt_cuota_credito >= 0 And txt_cuota_credito <> "" Then

Dim fec1
Dim hora1

fec1 = Format(Date, "yyyy/mm/dd")
txt_fecha_actual = fec1

hora1 = hora
txt_hora_actual = Time


' La conexión a la base de datos
        
    Call conectarBD


   irespuesta = MsgBox("¿Esta Seguro Que Desea Guardar La Evaluacion Final?", vbYesNo)
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

'----------------------------------------------------------
    
  
    ssql = "INSERT INTO TBL_MICRO_PERFIL_RIESGO_CLIENTE " _
    & "([Rut_Cliente], [n_solicitud],[Dv],[Cliente_Nuevo], [Bancarizado],[Antiguedad_Banco],[mora_promedio_dias_BD],[mora_maxima_dias_BD]," _
    & " [R_Tipo_Cliente],[Registro_Ventas],[R_FINAL_PERFIL],[fecha_ingreso],[hora_ingreso],[Metodologia_asignada])" _
    & " VALUES (('" & rut_cliente_ficha & "'), ('" & txt_n_solicitud & "')" _
    & ",('" & dv_cliente_ficha & "') , ('" & Cliente_Nuevo_evaluacion & "'), ('" & Bancarizado_evaluacion & "'),('" & Antiguedad_banco_evaluacion & "')" _
    & ",('" & mora_promedio_dias_BD_evaluacion & "'), ('" & mora_maxima_dias_BD_evaluacion & "'),('" & txt_tipo_cliente_evaluacion & "')" _
    & ",('" & registros_ventas_evaluacion & "'), ('" & R_Final_Perfil_evaluacion & "'),('" & txt_fecha_actual & "'), ('" & txt_hora_actual & "'),('" & metodologia_asignada & "'))"
    
    cnn.Execute ssql
    
    
    ssql = "INSERT INTO TBL_MICRO_METODOLOGIA_ACTIVO_CIRCULANTE " _
& " ([RUT_CLIENTE], [N_SOLICITUD],[DV]," _
& " [compra_promedio_mensual],[veces_compra_mes],[compra_total_mensual_ctm],[caja_banco],[materia_prima],[mercaderias],[cuenta_cobrar],[otros_activos_circulantes],[total_activo_circulante],[producto1],[producto2],[producto3],[producto4],[producto5],[precio_venta1],[precio_venta2],[precio_venta3],[precio_venta4],[precio_venta5],[materia_prima1],[materia_prima2],[materia_prima3],[materia_prima4],[materia_prima5],[mano_obra1],[mano_obra2],[mano_obra3],[mano_obra4],[mano_obra5],[incidencia_ventas1],[incidencia_ventas2],[incidencia_ventas3],[incidencia_ventas4],[incidencia_ventas5],[r_cvcmo1],[r_cvcmo2],[r_cvcmo3],[r_cvcmo4],[r_cvcmo5],[r_cvsmo1],[r_cvsmo2],[r_cvsmo3],[r_cvsmo4],[r_cvsmo5],[r_cvppcmo1],[r_cvppcmo2],[r_cvppcmo3],[r_cvppcmo4],[r_cvppcmo5],[r_cvppsmo1],[r_cvppsmo2],[r_cvppsmo3],[r_cvppsmo4],[r_cvppsmo5],[r_Subtotal_costo_variable],[r_Subtotal_x1_costo_variable],[r_compra_total_mensual]," _
& " [r_venta_total_alto],[r_venta_total_medio],[r_venta_total_bajo],[r_compra_total_max_corregida],[r_venta_total_alto_corregida],[r_venta_total_medio_corregida],[r_venta_total_bajo_corregida],[arriendo_micro],[sueldos],[movilizacion],[servicios_basicos],[contador],[lubricantes],[neumaticos],[afinamientos],[patentes_seguros],[otros_costos_fijos],[total_costos_fijos],[valor_uf],[n_grupo_familiar],[arriendo_vivienda_Gastos_Fam],[gastos_indicado_cliente],[total_gasto_familiar],[liquidacion_sueldo],[jubilacion],[montepio],[arriendo_vivienda_Otro_Ing],[ingreso_segunda_microempresa],[boleta_honorario],[total_otros_ingresos]," _
& " [acreedor1_deuda],[acreedor2_deuda],[acreedor3_deuda],[acreedor4_deuda],[acreedor5_deuda],[acreedor6_deuda],[tipo_producto1_deuda],[tipo_producto2_deuda],[tipo_producto3_deuda],[tipo_producto4_deuda],[tipo_producto5_deuda],[tipo_producto6_deuda],[saldo_pendiente1_deuda],[saldo_pendiente2_deuda],[saldo_pendiente3_deuda],[saldo_pendiente4_deuda],[saldo_pendiente5_deuda],[saldo_pendiente6_deuda],[monto_cuota1_deuda]," _
& " [monto_cuota2_deuda],[monto_cuota3_deuda],[monto_cuota4_deuda],[monto_cuota5_deuda],[monto_cuota6_deuda],[cuotas_pactadas1_deuda],[cuotas_pactadas2_deuda],[cuotas_pactadas3_deuda],[cuotas_pactadas4_deuda],[cuotas_pactadas5_deuda],[cuotas_pactadas6_deuda],[cuotas_pendientes1_deuda],[cuotas_pendientes2_deuda],[cuotas_pendientes3_deuda],[cuotas_pendientes4_deuda],[cuotas_pendientes5_deuda],[cuotas_pendientes6_deuda],[prepaga_cuota1_deuda],[prepaga_cuota2_deuda],[prepaga_cuota3_deuda],[prepaga_cuota4_deuda],[prepaga_cuota5_deuda],[prepaga_cuota6_deuda],[total_saldo_pendiente_deuda],[total_deudas],[numero_meses_alto_flujo],[numero_meses_medio_flujo],[numero_meses_bajo_flujo],[vta_formal_promedio_mes_alto_flujo],[vta_formal_promedio_mes_medio_flujo],[vta_formal_promedio_mes_bajo_flujo],[vta_informal_promedio_mes_alto_flujo],[vta_informal_promedio_mes_medio_flujo],[vta_informal_promedio_mes_bajo_flujo],[Venta_Total_Promedio_Mes_Alto_flujo]," _
& " [Venta_Total_Promedio_Mes_medio_flujo],[Venta_Total_Promedio_Mes_bajo_flujo],[resultado_operacional_alto_flujo],[resultado_operacional_medio_flujo],[resultado_operacional_bajo_flujo],[capacidad_pago_mes_alto_flujo],[capacidad_pago_mes_medio_flujo],[capacidad_pago_mes_bajo_flujo],[cap_pago_corregida_ajus_mes_alto_flujo],[cap_pago_corregida_ajus_mes_medio_flujo],[cap_pago_corregida_ajus_mes_bajo_flujo],[cap_pago_promedio_corregida_ajustada_flujo],[monto_maximo_credito_flujo],[cuota_credito_flujo],[mto_bruto_solicitado_cliente_flujo],[resolucion_credito_cuota_flujo],[resolucion_credito_monto_flujo],[venta_total_promedio_anual],[fecha_ingreso],[hora_ingreso],[impuesto],[venta_formal_maxima],[leverage],[tipo_credito_deuda1],[tipo_credito_deuda2],[tipo_credito_deuda3],[tipo_credito_deuda4],[tipo_credito_deuda5],[tipo_credito_deuda6],[total_saldo_pendiente_consumo],[total_deudas_consumo],[total_saldo_pendiente_comercial],[total_deudas_comercial],[saldo_deuda_con_prepago_consumo]," _
& " [saldo_deuda_con_prepago_comercial],[mto_cuota_con_prepago_consumo],[mto_cuota_con_prepago_comercial],[saldo_deuda_sin_prepago_consumo],[saldo_deuda_sin_prepago_comercial],[mto_cuota_sin_prepago_comercial],[mto_cuota_sin_prepago_consumo])" _
& " VALUES (('" & txt_rut_cliente & "'),('" & txt_n_solicitud & "'),('" & txt_dv & "'),('" & txt_compra_promedio_mensual & "'),('" & txt_veces_compra_mes & "'),('" & txt_compra_total_mensual_ctm & "'),('" & txt_caja_banco & "'),('" & txt_materia_primas & "'),('" & txt_mercaderias & "'),('" & txt_cuenta_cobrar & "'),('" & txt_otros_activos_circulantes & "'),('" & txt_total_activos_circulantes & "'),('" & txt_producto1 & "'),('" & txt_producto2 & "'),('" & txt_producto3 & "'),('" & txt_producto4 & "'),('" & txt_producto5 & "'),('" & txt_precio_venta1 & "'),('" & txt_precio_venta2 & "'),('" & txt_precio_venta3 & "'),('" & txt_precio_venta4 & "'),('" & txt_precio_venta5 & "'),('" & txt_materia_prima1 & "'),('" & txt_materia_prima2 & "'),('" & txt_materia_prima3 & "'),('" & txt_materia_prima4 & "'),('" & txt_materia_prima5 & "'),('" & txt_mano_obra1 & "'),('" & txt_mano_obra2 & "'),('" & txt_mano_obra3 & "')" _
& ",('" & txt_mano_obra4 & "'),('" & txt_mano_obra5 & "'),('" & txt_incidencia_ventas1 & "'),('" & txt_incidencia_ventas2 & "'),('" & txt_incidencia_ventas3 & "'),('" & txt_incidencia_ventas4 & "'),('" & txt_incidencia_ventas5 & "'),('" & txt_r_cvcmo1 & "'),('" & txt_r_cvcmo2 & "'),('" & txt_r_cvcmo3 & "'),('" & txt_r_cvcmo4 & "'),('" & txt_r_cvcmo5 & "'),('" & txt_r_cvsmo1 & "'),('" & txt_r_cvsmo2 & "'),('" & txt_r_cvsmo3 & "'),('" & txt_r_cvsmo4 & "'),('" & txt_r_cvsmo5 & "'),('" & txt_r_cvppcmo1 & "'),('" & txt_r_cvppcmo2 & "'),('" & txt_r_cvppcmo3 & "'),('" & txt_r_cvppcmo4 & "'),('" & txt_r_cvppcmo5 & "'),('" & txt_r_cvppsmo1 & "'),('" & txt_r_cvppsmo2 & "'),('" & txt_r_cvppsmo3 & "'),('" & txt_r_cvppsmo4 & "'),('" & txt_r_cvppsmo5 & "'),('" & txt_Sub_Total & "'),('" & txt_Sub_Total_x1 & "'),('" & txt_compra_total_mensual & "'),('" & txt_venta_total_alto & "'),('" & txt_venta_total_medio & "'),('" & txt_venta_total_bajo & "') " _
& ",('" & txt_compra_total_max_corregida & "'),('" & txt_venta_total_mes_alto_corregida & "'),('" & txt_venta_total_mes_medio_corregida & "'),('" & txt_venta_total_mes_bajo_corregida & "'),('" & txt_arriendo_micro & "'),('" & txt_sueldos & "'),('" & txt_movilizacion & "'),('" & txt_servicios_basicos & "'),('" & txt_contador & "'),('" & txt_lubricantes & "'),('" & txt_neumaticos & "'),('" & txt_afinamientos & "'),('" & txt_patentes_seguros & "'),('" & txt_otros_costos_fijos & "'),('" & txt_total_costos_fijos & "'),('" & txt_valor_uf & "'),('" & txt_n_grupo_familiar & "'),('" & txt_arriendo_vivienda & "'),('" & txt_gastos_indicado_cliente & "'),('" & txt_total_gasto_familiar & "'),('" & txt_liquidacion_sueldo & "'),('" & txt_jubilacion & "'),('" & txt_montepio & "'),('" & txt_arriendo_vivienda1 & "'),('" & txt_ingreso_segunda_microempresa & "'),('" & txt_boleta_honorario & "'),('" & txt_total_otros_ingresos & "'),('" & txt_acreedor1 & "'),('" & txt_acreedor2 & "'),('" & txt_acreedor3 & "')" _
& ",('" & txt_acreedor4 & "'),('" & txt_acreedor5 & "'),('" & txt_acreedor6 & "'),('" & txt_tipo_producto1 & "'),('" & txt_tipo_producto2 & "'),('" & txt_tipo_producto3 & "'),('" & txt_tipo_producto4 & "'),('" & txt_tipo_producto5 & "'),('" & txt_tipo_producto6 & "'),('" & txt_saldo_pendiente1 & "'),('" & txt_saldo_pendiente2 & "'),('" & txt_saldo_pendiente3 & "'),('" & txt_saldo_pendiente4 & "'),('" & txt_saldo_pendiente5 & "'),('" & txt_saldo_pendiente6 & "'),('" & txt_monto_cuota1 & "'),('" & txt_monto_cuota2 & "'),('" & txt_monto_cuota3 & "'),('" & txt_monto_cuota4 & "'),('" & txt_monto_cuota5 & "'),('" & txt_monto_cuota6 & "'),('" & txt_cuotas_pactadas1 & "'),('" & txt_cuotas_pactadas2 & "'),('" & txt_cuotas_pactadas3 & "'),('" & txt_cuotas_pactadas4 & "'),('" & txt_cuotas_pactadas5 & "'),('" & txt_cuotas_pactadas6 & "'),('" & txt_cuotas_pendientes1 & "'),('" & txt_cuotas_pendientes2 & "'),('" & txt_cuotas_pendientes3 & "')" _
& ",('" & txt_cuotas_pendientes4 & "'),('" & txt_cuotas_pendientes5 & "'),('" & txt_cuotas_pendientes6 & "'),('" & cbx_prepaga_deuda1 & "'),('" & cbx_prepaga_deuda2 & "'),('" & cbx_prepaga_deuda3 & "'),('" & cbx_prepaga_deuda4 & "'),('" & cbx_prepaga_deuda5 & "'),('" & cbx_prepaga_deuda6 & "'),('" & txt_total_saldo_pendiente & "'),('" & txt_total_deudas & "'),('" & numero_meses_tipo_mes_alto & "'),('" & numero_meses_tipo_mes_medio & "'),('" & numero_meses_tipo_mes_bajo & "'),('" & txt_vta_formal_promedio_mes_alto & "'),('" & txt_vta_formal_promedio_mes_medio & "'),('" & txt_vta_formal_promedio_mes_bajo & "'),('" & txt_vta_informal_promedio_mes_alto & "'),('" & txt_vta_informal_promedio_mes_medio & "'),('" & txt_vta_informal_promedio_mes_bajo & "'),('" & txt_Venta_Total_Promedio_Mes_Alto & "'),('" & txt_Venta_Total_Promedio_Mes_Medio & "'),('" & txt_Venta_Total_Promedio_Mes_Bajo & "'),('" & txt_resultado_operacional_mes_alto & "')" _
& ",('" & txt_resultado_operacional_mes_medio & "'),('" & txt_resultado_operacional_mes_bajo & "'),('" & txt_capacidad_pago_mes_alto & "'),('" & txt_capacidad_pago_mes_medio & "'),('" & txt_capacidad_pago_mes_bajo & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_alto & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_medio & "'),('" & txt_capacidad_pago_corregida_ajustada_mes_bajo & "'),('" & txt_capacidad_pago_promedio_corregida_ajustada & "'),('" & txt_monto_maximo_credito & "'),('" & txt_cuota_credito & "'),('" & txt_mto_bruto_sol_cliente & "'),('" & txt_resolucion_credito_por_cuota & "'),('" & txt_aprobacion & "'), ('" & txt_venta_total_promedio_anual & "'),('" & txt_fecha_actual & "'), ('" & txt_hora_actual & "'), ('" & txt_impuesto & "'),('" & txt_venta_formal_maxima & "'),('" & txt_leverage & "'),('" & cbx_tipo_credito_deuda1 & "'),('" & cbx_tipo_credito_deuda2 & "'),('" & cbx_tipo_credito_deuda3 & "'),('" & cbx_tipo_credito_deuda4 & "'),('" & cbx_tipo_credito_deuda5 & "')" _
& ",('" & cbx_tipo_credito_deuda6 & "'),('" & txt_total_saldo_pendiente_consumo & "'),('" & txt_total_deudas_consumo & "'),('" & txt_total_saldo_pendiente_comercial & "'),('" & txt_total_deudas_comercial & "'),('" & txt_saldo_deuda_con_prepago_consumo & "'),('" & txt_saldo_deuda_con_prepago_comercial & "'),('" & txt_mto_cuota_con_prepago_consumo & "'),('" & txt_mto_cuota_con_prepago_comercial & "'),('" & txt_saldo_deuda_sin_prepago_consumo & "'),('" & txt_saldo_deuda_sin_prepago_comercial & "'),('" & txt_mto_cuota_sin_prepago_comercial & "'),('" & txt_mto_cuota_sin_prepago_consumo & "'))"

    cnn.Execute ssql
    
    
    cmd_resumen_Estado_Rechazo.Enabled = True
    cmd_guardar_evaluacion.Enabled = False
    
 End If
Else
    MsgBox "Debe Ingresar Cuota Comercial y Monto Solicitado Por El Cliente"
 End If


'''se inhibe campos de cuota y mto comercial hasta que no se cree un nuevo numero de solicitud en la ficha

txt_cuota_credito.Locked = True
txt_mto_bruto_sol_cliente.Locked = True


End Sub

Private Sub cmd_imprimir1_meto_ac_Click()
Metodologia_Activo_Circulante.PrintForm
End Sub

Private Sub cmd_imprimir2_meto_ac_Click()
Metodologia_Activo_Circulante.PrintForm
End Sub

Private Sub cmd_imprimir3_meto_ac_Click()
Metodologia_Activo_Circulante.PrintForm
End Sub

Private Sub cmd_imprimir4_meto_ac_Click()
Metodologia_Activo_Circulante.PrintForm
End Sub

Private Sub cmd_salir_del_sistema_Click()
    ActiveWorkbook.Save
    Workbooks("Microempresas_1401.xls").Close
    Application.Quit
End Sub

Private Sub cmd_resumen_Estado_Rechazo_Click()

Estado_Resolucion_Final.cmd_guardar_evaluacion.Enabled = False
'Estado_Resolucion_Final.cmd_volver_pag_anterior.Enabled = False
Estado_Resolucion_Final.cmd_carta_rechazo.Enabled = False
Estado_Resolucion_Final.Imprimir_resolucion_f.Enabled = False
Estado_Resolucion_Final.cmd_volver_evaluacion.Enabled = False


Call conectarBD

Metodologia_Activo_Circulante.Hide

Estado_Resolucion_Final.txt_resultado_APROBADO_final_cred.Enabled = False
Estado_Resolucion_Final.txt_resultado_RECHAZADO_final_cred.Enabled = False


Estado_Resolucion_Final.TXT_ESTADO_METODOLOGIA_OCUPADA = "Activo Circulante"


''''  TRAE LA ULTIMA SOLICITUD rut evaluando

        ssql = "select top 1 n_solicitud, fecha_ingreso, hora_ingreso" _
        & " from TBL_MICRO_ficha_cliente" _
        & " where rut_cliente = '" & txt_rut_cliente & "'" _
        & " order by n_solicitud desc" _
        
        Set rst = cnn.Execute(ssql, , adCmdText)
        
        Estado_Resolucion_Final.txt_n_solicitud = rst!n_solicitud
        Estado_Resolucion_Final.txt_fecha_actual = rst!FECHA_INGRESO
        Estado_Resolucion_Final.txt_hora_actual = rst!HORA_INGRESO
        

'''' RUT CLIENTE - NUMERO SOLICITUD - FECHA ACTUAL

    Estado_Resolucion_Final.txt_rut_cliente = Ficha_Cliente_Micro.txt_rut_cliente
    Estado_Resolucion_Final.txt_dv = Ficha_Cliente_Micro.txt_dv
    Estado_Resolucion_Final.txt_fecha_actual = Ficha_Cliente_Micro.txt_fecha_ingreso_compara
    Estado_Resolucion_Final.txt_hora_actual = Ficha_Cliente_Micro.txt_hora_actual
    

''''MORAS SBIF
    Estado_Resolucion_Final.txt_r_f_mora_directa_SBIF = Ficha_Cliente_Micro.txt_r_mora_sbif
    Estado_Resolucion_Final.txt_r_f_vdo_directo_SBIF = Ficha_Cliente_Micro.cbx_r_venc_cast_SBIF
    Estado_Resolucion_Final.txt_r_f_cast_directo_SBIF = Ficha_Cliente_Micro.cbx_r_Mora_Total_Sbif
    Estado_Resolucion_Final.txt_r_f_vdo_indirecto_SBIF = Ficha_Cliente_Micro.cbx_r_venc_cast_SBIF_indirecta
    Estado_Resolucion_Final.txt_r_f_cast_indirecto_SBIF = Ficha_Cliente_Micro.cbx_r_Mora_Total_Sbif_indirecta

'MORAS INTERNAS

    Estado_Resolucion_Final.txt_r_f_mora_directa = Ficha_Cliente_Micro.txt_r_mora_directa_interna
    Estado_Resolucion_Final.txt_r_f_Vencido_directo = Ficha_Cliente_Micro.txt_r_Vencido_directo_interna
    Estado_Resolucion_Final.txt_r_f_castigo_directo = Ficha_Cliente_Micro.txt_r_castigo_directo_interna

'solo ACTIVO circylante
    Estado_Resolucion_Final.txt_r_f_compra_tot_mensual = Metodologia_Activo_Circulante.txt_r_compra_activo_AC
'------------------

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
    'Estado_Resolucion_Final.txt_r_f_edad = Ficha_Cliente_Micro.txt_r_edad
    Estado_Resolucion_Final.txt_r_f_antiguedad_giro = Ficha_Cliente_Micro.txt_r_meses_antiguedad
    
    
    '##########################################################################################
    'Cambiado la fila para que tome el R o A de el RI de Riesgo
    
    'Estado_Resolucion_Final.txt_r_f_ir_sinac = Ficha_Cliente_Micro.txt_r_predictor_score_dicom
    Estado_Resolucion_Final.txt_r_f_ir_sinac = Evaluacion_Perfil.txt_r_dicom_tipo_cliente
    Estado_Resolucion_Final.txt_r_f_ir_tipo_cliente = Evaluacion_Perfil.txt_r_dicom_tipo_cliente
    '##########################################################################################
    
    Estado_Resolucion_Final.txt_r_f_deuda_sbif_declarada = Metodologia_Activo_Circulante.txt_r_sbif_declarada
    Estado_Resolucion_Final.txt_r_f_nivel_vta_inf_min = Metodologia_Activo_Circulante.txt_r_venta_total_min
    Estado_Resolucion_Final.txt_r_f_nivel_vta_sup_max = Metodologia_Activo_Circulante.txt_r_venta_total_max
    Estado_Resolucion_Final.txt_r_f_capacidad_pago = Metodologia_Activo_Circulante.txt_r_capacidad_pago
    'Estado_Resolucion_Final.txt_r_f_costo_fijo_rub_trasp = Metodologia_Activo_Circulante.txt_valida_costos_fijos
    Estado_Resolucion_Final.txt_r_f_costo_variable_ponde = Metodologia_Activo_Circulante.txt_r_promedio_ponderado
    'Estado_Resolucion_Final.txt_r_f_compra_tot_mensual = Metodologia_Activo_Circulante.txt_r_nota_superar60
    Estado_Resolucion_Final.txt_r_f_leverage = Metodologia_Activo_Circulante.txt_r_leverage
    Estado_Resolucion_Final.txt_r_f_costo_variable_ponde = Metodologia_Activo_Circulante.txt_r_promedio_ponderado

    
    'Recodifica las edades minima y maxima para un cliente empresa
    
    If Ficha_Cliente_Micro.txt_rut_cliente >= 45000000 Then
       
       Estado_Resolucion_Final.txt_r_f_edad = "N/A"
       Estado_Resolucion_Final.txt_r_f_edad_maxima = "N/A"
       Estado_Resolucion_Final.txt_r_f_ir_sinac = "ZG"
       Estado_Resolucion_Final.txt_r_f_ir_tipo_cliente = "ZG"
       
    End If
    

Estado_Resolucion_Final.Show

End Sub

Private Sub cmd_volver_evaluacion_Click()
MsgBox "Recuerda Que Al Volver y Cambiar Datos Debes Recalcular Los Campos"

Metodologia_Activo_Circulante.Hide
Evaluacion_Perfil.Show
End Sub

Private Sub cmd_volver_ficha_Click()
MsgBox "Recuerda Que Al Volver y Cambiar Datos Debes Recalcular Los Campos"
Ficha_Cliente_Micro.cmd_Menu_Evaluacion.Enabled = False
Metodologia_Activo_Circulante.cmd_guardar_evaluacion.Enabled = True
Metodologia_Activo_Circulante.Hide
Ficha_Cliente_Micro.Show
End Sub

Private Sub CommandButton1_Click()
Metodologia_Activo_Circulante.Hide
Estado_Resolucion_Final.Show
End Sub

Private Sub Label241_Click()

End Sub

Private Sub Label37_Click()

End Sub

Private Sub numero_meses_tipo_mes_alto_Change()

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

Private Sub txt_aprobacion_Change()

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

Private Sub txt_compra_promedio_mensual_Change()

End Sub

Private Sub txt_credito_hipotecario_Change()
'txt_credito_hipotecario = Format(txt_credito_hipotecario, "##,##")
End Sub
Private Sub txt_cuota_credito_AfterUpdate()
If txt_cuota_credito * 1 > txt_capacidad_pago_promedio_corregida_ajustada * 1 Then
  txt_r_capacidad_pago = "R"
  'lbl_accion.BackColor = &HFF&       'rojo
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else
  
  txt_r_capacidad_pago = "A"
  'lbl_accion.BackColor = &HC000&
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
End If
End Sub



Private Sub txt_cuotas_pactadas1_Change()

End Sub

Private Sub txt_cuotas_pendientes1_Change()

End Sub

Private Sub txt_cupo_linea_credito_Change()

'txt_cupo_linea_credito = Format(txt_cupo_linea_credito, "##,##")
End Sub

Private Sub txt_deuda_comercial_Change()
'txt_deuda_comercial = Format(txt_deuda_comercial, "##,##")
End Sub

Private Sub txt_deuda_consumo_indirecta_Change()
'txt_deuda_consumo_indirecta = Format(txt_deuda_consumo_indirecta, "##,##")
End Sub

Private Sub txt_deuda_d10_comercial_Change()
'txt_deuda_d10_comercial = Format(txt_deuda_d10_comercial, "##,##")
End Sub

Private Sub txt_deuda_d10_consumo_Change()
'txt_deuda_d10_consumo = Format(txt_deuda_d10_consumo, "##,##")
End Sub

Private Sub txt_deuda_d10_hipotecario_Change()
'txt_deuda_d10_hipotecario = Format(txt_deuda_d10_hipotecario, "##,##")
End Sub

Private Sub txt_deuda_d10_linea_Change()
'txt_deuda_d10_linea = Format(txt_deuda_d10_linea, "##,##")
End Sub

Private Sub txt_deuda_indirecta_vigente_Change()
'txt_deuda_indirecta_vigente = Format(txt_deuda_indirecta_vigente, "##,##")
End Sub

Private Sub txt_deuda_vigente_directa_Change()
'txt_deuda_vigente_directa = Format(txt_deuda_vigente_directa, "##,##")
End Sub

Private Sub txt_deudas_directas_vig_Change()
'txt_deudas_directas_vig = Format(txt_deudas_directas_vig, "##,##")
End Sub



Private Sub txt_fecha_actual_Change()

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

cbx_tipo_credito_deuda1.Visible = True

cbx_tipo_credito_deuda2.Visible = False
cbx_tipo_credito_deuda3.Visible = False
cbx_tipo_credito_deuda4.Visible = False
cbx_tipo_credito_deuda5.Visible = False
cbx_tipo_credito_deuda6.Visible = False

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

cbx_tipo_credito_deuda1.Visible = True
cbx_tipo_credito_deuda2.Visible = True

txt_acreedor1.Visible = True
txt_tipo_producto1.Visible = True
txt_monto_cuota1.Visible = True
txt_cuotas_pactadas1.Visible = True
txt_cuotas_pendientes1.Visible = True
txt_saldo_pendiente1.Visible = True
cbx_prepaga_deuda1.Visible = True

cbx_tipo_credito_deuda3.Visible = False
cbx_tipo_credito_deuda4.Visible = False
cbx_tipo_credito_deuda5.Visible = False
cbx_tipo_credito_deuda6.Visible = False


ElseIf txt_ingreso_cantidad_deudas = 3 Then

txt_acreedor3.Visible = True
txt_tipo_producto3.Visible = True
txt_monto_cuota3.Visible = True
txt_cuotas_pactadas3.Visible = True
txt_cuotas_pendientes3.Visible = True
txt_saldo_pendiente3.Visible = True
cbx_prepaga_deuda3.Visible = True

cbx_tipo_credito_deuda1.Visible = True
cbx_tipo_credito_deuda2.Visible = True
cbx_tipo_credito_deuda3.Visible = True


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

cbx_tipo_credito_deuda1.Visible = True
cbx_tipo_credito_deuda2.Visible = True
cbx_tipo_credito_deuda3.Visible = True
cbx_tipo_credito_deuda4.Visible = True


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

cbx_tipo_credito_deuda1.Visible = True
cbx_tipo_credito_deuda2.Visible = True
cbx_tipo_credito_deuda3.Visible = True
cbx_tipo_credito_deuda4.Visible = True
cbx_tipo_credito_deuda5.Visible = True

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

cbx_tipo_credito_deuda1.Visible = True
cbx_tipo_credito_deuda2.Visible = True
cbx_tipo_credito_deuda3.Visible = True
cbx_tipo_credito_deuda4.Visible = True
cbx_tipo_credito_deuda5.Visible = True
cbx_tipo_credito_deuda6.Visible = True

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



Private Sub txt_monto_cuota1_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_monto_cuota2_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub txt_monto_cuota3_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_monto_cuota4_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_monto_cuota5_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub txt_monto_cuota6_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub txt_monto_maximo_credito_Change()

End Sub

Private Sub txt_mto_bruto_sol_cliente_AfterUpdate()
If txt_mto_bruto_sol_cliente * 1 > txt_monto_maximo_credito * 1 Then
    txt_r_leverage = "R"
  'lbl_accion.BackColor = &HFF&       'rojo
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
  
  Else
  
  txt_r_leverage = "A"
  'lbl_accion.BackColor = &HC000&
  'lbl_accion.ForeColor = &H8000000E  'blanco
  'lbl_accion.BorderStyle = fmBorderStyleSingle 'con bordes
End If

If txt_mto_bruto_sol_cliente > Menu_Principal_Micro.txt_monto_aut_micro Then
'Menu_Principal_Micro.txt_monto_aut_micro Then
    
    txt_r_mto_maximo_aut = "ZG"
    Estado_Resolucion_Final.txt_r_f_mto_maximo_aut = "ZG"
    
    Else
    
    txt_r_mto_maximo_aut = "A"
    Estado_Resolucion_Final.txt_r_f_mto_maximo_aut = "A"
    
End If

End Sub

Private Sub txt_n_grupo_familiar_Change()

End Sub

Private Sub txt_precio_venta1_Change()

End Sub

Private Sub txt_producto1_Change()
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

Private Sub txt_r_capacidad_pago_Change()

End Sub

Private Sub txt_r_leverage_Change()

End Sub

Private Sub txt_r_sbif_declarada_Change()

End Sub

Private Sub txt_r_venta_total_max_Change()

End Sub

Private Sub txt_r_venta_total_min_Change()

End Sub

Private Sub txt_resolucion_credito_por_cuota_Change()

End Sub

Private Sub txt_saldo_pendiente1_Change()

txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_saldo_pendiente2_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_saldo_pendiente3_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub txt_saldo_pendiente4_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
End Sub

Private Sub txt_saldo_pendiente5_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0

End Sub

Private Sub txt_saldo_pendiente6_Change()
txt_total_saldo_pendiente = 0
txt_total_deudas = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_comercial = 0
txt_total_deudas_comercial = 0

txt_sumar_mto_cuota1 = 0
txt_no_sumar_mto_cuota1 = 0
txt_sumar_mto_cuota1_consumo = 0
txt_no_sumar_mto_cuota1_consumo = 0
txt_sumar_mto_cuota1_comercial = 0
txt_no_sumar_mto_cuota1_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota2 = 0
txt_no_sumar_mto_cuota2 = 0
txt_sumar_mto_cuota2_consumo = 0
txt_no_sumar_mto_cuota2_consumo = 0
txt_sumar_mto_cuota2_comercial = 0
txt_no_sumar_mto_cuota2_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota3 = 0
txt_no_sumar_mto_cuota3 = 0
txt_sumar_mto_cuota3_consumo = 0
txt_no_sumar_mto_cuota3_consumo = 0
txt_sumar_mto_cuota3_comercial = 0
txt_no_sumar_mto_cuota3_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota4 = 0
txt_no_sumar_mto_cuota4 = 0
txt_sumar_mto_cuota4_consumo = 0
txt_no_sumar_mto_cuota4_consumo = 0
txt_sumar_mto_cuota4_comercial = 0
txt_no_sumar_mto_cuota4_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota5 = 0
txt_no_sumar_mto_cuota5 = 0
txt_sumar_mto_cuota5_consumo = 0
txt_no_sumar_mto_cuota5_consumo = 0
txt_sumar_mto_cuota5_comercial = 0
txt_no_sumar_mto_cuota5_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_cuota6 = 0
txt_no_sumar_mto_cuota6 = 0
txt_sumar_mto_cuota6_consumo = 0
txt_no_sumar_mto_cuota6_consumo = 0
txt_sumar_mto_cuota6_comercial = 0
txt_no_sumar_mto_cuota6_comercial = 0
txt_total_deudas_consumo = 0
txt_total_saldo_pendiente_consumo = 0
txt_total_deudas_consumo_comercial = 0
txt_total_saldo_pendiente_consumo_comercial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_sumar_mto_deuda1_consumo = 0
txt_sumar_mto_deuda2_consumo = 0
txt_sumar_mto_deuda3_consumo = 0
txt_sumar_mto_deuda4_consumo = 0
txt_sumar_mto_deuda5_consumo = 0
txt_sumar_mto_deuda6_consumo = 0

txt_sumar_mto_deuda1_comecial = 0
txt_sumar_mto_deuda2_comecial = 0
txt_sumar_mto_deuda3_comecial = 0
txt_sumar_mto_deuda4_comecial = 0
txt_sumar_mto_deuda5_comecial = 0
txt_sumar_mto_deuda6_comecial = 0

txt_saldo_deuda_con_prepago_consumo = 0
txt_mto_cuota_con_prepago_consumo = 0
txt_saldo_deuda_con_prepago_comercial = 0
txt_mto_cuota_con_prepago_comercial = 0
txt_saldo_deuda_sin_prepago_consumo = 0
txt_saldo_deuda_sin_prepago_comercial = 0
txt_mto_cuota_sin_prepago_consumo = 0
txt_mto_cuota_sin_prepago_comercial = 0


txt_saldo_deuda_con_prepago_consumo1 = 0
txt_mto_cuota_con_prepago_consumo1 = 0
txt_saldo_deuda_con_prepago_comercial1 = 0
txt_mto_cuota_con_prepago_comercial1 = 0
txt_saldo_deuda_sin_prepago_consumo1 = 0
txt_mto_cuota_sin_prepago_consumo1 = 0
txt_saldo_deuda_sin_prepago_comercial1 = 0
txt_mto_cuota_sin_prepago_comercial1 = 0

txt_saldo_deuda_con_prepago_consumo2 = 0
txt_mto_cuota_con_prepago_consumo2 = 0
txt_saldo_deuda_con_prepago_comercial2 = 0
txt_mto_cuota_con_prepago_comercial2 = 0
txt_saldo_deuda_sin_prepago_consumo2 = 0
txt_mto_cuota_sin_prepago_consumo2 = 0
txt_saldo_deuda_sin_prepago_comercial2 = 0
txt_mto_cuota_sin_prepago_comercial2 = 0

txt_saldo_deuda_con_prepago_consumo3 = 0
txt_mto_cuota_con_prepago_consumo3 = 0
txt_saldo_deuda_con_prepago_comercial3 = 0
txt_mto_cuota_con_prepago_comercial3 = 0
txt_saldo_deuda_sin_prepago_consumo3 = 0
txt_mto_cuota_sin_prepago_consumo3 = 0
txt_saldo_deuda_sin_prepago_comercial3 = 0
txt_mto_cuota_sin_prepago_comercial3 = 0

txt_saldo_deuda_con_prepago_consumo4 = 0
txt_mto_cuota_con_prepago_consumo4 = 0
txt_saldo_deuda_con_prepago_comercial4 = 0
txt_mto_cuota_con_prepago_comercial4 = 0
txt_saldo_deuda_sin_prepago_consumo4 = 0
txt_mto_cuota_sin_prepago_consumo4 = 0
txt_saldo_deuda_sin_prepago_comercial4 = 0
txt_mto_cuota_sin_prepago_comercial4 = 0

txt_saldo_deuda_con_prepago_consumo5 = 0
txt_mto_cuota_con_prepago_consumo5 = 0
txt_saldo_deuda_con_prepago_comercial5 = 0
txt_mto_cuota_con_prepago_comercial5 = 0
txt_saldo_deuda_sin_prepago_consumo5 = 0
txt_mto_cuota_sin_prepago_consumo5 = 0
txt_saldo_deuda_sin_prepago_comercial5 = 0
txt_mto_cuota_sin_prepago_comercial5 = 0

txt_saldo_deuda_con_prepago_consumo6 = 0
txt_mto_cuota_con_prepago_consumo6 = 0
txt_saldo_deuda_con_prepago_comercial6 = 0
txt_mto_cuota_con_prepago_comercial6 = 0
txt_saldo_deuda_sin_prepago_consumo6 = 0
txt_mto_cuota_sin_prepago_consumo6 = 0
txt_saldo_deuda_sin_prepago_comercial6 = 0
txt_mto_cuota_sin_prepago_comercial6 = 0
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

Private Sub txt_total_activos_circulantes_Change()

End Sub

Private Sub txt_total_deuda_d10_Change()
'txt_total_deuda_d10 = Format(txt_total_deuda_d10, "##,##")
End Sub

Private Sub txt_total_deudas_sbif_Change()
'txt_total_deudas_sbif = Format(txt_total_deudas_sbif, "##,##")
End Sub

Private Sub txt_total_saldo_pendiente_Change()
'txt_total_saldo_pendiente = Format(txt_total_saldo_pendiente, "##,##")
End Sub

Private Sub txt_valor_uf_Change()

End Sub

Private Sub txt_venta_formal_maxima_Change()

If (txt_vta_formal_promedio_mes_alto * 1 * numero_meses_tipo_mes_alto * 1 + txt_vta_formal_promedio_mes_medio * 1 * numero_meses_tipo_mes_medio * 1 + txt_vta_formal_promedio_mes_bajo * 1 * numero_meses_tipo_mes_bajo * 1) / Metodologia_Activo_Circulante.txt_valor_uf * 1 > 2400 Then
txt_r_venta_total_max = "ZG"
Else
txt_r_venta_total_max = "A"
End If


'(txt_venta_total * 1 / Metodologia_Activo_Circulante.txt_valor_uf * 1) < 120

End Sub

Private Sub txt_venta_total_Change()



'If (txt_venta_total * 1 / Metodologia_Activo_Circulante.txt_valor_uf * 1) > 2400 Then
 '       txt_r_venta_total_max = "R"
'Else
'        txt_r_venta_total_max = "A"
'End If

If txt_venta_total = "" Then
   txt_venta_total = 0
End If

If (txt_venta_total * 1 / Metodologia_Activo_Circulante.txt_valor_uf * 1) < 120 Then

        txt_r_venta_total_min = "R"
Else
        txt_r_venta_total_min = "A"
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
MsgBox ("Boton Deshabilitado Ocupe Opciones De Menu")
Cancel = True
End If
End Sub

Private Sub UserForm_Initialize()

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

cbx_tipo_credito_deuda1.AddItem "Consumo"
cbx_tipo_credito_deuda1.AddItem "Comercial"

cbx_tipo_credito_deuda2.AddItem "Consumo"
cbx_tipo_credito_deuda2.AddItem "Comercial"

cbx_tipo_credito_deuda3.AddItem "Consumo"
cbx_tipo_credito_deuda3.AddItem "Comercial"

cbx_tipo_credito_deuda4.AddItem "Consumo"
cbx_tipo_credito_deuda4.AddItem "Comercial"

cbx_tipo_credito_deuda5.AddItem "Consumo"
cbx_tipo_credito_deuda5.AddItem "Comercial"

cbx_tipo_credito_deuda6.AddItem "Consumo"
cbx_tipo_credito_deuda6.AddItem "Comercial"

End Sub

Public Sub CALCULO_FLUJO_CAJA_ALTO()





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 <> 12 Then
  
MsgBox "La Suma de Meses Debe Ser igual a 12 Revise ..."

Else

   
txt_costo_fijo_mes_alto = Val(txt_total_costos_fijos)

txt_gastos_familiares_mes_alto = Val(txt_total_gasto_familiar)


txt_otros_ingresos_mes_alto = Val(txt_total_otros_ingresos)


txt_Deudas_flujo_caja_mes_alto = Val(txt_total_deudas)




If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 > 12 Then
  MsgBox "La Suma Del Ingreso de Meses NO puede ser mayor a 12 ... Revise"
Else

txt_Venta_Total_Promedio_Mes_Alto = Int(txt_venta_total_mes_alto_corregida)


''''''
  ''' construIDO 14-04
txt_vta_formal_promedio_mes_alto = Int(Evaluacion_Perfil.txt_registro_ventas_var * txt_Venta_Total_Promedio_Mes_Alto)


txt_vta_informal_promedio_mes_alto = Int(Evaluacion_Perfil.txt_registro_ventas_dif1_var * txt_Venta_Total_Promedio_Mes_Alto)


txt_venta_total_promedio_anual = Int(((txt_Venta_Total_Promedio_Mes_Alto * numero_meses_tipo_mes_alto) + (txt_Venta_Total_Promedio_Mes_Medio * numero_meses_tipo_mes_medio) + (txt_Venta_Total_Promedio_Mes_Bajo * numero_meses_tipo_mes_bajo)) / 12)


''''''

txt_costo_variable_mes_alto = Int(txt_Venta_Total_Promedio_Mes_Alto * txt_Sub_Total_x1)


txt_resultado_operacional_mes_alto = (txt_Venta_Total_Promedio_Mes_Alto) - (txt_costo_variable_mes_alto) - (txt_costo_fijo_mes_alto)


txt_capacidad_pago_mes_alto = (txt_resultado_operacional_mes_alto) * 1 + (txt_otros_ingresos_mes_alto) * 1 + (txt_segunda_microempresa_mes_alto) * 1 - (txt_Deudas_flujo_caja_mes_alto) * 1 - (txt_gastos_familiares_mes_alto) * 1



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

   'txt_factor = 0.8
   txt_factor_consumo = 0
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5


ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6

ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

  'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5
   
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

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


txt_capacidad_pago_corregida_consumo_mes_alto = Int(txt_capacidad_pago_mes_alto * txt_factor_consumo)


'''''''' CALCULO DE CAPACIDAD DE PAGO SEGUN COMBINACIONES MESES DE VENTAS MEDIOS ALTOS BAJOS
'--------------------------------------------------

'If numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio = 12 And numero_meses_tipo_mes_bajo = 0 Then



If txt_capacidad_pago_corregida_consumo_mes_alto = "" Then
 txt_capacidad_pago_corregida_consumo_mes_alto = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_medio = "" Then
 txt_capacidad_pago_corregida_consumo_mes_medio = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_bajo = "" Then
 txt_capacidad_pago_corregida_consumo_mes_bajo = 0
 End If
 
'''****inibhir !!!! 23-03-2012
'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 1)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 1)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo = 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 3)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 3)

End If


'txt_monto_maximo_credito = Int(txt_capacidad_pago_promedio_corregida_ajustada * txt_leverage)
'txt_monto_maximo_consumo = Int(txt_capacidad_pago_promedio_corregida_consumo * txt_leverage)

txt_costo_variable_mes_alto = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Alto)
txt_costo_variable_mes_medio = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Medio)
txt_costo_variable_mes_bajo = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Bajo)
  
txt_venta_total = Int(txt_venta_total_promedio_anual * 12) * 1
  

End If

cmd_calcular_flujo_Caja.Enabled = True
cmd_calcular_resolucion_cred.Enabled = True

End Sub


Public Sub CALCULO_FLUJO_CAJA_MEDIO()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 <> 12 Then
  
MsgBox "La Suma de Meses Debe Ser igual a 12 Revise ..."

Else

   
txt_costo_fijo_mes_medio = Val(txt_total_costos_fijos)

txt_gastos_familiares_mes_medio = Val(txt_total_gasto_familiar)

txt_otros_ingresos_mes_medio = Val(txt_total_otros_ingresos)

txt_Deudas_flujo_caja_mes_medio = Val(txt_total_deudas)


If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 > 12 Then
  MsgBox "La Suma Del Ingreso de Meses NO puede ser mayor a 12 ... Revise"
Else


txt_Venta_Total_Promedio_Mes_Medio = Int(txt_venta_total_mes_medio_corregida)


''''''
  ''' construIDO 14-04

txt_vta_formal_promedio_mes_medio = Int(Evaluacion_Perfil.txt_registro_ventas_var * txt_Venta_Total_Promedio_Mes_Medio)


txt_vta_informal_promedio_mes_medio = Int(Evaluacion_Perfil.txt_registro_ventas_dif1_var * txt_Venta_Total_Promedio_Mes_Medio)


txt_venta_total_promedio_anual = Int(((txt_Venta_Total_Promedio_Mes_Alto * numero_meses_tipo_mes_alto) + (txt_Venta_Total_Promedio_Mes_Medio * numero_meses_tipo_mes_medio) + (txt_Venta_Total_Promedio_Mes_Bajo * numero_meses_tipo_mes_bajo)) / 12)


''''''


txt_costo_variable_mes_medio = Int(txt_Venta_Total_Promedio_Mes_Medio * txt_Sub_Total_x1)



txt_resultado_operacional_mes_medio = (txt_Venta_Total_Promedio_Mes_Medio) - (txt_costo_variable_mes_medio) - (txt_costo_fijo_mes_medio)



txt_capacidad_pago_mes_medio = (txt_resultado_operacional_mes_medio) * 1 + (txt_otros_ingresos_mes_medio) * 1 + (txt_segunda_microempresa_mes_medio) * 1 - (txt_Deudas_flujo_caja_mes_medio) * 1 - (txt_gastos_familiares_mes_medio) * 1



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

   'txt_factor = 0.8
   txt_factor_consumo = 0
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5


ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6

ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5
   
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

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


txt_capacidad_pago_corregida_ajustada_mes_medio = Int(txt_capacidad_pago_mes_medio * txt_factor)



txt_capacidad_pago_corregida_consumo_mes_medio = Int(txt_capacidad_pago_mes_medio * txt_factor_consumo)


'''''''' CALCULO DE CAPACIDAD DE PAGO SEGUN COMBINACIONES MESES DE VENTAS MEDIOS ALTOS BAJOS
'--------------------------------------------------

'If numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio = 12 And numero_meses_tipo_mes_bajo = 0 Then


If txt_capacidad_pago_corregida_consumo_mes_alto = "" Then
 txt_capacidad_pago_corregida_consumo_mes_alto = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_medio = "" Then
 txt_capacidad_pago_corregida_consumo_mes_medio = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_bajo = "" Then
 txt_capacidad_pago_corregida_consumo_mes_bajo = 0
 End If




'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 1)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 1)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo = 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 3)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 3)

End If


'txt_monto_maximo_credito = Int(txt_capacidad_pago_promedio_corregida_ajustada * txt_leverage)
'txt_monto_maximo_consumo = Int(txt_capacidad_pago_promedio_corregida_consumo * txt_leverage)


txt_costo_variable_mes_medio = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Medio)

  
txt_venta_total = Int(txt_venta_total_promedio_anual * 12) * 1
  

End If

cmd_calcular_flujo_Caja.Enabled = True
cmd_calcular_resolucion_cred.Enabled = True
End Sub

Public Sub CALCULO_FLUJO_CAJA_BAJO()



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 <> 12 Then
  
MsgBox "La Suma de Meses Debe Ser igual a 12 Revise ..."

Else

txt_costo_fijo_mes_bajo = Val(txt_total_costos_fijos)
txt_gastos_familiares_mes_bajo = Val(txt_total_gasto_familiar)
txt_otros_ingresos_mes_bajo = Val(txt_total_otros_ingresos)
txt_Deudas_flujo_caja_mes_bajo = Val(txt_total_deudas)

If Val(numero_meses_tipo_mes_alto) * 1 + Val(numero_meses_tipo_mes_medio) * 1 + Val(numero_meses_tipo_mes_bajo) * 1 > 12 Then
  MsgBox "La Suma Del Ingreso de Meses NO puede ser mayor a 12 ... Revise"
Else

txt_Venta_Total_Promedio_Mes_Bajo = Int(txt_venta_total_mes_bajo_corregida)

''''''
  ''' construIDO 14-04

txt_vta_formal_promedio_mes_bajo = Int(Evaluacion_Perfil.txt_registro_ventas_var * txt_Venta_Total_Promedio_Mes_Bajo)
txt_vta_informal_promedio_mes_bajo = Int(Evaluacion_Perfil.txt_registro_ventas_dif1_var * txt_Venta_Total_Promedio_Mes_Bajo)

txt_venta_total_promedio_anual = Int(((txt_Venta_Total_Promedio_Mes_Alto * numero_meses_tipo_mes_alto) + (txt_Venta_Total_Promedio_Mes_Medio * numero_meses_tipo_mes_medio) + (txt_Venta_Total_Promedio_Mes_Bajo * numero_meses_tipo_mes_bajo)) / 12)


''''''

txt_costo_variable_mes_bajo = Int(txt_Venta_Total_Promedio_Mes_Bajo * txt_Sub_Total_x1)

txt_resultado_operacional_mes_bajo = (txt_Venta_Total_Promedio_Mes_Bajo) - (txt_costo_variable_mes_bajo) - (txt_costo_fijo_mes_bajo)
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

   'txt_factor = 0.8
   txt_factor_consumo = 0
   'txt_leverage = 7
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Con Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5


ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Excelente" Then

   'txt_factor = 0.7
   txt_factor_consumo = 0
   'txt_leverage = 6

ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Bueno" Then

   'txt_factor = 0.6
   txt_factor_consumo = 0
   'txt_leverage = 5
   
   
ElseIf txt_tipo_cliente_form_evaluacion = "Nuevo Sin Historia Sbif" And txt_tipo_riesgo_form_evaluacion = "Regular" Then

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

txt_capacidad_pago_corregida_ajustada_mes_bajo = Int(txt_capacidad_pago_mes_bajo * txt_factor)
txt_capacidad_pago_corregida_consumo_mes_bajo = Int(txt_capacidad_pago_mes_bajo * txt_factor_consumo)

'''''''' CALCULO DE CAPACIDAD DE PAGO SEGUN COMBINACIONES MESES DE VENTAS MEDIOS ALTOS BAJOS
'--------------------------------------------------


If txt_capacidad_pago_corregida_consumo_mes_alto = "" Then
 txt_capacidad_pago_corregida_consumo_mes_alto = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_medio = "" Then
 txt_capacidad_pago_corregida_consumo_mes_medio = 0
 End If
 
If txt_capacidad_pago_corregida_consumo_mes_bajo = "" Then
 txt_capacidad_pago_corregida_consumo_mes_bajo = 0
 End If


'If numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio = 12 And numero_meses_tipo_mes_bajo = 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 1)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 1)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo = 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto = 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 2)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 2)

'ElseIf numero_meses_tipo_mes_alto > 0 And numero_meses_tipo_mes_medio > 0 And numero_meses_tipo_mes_bajo > 0 Then

'txt_capacidad_pago_promedio_corregida_ajustada = Int(((txt_capacidad_pago_corregida_ajustada_mes_alto) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_medio) * 1 + (txt_capacidad_pago_corregida_ajustada_mes_bajo) * 1) / 3)
'txt_capacidad_pago_promedio_corregida_consumo = Int(((txt_capacidad_pago_corregida_consumo_mes_alto) * 1 + (txt_capacidad_pago_corregida_consumo_mes_medio) * 1 + (txt_capacidad_pago_corregida_consumo_mes_bajo) * 1) / 3)

End If


'txt_monto_maximo_credito = Int(txt_capacidad_pago_promedio_corregida_ajustada * txt_leverage)
'txt_monto_maximo_consumo = Int(txt_capacidad_pago_promedio_corregida_consumo * txt_leverage)

txt_costo_variable_mes_bajo = Int(txt_Sub_Total_x1 * txt_Venta_Total_Promedio_Mes_Bajo)
  
txt_venta_total = Int(txt_venta_total_promedio_anual * 12) * 1
  

End If

cmd_calcular_flujo_Caja.Enabled = True
cmd_calcular_resolucion_cred.Enabled = True
End Sub
