VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Carta_Cliente 
   Caption         =   "::::: Informe Rechazo"
   ClientHeight    =   9255.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585.001
   OleObjectBlob   =   "Carta_Cliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Carta_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_imprimir_Click()

Carta_Cliente.cmd_mostrar_rechazos.Visible = False
'Carta_Cliente.cmd_menu_rechazos.Visible = False
Carta_Cliente.cmd_volver_estado_resolucion.Visible = False
Carta_Cliente.cmd_imprimir.Visible = False
Carta_Cliente.cmd_salir_sistema.Visible = False

Carta_Cliente.PrintForm

Carta_Cliente.cmd_mostrar_rechazos.Visible = True
'Carta_Cliente.cmd_menu_rechazos.Visible = True
Carta_Cliente.cmd_volver_estado_resolucion.Visible = True
Carta_Cliente.cmd_imprimir.Visible = True
Carta_Cliente.cmd_salir_sistema.Visible = True

End Sub

Private Sub cmd_mostrar_rechazos_Click()

'cmd_imprimir.Enabled = False


Dim fec1
fec1 = Format(Date, "dd/mm/yyyy")
txt_fecha_dia = fec1


Call conectarBD


        ssql = "select cod9,cod10,cod11,cod13,cod14,cod15,cod16,cod18" _
        & " from TBL_MICRO_maestro_RECHAZOS_SERNAC_F" _
        & " where rut_cliente = '" & txt_rut_cliente & "'" _
        & " and n_Solicitud = '" & txt_n_solicitud & "'" _
        
        Set rst = cnn.Execute(ssql, , adCmdText)
        
    If Not rst.EOF Then
        
        If rst!Cod9 <> 0 Then
        'txt_aceptado9.Visible = False
        txt_rechazado9.Visible = True
        txt_resultado_rechazo1.Visible = True
        Else
        'txt_aceptado9.Visible = True
        End If
        
        If rst!Cod10 <> 0 Then
        'txt_aceptado10.Visible = False
        txt_rechazado10.Visible = True
        txt_resultado_rechazo2.Visible = True
        Else
        'txt_aceptado10.Visible = True
        End If
        
        If rst!Cod11 <> 0 Then
        'txt_aceptado11.Visible = False
        txt_rechazado11.Visible = True
        txt_resultado_rechazo3.Visible = True
        Else
        'txt_aceptado11.Visible = True
        End If
                
        If rst!Cod13 <> 0 Then
        'txt_aceptado13.Visible = False
        txt_rechazado13.Visible = True
        txt_resultado_rechazo4.Visible = True
        Else
        'txt_aceptado13.Visible = True
        End If
                        
        If rst!Cod14 <> 0 Then
        'txt_aceptado14.Visible = False
        txt_rechazado14.Visible = True
        txt_resultado_rechazo5.Visible = True
        Else
        'txt_aceptado14.Visible = True
        End If
                                
        If rst!Cod15 <> 0 Then
        'txt_aceptado15.Visible = False
        txt_rechazado15.Visible = True
        txt_resultado_rechazo6.Visible = True
        Else
        'txt_aceptado15.Visible = True
        End If
                                        
        If rst!Cod16 <> 0 Then
        'txt_aceptado16.Visible = False
        txt_rechazado16.Visible = True
        txt_resultado_rechazo7.Visible = True
        Else
        'txt_aceptado16.Visible = True
        End If
                                                
        If rst!Cod18 <> 0 Then
        'txt_aceptado18.Visible = False
        txt_rechazado18.Visible = True
        txt_resultado_rechazo8.Visible = True
        Else
        'txt_aceptado18.Visible = True
        End If
        
        'cmd_imprimir.Enabled = True
         
    End If

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub txt_cod1_Change()
If txt_cod1 = 0 Then
txt_cod1 = Empty
End If
End Sub

Private Sub txt_cod2_Change()

If txt_cod2 = 0 Then
txt_cod2 = Empty
End If

End Sub

Private Sub txt_cod3_Change()
If txt_cod3 = 0 Then
txt_cod3 = Empty
End If
End Sub

Private Sub txt_cod4_Change()
If txt_cod4 = 0 Then
txt_cod4 = Empty
End If
End Sub

Private Sub txt_cod5_Change()
If txt_cod5 = 0 Then
txt_cod5 = Empty
End If
End Sub

Private Sub txt_cod6_Change()
If txt_cod6 = 0 Then
txt_cod6 = Empty
End If
End Sub

Private Sub txt_cod7_Change()
If txt_cod7 = 0 Then
txt_cod7 = Empty
End If
End Sub

Private Sub txt_cod8_Change()
If txt_cod8 = 0 Then
txt_cod8 = Empty
End If
End Sub

Private Sub txt_glosa1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If txt_cod1 = "9" Then
   txt_glosa1 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa1 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa1 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa1 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa1 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa1 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa1 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa1 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa3 = ""

End If
End Sub

Private Sub txt_glosa1_Change()

End Sub

Private Sub txt_glosa2_Change()

End Sub

Private Sub txt_glosa3_Change()
If txt_cod1 = "9" Then
   txt_glosa3 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa3 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa3 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa3 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa3 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa3 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa3 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa3 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa3 = ""

End If
End Sub

Private Sub txt_glosa4_Change()
If txt_cod1 = "9" Then
   txt_glosa4 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa4 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa4 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa4 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa4 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa4 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa4 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa4 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa4 = ""

End If
End Sub

Private Sub txt_glosa5_Change()
If txt_cod1 = "9" Then
   txt_glosa5 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa5 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa5 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa5 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa5 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa5 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa5 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa5 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa5 = ""

End If
End Sub

Private Sub txt_glosa6_Change()
If txt_cod1 = "9" Then
   txt_glosa6 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa6 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa6 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa6 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa6 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa6 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa6 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa6 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa6 = ""

End If
End Sub

Private Sub txt_glosa7_Change()
If txt_cod1 = "9" Then
   txt_glosa7 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa7 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa7 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa7 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa7 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa7 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa7 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa7 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa7 = ""

End If
End Sub

Private Sub txt_glosa8_Change()
If txt_cod1 = "9" Then
   txt_glosa8 = "Morosidad o Protestos Vigentes"
    
        ElseIf txt_cod1 = "10" Then
        txt_glosa8 = "Excesiva Carga Financiera o de Endeudamiento"
       
        ElseIf txt_cod1 = "11" Then
        txt_glosa8 = "Incumplimiento Previo"
    
        ElseIf txt_cod1 = "13" Then
        txt_glosa8 = "Incumplimiento en Parametros de politica de creditos"
    
        ElseIf txt_cod1 = "14" Then
        txt_glosa8 = "Incumplimiento en Parametros de Score"
    
        ElseIf txt_cod1 = "15" Then
        txt_glosa8 = "Incumplimiento en Parametros de Edad"
    
        ElseIf txt_cod1 = "16" Then
        txt_glosa8 = "Incumplimiento en Parametros Renta"
    
        ElseIf txt_cod1 = "18" Then
        txt_glosa8 = "Insuficiencia de Garantias"
    
Else
    
        txt_glosa8 = ""

End If
End Sub

Private Sub cmd_salir_sistema_Click()
    Workbooks("Sistema_Evaluacion_Gestion_Micro1.xls").Close
    Application.Quit
End Sub

Private Sub cmd_volver_estado_resolucion_Click()

    Unload Estado_Resolucion_Final
    Unload Carta_Cliente
    Unload Ficha_Cliente_Micro
    Unload Evaluacion_Perfil
    Unload Metodologia_Activo_Circulante
    Unload Metodologia_IVA1
    Unload Metodologia_Maxima_Prod

    Menu_Principal_Micro.Show

End Sub

Private Sub Image1_Click()

End Sub
