VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Datos_Sinacofi 
   Caption         =   "::::::::: Ingreso Datos De Sinacofi"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   OleObjectBlob   =   "Datos_Sinacofi.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Datos_Sinacofi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub cbx_protesto_Change()

End Sub

Private Sub cmd_ingreso_datos_sinacofi_Click()

If txt_rut_cliente <> "" And txt_dv <> "" And cbx_protesto <> "" And cbx_mora <> "" And cbx_boletin <> "" And _
   txt_score <> "" And txt_cod_observacion <> "" Then


Call conectarBD

    ssql = "select rut" _
            & " from TBL_MICRO_SINACOFI" _
            & " where RUT = '" & txt_rut_cliente & "'"
    
    Set rst = cnn.Execute(ssql, , adCmdText)
 

If rst.EOF Then
        
    ssql = "INSERT INTO TBL_MICRO_SINACOFI " _
    & "([Rut], [DV],[PROTESTO], [MORA],[INFRACCION_PREV],[COD_OBSERVACION],[SCORE],[FECHA_CONSULTA])" _
    & " VALUES (('" & txt_rut_cliente & "')," _
    & "('" & txt_dv & "') , ('" & cbx_protesto & "') ,('" & cbx_mora & "'), ('" & cbx_boletin & "'),('" & txt_cod_observacion & "'),('" & txt_score & "'),convert(varchar,getdate(),111))"
    
    cnn.Execute ssql
    
    MsgBox "Ingresado Correctamente"
    
    txt_rut_cliente.Enabled = False
    txt_dv.Enabled = False
    
    cbx_protesto = Empty
    cbx_mora = Empty
    cbx_boletin = Empty
    
    
    txt_score = Empty
    txt_cod_observacion = Empty
    
    Datos_Sinacofi.cbx_mora.Enabled = False
    Datos_Sinacofi.cbx_protesto.Enabled = False
    Datos_Sinacofi.cbx_boletin.Enabled = False
    Datos_Sinacofi.txt_score.Enabled = False
    Datos_Sinacofi.txt_cod_observacion.Enabled = False
    
    
    
Else

    ssql = "UPDATE TBL_MICRO_SINACOFI " _
    & " SET protesto = '" & cbx_protesto & "' , mora = '" & cbx_mora & "', infraccion_prev ='" & cbx_boletin & "', cod_observacion = '" & txt_cod_observacion & "' ,score = '" & txt_score & "', fecha_consulta = convert(varchar,getdate(),111)" _
    & " where rut = '" & txt_rut_cliente & "'"
    cnn.Execute ssql

    
    MsgBox "Actualizado Correctamente"
    
    txt_rut_cliente.Enabled = False
    txt_dv.Enabled = False
    
    cbx_protesto = Empty
    cbx_mora = Empty
    cbx_boletin = Empty
    
    txt_score = Empty
    txt_cod_observacion = Empty
  
    Datos_Sinacofi.cbx_mora.Enabled = False
    Datos_Sinacofi.cbx_protesto.Enabled = False
    Datos_Sinacofi.cbx_boletin.Enabled = False
    Datos_Sinacofi.txt_score.Enabled = False
    Datos_Sinacofi.txt_cod_observacion.Enabled = False
    
  
    
End If

    
Else

MsgBox "Para grabar correctamente debe ingresar todos los datos", vbCritical

End If
    
End Sub



Private Sub cmd_reingreso_Click()

    Datos_Sinacofi.cbx_mora.Clear
    Datos_Sinacofi.cbx_protesto.Clear
    Datos_Sinacofi.cbx_boletin.Clear
    
    cbx_mora.AddItem "Cumple"
    cbx_mora.AddItem "No Cumple"
    
    cbx_protesto.AddItem "Cumple"
    cbx_protesto.AddItem "No Cumple"
    
    cbx_boletin.AddItem "Cumple"
    cbx_boletin.AddItem "No Cumple"

    txt_rut_cliente.Enabled = True
    txt_dv.Enabled = True

    Datos_Sinacofi.txt_rut_cliente = Empty
    Datos_Sinacofi.txt_dv = Empty
    Datos_Sinacofi.cbx_protesto.Enabled = True
    Datos_Sinacofi.cbx_mora.Enabled = True
    Datos_Sinacofi.cbx_boletin.Enabled = True
    Datos_Sinacofi.txt_score.Enabled = True
    Datos_Sinacofi.txt_cod_observacion.Enabled = True
    
    Datos_Sinacofi.txt_rut_cliente.SetFocus

End Sub

Private Sub cmd_volver_menu_p_Click()
Datos_Sinacofi.Hide
Menu_Principal_Micro.Show
End Sub

Private Sub txt_cod_observacion_AfterUpdate()

    If Not IsNumeric(txt_cod_observacion) Or txt_cod_observacion > 15 Or txt_cod_observacion = 1 Then
        diga = MsgBox("El Código de Observación Debe ser Numerico o NO Corresponde. Favor Ingrese Solo Números", vbOKOnly)
        
        txt_cod_observacion = Empty
        txt_cod_observacion.SetFocus
      
    End If

End Sub


Private Sub txt_dv_AfterUpdate()

    Datos_Sinacofi.cbx_mora.Enabled = False
    Datos_Sinacofi.cbx_protesto.Enabled = False
    Datos_Sinacofi.cbx_boletin.Enabled = False
    Datos_Sinacofi.txt_score.Enabled = False
    Datos_Sinacofi.txt_cod_observacion.Enabled = False

    If txt_dv <> txt_dv_compara Then
        MsgBox ("Rut Invalido Revise ...")

Else

    Datos_Sinacofi.cbx_mora.Enabled = True
    Datos_Sinacofi.cbx_protesto.Enabled = True
    Datos_Sinacofi.cbx_boletin.Enabled = True
    Datos_Sinacofi.txt_score.Enabled = True
    Datos_Sinacofi.txt_cod_observacion.Enabled = True
    
    Datos_Sinacofi.cbx_mora.SetFocus

End If

End Sub

Private Sub txt_dv_Change()
    
    txt_estado_credito = Empty

    Dim I As Integer

    txt_dv = UCase(txt_dv)
    I = Len(txt_dv)
    txt_dv.SelStart = I
End Sub

Private Sub txt_rut_cliente_Change()
    
    Dim diga As Variant
       
    If Not IsNumeric(txt_rut_cliente) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_cliente = Empty
      End If
      
  
  
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_rut_cliente = Replace(txt_rut_cliente, "-", "")
    txt_rut_cliente = Replace(txt_rut_cliente, ".", "")
    txt_rut_cliente = Replace(txt_rut_cliente, ",", "")
    txt_rut_cliente = Replace(txt_rut_cliente, " ", "")
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
    'fin digito verificador
End Sub

Private Sub txt_score_AfterUpdate()

    If Not IsNumeric(txt_score) Or txt_score > 999 Then
        diga = MsgBox("El Indicador de Riesgo Debe Ser Numérico y el No Matoy a 999. Favor Ingrese Solo Números", vbOKOnly)
        
        txt_score = Empty
        txt_score.SetFocus
      
    End If

End Sub


