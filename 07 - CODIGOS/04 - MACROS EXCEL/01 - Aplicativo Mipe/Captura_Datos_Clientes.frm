VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Captura_Datos_Clientes 
   Caption         =   ":::: Captura Datos Microempresa"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   OleObjectBlob   =   "Captura_Datos_Clientes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Captura_Datos_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cbx_codigo_sucursal_Change()

cbx_cod_ejecutivo.Clear

Call conectarBD
ssql = "select codigo_ejecutivo +'      '+ nombre_ejecutivo +' '+ apellido_ejecutivo as EJECUTIVO FROM TBL_ejecutivo " _
    & " WHERE (CODIGO_EJECUTIVO <>9999 and CODIGO_EJECUTIVO <>999) " _
    & " and (cargo_ejecutivo ='EJECUTIVO MICROEMPRESA' OR  cargo_ejecutivo ='EVALUADOR MICROEMPRESA' OR cargo_ejecutivo = 'EJECUTIVO TLMK')" _
    & " AND '" & cbx_codigo_sucursal & "' = codigo_sucursal" _
    & " ORDER BY codigo_sucursal, CODIGO_EJECUTIVO"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_cod_ejecutivo.AddItem rst!EJECUTIVO
        rst.MoveNext
    Loop

End Sub

Private Sub cmd_salir_sistema_Click()
    Workbooks("Sistema_Evaluacion_Gestion_Micro.xls").Close
    Application.Quit
End Sub


Private Sub DV_Txt_Change()
Dim I As Integer

DV_Txt = UCase(DV_Txt)
I = Len(DV_Txt)
DV_Txt.SelStart = I
End Sub



Private Sub Label22_Click()

End Sub

Private Sub Nombre_Cliente_txt_Change()
Dim I As Integer

Nombre_Cliente_txt = UCase(Nombre_Cliente_txt)
I = Len(Nombre_Cliente_txt)
Nombre_Cliente_txt.SelStart = I

End Sub

Private Sub Apel_Paterno_txt_Change()
Dim I As Integer

Apel_Paterno_txt = UCase(Apel_Paterno_txt)
I = Len(Apel_Paterno_txt)
Apel_Paterno_txt.SelStart = I

End Sub
Private Sub Apel_Materno_txt_Change()
Dim I As Integer

Apel_Materno_txt = UCase(Apel_Materno_txt)
I = Len(Apel_Materno_txt)
Apel_Materno_txt.SelStart = I

End Sub
Private Sub CALLE_txt_Change()

CALLE_txt = UCase(CALLE_txt)
I = Len(CALLE_txt)
CALLE_txt.SelStart = I

End Sub

Private Sub Pobla_txt_Change()

Pobla_txt = UCase(Pobla_txt)
I = Len(Pobla_txt)
Pobla_txt.SelStart = I

End Sub
Private Sub comuna_txt_Change()

Comuna_txt = UCase(Comuna_txt)
I = Len(Comuna_txt)
Comuna_txt.SelStart = I

End Sub
Private Sub email_txt_Change()
Dim I As Integer

email_txt = LCase(email_txt)
I = Len(email_txt)
email_txt.SelStart = I
End Sub

Private Sub cbx_motivo_ingreso_click()
Dim fecha
fecha = Date
txt_fecha_actual = fecha

End Sub


Private Sub cbx_motivo_ingreso_Change()
If cbx_motivo_ingreso = "Agrega Telefono" Or cbx_motivo_ingreso = "Telefono Erroneo" Then
Call LIMPIA_CAMPOS
Area1_txt.Visible = True
Area1_txt.Enabled = True
Telef1_txt.Visible = True
Telef1_txt.Enabled = True
Area2_txt.Visible = True
Area2_txt.Enabled = True
Telef2_txt.Visible = True
Telef2_txt.Enabled = True
Area3_txt.Visible = True
Area3_txt.Enabled = True
Telef3_txt.Visible = True
Telef3_txt.Enabled = True
Nombre_Cliente_txt.Visible = False
Apel_Paterno_txt.Visible = False
Apel_Materno_txt.Visible = False
CALLE_txt.Visible = False
Numero_txt.Visible = False
Dpto_txt.Visible = False
Pobla_txt.Visible = False
Comuna_txt.Visible = False
email_txt.Visible = False
MsgBox ("Los Celulares Deben Ser Ingresado Con Codigo Area 9 y El Número De Celular Debe Ser De 8 Dígitos")

ElseIf cbx_motivo_ingreso = "Agrega Direccion" Then
Call LIMPIA_CAMPOS
CALLE_txt.Visible = True
Numero_txt.Visible = True
Dpto_txt.Visible = True
Pobla_txt.Visible = True
Comuna_txt.Visible = True
CALLE_txt.Enabled = True
Numero_txt.Enabled = True
Dpto_txt.Enabled = True
Pobla_txt.Enabled = True
Comuna_txt.Enabled = True

Nombre_Cliente_txt.Visible = False
Apel_Paterno_txt.Visible = False
Apel_Materno_txt.Visible = False
Area1_txt.Visible = False
Telef1_txt.Visible = False
Area2_txt.Visible = False
Telef2_txt.Visible = False
Area3_txt.Visible = False
Telef3_txt.Visible = False
email_txt.Visible = False

ElseIf cbx_motivo_ingreso = "Direccion Erronea" Then
Call LIMPIA_CAMPOS
CALLE_txt.Visible = False
Numero_txt.Visible = False
Dpto_txt.Visible = False
Pobla_txt.Visible = False
Comuna_txt.Visible = False
CALLE_txt.Enabled = False
Numero_txt.Enabled = False
Dpto_txt.Enabled = False
Pobla_txt.Enabled = False
Comuna_txt.Enabled = False

Nombre_Cliente_txt.Visible = False
Apel_Paterno_txt.Visible = False
Apel_Materno_txt.Visible = False
Area1_txt.Visible = False
Telef1_txt.Visible = False
Area2_txt.Visible = False
Telef2_txt.Visible = False
Area3_txt.Visible = False
Telef3_txt.Visible = False
email_txt.Visible = False


ElseIf cbx_motivo_ingreso = "Agrega Telef. y Direc." Then
Call LIMPIA_CAMPOS
CALLE_txt.Visible = True
Numero_txt.Visible = True
Dpto_txt.Visible = True
Pobla_txt.Visible = True
Comuna_txt.Visible = True
Area1_txt.Visible = True
Telef1_txt.Visible = True
Area2_txt.Visible = True
Telef2_txt.Visible = True
Area3_txt.Visible = True
Telef3_txt.Visible = True

CALLE_txt.Enabled = True
Numero_txt.Enabled = True
Dpto_txt.Enabled = True
Pobla_txt.Enabled = True
Comuna_txt.Enabled = True
Area1_txt.Enabled = True
Telef1_txt.Enabled = True
Area2_txt.Enabled = True
Telef2_txt.Enabled = True
Area3_txt.Enabled = True
Telef3_txt.Enabled = True

Nombre_Cliente_txt.Visible = False
Apel_Paterno_txt.Visible = False
Apel_Materno_txt.Visible = False
email_txt.Visible = False

MsgBox ("Los Celulares Deben Ser Ingresado Con Codigo Area 9 y El Número De Celular Debe Ser De 8 Dígitos")


ElseIf cbx_motivo_ingreso = "Agrega E-Mail" Then
Call LIMPIA_CAMPOS
Nombre_Cliente_txt.Visible = True
Apel_Paterno_txt.Visible = True
Apel_Materno_txt.Visible = True
email_txt.Visible = True

Nombre_Cliente_txt.Enabled = True
Apel_Paterno_txt.Enabled = True
Apel_Materno_txt.Enabled = True
email_txt.Enabled = True

CALLE_txt.Visible = False
Numero_txt.Visible = False
Dpto_txt.Visible = False
Pobla_txt.Visible = False
Comuna_txt.Visible = False
Area1_txt.Visible = False
Telef1_txt.Visible = False
Area2_txt.Visible = False
Telef2_txt.Visible = False
Area3_txt.Visible = False
Telef3_txt.Visible = False

ElseIf cbx_motivo_ingreso = "Fallecido" Then
Call LIMPIA_CAMPOS
Nombre_Cliente_txt.Visible = True
Apel_Paterno_txt.Visible = True
Apel_Materno_txt.Visible = True

Nombre_Cliente_txt.Enabled = True
Apel_Paterno_txt.Enabled = True
Apel_Materno_txt.Enabled = True

CALLE_txt.Visible = False
Numero_txt.Visible = False
Dpto_txt.Visible = False
Pobla_txt.Visible = False
Comuna_txt.Visible = False
Area1_txt.Visible = False
Telef1_txt.Visible = False
Area2_txt.Visible = False
Telef2_txt.Visible = False
Area3_txt.Visible = False
Telef3_txt.Visible = False
email_txt.Visible = False

ElseIf cbx_motivo_ingreso = "Telef. y Direc. Erronea" Then
Call LIMPIA_CAMPOS
CALLE_txt.Visible = False
Numero_txt.Visible = False
Dpto_txt.Visible = False
Pobla_txt.Visible = False
Comuna_txt.Visible = False
Area1_txt.Visible = True
Telef1_txt.Visible = True
Area2_txt.Visible = True
Telef2_txt.Visible = True
Area3_txt.Visible = True
Telef3_txt.Visible = True

CALLE_txt.Enabled = False
Numero_txt.Enabled = False
Dpto_txt.Enabled = False
Pobla_txt.Enabled = False
Comuna_txt.Enabled = False
Area1_txt.Enabled = True
Telef1_txt.Enabled = True
Area2_txt.Enabled = True
Telef2_txt.Enabled = True
Area3_txt.Enabled = True
Telef3_txt.Enabled = True

Nombre_Cliente_txt.Visible = False
Apel_Paterno_txt.Visible = False
Apel_Materno_txt.Visible = False
email_txt.Visible = False

MsgBox ("Los Celulares Deben Ser Ingresado Con Codigo Area 9 y El Número De Celular Debe Ser De 8 Dígitos")


End If



End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

      If CloseMode = vbFormControlMenu Then
            'ActiveWorkbook.Save
            Workbooks("Sistema_Microempresa.xls").Close
            Application.Quit
            Cancel = True
      End If

 End Sub
 
 
 Private Sub Rut_Cliente_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   
    Dim diga As Variant
    
    If Not IsNumeric(Rut_Cliente_Txt) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        Rut_Cliente_Txt = Empty
      
      End If
      
   
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    Rut_Cliente_Txt = Replace(Rut_Cliente_Txt, "-", "")
    Vari3 = 2
    For I = 0 To Len(Rut_Cliente_Txt) - 1
     If Left(Right(Rut_Cliente_Txt, I + 1), 1) <> "." Then
      Vari1 = Vari1 + Left(Right(Rut_Cliente_Txt, I + 1), 1) * Vari3
      Vari2 = Vari1 Mod 11
      Select Case Vari2
       Case 0
        dv_compara_txt.Text = "0"
       Case 1
        dv_compara_txt.Text = "K"
       Case Else
        dv_compara_txt.Text = 11 - Vari2
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
  
 
  
    Private Sub area1_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
   
    
    If Not IsNumeric(Area1_txt) Then
        diga = MsgBox("El código De Área Debe Ser Numérico. Favor Ingresar Sólo Números", vbOKOnly)
        Area1_txt = Empty
   
   End If
      
    
    Area1_txt.SetFocus
    
  End Sub
      Private Sub telef1_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
    Dim largo_celular As Integer
    
    
    If Not IsNumeric(Telef1_txt) Then
        diga = MsgBox("El Télefono Debe Ser Numérico. Favor ingresar Sólo Números", vbOKOnly)
        Telef1_txt = Empty
    End If
    Telef1_txt.SetFocus
  End Sub
      Private Sub area2_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
    
    If Not IsNumeric(Area2_txt) Then
        diga = MsgBox("El código De Área Debe Ser Numérico. Favor Ingresar Sólo Números", vbOKOnly)
        Area2_txt = Empty
    End If
    Area2_txt.SetFocus
  End Sub
      Private Sub telef2_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
    
    If Not IsNumeric(Telef2_txt) Then
        diga = MsgBox("El Télefono Debe Ser Numérico. Favor ingresar Sólo Números", vbOKOnly)
        Telef2_txt = Empty
    End If
    Telef2_txt.SetFocus
  End Sub
    Private Sub area3_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
    
    If Not IsNumeric(Area3_txt) Then
        diga = MsgBox("El código De Área Debe Ser Numérico. Favor Ingresar Sólo Números", vbOKOnly)
        Area3_txt = Empty
    End If
    Area3_txt.SetFocus
  End Sub
   Private Sub telef3_Txt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim diga As Variant
    
    If Not IsNumeric(Telef3_txt) Then
        diga = MsgBox("El Télefono Debe Ser Numérico. Favor ingresar Sólo Números", vbOKOnly)
        Telef3_txt = Empty
    End If
    Telef3_txt.SetFocus
  End Sub
 
Private Sub UserForm_Activate()
cbx_motivo_ingreso.AddItem "Agrega Telefono"
cbx_motivo_ingreso.AddItem "Telefono Erroneo"
cbx_motivo_ingreso.AddItem "Agrega Direccion"
'cbx_motivo_ingreso.AddItem "Direccion Erronea"
cbx_motivo_ingreso.AddItem "Agrega Telef. y Direc."
'cbx_motivo_ingreso.AddItem "Telef. y Direc. Erronea"
cbx_motivo_ingreso.AddItem "Agrega E-Mail"
cbx_motivo_ingreso.AddItem "Cliente Dependiente"
cbx_motivo_ingreso.AddItem "Fallecido"
cbx_motivo_ingreso.AddItem "No acredita Ingresos"



' DESACTIVA TODAS LOS TEXT

Nombre_Cliente_txt.Enabled = False
Apel_Paterno_txt.Enabled = False
Apel_Materno_txt.Enabled = False
CALLE_txt.Enabled = False
Numero_txt.Enabled = False
Dpto_txt.Enabled = False
Pobla_txt.Enabled = False
Comuna_txt.Enabled = False
Area1_txt.Enabled = False
Telef1_txt.Enabled = False
Area2_txt.Enabled = False
Telef2_txt.Enabled = False
Area3_txt.Enabled = False
Telef3_txt.Enabled = False
email_txt.Enabled = False

End Sub

Private Sub UserForm_Initialize()

Call conectarBD

    
    '''''''''' TRAE CODIGO_SUCURSALES '''''''''
    ssql = "select distinct(codigo_sucursal) FROM TBL_ejecutivo " _
    & " where (cargo_ejecutivo ='EJECUTIVO MICROEMPRESA' OR  cargo_ejecutivo ='EVALUADOR MICROEMPRESA' or cargo_ejecutivo='EJECUTIVO TLMK')" _
    & " ORDER BY codigo_sucursal"
              
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Do Until rst.EOF
        cbx_codigo_sucursal.AddItem rst!codigo_sucursal
        rst.MoveNext
    Loop
    
    End Sub

Private Sub insertar_cm_Click()

    Call conectarBD


    If DV_Txt = dv_compara_txt Then
    If cbx_cod_ejecutivo <> "" And Rut_Cliente_Txt <> "" And DV_Txt <> "" And cbx_motivo_ingreso <> "" Then
    
    'INSERTA DATOS A TABLA SQL

    'On Error GoTo MostrarError ' **** Verifica si al error en el procedimiento proxima lineas

    
    ssql = "INSERT INTO TBL_GESTION_CLIENTE_SUCURSAL " _
    & "([cod_ejecutivo], " _
    & " [sucursal], " _
    & " [estado_gestion]," _
    & " [rut_cliente]," _
    & " [dv]," _
    & " [nombre_cliente]," _
    & " [apellido_paterno]," _
    & " [apellido_materno],[calle],[numero],[dpto],[villa],[comuna],[cod1],[telef1]" _
    & " ,[cod2],[telef2],[cod3],[telef3], [email],[fecha_ingreso],[Negocio])" _
    & " VALUES (substring('" & cbx_cod_ejecutivo & "',1,4)" _
    & ", '" & cbx_codigo_sucursal & "' , '" & cbx_motivo_ingreso & "' , " & Rut_Cliente_Txt & " , '" & DV_Txt & "' " _
    & ", '" & Nombre_Cliente_txt & "' , '" & Apel_Paterno_txt & "' , '" & Apel_Materno_txt & "' " _
    & ", '" & CALLE_txt & "' , '" & Numero_txt & "' , '" & Dpto_txt & "' " _
    & ", '" & Pobla_txt & "' , '" & Comuna_txt & "' " _
    & ", '" & Area1_txt & "'" _
    & ", '" & Telef1_txt & "'" _
    & ", '" & Area2_txt & "'" _
    & ", '" & Telef2_txt & "'" _
    & ", '" & Area3_txt & "'" _
    & ", '" & Telef3_txt & "'" _
    & ", '" & email_txt & "'" _
    & ", '" & txt_fecha_actual & "', 'Micro')"
    
  
    cnn.Execute ssql
        
        
        cbx_cod_ejecutivo = Empty
        cbx_motivo_ingreso = Empty
        cbx_codigo_sucursal = Empty
        Rut_Cliente_Txt = Empty
        DV_Txt = Empty
        Nombre_Cliente_txt = Empty
        Apel_Paterno_txt = Empty
        Apel_Materno_txt = Empty
        CALLE_txt = Empty
        Numero_txt = Empty
        Dpto_txt = Empty
        Pobla_txt = Empty
        Comuna_txt = Empty
        Area1_txt = Empty
        Telef1_txt = Empty
        Area2_txt = Empty
        Telef2_txt = Empty
        Area3_txt = Empty
        Telef3_txt = Empty
        email_txt = Empty
        
        Exit Sub
'MostrarError:
 '  MsgBox "No Tiene Acceso Para Ingresar Datos... Avise A Su Agente Para Revisar Mantenedor y Permisos a la Base De Datos", vbCritical
'-----
        
        
    Else
        MsgBox "No Cumple Con El Mínimo De Datos Solictado. Favor Revise"
    End If
Else
    MsgBox "Rut O Dígito Mal Ingresado. Favor Reingrese Datos"
    End If

End Sub
    
    
Private Sub Cerrar_Click()
Unload Captura_Datos_Clientes
Call Menu_Principal_Micro.Show

End Sub

Public Function LIMPIA_CAMPOS()
CALLE_txt = Empty
Numero_txt = Empty
Dpto_txt = Empty
Pobla_txt = Empty
Comuna_txt = Empty
Area1_txt = Empty
Telef1_txt = Empty
Area2_txt = Empty
Telef2_txt = Empty
Area3_txt = Empty
Telef3_txt = Empty
email_txt = Empty
End Function
