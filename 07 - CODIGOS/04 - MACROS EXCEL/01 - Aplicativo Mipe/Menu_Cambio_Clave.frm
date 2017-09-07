VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Cambio_Clave 
   Caption         =   "::::::::Cambiar Password"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   OleObjectBlob   =   "Menu_Cambio_Clave.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Menu_Cambio_Clave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cambiar_clave_Click()

If txt_acceso_password <> "" Then

   
    Call conectarBD

      
       
        ssql = " UPDATE TBL_passwd " _
        & " SET passwd = '" & txt_acceso_password & "'" _
        & " where rut_ejecutivo = '" & txt_acceso_rut_ejecutivo & "'"
        
                             
        Dim irespuesta As Integer
        irespuesta = MsgBox("¿Esta Seguro Que Desea Cambiar Su Contraseña?", vbYesNo)
        
        If irespuesta = vbYes Then
        
            Set rst = cnn.Execute(ssql, , adCmdText)
                    
            MsgBox "La Contraseña ha sido cambiada Satisfatoriamente al rut"
            
            Menu_Cambio_Clave.Hide
            Menu_Principal_Micro.Show
        
            Else
        
            txt_acceso_rut_ejecutivo = Empty
            txt_dv = Empty
            txt_acceso_password = Empty

        
        End If

Else
        MsgBox ("La Contraseña Debe ser Diferente a Blanco")

End If
End Sub


Private Sub cmd_volver_acceso_principal_Click()

Unload Menu_Cambio_Clave
Acceso_Principal.txt_acceso_password = Empty
Acceso_Principal.Show

End Sub

Private Sub Image1_Click()

End Sub

Private Sub txt_acceso_password_Change()

Dim Largo_Passwd As Integer
    Largo_Passwd = Len(txt_acceso_password)
     
     If Largo_Passwd > 10 Then
        MsgBox ("La Clave Tiene Mas Caracter De Lo Permitido")
        txt_acceso_password = Empty
     
     Else
        txt_acceso_rut_ejecutivo = rut_cliente_var
        txt_dv = dv_cliente_var

        Dim I As Integer
        txt_acceso_password = UCase(txt_acceso_password)
        I = Len(txt_acceso_password)
        txt_acceso_password.SelStart = I
    End If

End Sub

