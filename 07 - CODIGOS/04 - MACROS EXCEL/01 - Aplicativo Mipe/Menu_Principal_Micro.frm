VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Principal_Micro 
   Caption         =   "::: Menu Principal Sistema Microempresas"
   ClientHeight    =   9105.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615.001
   OleObjectBlob   =   "Menu_Principal_Micro.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "Menu_Principal_Micro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_evaluacion_Click()
Unload Menu_Principal_Micro
Evaluacion_Perfil.Show
End Sub

Private Sub cmd_cambio_password_Click()

Menu_Cambio_Clave.txt_acceso_rut_ejecutivo = Empty
Menu_Cambio_Clave.txt_dv = Empty
Menu_Cambio_Clave.txt_acceso_password = Empty

Menu_Principal_Micro.Hide
Menu_Cambio_Clave.Show
End Sub

Private Sub cmd_ficha_Click()

Call conectarBD
   ''''DISPONIBILIDAD DEL SISTEMA
    
        ssql = "select Marca_ok  from TBL_MICRO_MARCA_DISPONIBLE"
        Set rst = cnn.Execute(ssql, , adCmdText)
    
If rst!MARCA_OK = 0 Then


Menu_Principal_Micro.Hide
Ficha_Cliente_Micro.Show
txt_rut_cliente.Enabled = True
txt_dv.Enabled = True


Else
MsgBox "EL SISTEMA ESTA EN PROCESO DE ACTUALIZACION DE DATOS INTENTELO EN UNOS MOMENTOS MAS", vbCritical

End If ''''' PRIMER IF PARA EL INGRESO AL SISTEMA


End Sub

Private Sub Cerrar_Aplicacion_Click()
    'ActiveWorkbook.Save
    ActiveWorkbook.Close ""
    Excel.Application.Quit
End Sub

Private Sub cmd_ingreso_sinacofi_Click()

Menu_Principal_Micro.Hide

    Datos_Sinacofi.cbx_mora.Clear
    Datos_Sinacofi.cbx_protesto.Clear
    Datos_Sinacofi.cbx_boletin.Clear

    Datos_Sinacofi.cbx_mora.AddItem "Cumple"
    Datos_Sinacofi.cbx_mora.AddItem "No Cumple"

    Datos_Sinacofi.cbx_protesto.AddItem "Cumple"
    Datos_Sinacofi.cbx_protesto.AddItem "No Cumple"

    Datos_Sinacofi.cbx_boletin.AddItem "Cumple"
    Datos_Sinacofi.cbx_boletin.AddItem "No Cumple"

    Datos_Sinacofi.Show

End Sub

Private Sub cmd_princ_captura_Click()
Menu_Principal_Micro.Hide
Captura_Datos_Clientes.Show
End Sub

Private Sub cmd_rechazos_sernac_Click()

Menu_Principal_Micro.Hide
Visualizador_Rechazos_Sernac.Show

End Sub

Private Sub cmd_reconectar_Click()
Acceso_Principal.txt_acceso_rut_ejecutivo = Empty
Acceso_Principal.txt_acceso_password = Empty
Acceso_Principal.txt_dv = Empty

Menu_Principal_Micro.Hide
Acceso_Principal.Show
Acceso_Principal.txt_acceso_rut_ejecutivo.SetFocus
End Sub

Private Sub cmd_ver_evaluacion_Click()
Menu_Principal_Micro.Hide
Visualizador_Inicial.Show
End Sub

Private Sub cmd_visualizador_micro_Click()
Menu_Principal_Micro.Hide
Visualizador_Cliente.Show
End Sub

Private Sub CommandButton1_Click()



End Sub

Private Sub Image1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
