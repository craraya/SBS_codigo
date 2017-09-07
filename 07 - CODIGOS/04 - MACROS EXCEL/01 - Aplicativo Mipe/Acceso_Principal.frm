VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Acceso_Principal 
   Caption         =   ":::: Menu Acceso"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595.001
   OleObjectBlob   =   "Acceso_Principal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Acceso_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_volver_menu_principal_Click()

    'ActiveWorkbook.Save
    ActiveWorkbook.Close ""
    Excel.Application.Quit
End Sub



Private Sub txt_acceso_rut_ejecutivo_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   
    Dim diga As Variant
    
    If Not IsNumeric(txt_acceso_rut_ejecutivo) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_acceso_rut_ejecutivo = Empty
      
      End If
      
   
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_acceso_rut_ejecutivo = Replace(txt_acceso_rut_ejecutivo, "-", "")
    txt_acceso_rut_ejecutivo = Replace(txt_acceso_rut_ejecutivo, ".", "")
    Vari3 = 2
    For I = 0 To Len(txt_acceso_rut_ejecutivo) - 1
     If Left(Right(txt_acceso_rut_ejecutivo, I + 1), 1) <> "." Then
      Vari1 = Vari1 + Left(Right(txt_acceso_rut_ejecutivo, I + 1), 1) * Vari3
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

Private Sub cmd_ingresar_sistema_Click()

    '''''Llamado a la CNX_BD
    Call conectarBD
    
    
    
    hora1 = hora
    txt_hora_actual = Time

    Dim fec1
    'Dim hora1

    fec1 = Format(Date, "yyyy/mm/dd")
    txt_fecha_actual = fec1

    'hora1 = hora
    'txt_hora_actual = Time
    
    Dim nombre_completo As String
    
    Dim txt_clave_compara As Variant
    Dim cargo_compara As Variant

    rut_cliente_var = txt_acceso_rut_ejecutivo
    dv_cliente_var = txt_dv
 
    If txt_dv = txt_dv_compara Then


        
    ''''' Chequea Version Aplicativo ''''''
        ssql = "select max(Numero_Version) AS Numero_Version from tbl_version_apl_MICRO"
        Set rst = cnn.Execute(ssql, , adCmdText)
      
    '''' Fin Chequeo
   
    If txt_version * 1 = rst!numero_version Then
        
        ssql = "SELECT a.rut_ejecutivo, a.passwd, b.cargo_ejecutivo, b.Nombre_Ejecutivo+' '+b.Apellido_Ejecutivo as Nombre_Eje, b.aut_excepcion_micro, monto_aut_evaluador,estado_ejecutivo" _
        & " FROM TBL_passwd a, tbl_ejecutivo b " _
        & " where a.rut_ejecutivo = b.rut_ejecutivo" _
        & " and a.rut_ejecutivo = '" & txt_acceso_rut_ejecutivo & "'" _
        & " group by a.rut_ejecutivo, a.passwd,b.cargo_ejecutivo, b.nombre_ejecutivo, b.apellido_ejecutivo,b.aut_excepcion_micro,monto_aut_evaluador,estado_ejecutivo"
        
        Set rst = cnn.Execute(ssql, , adCmdText)
      
        If rst.EOF Then
               MsgBox ("Ejecutivo No Registrado En El Mantenedor... Solicitar Ingreso A Su Agente")
               txt_acceso_rut_ejecutivo = Empty
               txt_dv = Empty
               txt_acceso_password = Empty
               
        ElseIf txt_acceso_password <> rst!passwd Then
               MsgBox ("Clave Erronea ... Reintente")
               txt_acceso_password = Empty
               txt_acceso_password.SetFocus
               
        ElseIf rst!estado_ejecutivo = "NO ACTIVO" Then
               MsgBox ("Usuario NO PUEDE INGRESAR AL SISTEMA en Estado Inactivo... Avisar a su Agente")
               txt_acceso_password = Empty
               txt_acceso_password.SetFocus
               
        ElseIf rst!cargo_ejecutivo = "EJECUTIVO MICROEMPRESA" And rst!aut_excepcion_micro = 0 Then
        
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
            
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
            
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.Show

                
                
                
        ElseIf rst!cargo_ejecutivo = "EJECUTIVO MICROEMPRESA" And rst!aut_excepcion_micro = 1 Then
        
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
            
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
            
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Ficha_Cliente_Micro.lbl_tipo_excepcion.Visible = True
                Ficha_Cliente_Micro.lbl_ejecutivo_evaluador.Visible = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                
                Ficha_Cliente_Micro.cbx_tipo_excepcion.Visible = True
                
                Menu_Principal_Micro.Show
                
               
        ElseIf rst!cargo_ejecutivo = "EVALUADOR MICROEMPRESA" And rst!aut_excepcion_micro = 1 Then
        
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
            
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                
                Ficha_Cliente_Micro.lbl_tipo_excepcion.Visible = True
                Ficha_Cliente_Micro.lbl_ejecutivo_evaluador.Visible = True
                
                Ficha_Cliente_Micro.cbx_tipo_excepcion.Visible = True
                'Ficha_Cliente_Micro.cbx_ejecutivo_excepcion.Visible = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                
                Menu_Principal_Micro.Show
                

                
                
        ElseIf rst!cargo_ejecutivo = "EVALUADOR MICROEMPRESA" And rst!aut_excepcion_micro = 0 Then
        
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
            
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.Show

               
        ElseIf rst!cargo_ejecutivo = "SIC" Then
                
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
                
                 Acceso_Principal.Hide
        
                Menu_Principal_Micro.cmd_ficha.Enabled = False
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = False
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = True
                Menu_Principal_Micro.Show
                
        ElseIf rst!cargo_ejecutivo = "SIC_ADJ_MICRO" Then
                
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
                
                 Acceso_Principal.Hide
        
                Menu_Principal_Micro.cmd_ficha.Enabled = False
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = False
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = True
                Menu_Principal_Micro.Show
                
                
        ElseIf rst!cargo_ejecutivo = "RIESGO" Then

                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
        
                 Acceso_Principal.Hide
        
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = False
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.Show
   
        ElseIf rst!cargo_ejecutivo = "AGENTE SUCURSAL" Then
            
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador

                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                'Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = True 'Modificacion 2015-06-30, Solicitada por R. Catalan. Prgm: Jose Pardo
                Menu_Principal_Micro.Show
                
        ElseIf rst!cargo_ejecutivo = "AGENTE SUCURSAL ESP" Then
            
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador

                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = True
                Menu_Principal_Micro.Show
   
                       
                ElseIf rst!cargo_ejecutivo = "ADMINISTRADOR" Then
                
                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador

                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = True
                
                Ficha_Cliente_Micro.lbl_tipo_excepcion.Visible = True
                Ficha_Cliente_Micro.lbl_ejecutivo_evaluador.Visible = True
                
                Ficha_Cliente_Micro.cbx_tipo_excepcion.Visible = True
                Ficha_Cliente_Micro.cbx_ejecutivo_excepcion.Visible = True
                
                Menu_Principal_Micro.Show
   
        ElseIf rst!cargo_ejecutivo = "ZONAL" Then

                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                Menu_Principal_Micro.cmd_ficha.Enabled = False
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.Show
                
                
        ElseIf rst!cargo_ejecutivo = "EJECUTIVO TLMK" Then

                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                Menu_Principal_Micro.cmd_ficha.Enabled = False
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = False
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                
                Menu_Principal_Micro.Show
   
   
           ElseIf rst!cargo_ejecutivo = "AUDITORIA" Then

                Menu_Principal_Micro.txt_nombre_ejecutivo_presentacion = rst!nombre_eje
                Menu_Principal_Micro.txt_rut_ejecutivo_presentacion = rst!rut_ejecutivo
                Menu_Principal_Micro.txt_monto_aut_micro = rst!monto_aut_evaluador
                
                ssql = "UPDATE TBL_PASSWD set Fecha_Ultima_Cnx = ' " & txt_fecha_actual & " ', hora_ultima_cnx = ' " & txt_hora_actual & " '" _
                & " where rut_ejecutivo      = '" & txt_acceso_rut_ejecutivo & "'"
            
            
                cnn.Execute ssql
        
                Acceso_Principal.Hide
                
                Menu_Principal_Micro.cmd_ficha.Enabled = True
                Menu_Principal_Micro.cmd_ver_evaluacion.Enabled = True
                Menu_Principal_Micro.cmd_visualizador_micro.Enabled = True
                Menu_Principal_Micro.cmd_princ_captura.Enabled = True
                Menu_Principal_Micro.cmd_ingreso_sinacofi.Enabled = False
                Menu_Principal_Micro.Show
                
                Menu_Principal_Micro.Show
   
               
        ElseIf rst!cargo_ejecutivo = "EJECUTIVO CONSUMER" Then
                
                cnn.Execute ssql
        
               MsgBox ("Este Tipo De Ejecutivo No Tiene Acceso A Esta Aplicacion")
       
                
       
        ElseIf rst!cargo_ejecutivo = "EJECUTIVO COMERCIAL TERRENO" Then
   
                cnn.Execute ssql
                
                MsgBox ("Este Tipo De Ejecutivo No Tiene Acceso A Esta Aplicacion")
       
               
               rst.MoveNext

        End If
        
    Else
   MsgBox "Se han Incorporado Nuevos Cambios En El Aplicativo, Favor bajar Versión Actualizada de Intranet"

    End If
 Else
    MsgBox ("Rut No Válido ... ")

   End If
   

End Sub

Private Sub txt_acceso_password_Change()
Dim I As Integer
txt_acceso_password = UCase(txt_acceso_password)
I = Len(txt_acceso_password)
txt_acceso_password.SelStart = I
End Sub

Private Sub txt_dv_Change()
Dim I As Integer
txt_dv = UCase(txt_dv)
I = Len(txt_dv)
txt_dv.SelStart = I
End Sub

