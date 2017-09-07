VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Visualizador_Cliente 
   Caption         =   "::::::Visualizador Clientes"
   ClientHeight    =   8235.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   OleObjectBlob   =   "Visualizador_Cliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Visualizador_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub cmd_buscar_Click()
    
    If txt_dv_compara = txt_dv Then
    

    If txt_dv_compara = txt_dv And txt_rut_cliente <> "" Then
    

    Call conectarBD
    
   
    ssql = "SELECT RUT_NUM,DV,TIPO_CLIENTE,GIRO,NOMBRE_CLIENTE,ID_EJECUTIVO_ASIGNADO,EJECUTIVO_ASIGNADO," _
        & " COD_SUC,ZONA_SUCURSAL,NOMBRE_SUCURSAL,OFERTA_PREEVALUADA,SCORE " _
        & " FROM TBL_carga_campana_me " _
        & " where rut_num = '" & txt_rut_cliente & "'"
       
               
    Set rst = cnn.Execute(ssql, , adCmdText)
        
    Dim txt_rut_consul As Variant
      
        If rst.EOF Then
           
        MsgBox ("Rut No Esta en la Base De Campaña")
        
        Else
        
              txt_tipo_cliente = rst!tipo_cliente
              txt_giro = rst!giro
              txt_nombre_cliente = rst!nombre_cliente
              'txt_id_ejecutivo_asig = rst!id_ejecutivo_asignado
              txt_ejecutivo_asig = rst!Ejecutivo_Asignado
              txt_cod_sucursal = rst!cod_suc
              txt_nombre_zona = rst!zona_sucursal
              txt_nombre_sucursal = rst!Nombre_Sucursal
              txt_oferta = rst!oferta_preevaluada
              txt_score = rst!score
            
            End If
          
        
        End If

    Else
        MsgBox " Rut o Digito Verificador Incorrecto"
End If
End Sub



Private Sub cmd_salir_sistema_Click()
    Workbooks("Sistema_Evaluacion_Gestion_Micro.xls").Close
    Application.Quit
End Sub

Private Sub cmd_volver_menu_princ_Click()
Unload Visualizador_Cliente
Call Menu_Principal_Micro.Show
End Sub

Private Sub txt_dv_Change()
Dim I As Integer
txt_dv = UCase(txt_dv)
I = Len(txt_dv)
txt_dv.SelStart = I
End Sub

Private Sub txt_oferta_Change()
txt_oferta = Format(txt_oferta, "##,##")
End Sub

Private Sub txt_rut_cliente_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
   
    Dim diga As Variant
    
    If Not IsNumeric(txt_rut_cliente) Then
        diga = MsgBox("El Rut Debe Ser Numérico. Favor Ingrese Solo Números", vbOKOnly)
        txt_rut_cliente = Empty
      
      End If
   
' ********** CALCULO DE DIGITO VERIFICADO *************
    Dim Vari1, Vari2, Vari3, I As Integer
    txt_rut_cliente = Replace(txt_rut_cliente, "-", "")
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
