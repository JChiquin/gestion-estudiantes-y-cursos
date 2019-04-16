VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmUsuario 
   Caption         =   "Usuarios"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10755
   Icon            =   "frmUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtUsuario 
      Height          =   375
      Left            =   5520
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox TxtContraseña 
      Height          =   375
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   16
      Top             =   3360
      Width           =   3000
   End
   Begin VB.TextBox TxtConfirmar 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtNueva 
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   1920
      Width           =   3000
   End
   Begin VB.TextBox txtapellido 
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   2640
      Width           =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11040
      Top             =   1800
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   4200
      Picture         =   "frmUsuario.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1700
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   6120
      Picture         =   "frmUsuario.frx":0D9C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1700
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   8040
      Picture         =   "frmUsuario.frx":2AB2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1700
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   9960
      Picture         =   "frmUsuario.frx":4888
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1700
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   6120
      Picture         =   "frmUsuario.frx":7439
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1700
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   8040
      Picture         =   "frmUsuario.frx":976B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1700
   End
   Begin VB.Frame Frameusuario 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1320
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
      Begin VB.OptionButton Opt3 
         BackColor       =   &H8000000D&
         Caption         =   "Usuario Tipo 3"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H8000000D&
         Caption         =   "Usuario Tipo 2"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H8000000D&
         Caption         =   "Usuario Tipo 1"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   630
      Left            =   1920
      TabIndex        =   29
      Top             =   480
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1111
      Image           =   "frmUsuario.frx":B074
      Props           =   5
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   28
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label LblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Usuario "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   27
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label LblContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   26
      Top             =   3360
      Width           =   1605
   End
   Begin VB.Label LblConfirmar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Contraseña :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label lblError1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Mínimo 4 Digitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblError3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Seleccionar un Tipo de usario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   8640
      Picture         =   "frmUsuario.frx":B858
      Stretch         =   -1  'True
      ToolTipText     =   "Contraseñas coinciden"
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   8640
      Picture         =   "frmUsuario.frx":BE24
      Stretch         =   -1  'True
      ToolTipText     =   "Contraseña no coinciden"
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNueva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Nueva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   21
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   20
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblErrorNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* debe colocar un nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblErrorApellido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* debe colocar un apellido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image Image4 
      Height          =   1815
      Left            =   11880
      Picture         =   "frmUsuario.frx":C2FC
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   13080
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   12960
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   9000
      Left            =   0
      Picture         =   "frmUsuario.frx":F97C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()

        'buscar un usuario
        Set rs = New ADODB.Recordset
        sql = " select * from tusuarios where uusuario = '" & LCase(TxtUsuario.Text) & "' and uestatus= 'A'  "
        rs.Open sql, db, adOpenStatic

    If (rs.EOF) Then
      Set rs = New ADODB.Recordset
        sql = " select * from tusuarios where uusuario = '" & LCase(TxtUsuario.Text) & "' and uestatus= 'E'  "
        rs.Open sql, db, adOpenStatic

    If (rs.EOF) Then
      resp = MsgBox("El Usuario no esta registrado. Desea Incluirlo?", vbInformation + vbYesNo, "Inclusion")
    If (resp = vbYes) Then
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = True
    TxtUsuario.Enabled = False
    TxtContraseña.Enabled = True
    TxtConfirmar.Visible = True
    txtnombre.Enabled = True
    txtapellido.Enabled = True
    LblConfirmar.Visible = True
    Opt1.Enabled = True
    Opt2.Enabled = True
    Opt3.Enabled = True
    txtnombre.SetFocus
    TxtContraseña.PasswordChar = "*"
    End If
    Else
    TxtUsuario.Text = rs!uusuario
    TxtContraseña.Text = rs!upassword
    txtnombre.Text = rs!unombres
    txtapellido.Text = rs!uapellidos
'Activar las Opciones

    Select Case rs!utipousuario
        Case 1
        Opt1.Value = True
        Case 2
        Opt2.Value = True
        Case 3
        Opt3.Value = True
    End Select
resp = MsgBox("Usuario inactivo. ¿Desea activarlo?", vbYesNo + vbQuestion, "Usuario")
                If (resp = vbYes) Then
                    sql = "update  tusuarios set uestatus = 'A' where uusuario = '" & LCase(TxtUsuario.Text) & "'"
                    db.Execute sql, SOpt
                    MsgBox "Activacion Exitosa", vbInformation, "Usuario"
                    cmdBuscar.Enabled = False
                    cmdModificar.Enabled = True
                    cmdEliminar.Enabled = True
                    TxtUsuario.Enabled = False
                Else
                    CmdCancelar_Click
                End If
End If
Else

    cmdIncluir.Enabled = False
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdCancelar.Enabled = True
    cmdBuscar.Enabled = False
    TxtUsuario.Text = rs!uusuario
    TxtContraseña.Text = rs!upassword
    txtnombre.Text = rs!unombres
    txtapellido.Text = rs!uapellidos
'Activar las Opciones

    Select Case rs!utipousuario
        Case 1
        Opt1.Value = True
        Case 2
        Opt2.Value = True
        Case 3
        Opt3.Value = True
    End Select

    TxtUsuario.Enabled = False
    TxtContraseña.Enabled = False
    Opt1.Enabled = False
    Opt2.Enabled = False
    Opt3.Enabled = False

End If
'fin de buscar usurio
End Sub

Private Sub CmdCancelar_Click()
    
    TxtUsuario.Text = ""
    TxtContraseña.Text = ""
    TxtConfirmar.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    Opt1.Value = False
    Opt2.Value = False
    Opt3.Value = False
    Opt1.Enabled = False
    Opt2.Enabled = False
    Opt3.Enabled = False
    TxtUsuario.Enabled = True
    TxtContraseña.Enabled = False
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    TxtUsuario.SetFocus
    TxtConfirmar.Visible = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    LblConfirmar.Visible = False
    Image1.Visible = False
    TxtContraseña.PasswordChar = ""
    lblError1.Visible = False
    lblError3.Visible = False
    lblErrorNombre.Visible = False
    lblErrorApellido.Visible = False
    TxtConfirmar.Enabled = False
    Image2.Visible = False
    lblNueva.Visible = False
    txtNueva.Visible = False
    LblContraseña.Caption = "Contraseña"
    LblConfirmar.Caption = "Confirmar Contraseña"
    cmdModificar.Caption = "Modificar"
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
End Sub

Private Sub CmdEliminar_Click()

        'eliminacion logica de usuarios
         resp = MsgBox("¿Está seguro que desea eliminar el instructor?", vbYesNo + vbQuestion, "Eliminar Usuario")
    If (resp = vbYes) Then
        sql = "update tusuarios set uestatus = 'E' where uusuario = '" & TxtUsuario.Text & "' "
        db.Execute sql, SOpt
        CmdCancelar_Click
        MsgBox "Eliminacion Exitosa", vbInformation, "Eliminar Usuario"
    End If


End Sub

Private Sub CmdIncluir_Click()
        'incluir nuevo usuario
        'validacion de campos vacios y errores
        Dim tipo As String
        
            If (Opt1.Value = True) Then
                tipo = "1"
            ElseIf (Opt2.Value = True) Then
                tipo = "2"
            Else
                tipo = "3"
            End If
            
        If (txtnombre.Text = "") Then
            lblErrorNombre.Visible = True
        End If

        If (txtapellido.Text = "") Then
            lblErrorApellido.Visible = True
        End If
        
        If (Len(TxtContraseña.Text) < 4) Then
            lblError1.Visible = True
            Else
            lblError1.Visible = False
        End If
            
            
        If (Opt1.Value = False And Opt2.Value = False And Opt3.Value = False) Then
            lblError3.Visible = True
            Else
            lblError3.Visible = False
        End If
    
    If (lblErrorNombre.Visible) Then
        txtnombre.SetFocus
        Exit Sub
    End If
    
    If (lblErrorApellido.Visible) Then
        txtapellido.SetFocus
        Exit Sub
    End If
    
    If lblError1.Visible = True Then
    TxtContraseña.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    
       
    If Image1.Visible = False Then
    TxtConfirmar.SetFocus
    Image2.Visible = True
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    'fin de validacion de campos vacios y errores
    
    If Not (lblErrorNombre.Visible Or lblErrorApellido.Visible Or lblError1.Visible Or Image1.Visible = False Or lblError3.Visible) Then  'Si no hay un label de error visible, es porque está listo para incluir.
    sql = "Insert Into TUsuarios values ('" & LCase(TxtUsuario.Text) & "' , '" & (TxtContraseña.Text) & "' , '" & (tipo) & "','A','" & LCase(txtnombre.Text) & "','" & LCase(txtapellido.Text) & "')"
    db.Execute sql, SOpt
    rs.Close
    
    MsgBox "Usuario: " + TxtUsuario.Text + " esta Registrado", vbInformation, "Usuario"
    CmdCancelar_Click
    ' fin de inlucion de usuario
    End If
End Sub

Private Sub CmdModificar_Click()
    'inicio de modificar usuario
    If (cmdModificar.Caption = "Modificar") Then
        
        txtnombre.Enabled = True
        txtapellido.Enabled = True
        Opt1.Enabled = True
        Opt2.Enabled = True
        Opt3.Enabled = True
        TxtConfirmar.Visible = True
        LblConfirmar.Visible = True
        cmdModificar.Caption = "Guardar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Guardar.jpg")
        LblContraseña.Caption = "Contraseña Actual"
        LblConfirmar.Caption = "Confirmar Contraseña Nueva"
        lblNueva.Visible = True
        txtNueva.Visible = True
        Exit Sub
    End If
    'validacion de campos vacios y errores
    
     If (txtnombre.Text = "") Then
            lblErrorNombre.Visible = True
        End If

        If (txtapellido.Text = "") Then
            lblErrorApellido.Visible = True
        End If
        
        If (Len(TxtContraseña.Text) < 4) Then
            lblError1.Visible = True
            Else
            lblError1.Visible = False
        End If
        
    Dim tipo As String
        
            If (Opt1.Value = True) Then
                tipo = "1"
             ElseIf (Opt2.Value = True) Then
                tipo = "2"
            Else
                tipo = "3"
            End If
            
    If (Len(txtNueva.Text) < 4) Then
            lblError1.Visible = True
            Else
            lblError1.Visible = False
        End If
            
            
        If (Opt1.Value = False And Opt2.Value = False And Opt3.Value = False) Then
            lblError3.Visible = True
            Else
            lblError3.Visible = False
        End If
     
    If (lblErrorNombre.Visible) Then
        txtnombre.SetFocus
        Exit Sub
    End If
    
    If (lblErrorApellido.Visible) Then
        txtapellido.SetFocus
        Exit Sub
    End If
    
    If lblError1.Visible = True Then
    txtNueva.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    
       
    If Image1.Visible = False And txtNueva.Visible = True Then
    TxtConfirmar.SetFocus
    Image2.Visible = True
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    'fin validacion de campos vacios y errores
    If Not (lblErrorNombre.Visible Or lblErrorApellido.Visible Or lblError1.Visible Or Image1.Visible = False Or lblError3.Visible) Then  'Si no hay un label de error visible, es porque está listo para incluir.
    'verifica si no hubo algun cambio
    Set rs = New ADODB.Recordset
    sql = "Select * from tusuarios where Uusuario = '" & TxtUsuario.Text & "' and Uestatus= ('A')"
    rs.Open sql, db, adOpenStatic
    
    If (rs!upassword = txtNueva.Text And rs!utipousuario = tipo And rs!unombres = txtnombre.Text And rs!uapellidos = txtapellido.Text) Then
    MsgBox "No hubieron cambios.", vbInformation, "Modificar Usuarios"
    cmdModificar.Caption = "Modificar"
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    
        txtnombre.Enabled = False
        txtapellido.Enabled = False
        Opt1.Enabled = False
        Opt2.Enabled = False
        Opt3.Enabled = False
        lblNueva.Visible = False
        LblConfirmar.Visible = False
        txtNueva.Visible = False
        txtNueva.Text = ""
        TxtConfirmar.Visible = False
        Image1.Visible = False
    Exit Sub
    End If
    'modificacion con cambios
    resp = MsgBox("¿Estás seguro que quieres Modificar los Datos ?", vbYesNo + vbQuestion, "Modificar Usuario")
    If (resp = vbYes) Then
        sql = "update Tusuarios set upassword = '" & txtNueva.Text & "', utipousuario = '" & tipo & "', unombres = '" & txtnombre.Text & "', uapellidos = '" & txtapellido.Text & "' where uusuario='" & TxtUsuario.Text & "' and uestatus='A' "
        db.Execute sql, SOpt
        MsgBox "Usuario modificado Exitosamente!", vbInformation, "Modificar Usuario"
        cmdModificar.Caption = "Modificar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
        txtnombre.Enabled = False
        txtapellido.Enabled = False
        lblNueva.Visible = False
        LblConfirmar.Visible = False
        txtNueva.Visible = False
        TxtContraseña.Text = txtNueva.Text
        txtNueva.Text = ""
        TxtConfirmar.Visible = False
        Image1.Visible = False
        Opt1.Enabled = False
        Opt2.Enabled = False
        Opt3.Enabled = False
        
    End If
    End If
    'fin de modificar usuarios
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    

End Sub

Private Sub Form_Activate()
'Para que la imagen de fondo se autoajuste al tamaño del form MAXIMIZADO
If (Image3.Height < (Me.Height - 300)) Then
    Image3.Height = Me.Height - 300
End If
If (Image3.Width < (Me.Width - 100)) Then
    Image3.Width = Me.Width - 100
End If
'fin autoajustar
End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub
Private Sub Form_Load()

    TxtContraseña.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    Opt1.Enabled = False
    Opt2.Enabled = False
    Opt3.Enabled = False
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    cmdModificar.Caption = "Modificar"


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPrincipal.Show
End Sub

Private Sub Opt1_Click()
If Opt1.Value = True Then
lblError3.Visible = False
End If
End Sub

Private Sub Opt2_Click()
If Opt2.Value = True Then
lblError3.Visible = False
End If
End Sub

Private Sub Opt3_Click()
If Opt3.Value = True Then
lblError3.Visible = False
End If
End Sub

Private Sub txtapellido_Change()
If (txtapellido.Text <> "") Then
lblErrorApellido.Visible = False
End If
End Sub

Private Sub txtapellido_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtConfirmar_Change()
    
     If (txtNueva.Visible = True) Then                 ' confirmar la contrasena nueva
         If (txtNueva.Text <> TxtConfirmar.Text) Then  ' es correcta
            Image1.Visible = False                     '
            Image2.Visible = False
         Else
            Image1.Visible = True
            Image2.Visible = False
         End If
     End If
     
     If (txtNueva.Visible = False) Then                         'confirmar la contrasena
         If (TxtContraseña.Text <> TxtConfirmar.Text) Then      'del usuario nuevo al incluirlo
            Image1.Visible = False
            Image2.Visible = False
         Else
            Image1.Visible = True
            Image2.Visible = False
         End If
     End If
        If TxtConfirmar.PasswordChar = "*" Then
        TxtConfirmar.FontSize = 12
        Else
        TxtConfirmar.FontSize = 8
        End If
End Sub

Private Sub TxtConfirmar_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32) Then
KeyAscii = 0
End If
End Sub

Private Sub TxtContraseña_Change()
If (TxtContraseña.Enabled = True And Len(TxtContraseña.Text) >= 4) Then
lblError1.Visible = False                                              'para habilitar la caja ta texto de confirmar contrasena
TxtConfirmar.Enabled = True
Else

TxtConfirmar.Enabled = False
TxtConfirmar.Text = ""
Image2.Visible = False
End If

If TxtContraseña.PasswordChar = "*" Then  'muestra la contrasena en **
TxtContraseña.FontSize = 12
Else
TxtContraseña.FontSize = 8
End If

If (txtNueva.Visible = False) Then
         If (TxtContraseña.Text <> TxtConfirmar.Text Or txtNueva.Text = "") Then ' ocultar o mostrar imagen de confirmar contrasena
            Image1.Visible = False
            Image2.Visible = False
         Else
            Image1.Visible = True
            Image2.Visible = False
         End If
     End If
End Sub

Private Sub TxtContraseña_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
KeyAscii = 0
End If
End Sub

Private Sub txtNombre_Change()
If (txtnombre.Text <> "") Then
lblErrorNombre.Visible = False
End If
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNueva_Change()
If (txtNueva.Enabled = True And Len(txtNueva.Text) >= 4) Then
lblError1.Visible = False                                     'para habilitar la caja ta texto de confirmar contrasena
TxtConfirmar.Enabled = True
Else

TxtConfirmar.Enabled = False
TxtConfirmar.Text = ""
Image2.Visible = False
End If

If (txtNueva.Visible = True) Then
         If (txtNueva.Text <> TxtConfirmar.Text Or txtNueva.Text = "") Then  ' ocultar o mostrar imagen de confirmar contrasena
            Image1.Visible = False
            Image2.Visible = False
         Else
            Image1.Visible = True
            Image2.Visible = False
         End If
     End If
End Sub

Private Sub txtNueva_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
KeyAscii = 0
End If
End Sub

Private Sub TxtUsuario_Change()

        If (TxtUsuario.Text <> "") Then  'habilita el buscar y el cancelar si txtUsuario es distinto de vacio
            cmdBuscar.Enabled = True
            cmdCancelar.Enabled = True
        Else
            cmdBuscar.Enabled = False
            cmdCancelar.Enabled = False
            End If
            
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 And cmdBuscar.Enabled = True) Then
CmdBuscar_Click
End If
If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
KeyAscii = 0
End If
End Sub
