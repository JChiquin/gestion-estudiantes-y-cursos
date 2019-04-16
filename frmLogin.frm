VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio sesión "
   ClientHeight    =   2145
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4155
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1267.337
   ScaleMode       =   0  'User
   ScaleWidth      =   3901.319
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ConexionOracle"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ConexionOracle"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      ItemData        =   "frmLogin.frx":058A
      Left            =   1290
      List            =   "frmLogin.frx":058C
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton cmdVer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   195
      Left            =   3240
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Mostrar contraseña"
      Top             =   799
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   1
      Top             =   375
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000014&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H80000009&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2340
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1500
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2325
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Erróneos. Por favor intente de nuevo."
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   390
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   780
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   4443
      Left            =   0
      Picture         =   "frmLogin.frx":058E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4521
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    'comprobar si el usuario es correcto
    Set rs = New ADODB.Recordset
    sql = "select uusuario from tusuarios where uusuario= '" & txtUserName.Text & "' and uestatus='A'"
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        If (txtUserName.Text = "") Then
            lblError.Caption = "Ingrese un usuario"
            lblError.Visible = True
            txtUserName.SetFocus
            Exit Sub
        End If
    lblError.Caption = "No existe ese usuario"
    lblError.Visible = True
    txtUserName.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    'fin comprobar nombre
    
    'comprobar contraseña
    Set rs = New ADODB.Recordset
    sql = "select upassword from tusuarios where upassword='" & LCase(txtPassword.Text) & "' and uusuario= '" & LCase(txtUserName.Text) & "' and uestatus='A'"
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        If (txtPassword.Text = "") Then
            lblError.Caption = "Ingrese una contraseña"
            lblError.Visible = True
            txtPassword.SetFocus
            Exit Sub
         End If
    lblError.Caption = "Contraseña incorrecta"
    lblError.Visible = True
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    End If
    'fin comprobar contraseña
    
    'recordar usuarios
    Set rs = New ADODB.Recordset
    sql = "select *from trecordarusuario where usuarios= '" & LCase(txtUserName.Text) & "' "
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        sql = "Insert Into TRecordarUsuario values ('" & LCase(txtUserName.Text) & "' )"
        db.Execute sql, SOpt
    End If
    'fin recordar usuarios
    
    Set rs = New ADODB.Recordset
    sql = "select * from tusuarios where upassword='" & LCase(txtPassword.Text) & "' and uusuario= '" & LCase(txtUserName.Text) & "' and uestatus='A'"
    rs.Open sql, db, adOpenStatic
    
    With frmPrincipal 'Para no escribir siempre "frmPrincipal." ahora basta con solo poner el punto
        Select Case rs!utipousuario
            Case 1
                .Show
                .mReportes.Visible = False
                .mUsuarios.Visible = False
                .cmdReporteCursos.Visible = False
                .cmdReporteGrupos.Visible = False
                .cmdUsuarios.Visible = False
                .imgGruposCulminar.Visible = False
                .imgAcento.Visible = False
                .DataGrid2.Visible = False
            Case 2
                .Show
            Case 3
                .Show
                .mUsuarios.Visible = False
                .mMenu.Visible = False
                .cmdCursos.Visible = False
                .cmdGrupos.Visible = False
                .cmdInstructores.Visible = False
                .cmdUsuarios.Visible = False
        End Select
    End With
    
    Me.Hide
    
    Dim apellidos() As String, nombres() As String 'apellidos y nombres son vectores
    nombres = Split(rs!unombres) 'antes de split: "Juan Carlos Luis" split me retorna un vector con un nombre en cada casilla:  {"Juan","Carlos","Luis"}
    apellidos = Split(rs!uapellidos)
    nombreUser = nombres(0) 'tomo únicamente el primer nombre
    apellidoUser = apellidos(0) 'tomo únicamente el primer apellido.
    
    txtPassword.Text = ""
    txtUserName.Text = ""
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub cmdVer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPassword.PasswordChar = ""
txtPassword.FontSize = 8

End Sub

Private Sub cmdVer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPassword.PasswordChar = "*"
txtPassword.FontSize = 12
End Sub

Private Sub Command1_Click()
frmPrincipal.Show
Me.Hide
End Sub

Private Sub Form_Load()
Call Conexion
txtPassword.FontSize = 12
CulminarGrupos 'se llama el procedimiento que culmina grupos
'Al ser este el primer formulario en abrirse con el programa, cada vez que abran el programa se actualizarán los grupos.


End Sub

Private Sub Form_Unload(Cancel As Integer)
db.Close
End
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 32 Or KeyCode = 13) Then
txtUserName.Text = List1.Text
SendKeys "{End}"
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtUserName.Text = List1.Text
List1.Visible = False
End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Now
'frmLogin.Caption = "Inicio de Sesión " & Now

End Sub

Private Sub txtPassword_Change()
lblError.Visible = False
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
KeyAscii = 0
End If
End Sub

Private Sub txtUserName_Change()
lblError.Visible = False
'Generar lista de usuarios recorados
    Set rs = New ADODB.Recordset
    sql = "select *from trecordarusuario where usuarios like '" & LCase(txtUserName.Text) & "%' "
    rs.Open sql, db, adOpenStatic
    If Not rs.EOF Then
        List1.Clear
        Dim X As Integer
        X = 0
        Do
            List1.List(X) = rs!usuarios
            rs.MoveNext
            X = X + 1
        Loop Until rs.EOF
        List1.Height = List1.ListCount * 150.662
        List1.Visible = True
        If (List1.List(0) = txtUserName.Text) Then
            List1.Visible = False
        End If
    Else
        List1.Visible = False
    End If
    
    If (txtUserName.Text = "") Then
        List1.Visible = False
        List1.Clear
        List1.Height = 150.662
    End If
'fin generar lista usuarios recordados

    If List1.Visible = False Then
        cmdOK.Default = True
    ElseIf (List1.Visible = True) Then
        cmdOK.Default = False
    End If
    
    
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 40 And List1.Visible) Then
List1.SetFocus
SendKeys "{DOWN}"
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
KeyAscii = 0
End If
End Sub
