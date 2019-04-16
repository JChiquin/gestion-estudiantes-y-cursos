VERSION 5.00
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmCursos 
   BackColor       =   &H00004000&
   Caption         =   "Actualizacion de Cursos"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "frmCursos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6525
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   13
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   12
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtHoras 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   11
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtCant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   10
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox txtCosto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   9
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11160
      Top             =   1680
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H8000000A&
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
      Left            =   5640
      Picture         =   "frmCursos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H8000000A&
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
      Left            =   3600
      Picture         =   "frmCursos.frx":1E93
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H8000000A&
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
      Left            =   7680
      Picture         =   "frmCursos.frx":41C5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H8000000A&
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
      Left            =   5640
      Picture         =   "frmCursos.frx":6D76
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdIncluir 
      BackColor       =   &H8000000A&
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
      Left            =   3600
      Picture         =   "frmCursos.frx":8B4C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H8000000A&
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
      Left            =   1560
      Picture         =   "frmCursos.frx":A862
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1700
   End
   Begin VB.Label lblErrorCosto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Agregue un costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblErrorParticipantes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Debe estar entre 10 y 30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblErrorHoras 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Debe estar entre 20 y 120 horas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblErrorNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Indique Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   2160
      TabIndex        =   18
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   2880
      Width           =   1140
   End
   Begin VB.Label lblHoras 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horas/Curso"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   3480
      Width           =   1755
   End
   Begin VB.Label lblCant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Participantes"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   4080
      Width           =   1800
   End
   Begin VB.Label lblCosto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   4680
      Width           =   810
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   555
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   979
      Image           =   "frmCursos.frx":B074
      Props           =   5
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   12000
      Picture         =   "frmCursos.frx":B7B4
      Top             =   960
      Width           =   885
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
      Left            =   13080
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
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
      Left            =   13200
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   7995
      Left            =   0
      Picture         =   "frmCursos.frx":EE34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10515
   End
End
Attribute VB_Name = "frmCursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function FormatoDecimalPunto(costo As String) As String ' transforma "Bs. 12.340,701" en "12340.701" por ejemplo
    Dim variable As String
    variable = Mid(txtCosto.Text, 5) 'Para que quite "Bs. " del numero. "12.340,701"
    variable = CDbl(variable) 'Quita los puntos separadores de mil. "12340,701"
    variable = Replace(variable, ",", ".") 'Reemplaza la coma por un punto, esto es porque para sql el punto es el separador de decimal (como en las calculadoras) "12340.701"
    FormatoDecimalPunto = variable 'return variable

End Function

Private Sub CmdBuscar_Click()
    'Buscando un curso
    Set rs = New ADODB.Recordset
    sql = "Select * from TCursos where ccodigo = '" & txtcodigo.Text & "' and cestatus= ('A')"
    rs.Open sql, db, adOpenStatic
    
    If rs.EOF Then
        Set rs = New ADODB.Recordset
        sql = "Select * from TCursos where ccodigo= '" & txtcodigo.Text & "' and cestatus= ('E')"
        rs.Open sql, db, adOpenStatic
    
        If rs.EOF Then
            resp = MsgBox("No está registrado el curso. ¿Desea Incluirlo? ", vbYesNo + vbQuestion, "Cursos")
            If (resp = vbYes) Then
                txtcodigo.Enabled = False
                txtnombre.Enabled = True
                txtHoras.Enabled = True
                txtCant.Enabled = True
                txtCosto.Enabled = True
                txtnombre.SetFocus
                cmdBuscar.Enabled = False
                cmdIncluir.Enabled = True
                cmdCancelar.Enabled = True
            End If
        Else
            txtcodigo.Text = rs!ccodigo
            txtnombre.Text = rs!cnombre
            txtHoras.Text = rs!cduracionH
            txtCant.Text = rs!ccantidadP
            txtCosto.Text = Format(rs!ccosto, "\B\s\. #,##0.000") 'separador de miles y 3 decimales.
            
            resp = MsgBox("Curso inactivo. ¿Desea activarlo?", vbYesNo + vbQuestion, "Cursos")
            If (resp = vbYes) Then
                sql = "update TCursos set cestatus = 'A' where ccodigo = '" & txtcodigo.Text & "' "
                db.Execute sql, SOpt
                MsgBox "Activacion Exitosa", vbInformation, "Cursos"
                txtcodigo.Enabled = False
                cmdBuscar.Enabled = False
                cmdModificar.Enabled = True
                cmdEliminar.Enabled = True
                cmdCancelar.Enabled = True
            Else
                CmdCancelar_Click
            End If
        End If
    Else
        txtnombre.Text = rs!cnombre
        txtHoras.Text = rs!cduracionH
        txtCant.Text = rs!ccantidadP
        txtCosto.Text = Format(rs!ccosto, "\B\s\. #,##0.000") 'separador de miles y 3 decimales.
        txtcodigo.Enabled = False
        cmdBuscar.Enabled = False
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = True
    End If
    
    rs.Close
    
End Sub

Private Sub CmdCancelar_Click()


    txtnombre.Enabled = False
    txtHoras.Enabled = False
    txtCant.Enabled = False
    txtCosto.Enabled = False
    txtcodigo.Enabled = True
    txtcodigo.Text = ""
    txtnombre.Text = ""
    txtHoras.Text = ""
    txtCant.Text = ""
    txtCosto.Text = ""
    txtcodigo.SetFocus
    cmdCancelar.Enabled = False
    cmdIncluir.Enabled = False
    cmdEliminar.Enabled = False
    cmdModificar.Enabled = False
    cmdBuscar.Enabled = False
    cmdModificar.Caption = "Modificar"
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    
    lblErrorNombre.Visible = False
    lblErrorParticipantes.Visible = False
    lblErrorHoras.Visible = False
    lblErrorCosto.Visible = False
    

End Sub

Private Sub CmdEliminar_Click()

    resp = MsgBox("¿Está seguro que desea eliminar el Curso?", vbYesNo + vbQuestion, "Eliminar cursos")
    If (resp = vbYes) Then
    sql = "update TCursos set cestatus = 'E' where ccodigo = '" & txtcodigo.Text & "' "
    db.Execute sql, SOpt
    CmdCancelar_Click
    MsgBox "Eliminacion Exitosa", vbInformation, "Cursos"
    End If

End Sub

Private Sub CmdIncluir_Click()

    If (txtCosto.Text = "" Or Val(Mid(txtCosto.Text, 4)) <= 0) Then
        lblErrorCosto.Visible = True
    End If
    If (txtHoras.Text = "") Then
        lblErrorHoras.Caption = "Indique horas"
        lblErrorHoras.Visible = True
    End If
    If (txtCant.Text = "") Then
        lblErrorParticipantes.Caption = "Indique cantidad de participantes"
        lblErrorParticipantes.Visible = True
    End If
    If (txtnombre.Text = "") Then
        lblErrorNombre.Caption = "Indique nombre"
        lblErrorNombre.Visible = True
    End If
    
    
    If (Val(txtHoras.Text) < 20) Or (Val(txtHoras.Text) > 120) Then
        lblErrorHoras.Caption = "Debe estar entre 20 y 120."
        lblErrorHoras.Visible = True
    End If
    
    If (Val(txtCant.Text) < 10) Or (Val(txtCant.Text) > 30) Then
        lblErrorParticipantes.Caption = "Debe estar entre 10 y 30"
        lblErrorParticipantes.Visible = True
    End If
    
    If (lblErrorNombre.Visible) Then
        txtnombre.SetFocus
    Exit Sub
    End If

    If (lblErrorHoras.Visible) Then
        txtHoras.SetFocus
        Exit Sub
    End If
    
    If (lblErrorParticipantes.Visible) Then
        txtCant.SetFocus
        Exit Sub
    End If
    
    If (lblErrorCosto.Visible) Then
        txtCosto.SetFocus
        Exit Sub
    End If
    
    If Not (lblErrorNombre.Visible Or lblErrorHoras.Visible Or lblErrorParticipantes.Visible Or lblErrorCosto.Visible) Then
        sql = "Insert Into TCursos values" _
        & "('" & (txtcodigo.Text) & "', '" & (txtnombre.Text) & "', '" & (txtHoras.Text) & "', '" & (txtCant.Text) & "', '" & FormatoDecimalPunto(txtCosto.Text) & "', 'A')" 'Acá se usó la función definida FormatoDecimalPunto()
        db.Execute sql, SOpt
        MsgBox "Curso Incluido", vbInformation, "Curso"
        cmdIncluir.Enabled = False
        cmdBuscar.Enabled = False
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        txtnombre.Enabled = False
        txtHoras.Enabled = False
        txtCant.Enabled = False
        txtCosto.Enabled = False
    End If
    
    
End Sub

Private Sub CmdModificar_Click()

    If (cmdModificar.Caption = "Modificar") Then
        txtcodigo.Enabled = False
        txtnombre.Enabled = True
        txtHoras.Enabled = True
        txtCant.Enabled = True
        txtCosto.Enabled = True
        cmdModificar.Caption = "Guardar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Guardar.jpg")
        Exit Sub
    End If

    If (txtCosto.Text = "" Or Val(Mid(txtCosto.Text, 4)) <= 0) Then
        lblErrorCosto.Visible = True
    End If
    If (txtHoras.Text = "") Then
        lblErrorHoras.Caption = "Indique horas"
        lblErrorHoras.Visible = True
    End If
    If (txtCant.Text = "") Then
        lblErrorParticipantes.Caption = "Indique cantidad de participantes"
        lblErrorParticipantes.Visible = True
    End If
    If (txtnombre.Text = "") Then
        lblErrorNombre.Caption = "Indique nombre"
        lblErrorNombre.Visible = True
    End If
    
    
    If (Val(txtHoras.Text) < 20) Or (Val(txtHoras.Text) > 120) Then
        lblErrorHoras.Caption = "Debe estar entre 20 y 120."
        lblErrorHoras.Visible = True
    End If
    
    If (Val(txtCant.Text) < 10) Or (Val(txtCant.Text) > 30) Then
        lblErrorParticipantes.Caption = "Debe estar entre 10 y 30"
        lblErrorParticipantes.Visible = True
    End If
    
    If (lblErrorNombre.Visible) Then
        txtnombre.SetFocus
    Exit Sub
    End If

    If (lblErrorHoras.Visible) Then
        txtHoras.SetFocus
        Exit Sub
    End If
    
    If (lblErrorParticipantes.Visible) Then
        txtCant.SetFocus
        Exit Sub
    End If
    
    If (lblErrorCosto.Visible) Then
        txtCosto.SetFocus
        Exit Sub
    End If
    
    If Not (lblErrorNombre.Visible Or lblErrorHoras.Visible Or lblErrorParticipantes.Visible Or lblErrorCosto.Visible) Then
        Set rs = New ADODB.Recordset
        sql = "Select * from TCursos where ccodigo = '" & txtcodigo.Text & "' and cestatus= ('A')"
        rs.Open sql, db, adOpenStatic
        If (rs!cnombre = txtnombre.Text And rs!cduracionH = txtHoras.Text And rs!ccantidadP = txtCant.Text And rs!ccosto = CDbl(Mid(txtCosto.Text, 5))) Then
            MsgBox "No hubieron cambios.", vbExclamation, "Modificar cursos"
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            txtnombre.Enabled = False
            txtHoras.Enabled = False
            txtCant.Enabled = False
            txtCosto.Enabled = False
            Exit Sub
        End If
        
        resp = MsgBox("¿Estás seguro que quieres Modificar los Datos ?", vbYesNo + vbQuestion, "Modificar curso")
        If (resp = vbYes) Then
            sql = "update TCursos set cnombre = '" & txtnombre.Text & "', cduracionH = '" & txtHoras.Text & "', ccantidadP = '" & txtCant.Text & "', ccosto = '" & FormatoDecimalPunto(txtCosto.Text) & "' where ccodigo = '" & txtcodigo.Text & "' " 'Acá se usó la función definida FormatoDecimalPunto()
            db.Execute sql, SOpt
            MsgBox "Curso modificado Exitosamente!", vbInformation, "Modificar curso"
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            txtnombre.Enabled = False
            txtHoras.Enabled = False
            txtCant.Enabled = False
            txtCosto.Enabled = False
        Else
            CmdBuscar_Click
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            txtnombre.Enabled = False
            txtHoras.Enabled = False
            txtCant.Enabled = False
            txtCosto.Enabled = False
        End If
    End If
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Activate()
'Para que la imagen de fondo se autoajuste al tamaño del form MAXIMIZADO
If (Image1.Height < (Me.Height - 300)) Then
    Image1.Height = Me.Height - 300
End If
If (Image1.Width < (Me.Width - 100)) Then
    Image1.Width = Me.Width - 100
End If
'fin autoajustar
End Sub

Private Sub Form_Load()

    'Validaciones de inicio
    txtnombre.Enabled = False
    txtHoras.Enabled = False
    txtCant.Enabled = False
    txtCosto.Enabled = False
    txtnombre.Text = ""
    txtHoras.Text = ""
    txtCant.Text = ""
    txtCosto.Text = ""
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    txtcodigo.MaxLength = 4
    txtnombre.MaxLength = 30
    txtHoras.MaxLength = 3
    txtCant.MaxLength = 2
    txtCosto.MaxLength = 16
    
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    cmdModificar.Caption = "Modificar"
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPrincipal.Show

End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm AM/PM")
End Sub

Private Sub txtCant_Change()
If (txtCant.Text <> "") Then
    lblErrorParticipantes.Visible = False
End If
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
    KeyAscii = 0
    End If

End Sub

Private Sub txtCodigo_Change()

    If (txtcodigo.Text <> "") Then
    cmdBuscar.Enabled = True
    Else
    cmdBuscar.Enabled = False
    End If
    

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
    KeyAscii = 0
End If
    If (KeyAscii = 13 And cmdBuscar.Enabled) Then
    CmdBuscar_Click
    End If

End Sub

Private Sub txtCosto_Change()
If (txtCosto.Text <> "") Then
    lblErrorCosto.Visible = False
End If
End Sub

Private Sub txtCosto_GotFocus()

txtCosto.Text = Mid(txtCosto.Text, 5)
txtCosto.SelStart = Len(txtCosto.Text)
End Sub

Private Sub txtCosto_LostFocus()
    txtCosto.Text = Format(txtCosto.Text, "\B\s\. #,##0.000") ' separador de miles y 3 decimales.
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 44) Then 'Permite números, borrar y comas.
        KeyAscii = 0
    End If
    
    If (txtCosto.SelStart = 0 And KeyAscii = 44) Then  'Debe ser un número el primer caracter, no puede empezar con coma.
        KeyAscii = 0
    End If
        
    If (KeyAscii = 44 And InStr(1, txtCosto.Text, ",") <> 0) Then 'Esto es para que sólo se pueda poner una coma.
        'InStr retorna la posición de "," en la cadena txtCosto.text, si no la encuentra la posición es 0
        'Osea que si posicion<>0 es porque ya hay una coma en el número.
        KeyAscii = 0
    End If
    
    If (InStr(1, txtCosto.Text, ",") <> 0) Then 'Si ya hay una coma
        If (txtCosto.SelStart >= InStr(1, txtCosto.Text, ",") And Len(txtCosto.Text) > (InStr(1, txtCosto.Text, ",") + 2) And Not KeyAscii = 8) Then 'Permite solo 3 numeros luego de la coma
            'Si está escribiendo después de la coma y ya el número tiene 3 decimales que no lo permita (salvo que le dé a borrar)
            KeyAscii = 0
        End If
    End If
    

        
    

End Sub

Private Sub txtHoras_Change()
If (txtHoras.Text <> "") Then
    lblErrorHoras.Visible = False
End If
End Sub

Private Sub txtHoras_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
    KeyAscii = 0
    End If

End Sub

Private Sub txtNombre_Change()
If (txtnombre.Text <> "") Then
    lblErrorNombre.Visible = False
End If
End Sub

