VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmGrupos 
   Caption         =   "Grupos"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "frmGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbTurnos 
      Height          =   315
      Left            =   5640
      TabIndex        =   11
      Top             =   4560
      Width           =   3375
   End
   Begin VB.ComboBox cmbCursos 
      Height          =   315
      ItemData        =   "frmGrupos.frx":058A
      Left            =   5640
      List            =   "frmGrupos.frx":058C
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox cmbInstructores 
      Height          =   315
      ItemData        =   "frmGrupos.frx":058E
      Left            =   5640
      List            =   "frmGrupos.frx":0590
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin VB.ComboBox cmbGrupos 
      Height          =   315
      ItemData        =   "frmGrupos.frx":0592
      Left            =   5640
      List            =   "frmGrupos.frx":0594
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11040
      Top             =   1800
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H8000000B&
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
      Left            =   1920
      Picture         =   "frmGrupos.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdIncluir 
      BackColor       =   &H8000000B&
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
      Left            =   4080
      Picture         =   "frmGrupos.frx":0DA8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H8000000B&
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
      Left            =   6240
      Picture         =   "frmGrupos.frx":2ABE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H8000000B&
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
      Left            =   8400
      Picture         =   "frmGrupos.frx":4894
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1700
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H8000000B&
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
      Left            =   4080
      Picture         =   "frmGrupos.frx":7445
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H8000000B&
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
      Left            =   6240
      Picture         =   "frmGrupos.frx":9777
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   1700
   End
   Begin MSComCtl2.DTPicker dtpFechaini 
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      Top             =   3360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48955393
      CurrentDate     =   42432
   End
   Begin MSComCtl2.DTPicker dtpFechacul 
      Height          =   315
      Left            =   5640
      TabIndex        =   13
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48955393
      CurrentDate     =   42432
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   690
      Left            =   2280
      TabIndex        =   23
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1217
      Image           =   "frmGrupos.frx":B080
      Props           =   5
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de grupo"
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
      Left            =   2400
      TabIndex        =   22
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Label lblCodigoC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de curso"
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
      Left            =   2400
      TabIndex        =   21
      Top             =   1560
      Width           =   2340
   End
   Begin VB.Label lblFechaini 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial"
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
      Left            =   2400
      TabIndex        =   20
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label lblFechacul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha culminación"
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
      Left            =   2400
      TabIndex        =   19
      Top             =   3960
      Width           =   2685
   End
   Begin VB.Label lblTurno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turno"
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
      Left            =   2400
      TabIndex        =   18
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblInstructor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor"
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
      Left            =   2400
      TabIndex        =   17
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label lblErrorInstructor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9120
      TabIndex        =   16
      Top             =   2760
      Width           =   105
   End
   Begin VB.Label lblErrorTurno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9120
      TabIndex        =   15
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label lblErrorGrupoInstructor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*El Instructor ya tiene un grupo ese turno y fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2970
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   11880
      Picture         =   "frmGrupos.frx":B7E6
      Top             =   960
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
      Left            =   13020
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmGrupos.frx":EE66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12720
   End
End
Attribute VB_Name = "frmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub activacionEliminado()

    Set rs = New ADODB.Recordset
    sql = "select * from tgrupos where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='A' and not gcodigo='" & Mid(cmbGrupos.Text, 1, 4) & "' and not gcodigoc='" & Mid(cmbCursos.Text, 1, 4) & "' " _
    & "union " _
    & "select * from tgruposculminados where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='C'"
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        lblErrorGrupoInstructor.Visible = False
    Else
        lblErrorGrupoInstructor.Visible = True
        Exit Sub
    End If
    
    If (dtpFechacul.Value < Date) Then 'Si al modificar el grupo hace que este culmine
        resp = MsgBox("El grupo pasará a culminado, ¿Desea modificar de todos modos?", vbExclamation + vbYesNo, "Modificar Grupos")
        If (resp = vbYes) Then
            sql = "update TGrupos set gcodigoi = '" & Mid(cmbInstructores.Text, 1, 4) & "', gfechaini = to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), gfechacul = to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), gcodturno = '" & Mid(cmbTurnos.Text, 1, 4) & "', gestatus='A' where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus='E' "
            db.Execute sql, SOpt
            MsgBox "Grupo modificado Exitosamente!", vbInformation, "Modificar Grupos"
            CulminarGrupos 'Actualiza el grupo porque modificó la fecha haciendo que culmine.
            CmdCancelar_Click
        End If
    Exit Sub 'si resp=vbNo, simplemente no lo modificará, dándole la posibilidad de cambiar las fecha para modificarlo bien.
    End If
    
    sql = "update TGrupos set gcodigoi = '" & Mid(cmbInstructores.Text, 1, 4) & "', gfechaini = to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), gfechacul = to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), gcodturno = '" & Mid(cmbTurnos.Text, 1, 4) & "', gestatus='A' where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus='E'"
    db.Execute sql, SOpt
    MsgBox "Grupo reactivado Exitosamente!", vbInformation, "Modificar Grupos"
    cmdModificar.Caption = "Modificar"
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    CmdBuscar_Click
    cmbInstructores.Enabled = False
    dtpFechaini.Enabled = False
    dtpFechacul.Enabled = False
    cmbTurnos.Enabled = False
    
End Sub

Private Sub cmbCursos_Click()
    'llenar combobox grupos
        cmbGrupos.Enabled = True
        cmbGrupos.Clear
        Set rs = New ADODB.Recordset
        sql = "select gcodigo,gestatus from TGrupos, TCursos where gcodigoc= '" & Mid(cmbCursos.Text, 1, 4) & "' and ccodigo= '" & Mid(cmbCursos.Text, 1, 4) & "' and cestatus='A' and gestatus IN ('A','E') "
        rs.Open sql, db, adOpenStatic
        If Not rs.EOF Then
            Do
                Dim estatus As String
                Select Case rs!gestatus
                Case "A"
                    estatus = "Activo"
                Case "E"
                    estatus = "Eliminado"
                End Select
                
                cmbGrupos.AddItem (rs!gcodigo & " - " & estatus)
                rs.MoveNext
            Loop Until rs.EOF
        End If
        
        If (cmbGrupos.ListCount = 0) Then
            cmbGrupos.Text = "Sin grupos"
            cmbGrupos.Enabled = False
        End If
    'fin llenar combobox grupos
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = True
    cmdCancelar.Enabled = True
    
End Sub

Private Sub cmbCursos_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46) Then 'Inhabilita el botón Supr, al parecer con el evento keypress no es suficiente.
KeyCode = 0
End If
End Sub

Private Sub cmbCursos_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbInstructor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbGrupos_Click()
cmdBuscar.Enabled = True
cmdIncluir.Enabled = False

End Sub

Private Sub cmbGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46) Then 'Inhabilita el botón Supr, al parecer con el evento keypress no es suficiente.
KeyCode = 0
End If
End Sub

Private Sub cmbGrupos_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 13) Then
KeyAscii = 0
Else
CmdBuscar_Click
End If
End Sub

Private Sub cmbInstructores_Click()
lblErrorInstructor.Visible = False
End Sub

Private Sub cmbInstructores_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46) Then 'Inhabilita el botón Supr, al parecer con el evento keypress no es suficiente.
KeyCode = 0
End If
End Sub

Private Sub cmbTurnos_Click()
lblErrorTurno.Visible = False
End Sub

Private Sub cmbTurnos_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46) Then 'Inhabilita el botón Supr, al parecer con el evento keypress no es suficiente.
KeyCode = 0
End If
End Sub

Private Sub cmbTurnos_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CmdBuscar_Click()
'Buscando un grupo
    Set rs = New ADODB.Recordset
    sql = "Select * from TGrupos where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus IN ('A','E')"
    rs.Open sql, db, adOpenStatic
        
    If rs.EOF Then
        Exit Sub
    End If
    
        For X = 0 To cmbCursos.ListCount
            If (rs!gcodturno = Mid(cmbTurnos.List(X), 1, 4)) Then
                cmbTurnos.Text = cmbTurnos.List(X)
            End If
        Next X
        
        
        Dim estatus As String
        Select Case rs!gestatus
            Case "A"
                estatus = "Activo"
            Case "E"
                estatus = "Eliminado"
        End Select
        cmbGrupos.Text = rs!gcodigo & " - " & estatus
        'For X = 0 To cmbGrupos.ListCount
           ' If (rs!gcodigo = Mid(cmbGrupos.List(X), 1, 4)) Then
            '    cmbGrupos.Text = cmbGrupos.List(X) & rs!gestatus
           ' End If
       ' Next X

        
        For X = 0 To cmbInstructores.ListCount
            If (rs!gcodigoi = Mid(cmbInstructores.List(X), 1, 4)) Then
                cmbInstructores.Text = cmbInstructores.List(X)
            End If
        Next X
        
        dtpFechaini.Value = rs!gfechaini
        dtpFechacul.Value = rs!gfechacul

        For X = 0 To cmbTurnos.ListCount
            If (rs!gcodturno = Mid(cmbTurnos.List(X), 1, 4)) Then
                cmbTurnos.Text = cmbTurnos.List(X)
            End If
        Next X
        
        If (rs!gestatus = "E") Then
            resp = MsgBox("Grupo eliminado, ¿Desea reactivarlo?", vbYesNo + vbQuestion, "Grupos")
            If (resp = vbYes) Then
                cmbCursos.Enabled = False
                cmbGrupos.Enabled = False
                cmdBuscar.Enabled = False
                cmdIncluir.Enabled = False
                cmdModificar.Enabled = True
                cmdCancelar.Enabled = True
                CmdModificar_Click 'No se activa el grupo de una vez, porque primero debe pasar la validación del instructor con lo del turno y fechas.
                rs.Close
                Exit Sub
            Else 'Si dice que no simplemente quedan los datos en pantalla pero no puede hacer más nada salvo Cancelar (limpiar) o salir.
                cmbCursos.Enabled = False
                cmbGrupos.Enabled = False
                cmdIncluir.Enabled = False
                cmdBuscar.Enabled = False
                cmdCancelar.Enabled = True
                rs.Close
                Exit Sub
            End If
        End If
               
        cmbCursos.Enabled = False
        cmbGrupos.Enabled = False
        cmdBuscar.Enabled = False
        cmdIncluir.Enabled = False
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = True
        
    rs.Close
    'SendKeys "{TAB}" 'Oprime la tecla tabulador para que no quede el focus en el botón salir.

End Sub

Private Sub CmdCancelar_Click()


cmbCursos.Text = ""
cmbGrupos.Text = ""
cmbInstructores.Text = ""
dtpFechacul.MinDate = dtpFechaini.MinDate
dtpFechaini.Value = Date
dtpFechacul.Value = Date
cmbTurnos.Text = ""


cmdBuscar.Enabled = False
cmdIncluir.Enabled = False
cmdModificar.Enabled = False
cmdEliminar.Enabled = False
cmdCancelar.Enabled = False

cmbCursos.Enabled = True
cmbGrupos.Enabled = False
cmbInstructores.Enabled = False
dtpFechaini.Enabled = False
dtpFechacul.Enabled = False
cmbTurnos.Enabled = False


cmbCursos.SetFocus
cmdModificar.Caption = "Modificar"
cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
cmdIncluir.Caption = "Incluir"
lblErrorInstructor.Visible = False
lblErrorTurno.Visible = False


End Sub

Private Sub CmdEliminar_Click()
    resp = MsgBox("¿Está seguro que desea eliminar el grupo?", vbYesNo + vbQuestion, "Eliminar grupo")
    If (resp = vbYes) Then
        sql = "update TGrupos set gestatus = 'E' where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "'"
        db.Execute sql, SOpt
        CmdCancelar_Click
        MsgBox "Eliminacion Exitosa", vbInformation, "Eliminar grupos"
    End If
End Sub

Private Sub CmdIncluir_Click()
    If (cmdIncluir.Caption = "Incluir") Then
   
        If (cmbGrupos.ListCount = 5) Then
            MsgBox "El curso " & Mid(cmbCursos.Text, 8) & " ya tiene asignado 5 grupos", vbExclamation, "Incluir grupo"
            Exit Sub
        End If
        cmbCursos.Enabled = False
        cmbGrupos.Enabled = False
        Set rs = New ADODB.Recordset
        sql = "Select gcodigo from TGrupos where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gestatus IN ('A','E') order by gcodigo"
        rs.Open sql, db, adOpenStatic
        If rs.EOF Then
            cmbGrupos.Text = "0001"
        Else
            Dim X As String
            X = 0
            Do
                X = Format(X + 1, "000#") ' La primera vez x="0001", luego x="0002"...
                If Not (rs!gcodigo = X) Then 'Si el codigo del primer(segundo,tercer...) registro no es 0001(0002,0003...), es porque ese es el codigo que falta.
                    cmbGrupos.Text = X
                    Exit Do 'Si ya se encontró el codigo que falta, que salga del ciclo.
                End If
                rs.MoveNext 'Se mueve al siguiente registro.
                On Error Resume Next 'Si surge un error que lo ignore y siga.
                'Supongan que sólo existe el grupo 0001, el siguiente a incluir es el 0002. Al primer ciclo 0001=0001(no entra al if), luego rs.movenext
                'pero no existe un siguiente registro (ya está EOF) dará error, no importa, que siga el ciclo hará el if
                'como en rs!gcodigo no hay nada(null), null es distinto a 0002, por tanto el codigo que falta es 0002.
            Loop Until X = "0005"
        End If
        
        dtpFechaini.Enabled = True
        dtpFechacul.Enabled = True
        cmbInstructores.Enabled = True
        cmbTurnos.Enabled = True
        cmdBuscar.Enabled = False
        cmdCancelar.Enabled = True
        cmdIncluir.Caption = "Guardar"
        dtpFechacul.MinDate = dtpFechaini.Value
        Exit Sub
    End If
    
    'Inicio colocar labels de errores
    If (cmbInstructores.Text = "" Or cmbTurnos.Text = "") Then
        If (cmbInstructores.Text = "") Then
            cmbInstructores.SetFocus
            lblErrorInstructor.Caption = "* Debe colocar instructor."
            lblErrorInstructor.Visible = True
        Else
            lblErrorInstructor.Visible = False
        End If
        If (cmbTurnos.Text = "") Then
            cmbTurnos.SetFocus
            lblErrorTurno.Caption = "* Debe colocar el turno."
            lblErrorTurno.Visible = True
        Else
            lblErrorTurno.Visible = False
        End If
        If (cmbInstructores.Text = "" And cmbTurnos.Text = "") Then
            cmbInstructores.SetFocus
        End If
    Exit Sub
    End If
    
    'Para que no incluyan un grupo con un instructor que ya tiene un grupo en el mismo turno y mismo rango de fechas.
    'Obvio un instructor no puede estar en dos lugares a la vez.
    Set rs = New ADODB.Recordset
    'El where es sencillamente para buscar un (o varios) grupo de los grupos existentes, si ALGUNO de ellos sus fechas se INTERCEPTAN con las fechas del nuevo grupo no permitimos incluir ese nuevo grupo.
    sql = "select * from tgrupos where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='A' " _
    & "union " _
    & "select * from tgruposculminados where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='C'"
    'Se usó el comando sql union para unir en un solo listado el resultado de la tabla tgrupo y tgruposculminados
    'Esto es porque si quieren incluir un nuevo grupo culminado podría chocar con un grupo culminado ya existente
    'O si incluirán un grupo activo pero que su fecha de inicio ya pasó, esa fecha de inicio podría chocar con un grupo recientemene culminado.
    'En fin, con esto se compara el nuevo grupo con TODOS los grupos existentes (salvo los eliminados) para que no choquen entre sí.
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
    lblErrorGrupoInstructor.Visible = False
    Else
    lblErrorGrupoInstructor.Visible = True
    Exit Sub
    End If
    
    'Fin colocar labels de errores
    lblErrorGrupoInstructor.Visible = False
    lblErrorInstructor.Visible = False
    lblErrorTurno.Visible = False
    
    If (cmdIncluir.Caption = "Guardar") Then
    
        If (dtpFechacul.Value < Date) Then 'Si incluye un grupo ya culminado.
            resp = MsgBox("El grupo se incluirá como culminado, ¿Desea incluir de todos modos?", vbExclamation + vbYesNo, "Incluir Grupos")
            If (resp = vbYes) Then
                sql = "Insert Into TGrupos values" _
                & "('" & Mid(cmbGrupos.Text, 1, 4) & "', '" & Mid(cmbCursos.Text, 1, 4) & "', '" & Mid(cmbInstructores.Text, 1, 4) & "', to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), '" & Mid(cmbTurnos.Text, 1, 4) & "', 'A')" 'otra forma sería: substr('" & (cmbTurnos.Text) & "',1,4)
                db.Execute sql, SOpt
                CulminarGrupos 'Actualiza grupos porque incluyó un grupo culminado.
                MsgBox "Grupo Incluido como culminado", vbInformation, "Grupos"
                CmdCancelar_Click
            End If
        Exit Sub 'si resp=vbNo, simplemente no lo incluirá, dándole la posibilidad de cambiar las fecha para incluirlo bien.
        End If
            
        sql = "Insert Into TGrupos values" _
        & "('" & Mid(cmbGrupos.Text, 1, 4) & "', '" & Mid(cmbCursos.Text, 1, 4) & "', '" & Mid(cmbInstructores.Text, 1, 4) & "', to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), '" & Mid(cmbTurnos.Text, 1, 4) & "', 'A')" 'otra forma sería: substr('" & (cmbTurnos.Text) & "',1,4)
        db.Execute sql, SOpt
        MsgBox "Grupo Incluido", vbInformation, "Grupos"
        CmdBuscar_Click
        cmdIncluir.Enabled = False
        cmdBuscar.Enabled = False
        cmbCursos.Enabled = False
        cmbInstructores.Enabled = False
        dtpFechaini.Enabled = False
        dtpFechacul.Enabled = False
        cmbTurnos.Enabled = False
        cmdIncluir.Caption = "Incluir"
    End If
    

End Sub

Private Sub CmdModificar_Click()

    If (cmdModificar.Caption = "Modificar") Then
        cmbInstructores.Enabled = True
        dtpFechaini.Enabled = True
        dtpFechacul.Enabled = True
        cmbTurnos.Enabled = True
        cmdModificar.Caption = "Guardar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Guardar.jpg")
        dtpFechacul.MinDate = dtpFechaini.Value
        Exit Sub
    End If
    
    'Modificar especial para la reactivación de un grupo eliminado.
    Set rs = New ADODB.Recordset
    sql = "Select * from TGrupos where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus = ('E') "
    rs.Open sql, db, adOpenStatic
    If Not rs.EOF Then 'Si el que se está modificando es un eliminado
        activacionEliminado
        Exit Sub
    End If
    
    
    
    
    Set rs = New ADODB.Recordset
    sql = "Select * from TGrupos where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus = ('A') "
    rs.Open sql, db, adOpenStatic
        If (rs!gcodigoi = Mid(cmbInstructores.Text, 1, 4) And rs!gfechaini = dtpFechaini.Value And rs!gfechacul = dtpFechacul.Value And rs!gcodturno = Mid(cmbTurnos.Text, 1, 4)) Then
            MsgBox "No hubieron cambios.", vbExclamation, "Modificar grupo"
            cmbInstructores.Enabled = False
            dtpFechaini.Enabled = False
            dtpFechacul.Enabled = False
            cmbTurnos.Enabled = False
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            Exit Sub
        End If
    
    '>>>>LEER PRIMERO ESTA VALIDACIÓN EN EL INCLUIR<<<<<
    'Para que no modifiquen un grupo colocando un instructor que ya tiene un grupo en el mismo turno y mismo rango de fechas.
    Set rs = New ADODB.Recordset
    sql = "select * from tgrupos where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='A' and not gcodigo='" & Mid(cmbGrupos.Text, 1, 4) & "' and not gcodigoc='" & Mid(cmbCursos.Text, 1, 4) & "' " _
    & "union " _
    & "select * from tgruposculminados where ((gfechaini between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') or gfechacul between to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and to_date('" & dtpFechacul.Value & "','dd/mm/yyyy')) or (gfechaini <= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul >= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'))) and gcodigoi= '" & Mid(cmbInstructores.Text, 1, 4) & "' and gcodturno= '" & Mid(cmbTurnos.Text, 1, 4) & "' and gestatus='C'"
    'En el modificar hay un caso especial, porque NO hay que tomar en cuenta el grupo que se está modificando.
    'No tiene sentido que modifique un grupo y me diga "el instructor ya tiene un GRUPO en ese turno"
    'Siendo que estoy modificado ESE grupo.
    'Por eso en el final del where de TGrupos (grupos activos) está "and not gcodigo=..... and not gcodigoc=....."
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        lblErrorGrupoInstructor.Visible = False
    Else
        lblErrorGrupoInstructor.Visible = True
        Exit Sub
    End If
    
    
    If (dtpFechacul.Value < Date) Then 'Si al modificar el grupo hace que este culmine
        resp = MsgBox("El grupo pasará a culminado, ¿Desea modificar de todos modos?", vbExclamation + vbYesNo, "Modificar Grupos")
        If (resp = vbYes) Then
            sql = "update TGrupos set gcodigoi = '" & Mid(cmbInstructores.Text, 1, 4) & "', gfechaini = to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), gfechacul = to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), gcodturno = '" & Mid(cmbTurnos.Text, 1, 4) & "' where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus='A' "
            db.Execute sql, SOpt
            MsgBox "Grupo modificado Exitosamente!", vbInformation, "Modificar Grupos"
            CulminarGrupos 'Actualiza el grupo porque modificó la fecha haciendo que culmine.
            CmdCancelar_Click
        End If
    Exit Sub 'si resp=vbNo, simplemente no lo modificará, dándole la posibilidad de cambiar las fecha para modificarlo bien.
    End If
    
    resp = MsgBox("¿Estás seguro que quieres Modificar los Datos ?", vbYesNo + vbQuestion, "Modificar grupo")
    If (resp = vbYes) Then
        sql = "update TGrupos set gcodigoi = '" & Mid(cmbInstructores.Text, 1, 4) & "', gfechaini = to_date('" & dtpFechaini.Value & "','dd/mm/yyyy'), gfechacul = to_date('" & dtpFechacul.Value & "','dd/mm/yyyy'), gcodturno = '" & Mid(cmbTurnos.Text, 1, 4) & "' where gcodigoc = '" & Mid(cmbCursos.Text, 1, 4) & "' and gcodigo= '" & Mid(cmbGrupos.Text, 1, 4) & "' and gestatus='A'"
        db.Execute sql, SOpt
        MsgBox "Grupo modificado Exitosamente!", vbInformation, "Modificar Grupos"
        cmdModificar.Caption = "Modificar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
        CmdBuscar_Click
        cmbInstructores.Enabled = False
        dtpFechaini.Enabled = False
        dtpFechacul.Enabled = False
        cmbTurnos.Enabled = False
    Else
        CmdBuscar_Click
        cmdModificar.Caption = "Modificar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
        cmbInstructores.Enabled = False
        dtpFechaini.Enabled = False
        dtpFechacul.Enabled = False
        cmbTurnos.Enabled = False
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dtpFechaini_Change()
If (cmdIncluir.Caption = "Guardar" Or cmdModificar.Caption = "Guardar") Then
dtpFechacul.MinDate = dtpFechaini.Value 'obliga a fecha de culminación a ser mayor que la de inicio.
Else
dtpFechacul.MinDate = dtpFechaini.MinDate
End If

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

Private Sub frameGrupos_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm AM/PM")
End Sub
Private Sub Form_Load()

cmdBuscar.Enabled = False
cmdIncluir.Enabled = False
cmdModificar.Enabled = False
cmdEliminar.Enabled = False
cmdCancelar.Enabled = False
cmbTurnos.Enabled = False
cmbInstructores.Enabled = False
cmbGrupos.Enabled = False
dtpFechaini.Enabled = False
dtpFechacul.Enabled = False


'llenar combobox turnos
    Set rs = New ADODB.Recordset
    sql = "select * from Tturnos where testatus='A'"
    rs.Open sql, db, adOpenStatic
    If Not rs.EOF Then
        Do
            cmbTurnos.AddItem (rs!tcodigo & " - " & rs!tturno)
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
'fin llenar combobox turnos

'llenar combobox instructores
    Set rs = New ADODB.Recordset
    sql = "select * from TInstructores where iestatus='A'"
    rs.Open sql, db, adOpenStatic
    If Not rs.EOF Then
        Do
            cmbInstructores.AddItem (rs!icodigo & " - " & rs!inombres & " " & rs!iapellidos)
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
'fin llenar combobox instructores

'llenar combobox cursos
    Set rs = New ADODB.Recordset
    sql = "select * from TCursos where cestatus='A'"
    rs.Open sql, db, adOpenStatic
    If Not rs.EOF Then
        Do
            cmbCursos.AddItem (rs!ccodigo & " - " & rs!cnombre)
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
'fin llenar combobox cursos

dtpFechaini.Value = Date
dtpFechacul.Value = Date
cmdModificar.Caption = "Modificar"
cmdIncluir.ToolTipText = "Incluir nuevo grupo"
lblErrorInstructor.Visible = False
lblErrorTurno.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPrincipal.Show
End Sub

