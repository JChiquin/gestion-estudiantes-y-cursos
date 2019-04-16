VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmInstructor 
   Caption         =   "Instructor"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   Icon            =   "frmInstructor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcodigo 
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox txtcedula 
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   2160
      Width           =   3000
   End
   Begin VB.TextBox txtnombres 
      Height          =   315
      Left            =   5160
      TabIndex        =   13
      Top             =   2760
      Width           =   3000
   End
   Begin VB.TextBox txtapellidos 
      Height          =   315
      Left            =   5160
      TabIndex        =   12
      Top             =   3360
      Width           =   3000
   End
   Begin VB.TextBox txtdireccion 
      Height          =   315
      Left            =   5160
      TabIndex        =   11
      Top             =   5160
      Width           =   3000
   End
   Begin VB.TextBox txttelefono 
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      Top             =   5760
      Width           =   3000
   End
   Begin VB.ComboBox comboSexo 
      Height          =   315
      ItemData        =   "frmInstructor.frx":058A
      Left            =   5160
      List            =   "frmInstructor.frx":0594
      TabIndex        =   8
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11040
      Top             =   1920
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
      Left            =   2880
      Picture         =   "frmInstructor.frx":05AD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
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
      Left            =   5040
      Picture         =   "frmInstructor.frx":0DBF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
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
      Left            =   7200
      Picture         =   "frmInstructor.frx":2AD5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
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
      Left            =   9360
      Picture         =   "frmInstructor.frx":48AB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
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
      Left            =   5040
      Picture         =   "frmInstructor.frx":745C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1700
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
      Left            =   7200
      Picture         =   "frmInstructor.frx":978E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8520
      Width           =   1700
   End
   Begin MSComCtl2.DTPicker dtpnacimiento 
      Height          =   315
      Left            =   5160
      TabIndex        =   9
      Top             =   4560
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48955393
      CurrentDate     =   42404
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   510
      Left            =   2280
      TabIndex        =   31
      Top             =   720
      Width           =   7485
      _ExtentX        =   13044
      _ExtentY        =   900
      Image           =   "frmInstructor.frx":B097
      Props           =   5
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
      Left            =   3240
      TabIndex        =   30
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label lblcedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
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
      Left            =   3240
      TabIndex        =   29
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label lblNombres 
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
      Left            =   3240
      TabIndex        =   28
      Top             =   2760
      Width           =   1290
   End
   Begin VB.Label lblApellidos 
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
      Left            =   3240
      TabIndex        =   27
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Label lblFechanac 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha/nac"
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
      Left            =   3240
      TabIndex        =   26
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label lblsexo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sexo"
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
      Left            =   3240
      TabIndex        =   25
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblDireccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
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
      Left            =   3240
      TabIndex        =   24
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Label lbltelefono 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
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
      Left            =   3240
      TabIndex        =   23
      Top             =   5760
      Width           =   1260
   End
   Begin VB.Label lblErrorCedula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorNombres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorApellidos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorSexo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorDireccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblErrorTelefono 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   8280
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   11880
      Picture         =   "frmInstructor.frx":B903
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
      Left            =   12915
      TabIndex        =   7
      Top             =   2280
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
      Left            =   12915
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmInstructor.frx":EF83
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmInstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function edad(fechaNacimiento As Date) As Integer
    Dim diasVivo As Long
    diasVivo = Date - fechaNacimiento
    edad = Int(diasVivo / 365) ' Int() quita los decimales, como Trunc() en sql
End Function
Private Sub CmdBuscar_Click()
    'Buscando un instructor
    Set rs = New ADODB.Recordset
    sql = "Select * from TInstructores where icodigo = '" & txtcodigo.Text & "' and iestatus= ('A')"
    rs.Open sql, db, adOpenStatic
    
    If rs.EOF Then
        Set rs = New ADODB.Recordset
        sql = "Select * from TInstructores where icodigo= '" & txtcodigo.Text & "' and iestatus= ('E')"
        rs.Open sql, db, adOpenStatic
    
        If rs.EOF Then
            resp = MsgBox("No está registrado el instructor. ¿Desea Incluirlo? ", vbYesNo + vbQuestion, "Incluir instructor")
            If (resp = vbYes) Then
                txtcodigo.Enabled = True
                txtcedula.Enabled = True
                txtnombres.Enabled = True
                txtapellidos.Enabled = True
                dtpnacimiento.Enabled = True
                txtdireccion.Enabled = True
                txttelefono.Enabled = True
                comboSexo.Enabled = True
                cmdIncluir.Enabled = True
                cmdCancelar.Enabled = True
                cmdBuscar.Enabled = False
                txtcedula.SetFocus
            End If
        Else
                txtcedula.Text = rs!icedula
                txtnombres.Text = rs!inombres
                txtapellidos.Text = rs!iapellidos
                txtdireccion.Text = rs!idireccion
                dtpnacimiento.Value = rs!ifechanac
                txttelefono.Text = rs!itelefono
                If (rs!isexo = "M") Then
                    comboSexo.Text = "Masculino"
                Else
                    comboSexo.Text = "Femenino"
                End If
                
                resp = MsgBox("Instructor inactivo. ¿Desea activarlo?", vbYesNo + vbQuestion, "Instructor")
                If (resp = vbYes) Then
                    sql = "update TInstructores set iestatus = 'A' where icodigo = '" & txtcodigo.Text & "' "
                    db.Execute sql, SOpt
                    MsgBox "Activacion Exitosa", vbInformation, "Intructores"
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
        txtcedula.Text = rs!icedula
        txtnombres.Text = rs!inombres
        txtapellidos.Text = rs!iapellidos
        txtdireccion.Text = rs!idireccion
        dtpnacimiento.Value = rs!ifechanac
        txttelefono.Text = rs!itelefono
        If (rs!isexo = "M") Then
            comboSexo.Text = "Masculino"
        Else
            comboSexo.Text = "Femenino"
        End If
        txtcodigo.Enabled = False
        cmdBuscar.Enabled = False
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = True

    End If
    
    rs.Close
End Sub

Private Sub CmdCancelar_Click()

    cmdIncluir.Enabled = False
    cmdEliminar.Enabled = False
    cmdModificar.Enabled = False
    cmdBuscar.Enabled = False
    cmdCancelar.Enabled = False
    txtcodigo.Text = ""
    txtcedula.Text = ""
    txtnombres.Text = ""
    txtapellidos.Text = ""
    comboSexo.Text = ""
    txtdireccion.Text = ""
    txttelefono.Text = ""
    dtpnacimiento.Value = Date
    txtcodigo.Enabled = True
    txtcodigo.SetFocus
    txtcedula.Enabled = False
    txtnombres.Enabled = False
    txtapellidos.Enabled = False
    dtpnacimiento.Enabled = False
    txtdireccion.Enabled = False
    txttelefono.Enabled = False
    comboSexo.Enabled = False
    cmdModificar.Caption = "Modificar"
    lblErrorCedula.Visible = False
    lblErrorNombres.Visible = False
    lblErrorApellidos.Visible = False
    lblErrorSexo.Visible = False
    lblErrorFecha.Visible = False
    lblErrorDireccion.Visible = False
    lblErrorTelefono.Visible = False
    
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    cmdModificar.Caption = "Modificar"
    



End Sub

Private Sub CmdEliminar_Click()
    'variable que guarda el dia actual del sistema
    'Dim DiaActual As String
    'DiaActual = Format(DateTime.Date, "medium date")
    'resp = MsgBox("¿Está seguro que desea eliminar el instructor?", vbYesNo + vbQuestion, "Eliminar instructor")
    
    'validando instructor en curso activo
    'If (resp = vbYes) Then
        'Buscando instructor en curso
       ' Set rs = New ADODB.Recordset
        'sql = "select * from TGrupos, TInstructores,TCursos where icodigo = (select distinct gcodigoi from TGrupos where gcodigoi = '" & txtcodigo.Text & "') and gcodigoc = ccodigo and gestatus = 'A'"
        'rs.Open sql, db, adOpenStatic
        'DiaCulminacion = Format(rs!gfechacul, "Medium date")
        
        'comparando fechas para saber si el curso termino
        'If (DiaActual <= DiaCulminacion) Then
         '   MsgBox "Accion no valida! El curso aun no ha terminado", vbExclamation
         '   Exit Sub
        'Else
        
        
        resp = MsgBox("¿Está seguro que desea eliminar el instructor?", vbYesNo + vbQuestion, "Eliminar instructor")
        If (resp = vbYes) Then
            Set rs = New ADODB.Recordset
            sql = "select cnombre,gcodigo from TGrupos,tCursos where gcodigoi= '" & txtcodigo.Text & "' and gcodigoc=ccodigo and gestatus='A' and cestatus='A' order by cnombre,gcodigo"
            rs.Open sql, db, adOpenStatic
            
            If rs.EOF Then
                sql = "update TInstructores set iestatus = 'E' where icodigo = '" & txtcodigo.Text & "' "
                db.Execute sql, SOpt
                CmdCancelar_Click
                MsgBox "Eliminacion Exitosa", vbInformation, "Eliminar instructores"
            Else
                Dim Grupos As String
                Do
                    Grupos = Grupos & rs!cnombre & " Grupo: " & rs!gcodigo & vbCrLf
                    rs.MoveNext
                Loop Until rs.EOF
                MsgBox "Acción no valida! El instructor " & txtnombres.Text & " tiene los siguientes grupos activos: " & vbCrLf & vbCrLf & Grupos, vbExclamation, "Eliminar instructor"
            End If
        End If
End Sub

Private Sub CmdIncluir_Click()


    If (comboSexo.Text = "Masculino") Then
        isexoLetra = "M"
    Else
        isexoLetra = "F"
    End If
    
    'Inicio colocar labels de errores
    Set rs = New ADODB.Recordset
    sql = "Select icedula from TInstructores where icedula = '" & txtcedula.Text & "' "
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        lblErrorCedula.Visible = False
    Else
        lblErrorCedula.Caption = "* Cédula ya existe"
        lblErrorCedula.Visible = True
    End If
    
    If (txtcedula.Text = "") Then
        lblErrorCedula.Caption = "* Debe colocar cédula"
        lblErrorCedula.Visible = True
    End If
    
    If (txtnombres.Text = "") Then
        lblErrorNombres.Caption = "* Debe colocar nombre"
        lblErrorNombres.Visible = True
    End If

    If (txtapellidos.Text = "") Then
        lblErrorApellidos.Caption = "* Debe colocar apellido"
        lblErrorApellidos.Visible = True
    End If
    
    If (comboSexo.Text = "") Then
        lblErrorSexo.Caption = "* Debe colocar sexo"
        lblErrorSexo.Visible = True
    End If
    
    If (txtdireccion.Text = "") Then
        lblErrorDireccion.Caption = "* Debe colocar dirección"
        lblErrorDireccion.Visible = True
    End If
    
    If (txttelefono.Text = "") Then
        lblErrorTelefono.Caption = "* Debe colocar teléfono"
        lblErrorTelefono.Visible = True
    End If
    
    If (comboSexo.Text = "Masculino") Then
        If Not (edad(dtpnacimiento.Value) >= 25 And edad(dtpnacimiento.Value) <= 40) Then
            lblErrorFecha.Caption = "* Si es hombre, debe tener entre 25-40 años"
            lblErrorFecha.Visible = True
        Else
            lblErrorFecha.Visible = False
        End If
    End If

    If (comboSexo.Text = "Femenino") Then
        If Not (edad(dtpnacimiento.Value) >= 20 And edad(dtpnacimiento.Value) <= 30) Then
            lblErrorFecha.Caption = "* Si es mujer, debe tener entre 20-30 años"
            lblErrorFecha.Visible = True
        Else
            lblErrorFecha.Visible = False
        End If
    End If
    'Fin colocar labels de errores
    
    'inicio hacer focus al label que dé error
    If (lblErrorCedula.Visible) Then
        txtcedula.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If (lblErrorNombres.Visible) Then
        txtnombres.SetFocus
        Exit Sub
    End If
    
    If (lblErrorApellidos.Visible) Then
        txtapellidos.SetFocus
        Exit Sub
    End If
    
    If (lblErrorSexo.Visible) Then
        comboSexo.SetFocus
        Exit Sub
    End If
    
    If (lblErrorFecha.Visible) Then
        dtpnacimiento.SetFocus
        Exit Sub
    End If
    
    If (lblErrorDireccion.Visible) Then
        txtdireccion.SetFocus
        Exit Sub
    End If
    
    If (lblErrorTelefono.Visible) Then
        txttelefono.SetFocus
        Exit Sub
    End If
    'fin hacer focus al label que dé error
    
    If Not (lblErrorFecha.Visible Or lblErrorCedula.Visible Or lblErrorNombres.Visible Or lblErrorApellidos.Visible Or lblErrorSexo.Visible Or lblErrorDireccion.Visible Or lblErrorTelefono.Visible) Then 'Si no hay un label de error visible, es porque está listo para incluir.
        sql = "Insert Into TInstructores values" _
        & "('" & (txtcodigo.Text) & "', '" & (txtcedula.Text) & "', '" & (txtnombres.Text) & "', '" & (txtapellidos.Text) & "', to_date('" & dtpnacimiento.Value & "','dd/mm/yyyy'), '" & (txtdireccion.Text) & "','" & isexoLetra & "', '" & (txttelefono.Text) & "', 'A')"
        db.Execute sql, SOpt
        MsgBox "Instructor Incluido", vbInformation, "Instructor"
        txtcodigo.Enabled = False
        txtcedula.Enabled = False
        txtnombres.Enabled = False
        txtapellidos.Enabled = False
        dtpnacimiento.Enabled = False
        txtdireccion.Enabled = False
        txttelefono.Enabled = False
        comboSexo.Enabled = False
        cmdBuscar.Enabled = False
        cmdIncluir.Enabled = False
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = True
    End If
    
End Sub

Private Sub CmdModificar_Click()

    If (cmdModificar.Caption = "Modificar") Then
        txtnombres.Enabled = True
        txtapellidos.Enabled = True
        dtpnacimiento.Enabled = True
        txtdireccion.Enabled = True
        txttelefono.Enabled = True
        comboSexo.Enabled = True
        cmdModificar.Caption = "Guardar"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Guardar.jpg")
        isexo = comboSexo.Text
        Exit Sub
    End If
    
    If (comboSexo.Text = "Masculino") Then
        isexoLetra = "M"
    Else
        isexoLetra = "F"
    End If

    'Inicio colocar labels de errores
    Set rs = New ADODB.Recordset
    sql = "Select icedula from TInstructores where icedula = '" & txtcedula.Text & "' "
    rs.Open sql, db, adOpenStatic
    If rs.EOF Then
        lblErrorCedula.Visible = False
    Else
        lblErrorCedula.Caption = "* Cédula ya existe"
        lblErrorCedula.Visible = True
    End If
    
    If (txtcedula.Text = "") Then
        lblErrorCedula.Caption = "* Debe colocar cédula"
        lblErrorCedula.Visible = True
    End If
    
    If (txtnombres.Text = "") Then
        lblErrorNombres.Caption = "* Debe colocar nombre"
        lblErrorNombres.Visible = True
    End If

    If (txtapellidos.Text = "") Then
        lblErrorApellidos.Caption = "* Debe colocar apellido"
        lblErrorApellidos.Visible = True
    End If
    
    If (comboSexo.Text = "") Then
        lblErrorSexo.Caption = "* Debe colocar sexo"
        lblErrorSexo.Visible = True
    End If
    
    If (txtdireccion.Text = "") Then
        lblErrorDireccion.Caption = "* Debe colocar dirección"
        lblErrorDireccion.Visible = True
    End If
    
    If (txttelefono.Text = "") Then
        lblErrorTelefono.Caption = "* Debe colocar teléfono"
        lblErrorTelefono.Visible = True
    End If
    
    If (comboSexo.Text = "Masculino") Then
        If Not (edad(dtpnacimiento.Value) >= 25 And edad(dtpnacimiento.Value) <= 40) Then
            lblErrorFecha.Caption = "* Si es hombre, debe tener entre 25-40 años"
            lblErrorFecha.Visible = True
        Else
            lblErrorFecha.Visible = False
        End If
    End If

    If (comboSexo.Text = "Femenino") Then
        If Not (edad(dtpnacimiento.Value) >= 20 And edad(dtpnacimiento.Value) <= 30) Then
            lblErrorFecha.Caption = "* Si es mujer, debe tener entre 20-30 años"
            lblErrorFecha.Visible = True
        Else
            lblErrorFecha.Visible = False
        End If
    End If
    'Fin colocar labels de errores
    
    'inicio hacer focus al label que dé error
    If (lblErrorCedula.Visible) Then
        txtcedula.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If (lblErrorNombres.Visible) Then
        txtnombres.SetFocus
        Exit Sub
    End If
    
    If (lblErrorApellidos.Visible) Then
        txtapellidos.SetFocus
        Exit Sub
    End If
    
    If (lblErrorSexo.Visible) Then
        comboSexo.SetFocus
        Exit Sub
    End If
    
    If (lblErrorFecha.Visible) Then
        dtpnacimiento.SetFocus
        Exit Sub
    End If
    
    If (lblErrorDireccion.Visible) Then
        txtdireccion.SetFocus
        Exit Sub
    End If
    
    If (lblErrorTelefono.Visible) Then
        txttelefono.SetFocus
        Exit Sub
    End If
    'fin hacer focus al label que dé error
    
    If Not (lblErrorFecha.Visible Or lblErrorCedula.Visible Or lblErrorNombres.Visible Or lblErrorApellidos.Visible Or lblErrorSexo.Visible Or lblErrorDireccion.Visible Or lblErrorTelefono.Visible) Then 'Si no hay un label de error visible, es porque está listo para incluir.
        Set rs = New ADODB.Recordset
        sql = "Select * from TInstructores where icodigo = '" & txtcodigo.Text & "' and iestatus= ('A')"
        rs.Open sql, db, adOpenStatic
        
        If (rs!inombres = txtnombres.Text And rs!iapellidos = txtapellidos.Text And rs!ifechanac = dtpnacimiento.Value And isexo = comboSexo.Text And rs!idireccion = txtdireccion.Text And rs!itelefono = txttelefono.Text) Then
        MsgBox "No hubieron cambios.", vbExclamation, "Modificar Instructores"
        cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
        cmdModificar.Caption = "Modificar"
        txtnombres.Enabled = False
        txtapellidos.Enabled = False
        dtpnacimiento.Enabled = False
        txtdireccion.Enabled = False
        txttelefono.Enabled = False
        comboSexo.Enabled = False
        Exit Sub
        End If
        
        resp = MsgBox("¿Estás seguro que quieres Modificar los Datos ?", vbYesNo + vbQuestion, "Modificar instructores")
        If (resp = vbYes) Then
            sql = "update TInstructores set inombres = '" & txtnombres.Text & "', iapellidos = '" & txtapellidos.Text & "', ifechanac = to_date('" & dtpnacimiento.Value & "','dd/mm/yyyy'), isexo = '" & isexoLetra & "', idireccion = '" & txtdireccion.Text & "', itelefono = '" & txttelefono.Text & "' where icodigo = '" & txtcodigo.Text & "' "
            db.Execute sql, SOpt
            MsgBox "Instructor modificado Exitosamente!", vbInformation, "Modificar instructor"
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            txtnombres.Enabled = False
            txtapellidos.Enabled = False
            dtpnacimiento.Enabled = False
            txtdireccion.Enabled = False
            txttelefono.Enabled = False
            comboSexo.Enabled = False
        Else
            CmdBuscar_Click
            cmdModificar.Caption = "Modificar"
            cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
            txtnombres.Enabled = False
            txtapellidos.Enabled = False
            dtpnacimiento.Enabled = False
            txtdireccion.Enabled = False
            txttelefono.Enabled = False
            comboSexo.Enabled = False
        End If
    End If

End Sub

Private Sub comboSexo_Click()
lblErrorSexo.Visible = False
End Sub

Private Sub comboSexo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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

Private Sub Timer1_Timer()
    lblFecha.Caption = Date
    lblHora.Caption = Format(Time, "hh:mm AM/PM")
End Sub
Private Sub Form_Load()
    lblErrorCedula.Visible = False
    lblErrorFecha.Visible = False
    txtcedula.Enabled = False
    txtnombres.Enabled = False
    txtapellidos.Enabled = False
    dtpnacimiento.Enabled = False
    comboSexo.Enabled = False
    txtdireccion.Enabled = False
    txttelefono.Enabled = False
    txtcodigo.Text = ""
    txtcedula.Text = ""
    txtnombres.Text = ""
    comboSexo.Text = ""
    txtapellidos.Text = ""
    txtdireccion.Text = ""
    txttelefono.Text = ""
    dtpnacimiento.Value = Date
    dtpnacimiento.MaxDate = Date
    cmdBuscar.Enabled = False
    cmdIncluir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    txtcodigo.MaxLength = 4
    txtcedula.MaxLength = 10
    txtnombres.MaxLength = 30
    txtapellidos.MaxLength = 30
    txtdireccion.MaxLength = 30
    txttelefono.MaxLength = 15
    
    cmdModificar.Picture = LoadPicture(App.Path & "\imágenes\Modificar.jpg")
    cmdModificar.Caption = "Modificar"
        
End Sub

Private Sub cmdSalir_Click()
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPrincipal.Show
End Sub


Private Sub txtapellidos_Change()
If (txtapellidos.Text <> "") Then
lblErrorApellidos.Visible = False
End If
End Sub

Private Sub txtapellidos_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtcedula_Change()
If (Len(txtcedula.Text) = 1) Then
lblErrorCedula.Visible = False
End If
End Sub

Private Sub txtcedula_GotFocus()
txtcedula.MaxLength = 8
If (txtcedula.Text <> "") Then
txtcedula.Text = CLng(txtcedula.Text) 'Quita los puntos a la cedula
txtcedula.SelStart = Len(txtcedula.Text)
End If
End Sub

Private Sub txtcedula_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtcedula_LostFocus()
If (txtcedula.Text <> "") Then
txtcedula.MaxLength = 10
txtcedula.Text = FormatNumber(txtcedula.Text, 0, vbFalse, vbFalse, vbTrue) 'pone los puntos a la cedula
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
    
    If (KeyAscii = 13 And cmdBuscar.Enabled) Then
        CmdBuscar_Click
    End If
    If (KeyAscii = 32) Then 'Inhabilitar barra espaciadora
        KeyAscii = 0
    End If
End Sub

Private Sub txtdireccion_Change()
If (txtdireccion.Text <> "") Then
lblErrorDireccion.Visible = False
End If
End Sub

Private Sub txtnombres_Change()
If (txtnombres.Text <> "") Then
lblErrorNombres.Visible = False
End If
End Sub

Private Sub txtnombres_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttelefono_Change()
If (txttelefono.Text <> "") Then
lblErrorTelefono.Visible = False
End If
End Sub

Private Sub txttelefono_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43) Then
        KeyAscii = 0
    End If
End Sub
