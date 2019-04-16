VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00004000&
   Caption         =   "To Infinity And Beyond C.A."
   ClientHeight    =   10710
   ClientLeft      =   -2610
   ClientTop       =   -2130
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   450
      Left            =   6720
      TabIndex        =   9
      Top             =   5280
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   794
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   11880
      Top             =   360
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   1575
      Left            =   10440
      Picture         =   "frmUsuarioTipo2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1875
   End
   Begin VB.CommandButton cmdReporteGrupos 
      Height          =   2655
      Left            =   3960
      Picture         =   "frmUsuarioTipo2.frx":13FB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   2100
   End
   Begin VB.CommandButton cmdGrupos 
      Height          =   2295
      Left            =   5880
      Picture         =   "frmUsuarioTipo2.frx":34AD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdReporteCursos 
      Height          =   2655
      Left            =   1560
      Picture         =   "frmUsuarioTipo2.frx":4CF8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1980
   End
   Begin VB.CommandButton cmdUsuarios 
      Height          =   2295
      Left            =   2520
      Picture         =   "frmUsuarioTipo2.frx":69DC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton cmdInstructores 
      Height          =   2295
      Left            =   2520
      Picture         =   "frmUsuarioTipo2.frx":F524
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdCursos 
      Height          =   2295
      Left            =   480
      Picture         =   "frmUsuarioTipo2.frx":113AC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   11640
      Picture         =   "frmUsuarioTipo2.frx":12C58
      Top             =   840
      Width           =   885
   End
   Begin VB.Image imgGruposCulminar 
      Height          =   480
      Left            =   7680
      Picture         =   "frmUsuarioTipo2.frx":162D8
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6180
   End
   Begin VB.Label lblBienvenido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido, "
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
      Left            =   9000
      TabIndex        =   10
      Top             =   3240
      Width           =   2190
   End
   Begin VB.Image imgAcento 
      Height          =   180
      Left            =   9960
      Picture         =   "frmUsuarioTipo2.frx":20AD4
      Top             =   4320
      Width           =   225
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
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
      Left            =   12840
      TabIndex        =   8
      Top             =   1920
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
      Left            =   12840
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6015
      Left            =   0
      Picture         =   "frmUsuarioTipo2.frx":20F99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14205
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menú"
      Begin VB.Menu mCursos 
         Caption         =   "Cursos"
      End
      Begin VB.Menu mInstructores 
         Caption         =   "Instructores"
      End
      Begin VB.Menu mGrupos 
         Caption         =   "Grupos"
      End
      Begin VB.Menu mUsuarios 
         Caption         =   "Administrador de Usuarios"
      End
   End
   Begin VB.Menu mReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mCulminados 
         Caption         =   "Grupos Culminados"
      End
      Begin VB.Menu mListado 
         Caption         =   "Listado de Cursos"
      End
   End
   Begin VB.Menu mSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCursos_Click()
Me.Hide
frmCursos.Show
End Sub

Private Sub cmdGrupos_Click()
Me.Hide
frmGrupos.Show
End Sub

Private Sub cmdInstructores_Click()
Me.Hide
frmInstructor.Show
End Sub

Private Sub cmdReporteCursos_Click()
Me.Hide
frmListadoCursos.Show
End Sub

Private Sub cmdReporteGrupos_Click()
Me.Hide
frmCulminado.Show
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdUsuarios_Click()
Me.Hide
frmUsuario.Show
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Activate() 'Este evento se activa cada vez que el formulario pasa a ser la ventana activa
'Usamos este evento y no Form Load porque en el load aún el form no está maximizado.
'Además el form Load se ejecuta sólo cuando se carga, si el form está oculto y lo activo con form.show el evento Load no se ejecuta

lblBienvenido.Caption = "Bienvenido, " & nombreUser & " " & apellidoUser

'Para que la imagen de fondo se autoajuste al tamaño del form MAXIMIZADO
If (Image1.Height < (Me.Height - 1000)) Then
    Image1.Height = Me.Height - 1000
End If
If (Image1.Width < (Me.Width - 100)) Then
    Image1.Width = Me.Width - 100
End If
'fin autoajustar

'llenar datagrid grupos proximos a culminar
    Set rs = New ADODB.Recordset
    'sql = "select cnombre  as Curso, gcodigo as Grupo, gfechaini as Fecha_Inicio,gfechacul as Fecha_Culminación,decode(trunc(gfechacul - (sysdate -1)),0,'Hoy Culmina el Grupo ',trunc(gfechacul - (sysdate - 1)))as Dias_Restantes from TGrupos,TCursos where gcodigoc=ccodigo and gfechacul >= sysdate and trunc(gfechacul - sysdate)<= 7 and gestatus = 'A' order by gfechacul desc"
    sql = "select gcodigo as Grupo, cnombre as Curso,  gfechaini as Fecha_Inicio, gfechacul as Fecha_Culminación,decode(to_date(gfechacul,'dd/mm/yyyy')-to_date(sysdate,'dd/mm/yyyy'),0,'Hoy culmina',1,to_date(gfechacul,'dd/mm/yyyy')-to_date(sysdate,'dd/mm/yyyy')||' Día',to_date(gfechacul,'dd/mm/yyyy')-to_date(sysdate,'dd/mm/yyyy')||' Días') as Dias_Restantes from tgrupos,tcursos where gestatus='A' and (to_date(gfechacul,'dd/mm/yyyy')-to_date(sysdate,'dd/mm/yyyy'))<=7 and gcodigoc=ccodigo order by gfechacul"
    rs.Open sql, db, adOpenStatic
    
    If rs.EOF Then
   Set DataGrid2.DataSource = rs
    Else
    Set DataGrid2.DataSource = rs
    DataGrid2.Height = 450 + (rs.RecordCount - 1) * 240 'Para que se autoajuste a la cantidad de filas.
    End If
'fin llenar datagrid grupos proximos a culminar

End Sub


Private Sub Form_Unload(Cancel As Integer)
resp = MsgBox(" ¿Está seguro que desea SALIR? ", vbInformation + vbYesNo)
If (resp = vbNo) Then
    Cancel = 1
ElseIf (resp = vbYes) Then
    frmLogin.Show
End If
End Sub


Private Sub mCulminados_Click()
Me.Hide
frmCulminado.Show
End Sub

Private Sub mCursos_Click()
Me.Hide
frmCursos.Show
End Sub

Private Sub mGrupos_Click()
Me.Hide
frmGrupos.Show
End Sub

Private Sub mInstructores_Click()
Me.Hide
frmInstructor.Show
End Sub

Private Sub mListado_Click()
Me.Hide
frmListadoCursos.Show
End Sub

Private Sub mSalir_Click()
Unload Me
End Sub

Private Sub mUsuarios_Click()
Me.Hide
frmUsuario.Show
End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm AM/PM")
End Sub
