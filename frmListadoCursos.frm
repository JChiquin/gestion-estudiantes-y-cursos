VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmListadoCursos 
   Caption         =   "Listado Cursos"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   Icon            =   "frmListadoCursos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7470
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11640
      Top             =   1440
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H8000000A&
      Caption         =   "Mostrar"
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
      Left            =   3840
      Picture         =   "frmListadoCursos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
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
      Left            =   9240
      Picture         =   "frmListadoCursos.frx":1EDF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H8000000A&
      Caption         =   "Imprimir"
      Enabled         =   0   'False
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
      Left            =   6720
      Picture         =   "frmListadoCursos.frx":37E8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1700
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   705
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1244
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Franklin Gothic Medium"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New Baltic"
         Size            =   12
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Código_Curso"
         Caption         =   "Código Curso"
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
         DataField       =   "Curso"
         Caption         =   "Curso"
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
      BeginProperty Column02 
         DataField       =   "Activos"
         Caption         =   "Activos"
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
      BeginProperty Column03 
         DataField       =   "Culminados"
         Caption         =   "Culminados"
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
      BeginProperty Column04 
         DataField       =   "Total"
         Caption         =   "Total"
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
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2894,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794,835
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFechaini 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48955393
      CurrentDate     =   42432
   End
   Begin MSComCtl2.DTPicker dtpFechacul 
      Height          =   315
      Left            =   8400
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48955393
      CurrentDate     =   42432
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   510
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   900
      Image           =   "frmListadoCursos.frx":5720
      Props           =   5
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   12120
      Picture         =   "frmListadoCursos.frx":6169
      Top             =   720
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
      Left            =   13200
      TabIndex        =   9
      Top             =   1800
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
      Left            =   13200
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmListadoCursos.frx":97E9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Culminación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   1605
   End
End
Attribute VB_Name = "frmListadoCursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Listado() As String
    
    'codigo sql que muestra el listado entre el rango de fechas seleccionado.
    Dim sqlistado As String
    'sqlistado = "select ccodigo as Codigo_Curso, cnombre as Curso, count(*) as Cantidad_de_Cursos from tgrupos, tcursos where gfechaini >= '" & Format(dtpFechaini.Value, "medium date") & "' and gfechacul <= '" & Format(dtpFechacul.Value, "medium date") & "' and gcodigoc = ccodigo and gestatus = 'A' group by ccodigo, cnombre"
    'sqlistado = "select ccodigo as Código_Curso, cnombre as Curso, count(*) as Total from tgrupos, tcursos where gfechaini >= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul<= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') and gcodigoc = ccodigo and gestatus = 'A' group by ccodigo, cnombre order by ccodigo"
    sqlistado = "select ccodigo as Código_Curso, cnombre as Curso, count(decode(gestatus,'A',1,null)) as Activos, count(decode(gestatus,'C',1,null)) as Culminados, count(*) as Total " _
    & "from( " _
    & "select ccodigo, cnombre, gestatus from tgrupos, tcursos where gfechaini >= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul<= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') and gcodigoc = ccodigo and gestatus = 'A' " _
    & "union all " _
    & "select ccodigo, cnombre, gestatus from tgruposculminados, tcursos where gfechaini >= to_date('" & dtpFechaini.Value & "','dd/mm/yyyy') and gfechacul<= to_date('" & dtpFechacul.Value & "','dd/mm/yyyy') and gcodigoc = ccodigo and gestatus = 'C') " _
    & "group by ccodigo, cnombre"
    'Lo que se hizo fue un select de un listado hecho con dos select (union)
    Listado = sqlistado
End Function

Private Sub cmdImprimir_Click()
    Set rs = New ADODB.Recordset
    sql = Listado()
    rs.Open sql, db, adOpenStatic
    Set reportCursos.DataSource = rs
    'Para indicar en el encabezado el periodo de fechas que se escogio.
    reportCursos.Sections("Sección4").Controls.Item("fechas").Caption = "En el periodo: " & dtpFechaini.Value & " - " & dtpFechacul.Value
    'Para indiciar en el pie de página el usuario que imprimió el reporte.
    reportCursos.Sections("Sección3").Controls.Item("Etiqueta12").Caption = "Impreso por: " & nombreUser & " " & apellidoUser
    
    'Para colocar los totales al final de cada columna
    Do
        activos = activos + rs!activos
        culminados = culminados + rs!culminados
        total = total + rs!total
        rs.MoveNext
    Loop Until rs.EOF
    reportCursos.Sections("Sección5").Controls.Item("activos").Caption = "Total: " & activos
    reportCursos.Sections("Sección5").Controls.Item("culminados").Caption = "Total: " & culminados
    reportCursos.Sections("Sección5").Controls.Item("total").Caption = "Total: " & total
    'fin colocar totales al final.
    
    
    reportCursos.Show 'mostrar el reporte
    
End Sub

Private Sub cmdMostrar_Click()

    'Mostrando el listado de cursos dependiendo del rango entre fechas
    'If cmdMostrar.Caption = "Mostrar" Then
     
    If cmdMostrar.Caption = "Limpiar" Then
        Set DataGrid1.DataSource = Nothing
        DataGrid1.Refresh
        DataGrid1.Height = 705
        dtpFechaini.Enabled = True
        dtpFechacul.Enabled = True
        dtpFechaini.Value = Date
        dtpFechacul.Value = Date
        dtpFechacul.MinDate = Date
        cmdMostrar.Caption = "Mostrar"
        cmdMostrar.Picture = LoadPicture(App.Path & "\imágenes\Reporte.jpg")
        cmdImprimir.Enabled = False
    ElseIf cmdMostrar.Caption = "Mostrar" Then
        Set rs = New ADODB.Recordset
           
        sql = Listado()
        rs.Open sql, db, adOpenStatic
    
        If rs.EOF Then
            MsgBox "No hay datos para mostrar", vbExclamation
            Set DataGrid1.DataSource = rs
            cmdImprimir.Enabled = False
            Exit Sub
        Else
            Set DataGrid1.DataSource = rs
            dtpFechaini.Enabled = False
            dtpFechacul.Enabled = False
            cmdMostrar.Caption = "Limpiar"
            cmdMostrar.Picture = LoadPicture(App.Path & "\imágenes\Cancelar2.jpg")
            DataGrid1.Height = DataGrid1.Height + (rs.RecordCount - 1) * 345 'Para que se autoajuste a la cantidad de filas.
            cmdImprimir.Enabled = True
            Exit Sub
        End If
    
    End If
End Sub

Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm AM/PM")
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub dtpFechaini_Change()
dtpFechacul.MinDate = dtpFechaini.Value
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
dtpFechacul.Value = Date
dtpFechaini.Value = Date
dtpFechacul.MinDate = dtpFechaini.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPrincipal.Show
End Sub

