VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{9156C8F9-B397-4DEF-8AC5-5966221A134A}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmCulminado 
   Caption         =   "Culminación de Gupros"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCulminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   13440
      Top             =   2040
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
      Left            =   6720
      Picture         =   "frmCulminado.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1700
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H8000000A&
      Caption         =   "Imprimir"
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
      Picture         =   "frmCulminado.frx":1E93
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1700
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   705
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   13140
      _ExtentX        =   23178
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Grupo"
         Caption         =   "Grupo"
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
         DataField       =   "Fecha_Inicio"
         Caption         =   "Fecha de Inicio"
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
         DataField       =   "Fecha_Culminación"
         Caption         =   "Fecha de Culminación"
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
         DataField       =   "Turno"
         Caption         =   "Turno"
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
      BeginProperty Column05 
         DataField       =   "Días_Culminado"
         Caption         =   "Días Culminado"
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
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2894,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2700,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2700,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995,024
         EndProperty
      EndProperty
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   630
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   1111
      Image           =   "frmCulminado.frx":3DCB
      Props           =   5
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   12240
      Picture         =   "frmCulminado.frx":45EB
      Top             =   480
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
      Left            =   13320
      TabIndex        =   4
      Top             =   1560
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
      Left            =   13320
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   6675
      Left            =   0
      Picture         =   "frmCulminado.frx":7C6B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9315
   End
End
Attribute VB_Name = "frmCulminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimir_Click()
    Set rs = New ADODB.Recordset
    sql = "select gcodigo as Grupo, cnombre as Curso, gfechaini as Fecha_Inicio, gfechacul as Fecha_Culminación, tturno as Turno, decode(to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy'),1,to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy')||' Día',to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy')||' Días') as Días_Culminado from tgruposculminados,tcursos,tturnos where gestatus='C' and gcodigoc=ccodigo and gcodturno=tcodigo order by gfechacul desc"
    rs.Open sql, db, adOpenStatic
    
    Set reportGruposCulminados.DataSource = rs
    'Para indiciar en el pie de página el usuario que imprimió el reporte.
    reportGruposCulminados.Sections("Sección3").Controls.Item("Etiqueta12").Caption = "Impreso por: " & nombreUser & " " & apellidoUser
    reportGruposCulminados.Show
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()

End Sub
Private Sub Timer1_Timer()
lblFecha.Caption = Date
lblHora.Caption = Format(Time, "hh:mm AM/PM")
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

    Set rs = New ADODB.Recordset
    'sql = "select cnombre  as Curso, gcodigo as Grupo, gfechaini as Fecha_Inicio,gfechacul as Fecha_Culminación,trunc(sysdate - gfechacul) as Dias_de_Culminado from TGruposCulminados,TCursos where gcodigoc=ccodigo and gfechacul < sysdate and trunc(sysdate - gfechacul)<= 7 and gestatus = 'C' order by gfechacul desc"
    sql = "select gcodigo as Grupo, cnombre as Curso, gfechaini as Fecha_Inicio, gfechacul as Fecha_Culminación, tturno as Turno, decode(to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy'),1,to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy')||' Día',to_date(sysdate,'dd/mm/yyyy')-to_date(gfechacul,'dd/mm/yyyy')||' Días') as Días_Culminado from tgruposculminados,tcursos,tturnos where gestatus='C' and gcodigoc=ccodigo and gcodturno=tcodigo order by gfechacul desc"
    rs.Open sql, db, adOpenStatic
    
    If rs.EOF Then
        cmdImprimir.Enabled = False
    Else
        Set DataGrid2.DataSource = rs
        DataGrid2.Height = DataGrid2.Height + (rs.RecordCount - 1) * 345 'Para que se autoajuste a la cantidad de filas.
        cmdImprimir.Enabled = True
    End If
    
    If (rs.RecordCount > 12) Then
        DataGrid2.Height = 705 + (12 - 1) * 345
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPrincipal.Show
End Sub

