VERSION 5.00
Begin VB.Form frmPresentacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPresentacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPresentacion.frx":058A
   ScaleHeight     =   5340
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   150
      Left            =   2760
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   1080
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   725
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   7080
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   715
         Left            =   0
         Top             =   0
         Width           =   7065
      End
      Begin VB.Label lblPorcentaje 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   150
      End
      Begin VB.Label lblCargar2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   75
      End
      Begin VB.Label lblCargar1 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Height          =   675
         Left            =   0
         TabIndex        =   1
         Top             =   20
         Width           =   15
      End
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compañía"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5280
      TabIndex        =   12
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5280
      TabIndex        =   11
      Top             =   3480
      Width           =   690
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advertencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   885
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plataforma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Infinity And Beyond C.A."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   4140
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto de la compañía"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
      Width           =   930
   End
   Begin VB.Image imgLogo 
      Height          =   2385
      Left            =   840
      Picture         =   "frmPresentacion.frx":5274
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblCargar3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   45
   End
End
Attribute VB_Name = "frmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
    'Marchoso1.FileName = App.Path & "\imágenes\Binario.gif"
    
End Sub

Private Sub lblCargar2_Change()
If lblCargar2.Caption = 100 Then
lblPorcentaje.Left = lblPorcentaje.Left + 70
End If
End Sub



Private Sub Timer1_Timer()
Static m As Integer
m = m + 1
lblCargar2.Caption = m
If m = 100 Then
Timer1.Enabled = False

End If
End Sub

Private Sub Timer2_Timer()
lblCargar1.Width = lblCargar1.Width + 26
If (lblCargar1.Width > 7080) Then
Timer2.Enabled = False
Unload Me
frmLogin.Show
End If

End Sub

Private Sub Timer3_Timer()
Static c As Integer
c = c + 1
If lblCargar2.Caption = "100" Then
lblCargar3.Caption = "Descarga Completa"
Else

If c = 1 Then
lblCargar3.Caption = "Descargando Paquetes"
ElseIf c = 2 Then
lblCargar3.Caption = "Librerias //lbl"
ElseIf c = 3 Then
lblCargar3.Caption = "Librerias //msc"
ElseIf c = 4 Then
lblCargar3.Caption = "Librerias //zxc"
ElseIf c = 5 Then
lblCargar3.Caption = "Librerias //er33"
ElseIf c = 6 Then
lblCargar3.Caption = "Librerias //uuy56"
ElseIf c = 7 Then
lblCargar3.Caption = "Librerias //xc32"
ElseIf c = 7 Then
lblCargar3.Caption = "Librerias //xc32"
ElseIf c = 8 Then
lblCargar3.Caption = "Abriendo Formularios"
ElseIf c = 9 Then
lblCargar3.Caption = "FrmIntructores"
ElseIf c = 10 Then
lblCargar3.Caption = "FrmGrupos"
ElseIf c = 11 Then
lblCargar3.Caption = "Actualizando Datos"
ElseIf c = 12 Then
lblCargar3.Caption = "Archivos Fuentes"
ElseIf c = 13 Then
lblCargar3.Caption = "Archivos de Sistema"
ElseIf c = 14 Then
lblCargar3.Caption = "Archivo //e3e3"
ElseIf c = 15 Then
lblCargar3.Caption = "Archivo //23vb"

Else: c = 16
c = 0
End If
End If
End Sub
