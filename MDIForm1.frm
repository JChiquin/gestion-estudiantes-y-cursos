VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9345
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Menu mMenu 
      Caption         =   "Menú"
      Begin VB.Menu mCursos 
         Caption         =   "Cursos"
      End
      Begin VB.Menu mGrupos 
         Caption         =   "Grupos"
      End
      Begin VB.Menu mInstructores 
         Caption         =   "Instructores"
      End
      Begin VB.Menu mUsuarios 
         Caption         =   "Administrador de Usuarios"
      End
   End
   Begin VB.Menu mReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mCulminado 
         Caption         =   "Grupos Culminados"
      End
      Begin VB.Menu mListado 
         Caption         =   "Listado de Cursos"
      End
   End
   Begin VB.Menu mSalir 
      Caption         =   "Salir"
   End
   Begin VB.Menu Prueba 
      Caption         =   "Prueba"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mCursos_Click()
    frmCursos.Show
End Sub

Private Sub mGrupos_Click()
frmGrupos.Show
End Sub

Private Sub mCulminado_Click()
frmCulminado.Show
End Sub

Private Sub mInstructores_Click()
frmInstructor.Show
End Sub

Private Sub mListado_Click()
frmListadoCursos.Show
End Sub

Private Sub MDIForm_Load()
    Set db = New ADODB.Connection
    db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=conexionOracle"
    SOpt = dbSQLPassThrough
    Me.Hide
    frmLogin.Show

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
resp = MsgBox(" ¿Está seguro que desea SALIR? ", vbInformation + vbYesNo)
If (resp = vbNo) Then
    Cancel = 1
ElseIf (resp = vbYes) Then
    frmLogin.Show
End If
End Sub

Private Sub mUsuarios_Click()
frmUsuario.Show
End Sub

Private Sub mSalir_Click()
Unload Me
End Sub

Private Sub Prueba_Click()
frmUsuarioTipo2.Show
End Sub
