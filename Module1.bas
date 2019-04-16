Attribute VB_Name = "Module1"
Public db As ADODB.Connection
Public rs As ADODB.Recordset
Public as1 As ADODB.Recordset

Public SOpt As Variant
Public cm As ADODB.Command

Public sql As String

Public nombreUser As String 'Variable que almacena el primer nombre del usuario que inici� la sesi�n.
Public apellidoUser As String 'Variable que almacena el primer apellido del usuario que inici� la sesi�n.


'variables para instructores
Public isexoLetra As String
Public isexo As String
Public Sub MoverGruposCulminados() 'Saca los grupos culminados de TGrupos a TGruposculminados
'La tabla TGruposculminados es exactamente igual a tgrupos.
sql = "insert into tgruposculminados select *from tgrupos where gestatus='C'" 'incluye en tgruposculminados los grupos culminados de TGrupos.
db.Execute sql, SOpt
sql = "delete tgrupos where gestatus='C'" 'Elimina los grupos culminados de TGrupos.
db.Execute sql, SOpt
End Sub
Public Sub CulminarGrupos() 'Esto actualizar� el estatus de los grupos.
'si la fecha de culminaci�n ya pas� (menor a fecha de hoy) el estatus ser� C
sql = "update TGrupos set gestatus='C' where to_date(gfechacul,'dd/mm/yyyy') < to_date(sysdate,'dd/mm/yyyy') and gestatus='A'"
'est� en menor estricto, porque aunque el grupo culmine hoy en realidad lo har� al final del d�a. Hay que preguntar esto.
db.Execute sql, SOpt
MoverGruposCulminados 'Ni bien culmine un grupo, que lo saque de TGrupos a TGruposculminados
End Sub
Public Sub Conexion()
    Set db = New ADODB.Connection
    db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=conexionOracle"
    SOpt = dbSQLPassThrough
End Sub





