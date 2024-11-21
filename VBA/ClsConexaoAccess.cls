VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsConexaoAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Habilitar Biblioteca: Microsoft ActiveX Data Object 6.1 Library
' Habilitar Biblioteca: Microsoft Scripting Runtime

Private Type ClassType
    conexaoDB           As ADODB.Connection     ' Conexão com o banco de dados
    rs                  As ADODB.recordSet      ' Tabela virtual
    caminhoDB           As String
End Type

Private self As ClassType

Private Sub Class_Initialize()
    caminhoDB = ThisWorkbook.Path & "\stgmusic_db.accdb"
End Sub

' Métodos assessores:------------------------------------------------------

Private Property Get caminhoDB() As String
    Set caminhoDB = self.caminhoDB
End Property

Private Property Set caminhoDB(value As String)
    Set self.caminhoDB = value
End Property

Private Property Get conexaoDB() As ADODB.Connection
    Set conexaoDB = self.conexaoDB
End Property

Private Property Set conexaoDB(obj As ADODB.Connection)
    Set self.conexaoDB = obj
End Property

Private Property Get rs() As ADODB.recordSet
    Set rs = self.rs
End Property

Private Property Set rs(obj As ADODB.recordSet)
    Set self.rs = obj
End Property

'Métodos:-------------------------------------------------------------------

Public Function ConectarDB()
    Dim Provider As String, DataSource As String, ConnectionString As String
    If conexaoDB Is Nothing Then                        ' Verifica se a variável existe
        If conexaoDB.State <> 1 Then                    ' Verifica se o banco de dados está aberto
            Set conexaoDB = New ADODB.Connection        ' Cria o objeto de conexão com o Access.
            Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
            DataSource = "Data Source=" & caminhoDB
            ConnectionString = Provider & DataSource    ' Driver de concexão com o Access.
            conexaoDB.Open ConnectionString             ' Abri a concexão com o banco de dados Access.
        End If
    End If
End Function

Public Function DesconectarRS()
    If Not rs Is Nothing Then       ' Verifica se a variável existe
        If rs.State = 1 Then        ' Verifica se o banco de dados está aberto
            rs.Close
        End If
    End If
End Function

Public Function DesconectarDB()
    If Not conexaoDB Is Nothing Then
        If conexaoDB.State = 1 Then
            conexaoDB.Close
        End If
    End If
End Function

Public Function InserirRegistro(tabela As String, dictRegistro As Scripting.Dictionary)
    Call DesconectarRS
    
    Set rs = New ADODB.recordSet
    rs.Open tabela, conexaoDB, adOpenKeyset, adLockOptimistic   ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio). Abre a conexão com a RecordSet.
    rs.AddNew                                                   ' Adiciona um novo registro à RecordSet
    
    Dim campo As Variant
    For Each campo In dictRegistro
        rs(campo) = dictRegistro(campo)
    Next campo
    
    Call UpdateRS
End Function

Public Function ConsultarRegistro(sql As String, Optional planilha As Worksheet, Optional celula As String) As Boolean
    Call DesconectarRS
    Set rs = New ADODB.recordSet
    rs.Open sql, conexaoDB, adOpenKeyset, adLockReadOnly    ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio)
    If rs.BOF = True And rs.EOF = True Then                 ' Retorna True or False
        ConsultarRegistro = False
    Else
        ConsultarRegistro = True
        If Not planilha Is Nothing And celula <> "" Then
            planilha.Range(celula).CopyFromRecordset rs
        End If
    End If
End Function

Public Function EditarRegistro(sql As String)
    Call DesconectarRS
    Set rs = New ADODB.recordSet
    rs.Open sql, conexaoDB, adOpenKeyset, adLockOptimistic  ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio).
End Function

Function DeletarRegistro(sql As String)
    Call DesconectarRS
    Set rs = New ADODB.recordSet
    rs.Open sql, conexaoDB, adOpenKeyset, adLockOptimistic  ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio).
    rs.Delete                                               ' Deletar registro
End Function

Function UpdateRS()
    rs.Update   'Atualiza a tabela
End Function
