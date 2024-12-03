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
    rs                  As ADODB.Recordset      ' Tabela virtual
    caminhoDB           As String
End Type

Private self As ClassType

Private Sub Class_Initialize()
    caminhoDB = ThisWorkbook.Path & "\stgmusic_db.accdb"
End Sub

' Métodos assessores:------------------------------------------------------

Private Property Get caminhoDB() As String
    caminhoDB = self.caminhoDB
End Property

Private Property Let caminhoDB(valor As String)
    self.caminhoDB = valor
End Property

Private Property Get conexaoDB() As ADODB.Connection
    Set conexaoDB = self.conexaoDB
End Property

Private Property Set conexaoDB(obj As ADODB.Connection)
    Set self.conexaoDB = obj
End Property

Public Property Get rs() As ADODB.Recordset
    Set rs = self.rs
End Property

Private Property Set rs(obj As ADODB.Recordset)
    Set self.rs = obj
End Property

'Métodos:-------------------------------------------------------------------

Public Function ConectarDB()
    Dim Provider As String, DataSource As String, ConnectionString As String
    Set conexaoDB = New ADODB.Connection        ' Cria o objeto de conexão com o Access.
    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
    DataSource = "Data Source=" & caminhoDB
    ConnectionString = Provider & DataSource    ' Driver de concexão com o Access.
    conexaoDB.Open ConnectionString             ' Abri a concexão com o banco de dados Access.
End Function

Public Function DesconectarDB()
    If Not conexaoDB Is Nothing Then    ' Verifica se a variável existe.
        If conexaoDB.State = 1 Then     ' Verifica se o banco de dados está aberto.
            conexaoDB.Close             ' Fecha a conexão com o banco.
        End If
    End If
End Function

Public Function ConectarRS()
    Set rs = New ADODB.Recordset
End Function

Public Function DesconectarRS()
    If Not rs Is Nothing Then       ' Verifica se a variável existe
        If rs.State = 1 Then        ' Verifica se o banco de dados está aberto
            rs.Close                ' Fecha a conexão com a RecordSet.
        End If
    End If
End Function

Public Function InserirRegistro(tabela As String, dictRegistro As Scripting.Dictionary)
    Dim campo As Variant
    
    rs.Open tabela, conexaoDB, adOpenKeyset, adLockOptimistic   ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio). Abre a conexão com a RecordSet.
    rs.AddNew                                                   ' Adiciona um novo registro à RecordSet

    For Each campo In dictRegistro
        rs(campo) = dictRegistro(campo)
    Next campo
End Function

Public Function ConsultarRegistro(sql As String, Optional planilha As Worksheet, Optional celulaRef As String) As Boolean
    rs.Open sql, conexaoDB, adOpenKeyset, adLockReadOnly    ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio)
    If rs.BOF = True And rs.EOF = True Then                 ' Retorna True or False
        ConsultarRegistro = False
    Else
        ConsultarRegistro = True
        If Not planilha Is Nothing And celulaRef <> "" Then
            planilha.Range(celulaRef).CopyFromRecordset rs
        End If
    End If
End Function

Public Function DBParaForm(dictRegistro As Scripting.Dictionary, pageForm As MSForms.Page)
    Dim campoDB As Variant, campoForm As MSForms.Control
    For Each campoDB In dictRegistro
        For Each campoForm In pageForm.Controls
            If dictRegistro(campoDB) = campoForm.Name Then
                If TypeName(campoForm) = "TextBox" Or TypeName(campoForm) = "ComboBox" Then
                    campoForm.value = rs(campoDB)
                ElseIf TypeName(campoForm) = "CommandButton" And InStr(campoForm.Name, "CaminhoAnexo") > 0 Then
                    campoForm.Caption = rs(campoDB)
                End If
            End If
        Next campoForm
    Next campoDB
End Function

Public Function EditarRegistro(tabela As String, dictRegistro As Scripting.Dictionary)
    Dim sql As String, campo As Variant
    
    sql = "SELECT * FROM " & tabela & " WHERE ID =" & dictRegistro("ID")
    rs.Open sql, conexaoDB, adOpenKeyset, adLockOptimistic  ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio).

    For Each campo In dictRegistro
        If dictRegistro(campo) <> Empty And campo <> "ID" Then
            rs(campo) = dictRegistro(campo)
        End If
    Next campo
End Function

Public Function DeletarRegistro(sql As String)
    rs.Open sql, conexaoDB, adOpenKeyset, adLockOptimistic  ' (tabela, conexão ativa, tipo de cursor, tipo de bloqueio).
    rs.Delete                                               ' Deletar registro
End Function

Public Function UpdateRS()
    rs.Update   'Atualiza a tabela
End Function
