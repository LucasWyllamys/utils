VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: clsDatabaseConnection - Conexão com Banco de Dados via ADODB
' Descrição: Esta classe permite realizar conexões com bancos de dados utilizando a biblioteca ADODB.
'   - Suporte aos bancos: MySQL, SQL Server e Access.
'   - Montagem automática da string de conexão.
'   - Execução de comandos SQL: SELECT, INSERT, UPDATE, DELETE.
'   - Gerenciamento de objetos ADODB.Connection, ADODB.Command e ADODB.Recordset.
' Versão: 2.1.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 29/10/2025
' Atualizado em: 30/10/2025
' Pré-requisitos:
'   - Ativar biblioteca:
'       - Microsoft ActiveX Data Objects 6.1 Library: No Editor VBA: Ferramentas > Referências > Microsoft ActiveX Data Objects 6.1 Library
'   - Importar os módulos:
'       - PublicConstants: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
' Observações:
'   - A propriedade `rs` armazena o resultado de consultas SELECT.
'   - A propriedade `cmd` é usada internamente para executar comandos SQL.
'   - A propriedade `conn` representa a conexão ativa com o banco de dados.
'   - A string de conexão é montada automaticamente com base no tipo de banco informado.
' Exemplo de uso:
'   Dim db As clsDatabaseConnection
'   Set db = New clsDatabaseConnection
'
'   db.Initialize tipoDeBanco, "database", "server" , "usuario", "senha", "provider"
'
'   If db.OpenDb Then
'       db.ExecuteNonQuery "UPDATE Tabela SET Nome = 'Lucas' WHERE ID = 1"
'       db.ExecuteQuery "SELECT * FROM Tabela"
'       Do While Not db.rs.EOF
'           Debug.Print db.rs.Fields("Nome").Value
'           db.rs.MoveNext
'       Loop
'       db.CloseConn
'   End If
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================

Private pConn As ADODB.Connection
Private pRs As ADODB.Recordset
Private pCmd As ADODB.Command

Private pServer As String
Private pDatabase As String
Private pUserId As String
Private pPassword As String
Private pProvider As String
Private pDriver As String
Private pConnStr As String

'========================================= Enum Tipo de Banco ========================================

Public Enum DatabaseType
    dbMySQL = 1
    dbSQLServer = 2
    dbAccess = 3
End Enum

'========================================= Métodos Assessores ========================================

Private Property Get conn() As ADODB.Connection: Set conn = pConn: End Property

Private Property Set conn(object As ADODB.Connection): Set pConn = object: End Property


Public Property Get rs() As ADODB.Recordset: Set rs = pRs: End Property

Public Property Set rs(object As ADODB.Recordset): Set pRs = object: End Property


Private Property Get cmd() As ADODB.Command: Set cmd = pCmd: End Property

Private Property Set cmd(object As ADODB.Command): Set pCmd = object: End Property


Private Property Get connStr() As String: connStr = pConnStr: End Property

Private Property Let connStr(value As String): pConnStr = value: End Property

'=========================================== Inicialização ===========================================

Public Sub Initialize(ByVal dbType As DatabaseType, _
                        ByVal strDatabase As String, _
                        Optional ByVal strServer As String = "", _
                        Optional ByVal strUserId As String = "", _
                        Optional ByVal strPassword As String = "", _
                        Optional ByVal strProvider As String = "", _
                        Optional ByVal strDriver As String = "")
    
    pServer = strServer
    pDatabase = strDatabase
    pUserId = strUserId
    pPassword = strPassword
    pProvider = strProvider
    pDriver = strDriver
    
    AssembleConnectionString dbType
End Sub

'=========================================== Métodos Públicos ===========================================

' Abre a conexão com o banco
Public Function OpenConn() As Boolean
    On Error GoTo ErrorHandler

    Set conn = New ADODB.Connection
    conn.ConnectionTimeout = 30 ' Tempo limite da conexão
    conn.Open connStr
    
    OpenConn = True
    Exit Function
ErrorHandler:
    Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Falha ao conectar ao banco de dados (" & Err.Number & "): " & Err.Description
    OpenConn = False
End Function

' Verifica se o banco de dados está conectado
Public Function IsConnected() As Boolean
    IsConnected = Not conn Is Nothing And conn.State = adStateOpen
End Function

' Use os comando SELECT para retornar uma Recordset na propriedade rs
Public Function ExecuteQuery(ByVal sql As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão NÃO está ativa (" & Err.Number & "): " & Err.Description
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 60 ' Tempo limite para execução do SELECT
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = sql
        Set rs = .Execute
    End With
    
    ExecuteQuery = True
    Exit Function
ErrorHandler:
    Err.Raise ERR_DB_QUERY_FAILED, "clsDatabaseConnection", "Erro ao executar SELECT (" & Err.Number & "): " & Err.Description
    ExecuteQuery = False
End Function

' Use os comandos INSERT, UPDATE, DELETE, retorna o ID do registro (exceto para o DELETE)
Public Sub ExecuteNonQuery(ByVal sql As String)
    On Error GoTo ErrorHandler
    
    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão NÃO está ativa: " & Err.Description
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 60 ' Tempo limite para INSERT/UPDATE/DELETE
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = sql
    End With
    
    cmd.Execute
    Exit Sub
ErrorHandler:
    Err.Raise ERR_DB_QUERY_FAILED, "clsDatabaseConnection", "Erro ao executar comando SQL (" & Err.Number & "): " & Err.Description
End Sub

' Retorna o ID do ultimo registro inserido
Public Function GetLastInsertId() As Long
    On Error GoTo ErrorHandler
    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão não está ativa."

    If ExecuteQuery("SELECT @@IDENTITY") Then
        If Not rs Is Nothing And Not rs.EOF Then
            GetLastInsertId = rs.fields(0).value
        End If
    End If
    Exit Function
ErrorHandler:
    GetLastInsertId = -1
End Function

' Fecha a conexão com a recordset e o banco de dados
Public Sub CloseConn()
    Set cmd = Nothing ' Esvazia o objeto cmd
    CloseRs
    CloseDb
End Sub

'=========================================== Métodos Privados ===========================================

' Monta a string de conexão
Private Sub AssembleConnectionString(dbType As DatabaseType)
    Select Case dbType
        Case dbMySQL
            connStr = Join(Array( _
                "Driver={" & pDriver & "}", _
                "Server=" & pServer, _
                "Database=" & pDatabase, _
                "User=" & pUserId, _
                "Password=" & pPassword, _
                "Option=3"), ";")

        Case dbSQLServer
            connStr = Join(Array( _
                "Provider=" & pProvider, _
                "Data Source=" & pServer, _
                "Initial Catalog=" & pDatabase, _
                "User ID=" & pUserId, _
                "Password=" & pPassword), ";")

        Case dbAccess
            connStr = Join(Array( _
                "Provider=" & pProvider, _
                "Data Source=" & pDatabase), ";")
                
        Case Else
            Err.Raise ERR_DB_INVALID_CONNECTION_STRING, "clsDatabaseConnection", "Tipo de banco não suportado: " & dbType
    End Select
End Sub

' Fecha a conexão com a recordset
Private Sub CloseRs()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

' Fecha a conexão com o database
Private Sub CloseDb()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub
