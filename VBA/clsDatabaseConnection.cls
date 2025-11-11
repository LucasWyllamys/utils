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
' Versão: 2.4.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 29/10/2025
' Atualizado em: 11/11/2025
'
' Descrição:
'   Esta classe encapsula a lógica de conexão e execução de comandos SQL em bancos de dados utilizando a
'   biblioteca ADODB. Ideal para aplicações VBA que exigem integração com bancos externos.
'
' Recursos:
'   - Suporte aos bancos: MySQL, SQL Server e Access.
'   - Montagem automática da string de conexão com validação dos campos obrigatórios.
'   - Execução de comandos SQL: SELECT, INSERT, UPDATE, DELETE.
'   - Retorno do ID do último registro inserido (SQL Server).
'   - Gerenciamento seguro dos objetos ADODB.Connection, ADODB.Command e ADODB.Recordset.
'   - Métodos para verificar conexão ativa e fechar recursos.
'
' Pré-requisitos:
'   - Ativar biblioteca:
'       - Microsoft ActiveX Data Objects 6.1 Library
'         No Editor VBA: Ferramentas > Referências > Microsoft ActiveX Data Objects 6.1 Library
'
'   - Importar os módulos auxiliares:
'       - PublicConstants: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
'
' Observações:
'   - A propriedade `rs` armazena o resultado de consultas SELECT.
'   - A propriedade `cmd` é usada internamente para executar comandos SQL.
'   - A propriedade `conn` representa a conexão ativa com o banco de dados.
'   - A propriedade `connStr` contém a string de conexão montada automaticamente.
'   - A validação dos campos obrigatórios é feita de forma centralizada.
'
' Exemplo de uso:
'   Dim db As clsDatabaseConnection
'   Set db = New clsDatabaseConnection
'
'   db.Initialize dbSQLServer, _
'       strDataSource:="", _
'       strDatabase:="MinhaBase", _
'       strServer:="localhost", _
'       strUserId:="usuario", _
'       strPassword:="senha", _
'       strProvider:="SQLOLEDB"
'
'   If db.OpenConn Then
'       db.ExecuteNonQuery "UPDATE Tabela SET Nome = 'Lucas' WHERE ID = 1"
'       db.ExecuteQuery "SELECT * FROM Tabela"
'
'       Do While Not db.rs.EOF
'           Debug.Print db.rs.Fields("Nome").Value
'           db.rs.MoveNext
'       Loop
'
'       db.CloseConn
'   End If
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================

Private pConn As ADODB.Connection
Private pRs As ADODB.Recordset
Private pCmd As ADODB.Command

Private pServer As String
Private pDataSource As String
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
    dbSAPHANA = 4
End Enum

'========================================= Métodos Assessores ========================================

Private Property Get conn() As ADODB.Connection: Set conn = pConn: End Property

Private Property Set conn(object As ADODB.Connection): Set pConn = object: End Property


Friend Property Get rs() As ADODB.Recordset: Set rs = pRs: End Property

Friend Property Set rs(object As ADODB.Recordset): Set pRs = object: End Property


Private Property Get cmd() As ADODB.Command: Set cmd = pCmd: End Property

Private Property Set cmd(object As ADODB.Command): Set pCmd = object: End Property


Friend Property Get connStr() As String: connStr = pConnStr: End Property

Private Property Let connStr(Value As String): pConnStr = Value: End Property

'=========================================== Inicialização ===========================================

Friend Sub Initialize(ByVal dbType As DatabaseType, _
                        Optional ByVal strDataSource As String = "", _
                        Optional ByVal strDatabase As String = "", _
                        Optional ByVal strServer As String = "", _
                        Optional ByVal strUserId As String = "", _
                        Optional ByVal strPassword As String = "", _
                        Optional ByVal strProvider As String = "", _
                        Optional ByVal strDriver As String = "")
    
    pServer = strServer
    pDataSource = strDataSource
    pDatabase = strDatabase
    pUserId = strUserId
    pPassword = strPassword
    pProvider = strProvider
    pDriver = strDriver
    
    AssembleConnectionString dbType
End Sub

Private Sub Class_Terminate()
    CloseConn
End Sub

'=========================================== Métodos Públicos ===========================================

' Abre a conexão com o banco
Friend Function OpenConn() As Boolean
    'On Error GoTo ErrorHandler

    Set conn = New ADODB.Connection
    conn.ConnectionTimeout = 30 ' Tempo limite da conexão
    conn.Open connStr
    
    OpenConn = True
    Exit Function
ErrorHandler:
    Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Falha ao conectar ao banco de dados: " & connStr
    OpenConn = False
End Function

' Verifica se o banco de dados está conectado
Friend Function IsConnected() As Boolean
    IsConnected = Not conn Is Nothing And conn.State = adStateOpen
End Function

' Use os comando SELECT para retornar uma Recordset na propriedade rs
Friend Function ExecuteQuery(ByVal sql As String) As Boolean

    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão NÃO está ativa."
    
    On Error GoTo ErrorHandler
    
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
    Err.Raise ERR_DB_QUERY_FAILED, "clsDatabaseConnection", "Erro ao executar comando: " & sql
    ExecuteQuery = False
End Function

' Use os comandos INSERT, UPDATE, DELETE, retorna o ID do registro (exceto para o DELETE)
Friend Sub ExecuteNonQuery(ByVal sql As String)
    On Error GoTo ErrorHandler
    
    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão NÃO está ativa."
    
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
    Err.Raise ERR_DB_QUERY_FAILED, "clsDatabaseConnection", "Erro ao executar comando SQL: " & sql
End Sub

' Retorna o ID do ultimo registro inserido
Friend Function GetLastInsertId() As Long
    On Error GoTo ErrorHandler
    If Not IsConnected Then Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Conexão não está ativa."

    If ExecuteQuery("SELECT SCOPE_IDENTITY()") Then
        If Not rs Is Nothing And Not rs.EOF Then
            GetLastInsertId = rs.fields(0).Value
        End If
    End If
    Exit Function
ErrorHandler:
    GetLastInsertId = -1
End Function

' Fecha a conexão com a recordset e o banco de dados
Friend Sub CloseConn()
    CloseCmd
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
            ValidateRequiredFields Array(pDriver, pServer, pDatabase, pUserId, pPassword)
        Case dbSQLServer
            connStr = Join(Array( _
                "Provider=" & pProvider, _
                "Data Source=" & pServer, _
                "Initial Catalog=" & pDatabase, _
                "User ID=" & pUserId, _
                "Password=" & pPassword), ";")
            ValidateRequiredFields Array(pProvider, pServer, pDatabase, pUserId, pPassword)
        Case dbAccess
            connStr = Join(Array( _
                "Provider=" & pProvider, _
                "Data Source=" & pDataSource), ";")
            ValidateRequiredFields Array(pProvider, pDataSource)
        Case dbSAPHANA
            connStr = Join(Array( _
                "Driver={" & pDriver & "}", _
                "ServerNode=" & pServer, _
                "DatabaseName=" & pDatabase, _
                "UID=" & pUserId, _
                "PWD=" & pPassword, _
                "SSLValidateCertificate=No"), ";")
            ValidateRequiredFields Array(pDriver, pServer, pDatabase, pUserId, pPassword)
        Case Else
            Err.Raise ERR_DB_INVALID_CONNECTION_STRING, "clsDatabaseConnection", "Tipo de banco não suportado: " & dbType
    End Select
End Sub

Private Sub ValidateRequiredFields(ByVal fields As Variant)
    Dim i As Integer
    
    For i = LBound(fields) To UBound(fields)
        If Trim(fields(i)) = "" Then
            Err.Raise ERR_DB_INVALID_CONNECTION_STRING, "clsDatabaseConnection", "String de conexão inválida: " & connStr
        End If
    Next i
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

' Fecha a conexão com o Command
Private Sub CloseCmd()
    If Not cmd Is Nothing Then Set cmd = Nothing
End Sub
