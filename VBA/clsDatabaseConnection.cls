VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================== Propriedades Privadas ========================================

Private pConn As ADODB.Connection
Private pRs As ADODB.Recordset

Private pServer As String
Private pDatabase As String
Private pUserId As String
Private pPassword As String
Private pProvider As String
Private pDriver As String
Private pConnStr As String

'========================================= Métodos Assessores ========================================

Private Property Get conn() As ADODB.Connection: Set conn = pConn: End Property

Private Property Set conn(object As ADODB.Connection): Set pConn = object: End Property

Public Property Get rs() As ADODB.Recordset: Set rs = pRs: End Property

Public Property Set rs(object As ADODB.Recordset): Set pRs = object: End Property

Private Property Get server() As String: server = pServer: End Property

Private Property Let server(value As String): pServer = value: End Property

Private Property Get database() As String: database = pDatabase: End Property

Private Property Let database(value As String): pDatabase = value: End Property

Private Property Get userId() As String: userId = pUserId: End Property

Private Property Let userId(value As String): pUserId = value: End Property

Private Property Get password() As String: password = pPassword: End Property

Private Property Let password(value As String): pPassword = value: End Property

Private Property Get provider() As String: provider = pProvider: End Property

Private Property Let provider(value As String): pProvider = value: End Property

Private Property Get driver() As String: driver = pDriver: End Property

Private Property Let driver(value As String): pDriver = value: End Property

Private Property Get connStr() As String: connStr = pConnStr: End Property

Private Property Let connStr(value As String): pConnStr = value: End Property

'=========================================== Inicialização ===========================================

Public Sub Initialize(ByVal tipoBanco As String, _
                        ByVal strServer As String, _
                        ByVal strDatabase As String, _
                        ByVal strUserId As String, _
                        ByVal strPassword As String, _
                        Optional ByVal strProvider As String = "", _
                        Optional ByVal strDriver As String = "")
    
    server = strServer
    database = strDatabase
    userId = strUserId
    password = strPassword
    provider = strProvider
    driver = strDriver
    
    AssembleConnectionString tipoBanco
End Sub

'=========================================== Métodos Públicos ===========================================

' Abre a conexão
Public Function OpenDb() As Boolean
    On Error GoTo ErrorHandler

    Set conn = New ADODB.Connection
    conn.Open connStr
    
    OpenDb = True
    Exit Function
ErrorHandler:
   Err.Raise ERR_DB_CONNECTION_FAILED, "clsDatabaseConnection", "Falha ao conectar ao banco de dados (" & Err.Number & "): " & Err.Description
    OpenDb = False
End Function

' Executa SELECT e retorna Recordset
Public Function ExecuteSql(ByVal sql As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly '(tabela, conexão ativa, tipo de cursor, tipo de bloqueio)
    Set ExecuteSql = rs
    
    Exit Function
ErrorHandler:
    Err.Raise ERR_DB_QUERY_FAILED, "clsDatabaseConnection", "Falha ao executar a consulta SQL (" & Err.Number & "): " & Err.Description
End Function

' Fecha a conexão com a recordset
Public Sub CloseRs()
    If Not rs Is Nothing Then ' Verifica se o objeto está instanciado
        If rs.State = adStateOpen Then 'Verifica se a recordset está aberta
            rs.Close
        End If
        Set rs = Nothing
    End If
End Sub

' Fecha a conexão com a recordset e o banco de dados
Public Sub CloseDb()
    ' Fecha a conexão com a recordset
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    ' Fecha a conexão com o database
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub

'=========================================== Métodos Privados ===========================================

' Monta a string de conexão
Private Sub AssembleConnectionString(tipoBanco As String)
    tipoBanco = LCase(tipoBanco)
    
    Select Case tipoBanco
        Case "mysql"
            connStr = _
            "Driver={" & driver & "};" & _
            "Server=" & server & ";" & _
            "Database=" & database & ";" & _
            "User=" & userId & ";" & _
            "Password=" & password & ";" & _
            "Option=3;"
        Case "sqlserver"
            connStr = _
            "Provider=" & provider & ";" & _
            "Data Source=" & server & ";" & _
            "Initial Catalog=" & database & ";" & _
            "User ID=" & userId & ";" & _
            "Password=" & password & ";"
        Case "access"
            connStr = _
            "Provider=" & provider & ";" & _
            "Data Source=" & database & ";"
        Case Else
            Err.Raise ERR_DB_INVALID_CONNECTION_STRING, "clsDatabaseConnection", "Tipo de banco não suportado (" & Err.Number & "): " & Err.Description
    End Select
End Sub
