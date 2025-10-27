Attribute VB_Name = "PublicConstants"
Option Explicit

' ================================= Manipula��o de Arquivo =================================
Public Enum FileErrorCodes
    ERR_FILE_NOT_FOUND = 1000 ' Erro de arquivo n�o encontrado
    ERR_READ_FILE_FAILED = 1001 ' Erro ao ler arquivo
    ERR_WRITE_FILE_FAILED = 1002  ' Erro ao escrever no arquivo
    ERR_SAVE_FILE_FAILED = 1003  ' Erro ao salvar o arquivo
End Enum

' ================================= Manipula��o de E-mail =================================
Public Enum EmailErrorCodes
    ERR_OUTLOOK_INIT_FAILED = 1100 ' Erro ao inicializar o Outlook
    ERR_EMAIL_SEND_FAILED = 1101 ' Erro de envio de e-mail
End Enum

' ================================= Banco de Dados =================================
Public Enum DBConnectionErrorCodes
    ERR_DB_CONNECTION_FAILED = 1200      ' Falha ao conectar ao banco de dados
    ERR_DB_TIMEOUT = 1201                ' Tempo de conex�o excedido
    ERR_DB_AUTHENTICATION_FAILED = 1202  ' Falha na autentica��o com o banco de dados
    ERR_DB_QUERY_FAILED = 1203           ' Erro ao executar a consulta SQL
    ERR_DB_DISCONNECTED = 1204           ' Conex�o com o banco de dados foi perdida
    ERR_DB_DRIVER_NOT_FOUND = 1205       ' Driver de banco de dados n�o encontrado
    ERR_DB_INVALID_CONNECTION_STRING = 1206 ' String de conex�o inv�lida
    ERR_DB_TRANSACTION_FAILED = 1207     ' Falha na transa��o com o banco de dados
End Enum

Public Enum DBDataOperationErrorCodes
    ERR_DB_INSERT_FAILED = 1300       ' Falha ao inserir dados no banco
    ERR_DB_UPDATE_FAILED = 1301       ' Falha ao atualizar dados no banco
    ERR_DB_DELETE_FAILED = 1302       ' Falha ao excluir dados do banco
    ERR_DB_DUPLICATE_ENTRY = 1303     ' Tentativa de inserir dados duplicados
    ERR_DB_CONSTRAINT_VIOLATION = 1304 ' Viola��o de restri��o (chave prim�ria, estrangeira, etc.)
    ERR_DB_NULL_VALUE = 1305          ' Valor nulo em campo obrigat�rio
    ERR_DB_DATA_TYPE_MISMATCH = 1306  ' Tipo de dado incompat�vel
    ERR_DB_RECORD_NOT_FOUND = 1307    ' Registro n�o encontrado para atualiza��o ou exclus�o
End Enum
