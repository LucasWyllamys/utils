Attribute VB_Name = "PublicConstants"
'================================================================================
' Módulo VBA: PublicConstants - Constantes públicas
' Versão: 1.3.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 29/10/2025
' Atualizado em: 17/11/2025
'================================================================================

Option Explicit

' ================================= Manipulação de Arquivo =================================
Public Enum FileErrorCodes
    ERR_FILE_NOT_FOUND = 1000 ' Erro de arquivo não encontrado
    ERR_READ_FILE_FAILED = 1001 ' Erro ao ler arquivo
    ERR_WRITE_FILE_FAILED = 1002  ' Erro ao escrever no arquivo
    ERR_SAVE_FILE_FAILED = 1003  ' Erro ao salvar o arquivo
    ERR_INVALID_FILE_TYPE = 1004  ' Tipo de arquivo inválido
End Enum

' ================================= Manipulação de E-mail =================================
Public Enum EmailErrorCodes
    ERR_OUTLOOK_INIT_FAILED = 1050 ' Erro ao inicializar o Outlook
    ERR_EMAIL_SEND_FAILED = 1051 ' Erro de envio de e-mail
End Enum

' ================================= Manipulação do Word =================================
Public Enum WordErrorCodes
    ERR_WORD_INIT_FAILED = 1200 ' Erro ao inicializar o Word
End Enum

' ================================= Manipulação do Excel =================================
Public Enum ExcelErrorCodes
    ERR_EXCEL_INIT_FAILED = 1250 ' Erro ao inicializar o Excel
End Enum

' ================================= Banco de Dados =================================
Public Enum DBConnectionErrorCodes
    ERR_DB_CONNECTION_FAILED = 1300     ' Falha ao conectar ao banco de dados
    ERR_DB_TIMEOUT = 1301               ' Tempo de conexão excedido
    ERR_DB_AUTHENTICATION_FAILED = 1302  ' Falha na autenticação com o banco de dados
    ERR_DB_QUERY_FAILED = 1303           ' Erro ao executar a consulta SQL
    ERR_DB_DISCONNECTED = 1304           ' Conexão com o banco de dados foi perdida
    ERR_DB_DRIVER_NOT_FOUND = 1305      ' Driver de banco de dados não encontrado
    ERR_DB_INVALID_CONNECTION_STRING = 1306 ' String de conexão inválida
    ERR_DB_TRANSACTION_FAILED = 1307     ' Falha na transação com o banco de dados
End Enum

Public Enum DBDataOperationErrorCodes
    ERR_DB_INSERT_FAILED = 1350       ' Falha ao inserir dados no banco
    ERR_DB_UPDATE_FAILED = 1351       ' Falha ao atualizar dados no banco
    ERR_DB_DELETE_FAILED = 1352       ' Falha ao excluir dados do banco
    ERR_DB_DUPLICATE_ENTRY = 1353     ' Tentativa de inserir dados duplicados
    ERR_DB_CONSTRAINT_VIOLATION = 1354 ' Violação de restrição (chave primária, estrangeira, etc.)
    ERR_DB_NULL_VALUE = 1355          ' Valor nulo em campo obrigatório
    ERR_DB_DATA_TYPE_MISMATCH = 1356  ' Tipo de dado incompatível
    ERR_DB_RECORD_NOT_FOUND = 1357    ' Registro não encontrado para atualização ou exclusão
End Enum
