Attribute VB_Name = "PublicConstants"
'================================================================================
' Módulo VBA: PublicConstants - Constantes públicas
' Versão: 1.2.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 29/10/2025
' Atualizado em: 04/11/2025
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

' ================================= Banco de Dados =================================
Public Enum DBConnectionErrorCodes
    ERR_DB_CONNECTION_FAILED = 1250      ' Falha ao conectar ao banco de dados
    ERR_DB_TIMEOUT = 1251                ' Tempo de conexão excedido
    ERR_DB_AUTHENTICATION_FAILED = 1252  ' Falha na autenticação com o banco de dados
    ERR_DB_QUERY_FAILED = 1253           ' Erro ao executar a consulta SQL
    ERR_DB_DISCONNECTED = 1254           ' Conexão com o banco de dados foi perdida
    ERR_DB_DRIVER_NOT_FOUND = 1255       ' Driver de banco de dados não encontrado
    ERR_DB_INVALID_CONNECTION_STRING = 1256 ' String de conexão inválida
    ERR_DB_TRANSACTION_FAILED = 1257     ' Falha na transação com o banco de dados
End Enum

Public Enum DBDataOperationErrorCodes
    ERR_DB_INSERT_FAILED = 1300       ' Falha ao inserir dados no banco
    ERR_DB_UPDATE_FAILED = 1301       ' Falha ao atualizar dados no banco
    ERR_DB_DELETE_FAILED = 1302       ' Falha ao excluir dados do banco
    ERR_DB_DUPLICATE_ENTRY = 1303     ' Tentativa de inserir dados duplicados
    ERR_DB_CONSTRAINT_VIOLATION = 1304 ' Violação de restrição (chave primária, estrangeira, etc.)
    ERR_DB_NULL_VALUE = 1305          ' Valor nulo em campo obrigatório
    ERR_DB_DATA_TYPE_MISMATCH = 1306  ' Tipo de dado incompatível
    ERR_DB_RECORD_NOT_FOUND = 1307    ' Registro não encontrado para atualização ou exclusão
End Enum
