Attribute VB_Name = "PublicConstantsVariables"
Option Explicit

'======================================= CONSTANTES =========================================
'=================== Constantes de Erro ===================
' Manipulação de arquivo:
Public Enum FileErrorCodes
    ERR_FILE_NOT_FOUND = 1000 ' Erro de arquivo não encontrado
    ERR_READ_FILE_FAILED = 1001 ' Erro ao ler arquivo
    ERR_WRITE_FILE_FAILED = 1002  ' Erro ao escrever no arquivo
    ERR_SAVE_FILE_FAILED = 1003  ' Erro ao salvar o arquivo
End Enum
' Manipulação de E-mail:
Public Enum EmailErrorCodes
    ERR_OUTLOOK_INIT_FAILED = 1100 ' Erro ao inicializar o Outlook
    ERR_EMAIL_SEND_FAILED = 1101 ' Erro de envio de e-mail
End Enum
