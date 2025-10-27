VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: clsLogger - Gravação de logs
' Descrição: Esta classe permite a gravação de registros em um arquivo de log, com diferentes níveis de severidade
' (DEBUG, INFO, WARN, ERROR, CRITICAL), facilitando o diagnóstico e rastreamento de eventos e erros em aplicações VBA.
' Versão: 1.1.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 17/10/2025
' Atualizado em: 24/10/2025
' Pré-requisitos:
'   - O projeto VBA deve estar salvo em um diretório válido.
'   - Importar os módulos:
'       - FileManipulation: https://github.com/LucasWyllamys/utils/blob/main/VBA/FileManipulation.bas
'       - PublicConstants: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
' Exemplo de uso:
'   Dim logger As clsLogger
'   Set logger = New clsLogger
'
'   logger.SetLogFilePath "C:\Logs\meu_log.log"
'
'   logger.DebugLog "Mensagem de depuração"
'   logger.InfoLog "Mensagem informativa"
'   logger.WarnLog "Mensagem de alerta"
'   logger.ErrorLog "Mensagem de erro"
'   logger.CriticalLog "Mensagem crítica"
'
' Observações:
'   - O arquivo de log é criado automaticamente se não existir.
'   - Os parâmetros `step` e `total` são opcionais e indicam progresso de execução (ex: [0001/0010]).
'===========================================================================================================

Option Explicit

'========================================== Propriedades Privadas ==========================================

Private pLogFilePath As String ' Caminho completo do arquivo de log

'=========================================== Métodos Assessores ===========================================

Private Property Get LogFilePath() As String: LogFilePath = pLogFilePath: End Property

Private Property Let LogFilePath(value As String): pLogFilePath = value: End Property

'=========================================== Inicialização ===========================================

' Define o caminho padrão do arquivo de log ao instanciar a classe
Private Sub Class_Initialize()
    If ThisWorkbook.Path = "" Then
        Err.Raise ERR_FILE_NOT_FOUND, "ClsLogger", "O projeto precisa estar salvo para gerar o log (" & Err.Number & "): " & Err.Description ' Lançamento de erro
    Else
        LogFilePath = ThisWorkbook.Path & "\log.log" ' Define o caminho de arquivo log padrão
    End If
End Sub

'=========================================== Métodos Públicos ===========================================

' Permite definir um caminho personalizado para o arquivo de log
Public Sub SetLogFilePath(Optional filePath As String = "")
    If FileManipulation.FileExists(filePath) Then
        LogFilePath = filePath
    Else
        Err.Raise ERR_FILE_NOT_FOUND, "ClsLogger", "Arquivo não encontrado (" & Err.Number & "): " & Err.Description
    End If
End Sub

' Registra uma mensagem de depuração (diagnóstico técnico)
Public Sub DebugLog(message As String, Optional step As Integer, Optional total As Integer)
    Record SewLogLine("DEBUG", message, step, total)
End Sub

 ' Registra uma mensagem informativa (eventos importantes)
Public Sub InfoLog(message As String, Optional step As Integer, Optional total As Integer)
    Record SewLogLine("INFO", message, step, total)
End Sub

' Registra um alerta (possíveis problemas)
Public Sub WarnLog(message As String, Optional step As Integer, Optional total As Integer)
    Record SewLogLine("WARN", message, step, total)
End Sub

' Registra um erro (falhas que afetam funcionalidades)
Public Sub ErrorLog(message As String, Optional step As Integer, Optional total As Integer)
    Record SewLogLine("ERROR", message, step, total)
End Sub

' Registra um erro crítico (falhas graves que comprometem o sistema)
Public Sub CriticalLog(message As String, Optional step As Integer, Optional total As Integer)
    Record SewLogLine("CRITICAL", message, step, total)
End Sub

'=========================================== Métodos Privados ===========================================

' Monta a linha de log formatada com data, tipo, etapa e mensagem
Private Function SewLogLine(kind As String, message As String, Optional step As Integer = 0, Optional total As Integer = 0) As String
    Dim logLine As String
    
    logLine = "[" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "][" & kind & "]"
    If step <> 0 And total <> 0 Then
        logLine = logLine & "[" & Format(step, "0000") & "/" & Format(total, "0000") & "]"
    End If
    logLine = logLine & " " & message
    
    SewLogLine = logLine
End Function

' Escreve a mensagem no arquivo de log
Private Sub Record(message As String)
    FileManipulation.WriteFile LogFilePath, message, True
End Sub
