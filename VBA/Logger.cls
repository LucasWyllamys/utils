VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: Logger - Gravação de logs
' Descrição: Esta classe permite a gravação de registros em um arquivo de log, com diferentes níveis de severidade
' (DEBUG, INFO, WARN, ERROR, CRITICAL), facilitando o diagnóstico e rastreamento de eventos e erros em aplicações VBA.
' Versão: 1.3.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 17/10/2025
' Atualizado em: 02/02/2026
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

'=============================================== Constantes ===============================================

Public Enum FileErrorCodes
    ERR_FILE_NOT_FOUND = 1000 ' Erro de arquivo não encontrado
    ERR_READ_FILE_FAILED = 1001 ' Erro ao ler arquivo
    ERR_WRITE_FILE_FAILED = 1002  ' Erro ao escrever no arquivo
    ERR_SAVE_FILE_FAILED = 1003  ' Erro ao salvar o arquivo
    ERR_INVALID_FILE_TYPE = 1004  ' Tipo de arquivo inválido
End Enum

'========================================== Propriedades Privadas ==========================================

Private pLogFilePath As String ' Caminho completo do arquivo de log

'=========================================== Métodos Assessores ===========================================

Private Property Get LogFilePath() As String: LogFilePath = pLogFilePath: End Property

Private Property Let LogFilePath(Value As String): pLogFilePath = Value: End Property

'=========================================== Inicialização ===========================================

' Define o caminho padrão do arquivo de log ao instanciar a classe
Private Sub Class_Initialize()
    If ThisWorkbook.path = "" Then
        Err.Raise ERR_FILE_NOT_FOUND, "ClsLogger", "O projeto precisa estar salvo para gerar o log (" & Err.Number & "): " & Err.Description ' Lançamento de erro
    Else
        LogFilePath = ThisWorkbook.path & "\log.log" ' Define o caminho de arquivo log padrão
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

' Abre uma janela do PowerSheel para monitorar o arquivo de logs e retorna o PID (Processor ID) da janela
Public Function OpenLogInRealTime(Optional lastLines As Integer) As Long
    Dim cmd As String
    If lastLines > 0 Then
        cmd = "powershell -NoExit -Command " & _
                  """Get-Content '" & LogFilePath & "' -Tail " & lastLines & " -Wait"""
    Else
        cmd = "powershell -NoExit -Command " & _
                  """Get-Content '" & LogFilePath & "' -Wait"""
    End If
    ' Executa o comando e guarda o PID (Processor ID) retornado pelo Shell
    OpenLogInRealTime = Shell(cmd, vbMaximizedFocus)
End Function

Public Function CloseLogInRealTimeWMI(psPID As Long) As Long
    On Error Resume Next
    If psPID > 0 Then
        Dim obj As Object
        Set obj = GetObject("winmgmts:").Get("Win32_Process.Handle='" & CStr(psPID) & "'")
        If Not obj Is Nothing Then
            obj.Terminate 0
        End If
    End If
    On Error GoTo 0
    CloseLogInRealTimeWMI = 0
End Function

' Registra uma mensagem não formatada
Public Sub UnformattedMessage(message As String)
    Record message
End Sub

' Registra uma mensagem de depuração (diagnóstico técnico)
Public Sub DebugLog(message As String, Optional step As Long, Optional total As Long)
    Record FormatLogLine("DEBUG", message, step, total)
End Sub

 ' Registra uma mensagem informativa (eventos importantes)
Public Sub InfoLog(message As String, Optional step As Long, Optional total As Long)
    Record FormatLogLine("INFO", message, step, total)
End Sub

' Registra um alerta (possíveis problemas)
Public Sub WarnLog(message As String, Optional step As Long, Optional total As Long)
    Record FormatLogLine("WARN", message, step, total)
End Sub

' Registra um erro (falhas que afetam funcionalidades)
Public Sub ErrorLog(message As String, Optional step As Long, Optional total As Long)
    Record FormatLogLine("ERROR", message, step, total)
End Sub

' Registra um erro crítico (falhas graves que comprometem o sistema)
Public Sub CriticalLog(message As String, Optional step As Long, Optional total As Long)
    Record FormatLogLine("CRITICAL", message, step, total)
End Sub

'=========================================== Métodos Privados ===========================================

' Monta a linha de log formatada com data, tipo, etapa e mensagem
Private Function FormatLogLine(kind As String, message As String, Optional step As Long = 0, Optional total As Long = 0) As String
    Dim logLine As String
    
    logLine = "[" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "][" & kind & "]"
    If step <> 0 And total <> 0 Then
        logLine = logLine & "[" & step & " de " & total & "]"
    End If
    logLine = logLine & " " & message
    
    FormatLogLine = logLine
End Function

' Escreve a mensagem no arquivo de log
Private Sub Record(message As String)
    FileManipulation.WriteFile LogFilePath, message, True
End Sub

