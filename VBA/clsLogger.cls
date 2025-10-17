VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================ Constantes de Erro ============================================

Private Const ERR_LOG_INIT_FAILED As Long = 2000
Private Const ERR_LOG_WRITE_FAILED As Long = 2001

'=========================================== Propriedades Privadas ===========================================

Private pCaminhoArquivoLog As String

'=========================================== Métodos Assessores ===========================================

Private Property Get CaminhoArquivoLog() As String: CaminhoArquivoLog = pCaminhoArquivoLog: End Property

Private Property Let CaminhoArquivoLog(value As String): pCaminhoArquivoLog = value: End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo TratarErro
    CaminhoArquivoLog = ThisWorkbook.Path & "\log.log" ' Caminho de arquivo log padrão
    Exit Sub
TratarErro:
    Err.Raise ERR_LOG_INIT_FAILED, "clsLogger", "Erro ao definir caminho do arquivo de log: " & Err.Description ' Lançamento de erro
End Sub

'=========================================== Métodos Públicos ===========================================

'@description Define o caminho do arquivo de log
Public Sub SetCaminhoArquivoLog(Optional caminhoArquivo As String = "")
    If caminhoArquivo <> "" Then CaminhoArquivoLog = caminhoArquivo
End Sub

Public Sub Debug_(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Para informações detalhadas para diagnóstico técnico
    Registrar MontarLinhaLog("DEBUG", mensagem, etapa, total)
End Sub

Public Sub Info_(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Para eventos normais e importantes
    Registrar MontarLinhaLog("INFO", mensagem, etapa, total)
End Sub

Public Sub Warn_(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Para alertas sobre possíveis problemas
    Registrar MontarLinhaLog("WARN", mensagem, etapa, total)
End Sub

Public Sub Error_(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Para erros que afetam funcionalidades
    Registrar MontarLinhaLog("ERROR", mensagem, etapa, total)
End Sub

Public Sub Critical_(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Para erros graves que podem comprometer o sistema
    Registrar MontarLinhaLog("CRITICAL", mensagem, etapa, total)
End Sub

'=========================================== Métodos Privados ===========================================

Private Function MontarLinhaLog(tipo As String, mensagem As String, Optional etapa As Integer = 0, Optional total As Integer = 0) As String
    Dim linhaLog As String
    
    linhaLog = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "][" & tipo & "]"
    If etapa <> 0 And total <> 0 Then
        linhaLog = linhaLog & "[" & Format(etapa, "0000") & "/" & Format(total, "0000") & "]"
    End If
    linhaLog = linhaLog & " " & mensagem
    
    MontarLinhaLog = linhaLog
End Function

Private Sub Registrar(mensagem As String)
    Dim arquivo As Integer

    On Error GoTo TrataErro
        arquivo = FreeFile
        Open CaminhoArquivoLog For Append As #arquivo
        Print #arquivo, mensagem
        Close #arquivo
    Exit Sub
TrataErro:
    Err.Raise ERR_LOG_WRITE_FAILED, "clsLogger", "Erro ao registrar log: " & Err.Description
End Sub
