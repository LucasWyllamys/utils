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

'=========================================== Propriedades Privadas ===========================================

Private Type ClassType
    pCaminhoArquivoLog As String
End Type

Private self As ClassType

'=========================================== Métodos Assessores ===========================================

Private Property Get CaminhoArquivoLog() As String
    CaminhoArquivoLog = self.pCaminhoArquivoLog
End Property

Private Property Let CaminhoArquivoLog(value As String)
    self.pCaminhoArquivoLog = value
End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    CaminhoArquivoLog = ThisWorkbook.Path & "\log.log" ' Caminho de arquivo log padrão
End Sub

'=========================================== Métodos Públicos ===========================================

Public Sub SetCaminhoArquivoLog(Optional caminhoArquivo As String = "")
    If caminhoArquivo <> "" Then CaminhoArquivoLog = caminhoArquivo
End Sub

Public Sub Log(mensagem As String, Optional tipo As String = "INFO", Optional etapa As Integer, Optional total As Integer)
    Dim arquivo As Integer
    Dim linhaLog As String
    
    If etapa <> 0 And total <> 0 Then
        linhaLog = "[" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "][" & tipo & "][" & Format(etapa, "0000") & "/" & Format(total, "0000") & "] " & mensagem
    Else
        linhaLog = "[" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "][" & tipo & "] " & mensagem
    End If
    
    On Error Resume Next
    arquivo = FreeFile
    Open CaminhoArquivoLog For Append As #arquivo
    Print #arquivo, linhaLog
    Close #arquivo
    On Error GoTo 0
End Sub



