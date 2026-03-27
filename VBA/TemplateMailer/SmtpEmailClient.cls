VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmtpEmailClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Função: Enviar e-mail via servidor SMTP.
'
'   O que faz:
'       Configura host, porta, SSL, autenticação
'       Usa objeto CDO.Message para envio
'       Prepara html
'       Envia e-mail pela internet
'
'   Quando usar:
'       Gmail, Hotmail, Office365 via SMTP
'       Servidores próprios corporativos

Option Explicit
Implements IEmailClient

Private smtpHost As String
Private smtpPort As Long

Public Sub Init(host As String, port As Long)
    smtpHost = host
    smtpPort = port
End Sub

Private Sub IEmailClient_SendEmail(ByVal mail As Email)
    'Requer referência: Microsoft CDO for Windows 2000 Library
    
    Dim msg As CDO.message
    Set msg = New CDO.message
    
    With msg
        .Configuration.fields.item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Configuration.fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Configuration.fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Configuration.fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpHost
        .Configuration.fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
    
        .Configuration.fields.Update
    
        .To = Email.toAddress
        .subject = Email.subject
        .HTMLBod = Email.body
        .Send
    End With
End Sub
