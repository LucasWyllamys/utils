VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookEmailClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Função: Enviar e-mail usando a instalação do Outlook local.
'
'   O que faz:
'       Cria o e-mail no Outlook
'       Define To, Subject, HTML
'       Envia usando Outlook
'
'   Quando usar:
'       Usuários corporativos com Outlook instalado
'       Ambientes onde Outlook precisa registrar e-mails enviados

Option Explicit
Implements IEmailClient

Private Sub IEmailClient_SendEmail(ByVal mail As Email)
    Dim ol As Outlook.Application
    Dim item As Outlook.MailItem
    
    Set ol = New Outlook.Application
    Set item = ol.CreateItem(0)
    
    With item
        .To = mail.toAddress
        .subject = mail.subject
        .HTMLBody = mail.body
        .Send
    End With
End Sub
