VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmailClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Propósito: Definir como enviar um e-mail, independente do provedor (SMTP, Outlook, API externa, etc).
'
'   Problema que resolve:
'       Sem ela, você precisaria reescrever código sempre que quisesse trocar o provedor de envio.
'       Com ela, você muda o provedor sem tocar no resto da aplicação: EmailService -> depende da interface, não de Outlook ou SMTP.
'
'   Método obrigatório: SendEmail(mail As Email) – envia o e-mail.

Option Explicit

Public Sub SendEmail(ByVal mail As Email)
End Sub
