VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OftEmailTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   O que é: Um template baseado em arquivo .OFT do Outlook.
'
'   Função:
'       Abrir o template dentro do Outlook.
'       Permitir substituição de variáveis.
'
'   Por que existe?
'       OFT é um formato complexo — contém metadados, formatações exclusivas e recursos do Outlook.
'       Abrir um OFT como texto quebraria tudo.
'
'   Observação: template OFT precisa ser carregado via Outlook.

Option Explicit
Implements IEmailTemplate

Private outlookMail As Outlook.MailItem

Public Sub Init(mail As Outlook.MailItem)
    Set outlookMail = mail
End Sub

Private Function IEmailTemplate_Render(keysValues As Scripting.IDictionary) As String
    Dim body As String
    body = outlookMail.HTMLBody
    
    Dim key As Variant
    For Each key In keysValues.keys
        body = Replace(body, key, keysValues(key))
    Next key
    
    IEmailTemplate_Render = body
End Function
