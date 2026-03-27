VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OftEmailTemplateLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Função: Abrir o arquivo .OFT usando o Outlook:
'   Usa Application.CreateItemFromTemplate
'   Instancia um OftEmailTemplate

Option Explicit
Implements IEmailTemplateLoader

Private Function IEmailTemplateLoader_LoadTemplate(ByVal filePath As String) As IEmailTemplate
    Dim ol As Outlook.Application
    Dim mail As Outlook.MailItem
    
    Set ol = New Outlook.Application
    Set mail = ol.CreateItemFromTemplate(filePath)
    
    Dim template As New OftEmailTemplate
    template.Init mail
    Set IEmailTemplateLoader_LoadTemplate = template
End Function
