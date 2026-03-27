VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlEmailTemplateLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Função:
'       Ler um arquivo .html ou .htm da pasta e transformar em um template utilizável.
'
'   Responsabilidades:
'       Abrir arquivo
'       Ler conteúdo
'       Criar instância de HtmlEmailTemplate

Option Explicit
Implements IEmailTemplateLoader

Private Function IEmailTemplateLoader_LoadTemplate(ByVal filePath As String) As IEmailTemplate
    Dim fso As Scripting.FileSystemObject, content As String
    
    With fso.OpenTextFile(filePath)
        content = .ReadAll
        .Close
    End With
    
    Dim htmltemplate As New HtmlEmailTemplate
    htmltemplate.Init content
    Set IEmailTemplateLoader_LoadTemplate = htmltemplate
End Function
