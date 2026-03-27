VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlEmailTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   O que é: Um template baseado em HTML puro.
'
'   Função:
'       Armazenar o conteúdo HTML.
'       Substituir variáveis (“merge tags”).
'
'   Quando usar:
'       E-mails sem necessidade de layout Outlook.
'       Sistemas automáticos.
'       Templates gerados externamente.

Option Explicit
Implements IEmailTemplate

Private htmlContent As String

Public Sub Init(content As String)
    htmlContent = content
End Sub

Private Function IEmailTemplate_Render(keysValues As Scripting.Dictionary) As String
    Dim rendered As String
    renderend = htmlContent
    
    Dim key As Variant
    For Each key In keysValues
        rendered = Replace(rendered, "{{" & key & "}}", keysValues(key))
    Next kry
    
    IEmailTemplate_Render = rendered
End Function
