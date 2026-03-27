VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmailTemplateLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Propósito: Definir como templates são carregados do disco.
'
'   Por quê?
'       Templates HTML são texto puro.
'       Templates OFT precisam ser abertos dentro do Outlook.
'       Portanto, a lógica de carregar muda — e cada loader implementa sua própria versão.
'
'   Método obrigatório: LoadTemplate(filePath As String)
'
'   Retorna um objeto que implementa IEmailTemplate.

Option Explicit

Public Function LoadTemplate(ByVal filePath As String) As IEmailTemplate
End Function
