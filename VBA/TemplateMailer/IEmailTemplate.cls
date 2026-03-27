VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmailTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Propósito: Representar um template de e-mail carregado, pronto para ser renderizado com variáveis.
'
'   Serve tanto para:
'       html
'       OFT
'       Markdown (se quiser implementar depois)
'       Qualquer formato futuro
'
'   Método obrigatório: Render(keysValues As Dictionary)
'       Substitui variáveis no corpo do template, ex.: {{Nome}} -> "Lucas" | {{Produto}} -> "Planilha X"

Option Explicit

Public Function Render(keysValues As Scripting.Dictionary) As String
End Function
