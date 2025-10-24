Attribute VB_Name = "Unicode"
Option Explicit

' Função que converte sequências Unicode escapadas (como \u00E3) em caracteres reais (como ã)
Public Function DecodificarUnicode(texto As String) As String
    Dim ReGex As RegExp ' Declara um objeto RegExp para trabalhar com expressões regulares
    Dim matches As Object ' Objeto que armazenará os resultados encontrados pela expressão regular
    Dim codigo As String
    Dim caractere As String
    
    Set ReGex = New RegExp ' Instancia o objeto RegExp (requer referência à Microsoft VBScript Regular Expressions 5.5)

    ' Configurações da expressão regular
    With ReGex
        .Global = True ' Aplica a expressão a todas as ocorrências no texto
        .IgnoreCase = True ' Ignora maiúsculas/minúsculas
        .Pattern = "\\u([0-9A-F]{4})" ' Padrão para encontrar sequências do tipo \uXXXX (hexadecimal)
    End With
    
    Set matches = ReGex.Execute(texto) ' Executa a busca no texto

    Dim i As Integer
    ' Loop para substituir cada ocorrência de \uXXXX pelo caractere correspondente
    For i = matches.Count - 1 To 0 Step -1
        codigo = matches(i).SubMatches(0) ' Extrai o código hexadecimal (XXXX)
        caractere = ChrW("&H" & codigo) ' Converte o código hexadecimal em caractere Unicode
        texto = Replace(texto, matches(i).value, caractere) ' Substitui a sequência \uXXXX pelo caractere real
    Next i

    DecodificarUnicode = texto ' Retorna o texto com os caracteres convertidos
End Function
