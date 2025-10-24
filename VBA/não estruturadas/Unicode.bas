Attribute VB_Name = "Unicode"
Option Explicit

' Fun��o que converte sequ�ncias Unicode escapadas (como \u00E3) em caracteres reais (como �)
Public Function DecodificarUnicode(texto As String) As String
    Dim ReGex As RegExp ' Declara um objeto RegExp para trabalhar com express�es regulares
    Dim matches As Object ' Objeto que armazenar� os resultados encontrados pela express�o regular
    Dim codigo As String
    Dim caractere As String
    
    Set ReGex = New RegExp ' Instancia o objeto RegExp (requer refer�ncia � Microsoft VBScript Regular Expressions 5.5)

    ' Configura��es da express�o regular
    With ReGex
        .Global = True ' Aplica a express�o a todas as ocorr�ncias no texto
        .IgnoreCase = True ' Ignora mai�sculas/min�sculas
        .Pattern = "\\u([0-9A-F]{4})" ' Padr�o para encontrar sequ�ncias do tipo \uXXXX (hexadecimal)
    End With
    
    Set matches = ReGex.Execute(texto) ' Executa a busca no texto

    Dim i As Integer
    ' Loop para substituir cada ocorr�ncia de \uXXXX pelo caractere correspondente
    For i = matches.Count - 1 To 0 Step -1
        codigo = matches(i).SubMatches(0) ' Extrai o c�digo hexadecimal (XXXX)
        caractere = ChrW("&H" & codigo) ' Converte o c�digo hexadecimal em caractere Unicode
        texto = Replace(texto, matches(i).value, caractere) ' Substitui a sequ�ncia \uXXXX pelo caractere real
    Next i

    DecodificarUnicode = texto ' Retorna o texto com os caracteres convertidos
End Function
