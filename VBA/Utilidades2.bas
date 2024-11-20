Attribute VB_Name = "Testes"
Option Explicit

Public Function printDicionario(dict As Scripting.Dictionary)
    Dim chv1 As Variant, chv2 As Variant, chv3 As Variant
    Debug.Print "{"
    For Each chv1 In dict.Keys
        If chv1 <> Empty Then
            Debug.Print "   """ & chv1 & """: {"
            For Each chv2 In dict(chv1).Keys
                If Not IsObject(dict(chv1)(chv2)) Then  ' Verifica se tem um dicionário aninhado
                    Debug.Print "       """ & chv2 & """: " & """" & dict(chv1)(chv2) & ""","
                Else
                    Debug.Print "       """ & chv2 & """: {"
                    For Each chv3 In dict(chv1)(chv2).Keys
                        Debug.Print "           """ & chv3 & """: " & """" & dict(chv1)(chv2)(chv3) & ""","
                    Next chv3
                    Debug.Print "       },"
                End If
            Next chv2
        End If
    Next chv1
    ' Call ModificarArquivoTXT(ThisWorkbook.Path & "\log.txt", texto)
End Function

Public Function CriarArquivoTXT()
    Dim caminhoArquivo As String, arquivo As Integer, texto As String
    
    caminhoArquivo = ThisWorkbook.Path & "\log.txt"     ' Define o caminho e nome do arquivo
    
    arquivo = FreeFile                                  ' Retona um número de arquivo livre
    Open caminhoArquivo For Output As #arquivo          ' Abre o arquivo para saída (output)
    
    Close #arquivo                                      ' Fecha o arquivo
End Function

Public Function AbrirArquivoTXT(caminhoArquivo As String)
    Shell "notepad.exe " & caminhoArquivo, vbNormalFocus    ' Abre o arquivo
End Function

Public Function LerArquivoTXT(caminhoArquivo As String) As String   ' Retorna o texto do arquivo lido
    Dim arquivo As Integer, texto As String, linhaTexto As String, linha As Long

    arquivo = FreeFile                                  ' Retona um número de arquivo livre
    Open caminhoArquivo For Input As #arquivo           ' Abre o arquivo para entrada (Input)
    
    ' Lê cada linha do arquivo:
    linha = 1
    Do While Not EOF(arquivo)
        Line Input #arquivo, linhaTexto
        If linha = 1 Then
            texto = linhaTexto                          ' Exibe a linha do arquivo
        Else
            texto = texto & vbNewLine & linhaTexto      ' Exibe a linha do arquivo
        End If
        linha = linha + 1
    Loop
    
    Close #arquivo                                      ' Fecha o arquivo
    
    LerArquivoTXT = texto
End Function

Public Function ModificarArquivoTXT(caminhoArquivo As String, texto As String)
    Dim arquivo As Integer
    arquivo = FreeFile                              ' Retorna um número de arquivo livre
    Open caminhoArquivo For Append As #arquivo      ' Abre o arquivo para acréscimo (append)
    Print #arquivo, texto                           ' Escreve o texto no arquivo
    Close #arquivo                                  ' Fecha o arquivo
End Function






