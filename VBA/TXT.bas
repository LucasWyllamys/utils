Attribute VB_Name = "Testes"
Option Explicit

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






