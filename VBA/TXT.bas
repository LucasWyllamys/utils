Attribute VB_Name = "TXT"
Option Explicit

Public Function CriarArquivoTXT(caminhoArquivo As String)
    Dim Arquivo As Integer, texto As String
    
    Arquivo = FreeFile                                  ' Retona um número de arquivo livre
    Open caminhoArquivo For Output As #Arquivo          ' Abre o arquivo para saída (output)
    
    Close #Arquivo                                      ' Fecha o arquivo
End Function

Public Function AbrirArquivoTXT(caminhoArquivo As String)
    Shell "notepad.exe " & caminhoArquivo, vbNormalFocus    ' Abre o arquivo como txt
End Function

Public Function LerArquivoTXT(caminhoArquivo As String) As String   ' Retorna o texto do arquivo lido
    Dim Arquivo As Integer, texto As String, linhaTexto As String, linha As Long

    Arquivo = FreeFile                                  ' Retona um número de arquivo livre
    Open caminhoArquivo For Input As #Arquivo           ' Abre o arquivo para entrada (Input)
    
    ' Lê cada linha do arquivo:
    linha = 1
    Do While Not EOF(Arquivo)
        Line Input #Arquivo, linhaTexto
        If linha = 1 Then
            texto = linhaTexto                          ' Exibe a linha do arquivo
        Else
            texto = texto & vbNewLine & linhaTexto      ' Exibe a linha do arquivo
        End If
        linha = linha + 1
    Loop
    
    Close #Arquivo                                      ' Fecha o arquivo
    
    LerArquivoTXT = texto
End Function

Public Function ModificarArquivoTXT(caminhoArquivo As String, texto As String)
    Dim Arquivo As Integer
    
    Arquivo = FreeFile                              ' Retorna um número de arquivo livre
    Open caminhoArquivo For Append As #Arquivo      ' Abre o arquivo para acréscimo (append)
    Print #Arquivo, texto                           ' Escreve o texto no arquivo
    
    Close #Arquivo                                  ' Fecha o arquivo
End Function






