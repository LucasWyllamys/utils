Attribute VB_Name = "Utils"
Option Explicit

'================================================================================
' Módulo VBA: Utils
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 20/01/2026
' Atualizado em: 06/02/2026
' Versão: 1.2.0
' Habilitar bibliotecas:
'   Microsoft Scripting Runtime
'   Microsoft WinHTTP Services, version 5.1
'================================================================================

Public Function AbrirArquivoPowerShell(caminhoArquivo As String)
    Shell "powershell.exe -File " & caminhoArquivo, vbNormalFocus
End Function

' Descrição: Substitui chaves no texto por valores do dicionário
' Parâmetros:
'   - text: Texto no qual as chaves serão substituídas pelos valores
'   - keysValues: objeto Scripting.Dictionary com pares chave-valor
Public Function ReplaceKeys(text As String, keysValues As Scripting.dictionary) As String
    Dim key As Variant
    
    If Not keysValues Is Nothing Then ' Verifica se o dicionário não está vazio
        If text <> "" Then
            For Each key In keysValues.Keys   ' Itera sobre todos as chaves do dicionário
                text = Replace(text, key, keysValues(key)) ' Substitui os valores das respectivas chaves
            Next key
            ReplaceKeys = text
        End If
    End If
End Function

Public Function DividirTextoEmColecao( _
    ByVal texto As String, _
    Optional ByVal delimitador As String = ";") As Collection

    Dim partes() As String
    Dim resultado As New Collection
    Dim i As Long
    Dim item As String

    texto = Trim(texto)
    delimitador = CStr(delimitador)

    ' Se texto estiver vazio, retorna coleção vazia
    If Len(texto) = 0 Then
        Set DividirTextoEmColecao = resultado
        Exit Function
    End If

    ' Divide o texto pelo delimitador
    partes = Split(texto, delimitador)

    ' Percorre o array e adiciona os itens limpos
    For i = LBound(partes) To UBound(partes)
        item = Trim(partes(i))
        If Len(item) > 0 Then
            resultado.Add item
        End If
    Next i

    Set DividirTextoEmColecao = resultado
End Function

Public Function ValidaValor(valor As String, tipo As String) As Boolean
    ' Esta função valida os tipos de dados de acordo com o tipo informado.
    valor = Trim(Replace(Replace(Replace(valor, ".", ""), "-", ""), "/", ""))
    If tipo = "cpf" Then
        If Len(valor) = 11 Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    ElseIf tipo = "cep" Then
        If Len(valor) = 8 Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    ElseIf tipo = "cnpj" Then
        If Len(valor) = 14 Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    ElseIf tipo = "data" Then
        If IsDate(valor) Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    End If
End Function

Public Function FormataValor(valor As String, tipo As String) As String
    ' Esta função formata os dados de acordo com o tipo informado.
    If tipo = "cpf" Then
        FormataValor = Format(valor, "000\.000\.000\-00")
    ElseIf tipo = "cep" Then
        FormataValor = Format(valor, "00\.000\-000")
    ElseIf tipo = "cnpj" Then
        FormataValor = Format(valor, "00\.000\.000/0000\-00")
    End If
End Function

Public Function Aguardar(tempoEspera As String)
    'Formato tempoespera: 00:00:00
    Dim tempo As Double
    tempo = Now + TimeValue(tempoEspera)
    Application.Wait tempo
End Function

Public Function LimparFormatacao(valor As String) As String
    LimparFormatacao = Trim(Replace(Replace(Replace(valor, ".", ""), "-", ""), "/", ""))
End Function

Public Function AbrirSite(url As String)
    Shell "cmd /c start " & url, vbHide     ' Abre o link do site no navegador padrão.
End Function

Public Function GetUsuario() As String
    GetUsuario = Environ("USERNAME")
    ' GetUsuario = CreateObject("WScript.Network").UserName
End Function

Public Function LerRespostaComoUTF8(request As WinHttpRequest) As String
    ' Esta função usa ADODB.Stream para ler a resposta  de uma requisição HTTP como UTF-8.
    Dim stream As Object
    Dim responseText As String

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                             ' adTypeBinary
    stream.Open
    stream.Write request.ResponseBody
    stream.Position = 0
    stream.Type = 2                             ' adTypeText
    stream.Charset = "utf-8"
    responseText = stream.ReadText
    stream.Close
    Set stream = Nothing

    LerRespostaComoUTF8 = responseText
End Function
