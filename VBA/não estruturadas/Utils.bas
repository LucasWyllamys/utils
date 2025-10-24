Attribute VB_Name = "Utils"
Option Explicit

' Habilitar bibliotecas:-------------------------------------------------------------
' Microsoft Scripting Runtime
' Microsoft WinHTTP Services, version 5.1
'------------------------------------------------------------------------------------

Public Function LerRespostaComoUTF8(request As WinHttpRequest) As String
    ' Esta fun��o usa ADODB.Stream para ler a resposta  de uma requisi��o HTTP como UTF-8.
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

Public Function PrintDicionario(dicionario As Scripting.Dictionary, caminhoArquivoTXT As String)
    ' Esta fun��o imprime um dicion�rio, no formato JSON, em um arquivo TXT.
    Dim jsonStr As String
    
    jsonStr = JsonConverter.ConvertToJson(dicionario)           ' Converte o dicion�rio para json.
    jsonStr = Utils.DecodeUnicode(jsonStr)                      ' Decodifica os caracteres unicode.
    jsonStr = Utils.FormatJson(jsonStr)                         ' Formata a string para um formato mais leg�vel de JSON.
    
    If Not FilesFolders.ArquivoExiste(caminhoArquivoTXT) Then          ' Verifica se o arquivo TXT existe.
        Call TXT.CriarArquivoTXT(caminhoArquivoTXT)             ' Cria um arquivo TXT caso ele n�o exista.
    End If
    
    Call TXT.EscreverArquivoTXT(caminhoArquivoTXT, jsonStr)     ' Escreve o dicion�rio no formato JSON no arquivo TXT.
End Function

Public Function ValidaCamposObrigatorios(Campos As Collection) As Boolean
    ' Esta fun��o recebe uma cole��o de valores e avalia se algum est� vazio.
    Dim campo As Variant
    For Each campo In Campos
        If TypeName(campo) = "CheckBox" Then
            If campo.value = False Then
                ValidaCamposObrigatorios = False
                Exit For
            Else
                ValidaCamposObrigatorios = True
            End If
        Else
            If campo = Empty Then
                ValidaCamposObrigatorios = False
                Exit For
            Else
                ValidaCamposObrigatorios = True
            End If
        End If
    Next campo
End Function

Public Function ValidaValor(valor As String, tipo As String) As Boolean
    ' Esta fun��o valida os tipos de dados de acordo com o tipo informado.
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
    ' Esta fun��o formata os dados de acordo com o tipo informado.
    If tipo = "cpf" Then
        FormataValor = Format(valor, "000\.000\.000\-00")
    ElseIf tipo = "cep" Then
        FormataValor = Format(valor, "00\.000\-000")
    ElseIf tipo = "cnpj" Then
        FormataValor = Format(valor, "00\.000\.000/0000\-00")
    End If
End Function

Public Function LimparFormatacao(valor As String) As String
    LimparFormatacao = Trim(Replace(Replace(Replace(valor, ".", ""), "-", ""), "/", ""))
End Function

Public Function AbrirSite(url As String)
    Shell "cmd /c start " & url, vbHide     ' Abre o link do site no navegador padr�o.
End Function

Public Function GetUsuario() As String
    GetUsuario = Environ("USERNAME")
    ' GetUsuario = CreateObject("WScript.Network").UserName
End Function

Public Function OcultarFerramentasExcel1()
    Application.DisplayFullScreen = True
    Application.ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    Application.ActiveWindow.DisplayGridlines = False   ' Desativar grades.
End Function

Public Function OcultarFerramentasExcel2()
    Application.DisplayFullScreen = True
    Application.ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    Application.ActiveWindow.DisplayGridlines = True    ' Ativar grades.
End Function

Public Function ExibirFerramentasExcel1()
    Application.DisplayFullScreen = False
    Application.ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    Application.ActiveWindow.DisplayGridlines = True
End Function

Public Function ExibirFerramentasExcel2()
    ' Application.DisplayFullScreen = False
    ' Application.ActiveWindow.DisplayWorkbookTabs = True   ' Abas
    Application.DisplayFormulaBar = True                    ' Barra de f�rmulas
    ActiveWindow.DisplayHeadings = True                     ' T�tulos
    ' Application.ActiveWindow.DisplayGridlines = True      ' Grades
End Function

Public Function LinhaVazia(planilha As Worksheet, colunaRef As String) As Long
    LinhaVazia = planilha.Cells(Rows.Count, colunaRef).End(xlUp).Row + 1
End Function

Public Function UltimaLinha(planilha As Worksheet, colunaRef As String) As Long
    UltimaLinha = planilha.Cells(Rows.Count, colunaRef).End(xlUp).Row
End Function
