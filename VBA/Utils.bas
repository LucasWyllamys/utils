Attribute VB_Name = "Utils"
Option Explicit

Public Function LerRespostaComoUTF8(request As WinHttpRequest) As String
    ' Usa ADODB.Stream para ler a resposta como UTF-8.
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

Public Function ValidaCamposObrigatorios(campos As Collection) As Boolean
    Dim campo As Variant
    For Each campo In campos
        If campo = Empty Then
            ValidaCamposObrigatorios = False
            Exit For
        Else
            ValidaCamposObrigatorios = True
        End If
    Next campo
End Function

Public Function ValidaValor(value As String, tipo As String) As Boolean
    ' Valida os tipos de dados.
    value = Trim(Replace(Replace(Replace(value, ".", ""), "-", ""), "/", ""))
    If tipo = "cpf" Then
        If Len(value) = 11 Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    ElseIf tipo = "cep" Then
        If Len(value) = 8 Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    ElseIf tipo = "data" Then
        If IsDate(value) Then
            ValidaValor = True
        Else
            ValidaValor = False
        End If
    End If
End Function

Public Function FormataValor(value As String, tipo As String) As String
    ' Formata os dados de acordo com o tipo informado.
    If tipo = "cpf" Then
        FormataValor = Format(value, "000\.000\.000\-00")
    ElseIf tipo = "cep" Then
        FormataValor = Format(value, "00\.000\-000")
    End If
End Function

Public Function AbrirSite(url As String)
    Shell "cmd /c start " & url, vbHide     ' Abre o link do site no navegador padrão.
End Function

Public Function GetUsuario() As String
    GetUsuario = Environ("USERNAME")
    ' GetUsuario = CreateObject("WScript.Network").UserName
End Function

Public Function CriarPasta(caminhoPasta As String) As String
    Dim objPasta As Object
    Set objPasta = CreateObject("Scripting.FileSystemObject")   ' Cria o objeto no Windows.
    If Not objPasta.FolderExists(caminhoPasta) Then             ' Verifica se a pasta não existe
        Set objPasta = objPasta.CreateFolder(caminhoPasta)      ' Cria a pasta caso não exista
    End If
End Function

Public Function printDicionario(dicionario As Scripting.Dictionary, caminhoArquivoTXT As String)
    ' Imprime o dicionário especificado no formato JSON no arquivo txt informado.
    Dim chv1 As Variant, chv2 As Variant, chv3 As Variant
    Debug.Print "{"
    For Each chv1 In dicionario.Keys
        If chv1 <> Empty Then
            Debug.Print "   """ & chv1 & """: {"
            For Each chv2 In dicionario(chv1).Keys
                If Not IsObject(dicionario(chv1)(chv2)) Then  ' Verifica se tem um dicionário aninhado
                    Debug.Print "       """ & chv2 & """: " & """" & dicionario(chv1)(chv2) & ""","
                Else
                    Debug.Print "       """ & chv2 & """: {"
                    For Each chv3 In dicionario(chv1)(chv2).Keys
                        Debug.Print "           """ & chv3 & """: " & """" & dicionario(chv1)(chv2)(chv3) & ""","
                    Next chv3
                    Debug.Print "       },"
                End If
            Next chv2
        End If
    Next chv1
    ' Call ModificarArquivoTXT(caminhoArquivoTXT, texto)
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
    Application.DisplayFormulaBar = True                    ' Barra de fórmulas
    ActiveWindow.DisplayHeadings = True                     ' Títulos
    ' Application.ActiveWindow.DisplayGridlines = True      ' Grades
End Function

Public Function LinhaVazia(planilha As Worksheet, colunaRef As String) As Long
    LinhaVazia = planilha.Cells(Rows.Count, colunaRef).End(xlUp).Row + 1
End Function

Public Function UltimaLinha(planilha As Worksheet, colunaRef As String) As Long
    UltimaLinha = planilha.Cells(Rows.Count, colunaRef).End(xlUp).Row
End Function
