Attribute VB_Name = "Utilidades"
Option Explicit

Public Function LerRespostaComoUTF8(request As WinHttpRequest) As String
    Dim stream As Object
    Dim responseText As String
        
    ' Usar ADODB.Stream para ler a resposta como UTF-8:
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

Public Function ValidarValor(value As String, tipo As String) As Boolean
    If tipo = "cpf" Then
        value = Replace(Replace(value, ".", ""), "-")
        If Len(value) = 11 Then
            ValidarValor = True
        Else
            ValidarValor = False
        End If
    ElseIf tipo = "data" Then
        If IsDate(value) Then
            ValidarValor = True
        Else
            ValidarValor = False
        End If
    End If
End Function

Public Function FormatarValor(value As String, tipo As String) As String
    If tipo = "cpf" Then
        FormatarValor = Format(value, "000.000.000-00")
    End If
End Function

Public Function AbrirSite(url As String)
    Shell "cmd /c start " & url, vbHide     ' Abre o link do site no navegador padrão.
End Function

Public Function FileName() As String
    Dim caminho_arquivo As String
    'Exibe o diálogo de seleção de arquivo. Pega o hiperlink do arquivo selecionado.
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear                               ' Limpar filtros
        ' .Filters.Add "Excel", "*.xlsx"             ' Define um filtro para limitar os tipos de arquivos que podem ser selecionados
        .Title = "Selecione um arquivo"              ' Define o título do pop-up
        .AllowMultiSelect = False                    ' Definir como True se desejar permitir a seleção de vários arquivos

        If .Show = -1 Then                           ' Verifica se o usuário clicou em "Abrir"
            FileName = .SelectedItems(1)             ' Obtem o caminho do arquivo selecionado
        End If
    End With
End Function

Public Function GetUsuario() As String
    GetUsuario = Environ("USERNAME")
    ' GetUsuario = CreateObject("WScript.Network").UserName
End Function

Public Function AnexarArquivo(caminhoArquivo As String, caminhoArquivoDestino As String) As String
    On Error GoTo tratar
    
    If caminhoArquivo <> "" Then
        FileCopy caminhoArquivo, caminhoArquivoDestino   'Copia o arquivo do caminho indicado e cola no caminho destino
        AnexarArquivo = caminhoArquivoDestino
    End If
    
    Exit Function
tratar:
    On Error GoTo 0
    Call msg_erro
End Function

Public Function AbrirArquivo(caminhoArquivo As String)
    If Dir(caminhoArquivo, vbArchive) <> "" And caminhoArquivo <> "" Then  'Verifica se o arquivo existe antes de tentar abri-lo
        Shell "explorer.exe """ & caminho_arquivo & """"    'Abre o arquivo usando o aplicativo padrão para esse tipo de arquivo
    Else
        Call msg_arquivo_nao_encontrado
    End If
End Function

Public Function CriarPasta(caminhoPasta As String) As String
    Dim objPasta As Object
    
    On Error GoTo tratar
    
    Set objPasta = CreateObject("Scripting.FileSystemObject")   ' Cria o objeto no Windows.
    If Not objPasta.FolderExists(caminhoPasta) Then             ' Verifica se a pasta não existe
        Set objPasta = objPasta.CreateFolder(caminhoPasta)      ' Cria a pasta caso não exista
    End If
    
    Exit Function
tratar:
    On Error GoTo 0
    Call msg_erro_pasta
End Function

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
