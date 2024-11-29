Attribute VB_Name = "FilesFolders"
Option Explicit

Public Function ArquivoExiste(caminhoArquivo As String) As Boolean
    If Dir(caminhoArquivo, vbArchive) <> "" And caminhoArquivo <> "" Then  'Verifica se o arquivo existe antes de tentar abri-lo
        ArquivoExiste = True
    Else
        ArquivoExiste = False
    End If
End Function

Public Function AbrirArquivo(caminhoArquivo As String)
    ' Abre o arquivo usando o aplicativo padr�o para esse tipo de arquivo
    Shell "explorer.exe """ & caminhoArquivo & """"
End Function

Public Function GetCaminhoArquivo() As String
    ' Exibe o di�logo de sele��o de arquivo. Pega o hiperlink do arquivo selecionado.
    Dim caminho_arquivo As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear                              ' Limpar filtros
        ' .Filters.Add "Excel", "*.xlsx"            ' Define um filtro para limitar os tipos de arquivos que podem ser selecionados
        .Title = "Selecione um arquivo"             ' Define o t�tulo do pop-up
        .AllowMultiSelect = False                   ' Definir como True se desejar permitir a sele��o de v�rios arquivos

        If .Show = -1 Then                          ' Verifica se o usu�rio clicou em "Abrir"
            GetCaminhoArquivo = .SelectedItems(1)   ' Obtem o caminho do arquivo selecionado
        End If
    End With
End Function

Public Function AnexarArquivo(caminhoArquivo As String, caminhoArquivoDestino As String) As String
    ' Copia o arquivo do caminho indicado e cola no caminho destino
    If caminhoArquivo <> "" Then
        FileCopy caminhoArquivo, caminhoArquivoDestino
        AnexarArquivo = caminhoArquivoDestino
    End If
End Function

Public Function CriarPasta(collCaminhosPasta As Collection)
    Dim objPasta As New Scripting.FileSystemObject, caminhoPasta As Variant
    'Set objPasta = CreateObject("Scripting.FileSystemObject")   ' Cria o objeto no Windows.
    For Each caminhoPasta In collCaminhosPasta
        If Not objPasta.FolderExists(caminhoPasta) Then             ' Verifica se a pasta n�o existe
            Set objPasta = objPasta.CreateFolder(caminhoPasta)      ' Cria a pasta caso n�o exista
        End If
    Next caminhoPasta
End Function
