Attribute VB_Name = "FolderAndFileManipulator"
Option Explicit

'================================================================================
' Módulo VBA: FolderAndFileManipulator
' Versão: 2.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 21/01/2026
' Atualizado em: 02/03/2026
' Habilitar bibliotecas:
'   Microsoft Scripting Runtime
'================================================================================

Public Function FileExists(filePath As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    FileExists = fso.FileExists(filePath)
End Function

Public Function ReadFile(filePath As String) As String
    Dim fso As Scripting.FileSystemObject
    Dim fileStream As Scripting.TextStream
    
    On Error GoTo ErrorHandler
    
    Set fso = New Scripting.FileSystemObject
    Set fileStream = fso.OpenTextFile(filePath, ForReading)  ' Abre o arquivo em modo leitura.
    
    If fileStream.AtEndOfStream Then ' Verifica se o arquivo está vazio
        ReadFile = ""
    Else
        ReadFile = fileStream.ReadAll ' Lê todo o arquivo e atribui à variável.
    End If
    
    fileStream.Close ' Fecha o aquivo.
    
    Exit Function
ErrorHandler:
    Err.Raise ERR_READ_FILE_FAILED, "FolderAndFileManipulator", "Erro ao ler arquivo (" & Err.Number & "): " & Err.Description ' Lançamento de erro
End Function

Public Sub WriteFile(filePath As String, text As String, appendText As Boolean)
    Dim fso As Scripting.FileSystemObject
    Dim fileStream As Scripting.TextStream
    
    On Error GoTo ErrorHandler
    
    Set fso = New Scripting.FileSystemObject
    
    If appendText Then
        Set fileStream = fso.OpenTextFile(filePath, ForAppending, True) ' Abre o arquivo para escrita complementar (append)
    Else
        Set fileStream = fso.OpenTextFile(filePath, ForWriting, True) ' Abre o arquivo para escrita
    End If
    
    With fileStream
        .Write text & vbNewLine ' Escreve o texto no final do arquivo adicionando qubra de elinha
        .Close ' Fecha o arquivo
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_SAVE_FILE_FAILED, "FolderAndFileManipulator", "Erro ao salvar arquivo (" & Err.Number & "): " & Err.Description
End Sub

' Cria um arquivo de texto caso ele não exista
Public Sub CreateTextFile(filePath As String)
    Dim fso As New Scripting.FileSystemObject
    
    If Not fso.FileExists(filePath) Then
        fso.CreateTextFile filePath
    End If
End Sub

' Retorna a extensão do arquivo
Function GetFileExtension(filePath As String) As String
    Dim fso As New Scripting.FileSystemObject
    GetFileExtension = LCase(fso.GetExtensionName(filePath))
End Function

' Verifica se a pasta existe
Public Function FolderExists(folderPath As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    FolderExists = fso.FolderExists(folderPath)
End Function

' Remove caracteres inválidos em nomes de arquivo do Windows
Public Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalid As Variant, ch As Variant
    
    invalid = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SanitizeFileName = fileName
    
    For Each ch In invalid
        SanitizeFileName = Replace(SanitizeFileName, CStr(ch), "_")
    Next ch
    
    ' (opcional) aparar espaços e pontos finais
    SanitizeFileName = Trim$(SanitizeFileName)
    Do While Len(SanitizeFileName) > 0 And (Right$(SanitizeFileName, 1) = "." Or Right$(SanitizeFileName, 1) = " ")
        SanitizeFileName = Left$(SanitizeFileName, Len(SanitizeFileName) - 1)
    Loop
    
    If LenB(SanitizeFileName) = 0 Then
        SanitizeFileName = "arquivo"
    End If
End Function

' Monta path unindo pasta + nome sem extensão + ponto + extensão
Public Function BuildPath(ByVal folder As String, ByVal nameNoExt As String, ByVal ext As String) As String
    Dim sep As String
    sep = IIf(Right$(folder, 1) = "\" Or Right$(folder, 1) = "/", "", "\")
    BuildPath = folder & sep & nameNoExt & "." & ext
End Function

' Remove caracteres inválidos em nomes de arquivo do Windows
Public Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalid As Variant, ch As Variant
    
    invalid = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SanitizeFileName = fileName
    
    For Each ch In invalid
        SanitizeFileName = Replace(SanitizeFileName, CStr(ch), "_")
    Next ch
    
    ' (opcional) aparar espaços e pontos finais
    SanitizeFileName = Trim$(SanitizeFileName)
    Do While Len(SanitizeFileName) > 0 And (Right$(SanitizeFileName, 1) = "." Or Right$(SanitizeFileName, 1) = " ")
        SanitizeFileName = Left$(SanitizeFileName, Len(SanitizeFileName) - 1)
    Loop
    
    If LenB(SanitizeFileName) = 0 Then
        SanitizeFileName = "arquivo"
    End If
End Function

' Se existir, remove arquivo antes de salvar/exportar
Public Sub EnsureOverwrite(ByVal fullPath As String)
    On Error Resume Next
    If FolderAndFileManipulator.FileExists(fullPath) Then
        VBA.Kill fullPath
    End If
    On Error GoTo 0
End Sub

