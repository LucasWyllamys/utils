Attribute VB_Name = "FileManipulation"
Option Explicit

'==================================== Fun��es P�blicas ====================================

Public Function FileExists(filePath As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    FileExists = fso.FileExists(filePath)
End Function

Public Function ReadFile(filePath As String) As String
    Dim fso As Scripting.FileSystemObject
    Dim fileStream As Scripting.TextStream
    
    On Error GoTo ErrorHandler
    
    Set fso = New Scripting.FileSystemObject
    Set fileStream = fso.OpenTextFile(filePath, ForReading)  ' Abre o arquivo em modo leitura.
    
    If fileStream.AtEndOfStream Then ' Verifica se o arquivo est� vazio
        ReadFile = ""
    Else
        ReadFile = fileStream.ReadAll ' L� todo o arquivo e atribui � vari�vel.
    End If
    
    fileStream.Close ' Fecha o aquivo.
    
    Exit Function
ErrorHandler:
    Err.Raise ERR_READ_FILE_FAILED, "FileManipulation", "Erro ao ler arquivo (" & Err.Number & "): " & Err.Description ' Lan�amento de erro
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
    Err.Raise ERR_SAVE_FILE_FAILED, "FileManipulation", "Erro ao salvar arquivo (" & Err.Number & "): " & Err.Description
End Sub
