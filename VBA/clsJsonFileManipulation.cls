VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonFileManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: clsJsonFileManipulation - Manipulação de arquivos JSON
' Descrição: Esta classe permite carregar, salvar, limpar e converter objetos JSON no VBA utilizando o JsonConverter e o módulo auxiliar FileManipulation para operações de leitura e escrita de arquivos.
' Versão: 1.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 17/10/2025
' Atualizado em: 24/10/2025
' Pré-requisitos:
'   - O projeto VBA deve estar salvo em um diretório válido.
'   - Importar os módulos:
'       - JsonConverter: https://github.com/VBA-tools/VBA-JSON
'       - FileManipulation: https://github.com/LucasWyllamys/utils/blob/main/VBA/FileManipulation.bas
'       - PublicConstants: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
' Exemplo de uso:
'   Dim jsonHandler As ClsJsonFileManipulation
'   Set jsonHandler = New ClsJsonFileManipulation
'
'   jsonHandler.LoadFromFile "C:\Dados\config.json"
'   Debug.Print jsonHandler.GetJsonString
'
'   jsonHandler.jsonObj.Add "versao", "1.0.0"
'   jsonHandler.SaveToFile
' Observações:
'   - O arquivo JSON é criado automaticamente se não existir.
'   - O conteúdo do JSON é armazenado como um objeto Dictionary.
'   - A classe depende do JsonConverter para interpretar e gerar strings JSON.
'===========================================================================================================

Option Explicit

'========================================= Propriedades Privadas ==========================================

Private pJsonFilePath As String
Private pJsonObj As Scripting.dictionary

'============================================== Inicialização =============================================

Private Sub Class_Initialize()
    Set pJsonObj = New Scripting.dictionary
End Sub

'=========================================== Métodos Assessores ===========================================

Private Property Get JsonFilePath() As String: JsonFilePath = pJsonFilePath: End Property

Private Property Let JsonFilePath(value As String): pJsonFilePath = value: End Property

Public Property Get jsonObj() As Scripting.dictionary: Set jsonObj = pJsonObj: End Property

Public Property Set jsonObj(obj As Scripting.dictionary): Set pJsonObj = obj: End Property

'=========================================== Métodos Públicos ===========================================

' Lê um arquivo JSON e armazena como objeto Dictionary
' @param jsonFilePath Caminho completo do arquivo JSON
' @throws ERR_FILE_NOT_FOUND se o arquivo não existir
' @throws ERR_READ_FILE_FAILED se houver erro na leitura
Public Sub LoadFromFile(filePath As String)
    Dim jsonText As String
    
    ' Verifica se o arquivo existe
    If FileManipulation.FileExists(filePath) Then
        JsonFilePath = filePath
        jsonText = FileManipulation.ReadFile(JsonFilePath) ' Lê o arquivo
        
        If jsonText <> "" Then
            Set jsonObj = JsonConverter.ParseJson(jsonText) ' Converte o texto JSON em objeto para ser manipulado no VBA.
        End If
    Else
        Err.Raise ERR_FILE_NOT_FOUND, "ClsJsonFileManipulation", "Arquivo não encontrado!"
    End If
End Sub

' Limpa todo o arquivo JSON
Public Sub ResetJson()
    Set jsonObj = New Scripting.dictionary
End Sub

' Retorna o conteúdo do JSON como string formatada
' @return String JSON formatada com indentação
Public Function GetJsonString() As String
    GetJsonString = JsonConverter.ConvertToJson(jsonObj, Whitespace:=4)
End Function

' Salva o objeto JSON no arquivo atual
Public Sub SaveToFile()
    Dim jsonText As String
    
    jsonText = JsonConverter.ConvertToJson(jsonObj, Whitespace:=4) ' Converte o objeto em string de JSON.
    FileManipulation.WriteFile JsonFilePath, jsonText, False ' Salva o arquivo
End Sub
