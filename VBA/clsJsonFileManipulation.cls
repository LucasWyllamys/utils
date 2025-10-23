VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonFileManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================== Propriedades ===========================================

Private pCaminhoArquivo As String
Private pObjJson As Scripting.Dictionary

'=========================================== Métodos Assessores ===========================================

Private Property Get caminhoArquivo() As String
    caminhoArquivo = pCaminhoArquivo
End Property

Private Property Let caminhoArquivo(value As String)
    pCaminhoArquivo = value
End Property

Public Property Get objJson() As Scripting.Dictionary
    Set objJson = pObjJson
End Property

Public Property Set objJson(obj As Scripting.Dictionary)
    Set pObjJson = obj
End Property

'=========================================== Métodos Públicos ===========================================

' Lê o conteúdo do JSON e retorna como objeto
Public Sub LerJSON(caminhoArquivoJSON As String)
    Dim fso As New FileSystemObject
    Dim arquivo As TextStream
    Dim textoJson As String
    Dim json As Scripting.Dictionary
    
    caminhoArquivo = caminhoArquivoJSON
    
    Set arquivo = fso.OpenTextFile(caminhoArquivo, ForReading)  ' Abre o arquivo em modo leitura.
    textoJson = arquivo.ReadAll ' Lê todo o arquivo e atribui à variável.
    arquivo.Close ' Fecha o aquivo.
    
    Set json = JsonConverter.ParseJson(textoJson) ' Converte o texto JSON em objeto para ser manipulado no VBA.
    Set objJson = json ' Armazena na propriedade privada.
End Sub

' Limpa todo o arquivo JSON
Public Sub LimparJSON()
    Set objJson = New Scripting.Dictionary
    SalvarJSON
End Sub

' Retorna a string JSON formatada
Public Function ObterStringJSON() As String
    ObterStringJSON = JsonConverter.ConvertToJson(objJson, Whitespace:=4)
End Function

' Salva o objeto JSON no arquivo com codificação UTF-8 (acentos preservados)
Public Sub SalvarJSON()
    Dim fso As Scripting.FileSystemObject
    Dim arquivo As Scripting.TextStream
    Dim textoJson As String
    
    textoJson = JsonConverter.ConvertToJson(objJson, Whitespace:=4) ' Converte o objeto em string de JSON.
    
    Set fso = New Scripting.FileSystemObject
    Set arquivo = fso.CreateTextFile(caminhoArquivo, True)
    
    With arquivo
        .Write textoJson ' Escreve o texto JSON no arquivo
        .Close ' Fecha o arquivo
    End With
End Sub
