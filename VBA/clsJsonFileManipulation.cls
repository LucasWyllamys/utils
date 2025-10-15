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

Private Type ClassType
    caminhoArquivo As String
    objJson As Scripting.Dictionary
End Type

Private self As ClassType

'=========================================== Métodos Assessores ===========================================

Private Property Get caminhoArquivo() As String
    caminhoArquivo = self.caminhoArquivo
End Property

Private Property Let caminhoArquivo(value As String)
    self.caminhoArquivo = value
End Property

Private Property Get objJson() As Scripting.Dictionary
    Set objJson = self.objJson
End Property

Private Property Set objJson(obj As Scripting.Dictionary)
    Set self.objJson = obj
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

' Atualiza um valor de uma chave ou adiciona chave-valor no arquivo JSON
Public Sub DefinirValue(chave As String, valor As Variant)
    If objJson.Exists(chave) Then ' Verifica se a chave existe.
        objJson(chave) = valor ' Atualiza o valor da chave especificada.
    Else
        objJson.Add chave, valor 'Adiciona uma nova chave e valor.
    End If
    
    Call SalvarJSON
End Sub

' Retorna o valor de uma chave
Public Function ObterValor(chave As String) As Variant
    If objJson.Exists(chave) Then ' Verifica se a chave existe.
        ObterValor = objJson(chave)
    Else
        ObterValor = "Chave não encontrada."
    End If
End Function

' Retorna a string JSON formatada
Public Function ObterStringJSON() As String
    ObterStringJSON = JsonConverter.ConvertToJson(objJson, Whitespace:=4)
End Function

' Remove uma chave do JSON
Public Sub RemoverItem(chave As String)
    If objJson.Exists(chave) Then
        objJson.Remove chave
        SalvarJSON
    Else
        MsgBox "Chave não encontrada."
    End If
End Sub

' Limpa todo o arquivo JSON
Public Sub LimparJSON()
    Set objJson = New Scripting.Dictionary
    Call SalvarJSON
End Sub

'=========================================== Métodos Privados ===========================================

' Salva o objeto JSON no arquivo
Private Sub SalvarJSON()
    Dim fso As New FileSystemObject
    Dim arquivo As TextStream
    Dim textoJson As String
    
    textoJson = JsonConverter.ConvertToJson(objJson, Whitespace:=4) ' Converte o objeto em string de JSON.
    
    Set arquivo = fso.CreateTextFile(caminhoArquivo, True)
    arquivo.Write textoJson ' Escreve o texto JSON no arquivo
    arquivo.Close ' Fecha o arquivo
End Sub
