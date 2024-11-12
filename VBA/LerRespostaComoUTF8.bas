Attribute VB_Name = "LerRespostaComoUTF8"
Option Explicit

Public Function LerRespostaComoUTF8(Requisicao As WinHttpRequest) As String
    Dim Stream As Object
    Dim ResponseText As String
    ' Usar ADODB.Stream para ler a resposta como UTF-8
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 1 ' adTypeBinary
    Stream.Open
    Stream.Write Requisicao.ResponseBody
    Stream.Position = 0
    Stream.Type = 2 ' adTypeText
    Stream.Charset = "utf-8"
    ResponseText = Stream.ReadText
    Stream.Close
    Set Stream = Nothing
    
    LerRespostaComoUTF8 = ResponseText
End Function
