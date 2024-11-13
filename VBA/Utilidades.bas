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
