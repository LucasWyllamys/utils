VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITemplateProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ImportTemplate(ByVal templatePath As String)
End Sub

Public Function Export( _
    ByVal targetFolderPath As String, _
    ByVal fileName As String, _
    ByVal fileType As ExportFileType, _
    Optional ByVal keysValues As Scripting.Dictionary = Nothing) As String
End Function

' Reseta o documento carregado
Public Sub ResetTemplate()
End Sub

' Finaliza completamente a instância
Public Sub Finish()
End Sub


