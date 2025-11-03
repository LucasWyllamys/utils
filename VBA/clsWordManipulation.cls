VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWordManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================== Propriedades Privadas ========================================

Private pWdApp As Word.Application
Private pWdDoc As Word.Document

Private pTemplatePath As String

'========================================= Enum Tipo de Banco ========================================

Public Enum fileType
    docx = 1
    pdf = 2
End Enum

'========================================= Métodos Assessores ========================================

Private Property Get wdApp() As Word.Application: Set wdApp = pWdApp: End Property

Private Property Set wdApp(object As Word.Application): Set pWdApp = object: End Property


Private Property Get wdDoc() As Word.Document: Set wdDoc = pWdDoc: End Property

Private Property Set wdDoc(object As Word.Document): Set pWdDoc = object: End Property


Private Property Get templatePath() As String: templatePath = pTemplatePath: End Property

Private Property Let templatePath(value As String): pTemplatePath = value: End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    Set wdApp = New Word.Application
    wdApp.Visible = False
    wdApp.DisplayAlerts = False
End Sub

Private Sub Class_Terminate()
    Set wdApp = Nothing
    Set wdDoc = Nothing
End Sub

'=========================================== Métodos Públicos ===========================================

Public Sub ImportTemplate(strTemplatePath As String)
    If Dir(strTemplatePath) <> "" Then
        Set wdDoc = wdApp.Documents.Open(strTemplatePath, ReadOnly:=True)
        templatePath = strTemplatePath
    End If
End Sub

Public Sub ResetTemplatePath()
    ResetTemplatePath = ""
End Sub

Public Sub Replacekeys()

End Sub

Public Sub SaveAs(targetFolderPath As String, fileName As String, fileType As fileType)
    Select Case fileType
        Case docx
            wdDoc.SaveAs targetFolderPath & "\" & fileName & ".docx"
        Case pdf
            wdDoc.ExportAsFixedFormat OutputFileName:=targetFolderPath & "\" & fileName & ".pdf", ExportFormat:=17   'Exporta em PDF. 17 representa PDF
        Case Else
            Err.Raise ERR_INVALID_FILE_TYPE, "clsWordManipulation", "Tipo de arquivo não suportado: " & fileType
    End Select
    
    CloseWdDoc
End Sub

'=========================================== Métodos Privados ===========================================

Private Sub CloseWdDoc()
    If Not wdDoc Is Nothing Then
        wdDoc.Close SaveChanges:=False
        Set wdDoc = Nothing
    End If
End Sub

Private Sub CloseWdApp()
    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If
End Sub
