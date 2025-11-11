VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelTemplateProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================== Propriedades Privadas ========================================

Private pWbTemplate As Workbook
Private pWsTemplate As Worksheet

Private pTemplatePath As String

'====================================== Constantes =========================================

Private Const CLASS_NAME As String = "clsExcelTemplateProcessor"

'========================================= Métodos Assessores ========================================

Private Property Get wbTemplate() As Workbook
    Set wbTemplate = pWbTemplate
End Property

Private Property Set wbTemplate(object As Workbook)
    Set pWbTemplate = wbTemplate
End Property

Private Property Get wsTemplate() As Worksheet
    Set wsTemplate = pWsTemplate
End Property

Private Property Set wsTemplate(object As Worksheet)
    Set pWsTemplate = wsTemplate
End Property

Public Property Get templatePath() As String
    templatePath = pTemplatePath
End Property

Public Property Let templatePath(Value As String)
    pTemplatePath = Value
End Property

