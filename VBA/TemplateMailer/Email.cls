VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Estrutura de Dados (DTO - Data Transfer Object)
'   Representa os dados de um e-mail de forma organizada, limpa, desacoplada e coerente com OOP.
'
'   Carrega os dados do e-mail que será enviado.
'
'   Campos típicos:
'       ToAddress
'       Subject
'       Body
'
'   Propósito: Isolar os dados do e-mail em uma estrutura limpa e fácil de manipular.

Option Explicit

Public toAddress As String
Public subject As String
Public body As String
