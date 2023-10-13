VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Declenchement a l'ouverture du fichier
Public Sub Workbook_Open()
  Call general.DefineGlobalVariables()
End Sub

' Declenchement à l'activation d'une feuille
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  Call general.DefineGlobalVariables()
End Sub