Attribute VB_Name = "utils_sheets"
'namespace=vba-files/Module/Utils
Option Explicit

'  Retourne la dernier ligne non vide d'une colonne
Public Function LastNumberColNotEmpty(workSheetTitle As String, col As Integer) As Integer
  LastNumberColNotEmpty = Sheets(workSheetTitle).Cells(Rows.Count, col).End(xlUp).row
End Function
  
'  Retourne la dernier colonne non vide d'une ligne
Public Function LastNumberRowNotEmpty(workSheetTitle As String, row As Integer) As Integer
  LastNumberRowNotEmpty = Sheets(workSheetTitle).Cells(row, Columns.Count).End(xlToLeft).column
End Function

Public Sub clearBorders(sheetName As String, address As String)

  With Sheets(sheetName).Range(address)
  .Borders(xlEdgeBottom).LineStyle = xlNone
  .Borders(xlEdgeTop).LineStyle = xlNone
  .Borders(xlEdgeLeft).LineStyle = xlNone
  .Borders(xlEdgeRight).LineStyle = xlNone
  .Borders(xlInsideHorizontal).LineStyle = xlNone
  .Borders(xlInsideVertical).LineStyle = xlNone
  End With
  
End Sub