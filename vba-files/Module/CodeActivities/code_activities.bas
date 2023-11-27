Attribute VB_Name = "code_activities"
'namespace=vba-files/Module/CodeActivities
Option Explicit

' Action button
Public Sub initCodeActivities()

  ' 0 - clean plage
  Call code_activities.cleanCodeActivities()

  ' 1 - get all codesActivities
  Call code_activities.getAllCodeActivities()
End Sub

Private Sub cleanCodeActivities()
  Dim firstCol As Integer
  Dim firstRow As Integer
  Dim lastCol As Integer
  Dim lastRow As Integer

  firstRow = 5
  firstCol = 1

  lastCol = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_CODE_ACTIVITIES, firstRow)
  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_CODE_ACTIVITIES, firstCol)

  With Sheets(SHEET_NAME_CODE_ACTIVITIES).Range(Cells(4, firstCol), Cells(lastRow, lastCol))
  .unMerge
  .ClearContents
  .Interior.Color = xlNone
  .Font.Bold = False
  End With

  Call utils_sheets.clearBorders(SHEET_NAME_CODE_ACTIVITIES, Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol)).Address)

End Sub

Private Sub getAllCodeActivities()

  Dim wsConfig As Worksheet
  Dim wsCodeActivities As Worksheet
  Dim wbSite As Workbook
  Dim wsSite As Worksheet
  Dim strFichier As String
  Dim LastRow As Long, LastCol As Long, i As Long
  Dim nameOA As String

  'Referencez la feuille de configuration
  Set wsConfig = ThisWorkbook.Sheets("CONFIG")
  
  'Referencez la feuille de CODE_ACTIVITIES
  Set wsCodeActivities = ThisWorkbook.Sheets("CODE_ACTIVITES")
  
  'Desactivez les messages d'alerte
  Application.DisplayAlerts = False
  
  'Parcourir la plage des fichiers a partir de D5
  i = 5
  Do While wsConfig.Cells(i, 4).Value <> ""
    strFichier = wsConfig.Cells(i, 4).Value
  
    'Ouvrez le fichier du site
    Set wbSite = Workbooks.Open(strFichier)
    'on recupère le nom de l'ouvrage
    nameOA = wbSite.Sheets("CONFIG").Range("E36").value
    'Referencez la feuille CODE_ACTIVITIES du site
    Set wsSite = wbSite.Sheets("CODE_ACTIVITES")

    'Determinez la dernière colonne et ligne du site
    LastCol = wsSite.Cells(4, wsSite.Columns.Count).End(xlToLeft).Column
    LastRow = wsSite.Cells(wsSite.Rows.Count, 1).End(xlUp).Row
    
    'Determinez où ajouter les donnees dans ce classeur
    Dim DestRow As Long
    If i = 5 Then
      DestRow = 5
    Else
      DestRow = wsCodeActivities.Cells(wsCodeActivities.Rows.Count, 1).End(xlUp).Row + 4
    End If
    
    ' Copiez la plage pour la mise en forme
    wsSite.Range(wsSite.Cells(4, 1), wsSite.Cells(LastRow, LastCol)).Copy
    wsCodeActivities.Cells(DestRow, 1).PasteSpecial xlPasteAll

    ' Transferez les valeurs sans formules
    Dim r As Long, c As Long
    For r = 4 To LastRow
      For c = 1 To LastCol
        wsCodeActivities.Cells(DestRow + r - 4, c).Value = wsSite.Cells(r, c).Value
      Next c
    Next r
    
    wsCodeActivities.Cells(DestRow - 1, 1).NumberFormat = "General"
    wsCodeActivities.Cells(DestRow - 1, 1).Font.Bold = True
    wsCodeActivities.Cells(DestRow - 1, 1).Font.Size = 14
    wsCodeActivities.Cells(DestRow - 1, 1).HorizontalAlignment = xlLeft
    wsCodeActivities.Cells(DestRow - 1, 1).value = "OUVRAGE : " & nameOA

    'Fermez le fichier du site
    wbSite.Close SaveChanges:=False

    'Passez au fichier suivant dans la liste
    i = i + 1
  Loop
  'Reactivez les messages d'alerte
  Application.DisplayAlerts = True
  'Nettoyez le Presse-papiers pour liberer la memoire
  Application.CutCopyMode = False
End Sub

