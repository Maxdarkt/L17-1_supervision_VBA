Attribute VB_Name = "total_month"
'namespace=vba-files/Module/Total
Option Explicit

Dim nbDays As Integer
Dim workzones() As Variant
Dim arrNotWorkedDays() As Date

' Action button
Public Sub initTotalMonth()

  ' 0 - clean plage
  Call total_month.cleanTotalMonth()

  ' 1 - generer les dates et les ovurages
  Call total_month.generateDateMonth()

  ' 2 - get all hours
  Call total_month.ConsolidateTotalMois()

  ' 3 - coloration des jours non travailles
  Call total_month.colorNotWorkedDays()

  ' 4 - add sum by col
  Call total_month.addSumByCol()

  ' 5 - add sum by row by workzone
  Call total_month.addSumByRowByWorkzone()

  ' 6 - add sum by row
  Call total_month.addSumByRow()
End Sub

Public Sub cleanTotalMonth()
  Dim firstCol As Integer
  Dim firstRow As Integer
  Dim lastCol As Integer
  Dim lastRow As Integer

  firstRow = 1
  firstCol = 3

  lastCol = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_TOTAL_MONTH, 3)
  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_TOTAL_MONTH, firstCol)

  With Sheets(SHEET_NAME_TOTAL_MONTH).Range(Cells(firstRow, firstCol), Cells(lastRow + 1, lastCol + 1))
    .unMerge
    .ClearContents
    .Interior.Color = xlNone
    .Font.Bold = False
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .IndentLevel = 0
  End With

  Call utils_sheets.clearBorders(SHEET_NAME_TOTAL_MONTH, Range(Cells(firstRow, firstCol), Cells(lastRow + 1, lastCol + 1)).Address)

End Sub

' 1 - On genère les dates du mois
Public Sub generateDateMonth()
  ' Declaration de variables
  Dim arrDate() As String ' date saisie par l'utilisateur sous forme d'array
  Dim firstDay As Date
  Dim lastDay As Date
  Dim dateOfCol As Date

  ' Appeler la fonction pour recuperer la liste des ouvrages
  workzones = general.getListWorkzones()

  ' on recupère les cellules sous forme d'objet pour eviter les problème de format de date
  arrDate = Split(Sheets(SHEET_NAME_CONFIG).range("F5").Value, ".")

  firstDay = "01." + arrDate(1) + "." + arrDate(2)
  lastDay = utils_date.LastDayOfMonth(CLng(arrDate(1)), CLng(arrDate(2)))

  nbDays = utils_date.CalculateDurationBetweenDates(firstDay, lastDay)

  ' Declarations de variables pour boucle
  Dim i As Integer
  Dim j As Integer
  Dim firstRow As Integer
  Dim firstCol As Integer
  Dim varIsWorkingDay As Boolean
  Dim varIsDayNotWorked As Integer
  Dim typeCol As Integer

  arrNotWorkedDays = utils_worked_days.NotWorkedDays()

  firstRow = 4
  firstCol = 5

  For i = 1 To (day(nbDays) + 2)
    
    dateOfCol = DateAdd("d", i - 1, firstDay)

    ' On teste si ce jour est un jour travaille (semaine / week-end)
    varIsWorkingDay = utils_date.IsWorkingDay(dateOfCol)

    ' On teste si c'est un jour ferie national / projet
    varIsDayNotWorked = utils_worked_days.IsDayNotWorked(dateOfCol, arrNotWorkedDays())

    ' On boucle sur la liste des ouvrages
    For j = LBound(workzones) To UBound(workzones)
      For typeCol = 1 To 2
        ' Coloration de la cellule
        If varIsDayNotWorked > 0 Or varIsWorkingDay = False Then
          ' Si jour chôme
          ' Coloration des cellules jour non travaillé
          With sheets(SHEET_NAME_TOTAL_MONTH)
            .Cells(firstRow, firstCol).Interior.color = COLOR_CEL_TM_DAY_NOT_WORKED
            .Cells(firstRow - 1, firstCol).Interior.color = COLOR_CEL_TM_DAY_NOT_WORKED
            .Cells(firstRow - 2, firstCol).Interior.color = COLOR_CEL_TM_DAY_NOT_WORKED
          End With
        Else    ' Si jour travaille, couleur de la cellule par defaut
          ' Coloration des cellules jour travaillé
          With sheets(SHEET_NAME_TOTAL_MONTH)
            .Cells(firstRow, firstCol).Interior.color = COLOR_CEL_TM_DAY_WORKED
            .Cells(firstRow - 1, firstCol).Interior.color = COLOR_CEL_TM_DAY_WORKED
            .Cells(firstRow - 2, firstCol).Interior.color = COLOR_CEL_TM_DAY_WORKED
          End With
        End If

        ' Ecriture des cellules
        With Sheets(SHEET_NAME_TOTAL_MONTH)
          .Cells(firstRow, firstCol).Value = dateOfCol
          .Cells(firstRow, firstCol).NumberFormat = "dd"
          .Cells(firstRow, firstCol).Font.Bold = True
          .Cells(firstRow, firstCol).HorizontalAlignment = xlCenter
          .Cells(firstRow, firstCol).VerticalAlignment = xlCenter
          ' On ecrit le jour dans la cellule "J" ou "N"
          If typeCol = 1 Then
            .Cells(firstRow - 1, firstCol).Value = "J"
          Else
            .Cells(firstRow - 1, firstCol).Value = "N"
          End If
          .Cells(firstRow - 1, firstCol).Font.Bold = True
          .Cells(firstRow - 1, firstCol).HorizontalAlignment = xlCenter
          .Cells(firstRow - 1, firstCol).VerticalAlignment = xlCenter
          ' format date specialisee
          .Cells(firstRow - 2, firstCol).Formula = "=" & Replace(.Cells(firstRow, firstCol).Address, "$", "")
          .Cells(firstRow - 2, firstCol).NumberFormat = "ddd"
          .Cells(firstRow - 2, firstCol).HorizontalAlignment = xlCenter
          .Cells(firstRow - 2, firstCol).VerticalAlignment = xlCenter
          ' Nom de l'ouvrage
          .Cells(firstRow - 3, firstCol).Value = workzones(j, 1)
          .Cells(firstRow - 3, firstCol).HorizontalAlignment = xlCenter
          .Cells(firstRow - 3, firstCol).VerticalAlignment = xlCenter
          .Cells(firstRow - 3, firstCol).Orientation = 90
          .Cells(firstRow - 3, firstCol).Interior.Color = workzones(j, 2)
          .Cells(firstRow - 3, firstCol).Font.Color = RGB(0, 0, 0)
          ' format de cellule standard
        End With
        
        firstCol = firstCol + 1
        Next typeCol
      Next j
  Next i
End Sub

Sub ConsolidateTotalMois()

  Dim wsConfig As Worksheet, wsConsolidation As Worksheet
  Dim lastRowConfig As Long, i As Long
  Dim wbSite As Workbook, wsSite As Worksheet
  Dim lastRowSite As Long, j As Long, colOffset As Integer
  Dim employeeName As String
  Dim employeeCompany As String
  Dim Hours As Variant
  Dim wsSiteWorkZone As String
  DIm siteColOffset As Integer

  With Sheets(SHEET_NAME_TOTAL_MONTH).Range("C4")
    .Value = "NOM - PRENOM"
    .Font.Bold = True
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Interior.color = COLOR_CEL_READ_H1
    .IndentLevel = 1
  End With

  With Sheets(SHEET_NAME_TOTAL_MONTH).Range("D4")
    .Value = "ENTREPRISE"
    .Font.Bold = True
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Interior.color = COLOR_CEL_READ_H1
    .IndentLevel = 1
  End With

  ' Set references to worksheets
  Set wsConfig = ThisWorkbook.Sheets(SHEET_NAME_CONFIG)
  Set wsConsolidation = ThisWorkbook.Sheets(SHEET_NAME_TOTAL_MONTH)

  ' Find last row in CONFIG sheet
  lastRowConfig = wsConfig.Cells(wsConfig.Rows.Count, 4).End(xlUp).Row
  
  ' Loop through each site
  For i = 5 To lastRowConfig
    ' Desactiver les messages d'alerte
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
      ' Open site workbook
      Set wbSite = Workbooks.Open(wsConfig.Cells(i, 4).Value)
      Set wsSite = wbSite.Sheets(SHEET_NAME_TOTAL_MONTH)
      wsSiteWorkZone = wbSite.Sheets(SHEET_NAME_CONFIG).Range("E36").Value

      siteColOffset = general.getPositionWorkzonesInArray(wsSiteWorkZone, workzones)

      ' Find last row in site's TOTAL_MOIS sheet
      lastRowSite = 29

      ' Loop through each row (employee) in site's TOTAL_MOIS sheet
      For j = 5 To lastRowSite
        Dim EmployeeRow As Long
        Dim FoundCell As Range
        Dim typeCol As Integer

        employeeName = wsSite.Cells(j, 3).Value
        employeeCompany = wsSite.Cells(j, 4).Value

        If employeeName <> "" And Len(employeeName) > 1  Then

          Set FoundCell = wsConsolidation.Columns(3).Find(employeeName)
  
          ' Find or create row for this employee in consolidation sheet
          If Not FoundCell Is Nothing Then
            EmployeeRow = FoundCell.Row
          Else
            ' Always add to the next available row after the last filled row, ensuring it's at least row 4
            EmployeeRow = Application.WorksheetFunction.Max(4, wsConsolidation.Cells(wsConsolidation.Rows.Count, 3).End(xlUp).Row + 1)
            With wsConsolidation.Cells(EmployeeRow, 3)
              .Value = employeeName
              .Font.Bold = True
              .HorizontalAlignment = xlLeft
              .VerticalAlignment = xlCenter
              .IndentLevel = 1
              .Interior.Color = COLOR_CEL_READ_CONTENT
            End With
            With wsConsolidation.Cells(EmployeeRow, 4)
              .Value = employeeCompany
              .Font.Bold = False
              .HorizontalAlignment = xlLeft
              .VerticalAlignment = xlCenter
              .IndentLevel = 1
              .Interior.Color = COLOR_CEL_READ_CONTENT
            End With
          End If
  
          ' Copy hours for each day and site
          For colOffset = 0 To nbDays
            ' Colonne Jour et Nuit
            For typeCol = 1 To 2
              Dim adjustedCol As Integer

              adjustedCol = colOffset * (UBound(workzones) * 2) + 5 + ((siteColOffset - 1) * 2) + (typeCol - 1)

              Hours = wsSite.Cells(j, colOffset + (typeCol - 1) + 5).Value
  
              With wsConsolidation.Cells(EmployeeRow, adjustedCol)
                .Value = Hours
                .Font.Color = workzones(siteColOffset, 2)
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
              End With
            Next typeCol
          Next colOffset
        End If
      Next j
      ' Activer les messages d'alerte
      Application.DisplayAlerts = True
      Application.AskToUpdateLinks = True
      ' Close site workbook
      wbSite.Close SaveChanges:=False
  Next i
End Sub

Sub colorNotWorkedDays()
  Dim firstRow As Integer
  Dim firstCol As Integer
  Dim lastRow As Integer
  Dim lastCol As Integer
  Dim i As Integer
  Dim j As Integer
  Dim varIsDayNotWorked As Boolean

  firstRow = 3
  firstCol = 5

  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_TOTAL_MONTH, 3)
  lastCol = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_TOTAL_MONTH, lastRow)

  For j = firstCol To lastCol
    varIsDayNotWorked = colorCellIfDayIsNotWorked(firstRow, j, COLOR_CEL_TM_DAY_NOT_WORKED)
    If varIsDayNotWorked = True Then
      With Sheets(SHEET_NAME_TOTAL_MONTH).Range(Cells(firstRow + 1, j), Cells(lastRow, j))
        .Interior.color = COLOR_CEL_TM_DAY_NOT_WORKED
      End With
    End If
  Next j
End Sub

Sub addSumByCol()
  Dim firstRow As Integer
  Dim firstCol As Integer
  Dim lastRow As Integer
  Dim lastCol As Integer
  Dim i As Integer
  Dim j As Integer
  Dim Color As Variant

  firstRow = 4
  firstCol = 5

  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_TOTAL_MONTH, 3)
  lastCol = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_TOTAL_MONTH, lastRow)

  ' on ajoute le nom a la ligne
  With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(lastRow + 1, 4)
    .value = "TOTAL"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.color = COLOR_CEL_READ_CONTENT
  End With
  ' formule en bas de colonne
  For j = firstCol To lastCol
    color = Sheets(SHEET_NAME_TOTAL_MONTH).Cells(1, j).Interior.Color
    With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(lastRow + 1, j)
      .Formula = "=SUM(" & Replace(Cells(firstRow + 1, j).Address, "$", "") & ":" & Replace(Cells(lastRow, j).Address, "$", "") & ")"
      .NumberFormat = "0.00"
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .Font.Color = color
      .Interior.color = COLOR_CEL_READ_CONTENT
    End With
  Next j

End Sub

Sub addSumByRowByWorkzone()
  Dim firstRow As Integer
  Dim firstCol As Integer
  Dim lastRow As Integer
  Dim lastCol As Integer
  Dim i As Integer
  Dim j As Integer
  Dim Color As Variant

  firstRow = 3
  firstCol = 5

  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_TOTAL_MONTH, 3)
  lastCol = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_TOTAL_MONTH, lastRow)

  ' On boucle sur la liste des ouvrages
  For i = LBound(workzones) To UBound(workzones)
    color = workzones(i, 2)
    ' ligne 1
    ' la ligne 1 est laisse vide pour pouvoir compter la fin de mes colonnes du tableau de pointage
    ' et separee les colonnes de pointage des sous-totaux / totaux
    ' ligne 2
    With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(2, lastCol + i)
      .value = workzones(i, 1)
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .Interior.color = color
    End With
    ' ligne 3
    With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(3, lastCol + i)
      .value = "TOTAL"
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .Interior.color = color
    End With
    ' on definit la formule pour chaque ligne
    For j = firstRow + 1 To lastRow
      With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(j, lastCol + i)
        .Formula = getFormulaByRowByWorkzone(j, workzones(i, 1), firstCol, lastCol, SHEET_NAME_TOTAL_MONTH)
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = color
        .Interior.color = COLOR_CEL_READ_CONTENT
      End With
    Next j
  Next i
End Sub

Function getFormulaByRowByWorkzone(row As Integer, workzone As Variant, firstCol As Integer, lastCol As Integer, sheetName As String) As String
  Dim colArray() As Long

  ' SUMIF (range, criteria, [sum_range])
  getFormulaByRowByWorkzone = "=SUMIF(" & Cells(1, firstCol).Address & ":" & Cells(1, lastCol).Address & ",""" & workzone & """," & Replace(Cells(row, firstCol).Address, "$", "") & ":" & Replace(Cells(row, lastCol).Address, "$", "") & ")"
End Function

Sub addSumByRow()
  Dim firstColTotalByWorkzone As Integer
  Dim lastColTotalByWorkzone As Integer
  Dim firstRow As Integer
  Dim lastRow As Integer
  Dim i As Integer
  
  firstColTotalByWorkzone = utils_sheets.LastNumberRowNotEmpty(SHEET_NAME_TOTAL_MONTH, 1)
  lastColTotalByWorkzone = firstColTotalByWorkzone + UBound(workzones)
  ' on ajuste le depart de la colonne
  firstColTotalByWorkzone = firstColTotalByWorkzone + 1
  ' on nomme la colonne
  With Sheets(SHEET_NAME_TOTAL_MONTH).Range(Cells(2, lastColTotalByWorkzone + 1), Cells(3, lastColTotalByWorkzone + 1))
    .Merge
    .value = "TOTAL"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.color = COLOR_CEL_READ_CONTENT
  End With

  firstRow = 4
  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_TOTAL_MONTH, 3)
  ' on cree les formules
  For i = firstRow To lastRow
    With Sheets(SHEET_NAME_TOTAL_MONTH).Cells(i, lastColTotalByWorkzone + 1)
      .Formula = "=SUM(" & Replace(Cells(i, firstColTotalByWorkzone).Address, "$", "") & ":" & Replace(Cells(i, lastColTotalByWorkzone).Address, "$", "") & ")"
      .NumberFormat = "0.00"
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .Interior.color = COLOR_CEL_READ_CONTENT
    End With
  Next i

End Sub



