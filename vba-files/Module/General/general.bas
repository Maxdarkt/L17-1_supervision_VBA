Attribute VB_Name = "general"
'namespace=vba-files/Module/General
Option Explicit

' VARIABLES GLOBALES
' NAME SHEET
Public Const SHEET_NAME_CONFIG As String = "CONFIG"
Public Const SHEET_NAME_CODE_ACTIVITIES As String = "CODE_ACTIVITES"
Public Const SHEET_NAME_TOTAL_MONTH As String = "TOTAL_MOIS"
Public Const SHEET_NAME_SEND_EMAIL As String = "ENVOIE_MAIL"

' COLOR CELL
Public COLOR_CEL_WHITE As Long
Public COLOR_CEL_READ_H1 As Long
Public COLOR_CEL_READ_H2 As Long
Public COLOR_CEL_READ_CONTENT As Long
Public COLOR_CEL_INPUT As Long
Public COLOR_CEL_TASK_CODE As Long
Public COLOR_CEL_TM_DAY_WORKED As Long
Public COLOR_CEL_TM_DAY_NOT_WORKED As Long
Public COLOR_TAB_DAY_WORKED As Long
Public COLOR_TAB_DAY_NOT_WORKED As Long

' COLOR FONT
Public COLOR_FONT_TITLE_SUM_H1 As Long
Public COLOR_FONT_TITLE_SUM_H2 As Long

' ACTIVITIES
Public ARR_SELECTED_SUBACTIVITIES AS Collection
Public ARR_SELECTED_SUBACTIVITIES_ID() AS Variant
Public ARR_CODE_ACTIVITIES_INFOS AS Collection


Public  Sub DefineGlobalVariables()
  
  ' On définit les variables couleurs
  COLOR_CEL_WHITE = RGB(255, 255, 255)
  COLOR_CEL_READ_H1 = RGB(166, 166, 166)
  COLOR_CEL_READ_H2 = RGB(191, 191, 191)
  COLOR_CEL_READ_CONTENT = RGB(231, 230, 230)
  COLOR_CEL_INPUT = RGB(180, 198, 231)
  COLOR_CEL_TASK_CODE = RGB(255, 204, 0)
  COLOR_TAB_DAY_WORKED = RGB(51, 204, 204)
  COLOR_TAB_DAY_NOT_WORKED = RGB(255, 0, 0)

  COLOR_FONT_TITLE_SUM_H1 = RGB(112, 48, 160)
  COLOR_FONT_TITLE_SUM_H2 = RGB(255, 0, 0)

  COLOR_CEL_TM_DAY_WORKED = RGB(204, 255, 204)
  COLOR_CEL_TM_DAY_NOT_WORKED = RGB(83, 141, 213)
  
End Sub

Public Sub initConfig()
  Call general.cleanCells()
  Call general.initFilesByWorkzone()
  Call general.initWorkedDays()
End Sub

Public Sub cleanCells()
  Dim address As String
  Dim lastRow As Integer

  address = "E5:G10"

  Sheets(SHEET_NAME_CONFIG).Range(address).ClearContents

  address = "K5:P16"

  Sheets(SHEET_NAME_CONFIG).Range(address).ClearContents

End Sub

' ancienne fonction pour aller récupérer les fichiers dans un dossier
' Public Sub getAllExcelFiles()
'   Dim pathFolder As String
'   Dim filename As String
'   Dim pathFilename As String
'   Dim i As Integer
  
'   i = 0
  
'   ' Spécifiez le chemin du dossier contenant les fichiers Excel
'   pathFolder = utils_path.getPathFolder("reports\")
  
'   filename = Dir(pathFolder & "*.xls*")
'   ' Parcourez tous les fichiers Excel dans le dossier
'   Do While filename <> ""
'     ' Construisez le chemin complet du fichier
'     pathFilename = pathFolder & filename
'     ' On écrit les chemins des fichiers dans la feuille CONFIG
'     Sheets(SHEET_NAME_CONFIG).Cells(5 + i, 11).Value = pathFilename
'     ' Obtenez le prochain fichier Excel dans le dossier
'     filename = Dir
'     i = i + 1
'   Loop
' End Sub

Sub initFilesByWorkzone()

  Dim wsSupervisionConfig As Worksheet
  Dim LastRow As Long, i As Long
  Dim FilePath As String
  Dim wbSite As Workbook, wsSite As Worksheet
  Dim firstRow As Integer

  firstRow = 5
  
  ' Référence à la feuille CONFIG du classeur de consolidation
  Set wsSupervisionConfig = ThisWorkbook.Sheets(SHEET_NAME_CONFIG)
  
  ' Trouver la dernière ligne dans la colonne "LIEN DES FICHIERS"
  LastRow = wsSupervisionConfig.Cells(wsSupervisionConfig.Rows.Count, 4).End(xlUp).Row
  
  For i = firstRow To LastRow
    FilePath = wsSupervisionConfig.Cells(i, 4).Value
    
    ' Vérifier si le chemin est valide
    If FilePath <> "" Or Len(Dir(FilePath)) > 0 Then
      ' Ouvrir le fichier
      Set wbSite = Workbooks.Open(FilePath)
      Set wsSite = wbSite.Sheets(SHEET_NAME_CONFIG)
      
      ' Lire les données et les mettre dans le classeur de consolidation
      wsSupervisionConfig.Cells(i, 5).Value = wsSite.Range("E36").Value
      wsSupervisionConfig.Cells(i, 6).Value = wsSite.Range("F22").Value
      wsSupervisionConfig.Cells(i, 7).Value = wsSite.Range("F24").Value
      wsSupervisionConfig.Cells(i, 8).Value = wsSite.Range("E32").Value
      
      ' Fermer le fichier de site
      wbSite.Close SaveChanges:=False
    Else
      ' Si le chemin n'est pas valide, insérer un message d'erreur dans les colonnes "OUVRAGE", "DEBUT", et "FIN"
      wsSupervisionConfig.Cells(i, 5).Value = "Erreur: Lien invalide"
      wsSupervisionConfig.Cells(i, 6).Value = "Erreur: Lien invalide"
      wsSupervisionConfig.Cells(i, 7).Value = "Erreur: Lien invalide"
      wsSupervisionConfig.Cells(i, 8).Value = "Erreur: Lien invalide"
    End If
  Next i
End Sub

Sub initWorkedDays()
  Dim wsSupervisionConfig As Worksheet
  Dim wbSite As Workbook, wsSite As Worksheet
  Dim filePath As String

  filePath = Sheets(SHEET_NAME_CONFIG).Range("D5").Value

  Set wbSite = Workbooks.Open(FilePath)
  Set wsSite = wbSite.Sheets(SHEET_NAME_CONFIG)

  ' Référence à la feuille CONFIG du classeur de consolidation
  Set wsSupervisionConfig = ThisWorkbook.Sheets(SHEET_NAME_CONFIG)

  wsSite.Range("C7:H18").Copy
  wsSupervisionConfig.Range("K5:P16").PasteSpecial xlPasteValues
  
  'Fermez le fichier du site
  wbSite.Close SaveChanges:=False
End Sub

Function getListWorkzones() As Variant
  Dim ws As Worksheet
  Dim listWorkzones() As Variant
  Dim lastRow As Long
  Dim i As Long
  
  ' Spécifiez le nom de la feuille
  Set ws = ThisWorkbook.Sheets(SHEET_NAME_CONFIG)
  
  ' Trouver la dernière ligne avec des données dans la colonne E (à partir de la cellule E5)
  lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
  
  ' Redimensionnez le tableau pour correspondre au nombre d'ouvrages
  ReDim listWorkzones(1 To lastRow - 4, 1 To 2) ' Soustrayez 4 pour exclure les 4 premières lignes
  
  ' Parcourez les cellules dans la colonne E à partir de la cellule E5
  For i = 5 To lastRow
      ' Stockez la valeur de la cellule E dans la colonne 1 du tableau (ouvrage)
      listWorkzones(i - 4, 1) = ws.Cells(i, "E").Value
      ' Stockez la valeur de la cellule I dans la colonne 2 du tableau (code couleur)
      listWorkzones(i - 4, 2) = ws.Cells(i, "I").Interior.Color
  Next i
  
  ' Renvoyer le tableau comme résultat de la fonction
  getListWorkzones = listWorkzones
End Function

Function getPositionWorkzonesInArray(workzone As String, workzones() As Variant) As Integer
  Dim i As Integer
  
  For i = LBound(workzones) To UBound(workzones)
    If workzones(i, 1) = workzone Then
      getPositionWorkzonesInArray = i
      Exit Function
    End If
  Next i
End Function

