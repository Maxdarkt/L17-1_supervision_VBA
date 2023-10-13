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
  
  ' Call general.main()
End Sub

Public Sub main()
  Call general.cleanCells()
  Call general.getAllExcelFiles()
End Sub

Public Sub cleanCells()
  Dim address As String
  Dim lastRow As Integer

  lastRow = utils_sheets.LastNumberColNotEmpty(SHEET_NAME_CONFIG, 11)

  address = "K4:K" & lastRow

  Sheets(SHEET_NAME_CONFIG).Range(address).ClearContents

End Sub

Public Sub getAllExcelFiles()
  Dim pathFolder As String
  Dim filename As String
  Dim pathFilename As String
  Dim i As Integer
  
  i = 0
  
  ' Spécifiez le chemin du dossier contenant les fichiers Excel
  pathFolder = utils_path.getPathFolder("reports\")
  
  filename = Dir(pathFolder & "*.xls*")
  ' Parcourez tous les fichiers Excel dans le dossier
  Do While filename <> ""
    ' Construisez le chemin complet du fichier
    pathFilename = pathFolder & filename
    ' On écrit les chemins des fichiers dans la feuille CONFIG
    Sheets(SHEET_NAME_CONFIG).Cells(5 + i, 11).Value = pathFilename
    ' Obtenez le prochain fichier Excel dans le dossier
    filename = Dir
    i = i + 1
  Loop
End Sub