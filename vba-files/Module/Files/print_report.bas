Attribute VB_Name = "print_report"
'namespace=vba-files/Module/Files
Option Explicit

' Action button
Public Sub generatePDF()

  ' 0 - clean plage
  Call print_report.cleanGeneratePDF()
  
  ' 1 - generer les dates et les ovurages
  Call print_report.generateAllReportsInPDF()

End Sub

Sub cleanGeneratePDF()
  Dim address As String

  address = "C11:I25"

  With Sheets(SHEET_NAME_SEND_EMAIL).Range(address)
  .Value = ""
  .Interior.Color = COLOR_CEL_READ_CONTENT
  .Font.Bold = False
  .HorizontalAlignment = xlLeft
  .VerticalAlignment = xlCenter
  .IndentLevel = 0
  End With

End Sub

Sub generateAllReportsInPDF()
  Dim firstRowFiles As Integer
  Dim lastRowFiles As Integer
  Dim i As Integer
  Dim wbSite As Workbook
  Dim wsConfig As Worksheet
  Dim firstDay As Date
  Dim lastDay As Date
  Dim nbDays As Integer
  Dim workzone As String

  ' Référencer la feuille de configuration
  Set wsConfig = ThisWorkbook.Sheets(SHEET_NAME_CONFIG)
  ' Définir les dates
  firstDay = Sheets(SHEET_NAME_SEND_EMAIL).Range("D4").Value
  lastDay = Sheets(SHEET_NAME_SEND_EMAIL).Range("H4").Value
  ' Nombre de jours entre 2 dates
  nbDays = utils_date.CalculateDurationBetweenDates(firstDay, lastDay)

  firstRowFiles = 5

  lastRowFiles = LastNumberColNotEmpty(SHEET_NAME_CONFIG, 4)
  
  For i = firstRowFiles To lastRowFiles
    Dim filePathExcel As String
    Dim filePathPDF As String
    ' Désactiver les messages d'alerte
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    ' On définit le chemin d'accès au fichier excel et le nom de l'ouvrage
    filePathExcel = wsConfig.Cells(i, 4).Value
    workzone = wsConfig.Cells(i, 5).Value
    ' Open site workbook
    Set wbSite = Workbooks.Open(filePathExcel)
    ' on génère le rapport de poste et on reçoit le chemin d'accès au PDF
    filePathPDF = printManyReportsInPDF(wbSite, firstDay, lastDay, nbDays, workzone)
    ' On écrit le nom du fichier dans la colonne C
    ThisWorkbook.Sheets(SHEET_NAME_SEND_EMAIL).Cells(i + 6, 3).Value = filePathPDF
    ' Close site workbook
    wbSite.Close SaveChanges:=False
    ' Activer les messages d'alerte
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
  Next i

End Sub

' Imprimer le rapport de poste
Public Function printManyReportsInPDF(wbSite As Workbook, firstDay As Date, lastDay As Date, nbDays As Integer, workzone As String) As String
  Dim newWorkbook As Workbook
  Dim pdfFilename As String
  Dim outputPath As String
  Dim yearOfReport As String
  Dim nbWeek As Integer
  Dim address As String
  Dim i As Integer

  'Spécidier le nom du fichier PDF temporaire
  pdfFilename = "TempPDF.pdf"
  ' spécifier l'année
  yearOfReport = Format(firstDay, "yyyy")
  ' spécifier le numéro de semaine
  nbWeek = DatePart("ww", firstDay, vbMonday)
  ' Spécifiez le chemin de sortie pour le PDF combiné
  outputPath = print_report.getPathExportPDF("PDF\") & yearOfReport & "_S-" & nbWeek & "_Rapport_de_poste_OA" & workzone & ".pdf"
  ' print area
  address = "AA3:AR107"

  ' Créez un nouveau classeur temporaire pour imprimer les feuilles
  Workbooks.Add

  ' on boucle sur le nombre de jours
  For i = nbDays To 0 Step -1
    Dim arrDate() As String
    Dim dateOfSheet As Date
    Dim formatedDate As String
    Dim ws As Worksheet
    ' ici on reformatte la date pour le nom du fichier
    arrDate = Split(firstDay, ".")
    dateOfSheet = DateSerial(arrDate(2), arrDate(1), arrDate(0))
    dateOfSheet = DateAdd("d", i, firstDay)
    formatedDate = Format(dateOfSheet, "yyyy-mm-dd")

    Set ws = wbSite.Sheets("J" & Day(dateOfSheet))

    ws.Copy Before:=ActiveWorkbook.Sheets(1)
        
    ' Définissez la zone d'impression identique sur chaque feuille
    ActiveSheet.PageSetup.PrintArea = address ' Personnalisez cette plage selon vos besoins

  Next i

  ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
    
  ' Fermez le classeur temporaire sans enregistrer les modifications
  ActiveWorkbook.Close SaveChanges:=False
  ' Vérifier si le fichier existe déjà
  If Dir(outputPath) <> "" Then
      ' Supprimez le fichier existant s'il existe
      Kill outputPath
  End If
  ' Renommez le fichier PDF temporaire en PDF final
  Name pdfFileName As outputPath

  printManyReportsInPDF = outputPath

End Function


' obtenir le chemin d'un fichier de type c:/.../.../...
Public Function getPathExportPDF(folder As String)
  Dim localPathPDF As String
  Dim localPath As String
  Dim arrPath() As String
  Dim i As Integer

  localPath = LibFileTools.GetLocalPath(thisWorkbook.Path)

  arrPath = Split(localPath, "\")

  For i = 0 To UBound(arrPath) - 1
    localPathPDF = localPathPDF & arrPath(i) & "\"
  Next i

  getPathExportPDF = localPathPDF & folder

End Function