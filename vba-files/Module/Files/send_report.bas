Attribute VB_Name = "send_report"
'namespace=vba-files/Module/Files
Option Explicit

Dim localPathsPdf() As String

Public Sub sendReportsByEmail()
  Dim allPathsAreCorrects As Integer
  ' 1 - get all path of pdf
  Call send_report.getAllPathOfPDF()
  ' 2 - check if all path pdf are corrects
  allPathsAreCorrects = send_report.checkIfAllPathPdfAreCorrects()

  If allPathsAreCorrects > 0 Then
    MsgBox "Un chemin de fichier est incorrect",,"Alerte"
    Exit Sub
  End If
  ' 3 - get all information for edit a mail
  Call send_report.getAllInformationsForEditMail()
End Sub

Private Sub getAllPathOfPDF()
  Dim plage As Range
  Dim cell As Range
  Dim localPaths As String
  Dim i As Integer

  i = 0

  For Each cell In Sheets(SHEET_NAME_SEND_EMAIL).Range("C11:C25")
    If cell.Value <> "" Then
      ReDim Preserve localPathsPdf(i)
      localPathsPdf(i) = cell.Value
      i = i + 1
    End If
  Next cell 

End Sub

Private Function checkIfAllPathPdfAreCorrects() As Integer
  Dim i As Integer
  Dim allPathsAreCorrects As Integer

  allPathsAreCorrects = 0

  For i = 0 To UBound(localPathsPdf)
    If Dir(localPathsPdf(i)) = "" Then
      allPathsAreCorrects = allPathsAreCorrects + 1
    End If
  Next i
  checkIfAllPathPdfAreCorrects = allPathsAreCorrects
End Function

Private Sub getAllInformationsForEditMail()
  Dim expediteur As String
  Dim subject As String
  Dim body As String
  Dim localPathPDF As String
  Dim recipients As String
  Dim recipientsCc As String

  expediteur = Sheets(SHEET_NAME_SEND_EMAIL).Range("O29").Value
  subject = Sheets(SHEET_NAME_SEND_EMAIL).Range("O31").Value
  body = Sheets(SHEET_NAME_SEND_EMAIL).Range("M36").Value

  recipients = send_report.getRecipients()
  recipientsCc = send_report.getRecipientsCc()

  Call send_report.sendEmail(expediteur, recipients, recipientsCc, subject, body)
End Sub

' get all recipients in string
Private Function getRecipients() As String
  Dim recipients As String
  Dim row As Range
  Dim i As Integer

  For Each row In Sheets(SHEET_NAME_SEND_EMAIL).Range("Q5:R24").rows
    For i = 1 To row.Cells.Count
      If i = 1 Then 
        If row.Cells(i + 1).Value = "A" Then
          recipients = recipients & row.Cells(i).Value & ";"
        End If
      End If
    Next i 
  Next row

  getRecipients = recipients

End Function

' get all recipientsCc in string
Private Function getRecipientsCc() As String
  Dim recipients As String
  Dim row As Range
  Dim i As Integer

  For Each row In Sheets(SHEET_NAME_SEND_EMAIL).Range("Q6:R16").rows
    For i = 1 To row.Cells.Count
      If i = 1 Then 
        If row.Cells(i + 1).Value = "CC" Then
          recipients = recipients & row.Cells(i).Value & ";"
        End If
      End If
    Next i 
  Next row

  getRecipientsCc = recipients

End Function

Private Sub sendEmail(expediteur As String, recipients As String, recipientsCc As String, subject As string, body As String)
  Dim outlook As Object
  Dim outlookMail As Object
  Dim i As Integer

  Set outlook = CreateObject("Outlook.Application")
  Set outlookMail = outlook.CreateItem(0)

  On Error Resume Next
  With outlookMail
    .SentOnBehalfOfName = expediteur
    .To = recipients
    .CC = recipientsCc
    .Subject = subject
    .Body = body
    For i = 0 To UBound(localPathsPdf)
      .Attachments.Add localPathsPdf(i)
    Next i
    .Display   'or use .Send
  End With
  On Error GoTo 0

  Set outlookMail = Nothing
  Set outlook = Nothing
End Sub