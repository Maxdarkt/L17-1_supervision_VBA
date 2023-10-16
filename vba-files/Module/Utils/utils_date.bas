Attribute VB_Name = "utils_date"
'namepsace=vba-files/Module/Utils
Option Explicit

'LastDayOfMonth retourne le dernier jour d'un mois
'@Param month : mois en chiffre
'@Param year : annee en chiffre
Public Function LastDayOfMonth(month As Integer, year As Integer) As Date
  LastDayOfMonth = DateSerial(year, month + 1, 0)
End Function

'IsWorkingDay retourne :
'VRAI si c'est un jour travaille
'Faux si c'est un jour chome (Samedi ou Dimanche)
'Pas de verification de jour ferie
'@Param day : jour a analyser "05/04/2023"
Public Function IsWorkingDay(day As Date) As Boolean
  Dim response As Integer

  
  ' Debug.Print day
  response = Weekday(day)

  'WeekDay renvoie 1 pour Dimanche et 7 pour Samedi
  If response = 1 Or response = 7 Then
      IsWorkingDay = False
  Else
      IsWorkingDay = True
  End If
End Function

'Colore la feuille J1 si c'est un jour chome / travaille
'@Param firstDay : "01/07/2023"
Public Sub ColorTheFirstSheet(firstDay As Date)

  Dim varIsWorkingDay As Boolean
  
  varIsWorkingDay = IsWorkingDay(firstDay)
  
  If varIsWorkingDay = False Then
      Worksheets("J1").Tab.color = RGB(255, 0, 0)
  Else
      Worksheets("J1").Tab.color = RGB(51, 204, 204)
  End If
End Sub

Public Function CalculateDurationBetweenDates(firstDate As Date, lastDate As Date) As Integer
  Dim duration As Integer
  
  duration = DateDiff("d", firstDate, lastDate)
  
  CalculateDurationBetweenDates = duration

End Function
