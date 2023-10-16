Attribute VB_Name = "utils_worked_days"
'namepsace=vba-files/Module/Utils
Option Explicit

' Color un onglet si jour ferie nationale || projet
' Selectionne les champs sur la page config [1.1]
Public Function NotWorkedDays() As Date()
  ' Declaration de variables
  Dim days() As Date
  Dim i As Integer
  Dim Value As Variant
  i = 0

  ' Selection de la plage de cellules page CONFIG [1.1]
  For Each Value In Sheets(SHEET_NAME_CONFIG).range("K5:K16,N5:N16")
    ' Si la cellule n'est pas vide, on recupere la valeur dans le tableau
    If Not IsEmpty(Value) Then
      ' On ajoute une entree au tableau
      ReDim Preserve days(i)
      ' on stocke la valeur
      days(i) = Value

      'Debug.Print days(i)

      i = i + 1
  
    End If
  Next Value
  ' On stocke le tableau dans le même nom de variable de la fonction pour qu'elle retourne la valeur
  NotWorkedDays = days
End Function

' Est-ce un jour non travaille ?
' @Param day: Date
' @Param arrNotWorkedDays: Date()
' Retourne 0 si jour non travaille, 1 si jour travaille
Public Function IsDayNotWorked(day As Date, arrNotWorkedDays() As Date) As Integer
  ' Declaration de variables
  IsDayNotWorked = 0
  Dim i As Integer

  ' Debug.Print "Day is : " & day

  For i = LBound(arrNotWorkedDays) To UBound(arrNotWorkedDays)

    ' Debug.print arrNotWorkedDays(i)
    ' Si le jour courrant est egal a un jour ferie, on incremente de 1
    If arrNotWorkedDays(i) = day Then
      IsDayNotWorked = IsDayNotWorked + 1
    End If
  Next
End Function

' Coloration de la cellule si le jour n'est pas travaille (TOTAL MONTH)
Public Function colorCellIfDayIsNotWorked(row As Integer, col As Integer, color AS Long) As Boolean

  If color = cells(row, col).Interior.Color Then
    colorCellIfDayIsNotWorked = True
  Else
    colorCellIfDayIsNotWorked = False
  End If
End Function
