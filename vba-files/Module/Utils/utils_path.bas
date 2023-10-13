Attribute VB_Name = "utils_path"
'namespace=vba-files/Module/Utils
Option Explicit

' obtenir le chemin d'un fichier de type c:/.../.../...
Public Function getPathFolder(folder As String) As String
  Dim localPathPDF As String
  Dim localPath As String
  Dim arrPath() As String
  Dim i As Integer

  localPath = LibFileTools.GetLocalPath(thisWorkbook.Path)

  arrPath = Split(localPath, "\")

  For i = 0 To UBound(arrPath) - 1
    localPathPDF = localPathPDF & arrPath(i) & "\"
  Next i

  getPathFolder = localPathPDF & folder

End Function