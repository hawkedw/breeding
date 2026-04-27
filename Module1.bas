Attribute VB_Name = "Module1"
Option Explicit

' ----------------------------------------------------------------
' breedingSync.py - VBA bridge
' Shell не блокирует VBA => книга остаётся открытой и доступна
' Python цепляется к ней через GetActiveObject (win32com).
' ----------------------------------------------------------------

Private Sub RunPythonWithWb(ByVal action As String)
    Dim py      As String
    Dim script  As String
    Dim wb      As String
    Dim baseDir As String
    Dim outLog  As String
    Dim errLog  As String
    Dim cmd     As String

    ' --- пути (поправь py если Python в другом месте) ---
    py      = "C:\Python311\python.exe"
    baseDir = ThisWorkbook.Path
    script  = baseDir & "\breedingSync.py"
    wb      = ThisWorkbook.FullName

    outLog  = baseDir & "\" & action & "_stdout.log"
    errLog  = baseDir & "\" & action & "_stderr.log"
    ' ---------------------------------------------------

    ' cmd /c нужен чтобы работали >> редиректы
    ' внешние кавычки вокруг всей команды обязательны для cmd /c
    cmd = "cmd.exe /c " & Chr(34) & _
          Chr(34) & py & Chr(34) & " " & _
          Chr(34) & script & Chr(34) & " " & _
          action & " " & _
          Chr(34) & wb & Chr(34) & _
          " 1>>" & Chr(34) & outLog & Chr(34) & _
          " 2>>" & Chr(34) & errLog & Chr(34) & _
          Chr(34)

    Shell cmd, vbHide
End Sub


' --- публичные процедуры (вешай на кнопки) ---

Public Sub ImportRegistry()
    RunPythonWithWb "import_registry"
End Sub

Public Sub SubmitRegistry()
    RunPythonWithWb "submit_registry"
End Sub
