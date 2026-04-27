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
    Dim cmd     As String

    ' --- пути (поправь py если Python в другом месте) ---
    py     = "C:\Python311\python.exe"
    script = ThisWorkbook.Path & "\breedingSync.py"
    wb     = ThisWorkbook.FullName
    ' ---------------------------------------------------

    cmd = Chr(34) & py & Chr(34) & " " & _
          Chr(34) & script & Chr(34) & " " & _
          action & " " & _
          Chr(34) & wb & Chr(34)

    Shell "cmd.exe /c " & cmd, vbHide
End Sub


' --- публичные процедуры (вешай на кнопки) ---

Public Sub ImportRegistry()
    RunPythonWithWb "import_registry"
End Sub

Public Sub SubmitRegistry()
    RunPythonWithWb "submit_registry"
End Sub

Public Sub SaveRegistry()
    RunPythonWithWb "submit_registry"
End Sub
