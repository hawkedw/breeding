Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' breeding / breedingSync.py - VBA bridge
' RunPythonWithWb: запускает Python, передаёт action + путь к книге.
' Shell не блокирует VBA, поэтому книга остаётся доступна Python
' через GetActiveObject. DoEvents внутри Python выполнять не нужно.
' ============================================================

Private Sub RunPythonWithWb(ByVal action As String)
    Dim py      As String
    Dim script  As String
    Dim wb      As String
    Dim outLog  As String
    Dim errLog  As String
    Dim cmd     As String
    Dim baseDir As String

    ' ---------- пути (правь при необходимости) ----------
    py     = "C:\Python311\python.exe"
    ' Скрипт лежит рядом с книгой:
    baseDir = ThisWorkbook.Path
    script  = baseDir & "\breedingSync.py"
    wb      = ThisWorkbook.FullName

    outLog = baseDir & "\" & action & "_stdout.log"
    errLog = baseDir & "\" & action & "_stderr.log"
    ' ----------------------------------------------------

    ' cmd /c нужен чтобы работали >> редиректы
    ' Внешние кавычки вокруг всей команды нужны для cmd /c
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


' ============================================================
' Публичные кнопки / горячие клавиши
' ============================================================

Public Sub ImportRegistry()
    RunPythonWithWb "import_registry"
End Sub

Public Sub SubmitRegistry()
    RunPythonWithWb "submit_registry"
End Sub
