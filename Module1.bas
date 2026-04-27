Attribute VB_Name = "Module1"
Option Explicit

' ----------------------------------------------------------------
' breedingSync.py - VBA bridge
' RunPythonWithWb из Module2 запускает Python асинхронно,
' ждёт _done.flag, затем читает _report.txt и показывает MsgBox.
' ----------------------------------------------------------------

Private Sub RunAndReport(ByVal action As String)
    Dim rc         As Long
    Dim reportFile As String
    Dim fNum       As Integer
    Dim line       As String
    Dim msg        As String

    reportFile = ThisWorkbook.Path & "\_report.txt"

    ' удалить старый отчёт перед запуском
    On Error Resume Next
    Kill reportFile
    On Error GoTo 0

    rc = Module2.RunPythonWithWb(action)

    ' читаем отчёт
    If Dir(reportFile) <> "" Then
        fNum = FreeFile
        Open reportFile For Input As #fNum
        Do While Not EOF(fNum)
            Line Input #fNum, line
            msg = msg & line & vbCrLf
        Loop
        Close #fNum
        On Error Resume Next
        Kill reportFile
        On Error GoTo 0
    End If

    If rc <> 0 Then
        If Len(msg) = 0 Then msg = "Python завершился с ошибкой или таймаут."
        MsgBox msg, vbCritical, "breedingSync — ошибка"
    Else
        If Len(msg) = 0 Then msg = "Готово."
        MsgBox msg, vbInformation, "breedingSync"
    End If
End Sub


' --- публичные процедуры (вешай на кнопки) ---

Public Sub ImportRegistry()
    RunAndReport "import_registry"
End Sub

Public Sub SubmitRegistry()
    RunAndReport "submit_registry"
End Sub

Public Sub SaveRegistry()
    RunAndReport "submit_registry"
End Sub
