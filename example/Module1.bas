Attribute VB_Name = "Module1"
Option Explicit

Public Sub RunPythonWithLogs(ByVal action As String)
    Dim py As String, script As String, wb As String
    Dim outLog As String, errLog As String
    Dim cmd As String

    ' Ïóòè (ïîïðàâü ïðè íåîáõîäèìîñòè)
    py = "C:\Python311\python.exe"
    script = "F:\tables\milkQuality_Forms.py"
    wb = ThisWorkbook.FullName

    outLog = "F:\tables\" & action & "_stdout.log"
    errLog = "F:\tables\" & action & "_stderr.log"

    ' cmd.exe íóæåí, ÷òîáû ðàáîòàëè 1>> è 2>>
    cmd = "cmd.exe /c " & Chr(34) & _
          Chr(34) & py & Chr(34) & " " & Chr(34) & script & Chr(34) & " " & action & " " & Chr(34) & wb & Chr(34) & _
          " 1>>" & Chr(34) & outLog & Chr(34) & " 2>>" & Chr(34) & errLog & Chr(34) & _
          Chr(34)

    Shell cmd, vbHide
End Sub

' Òåñòîâûé çàïóñê (ìîæíî äåðãàòü âðó÷íóþ)
Public Sub Test_submit_f5()
    RunPythonWithLogs "submit_f5"
End Sub


