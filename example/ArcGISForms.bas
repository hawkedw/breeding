Attribute VB_Name = "ArcGISForms"
Option Explicit

' Путь к интерпретатору Python
Private Const PYTHON_EXE As String = "C:\Python311\python.exe"

' Имя .py-скрипта рядом с .xlsm
Private Const SCRIPT_NAME As String = "milkQuality_Forms.py"

Private Function RunPython(action As String) As Long
    Dim wbPath As String
    Dim scriptPath As String
    Dim cmd As String
    
    wbPath = ThisWorkbook.FullName
    scriptPath = ThisWorkbook.Path & "\" & SCRIPT_NAME
    
    If Dir(scriptPath, vbNormal) = "" Then
        MsgBox "Не найден Python-скрипт:" & vbCrLf & scriptPath, vbCritical, "milkQuality_Forms"
        RunPython = -1
        Exit Function
    End If
    
    If Dir(PYTHON_EXE, vbNormal) = "" Then
        MsgBox "Не найден интерпретатор Python:" & vbCrLf & PYTHON_EXE, vbCritical, "milkQuality_Forms"
        RunPython = -1
        Exit Function
    End If
    
    cmd = """" & PYTHON_EXE & """ " & _
          """" & scriptPath & """ " & _
          action & " " & _
          """" & wbPath & """"
    
    Debug.Print "RunPython:", cmd
    
    On Error Resume Next
    RunPython = Shell(cmd, vbNormalFocus)
    If Err.Number <> 0 Then
        MsgBox "Ошибка запуска Python:" & vbCrLf & Err.Description, vbCritical, "milkQuality_Forms"
        Err.Clear
    End If
End Function

' ====== Обработчик кнопок / макросы ======

Public Sub Import_Form1()
    DeleteSheetIfExists "Форма 1"
    RunPython "import_f1"
End Sub

Public Sub Submit_Form1()
    RunPython "submit_f1"
End Sub

Public Sub Import_Form2()
    DeleteSheetIfExists "Форма 2"
    RunPython "import_f2"
End Sub

Public Sub Submit_Form2()
    RunPython "submit_f2"
End Sub

Public Sub Import_Form5()
    DeleteSheetIfExists "Форма 5"
    RunPython "import_f5"
End Sub

Public Sub Submit_Form5()
    RunPython "submit_f5"
End Sub

' Удаление листа без подтверждения
Private Sub DeleteSheetIfExists(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub


