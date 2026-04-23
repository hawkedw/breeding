VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim nm As String, hdrRow As Long, dirtyCol As Long
    nm = sh.Name
    If nm <> "Форма 1" And nm <> "Форма 2" And nm <> "Форма 5" Then Exit Sub

    hdrRow = IIf(nm = "Форма 5", 3, 1)

    On Error Resume Next
    dirtyCol = sh.Rows(hdrRow).Find(What:="Dirty", _
                                    LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo 0
    If dirtyCol = 0 Then Exit Sub

    If Target.Row <= hdrRow Then Exit Sub
    If Not Intersect(Target, sh.Columns(dirtyCol)) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    sh.Cells(Target.Row, dirtyCol).Value = True
    Application.EnableEvents = True
End Sub


