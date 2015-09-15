Attribute VB_Name = "mRichtposition"
''''''''''''''''''''''
'@Author Jürg Weber
'@Date 26.05.2015
'@Version 0.1
''''''''''''''''''''''


Option Explicit

''Checks the RPU's
Sub rpu()
    Dim rpu As New RPUCheck
    rpu.Check
    MsgBox "RPU check erfolgreich!"
End Sub

''Splits a file into sub files and save them into the subfolder "output". This subfolder is located at the currents workbooks path.
Sub split()
    Dim iohepler As New IOHelper
    Dim startRng As Range
    Set startRng = Cells.Find("Kurzbe. Org. Einheit")
    Call iohepler.sort(ActiveWorkbook.Sheets(1), startRng)
    Call iohepler.splitAndSave(ActiveWorkbook, startRng.Offset(1, 0), 1, 2)
    MsgBox "Split erfolgreich!"
End Sub

''Merges the splitted file
Sub merge()
    Dim currWB As Workbook
    Set currWB = ActiveWorkbook
    Dim iohepler As New IOHelper
    Call iohepler.loadIntoFile(currWB.path & "\", "\output\")
    MsgBox "Merge erfolgreich!"
End Sub
