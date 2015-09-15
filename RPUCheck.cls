VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RPUCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


''''''''''''''''''''''
'@Author Jürg Weber
'@Date 26.05.2015
'@Version 0.1
''''''''''''''''''''''

Public Sub Check()
    Dim strStellennr As String
    Dim firstCell As String
    Dim lastCell As String
    Dim checkSheet As Worksheet
    Dim gk As String
    Dim rpu As String
    Dim nStell As String
    Dim i As Integer
    Dim cell As Range
    Dim gkCol As Integer
    Dim rpuCol As Integer
    Dim nStellCol As Integer
    Dim textCol As Integer
    
    Application.DisplayAlerts = False
               
    Set checkSheet = ActiveWorkbook.Sheets(1)
    firstCell = checkSheet.Cells.Find("Stellennr.").Offset(1, 0).Address
    lastCell = Cells(Range("A4").End(xlDown).Row, checkSheet.Cells.Find("StellenNr.").Column).Address
    
    With checkSheet
        gkCol = .Cells.Find("TrfGr").Column
        rpuCol = .Cells.Find("RPU").Column
        textCol = .Cells.Find("Abgleich GK und RPU").Column
        nStellCol = .Cells.Find("Neu Stellennr.").Column
        For Each cell In Range(firstCell & ":" & lastCell)
             gk = .Cells(cell.Row, gkCol).Value
             rpu = .Cells(cell.Row, rpuCol).Value
             nStell = .Cells(cell.Row, nStellCol).Value
             If nStell <> "" Then
                .Cells(cell.Row, textCol).Value = checkRpuSheet(.Cells(cell.Row, nStellCol), gk, rpu)
             Else
                .Cells(cell.Row, textCol).Value = checkRpuSheet(cell, gk, rpu)
             End If
        Next cell
    End With
    
    Application.DisplayAlerts = True
End Sub

''Checks whether an gk fits a given documented gk with rpu.
''@cell, here stands the funktionsnummer
''@gk, the Gehaltsklasse to check
''@rpu, the rpu to check
Private Function checkRpuSheet(cell As Range, gk As String, rpu As String) As String

    Dim rpuSheet As Worksheet
    Dim checkRow As Integer
    Dim i As Integer
    Dim gkCell As Range
    Dim rpuCell As Range
    Dim textGKnotmatch As String
    Dim textGKfurthermatch As String
    Dim textGKmatch As String
    Dim checkCol As Integer
    
    Set rpuSheet = ActiveWorkbook.Sheets(2)
    On Error GoTo errormsg:
    checkRow = rpuSheet.Cells.Find("Funktionsnummer").EntireColumn.Find(cell.Value).Row
    i = rpuSheet.Cells.Find("GK").Column
    
    Do While Not IsEmpty(rpuSheet.Cells(checkRow, i).Value)
        
        Set gkCell = rpuSheet.Cells(checkRow, i)
        Set rpuCell = rpuSheet.Cells(checkRow, i + 1)
        
        If gk = gkCell.Value And rpu = rpuCell.Value Then
            textGKmatch = "Match: " & "[" & gkCell.Value & "," & rpuCell.Value & "] "
           
        ElseIf gk = gkCell.Value Then
            If textGKfurthermatch = "" Then
                textGKfurthermatch = "Weitere Möglichkeiten: " & "[" & gkCell.Value & "," & rpuCell.Value & "] "
                textGKnotmatch = "Kein Match, Möglichkeiten: " & "[" & gkCell.Value & "," & rpuCell.Value & "] "
            Else
                textGKfurthermatch = textGKfurthermatch & "[" & gkCell.Value & "," & rpuCell.Value & "] "
                textGKnotmatch = textGKnotmatch & "[" & gkCell.Value & "," & rpuCell.Value & "] "
            End If
        Else
            If textGKnotmatch = "" Then
                textGKnotmatch = "Kein Match [" & gk & "," & rpu & "], Möglichkeiten: " & "[" & gkCell.Value & "," & rpuCell.Value & "] "
            Else
                textGKnotmatch = textGKnotmatch & "[" & gkCell.Value & "," & rpuCell.Value & "] "
            End If
        End If
        i = i + 2
    Loop
    
    If textGKmatch <> "" Then
        checkRpuSheet = textGKmatch & textGKfurthermatch
    Else
        checkRpuSheet = textGKnotmatch
    End If
    Exit Function
errormsg:
    checkRpuSheet = "Stelle nicht vorhanden: " & cell.Value
End Function




