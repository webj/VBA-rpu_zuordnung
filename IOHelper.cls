VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''
'
'Helper class for io events.
'
'@Author Jürg Weber
'@Date 04.06.2015
'@Version 0.2
''''''''''''''''''''''

Option Explicit

''@path where the merged file should be saved
''@foldername directory where the splited files are
Public Sub loadIntoFile(path As String, foldername As String)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim insertWB As Workbook
    Dim insertWS As Worksheet
    Dim tmpWB As Workbook
    Dim tmpWS As Worksheet
    Dim startRow As Integer
    Dim startCol As Integer
    Dim endRow As Integer
    Dim endCol As Integer
    Dim cpyRange As Range
    Dim insertRange As Range
    Dim headercell As Range
    Dim file As String
    Dim first As Boolean
    Dim mytitle As String
    Set insertWB = Workbooks.Add
    Set insertWS = insertWB.Worksheets(1)
    
    startRow = 4
    startCol = 1
    file = Dir(path & foldername)
    first = True
    
    While file <> ""
    
        Set tmpWB = Workbooks.Open(path & foldername & file)
        Set tmpWS = tmpWB.Worksheets(1)
        
        'remove autofilter
        If tmpWS.AutoFilterMode Then
            tmpWS.Rows.AutoFilter
        End If
        
        'add title
        If first Then
            Dim endAddress As String
            Set headercell = tmpWS.Range("A3")
            endAddress = headercell.End(xlToRight).Address
            Set cpyRange = tmpWS.Range(headercell, headercell.End(xlToRight))
            Set insertRange = insertWS.Range("A3", endAddress)
            Call copyRange(cpyRange, insertRange)
            Call copyColumWidth(cpyRange, insertRange)
            mytitle = tmpWS.Range("A1").Value
            With insertWS.Range("A1")
                .Value = Mid(mytitle, InStr(mytitle, "-") + 2)
                .Font.Size = 15
                .Font.Bold = True
            End With
            insertWS.Range("A2").Value = tmpWS.Range("A2").Value
            Call setHeader(insertWS, tmpWS)
            Call setFooter(insertWS, tmpWS)
            Call setPrintSettings(insertWS, tmpWS)
            first = False
        End If
            
        endRow = tmpWS.Cells(Rows.Count, 1).End(xlUp).Row
        endCol = tmpWS.UsedRange.Columns.Count

        Set cpyRange = tmpWS.Range(tmpWS.Cells(4, 1), tmpWS.Cells(endRow, endCol))
        Set insertRange = insertWS.Range(insertWS.Cells(startRow, startCol), insertWS.Cells((startRow + endRow - 4), (endCol)))
        Call copyRange(cpyRange, insertRange)
        Call copyRowWidth(cpyRange, insertRange)
        startRow = startRow + endRow - 3
        file = Dir
        tmpWB.Close
    Wend
    
    
    'delete empty columns
    Dim currRow As Range
    endRow = insertWS.Cells(Rows.Count, 1).End(xlUp).Row
    Set currRow = insertWS.Range("A4")
    
    Do While currRow.Row <= endRow
        If IsEmpty(currRow) Then
            Set currRow = currRow.Offset(1, 0)
            currRow.Offset(-1, 0).EntireRow.Delete
            endRow = endRow - 1
        Else
            Set currRow = currRow.Offset(1, 0)
        End If
        
    Loop
    insertWS.Select
    ActiveWindow.Zoom = 70
    insertWB.SaveAs (path & "merged.xlsx")
    insertWB.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

''Splits and saves a worksheet into many worksheets. To apply this function the list must be sorted by the splitting column.
''@myws, worksheet we wanna split
''@startcell, according this cell, we split and save
''@start, start of the compared substring
''@length, length of the comparing substring
Public Sub splitAndSave(mywb As Workbook, startcell As Range, start As Integer, length As Integer)
    Dim myws As Worksheet
    Dim checkRange As Range
    Dim cell As Range
    Dim headercell As Range
    Dim myoffset As Integer
    Application.ScreenUpdating = False
    myoffset = -1
    Set myws = mywb.Sheets(1)
    Set checkRange = Range(startcell, startcell.End(xlDown))
    Set headercell = startcell.Offset(-1, 0)
    For Each cell In checkRange
        If Mid(cell.Value, start, length) <> Mid(startcell.Value, start, length) Or IsEmpty(cell.Offset(1, 0)) Then
            If IsEmpty(cell.Offset(1, 0)) Then
                myoffset = 0
            End If
            Call copyRngInNewWB(myws, startcell, cell.Offset(myoffset, 0), headercell, getSaveTxt(mywb.Sheets(3), Mid(startcell.Value, start, length)))
            Set startcell = cell
        End If
    Next
    Application.ScreenUpdating = True
End Sub

''@myws, worksheet to sort
''@sortcell, start of sort range
Public Sub sort(myws As Worksheet, sortcell As Range)
    If Not myws.AutoFilterMode Then
        sortcell.AutoFilter
    End If
    myws.AutoFilter.sort.SortFields.Clear
    myws.AutoFilter.sort.SortFields.Add Key:=Range _
        (sortcell.Address), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With myws.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

''@myws, worksheet where copied data comes from
''@startcell, cell of copy start range
''@endcell, cell of copy end range
''@headercell, startcell with header infomration
''@title, title of the workbook
Private Function copyRngInNewWB(myws As Worksheet, startcell As Range, endcell As Range, headercell As Range, title As String)
    Application.ScreenUpdating = False
    Dim newWB As Workbook
    Dim newws As Worksheet
    Dim rowOffset As Integer
    Dim colOffset As Integer
    Dim cpyRng As Range
    Dim insertRng As Range
    Dim path As String
    path = ActiveWorkbook.path
    path = path & "/output"
    
    Set newWB = Workbooks.Add
    Set newws = newWB.Sheets(1)
    
    rowOffset = endcell.Row - startcell.Row
    colOffset = headercell.End(xlToRight).Column - startcell.Column
    
    Set cpyRng = myws.Range(headercell, headercell.End(xlToRight))
    Set insertRng = newws.Range(newws.Cells(3, 1), newws.Cells(3, 1 + colOffset))
    Call copyRange(cpyRng, insertRng)
    Call copyColumWidth(cpyRng, insertRng)
    
    Set cpyRng = myws.Range(startcell, myws.Cells(endcell.Row, headercell.End(xlToRight).Column))
    Set insertRng = newws.Range(newws.Cells(4, 1), newws.Cells(4 + rowOffset, 1 + colOffset))
    Call copyRange(cpyRng, insertRng)
    Call copyRowWidth(cpyRng, insertRng)
    
    Call setSheetTitle(newws, myws, title)
    Call setHeader(newws, myws)
    Call setFooter(newws, myws)
    Call setPrintSettings(newws, myws)

    Call createDirectory(path, "")
    newWB.SaveAs (path & "/" & title & ".xlsx")
    newWB.Close
    Application.ScreenUpdating = True
End Function

''Sets the title of a worksheet
''@myws worksheet which needs a title
''@title title of the worksheet
Private Function setSheetTitle(newws As Worksheet, myws As Worksheet, title As String)
    With newws.Range("A1")
        .Value = title & " - " & myws.Range("A1").Value
        .Font.Size = 15
        .Font.Bold = True
    End With
    With newws.Range("A2")
        .Value = myws.Range("A2").Value
        .Font.Size = 10
        .Font.Bold = True
    End With
    newws.Select
    ActiveWindow.Zoom = 70
End Function

''Sets the header of a given worksheet, according another worksheet
''@newws the sheet where the header has to be set
''@myws the sheet where the header settings comes from
Private Function setHeader(newws As Worksheet, myws As Worksheet)
    With newws.PageSetup
        .LeftHeader = myws.PageSetup.LeftHeader
        .CenterHeader = myws.PageSetup.CenterHeader
        .RightHeader = myws.PageSetup.RightHeader
    End With
End Function

''Sets the footer of a given worksheet, according another worksheet
''@newws the sheet where the footer has to be set
''@myws the sheet where the footer settings comes from
Private Function setFooter(newws As Worksheet, myws As Worksheet)
    With newws.PageSetup
        .LeftFooter = myws.PageSetup.LeftFooter
        .CenterFooter = myws.PageSetup.CenterFooter
        .RightFooter = myws.PageSetup.RightFooter
    End With
End Function

''Sets landscape and printsettings accoriding a given worksheet
''@newws the worksheet which has to be layoutet
''@myws, the worksheets where the layout comes from
Private Function setPrintSettings(newws As Worksheet, myws As Worksheet)
Application.PrintCommunication = False
    With newws.PageSetup
        .Orientation = myws.PageSetup.Orientation
        .FitToPagesWide = myws.PageSetup.FitToPagesWide
        .FitToPagesTall = myws.PageSetup.FitToPagesTall
    End With
Application.PrintCommunication = True
End Function

''Copies a chosen range from A to B. Furthermore it manages the style such as Font and Interior settings.
''@fromRng from where to copy
''@toRng target range
Private Function copyRange(fromRng As Range, toRng As Range)
    Dim tmpVal
    tmpVal = toRng
    fromRng.Copy toRng
End Function

''Copies the column width of one range to anoter
''@fromRng, the range from where the settincs comes
''@toRng, the range where the settings will be copied
Private Function copyColumWidth(fromRng As Range, toRng As Range)
    Dim i As Integer
    For i = 1 To fromRng.Cells.Count
        toRng.Cells(i).EntireColumn.ColumnWidth = fromRng.Cells(i).EntireColumn.ColumnWidth
    Next i
End Function

''Copies the row width from one range to another range.
''@fromRng, the range where the settings comes from
''@toRng, the row width which has to be set
Private Function copyRowWidth(fromRng As Range, toRng As Range)
    toRng.RowHeight = fromRng.RowHeight
End Function

''Creates a directory at a given path. The directory will only be created if it not already exists.
''@path the path to the location where the directory will be created
''@name the name of the directory to create
Private Function createDirectory(path As String, name As String)
    If Dir(path & "/" & name, vbDirectory) = "" Then
        MkDir path & "/" & name
    End If
End Function

''Returns the saving name of a splitted workbook.
''@myws, the worksheet where the saving names are in
''@number, the number which we look after
Private Function getSaveTxt(myws As Worksheet, number As String) As String
    getSaveTxt = myws.Cells.Find(number).Value
End Function


