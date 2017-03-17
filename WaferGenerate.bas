Attribute VB_Name = "WaferGenerate"
Option Explicit

Const RowBegin = 3
Const ColumnBegin = 2
Const RowEnd = 287
Const ColumnEnd = 53
Const DataRowBegin = 2
Const DataColumn = 2
Const DataRowEnd = 111
Const Data20SitesRowBegin = 2
Const Data20SitesColumn = 6
Const Data20SitesRowEnd = 12

Sub QuickClean()
Attribute QuickClean.VB_ProcData.VB_Invoke_Func = " \n14"
'
' QuickClean ¥¨¶°
'

'
    Sheets("Reference").Select
    Range("A1:BB288").Select
    Selection.Copy
    Sheets("Wafermap").Select
    Range("A1").Select
    ActiveSheet.Paste
End Sub

Sub WaferCellLocationMap()

    MsgBox ("X: " & ActiveCell.Cells.Column - 1 & ", Y: " & ActiveCell.Cells.Row - 1)

End Sub

Sub DrawDualSiteLocations()

    Dim RowCount As Long
    Dim XLoc, YLoc As Long
    
    Worksheets("Wafermap").Activate
    For RowCount = DataRowBegin To DataRowEnd
        XLoc = Worksheets("Location Tables").Cells(RowCount, DataColumn).Value
        YLoc = Worksheets("Location Tables").Cells(RowCount, DataColumn + 1).Value
        Cells(YLoc + 1, XLoc + 1).Interior.ColorIndex = 5
        Cells(YLoc + 1, XLoc + 2).Interior.ColorIndex = 5
    Next RowCount

End Sub

Sub Draw20SitesLocation()

    Dim RowCount As Long
    Dim XLoc, YLoc As Long
    Dim ColumnTemp, RowTemp As Long
    
    Worksheets("Wafermap").Activate
    For RowCount = Data20SitesRowBegin To Data20SitesRowEnd
        XLoc = Worksheets("Location Tables").Cells(RowCount, Data20SitesColumn).Value
        YLoc = Worksheets("Location Tables").Cells(RowCount, Data20SitesColumn + 1).Value
        For ColumnTemp = 0 To 3
            For RowTemp = 0 To 4
                Cells(YLoc + 1 + RowTemp, XLoc + 1 + ColumnTemp).Interior.ColorIndex = 6
            Next RowTemp
        Next ColumnTemp
    Next RowCount

End Sub

Sub ActiveCellProperty()

    MsgBox "Value: " & ActiveCell.Interior.ColorIndex

End Sub

Sub CleanCell()

    ActiveCell.Value = Null

End Sub

Sub CellLocation()

    MsgBox ActiveCell.Cells.Row & " " & ActiveCell.Cells.Column

End Sub

Sub CreateCleanMap()

    Dim X_count, Y_count As Long
    
    Worksheets("Wafermap").Activate
    For X_count = RowBegin To RowEnd
        For Y_count = ColumnBegin To ColumnEnd
            Cells(X_count, Y_count).Select
            If ActiveCell.Interior.ColorIndex <> 15 Then
                ActiveCell.Interior.ColorIndex = 4
                ActiveCell.Value = Null
            End If
        Next Y_count
    Next X_count
    
    Cells(1, 1).Select
    
End Sub

