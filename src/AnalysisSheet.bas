Sub AnalysisSheet()
    ' This macro sets up an Excel worksheet for analysis, identifies duplicates, and organizes key columns
    ' for further data analysis.

    Dim ws As Worksheet, dataSourceSheet As Worksheet
    Dim lastCol As Long, lastRow As Long, i As Long, row As Long
    Dim areaCol As Long, plantCol As Long, dateCol As Long
    Dim remarksCol As Long, categoryCol As Long, timeCol As Long
    Dim keyMetricValue As Double, categoryLabel As String, dynamicHeader As String
    Dim duplicateFound As Boolean, userResponse As VbMsgBoxResult
    
    ' Set references to the active worksheet and a data source worksheet, assumed to be named "DataSource"
    Set ws = ActiveSheet
    On Error Resume Next
    Set dataSourceSheet = ThisWorkbook.Sheets("DataSource")
    On Error GoTo 0
    
    If dataSourceSheet Is Nothing Then
        MsgBox "Data source sheet not found. Please ensure a sheet named 'DataSource' is available.", vbExclamation
        Exit Sub
    End If
    
    ' Find the last column with data in the active sheet
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Identify important columns by header name in row 1 (generalizable to other datasets)
    For i = 1 To lastCol
        Select Case Trim(UCase(ws.Cells(1, i).Value))
            Case "AREA": areaCol = i
            Case "PLANT": plantCol = i
            Case "DATE": dateCol = i
            Case "REMARKS": remarksCol = i
            Case "CATEGORY": categoryCol = i
            Case Else
                If InStr(1, ws.Cells(1, i).Value, "Time", vbTextCompare) > 0 And _
                   InStr(1, ws.Cells(1, i).Value, "hours", vbTextCompare) > 0 Then
                    timeCol = i
                End If
        End Select
    Next i
    
    ' Duplicate check across specified columns to identify and highlight repeated entries
    duplicateFound = False
    For row = 2 To ws.Cells(ws.Rows.Count, areaCol).End(xlUp).Row
        For i = row + 1 To ws.Cells(ws.Rows.Count, areaCol).End(xlUp).Row
            If ws.Cells(row, areaCol).Value = ws.Cells(i, areaCol).Value And _
               ws.Cells(row, plantCol).Value = ws.Cells(i, plantCol).Value And _
               ws.Cells(row, dateCol).Value = ws.Cells(i, dateCol).Value And _
               ws.Cells(row, remarksCol).Value = ws.Cells(i, remarksCol).Value And _
               ws.Cells(row, categoryCol).Value = ws.Cells(i, categoryCol).Value Then
                duplicateFound = True
                ws.Range(ws.Cells(i, areaCol), ws.Cells(i, 53)).Interior.Color = RGB(255, 200, 200) ' Highlight duplicate up to column BB
            End If
        Next i
    Next row
    
    ' Prompt the user if duplicates are found, with options to remove or keep duplicates
    If duplicateFound Then
        userResponse = MsgBox("Duplicate entries found. Would you like to keep only one of each duplicate and remove others?", vbQuestion + vbYesNoCancel, "Duplicate Entries")
        If userResponse = vbYes Then
            ' Remove duplicates by keeping only the first instance of each duplicate
            For row = ws.Cells(ws.Rows.Count, areaCol).End(xlUp).Row To 2 Step -1
                For i = row - 1 To 2 Step -1
                    If ws.Cells(row, areaCol).Value = ws.Cells(i, areaCol).Value And _
                       ws.Cells(row, plantCol).Value = ws.Cells(i, plantCol).Value And _
                       ws.Cells(row, dateCol).Value = ws.Cells(i, dateCol).Value And _
                       ws.Cells(row, remarksCol).Value = ws.Cells(i, remarksCol).Value And _
                       ws.Cells(row, categoryCol).Value = ws.Cells(i, categoryCol).Value Then
                        ws.Rows(row).Delete
                        Exit For
                    End If
                Next i
            Next row
            MsgBox "Duplicates removed. Only the first instance of each duplicate is kept.", vbInformation
        ElseIf userResponse = vbNo Then
            MsgBox "All duplicates have been left as they are and highlighted for your review.", vbInformation
        Else
            Exit Sub ' Cancel operation if the user selects "Cancel"
        End If
    End If

    ' Hide unnecessary columns, keeping only relevant columns for analysis
    For i = 1 To lastCol
        If i <> areaCol And i <> plantCol And i <> dateCol And i <> remarksCol And _
           i <> categoryCol And i <> timeCol Then
            ws.Columns(i).Hidden = True
        End If
    Next i

    ' Insert new columns for Key, % Loss Time, and Dynamic Column
    ws.Cells(1, timeCol + 1).EntireColumn.Insert
    ws.Cells(1, timeCol + 1).Value = "Key" ' General name for additional analysis column
    ws.Cells(1, timeCol + 2).EntireColumn.Insert
    ws.Cells(1, timeCol + 2).Value = "% Loss Time"
    
    ' Retrieve metric value from data source and set dynamic header
    categoryLabel = ws.Cells(2, categoryCol).Value ' Retrieve first entry in Category for use in analysis
    On Error Resume Next
    keyMetricValue = Application.WorksheetFunction.VLookup(categoryLabel, dataSourceSheet.Range("A:B"), 2, False)
    On Error GoTo 0
    If IsError(keyMetricValue) Or keyMetricValue = 0 Then
        MsgBox "Category '" & categoryLabel & "' not found in data source sheet.", vbExclamation
        Exit Sub
    End If
    dynamicHeader = "% of " & categoryLabel ' Dynamic header based on category
    
    ws.Cells(1, timeCol + 3).EntireColumn.Insert
    ws.Cells(1, timeCol + 3).Value = dynamicHeader ' Set dynamic header
    
    ' Set headers for Additional Analysis columns
    ws.Cells(1, 58).Value = "Assets"
    ws.Cells(1, 59).Value = "Counts"

    ' Calculate % Loss Time using the time metric for further analysis
    lastRow = ws.Cells(ws.Rows.Count, timeCol).End(xlUp).Row
    ws.Range(ws.Cells(2, timeCol + 2), ws.Cells(lastRow, timeCol + 2)).Formula = _
        "=" & ws.Cells(2, timeCol).Address(False, False) & "/SUM(" & ws.Range(ws.Cells(2, timeCol), ws.Cells(lastRow, timeCol)).Address & ")"

    ' Calculate % of Key Metric and format as percentage
    ws.Range(ws.Cells(2, timeCol + 3), ws.Cells(lastRow, timeCol + 3)).Formula = _
        "=" & ws.Cells(2, timeCol + 2).Address(False, False) & "*" & keyMetricValue
    ws.Range(ws.Cells(2, timeCol + 3), ws.Cells(lastRow, timeCol + 3)).NumberFormat = "0.000%" ' Display as percentage with 3 decimal places

    ' Hide the original Time(hours) and % Loss Time columns after calculation
    ws.Columns(timeCol).Hidden = True
    ws.Columns(timeCol + 2).Hidden = True ' Hide the "% Loss Time" column

    ' Sort data by the % of Key Metric column in descending order for analysis
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Cells(2, timeCol + 3), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, areaCol), ws.Cells(lastRow, timeCol + 3))
        .Header = xlYes
        .Apply
    End With

    MsgBox "Required columns and new analysis columns are now set up!", vbInformation
End Sub
