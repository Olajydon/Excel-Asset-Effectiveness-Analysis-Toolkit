Sub RunAnalysis()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Step 1: Prepare the Analysis Sheet
    ' This step organizes the sheet by setting up required columns and handling duplicates.
    Call PrepareAnalysisSheet
    
    ' Step 2: Extract Alphanumeric Identifiers
    ' Define the range for the column that contains the main text data (adjust as needed).
    Dim lastRow As Long, cell As Range
    lastRow = ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row ' Reference column containing source text, e.g., "Remarks"
    
    ' Clear the "Key" column (AZ) and populate it with extracted alphanumerics
    ws.Range("AZ2:AZ" & lastRow).ClearContents
    For Each cell In ws.Range("AZ2:AZ" & lastRow) ' Define "Key" column starting at AZ2
        cell.Value = ExtractAllAlphanumeric(ws.Cells(cell.Row, "Y").Value) ' Extract alphanumeric data from source column
    Next cell
    
    ' Step 3: Count Unique Alphanumeric Entries and Output Top Counts
    ' This step tallies and displays the most frequent alphanumeric entries.
    CountAlphanumerics
    
    ' Notify user upon completion
    MsgBox "Asset Effectiveness Analysis Completed!", vbInformation
End Sub
