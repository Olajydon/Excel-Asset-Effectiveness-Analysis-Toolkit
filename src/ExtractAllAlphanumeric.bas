Function ExtractAllAlphanumeric(str As String) As String
    Dim regex As Object, matches As Object, match As Object
    Dim result As String
    
    ' Create a new RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    ' Pattern to match both formats: letters and numbers mixed together
    regex.Pattern = "\b[A-Za-z0-9]+(?:[A-Za-z0-9]+)*\b"
    
    ' Get all matches
    Set matches = regex.Execute(str)
    
    ' Loop through matches and concatenate them, ensuring valid codes
    For Each match In matches
        ' Check if the match contains at least one letter and one number
        If match.Value Like "*[0-9]*" And match.Value Like "*[A-Za-z]*" Then
            If result = "" Then
                result = match.Value
            Else
                result = result & ", " & match.Value
            End If
        End If
    Next match
    
    ExtractAllAlphanumeric = result
End Function
