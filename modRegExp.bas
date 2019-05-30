Function ExtractTerms(strPhrase As String) As Variant
Dim RegExp As Object
Dim strPattern As String
Dim RegMatches As Object
Dim RegMatch As Object
Dim arrMatches()
Dim cnt As Long

    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.Pattern = "{(.*?)}"

    Set RegMatches = RegExp.Execute(strPhrase)
            
    If RegMatches.Count > 0 Then
        ReDim arrMatches(1 To RegMatches.Count)

        For Each RegMatch In RegMatches
            cnt = cnt + 1
            arrMatches(cnt) = RegMatch
        Next RegMatch

        ExtractTerms = arrMatches
    End If
    
End Function
