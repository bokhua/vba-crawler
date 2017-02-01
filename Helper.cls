Public Function MatchString(ByVal str As String, ByVal leftBound As String, ByVal rightBound As String, Optional ByVal includeBoundaries As Boolean) As Collection
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    mustEscapedCharacters = Split("\ . $ ^ { [ ( | ) * + ?", " ")
    
    leftBoundX = leftBound
    rightBoundX = rightBound
    
    For Each char In mustEscapedCharacters
        leftBoundX = Replace(leftBoundX, char, "\" + char)
        rightBoundX = Replace(rightBoundX, char, "\" + char)
    Next
    
    regex.Pattern = "(" + leftBoundX + ").*?(" + rightBoundX + ")"
   
    Set regexMatch = regex.Execute(str)
    If regexMatch.Count > 0 Then
        Dim result As New Collection
        For Each mat In regexMatch
            If includeBoundaries = True Then
                result.Add mat
            Else
                result.Add StrRemoveBoundaries(mat, leftBound, rightBound)
            End If
        Next
        Set MatchString = result
    End If
End Function

Public Function MatchStringFirst(ByVal str As String, ByVal leftBound As String, ByVal rightBound As String, Optional ByVal includeBoundaries As Boolean) As String
    Dim match As New Collection
    Set match = MatchString(str, leftBound, rightBound, includeBoundaries)
    If match.Count > 0 Then
        MatchStringFirst = match.Item(1)
    Else
        MatchStringFirst = ""
    End If
End Function

Public Function MatchStringLast(ByVal str As String, ByVal leftBound As String, ByVal rightBound As String, Optional ByVal includeBoundaries As Boolean) As String
    Dim match As New Collection
    Set match = MatchString(str, leftBound, rightBound, includeBoundaries)
    If match.Count > 0 Then
        MatchStringLast = match.Item(match.Count)
    Else
        MatchStringLast = ""
    End If
End Function

Public Function StrRemoveBoundaries(ByVal str As String, ByVal leftBound As String, ByVal rightBound As String) As String
    len1 = Len(leftBound)
    len2 = Len(rightBound)
    len3 = Len(str)
    StrRemoveBoundaries = Mid(str, len1 + 1, len3 - len1 - len2)
End Function
