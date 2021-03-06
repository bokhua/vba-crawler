' To use ConvertToHTMLDoc, you need add Microsoft HTML Object Library at Tools > References
' To use ConvertToXMLDOM, you need add Microsoft XML at Tools > References
Public Function HttpGet(ByVal url As String, Optional ByVal params As String) As String
    Dim winHttpReq As Object
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Len(params) > 0 Then
        url = url + "?" + params
    End If
    winHttpReq.Open "GET", url, False
    winHttpReq.send
    response = winHttpReq.responseText
    Set winHttpReq = Nothing
    HttpGet = response
End Function

' params: "var1=a&var2=b"
Public Function HttpPost(ByVal url As String, ByVal params As String) As String
    Dim winHttpReq As Object
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpReq.Open "POST", url, False
    winHttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    winHttpReq.send params
    response = winHttpReq.responseText
    Set winHttpReq = Nothing
    HttpPost = response
End Function

Public Function ConvertToHTMLDoc(ByVal html As String) As HTMLDocument
    Dim doc As New HTMLDocument
    doc.body.innerHTML = html
    Set ConvertToHTMLDoc = doc
End Function

Public Function ConvertToXMLDOM(ByVal xmlStr As String) As DOMDocument
    Dim xml As New DOMDocument
    xml.loadXML xmlStr
    Set ConvertToXMLDOM = xml
End Function

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

Public Function DecodeHTML(ByVal html As String) As String

    html = Replace(html, "&amp;", "&")
    html = Replace(html, "&nbsp;", " ")
    html = Replace(html, "&quot;", Chr(34))
    html = Replace(html, "&apos;", "'")
    html = Replace(html, "&lt;", "<")
    html = Replace(html, "&gt;", ">")
    
    DecodeHTML = html
End Function

Public Function StrRemoveBoundaries(ByVal str As String, ByVal leftBound As String, ByVal rightBound As String) As String
    len1 = Len(leftBound)
    len2 = Len(rightBound)
    len3 = Len(str)
    StrRemoveBoundaries = Mid(str, len1 + 1, len3 - len1 - len2)
End Function
