Public Function HttpGet(url As String, Optional params As String) As HTMLDocument
    Set HttpGet = ConvertToHTMLDoc(HttpGet2(url, params))
End Function

Public Function HttpPost(url As String, params As String) As HTMLDocument
    Set HttpPost = ConvertToHTMLDoc(HttpPost2(url, params))
End Function

Public Function HttpGet2(url As String, Optional params As String) As String
    Dim winHttpReq As New WinHttpRequest
    If Len(params) > 0 Then
        url = url + "?" + params
    End If
    winHttpReq.Open "GET", url, False
    winHttpReq.Send
    response = winHttpReq.responseText
    Set winHttpReq = Nothing
    HttpGet2 = response
End Function

Public Function HttpPost2(url As String, params As String) As String
    Dim winHttpReq As New WinHttpRequest
    winHttpReq.Open "POST", url, False
    winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    winHttpReq.Send params
    response = winHttpReq.responseText
    Set winHttpReq = Nothing
    HttpPost2 = response
End Function

Public Function ConvertToHTMLDoc(ByVal html As String) As HTMLDocument
    Dim doc As New HTMLDocument
    doc.body.innerHTML = html
    Set ConvertToHTMLDoc = doc
End Function
