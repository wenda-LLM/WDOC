Attribute VB_Name = "NewMacros"
Sub wenda()
 Dim selectedText As String
 Dim apiKey As String
 Dim response As Object, re As String
 Dim midString As String
 Dim ans As String
 If Selection.Type = wdSelectionNormal Then
     selectedText = Selection.Text
     selectedText = Replace(selectedText, ChrW$(13), "")
     URL = "http://127.0.0.1:17860/api/chat"
     Set response = CreateObject("MSXML2.XMLHTTP")
     response.Open "POST", URL, False
     response.setRequestHeader "Content-Type", "application/json"
     response.Send "{""prompt"":""" & selectedText & """,""temperature"":1.3,""top_p"":0.5,""max_length"":2048,""history"":[]}"
     ans = response.responseText
     Selection.EndKey
     Selection.Text = vbNewLine & ans
 Else
    Exit Sub
 End If
End Sub

Sub wenda_zq()
 Dim selectedText As String
 Dim apiKey As String
 Dim response As Object, re As String
 Dim midString As String
 Dim ans As String
 If Selection.Type = wdSelectionNormal Then
     selectedText = Selection.Text
     selectedText = Replace(selectedText, ChrW$(13), "")
     URL = "http://127.0.0.1:17861/"
     Set response = CreateObject("MSXML2.XMLHTTP")
     response.Open "POST", URL, False
     response.Send selectedText
     ans = response.responseText
     Selection.EndKey
     Selection.Text = vbNewLine & ans
 Else
    Exit Sub
 End If
End Sub
Sub wenda_find()
 Dim selectedText As String
 Dim apiKey As String
 Dim response As Object, re As String
 Dim midString As String
 Dim ans As String
 If Selection.Type = wdSelectionNormal Then
     selectedText = Selection.Text
     selectedText = Replace(selectedText, ChrW$(13), "")
     URL = "http://127.0.0.1:17861/find"
     Set response = CreateObject("MSXML2.XMLHTTP")
     response.Open "POST", URL, False
     response.Send selectedText
     ans = response.responseText
     Selection.EndKey
     Selection.Text = vbNewLine & ans
 Else
    Exit Sub
 End If
End Sub
Sub wenda_raw()
 Dim selectedText As String
 Dim apiKey As String
 Dim response As Object, re As String
 Dim midString As String
 Dim ans As String
 If Selection.Type = wdSelectionNormal Then
     selectedText = Selection.Text
     selectedText = Replace(selectedText, ChrW$(13), "")
     URL = "http://127.0.0.1:17861/raw"
     Set response = CreateObject("MSXML2.XMLHTTP")
     response.Open "POST", URL, False
     response.Send selectedText
     ans = response.responseText
     Selection.EndKey
     Selection.Text = ans
 Else
    Exit Sub
 End If
End Sub
