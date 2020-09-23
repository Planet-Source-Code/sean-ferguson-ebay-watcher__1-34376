Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOW = 5

Public showAlerts As Boolean

Public Function openURL(strURL As String)
    ShellExecute frmWatch.hwnd, "open", strURL, vbNullString, vbNullString, SW_SHOW
End Function

Public Function returnInfoArray(strHTMLData) As Variant
    On Error Resume Next
    Dim iArray(1 To 3) As String
    Dim sString As String
    Dim dPos As Long
    Dim dPos2 As Long
    sString = "<title>" & vbCrLf & "ebay item"
    dPos = InStr(1, LCase(strHTMLData), sString) + Len(sString)
    If dPos > Len(sString) Then
        sString = " - "
        dPos = InStr(dPos, LCase(strHTMLData), sString) + Len(sString)
        If dPos > Len(sString) Then
            sString = "</title>"
            dPos2 = InStr(dPos, LCase(strHTMLData), sString)
            iArray(1) = stripLineBreak(Mid(strHTMLData, dPos, dPos2 - dPos))
        Else
            iArray(1) = "(Info not found)"
        End If
    Else
        iArray(1) = "(Info not found)"
    End If
    sString = "# of bids"
    dPos = InStr(1, LCase(strHTMLData), sString) + Len(sString)
    If dPos > Len(sString) Then
        sString = "<td width=""45%""><b>"
        dPos = InStr(dPos, LCase(strHTMLData), sString) + Len(sString)
        If dPos > Len(sString) Then
            sString = "</b><font size=""2"">"
            dPos2 = InStr(dPos, LCase(strHTMLData), sString)
            iArray(2) = stripLineBreak(Mid(strHTMLData, dPos, dPos2 - dPos))
        Else
            iArray(2) = ""
        End If
    Else
        iArray(2) = ""
    End If
    sString = "font size=""2"">" & vbCrLf & "currently" & vbCrLf & "</font>"
    dPos = InStr(1, LCase(strHTMLData), sString) + Len(sString)
    If dPos > Len(sString) Then
        sString = "<td width=""31%""><b>"
        dPos = InStr(dPos, LCase(strHTMLData), sString) + Len(sString)
        If dPos > Len(sString) Then
            sString = "</b>"
            dPos2 = InStr(dPos, LCase(strHTMLData), sString)
            iArray(3) = stripLineBreak(Mid(strHTMLData, dPos, dPos2 - dPos))
        Else
            iArray(3) = ""
        End If
    Else
        iArray(3) = ""
    End If
    returnInfoArray = iArray
End Function

Public Function getUBound(ByRef arrArray) As Long
    On Error GoTo handleError
    getUBound = UBound(arrArray)
    Exit Function
    
handleError:
    getUBound = -1
End Function

Public Function stripLineBreak(strText) As String
    On Error Resume Next
    strText = CStr(Trim(strText))
    strText = Replace(strText, vbCrLf, " ")
    Do Until InStr(1, strText, "  ") = 0
        strText = Replace(strText, "  ", " ")
    Loop
    stripLineBreak = Trim(strText)
End Function
