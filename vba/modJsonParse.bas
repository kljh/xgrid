Attribute VB_Name = "modJsonParse"
Option Explicit

Function json_next_token(json As String, Optional pos As Long = 1)
    Dim p As Long, p0 As Long, p1 As Long, n As Long
    Dim c As String, token, token_is_String
    n = Len(json)
    p = pos
    c = Mid(json, p, 1)
    
    ' consume spaces
    Do While (c = " " Or c = vbNewLine Or c = Chr(10) Or c = Chr(13) Or c = vbTab) And p < n
        p = p + 1
        c = Mid(json, p, 1)
    Loop
    p0 = p
    
    token_is_String = c = """"
    
    If token_is_String Then
        ' consume string
        p = p + 1
        c = Mid(json, p, 1)
        Do While c <> """" And p < n
            If c = "\" Then
                p = p + 1
            End If
            p = p + 1
            c = Mid(json, p, 1)
        Loop
        Debug.Assert c = """"
        p = p + 1
    Else
        ' consume number/bool/id character (non spaces)
        p = p + 1
        c = Mid(json, p, 1)
        Do While InStrRev("_-0123456789.abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMONQRSTUVWXYZ", c) <> 0 And p < n
            p = p + 1
            c = Mid(json, p, 1)
        Loop
    End If
    p1 = p
    
    token = Mid(json, p0, p1 - p0)
    If token_is_String Then
        token = Replace(Replace(Replace(Mid(token, 2, Len(token) - 2), "\n", vbNewLine), "\t", vbTab), "\\", "\")
    ElseIf token = "true" Then
        token = True
    ElseIf token = "false" Then
        token = False
    ElseIf Len(token) > 1 Or Len(token) = 1 And InStrRev("0123456789", token) <> 0 Then
        token = CDbl(token)
    End If
    If token = "" Then
        Debug.Print "no more token"
    End If
    json_next_token = Array(token, p1)
End Function

Function json_parse(json As String)
    Dim pos As Long, tmp
    pos = 1
    tmp = json_parse_recurse(json, pos)
    
    On Error Resume Next
    json_parse = tmp(0)
    Set json_parse = tmp(0)
End Function

Function json_parse_recurse(json As String, pos As Long)
    Dim tmp, token As String, next_pos As Long
    tmp = json_next_token(json, pos)
    token = tmp(0)
    next_pos = tmp(1)
    
    If token = "{" Then
        json_parse_recurse = json_parse_object(json, pos)
    ElseIf token = "[" Then
        json_parse_recurse = json_parse_array(json, pos)
    Else
        json_parse_recurse = tmp
    End If
End Function

Function json_parse_object(json As String, Optional pos As Long)
    Dim obj
    Set obj = CreateObject("Scripting.Dictionary")
    
    Dim tmp, token As String, key As String, val, cln As String, sep As String
    tmp = json_next_token(json, pos)
    token = tmp(0)
    pos = tmp(1)
    Debug.Assert token = "{"
    
    ' early exit for empty object
    tmp = json_next_token(json, pos)
    val = tmp(0)
    If val = "}" Then
        json_parse_array = obj
        Exit Function
    End If

    Do While sep <> "}"
        tmp = json_parse_recurse(json, pos)
        key = tmp(0)
        pos = tmp(1)
    
        tmp = json_parse_recurse(json, pos)
        cln = tmp(0)
        pos = tmp(1)
        Debug.Assert sep = ":"
        
        tmp = json_parse_recurse(json, pos)
        val = tmp(0)
        pos = tmp(1)
        
        tmp = json_parse_recurse(json, pos)
        sep = tmp(0)
        pos = tmp(1)
        Debug.Assert sep = "," Or sep = "}"

        
        obj.Add key, val
    Loop
    
    json_parse_object = Array(obj, pos)
End Function

Function json_parse_array(json As String, Optional pos As Long)
    Dim tmp, token As String, val, sep As String
    tmp = json_next_token(json, pos)
    token = tmp(0)
    pos = tmp(1)
    Debug.Assert token = "["
    
    ' early exit for empty array
    tmp = json_next_token(json, pos)
    val = tmp(0)
    If val = "]" Then
        json_parse_array = Array(Array(), pos)
        Exit Function
    End If
    
    Dim arr
    Set arr = CreateObject("System.Collections.ArrayList")
    
    Do While sep <> "]"
        tmp = json_parse_recurse(json, pos)
        val = tmp(0)
        pos = tmp(1)
        
        tmp = json_parse_recurse(json, pos)
        sep = tmp(0)
        pos = tmp(1)
        Debug.Assert sep = "," Or sep = "]"
    
        arr.Add val
    Loop
    
    ReDim res(0 To arr.Count - 1)
    Dim i As Long
    For i = 0 To arr.Count - 1
        res(i) = arr(i)
    Next
    json_parse_array = Array(res, pos)
End Function

Sub json_next_token_test()
    Const json = "  { ""a"": [ 123, -4.56e-7], ""b"": true }"
    Dim tmp, token As String, pos As Long
    
    Debug.Print json
    tmp = json_next_token(json)
    token = tmp(0)
    pos = tmp(1)
    
    Do While token <> ""
        Debug.Print token
        tmp = json_next_token(json, pos)
        token = tmp(0)
        pos = tmp(1)
    Loop
    
End Sub

Sub json_parse_test()
    Const jsono = "  { ""a"": [ 123, -4.56e-7], ""b"": true }"
    Const jsona = "  [ 123, -4.56e-7, ""slaf"", true, [ 1, 2, 3] ] "
    
    Dim obj
    On Error Resume Next
    obj = json_parse(jsono)
    Set obj = json_parse(jsono)
    On Error GoTo 0
    
    Debug.Print vbs_val_to_json(obj, "  ")
End Sub
