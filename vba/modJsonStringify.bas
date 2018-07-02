Attribute VB_Name = "modJsonStringify"
Option Explicit

' To use Dictionary object type:
' Add a reference to Microsoft Scripting Runtime.

Const indentchar = "    "

Function vbs_val_to_json(v, Optional indent As String = "") As String
    Dim json As String, t As String, a As Boolean, d As Dictionary
    
    t = TypeName(v)
    a = Right(t, 2) = "()"
        
    If a Then
        json = vbs_array_to_json(v, indent)
    ElseIf t = "ArrayList" Then
        json = vbs_arraylist_to_json(v, indent)
    ElseIf t = "Dictionary" Then
        Set d = v
        json = vbs_dict_to_json(d, indent)
    ElseIf t = "Object" Then
        'Set d = New Dictionary
        If v Is d Then
            Set d = v
            json = vbs_dict_to_json(d, indent)
        Else
            json = "unknow " & t
        End If
    Else
        json = json & indent
        
        If t = "String" Then
            json = vbs_json_string("" & v)
        ElseIf t = "Double" Or t = "Integer" Or t = "Long" Then
            json = "" & v
        ElseIf t = "Boolean" Then
            json = IIf(v, "true", "false") & ""
        ElseIf t = "Date" Then
            json = """" & Format(v, "yyyy-mm-dd hh:mm:ss") & """"
        ElseIf t = "Null" Or t = "Nothing" Then
            json = "null"
        ElseIf IsEmpty(v) Then
            json = "null"
        Else
            json = "unknow type " & t
        End If
    End If
    
    vbs_val_to_json = json
End Function

Function vbs_array_to_json(v, Optional indent As String = "") As String
    Dim t As String, a As Boolean
    t = TypeName(v)
    a = Right(t, 2) = "()"
    Debug.Assert a
    
    ' check it is a 1D array
    On Error Resume Next
    Dim upper_bound_2
    upper_bound_2 = UBound(v, 2)
    On Error GoTo 0
    If Not IsEmpty(upper_bound_2) Then
        vbs_array_to_json = vbs_array_2d_to_json(v, indent)
        Exit Function
    End If
    
    
    Dim json As String, i As Long, n As Long
    n = UBound(v) - LBound(v) + 1
    
    For i = LBound(v) To UBound(v)
        json = json & indent & indentchar & vbs_val_to_json(v(i), indent & indentchar)
        json = json & IIf(i = UBound(v), "", ",") & vbNewLine
    Next i
    
    json = IIf(n = 0, "[]", "[" & vbNewLine & json & indent & "]")
    vbs_array_to_json = json
End Function

Function vbs_array_2d_to_json(v, Optional indent As String = "") As String
    Dim t As String, a As Boolean
    t = TypeName(v)
    a = Right(t, 2) = "()"
    Debug.Assert a
    
    ' check it is a 1D array
    On Error Resume Next
    Dim upper_bound_2, upper_bound_3
    upper_bound_2 = UBound(v, 2)
    upper_bound_3 = UBound(v, 3)
    On Error GoTo 0
    Debug.Assert Not IsEmpty(upper_bound_2)
    Debug.Assert IsEmpty(upper_bound_3)
    
    Dim json As String, json_row As String, i As Long, j As Long, n As Long, m As Long
    n = UBound(v) - LBound(v) + 1
    m = UBound(v, 2) - LBound(v, 2) + 1
    
    For i = LBound(v) To UBound(v)
        json_row = ""
        For j = LBound(v, 2) To UBound(v, 2)
            json_row = json_row & indent & indentchar & indentchar & vbs_val_to_json(v(i, j), indent & indentchar)
            json_row = json_row & IIf(j = UBound(v, 2), "", ",") & vbNewLine
        Next j
        json_row = IIf(m = 0, "[]", "[" & vbNewLine & json_row & indent & indentchar & "]")
        
        json = json & indent & indentchar & json_row
        json = json & IIf(i = UBound(v), "", ",") & vbNewLine
    Next i
    
    json = IIf(n = 0, "[]", "[" & vbNewLine & json & indent & "]")
    vbs_array_2d_to_json = json
End Function

Function vbs_arraylist_to_json(v, Optional indent As String = "") As String
    Dim t As String
    t = TypeName(v)
    Debug.Assert t = "ArrayList"
    
    Dim json As String, i As Long, n As Long
    n = v.Count
    
    For i = 0 To v.Count - 1
        json = json & indent & indentchar & vbs_val_to_json(v(i), indent & indentchar)
        json = json & IIf(i = n - 1, "", ",") & vbNewLine
    Next i
    
    json = IIf(n = 0, "[]", "[" & vbNewLine & json & indent & "]")
    vbs_arraylist_to_json = json
End Function

Function vbs_dict_to_json(d As Dictionary, Optional indent As String = "") As String
    Dim t As String, a As Boolean
    t = TypeName(d)
    Debug.Assert t = "Dictionary"
    
    Dim json As String, i As Long, n As Long, k, v, tmp
    n = d.Count
    json = "{" & IIf(n = 0, "", vbNewLine)
    
    Dim keys
    keys = d.keys
    'keys = vbs_sort(keys)
    
    For Each k In keys
        i = i + 1
        
        On Error Resume Next
        v = d(k)
        Set v = d(k)
        On Error GoTo 0
        
        json = json & indent & indentchar & vbs_json_string("" & k) & ": " & vbs_val_to_json(v, indent & indentchar)
        json = json & IIf(i = n, "", ",") & vbNewLine
    Next
    
    json = json & IIf(n = 0, "", indent) & "}"
    vbs_dict_to_json = json
End Function

Function vbs_json_string(s As String)
    vbs_json_string = """" & Replace(Replace(s, "\", "\\"), """", "\""") & """"
End Function

Function vbs_sort(arr)
    ' Copy to an .Net ArrayList
    Dim tmp, i, v
    Set tmp = CreateObject("System.Collections.ArrayList")
    For Each v In arr
        tmp.Add "" & v
    Next

    ' Sort
    Call tmp.Sort
    'Debug.Print "first element is now in position " & tmp.IndexOf(arr(LBound(arr)))
    
    ' Copy back to a VBA array
    If tmp.Count = 0 Then
        vbs_sort = Array()
    Else
        ReDim res(0 To tmp.Count - 1)
        For i = 0 To tmp.Count - 1
            res(i) = tmp(i)
        Next
        vbs_sort = res
    End If
End Function
