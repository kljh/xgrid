Attribute VB_Name = "modWalkDir"
Option Explicit

Sub WalkDirTest()
    
    Dim xlFilesRoot As String, xlSaveRoot As String, fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    xlFilesRoot = fso.GetAbsolutePathName(ThisWorkbook.path & "\..\..\..\vision\bin\xls") ' Vision XLL folder
    xlSaveRoot = fso.GetAbsolutePathName(ThisWorkbook.path & "\..\..\..\vision\bin\xjs")
    If Not fso.FolderExists(xlFilesRoot) Then
        xlFilesRoot = "C:\temp\"
        xlSaveRoot = "C:\temp\xjs"
    End If
    If Not fso.FolderExists(xlSaveRoot) Then
        MkDir xlSaveRoot
    End If
    Debug.Print "xlFilesRoot " & xlFilesRoot
    Debug.Print "xlSaveRoot " & xlSaveRoot
    
    Dim xlFiles, xlFileIn As String, xlFileOut As String, wbk As Workbook
    Set xlFiles = WalkDir(xlFilesRoot)
    Debug.Print "#xlFiles " & xlFiles.Count
            
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    'Application.AutomationSecurity = msoAutomationSecurityByUI  ' Default behavior: ask me whether I want to enable or disable macros
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    'Application.AutomationSecurity = msoAutomationSecurityLow

    Dim i As Long, n As Long, fileLenKB As Double
    n = xlFiles.Count
    For i = 0 To n - 1
        xlFileIn = xlFiles(i)
        fileLenKB = Round(FileLen(xlFileIn) / 1024)
        
        Debug.Print "< " & (i + 1) & "/" & n & " " & vbTab & fileLenKB & " KB " & vbTab & Replace(xlFileIn, xlFilesRoot, "")
        
        If fileLenKB > 4000 Then
            Debug.Print "> TOO BIG"
        Else
            Set wbk = Workbooks.Open(xlFileIn, ReadOnly = True)
            Application.Calculation = xlManual
            Application.CalculateBeforeSave = False
        
            xlFileOut = wbk.FullName & ".json"
            xlFileOut = xlSaveRoot & "\" & wbk.Name & ".json"  '  !!
        
            Call export_workbook_to_file(wbk, xlFileOut)
            wbk.Close (False)
            
            Debug.Print "> " & xlFileOut
        End If
    Next i
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AutomationSecurity = msoAutomationSecurityByUI
End Sub

Function WalkDir(path As String)
    Dim list
    Set list = CreateObject("System.Collections.ArrayList")
    
    Call WalkDirIter(list, path)
    
    Set WalkDir = list
End Function


Function WalkDirIter(ByRef list, path As String)
    Dim fNext As String, fullPath As String
    
    Dim ext As String, keep As Boolean
    fNext = Dir(path & "\*.*")
    Do While fNext <> ""
        fullPath = path & "\" & fNext
        ext = LCase(Mid(fNext, InStrRev(fNext, ".")))
        keep = (ext = ".xls") Or (ext = ".xlw") Or (ext = ".xlsx") Or (ext = ".xlsm")
        
        If keep Then
            list.Add fullPath
        End If
        fNext = Dir()
    Loop
    
    Dim subfolders
    Set subfolders = CreateObject("System.Collections.ArrayList")
    fNext = Dir(path & "\*", vbDirectory)
    Do While fNext <> ""
        fullPath = path & "\" & fNext
        If Left(fNext, 1) <> "." And GetAttr(fullPath) = vbDirectory Then
            subfolders.Add (fullPath)
        End If
        
        fNext = Dir()
    Loop
    
    Dim i As Long, n As Long, subfolder As String
    n = subfolders.Count
    For i = 0 To n - 1
        subfolder = subfolders(i)
        Call WalkDirIter(list, subfolder)
    Next i
    
End Function
