VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Load VBA modules from disk
    Call LoadVBAModules(ThisWorkbook.path, ThisWorkbook)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not ThisWorkbook.Saved Then
        If MsgBox(ThisWorkbook.Name & " not saved. Save ?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        ' Write VBA modules to disk, remove them, and save Excel file
        UnloadVBAModules ThisWorkbook.path, ThisWorkbook
        ThisWorkbook.Save
    End If
End Sub

Sub LoadVBAModulesTest()
    LoadVBAModules ThisWorkbook.path, ThisWorkbook
End Sub
Sub LoadVBAModules(path As String, wbk As Workbook)
    Dim VBProj As Object ' Object or VBIDE.VBProject
    On Error Resume Next
    Set VBProj = wbk.VBProject
    On Error GoTo 0
    
    If VBProj Is Nothing Then
        MsgBox "LoadVBAModules can't run: XLA not trusted." & vbNewLine & _
            "Go to Excel options, Macro security, Trust access to VB Project."
        Exit Sub
    End If
    
    
    Dim fso, folder, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
            VBProj.VBComponents.Import file.path
        End If
    Next
End Sub

Sub UnloadVBAModulesTest()
    UnloadVBAModules ThisWorkbook.path, ThisWorkbook
End Sub
Sub UnloadVBAModules(path As String, wbk As Workbook)
    Dim VBProj As VBIDE.VBProject
    On Error Resume Next
    Set VBProj = wbk.VBProject
    On Error GoTo 0
    
    If VBProj Is Nothing Then
        Exit Sub
    End If
    
    Dim component As VBIDE.VBComponent
    For Each component In VBProj.VBComponents
        ' Trim empty lines at the beginning and at the end of the code module
        If component.Type = vbext_ct_StdModule Then
            TrimEmptyLines component
            Call component.Export(path & "\" & component.Name & ".bas")
            VBProj.VBComponents.Remove component
        End If
    Next component
End Sub

' Trim empty lines at the beginning and at the end of the code module
Private Sub TrimEmptyLines(component As VBIDE.VBComponent)
    Dim bContinueTrim As Boolean
    
    ' Remove at the beginning
    bContinueTrim = True
    While bContinueTrim
        If component.CodeModule.CountOfLines = 0 Then
            bContinueTrim = False
        Else
            If component.CodeModule.Lines(1, 1) = "" Then
                component.CodeModule.DeleteLines 1
            Else
                bContinueTrim = False
            End If
        End If
    Wend
    
    ' Remove at the end
    bContinueTrim = True
    While bContinueTrim
        If component.CodeModule.CountOfLines = 0 Then
            bContinueTrim = False
        Else
            If component.CodeModule.Lines(component.CodeModule.CountOfLines, 1) = "" Then
                component.CodeModule.DeleteLines component.CodeModule.CountOfLines
            Else
                bContinueTrim = False
            End If
        End If
    Wend
End Sub


