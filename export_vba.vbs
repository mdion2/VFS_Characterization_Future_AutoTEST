'VBA Exporter
'Export VBA code from Excel to individual files for comparison, version control, or backups.
'Example: export_vba.vbs script.xlsm
'mwaterbu@ford.com
'Copyright: Ford Motor Company Limited

Option Explicit

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

Main

Sub Main
    Dim xl
    Dim fs
    Dim WBook
    Dim VBComp
    Dim Sfx
    Dim ExportFolder
    Dim arg
    Dim refFile
    Dim ref
    
    Set xl = CreateObject("Excel.Application")
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    xl.Visible = True 'Not necessary, but allows user to see some progress
    xl.AutomationSecurity = 3 'Do not allow macros to run when the file is opened
    xl.DisplayAlerts = False 'Do not display errors or other messages that could halt execution
    
    'Loop over each file passed in
    For Each arg in Wscript.Arguments
        Set WBook = xl.Workbooks.Open(Trim(arg))
        ExportFolder = WBook.Path & "\" & fs.GetBaseName(WBook.Name)
        'Create base folder if needed
        If Not fs.FolderExists(ExportFolder) Then
            fs.CreateFolder(ExportFolder)
        End If
        ExportFolder = ExportFolder & "\Exported_VBA"
        'Remove old exported files, if they exist
        If fs.FolderExists(ExportFolder) Then
            fs.DeleteFolder ExportFolder, True
        End If
        fs.CreateFolder(ExportFolder)
        'Export each module/form/class
        For Each VBComp In WBook.VBProject.VBComponents
            Select Case VBComp.Type
                Case vbext_ct_ClassModule, vbext_ct_Document
                    Sfx = ".cls"
                Case vbext_ct_MSForm
                    Sfx = ".frm"
                Case vbext_ct_StdModule
                    Sfx = ".bas"
                Case Else
                    Sfx = ""
            End Select
            If Sfx <> "" Then
                On Error Resume Next
                Err.Clear
                VBComp.Export ExportFolder & "\" & VBComp.Name & Sfx
                If Err.Number <> 0 Then
                    MsgBox "Failed to export " & ExportFolder & "\" & VBComp.Name & Sfx
                End If
                On Error Goto 0
            End If
        Next
        Set refFile = fs.CreateTextFile(ExportFolder & "\References.txt", True)
        'Export list of references being used
        For Each ref In WBook.VBProject.References
            refFile.WriteLine ref.FullPath
        Next
        refFile.Close
        WBook.Close False 'Do not save
    Next
    
    xl.Quit
End Sub
