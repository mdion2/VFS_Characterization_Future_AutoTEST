'VBA Importer
'Import VBA code from an export folder to Excel. To be used with export_vba.vbs.
'Example: import_vba.vbs script.xlsm
'mwaterbu@ford.com
'Copyright: Ford Motor Company Limited

Option Explicit

Main

Sub Main
    Dim xl
    Dim fs
    Dim fld
    Dim f
    Dim fn
    Dim WBook
    Dim ImportFolder
    Dim arg
    Dim refFile
    
    Set xl = CreateObject("Excel.Application")
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    xl.Visible = True 'Not necessary, but allows user to see some progress
    xl.AutomationSecurity = 3 'Do not allow macros to run when the file is opened
    xl.DisplayAlerts = False 'Do not display errors or other messages that could halt execution
    
    'Loop over each file passed in
    For Each arg in Wscript.Arguments
        Set WBook = xl.Workbooks.Open(Trim(arg),,,,,,True)
        ImportFolder = WBook.Path & "\" & fs.GetBaseName(WBook.Name) & "\Exported_VBA"
        If Not fs.FolderExists(ImportFolder) Then
            MsgBox "VBA folder not found for " & WBook.Name & ". Unable to import code."
        Else
            Set fld = fs.GetFolder(ImportFolder)
            'Import each file (module/form/class) from the associated folder
            For Each f In fld.Files
                fn = fs.GetBaseName(f.Name)
                On Error Resume Next
                Select Case f.Type
                    Case "CLS File"
                        'For class files, try to remove, then import, if not able to remove (Sheet, ThisWorkbook), clear code, then add text
                        If moduleExists(WBook.VBProject, fn) Then WBook.VBProject.VBComponents.Remove(WBook.VBProject.VBComponents(fn))
                        Err.Clear
                        If moduleExists(WBook.VBProject, fn) Then
                            With WBook.VBProject.VBComponents(fn).CodeModule
                                .DeleteLines 1, .CountOfLines
                                .AddFromFile(f.Path)
                                .DeleteLines 1, 4
                            End With
                        Else
                            WBook.VBProject.VBComponents.Import(f.Path)
                        End If
                    Case "BAS File", "FRM File"
                        'For modules and forms, remove the module if it exists, then import
                        If moduleExists(WBook.VBProject, fn) Then WBook.VBProject.VBComponents.Remove(WBook.VBProject.VBComponents(fn))
                        Err.Clear
                        If moduleExists(WBook.VBProject, fn) Then
                            MsgBox "Failed to import module " & f.Name & "."
                        Else
                            WBook.VBProject.VBComponents.Import(f.Path)
                            If f.Type = "FRM File" Then WBook.VBProject.VBComponents(fn).CodeModule.DeleteLines 1, 1
                        End If
                    Case Else 'FRX, TXT, etc
                        'Ignore
                End Select
                If Err.Number <> 0 Then
                    MsgBox "Failed to import module " & f.Name & "."
                End If
                On Error Goto 0
            Next
            
            If Not fs.FileExists(ImportFolder & "\References.txt") Then
                MsgBox "References file not found for " & WBook.Name & ". Unable to import references."
            Else
                Set refFile = fs.OpenTextFile(ImportFolder & "\References.txt")
                'Import list of references
                Do Until refFile.AtEndOfStream
                    fn = refFile.ReadLine
                    If referenceExists(WBook.VBProject, fn) Then
                        'Ignore
                    Else
                        On Error Resume Next
						WBook.VBProject.References.AddFromFile fn
						If Err.Number <> 0 Then
							MsgBox "Failed to import reference " & fn & "."
						End If
						On Error Goto 0
                    End If
                Loop
                refFile.Close
            End If
        End If
        WBook.Close True 'Save file
    Next
    
    xl.Quit
End Sub

'Check if module already exists
Function moduleExists(VBProj, checkName)
    Dim VBComp
    moduleExists = False
    For Each VBComp in VBProj.VBComponents
        If VBComp.Name = checkName Then
            moduleExists = True
            Exit For
        End If
    Next
End Function

'Check if reference already exists
Function referenceExists(VBProj, checkPath)
    Dim ref
    referenceExists = False
    For Each ref in VBProj.References
		If LCase(Right(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\"))) = LCase(Right(checkPath, Len(checkPath) - InStrRev(checkPath, "\"))) Then
            referenceExists = True
            Exit For
        End If
    Next
End Function
