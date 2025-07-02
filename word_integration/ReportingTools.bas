Attribute VB_Name = "ReportCompiler"
'// filepath: ReportingTools.bas
Option Explicit

'==================================================================================================
' Report Compiler VBA Module
'
' This module provides helper functions to insert placeholders and run the Python report compiler
' from within Microsoft Word.
'
' REQUIRES:
' To use the GetRelativePath function, you must enable the 'Microsoft Scripting Runtime'
' library. In the VBA Editor, go to Tools -> References, and check the box for
' "Microsoft Scripting Runtime".
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' PUBLIC PROCEDURES (Called by Ribbon Buttons)
'--------------------------------------------------------------------------------------------------

Public Sub InsertAppendixPlaceholder(control As IRibbonControl)
    ' Inserts a paragraph-based placeholder for merging a full PDF appendix.
    
    Dim pdfPath As String
    Dim relativePdfPath As String
    Dim placeholderText As String
    Dim cc As ContentControl
    
    ' Ensure the document is saved before creating a relative path.
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first to create a relative path for the appendix.", vbExclamation, "Save Document"
        Exit Sub
    End If
    
    ' Get the path to the PDF file from the user.
    pdfPath = GetPdfPath()
    If pdfPath = "" Then Exit Sub ' User cancelled
    
    ' Convert the absolute path to a relative path.
    relativePdfPath = GetRelativePath(pdfPath, ActiveDocument.Path)
    
    ' Construct the placeholder string.
    placeholderText = "[[INSERT: " & relativePdfPath & "]]"
    
    ' Insert the placeholder into a new paragraph and wrap it in a content control.
    With Selection
        .TypeParagraph
        Set cc = .Range.ContentControls.Add(wdContentControlText)
        With cc
            .Title = "Appendix Placeholder"
            .Tag = placeholderText
            .Range.Text = placeholderText
            .LockContents = True
        End With
        .MoveRight Unit:=wdCharacter, Count:=1
        .TypeParagraph
    End With
    
End Sub

Public Sub InsertOverlayPlaceholder(control As IRibbonControl)
    ' Inserts a table-based placeholder for overlaying a PDF page.
    
    Dim pdfPath As String
    Dim relativePdfPath As String
    Dim pageRange As String
    Dim cropText As String
    Dim placeholderText As String
    Dim tbl As Table
    Dim cc As ContentControl
    
    ' Ensure the document is saved before creating a relative path.
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first to create a relative path for the overlay.", vbExclamation, "Save Document"
        Exit Sub
    End If

    ' Get the path to the PDF file from the user.
    pdfPath = GetPdfPath()
    If pdfPath = "" Then Exit Sub ' User cancelled
    
    relativePdfPath = GetRelativePath(pdfPath, ActiveDocument.Path)
    
    ' Prompt for optional parameters.
    pageRange = InputBox("Enter an optional page range (e.g., 1-3,5). Leave blank for all pages.", "Overlay Page Selection")
    
    If MsgBox("Auto-crop the overlay to its content (removes whitespace)?", vbYesNo + vbQuestion, "Overlay Cropping") = vbNo Then
        cropText = ", crop=false"
    End If
    
    ' Construct the placeholder string.
    placeholderText = "[[OVERLAY: " & relativePdfPath
    If pageRange <> "" Then
        placeholderText = placeholderText & ", page=" & pageRange
    End If
    placeholderText = placeholderText & cropText & "]]"
    
    ' Insert a 1x1 table at the current selection.
    Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=1)
    
    ' Style the table to make it visible as a placeholder.
    With tbl.Borders
        .Enable = False

    End With
    
    ' Insert the placeholder text into the cell and wrap in a content control.
    Set cc = tbl.Cell(1, 1).Range.ContentControls.Add(wdContentControlText)
    With cc
        .Title = "Overlay Placeholder"
        .Tag = placeholderText
        .Range.Text = placeholderText
        .LockContents = True
    End With
    
End Sub

Public Sub RunReportCompiler(control As IRibbonControl)
    ' Saves the active document and executes the Python compiler script.
    
    ' --- CONFIGURATION: CHOOSE ONE ---
    ' For a deployed version (packaged .exe)
    ' Const COMPILER_PATH As String = "C:\Path\To\YourCompiler.exe"
    
    ' For development (running the raw Python script)
    Const COMPILER_PATH As String = "python C:\Users\p005452g\Source\report-compiler\main.py"
    ' ---------------------------------
    
    Dim doc As Document
    Dim inputPath As String
    Dim outputPath As String
    Dim cmdString As String
    
    Set doc = ActiveDocument
    
    ' Check if the document has been saved.
    If doc.Path = "" Then
        MsgBox "The document must be saved before the report can be compiled.", vbExclamation, "Save Document First"
        Exit Sub
    End If
    
    ' Save any pending changes.
    doc.Save
    
    ' Define input and output paths.
    inputPath = doc.FullName
    outputPath = Replace(doc.FullName, ".docx", ".pdf")
    
    ' Build the command string for the shell. Paths are wrapped in quotes.
    cmdString = COMPILER_PATH & " " & Chr(34) & inputPath & Chr(34) & " " & Chr(34) & outputPath & Chr(34)
    
    ' Execute the command. vbHide prevents the command window from flashing.
    On Error Resume Next
    Shell cmdString, vbHide
    If Err.Number <> 0 Then
        MsgBox "Failed to start the compiler. Please check the COMPILER_PATH constant in the 'RunReportCompiler' macro.", vbCritical, "Execution Error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Inform the user that the process has started.
    MsgBox "The report compiler has been started. This may take a moment. You will see the final PDF when it is complete.", vbInformation, "Compiler Started"
    
End Sub


'--------------------------------------------------------------------------------------------------
' PRIVATE HELPER FUNCTIONS
'--------------------------------------------------------------------------------------------------

Private Function GetPdfPath() As String
    ' Opens a file picker dialog for the user to select a PDF file.
    ' Returns the full path of the selected file, or an empty string if cancelled.
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a PDF File"
        .AllowMultiSelect = False
        
        ' Clear existing filters and add one for PDF files.
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        .Filters.Add "Word Files", "*.docx"
        
        ' Show the dialog and check if a file was selected.
        If .Show = True Then
            GetPdfPath = .SelectedItems(1)
        Else
            GetPdfPath = "" ' User cancelled
        End If
    End With
    
End Function

Private Function GetRelativePath(ByVal targetPath As String, ByVal basePath As String) As String
    ' Calculates the relative path from a base folder to a target file.
    ' Requires a reference to "Microsoft Scripting Runtime".
    
    Dim fso As Object
    Dim relativePath As String
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    relativePath = fso.GetFile(targetPath).Path
    
    ' Use the built-in (but poorly documented) RelativePath property if available
    ' This is a fallback for older systems; modern FSO should handle it.
    If fso.FolderExists(basePath) Then
        Dim tempFile As Object
        Set tempFile = fso.GetFile(targetPath)
        ' A trick to get relative path
        GetRelativePath = Mid(tempFile.Path, Len(fso.GetAbsolutePathName(basePath)) + 2)
        ' A more robust method using built-in functionality if available
        GetRelativePath = fso.GetFolder(basePath).ParentFolder.CreateTextFile("dummy.txt", True).ParentFolder.GetRelativePath(targetPath)
        fso.DeleteFile fso.GetFolder(basePath).ParentFolder.Path & "\dummy.txt"
    Else
        GetRelativePath = targetPath ' Fallback to absolute path
    End If
    
    Exit Function

ErrorHandler:
    ' If any error occurs (e.g., FSO not available), fall back to the absolute path.
    GetRelativePath = targetPath
End Function
