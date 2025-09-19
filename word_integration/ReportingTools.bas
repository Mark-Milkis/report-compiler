'// filepath: ReportingTools.bas
Option Explicit

'==================================================================================================
' Report Compiler VBA Module
'
' This module provides helper functions to insert placeholders and run the Python report compiler
' from within Microsoft Word.
'
' REQUIRES:
' This module depends on the LibFileTools module for OneDrive/SharePoint path handling.
' Make sure LibFileTools.bas is imported into your VBA project.
'
' LibFileTools provides robust file path operations that work with:
' - Local file paths
' - OneDrive synchronized folders  
' - SharePoint synchronized folders
' - UNC network paths
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' PUBLIC PROCEDURES (Called by Ribbon Buttons)
'--------------------------------------------------------------------------------------------------

Public Sub InsertAppendixPlaceholder(control As IRibbonControl)
    ' Inserts a paragraph-based placeholder for merging a full PDF appendix.
    
    Dim pdfPath As String
    Dim localDocPath As String
    Dim localPdfPath As String
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
    
    ' Convert OneDrive/SharePoint paths to local paths for both document and selected PDF
    localDocPath = GetLocalPath(ActiveDocument.Path, , True)
    localPdfPath = GetLocalPath(pdfPath, , True)
    
    ' Calculate relative path using LibFileTools
    relativePdfPath = GetRelativePath(localPdfPath, localDocPath)
    
    ' Construct the placeholder string.
    placeholderText = "[[INSERT: " & relativePdfPath & "]]"
    
    ' Insert the placeholder into a new paragraph as plain text (no content control).
    With Selection
        .TypeParagraph
        .TypeText Text:=placeholderText
    End With
    
End Sub

Public Sub InsertOverlayPlaceholder(control As IRibbonControl)
    ' Inserts a table-based placeholder for overlaying a PDF page.
    
    Dim pdfPath As String
    Dim localDocPath As String
    Dim localPdfPath As String
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
    
    ' Convert OneDrive/SharePoint paths to local paths for both document and selected PDF
    localDocPath = GetLocalPath(ActiveDocument.Path, , True)
    localPdfPath = GetLocalPath(pdfPath, , True)
    
    ' Calculate relative path using LibFileTools
    relativePdfPath = GetRelativePath(localPdfPath, localDocPath)
    
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
    
    ' Insert the placeholder text into the cell as plain text (no content control).
    tbl.Cell(1, 1).Range.Text = placeholderText
    
End Sub

Public Sub InsertPdfAsSvg(control As IRibbonControl)
    ' Converts a PDF page to SVG and inserts it as an image.
    
    Dim pdfPath As String
    Dim localDocPath As String
    Dim localPdfPath As String
    Dim pageNumber As String
    Dim intPageNumber As Integer
    Dim tempSvgFolder As String
    Dim tempSvgPath As String
    Dim cmdString As String
    Dim doc As Document
    Dim fso As Object
    Dim i As Integer
    Dim maxWait As Integer
    
    Set doc = ActiveDocument
    
    ' Ensure the document is saved before proceeding.
    If doc.Path = "" Then
        MsgBox "Please save the document first to create a temporary folder for SVG conversion.", vbExclamation, "Save Document"
        Exit Sub
    End If
    
    ' Get the path to the PDF file from the user.
    pdfPath = GetPdfPath()
    If pdfPath = "" Then Exit Sub ' User cancelled
    
    ' Convert OneDrive/SharePoint paths to local paths
    localDocPath = GetLocalPath(doc.Path, , True)
    localPdfPath = GetLocalPath(pdfPath, , True)
    
    ' Get the page range from the user.
    pageNumber = InputBox("Enter page number(s) to convert:" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "• 1 (single page)" & vbCrLf & _
                         "• 1-3 (pages 1 to 3)" & vbCrLf & _
                         "• 1,3,5 (specific pages)" & vbCrLf & _
                         "• all (all pages - default)", _
                         "PDF Page Selection", "all")
    If pageNumber = "" Then Exit Sub ' User cancelled
    
    ' Normalize input
    pageNumber = Trim(LCase(pageNumber))
    If pageNumber = "" Then pageNumber = "all"
    
    ' Create temporary folder using local document path
    tempSvgFolder = localDocPath & "\temp-svg"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(tempSvgFolder) Then
        fso.CreateFolder tempSvgFolder
    End If
    
    ' Define temporary SVG path (base name)
    tempSvgPath = tempSvgFolder & "\page.svg"
    
    ' Build the command string for SVG conversion using local PDF path
    cmdString = "uvx report-compiler svg-import --page " & pageNumber & " " & _
                Chr(34) & localPdfPath & Chr(34) & " " & Chr(34) & tempSvgPath & Chr(34)
    
    Debug.Print "Executing command: " & cmdString

    ' Execute the conversion command
    On Error Resume Next
    Shell cmdString, vbHide
    If Err.Number <> 0 Then
        MsgBox "Failed to start the PDF to SVG converter. Please check that uvx is installed and in your PATH.", vbCritical, "Execution Error"
        On Error GoTo 0
        GoTo Cleanup
    End If
    On Error GoTo 0
    
    ' Wait for SVG files to be created (up to 30 seconds)
    maxWait = 300 ' 30 seconds in 100ms increments
    For i = 1 To maxWait
        ' Check if any SVG files exist in the temp folder
        If fso.FolderExists(tempSvgFolder) Then
            If fso.GetFolder(tempSvgFolder).Files.Count > 0 Then
                Exit For
            End If
        End If
        ' Wait 100ms (Word VBA does not have Application.Wait)
        Dim waitTime As Single
        waitTime = Timer + 0.1
        Do While Timer < waitTime
            DoEvents
        Loop
    Next i
    
    ' Check if any files were created
    If Not fso.FolderExists(tempSvgFolder) Or fso.GetFolder(tempSvgFolder).Files.Count = 0 Then
        MsgBox "SVG conversion failed or timed out. Please check the PDF file and page specification.", vbExclamation, "Conversion Failed"
        GoTo Cleanup
    End If
    
    ' Insert all SVG files found in the temp folder
    Dim svgFiles As Object
    Dim svgFile As Object
    Dim insertedCount As Integer
    
    Set svgFiles = fso.GetFolder(tempSvgFolder).Files
    insertedCount = 0
    
    On Error GoTo InsertError
    For Each svgFile In svgFiles
        If LCase(Right(svgFile.Name, 4)) = ".svg" Then
            ' Insert each SVG as an image
            Selection.InlineShapes.AddPicture fileName:=svgFile.Path, LinkToFile:=False, SaveWithDocument:=True
            insertedCount = insertedCount + 1
            
            ' Add a line break after each image except the last one
            If insertedCount < svgFiles.Count Then
                Selection.TypeParagraph
            End If
        End If
    Next svgFile
    On Error GoTo 0
    
    If insertedCount > 0 Then
        MsgBox "Successfully inserted " & insertedCount & " PDF page(s) as SVG images!", vbInformation, "Success"
    Else
        MsgBox "No SVG files were created during conversion.", vbExclamation, "No Files Found"
    End If
    
    ' Clean up temporary files
Cleanup:
    On Error Resume Next
    If fso.FolderExists(tempSvgFolder) Then
        ' Delete all SVG files in the temp folder
        Dim tempFiles As Object
        Dim tempFile As Object
        Set tempFiles = fso.GetFolder(tempSvgFolder).Files
        For Each tempFile In tempFiles
            If LCase(Right(tempFile.Name, 4)) = ".svg" Then
                fso.DeleteFile tempFile.Path
            End If
        Next tempFile
        
        ' Delete folder if it's empty
        If fso.GetFolder(tempSvgFolder).Files.Count = 0 And fso.GetFolder(tempSvgFolder).SubFolders.Count = 0 Then
            fso.DeleteFolder tempSvgFolder
        End If
    End If
    On Error GoTo 0
    Exit Sub

InvalidPageNumber:
    MsgBox "Invalid page specification. Please enter a valid page number, range (1-3), list (1,3,5), or 'all'.", vbExclamation, "Invalid Input"
    Exit Sub

InsertError:
    MsgBox "Failed to insert one or more SVG images. Some files may be corrupted or in an unsupported format.", vbExclamation, "Insert Error"
    GoTo Cleanup
    
End Sub

Public Sub RunReportCompiler(control As IRibbonControl)
    ' Saves the active document and executes the Python compiler script.
    
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
    inputPath = GetLocalPath(doc.FullName)
    outputPath = Replace(inputPath, ".docx", ".pdf")
    
    ' Build the command string for the shell. Paths are wrapped in quotes.
    cmdString = "uvx report-compiler compile " & Chr(34) & inputPath & Chr(34) & " " & Chr(34) & outputPath & Chr(34)
    
    ' Execute the command. vbNormalFocus shows the console window so users can see progress.
    On Error Resume Next
    Shell cmdString, vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Failed to start the compiler. Please check that uvx is installed and in your PATH.", vbCritical, "Execution Error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Inform the user that the process has started.
    MsgBox "The report compiler has been started. You can monitor its progress in the console window.", vbInformation, "Compiler Started"
    
End Sub


'--------------------------------------------------------------------------------------------------
' PRIVATE HELPER FUNCTIONS
'--------------------------------------------------------------------------------------------------

Private Function GetPdfPath() As String
    ' Opens a file picker dialog for the user to select a PDF file.
    ' Returns the full path of the selected file, or an empty string if cancelled.
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a PDF File"
        .allowMultiSelect = False
        
        ' Clear existing filters and add one for PDF files.
        .filters.Clear
        .filters.Add "PDF Files", "*.pdf"
        .filters.Add "Word Files", "*.docx"
        
        ' Show the dialog and check if a file was selected.
        If .Show = True Then
            GetPdfPath = .SelectedItems(1)
        Else
            GetPdfPath = "" ' User cancelled
        End If
    End With
    
End Function


