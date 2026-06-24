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
    ' Launches the PDF Overlay dialog via the COM server. The dialog (Python/PySide6)
    ' shows page previews, lets the user pick a page range, then inserts the
    ' [[OVERLAY: ...]] placeholder back into this document at the bookmark below.

    Dim doc As Document
    Dim localDocPath As String
    Dim anchorName As String
    Dim compiler As Object

    Set doc = ActiveDocument

    ' Ensure the document is saved before creating a relative path.
    If doc.Path = "" Then
        MsgBox "Please save the document first to create a relative path for the overlay.", vbExclamation, "Save Document"
        Exit Sub
    End If

    ' Resolve OneDrive/SharePoint path to a local path (kept in LibFileTools).
    localDocPath = GetLocalPath(doc.FullName, , True)

    ' Drop a bookmark at the current selection so the dialog knows where to insert.
    anchorName = "RC_OverlayAnchor"
    On Error Resume Next
    doc.Bookmarks(anchorName).Delete
    On Error GoTo 0
    doc.Bookmarks.Add Name:=anchorName, Range:=Selection.Range

    ' Launch the dialog (returns immediately; Word stays responsive).
    On Error GoTo ComError
    Set compiler = CreateObject("ReportCompiler.Application")
    compiler.LaunchOverlayDialog localDocPath, anchorName
    Set compiler = Nothing
    Exit Sub

ComError:
    MsgBox "Could not reach the Report Compiler COM server." & vbCrLf & vbCrLf & _
           "Register it first by running:" & vbCrLf & _
           "    uvx report-compiler com-server register", _
           vbCritical, "COM Server Not Registered"
    Set compiler = Nothing
End Sub

Public Sub InsertImagePlaceholder(control As IRibbonControl)
    ' Inserts a table-based placeholder for embedding an image file.
    
    Dim imagePath As String
    Dim localDocPath As String
    Dim localImagePath As String
    Dim relativeImagePath As String
    Dim placeholderText As String
    Dim tbl As Table
    
    ' Ensure the document is saved before creating a relative path.
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first to create a relative path for the image.", vbExclamation, "Save Document"
        Exit Sub
    End If

    ' Get the path to the image file from the user.
    imagePath = GetImagePath()
    If imagePath = "" Then Exit Sub ' User cancelled
    
    ' Convert OneDrive/SharePoint paths to local paths for both document and selected image
    localDocPath = GetLocalPath(ActiveDocument.Path, , True)
    localImagePath = GetLocalPath(imagePath, , True)
    
    ' Calculate relative path using LibFileTools
    relativeImagePath = GetRelativePath(localImagePath, localDocPath)
    
    ' Construct the placeholder string.
    placeholderText = "[[IMAGE: " & relativeImagePath & "]]"
    
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
    Dim doc As Document
    Dim fso As Object
    Dim compiler As Object
    Dim jobId As String
    Dim jobStatus As String
    Dim waitTime As Single

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
    
    ' Convert via the Report Compiler COM server (async job + status polling),
    ' instead of shelling out. Word stays responsive while we poll.
    On Error GoTo ComError
    Set compiler = CreateObject("ReportCompiler.Application")
    jobId = compiler.SvgImportAsync(localPdfPath, tempSvgPath, pageNumber)

    Do
        ' Yield to Word and wait ~250ms between status checks (VBA has no Sleep).
        waitTime = Timer + 0.25
        Do While Timer < waitTime
            DoEvents
        Loop
        jobStatus = compiler.GetJobStatus(jobId)
    Loop While jobStatus = "pending" Or jobStatus = "running"
    On Error GoTo 0

    If jobStatus <> "succeeded" Then
        MsgBox "SVG conversion failed:" & vbCrLf & compiler.GetJobMessage(jobId), vbExclamation, "Conversion Failed"
        Set compiler = Nothing
        GoTo Cleanup
    End If
    Set compiler = Nothing

    ' Safety net: confirm files actually landed in the temp folder.
    If Not fso.FolderExists(tempSvgFolder) Or fso.GetFolder(tempSvgFolder).Files.Count = 0 Then
        MsgBox "Conversion reported success but no SVG files were found.", vbExclamation, "No Files Found"
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

ComError:
    MsgBox "Could not reach the Report Compiler COM server." & vbCrLf & vbCrLf & _
           "Register it first by running:" & vbCrLf & _
           "    uvx report-compiler com-server register", _
           vbCritical, "COM Server Not Registered"
    Set compiler = Nothing
    GoTo Cleanup

End Sub

Public Sub RunReportCompiler(control As IRibbonControl)
    ' Saves the active document and compiles it via the Report Compiler COM server.
    ' The COM server runs the compile on a background thread and we poll for status,
    ' so Word stays responsive and we can report real success/failure.

    Dim doc As Document
    Dim inputPath As String
    Dim outputPath As String
    Dim compiler As Object
    Dim jobId As String
    Dim jobStatus As String
    Dim waitTime As Single

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

    ' Connect to the Report Compiler COM server (registered per-user via
    ' 'uvx report-compiler com-server register').
    On Error GoTo ComError
    Set compiler = CreateObject("ReportCompiler.Application")

    ' Start the compile (returns immediately with a job id) and poll for completion.
    jobId = compiler.CompileAsync(inputPath, outputPath)

    Do
        ' Yield to Word and wait ~250ms between status checks (VBA has no Sleep).
        waitTime = Timer + 0.25
        Do While Timer < waitTime
            DoEvents
        Loop
        jobStatus = compiler.GetJobStatus(jobId)
    Loop While jobStatus = "pending" Or jobStatus = "running"

    On Error GoTo 0

    If jobStatus = "succeeded" Then
        MsgBox "Report compiled successfully:" & vbCrLf & outputPath, vbInformation, "Compile Complete"
    Else
        MsgBox "Report compilation failed:" & vbCrLf & compiler.GetJobMessage(jobId), vbCritical, "Compile Failed"
    End If

    ' Release the server (lets COM shut down the server process).
    Set compiler = Nothing
    Exit Sub

ComError:
    MsgBox "Could not reach the Report Compiler COM server." & vbCrLf & vbCrLf & _
           "Register it first by running:" & vbCrLf & _
           "    uvx report-compiler com-server register", _
           vbCritical, "COM Server Not Registered"
    Set compiler = Nothing
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

Private Function GetImagePath() As String
    ' Opens a file picker dialog for the user to select an image file.
    ' Returns the full path of the selected file, or an empty string if cancelled.
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select an Image File"
        .allowMultiSelect = False
        
        ' Clear existing filters and add common image formats.
        .filters.Clear
        .filters.Add "All Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff;*.tif;*.svg;*.emf;*.wmf"
        .filters.Add "PNG Files", "*.png"
        .filters.Add "JPEG Files", "*.jpg;*.jpeg"
        .filters.Add "GIF Files", "*.gif"
        .filters.Add "BMP Files", "*.bmp"
        .filters.Add "TIFF Files", "*.tiff;*.tif"
        .filters.Add "SVG Files", "*.svg"
        .filters.Add "EMF/WMF Files", "*.emf;*.wmf"
        
        ' Show the dialog and check if a file was selected.
        If .Show = True Then
            GetImagePath = .SelectedItems(1)
        Else
            GetImagePath = "" ' User cancelled
        End If
    End With

End Function


Public Sub OnOverlayViewChange(control As IRibbonControl, id As String, index As Integer)
    ' Ribbon "Overlay view" dropdown: 0 = Tags, 1 = Quick preview, 2 = Full preview.
    Dim mode As String
    Select Case index
        Case 1
            mode = "quick"
        Case 2
            mode = "full"
        Case Else
            mode = "tags"
    End Select
    ApplyOverlayView mode
End Sub

Public Sub RepairOverlays(control As IRibbonControl)
    ' Force every overlay back to its canonical tag form (recovery).
    ApplyOverlayView "tags"
End Sub

Private Sub ApplyOverlayView(mode As String)
    ' Drives the COM server to switch the whole document's overlay view.
    Dim doc As Document
    Dim localDocPath As String
    Dim compiler As Object

    Set doc = ActiveDocument
    If doc.Path = "" Then
        MsgBox "Please save the document first.", vbExclamation, "Save Document"
        Exit Sub
    End If
    localDocPath = GetLocalPath(doc.FullName, , True)

    On Error GoTo ComError
    Set compiler = CreateObject("ReportCompiler.Application")
    compiler.SetOverlayPreview localDocPath, mode
    Set compiler = Nothing
    Exit Sub

ComError:
    MsgBox "Could not reach the Report Compiler COM server." & vbCrLf & vbCrLf & _
           "Register it first by running:" & vbCrLf & _
           "    uvx report-compiler com-server register", _
           vbCritical, "COM Server Not Registered"
    Set compiler = Nothing
End Sub


