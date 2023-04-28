Private Sub Convert_Button_Click()

    Dim InputPath As String
    Dim OutputPath As String

    InputPath = InputFolder_Path_Label.Caption
    OutputPath = OutputFolder_Path_Label.Caption

    'check Input folder is selected
    If InputPath = "" Then
        MsgBox "Please select an Input folder"
        Exit Sub
    End If

    'check Output folder is selected
    If OutputPath = "" Then
        MsgBox "Please select an Ouput folder"
        Exit Sub
    End If

    'check both Input and Output folder is selected
    If InputPath <> "" And OutputPath <> "" Then
        ListFilesInFolder InputPath, OutputPath
        Exit Sub
    Else
        MsgBox "Please select a valid Input and Output folder"
        Exit Sub
    End If

End Sub


Private Sub Done_Button_Click()
    Unload Me
End Sub


Private Sub InputFolder_Browse_Button_Click()

    Dim DialogBox As FileDialog
    Dim PATH_InputFolder As String
    Set DialogBox = Application.FileDialog(msoFileDialogFolderPicker)

    DialogBox.Title = "Select Input folder"
    DialogBox.Filters.Clear
    DialogBox.Show

    'checks only 1 path is selected from folder picker dialog
    If DialogBox.SelectedItems.Count = 1 Then
        PATH_InputFolder = DialogBox.SelectedItems(1)
    End If

    'set "Input Folder Path" label caption to the selected path
    InputFolder_Path_Label.Caption = PATH_InputFolder

End Sub


Private Sub OutputFolder_Browse_Button_Click()

    Dim DialogBox As FileDialog
    Dim PATH_OutputFolder As String
    Set DialogBox = Application.FileDialog(msoFileDialogFolderPicker)

    DialogBox.Title = "Select Output folder"
    DialogBox.Filters.Clear
    DialogBox.Show

    'checks only 1 path is selected from folder picker dialog
    If DialogBox.SelectedItems.Count = 1 Then
        PATH_OutputFolder = DialogBox.SelectedItems(1)
    End If

    'set "Output Folder Path" label caption to the selected path
    OutputFolder_Path_Label.Caption = PATH_OutputFolder

End Sub


Sub ListFilesInFolder(InputPath, OutputPath)

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim counter As Integer

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(InputPath)

    For Each oFile In oFolder.Files
    
        'if file isn't a .docx file, we can't convert it, so skip it
        If Right(oFile.Name, 4) <> "docx" Then
            GoTo NextIteration
        End If
        
        counter = counter + 1
        FileName = Left(oFile.Name, Len(oFile.Name) - 4) 'slice off "docx" from end of word filename

        'combines folder path with filename and extension
        DOCX_Path_Full = InputPath + "\" + FileName + "docx"
        PDF_Path_Full = OutputPath + "\" + FileName + "pdf"

        'handles if PDF_Path_Full does or does not already exist
        If FileExists(PDF_Path_Full) = False Then
            ConvertDocxToPdf DOCX_Path_Full, PDF_Path_Full
        Else
            ConvertDocxToPdf DOCX_Path_Full, PathName_PlusOne(PDF_Path_Full) 'if filename already exists, use PathName_PlusOne function to append a number to filename so new filename is unique
        End If
        
NextIteration:
    Next oFile

    ' "Finished converting" message
    If counter = 1 Then
        MsgBox "Finished converting " + CStr(counter) + " file."
    Else
        MsgBox "Finished converting " + CStr(counter) + " files."
    End If

End Sub


Sub ConvertDocxToPdf(InputFile_Path_DOCX, OutputFile_Path_PDF)

    Dim wordApp As Object
    Dim wordDoc As Object

    ' Create a Word application object
    Set wordApp = GetObject(, "Word.Application")
    Set wordDoc = GetObject(InputFile_Path_DOCX)

    ' Save the document as PDF
    wordDoc.SaveAs OutputFile_Path_PDF, wdFormatPDF

    ' Close the document without saving changes
    wordDoc.Close wdDoNotSaveChanges

End Sub


Function FileExists(path) As Boolean

    If Dir(path) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function


Function PathName_PlusOne(PDF_path) As String
    'we enter the function already assuming PDF_path exists.
    'Example PDF_path:  C:\Documents\00__Projects\zz__Other\2023-04-18 - Task for Celia - MS Word Automation\Test\Output\21-022 Denver Dam 1 - 00 00 00 - Index.pdf
        
    Dim counter As Integer
    Dim newFileName As String
    Dim FileName As String

    counter = 0
    newFileName = PDF_path
    FileName = Left(PDF_path, Len(PDF_path) - 4) 'removes ".pdf"

    'while "file.pdf" exists, change name to "file_(1).pdf", "file_(2).pdf", "file_(3).pdf", ...etc
    While FileExists(newFileName) = True

        counter = counter + 1
        full_suffix = "_(" + CStr(counter) + ")"
        
        newFileName = FileName + full_suffix + ".pdf"
        
    Wend

    PathName_PlusOne = newFileName

End Function

