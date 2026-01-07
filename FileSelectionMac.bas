Attribute VB_Name = "FileSelectionMac"
'Used by CopyPrealert and CopyLentoAndRekka Subs, do not edit

Function Select_File_Or_Files_Mac(FolderPath As String) As String()
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim SplitPath As Variant
    Dim MySplit As Variant
    Dim N As Long, i As Long
    Dim FName As String
    Dim mybook As Workbook
    
    On Error Resume Next
    'MyPath = MacScript("return (path to documents folder) as String")
    MyPath = "Macintosh HD" & Replace(FolderPath, "/", ":")
    

    ' In the following statement, change true to false in the line "multiple
    ' selections allowed true" if you do not want to be able to select more
    ' than one file. Additionally, if you want to filter for multiple files, change
    ' {""com.microsoft.Excel.xls""} to
    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
    ' if you want to filter on xls and csv files, for example.
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFiles to (choose file of type " & _
             " {""com.microsoft.excel.xls"", ""org.openxmlformats.spreadsheetml.sheet"", ""org.openxmlformats.spreadsheetml.sheet.macroenabled"", ""public.comma-separated-values-text""}  " & _
               "with prompt ""Please select a file or files"" default location alias """ & _
               MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
               "set applescript's text item delimiters to """" " & vbNewLine & _
               "return theFiles"

    MyFiles = MacScript(MyScript)
    Dim returnList() As String
    On Error GoTo 0

    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        'MsgBox MyFiles
        MySplit = Split(MyFiles, ",")
        ReDim returnList(LBound(MySplit) To UBound(MySplit))
        For N = LBound(MySplit) To UBound(MySplit)

            returnList(N) = MySplit(N)

        Next N
        For i = LBound(returnList) To UBound(returnList)
            SplitPath = Split(returnList(i), ":")
            FileName = SplitPath(UBound(SplitPath))
            returnList(i) = FolderPath & FileName
        Next i
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
        Select_File_Or_Files_Mac = returnList
    Else
        ReDim returnList(0 To 0)
        returnList(0) = "False"
        Select_File_Or_Files_Mac = returnList
    End If
End Function



