Attribute VB_Name = "GetPrealertAndLentoRekka"
'Disables background processes to improve performance and UX
Sub DisableProcesses()
    With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
    End With
End Sub
Sub EnableProcesses()
     With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
    End With
End Sub
'Closes all workbooks in a collection
Sub CloseWorkbooks(WbColl As Collection)
    On Error Resume Next
    Dim wb As Workbook
    If WbColl.Count > 0 Then
        For Each wb In WbColl
            wb.Application.CutCopyMode = False
            wb.Close
        Next
    End If
    On Error GoTo 0
End Sub
'Gets the current date (dd.mm.yyyy)
Function GetCurrentDate()
    Dim Dt As Date
    Dt = Date
    GetCurrentDate = Dt
End Function
'Formats the date given into yyyymmd and converts it into string
Function FormatDate(Dt As Date)
    Dim FormattedDt As String
    FormattedDt = Format(Dt, "yyyymmd")
    FormatDate = FormattedDt
End Function
'Formats the date given into ddmmyyyy and converts it into string

Function FormatRekkaDate(Dt As Date)
    Dim FormattedDt As String
    FormattedDt = Format(Dt, "ddmmyyyy")
    FormatRekkaDate = FormattedDt
End Function

'Tries to set the variable Sheet to a worksheet in the target workbook, returns True if the worksheet exists and Sheet gains a value else it returns False
Function DoesSheetExist(wb As Workbook, WsName As String) As Boolean
    Dim Sheet As Worksheet
    On Error Resume Next
    Set Sheet = wb.Worksheets(WsName)
    On Error GoTo 0
    
    If Not Sheet Is Nothing Then DoesSheetExist = True
End Function

Function DoesColumnExist(wb As Workbook, WsName As String, ColName As String) As Boolean
    Dim Col As Range
    On Error Resume Next
    Set Col = wb.Worksheets(WsName).Columns(ColName)
    On Error GoTo 0
    
    If Not Col Is Nothing Then DoesColumnExist = True
    
End Function

'Finds the last row with data in a column
Function FindColumnLastRow(Worksheet As Worksheet, ColumnName As String) As Long
    Dim LastCell As Range
    Dim ColNum As Variant
    On Error Resume Next
    ColNum = Application.Match(ColumnName, Worksheet.Rows(1), 0)
    On Error GoTo 0
    If IsError(ColNum) Or IsEmpty(ColNum) Then
        FindColumnLastRow = 0
        Exit Function
    End If
    Set LastCell = Worksheet.Columns(ColNum).Find(What:="*", _
                    After:=Worksheet.Cells(1, ColNum), _
                    LookIn:=xlValues, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
    If LastCell Is Nothing Then
        FindColumnLastRow = 0
    Else
        FindColumnLastRow = LastCell.Row
    End If
   
End Function
Function FindLastColumn(wb As Workbook, WsName As String)
    Dim LastCell As Range
    Set LastCell = wb.Worksheets(WsName).Cells.Find(What:="*", _
                    After:=Cells(1, 1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
    
    If LastCell Is Nothing Then
        FindLastColumn = 0
    Else
        FindLastColumn = LastCell.Column
    End If
End Function
'Loops through the column headers to find a match with the argument given
Function DoesColumnHeaderExist(wb As Workbook, WsName As String, HeaderName As String) As Boolean
    Dim i As Integer
    Dim LastColumn As Integer
    
    LastColumn = FindLastColumn(wb, WsName)
    
    For i = 1 To LastColumn
        If wb.Worksheets(WsName).Cells(1, i).Value = HeaderName Then
            DoesColumnHeaderExist = True
            Exit Function
        End If
    Next i
        
    DoesColumnHeaderExist = False
    
End Function

'Loops through the TargetFolder and returns the path of the file that matches the FormattedDate

Function MatchDateToFile(FormattedDate As String, TargetFolderPath As String) As String
    Dim Fs, Folder, File As Object
    Dim TargetFilePath As String, FileName As String
    #If Mac Then
        FileName = Dir(TargetFolderPath & "*.*")
        
        Do While FileName <> ""
            If InStr(FileName, FormattedDate) > 0 Then
                TargetFilePath = TargetFolderPath & FileName
             End If
             FileName = Dir()
        Loop
    #Else
        Set Fs = CreateObject("Scripting.FileSystemObject")
        Set Folder = Fs.GetFolder(TargetFolderPath)
        For Each File In Folder.Files
            If InStr(File.name, FormattedDate) > 0 Then
                TargetFilePath = TargetFolderPath & "\" & File.name
            End If
        Next File
    #End If
    MatchDateToFile = TargetFilePath
End Function

'Adds paths of todays prealert and previous days files based on the value in the config worksheet
Sub AddPathsToPrealertColl(CurrDate As Date, PrealertFolder As String, PathColl As Collection)
    Dim PathCount As Integer, i As Integer
    Dim DateColl As New Collection
    Dim FormattedDate As String, FilePath As String
    Dim Dt As Variant, Dte As Variant
    Dim Count As Long
    
    Count = 0
    Dte = CurrDate
    PathCount = 1 + ThisWorkbook.Worksheets("Config").Range("B5").Value
    Do While Count < PathCount
        If Weekday(Dte, vbMonday) <= 5 Then
            DateColl.Add (Dte)
            Count = Count + 1
        End If
        Dte = DateAdd("d", -1, Dte)
    Loop
    For Each Dt In DateColl
        FilePath = MatchDateToFile(FormatDate(CDate(Dt)), PrealertFolder)
        If FilePath <> "" Then PathColl.Add (FilePath)
    Next Dt
    
End Sub
'Copies the entire contents of a worksheet from a workbook
Sub CopySheetFromWb(SrcWb As Workbook, SrcWsName As String, TrackingNumberColumnName As String, DestWsName As String)
    Dim LRow As Long, LColumn As Long, StartingRow As Long
    
    StartingRow = 1
    
    If DoesSheetExist(SrcWb, SrcWsName) Then
        LRow = FindColumnLastRow(SrcWb.Worksheets(SrcWsName), TrackingNumberColumnName)
        LColumn = FindLastColumn(SrcWb, SrcWsName)
    Else
        MsgBox "Kohdetiedostossa ei ole " & SrcWsName & " nimistä taulukkoa."
        Exit Sub
    End If
    
    If DoesColumnHeaderExist(ThisWorkbook, DestWsName, TrackingNumberColumnName) = True Then
        StartingRow = 2
    End If
    
    If LRow > 0 And LColumn > 0 Then
         With SrcWb.Worksheets(SrcWsName)
            .Range(.Cells(StartingRow, 1), .Cells(LRow, LColumn)).Copy
        End With
    Else
        Err.Raise vbObjectError + 2, , " Kopioitavaa dataa ei löytynyt kohteessa " & SrcWb.name & _
        " Tyhjennä RaakaDataAP-lista, jos sinne on siirtynyt osittaista dataa"
    End If
End Sub
'Pastes the contents of a worksheet to this workbook starting from the row with the first empty cell in TrackingNumberColumnName
Sub PasteSheet(SrcWb As Workbook, SrcWsName As String, DestWsName As String, TrackingNumberColumnName As String)
    
    Dim DestLRow As Long, DestLColumn As Long
    Dim SrcLRow As Long, SrcLCol As Long
    
    SrcLRow = FindColumnLastRow(SrcWb.Worksheets(SrcWsName), TrackingNumberColumnName)
    SrcLCol = FindLastColumn(SrcWb, SrcWsName)
    If DoesSheetExist(ThisWorkbook, DestWsName) = True Then
        If DoesColumnHeaderExist(ThisWorkbook, DestWsName, TrackingNumberColumnName) = True Then
            DestLRow = FindColumnLastRow(ThisWorkbook.Worksheets(DestWsName), TrackingNumberColumnName)
            DestLColumn = FindLastColumn(ThisWorkbook, DestWsName)
        Else
            DestLRow = 0
            DestLColumn = SrcLCol
        End If
    Else
        MsgBox "Kohdetaulukkoa ei ole olemassa"
        Exit Sub
    End If
    
    With ThisWorkbook.Worksheets(DestWsName)
            .Range(.Cells(DestLRow + 1, 1), .Cells(DestLRow + SrcLRow + 1, DestLColumn)).PasteSpecial xlPasteAll
            .Range(.Cells(DestLRow + 1, 1), .Cells(DestLRow + SrcLRow + 1, DestLColumn)).Columns.AutoFit
    End With

End Sub

Sub DisplayCopiedPaths(StartRange As Range, PathColl As Collection, Header As String)

    Dim Path As String
    Dim i As Long, StartRow As Long, StartCol As Long, LastRow As Long
    Dim Ws As Worksheet
    Dim Rng As Range, EmptyCell As Range
    Set Ws = ThisWorkbook.Worksheets("OHJAUSPANEELI")
    Dim Message As String
    
    Message = "Vain 10 polkua voidaan näyttää tiedostopoluissa." & vnNewLine & _
                " Tiedostot on todennäköisesti kopioitu onnistuneesti."
    
    If Header = "LentoRekka" Then
        StartRow = StartRange.Row
        StartCol = StartRange.Column
        Set Rng = Ws.Range(Ws.Cells(StartRow, StartCol), Ws.Cells(StartRow + 10, StartCol))
        Set EmptyCell = Rng.Find("", LookIn:=xlValues)
        LastRow = 15
    Else
        StartRow = StartRange.Row + 12
        StartCol = StartRange.Column
        Set Rng = Ws.Range(Ws.Cells(StartRow, StartCol), Ws.Cells(StartRow + 22, StartCol))
        Set EmptyCell = Rng.Find("", LookIn:=xlValues)
        LastRow = 27
    End If
    If EmptyCell Is Nothing Then
        MsgBox Message
        Exit Sub
    End If
    
    Ws.Cells(StartRow, StartCol).Value = Header
    For i = 0 To PathColl.Count - 1
        If EmptyCell.Row + i > LastRow Then
            MsgBox Message
            Exit Sub
        End If
        Ws.Cells(EmptyCell.Row + i, StartCol).Value = PathColl(i + 1)
    Next i
    
End Sub



'Cleanup subprocess for CopyPreAlert and CopyLentoRekka
Sub Cleanup(WbColl As Collection, TargetSheetName As String)

    On Error Resume Next
    Call CloseWorkbooks(WbColl)
    ThisWorkbook.Worksheets(TargetSheetName).Range("A1").Select
    ThisWorkbook.Worksheets("OHJAUSPANEELI").Activate
#If Not Mac Then
    ChDir "C:\"
#End If
    Call EnableProcesses
    On Error GoTo 0
End Sub
'Adds paths of todays Rekka file and previous days files based on the value in the config worksheet
Sub AddPathsToRekkaPathColl(CurrDate As Date, RekkaFolder As String, RekkaPathColl As Collection)

    Dim PathCount As Integer, i As Integer
    Dim DateColl As New Collection
    Dim FormattedDate As String, FilePath As String
    Dim Dt As Variant, Dte As Variant
    Dim Count As Long
    
    Count = 0
    Dte = CurrDate
    PathCount = 1 + ThisWorkbook.Worksheets("Config").Range("B11").Value
    Do While Count < PathCount
        If Weekday(Dte, vbMonday) <= 5 Then
            DateColl.Add (Dte)
            Count = Count + 1
        End If
        Dte = DateAdd("d", -1, Dte)
    Loop
    For Each Dt In DateColl
        FilePath = MatchDateToFile(FormatRekkaDate(CDate(Dt)), RekkaFolder)
        If FilePath <> "" Then RekkaPathColl.Add (FilePath)
    Next Dt
End Sub


Sub AddPathsToLentoPathColl(CurrDate As Date, LentoFolder As String, LentoPathColl As Collection)

    Dim PathCount As Integer, i As Integer
    Dim DateColl As New Collection
    Dim FormattedDate As String, FilePath As String
    Dim Dt As Variant, Dte As Variant
    Dim Count As Long
    
    Count = 0
    Dte = CurrDate
    PathCount = 1 + ThisWorkbook.Worksheets("Config").Range("B11").Value
    Do While Count < PathCount
        If Weekday(Dte, vbMonday) <= 5 Then
            DateColl.Add (Dte)
            Count = Count + 1
        End If
        Dte = DateAdd("d", -1, Dte)
    Loop
    For Each Dt In DateColl
        FilePath = MatchDateToFile(CDate(Dt), LentoFolder)
        If FilePath <> "" Then LentoPathColl.Add (FilePath)
    Next Dt
End Sub


'Automatically copies the prealerts from target folder and adds them to the "RaakaDataAP-lista" worksheet
Sub CopyPrealertAuto()

On Error GoTo ErrH
    Call DisableProcesses
    Dim PrealertFolder As String, Path As String, FormattedCurrDate As String
    Dim CurrDate As Date
    Dim i As Integer
    Dim wb As Workbook
    Dim PrealertPathColl As New Collection, PrealertColl As New Collection
    Dim SrcWsName As String
    
    SrcWsName = ThisWorkbook.Worksheets("Config").Range("B3").Value
    PrealertFolder = ThisWorkbook.Worksheets("Config").Range("B4")
    CurrDate = GetCurrentDate()
    AddPathsToPrealertColl CurrDate, PrealertFolder, PrealertPathColl
    
    If PrealertPathColl.Count = 0 Then
        MsgBox "Yhtäkään hakuehtoja täyttävää tiedostoa ei löytynyt. Tarkista tiedostonimien päivämäärät ja yritä uudelleen"
        GoTo ExitSub
    End If
    For i = 1 To PrealertPathColl.Count
        Set wb = Workbooks.Open(PrealertPathColl(i))
        PrealertColl.Add wb
    Next i
    
    For i = 1 To PrealertColl.Count
        Call CopySheetFromWb(PrealertColl(i), SrcWsName, "Trackingnumber", "RaakaDataAP-lista")
        Call PasteSheet(PrealertColl(i), SrcWsName, "RaakaDataAP-lista", "Trackingnumber")
    Next i
    
ExitSub:
    
    Call Cleanup(PrealertColl, "RaakaDataAP-lista")
    Call DisplayCopiedPaths(ThisWorkbook.Worksheets("OHJAUSPANEELI").Range("R5"), PrealertPathColl, "Prealert")
    
    Exit Sub

ErrH:
    MsgBox "Tapahtui virhe " & Err.Description
    Resume ExitSub
End Sub

'Copies the selected files to to the "RaakaDataAP-lista" worksheet
Sub CopyPrealertManual()

On Error GoTo ErrH
    Call DisableProcesses
    Dim PrealertFolder As String, Path As String, FormattedCurrDate As String, SrcWsName As String
    Dim CurrDate As Date
    Dim i As Long
    Dim wb As Workbook
    Dim PrealertPathColl As New Collection, PrealertColl As New Collection
    Dim Paths As Variant
    
    SrcWsName = ThisWorkbook.Worksheets("Config").Range("B3").Value
    
    PrealertFolder = ThisWorkbook.Worksheets("Config").Range("B4")
    CurrDate = GetCurrentDate()
    
    #If Mac Then
        Paths = FileSelectionMac.Select_File_Or_Files_Mac(PrealertFolder)
        If Paths(0) = "False" Then
            MsgBox "Tiedostopolkua ei määritely tai tiedostoa ei löydy"
            Exit Sub
        End If
        
    #Else
        Paths = Application.GetOpenFilename("Excel files (*.xlsx; *.xlsm; *.xls), *.xlsx;*.xlsm;*.xls", , , , True)
        If VarType(Paths) = vbBoolean Then
            MsgBox "Tiedostopolkua ei määritely tai tiedostoa ei löydy"
            Exit Sub
        End If
    #End If

    For i = LBound(Paths) To UBound(Paths)
        PrealertPathColl.Add Paths(i)
            
    Next i
    
    For i = 1 To PrealertPathColl.Count
        Set wb = Workbooks.Open(PrealertPathColl(i))
        PrealertColl.Add wb
    Next i
    
    For i = 1 To PrealertColl.Count
        Call CopySheetFromWb(PrealertColl(i), SrcWsName, "Trackingnumber", "RaakaDataAP-lista")
        Call PasteSheet(PrealertColl(i), SrcWsName, "RaakaDataAP-lista", "Trackingnumber")
    Next i
    
ExitSub:
    Call Cleanup(PrealertColl, "RaakaDataAP-lista")
    Call DisplayCopiedPaths(ThisWorkbook.Worksheets("OHJAUSPANEELI").Range("R5"), PrealertPathColl, "Prealert")
    Exit Sub

ErrH:
    MsgBox "Tapahtui virhe " & Err.Description
    Resume ExitSub
    
End Sub


Sub CopyLentoAndRekkaAuto()

On Error GoTo ErrH
    Call DisableProcesses
    Dim LentoFolder As String, RekkaFolder As String, Path As String, FormattedCurrDate As String, SrcWsName As String
    Dim CurrDate As Date
    Dim i As Integer
    Dim wb As Workbook
    Dim LentoPathColl As New Collection, RekkaPathColl As New Collection, LentoRekkaPathColl As New Collection, LentoRekkaColl As New Collection
    
    SrcWsName = ThisWorkbook.Worksheets("Config").Range("B8").Value
    RekkaFolder = ThisWorkbook.Worksheets("Config").Range("B10")
    LentoFolder = ThisWorkbook.Worksheets("Config").Range("B9")
    CurrDate = GetCurrentDate()
   
    AddPathsToRekkaPathColl CurrDate, RekkaFolder, RekkaPathColl
     If RekkaPathColl.Count = 0 Then
        MsgBox "Yhtäkään hakuehtoja täyttävää Rekka tiedostoa ei löytynyt. Tarkista tiedostonimien päivämäärät ja yritä uudelleen"
        GoTo ExitSub
    End If
    
    AddPathsToLentoPathColl CurrDate, LentoFolder, LentoPathColl
    If LentoPathColl.Count = 0 Then
        MsgBox "Yhtäkään hakuehtoja täyttävää Lento tiedostoa ei löytynyt. Tarkista tiedostonimien päivämäärät ja yritä uudelleen"
        GoTo ExitSub
    End If
    
    For i = 1 To RekkaPathColl.Count
        Set wb = Workbooks.Open(RekkaPathColl(i))
        LentoRekkaColl.Add wb
        LentoRekkaPathColl.Add RekkaPathColl(i)
    Next i
    
    For i = 1 To LentoPathColl.Count
        Set wb = Workbooks.Open(LentoPathColl(i))
        LentoRekkaColl.Add wb
        LentoRekkaPathColl.Add LentoPathColl(i)
    Next i
    
    For i = 1 To LentoRekkaColl.Count
        Call CopySheetFromWb(LentoRekkaColl(i), SrcWsName, "Tracking", "LentoRekka-Raaka")
        Call PasteSheet(LentoRekkaColl(i), SrcWsName, "LentoRekka-Raaka", "Tracking")
    Next i

ExitSub:

    Call Cleanup(LentoRekkaColl, "LentoRekka-Raaka")
    Call DisplayCopiedPaths(ThisWorkbook.Worksheets("OHJAUSPANEELI").Range("R5"), LentoRekkaPathColl, "LentoRekka")

    Exit Sub

ErrH:
    MsgBox "Tapahtui virhe " & Err.Description
    Resume ExitSub
End Sub

'Copies the selected files to to the "LentoRekka-Raaka" worksheet
Sub CopyLentoAndRekkaManual()

On Error GoTo ErrH
    Call DisableProcesses
    Dim LentoFolder As String, RekkaFolder As String, Path As String, FormattedCurrDate As String, SrcWsName As String
    Dim LentoSelectionTitle As String, RekkaSelectionTitle As String
    Dim CurrDate As Date
    Dim i As Integer
    Dim wb As Workbook
    Dim LentoPathColl As New Collection, RekkaPathColl As New Collection, LentoRekkaPathColl As New Collection, LentoRekkaColl As New Collection
    Dim LentoPathArr As Variant, RekkaPathArr As Variant
    
    LentoSelectionTitle = "Valitse yksi tai useampi Lento tiedosto"
    RekkaSelectionTitle = "Valitse yksi tai useampi Rekka tiedosto"
    SrcWsName = ThisWorkbook.Worksheets("Config").Range("B8").Value
    RekkaFolder = ThisWorkbook.Worksheets("Config").Range("B10")
    LentoFolder = ThisWorkbook.Worksheets("Config").Range("B9")
    CurrDate = GetCurrentDate()
    
    #If Mac Then
        
        RekkaPathArr = FileSelectionMac.Select_File_Or_Files_Mac(RekkaFolder)
        LentoPathArr = FileSelectionMac.Select_File_Or_Files_Mac(LentoFolder)
        If RekkaPathArr(0) = "False" And LentoPathArr(0) = "False" Then
            MsgBox "Tiedostopolkuja ei määritely tai tiedostoja ei löydy"
            Exit Sub
        End If
         
        If RekkaPathArr(0) <> "False" Then
            For i = LBound(RekkaPathArr) To UBound(RekkaPathArr)
                RekkaPathColl.Add RekkaPathArr(i)
            Next i
        End If
        
        If LentoPathArr(0) <> "False" Then
             For i = LBound(LentoPathArr) To UBound(LentoPathArr)
                LentoPathColl.Add LentoPathArr(i)
            Next i
        End If
        
        
    #Else
        ChDir RekkaFolder
        
        RekkaPathArr = Application.GetOpenFilename("Excel files (*.xlsx; *.xlsm; *.xls), *.xlsx;*.xlsm;*.xls", , RekkaSelectionTitle, , True)
            
        ChDir LentoFolder
        
        LentoPathArr = Application.GetOpenFilename("Excel files (*.xlsx; *.xlsm; *.xls), *.xlsx;*.xlsm;*.xls", , LentoSelectionTitle, , True)
        
        If VarType(LentoPathArr) = vbBoolean And VarType(RekkaPathArr) = vbBoolean Then
            MsgBox "Tiedostopolkuja ei määritelty tai tiedostoja ei löydy"
            GoTo ExitSub
        End If
        
        If VarType(RekkaPathArr) <> vbBoolean Then
            For i = LBound(RekkaPathArr) To UBound(RekkaPathArr)
                RekkaPathColl.Add RekkaPathArr(i)
            Next i
        End If
        
        If VarType(LentoPathArr) <> vbBoolean Then
             For i = LBound(LentoPathArr) To UBound(LentoPathArr)
                LentoPathColl.Add LentoPathArr(i)
            Next i
        End If
        
    #End If
    
        For i = 1 To RekkaPathColl.Count
            Set wb = Workbooks.Open(RekkaPathColl(i))
            LentoRekkaColl.Add wb
            LentoRekkaPathColl.Add RekkaPathColl(i)
        Next i
    
        For i = 1 To LentoPathColl.Count
            Set wb = Workbooks.Open(LentoPathColl(i))
            LentoRekkaColl.Add wb
            LentoRekkaPathColl.Add LentoPathColl(i)
        Next i
    
        For i = 1 To LentoRekkaColl.Count
            Call CopySheetFromWb(LentoRekkaColl(i), SrcWsName, "Tracking", "LentoRekka-Raaka")
            Call PasteSheet(LentoRekkaColl(i), SrcWsName, "LentoRekka-Raaka", "Tracking")
        Next i
    
ExitSub:

    Call Cleanup(LentoRekkaColl, "LentoRekka-Raaka")
    Call DisplayCopiedPaths(ThisWorkbook.Worksheets("OHJAUSPANEELI").Range("R5"), LentoRekkaPathColl, "LentoRekka")
    Exit Sub

ErrH:
    MsgBox "Tapahtui virhe " & Err.Description
    Resume ExitSub
        
End Sub


