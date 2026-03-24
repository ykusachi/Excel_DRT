Attribute VB_Name = "Module1"
Option Explicit

' --- Function to branch file selection based on the Operating System ---
Function WINorMAC() As Variant
    Dim MyFiles As Variant
    
    ' Test for the operating system (Check if it's NOT a Mac)
    If Not Application.OperatingSystem Like "*Mac*" Then
        
        ' Target: Windows OS
        MyFiles = Select_File_Or_Files_Windows
        
    Else
        
        ' Target: Mac OS (Test if running Excel 2011/Version 14 or higher)
        If Val(Application.Version) > 14 Then
            MyFiles = Select_File_Or_Files_Mac
        Else
            ' Error: Version not supported
            MsgBox "Error: This Mac Excel version is not supported.", vbCritical
            MyFiles = False
        End If
        
    End If
    
    ' Set the selected file(s) as the return value
    WINorMAC = MyFiles
    
End Function
    


' --- Function to display the file selection dialog on Windows ---
Function Select_File_Or_Files_Windows()
    Dim SaveDriveDir As String
    Dim MyPath As String
    Dim Fname As Variant
    Dim n As Long
    Dim FnameInLoop As String
    Dim mybook As Workbook

    ' Save the current directory to restore later
    SaveDriveDir = CurDir

    ' Set the target path to the application default
    MyPath = Application.DefaultFilePath

    ' Change current drive and directory to MyPath
    On Error Resume Next ' Avoid errors if the drive/path is invalid
    ChDrive MyPath
    ChDir MyPath
    On Error GoTo 0

    ' Open the file picker with Z filters and a custom title
    Fname = Application.GetOpenFilename( _
            FileFilter:="Z Files (*.z), *.z", _
            Title:="Select a file or files", _
            MultiSelect:=True)

    ' Restore the original drive and directory
    On Error Resume Next
    ChDrive SaveDriveDir
    ChDir SaveDriveDir
    On Error GoTo 0

    ' Return the selected file(s) (returns False if canceled)
    Select_File_Or_Files_Windows = Fname
    
End Function



' --- Function to display the file selection dialog on Mac using AppleScript ---
Function Select_File_Or_Files_Mac() As Variant
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim n As Long
    Dim Fname As String
    Dim mybook As Workbook

    On Error Resume Next
    ' Get the default path to the Documents folder
    MyPath = MacScript("return (path to documents folder) as String")
    
    ' Construct AppleScript to select files with .z extension
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
                "set theFiles to (choose file of type " & _
              " {""z""} " & _
                "with prompt ""Please select a .z file or files"" default location alias """ & _
                MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"

    ' Execute the AppleScript
    ' AppleScript を実行
    MyFiles = MacScript(MyScript)
    On Error GoTo 0
        
    ' Return the selected file(s) as the return value
    Select_File_Or_Files_Mac = MyFiles
    
End Function

    
    

' --- Function to check if a specific workbook is currently open ---
Function bIsBookOpen(ByRef szBookName As String) As Boolean
    ' Contributed by Rob Bovey
    
    ' Disable error handling to check for existence
    On Error Resume Next
    
    ' If the workbook is not found, the object will be Nothing
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
    
    ' Reset error handling
    On Error GoTo 0
End Function


' --- Function to split a full path into directory path and file name ---
' Returns: Array(Directory Path, File Name)
Function GetPathInfo(ByVal FullPath As String) As Variant
    Dim PathSeparator As String
    Dim LastSeparatorPos As Long
    Dim DirPath As String
    Dim FileName As String
    

    Dim i As Long
    Dim CheckChars As Variant
    
    ' List of separaton characters
    ' Chr(92) = \ , Chr(165) = ¥ , Chr(47) = /
    CheckChars = Array("\", "/", Chr(92), Chr(165), Chr(47))
    
    LastSeparatorPos = 0
    
    ' 
    For i = LBound(CheckChars) To UBound(CheckChars)
        Dim pos As Long
        pos = InStrRev(FullPath, CheckChars(i))
        If pos > LastSeparatorPos Then
            LastSeparatorPos = pos
        End If
    Next i
    

    
    If LastSeparatorPos > 0 Then
        ' Directory Path: Everything up to the last separator
        DirPath = Left(FullPath, LastSeparatorPos)
        
        ' File Name: Everything after the last separator
        FileName = Mid(FullPath, LastSeparatorPos + 1)
    Else
        ' If no separator is found, treat the whole path as the file name
        DirPath = ""
        FileName = FullPath
    End If
    
    ' Return as an array
    GetPathInfo = Array(DirPath, FileName)
    
End Function


' --- Subroutine to import CSV/Text files based on a list ---
Sub InsertTextCsvFiles()
    
    Dim targetSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim FilePath As String, FileName As String
    Dim NewSheet As Worksheet
    
    Set targetSheet = ActiveSheet
    ' Get the last row of the list in Column B
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Check if data exists in the list
    If lastRow < 2 Then
        MsgBox "List not found or contains no data.", vbExclamation
        Exit Sub
    End If
    
    ' Disable screen updates and alerts for performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through the list of files
    For i = 2 To lastRow
        FileName = targetSheet.Cells(i, "C").Value
        FilePath = targetSheet.Cells(i, "B").Value & FileName
        
        If FilePath <> "" And FileName <> "" Then
            
            ' --- 1. Sheet name generation logic ---
            Dim CleanSheetName As String
            
            ' Remove file extension
            If InStrRev(FileName, ".") > 0 Then
                CleanSheetName = Left(FileName, InStrRev(FileName, ".") - 1)
            Else
                CleanSheetName = FileName
            End If
            
            ' Replace illegal characters ( :  / ? * [ ] ) with underscore
            Dim illegalChars As Variant, charItem As Variant
            illegalChars = Array(":", "", "/", "?", "*", "[", "]")
            For Each charItem In illegalChars
                CleanSheetName = Replace(CleanSheetName, charItem, "_")
            Next charItem
            
            ' Trim to the last 25 characters to stay within Excel's limits
            If Len(CleanSheetName) > 25 Then
                CleanSheetName = Right(CleanSheetName, 25)
            End If
            
            ' Handle duplicate sheet names by adding a prefix
            Dim FinalSheetName As String
            Dim suffixIdx As Long
            Dim charList As String
            ' Prefix sequence: A-Z, then 0-9
            charList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            
            FinalSheetName = CleanSheetName
            suffixIdx = 0
            
            ' Loop until a unique sheet name is found
            Do While SheetExists(FinalSheetName)
                suffixIdx = suffixIdx + 1
                
                ' Add prefix (1 char + underscore) to the start of the name
                If suffixIdx <= Len(charList) Then
                    FinalSheetName = Mid(charList, suffixIdx, 1) & "_" & CleanSheetName
                Else
                    ' Use 2-digit number if pattern exceeds 36
                    FinalSheetName = Format(suffixIdx, "00") & "_" & CleanSheetName
                End If
                
                ' Ensure total length does not exceed 31 characters
                If Len(FinalSheetName) > 31 Then
                    FinalSheetName = Left(FinalSheetName, 31)
                End If
                
                If suffixIdx > 99 Then Exit Do ' Prevent infinite loop
            Loop
            
            ' --- Create New Sheet ---
            Set NewSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            NewSheet.Name = FinalSheetName
            
            ' --- Import Data via QueryTable ---
            On Error Resume Next
            With NewSheet.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=NewSheet.Range("A1"))
                .TextFilePlatform = 932 ' Shift-JIS
                .TextFileParseType = xlDelimited
                ' Check if CSV or Tab-delimited
                If UCase(Right(FileName, 4)) = ".CSV" Then
                    .TextFileCommaDelimiter = True
                Else
                    .TextFileTabDelimiter = True
                End If
                .Refresh BackgroundQuery:=False
            End With
            On Error GoTo 0
            
        End If
    Next i
    
    ' --- Post-processing Cleanup ---
    
    ' Restore screen updating and alerts
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Reset error handling
    On Error GoTo 0
    
    ' Call the data extraction procedure
    Call ExtractZData
    
    ' Return to the main "Top" sheet
    On Error Resume Next
    Dim wsList As Worksheet
    Set wsList = ThisWorkbook.Sheets("Top")
    
    If Not wsList Is Nothing Then
        wsList.Select
        wsList.Cells(1, 1).Select ' Place cursor at A1
    Else
        MsgBox "Sheet 'Top' was not found. Please check the sheet name.", vbCritical
    End If
    On Error GoTo 0

End Sub



' --- Function to check if a sheet with a specific name exists in the workbook ---
Function SheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    
    ' Disable error handling to attempt object assignment
    On Error Resume Next
    
    ' Try to set the worksheet object by name
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ' Reset error handling
    On Error GoTo 0
    
    ' If the object 'ws' is not Nothing, the sheet exists
    SheetExists = Not ws Is Nothing
    
End Function



' --- Subroutine to select multiple files and write their paths to the active sheet ---
Sub SelectFiles()
    Dim openWb As Workbook
    Dim openFileName As Variant, fileVar As Variant
    Dim InfoArray As Variant
    Dim WriteRow As Long ' Counter for writing info 
    Dim targetSheet As Worksheet ' Target sheet for writing 

    ' Branch process for Windows or Mac to select multiple files
    openFileName = WINorMAC
    
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' --- Case: Windows ---
        If IsEmpty(openFileName) Or openFileName(1) = False Then
            MsgBox "Action canceled by user." ' 
            Exit Sub
        End If
    Else
        ' --- Case: Mac ---
        If openFileName = "" Then
            MsgBox "Action canceled by user." ' 
            Exit Sub
        Else
            ' Split the string by commas and store into an array
            openFileName = Split(openFileName, ",")
        End If
    End If
    
    ' --- Start: Writing file information ---
    
    ' Set the active sheet as the destination
    Set targetSheet = ActiveSheet
    
    ' Set the starting row (e.g., Row 2 if there is a header)
    WriteRow = 1
    
    ' Loop through each selected file
    For Each fileVar In openFileName
        
        ' Path conversion for Mac environment
        If Application.OperatingSystem Like "*Mac*" Then
            ' Convert MacScript path format to a format recognizable by Workbooks.Open
            fileVar = Replace(Replace(fileVar, ":", "/"), "Macintosh HD", "")
        End If

        ' Retrieve file path information
        ' InfoArray(0) = Directory, InfoArray(1) = File Name
        InfoArray = GetPathInfo(CStr(fileVar))
        
        ' --- Write information to the sheet ---
        WriteRow = WriteRow + 1 ' Move to the next row 

        ' Column A: Serial Number 
        targetSheet.Cells(WriteRow, 1).Value = WriteRow - 1
        
        ' Column B: Directory of the selected file 
        targetSheet.Cells(WriteRow, 2).Value = InfoArray(0)
        
        ' Column C: File name of the selected file 
        targetSheet.Cells(WriteRow, 3).Value = InfoArray(1)
        
        ' Note: Original code for opening/closing files is commented out
        ' On Error Resume Next
        ' Workbooks.Open fileVar
        ' Set openWb = ActiveWorkbook
        ' ... [Processing] ...
        ' Application.DisplayAlerts = False
        ' If Not openWb Is Nothing Then openWb.Close
        ' Application.DisplayAlerts = True
        ' Set openWb = Nothing
        ' On Error GoTo 0
        
    Next fileVar
    
    ' --- End: Writing file information ---
    
    ' MsgBox "File information has been written to the active sheet."

End Sub



' --- Subroutine to extract frequency and impedance data from raw data sheets ---
Sub ExtractZData()
    Dim ws As Worksheet, extSheet As Worksheet
    Dim lastRow As Long, dataStartRow As Long, i As Long
    Dim targetName As String
    
    ' Disable screen updates for performance
    Application.ScreenUpdating = False
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Process sheets except those already ending in "ext" or the source list sheet
        If Not ws.Name Like "*ext" And ws.Name <> "Sheet1" And ws.Name <> "Top" Then
            
            dataStartRow = 0
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Search for the row containing header termination markers
            For i = 1 To lastRow
                Dim currentText As String
                currentText = ws.Cells(i, 1).Text
                
                If currentText Like "*End Comments*" Or currentText Like "*End Header*" Then
                    ' Data starts on the next row
                    dataStartRow = i + 1
                    Exit For
                End If
            Next i
            
            ' Proceed if data start row was found
            If dataStartRow > 0 And dataStartRow <= lastRow Then
                ' Adjust sheet name length to fit Excel's limit (31 chars)
                Dim safeBaseName As String
                safeBaseName = Left(ws.Name, 28)
                targetName = safeBaseName & "ext"
                
                ' Delete existing sheet with the same name if it exists
                On Error Resume Next
                Application.DisplayAlerts = False
                Sheets(targetName).Delete
                Application.DisplayAlerts = True
                On Error GoTo 0
                
                ' Add a new sheet after the current source sheet
                Set extSheet = ThisWorkbook.Sheets.Add(After:=ws)
                extSheet.Name = targetName
                
                ' Create Headers (Columns A to C)
                extSheet.Range("A1:C1").Value = Array("Freq(Hz)", "Z'", "Z''")
                
                ' Transfer (Copy) data values
                Dim rowCount As Long
                rowCount = lastRow - dataStartRow + 1
                
                ' Copy Freq (Col 1), Z' (Col 5), and Z'' (Col 6)
                ws.Cells(dataStartRow, 1).Resize(rowCount, 1).Copy extSheet.Range("A2") ' Freq
                ws.Cells(dataStartRow, 5).Resize(rowCount, 1).Copy extSheet.Range("B2") ' Z'
                ws.Cells(dataStartRow, 6).Resize(rowCount, 1).Copy extSheet.Range("C2") ' Z''
                
                ' Auto-fit columns for readability
                extSheet.Columns("A:C").AutoFit
            End If
            
        End If
    Next ws
    
    ' Restore screen updating
    Application.ScreenUpdating = True
    
    ' MsgBox "Data extraction completed.", vbInformation

End Sub



' --- Main Routine: Iterate through all "ext" sheets to analyze and aggregate results ---
Sub ProcessAllExtSheets()
    Dim ws As Worksheet
    Dim summarySheet As Worksheet
    Dim sName As String
    Dim colorIdx As Long
    Dim totalExtSheets As Long
    
    ' Keep ScreenUpdating enabled to reflect drawing progress
    Application.ScreenUpdating = True
    
    ' --- 1. Initialize the Summary Sheet ---
    sName = "Summary_Plots"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add a new summary sheet at the end
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    summarySheet.Name = sName
    
    ' Create the initial layout (Empty chart frames)
    Call ArrangeSummaryCharts(summarySheet)
    
    ' Count the number of target sheets
    totalExtSheets = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then totalExtSheets = totalExtSheets + 1
    Next ws
    
    If totalExtSheets = 0 Then
        MsgBox "No target 'ext' sheets found.", vbExclamation ' 
        Exit Sub
    End If

    ' --- 2. Analysis Loop ---
    colorIdx = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then
            
            ' Keep the summary sheet visible during analysis

            summarySheet.Activate
            
            ' Run DRT analysis on the target sheet

            ws.Activate
            On Error Resume Next
            Call ActiveSheetDRT_all
            On Error GoTo 0
            
            ' Update summary charts with the new data

            Call AddToMeasuredNyquist(ws, summarySheet, colorIdx)
            Call AddToCalcNyquist(ws, summarySheet, colorIdx)
            Call AddToDRTSpectrum(ws, summarySheet, colorIdx)
            
            ' Force refresh of the summary view

            summarySheet.Activate
            DoEvents
            
            colorIdx = colorIdx + 1
        End If
    Next ws
    
    ' --- 3. Final display adjustment ---

    summarySheet.Activate
    summarySheet.Range("A1").Select
    
    MsgBox "Analysis and summary creation completed for all sheets.", vbInformation

    
End Sub



' --- Function to return distinct colors for chart series ---

Function GetRGBColor(idx As Long) As Long
    Dim colors(0 To 19) As Long
    
    ' Define 20 distinct colors for scientific plotting

    colors(0) = RGB(255, 0, 0)      ' Red 
    colors(1) = RGB(0, 0, 255)      ' Blue 
    colors(2) = RGB(0, 128, 0)      ' Dark Green 
    colors(3) = RGB(255, 165, 0)    ' Orange 
    colors(4) = RGB(128, 0, 128)    ' Purple
    colors(5) = RGB(0, 255, 255)    ' Cyan 
    colors(6) = RGB(255, 20, 147)   ' Deep Pink 
    colors(7) = RGB(0, 100, 0)      ' Darker Green 
    colors(8) = RGB(139, 69, 19)    ' Saddle Brown 
    colors(9) = RGB(0, 0, 128)      ' Navy 
    colors(10) = RGB(255, 215, 0)   ' Gold 
    colors(11) = RGB(128, 128, 0)   ' Olive 
    colors(12) = RGB(255, 0, 255)   ' Magenta 
    colors(13) = RGB(75, 0, 130)    ' Indigo 
    colors(14) = RGB(0, 255, 0)     ' Lime Green
    colors(15) = RGB(165, 42, 42)   ' Brown 
    colors(16) = RGB(70, 130, 180)  ' Steel Blue 
    colors(17) = RGB(255, 127, 80)  ' Coral 
    colors(18) = RGB(47, 79, 79)    ' Dark Slate Gray
    colors(19) = RGB(0, 206, 209)   ' Turquoise
    
    ' Use Mod to cycle through the colors if idx exceeds 19

    GetRGBColor = colors(idx Mod 20)
End Function


' --- Subroutine to add measured Nyquist data to the summary chart ---

Sub AddToMeasuredNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Attempt to find the existing chart named "Chart_Measured"

    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Measured")
    On Error GoTo 0
    
    ' If the chart does not exist, create and initialize it
    If chtObj Is Nothing Then
        ' Position and size of the chart
        ' 
        Set chtObj = targetSheet.ChartObjects.Add(10, 10, 400, 350)
        chtObj.Name = "Chart_Measured"
        
        With chtObj.Chart
            .ChartType = xlXYScatter
            .HasTitle = True
            .ChartTitle.Text = "Measured Nyquist"
            
            ' X-Axis: Real Impedance (Z')
            ' 
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Z' / Ohm"
            
            ' Y-Axis: Negative Imaginary Impedance (-Z'')
            ' 
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "-Z'' / Ohm"
            .Axes(xlValue).ReversePlotOrder = True ' Standard EIS inversion 
            
            .HasLegend = True
        End With
    End If
    
    ' Add a new data series for the current worksheet
    ' 
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name
        .XValues = ws.Range("B2:B" & lastRow) ' Z' data
        .Values = ws.Range("C2:C" & lastRow)  ' Z'' data
        
        ' Set marker style and apply the distinct color
        ' 
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 4
        .Format.Fill.ForeColor.RGB = GetRGBColor(idx) ' Assign color 
        .Format.Line.Visible = msoFalse               ' Hide lines between points 
    End With
    
End Sub

' --- Subroutine to add Calculated Nyquist (Fit) data to the summary chart ---
' 
Sub AddToCalcNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    Dim xRange As Range, yRange As Range
    
    ' Extract rows where Column G (7th) is marked as "Used"
    ' 
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 7).Value) = "Used" Then
            If xRange Is Nothing Then
                Set xRange = ws.Cells(i, 9)  ' Column I (Z' Calc)
                Set yRange = ws.Cells(i, 10) ' Column J (-Z'' Calc)
            Else
                Set xRange = Union(xRange, ws.Cells(i, 9))
                Set yRange = Union(yRange, ws.Cells(i, 10))
            End If
        End If
    Next i
    
    ' Exit if no valid "Used" data is found
    ' 
    If xRange Is Nothing Then Exit Sub

    ' Attempt to find the existing chart "Chart_Calc"
    ' 
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Calc")
    On Error GoTo 0
    
    ' Create and initialize the chart if it doesn't exist
    ' 
    If chtObj Is Nothing Then
        Set chtObj = targetSheet.ChartObjects.Add(420, 10, 400, 350)
        chtObj.Name = "Chart_Calc"
        With chtObj.Chart
            .ChartType = xlXYScatterLinesNoMarkers
            .HasTitle = True: .ChartTitle.Text = "Calculated Nyquist (Fit)"
            .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Z' / Ohm"
            .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "-Z'' / Ohm"
            .Axes(xlValue).ReversePlotOrder = True: .HasLegend = True
        End With
    End If
    
    ' Add a new series with the distinct color
    ' 
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name
        .XValues = xRange
        .Values = yRange
        .Format.Line.ForeColor.RGB = GetRGBColor(idx)
        .Format.Line.Weight = 1.5
    End With
End Sub

' --- Subroutine to add DRT Spectrum data to the summary chart ---
' 
Sub AddToDRTSpectrum(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim j As Long, targetCol As Long, endRow As Long
    
    ' Identify the "Optimal" lambda column
    ' 
    targetCol = 0
    For j = 2 To ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        If ws.Cells(j, 11).Value = "Optimal" Then targetCol = 15 + (j - 1): Exit For
    Next j
    
    ' Exit if no optimal column is found
    ' 
    If targetCol = 0 Then Exit Sub
    
    ' Determine the data range for the frequency grid
    ' 
    For j = 2 To ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then endRow = j Else Exit For
    Next j
    If endRow > 10 Then endRow = endRow - 3
    
    ' Find or create the DRT Spectrum chart
    ' 
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_DRT")
    On Error GoTo 0
    
    If chtObj Is Nothing Then
        Set chtObj = targetSheet.ChartObjects.Add(830, 10, 450, 350)
        chtObj.Name = "Chart_DRT"
        With chtObj.Chart
            .ChartType = xlXYScatterLinesNoMarkers: .HasTitle = True: .ChartTitle.Text = "DRT Spectrum"
            .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Frequency (Hz)"
            .Axes(xlCategory).ScaleType = xlLogarithmic ' Standard Log scale for DRT 
            .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "g(tau) / Ohm"
            .HasLegend = True
        End With
    End If
    
    ' Add the DRT series
    ' 
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name
        .XValues = ws.Range(ws.Cells(2, 15), ws.Cells(endRow, 15))
        .Values = ws.Range(ws.Cells(2, targetCol), ws.Cells(endRow, targetCol))
        .Format.Line.ForeColor.RGB = GetRGBColor(idx)
        .Format.Line.Weight = 2
    End With
End Sub

' --- Layout function to organize all summary charts ---
' 
Sub ArrangeSummaryCharts(ws As Worksheet)
    Dim i As Long
    Dim names As Variant: names = Array("Chart_Measured", "Chart_Calc", "Chart_DRT")
    
    ' Loop through the three standard summary charts
    ' 
    For i = 0 To 2
        On Error Resume Next
        With ws.ChartObjects(names(i))
            ' Apply standard positioning and gridlines
            ' 
            .Left = i * 460 + 10: .Top = 10: .Width = 450: .Height = 350
            .Chart.Axes(xlCategory).HasMajorGridlines = True
            .Chart.Axes(xlValue).HasMajorGridlines = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
        On Error GoTo 0
    Next i
End Sub
