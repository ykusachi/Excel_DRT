Attribute VB_Name = "Module1"
Option Explicit

' --- Function to branch file selection based on the Operating System ---
' --- Windows‚ÆMac‚Åƒtƒ@ƒCƒ‹‘I‘ğˆ—‚ğØ‚è•ª‚¯‚éƒvƒƒV[ƒWƒƒ ---
Function WINorMAC() As Variant
    Dim MyFiles As Variant
    
    ' Test for the operating system (Check if it's NOT a Mac)
    ' OS‚Ìí—Ş‚ğ”»’èiMac‚Å‚È‚¢ê‡‚ÍWindows‚Æ‚İ‚È‚·j
    If Not Application.OperatingSystem Like "*Mac*" Then
        
        ' Target: Windows OS
        ' ‘ÎÛ: Windows
        MyFiles = Select_File_Or_Files_Windows
        
    Else
        
        ' Target: Mac OS (Test if running Excel 2011/Version 14 or higher)
        ' ‘ÎÛ: MaciExcel 2011/Version 14 ˆÈã‚©‚ğŠm”Fj
        If Val(Application.Version) > 14 Then
            MyFiles = Select_File_Or_Files_Mac
        Else
            ' Error: Version not supported
            ' ƒGƒ‰[: ƒTƒ|[ƒg‚³‚ê‚Ä‚¢‚È‚¢ƒo[ƒWƒ‡ƒ“
            MsgBox "Error: This Mac Excel version is not supported.", vbCritical
            MyFiles = False
        End If
        
    End If
    
    ' Set the selected file(s) as the return value
    ' ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚ğ–ß‚è’l‚Éİ’è‚·‚é
    WINorMAC = MyFiles
    
End Function
    


' --- Function to display the file selection dialog on Windows ---
' --- Windows‚Åƒtƒ@ƒCƒ‹‘I‘ğƒ_ƒCƒAƒƒO‚ğ•\¦‚·‚éŠÖ” ---
Function Select_File_Or_Files_Windows()
    Dim SaveDriveDir As String
    Dim MyPath As String
    Dim Fname As Variant
    Dim n As Long
    Dim FnameInLoop As String
    Dim mybook As Workbook

    ' Save the current directory to restore later
    ' Œ»İ‚ÌƒfƒBƒŒƒNƒgƒŠ‚ğ•Û‘¶iŒã‚Å•œŒ³‚·‚é‚½‚ßj
    SaveDriveDir = CurDir

    ' Set the target path to the application default
    ' ŠJ‚«‚½‚¢ƒtƒHƒ‹ƒ_‚ÌƒpƒX‚ğƒfƒtƒHƒ‹ƒg‚Éİ’è
    MyPath = Application.DefaultFilePath

    ' Change current drive and directory to MyPath
    ' ƒhƒ‰ƒCƒu‚ÆƒfƒBƒŒƒNƒgƒŠ‚ğMyPath‚É•ÏX
    On Error Resume Next ' Avoid errors if the drive/path is invalid
    ChDrive MyPath
    ChDir MyPath
    On Error GoTo 0

    ' Open the file picker with Excel filters and a custom title
    ' Excelƒtƒ@ƒCƒ‹ƒtƒBƒ‹ƒ^‚ÆƒJƒXƒ^ƒ€ƒ^ƒCƒgƒ‹‚Åƒtƒ@ƒCƒ‹‘I‘ğ‚ğŠJ‚­
    Fname = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.xls*), *.xls*", _
            Title:="Select a file or files", _
            MultiSelect:=True)

    ' Restore the original drive and directory
    ' ƒhƒ‰ƒCƒu‚ÆƒfƒBƒŒƒNƒgƒŠ‚ğŒ³‚ÌƒfƒBƒŒƒNƒgƒŠiSaveDriveDirj‚É–ß‚·
    On Error Resume Next
    ChDrive SaveDriveDir
    ChDir SaveDriveDir
    On Error GoTo 0

    ' Return the selected file(s) (returns False if canceled)
    ' ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚ğ–ß‚è’l‚Éİ’èiƒLƒƒƒ“ƒZƒ‹‚ÍFalsej
    Select_File_Or_Files_Windows = Fname
    
End Function



' --- Function to display the file selection dialog on Mac using AppleScript ---
' --- AppleScript‚ğg—p‚µ‚ÄMac‚Åƒtƒ@ƒCƒ‹‘I‘ğƒ_ƒCƒAƒƒO‚ğ•\¦‚·‚éŠÖ” ---
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
    ' ƒfƒtƒHƒ‹ƒg‚ÌƒpƒX‚Æ‚µ‚ÄƒhƒLƒ…ƒƒ“ƒgƒtƒHƒ‹ƒ_‚ğæ“¾
    MyPath = MacScript("return (path to documents folder) as String")
    
    ' Construct AppleScript to select files with .z extension
    ' .z Šg’£q‚Ìƒtƒ@ƒCƒ‹‚ğ‘I‘ğ‚·‚é‚½‚ß‚Ì AppleScript ‚ğ\’z
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
                "set theFiles to (choose file of type " & _
              " {""z""} " & _
                "with prompt ""Please select a .z file or files"" default location alias """ & _
                MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"

    ' Execute the AppleScript
    ' AppleScript ‚ğÀs
    MyFiles = MacScript(MyScript)
    On Error GoTo 0
        
    ' Return the selected file(s) as the return value
    ' ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚ğ–ß‚è’l‚Éİ’è‚·‚é
    Select_File_Or_Files_Mac = MyFiles
    
End Function

    
    

' --- Function to check if a specific workbook is currently open ---
' --- w’è‚µ‚½ƒ[ƒNƒuƒbƒN‚ªŒ»İŠJ‚¢‚Ä‚¢‚é‚©Šm”F‚·‚éŠÖ” ---
Function bIsBookOpen(ByRef szBookName As String) As Boolean
    ' Contributed by Rob Bovey
    
    ' Disable error handling to check for existence
    ' ‘¶İŠm”F‚Ì‚½‚ßAƒGƒ‰[ƒnƒ“ƒhƒŠƒ“ƒO‚ğˆê“I‚É–³Œø‰»
    On Error Resume Next
    
    ' If the workbook is not found, the object will be Nothing
    ' ƒ[ƒNƒuƒbƒN‚ªŒ©‚Â‚©‚ç‚È‚¢ê‡AƒIƒuƒWƒFƒNƒg‚Í Nothing ‚É‚È‚é
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
    
    ' Reset error handling
    ' ƒGƒ‰[ƒnƒ“ƒhƒŠƒ“ƒO‚ğƒŠƒZƒbƒg
    On Error GoTo 0
End Function


' --- Function to split a full path into directory path and file name ---
' --- ƒtƒ@ƒCƒ‹‚Ìƒtƒ‹ƒpƒX‚ğƒfƒBƒŒƒNƒgƒŠƒpƒX‚Æƒtƒ@ƒCƒ‹–¼‚É•ªŠ„‚·‚éŠÖ” ---
' Returns: Array(Directory Path, File Name)
' –ß‚è’l: Array(ƒfƒBƒŒƒNƒgƒŠ‚ÌƒpƒX, ƒtƒ@ƒCƒ‹–¼)
Function GetPathInfo(ByVal FullPath As String) As Variant
    Dim PathSeparator As String
    Dim LastSeparatorPos As Long
    Dim DirPath As String
    Dim FileName As String
    
    ' Determine path separator based on the Operating System
    ' OS‚É‚æ‚Á‚ÄƒpƒX‹æØ‚è•¶š‚ğ”»’f
    If Application.OperatingSystem Like "*Mac*" Then
        ' For Mac: Prioritize "/" but also check for ":"
        ' Mac OS: "/" ‚ğ—Dæ‚µA•K—v‚É‰‚¶‚Ä ":" ‚àƒ`ƒFƒbƒN‚·‚é
        PathSeparator = IIf(InStrRev(FullPath, "/") > 0, "/", ":")
    Else
        ' For Windows: Always Use "€"
        ' Windows: í‚É "€" ‚ğg—p
        PathSeparator = "€"
    End If
    
    ' Find the position of the last separator
    ' ÅŒã‚Ì‹æØ‚è•¶š‚ÌˆÊ’u‚ğæ“¾
    LastSeparatorPos = InStrRev(FullPath, PathSeparator)
    
    If LastSeparatorPos > 0 Then
        ' Directory Path: Everything up to the last separator
        ' ƒfƒBƒŒƒNƒgƒŠ‚ÌƒpƒX: ÅŒã‚Ì‹æØ‚è•¶š‚Ü‚Å
        DirPath = Left(FullPath, LastSeparatorPos)
        
        ' File Name: Everything after the last separator
        ' ƒtƒ@ƒCƒ‹–¼: ÅŒã‚Ì‹æØ‚è•¶š‚ÌŸ‚©‚çÅŒã‚Ü‚Å
        FileName = Mid(FullPath, LastSeparatorPos + 1)
    Else
        ' If no separator is found, treat the whole path as the file name
        ' ‹æØ‚è•¶š‚ªŒ©‚Â‚©‚ç‚È‚¢ê‡‚ÍAƒtƒ‹ƒpƒX‘S‘Ì‚ğƒtƒ@ƒCƒ‹–¼‚ÆŒ©‚È‚·
        DirPath = ""
        FileName = FullPath
    End If
    
    ' Return as an array
    ' ”z—ñ‚Æ‚µ‚Ä–ß‚è’l‚ğİ’è
    GetPathInfo = Array(DirPath, FileName)
    
End Function





' --- Subroutine to import CSV/Text files based on a list ---
' --- ƒŠƒXƒg‚ÉŠî‚Ã‚¢‚ÄCSV/ƒeƒLƒXƒgƒtƒ@ƒCƒ‹‚ğƒCƒ“ƒ|[ƒg‚·‚éƒTƒuƒvƒƒV[ƒWƒƒ ---
Sub InsertTextCsvFiles()
    
    Dim targetSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim FilePath As String, FileName As String
    Dim NewSheet As Worksheet
    
    Set targetSheet = ActiveSheet
    ' Get the last row of the list in Column B
    ' B—ñ‚ÌƒŠƒXƒg‚ÌÅIs‚ğæ“¾
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Check if data exists in the list
    ' ƒŠƒXƒg‚Éƒf[ƒ^‚ª‘¶İ‚·‚é‚©Šm”F
    If lastRow < 2 Then
        MsgBox "List not found or contains no data.", vbExclamation
        Exit Sub
    End If
    
    ' Disable screen updates and alerts for performance
    ' ƒpƒtƒH[ƒ}ƒ“ƒXŒüã‚Ì‚½‚ß‰æ–ÊXV‚ÆŒx‚ğ’â~
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through the list of files
    ' ƒtƒ@ƒCƒ‹ƒŠƒXƒg‚ğƒ‹[ƒvˆ—
    For i = 2 To lastRow
        FileName = targetSheet.Cells(i, "C").Value
        FilePath = targetSheet.Cells(i, "B").Value & FileName
        
        If FilePath <> "" And FileName <> "" Then
            
            ' --- 1. Sheet name generation logic ---
            ' --- 1. ƒV[ƒg–¼‚Ì¶¬ƒƒWƒbƒN ---
            Dim CleanSheetName As String
            
            ' Remove file extension
            ' Šg’£q‚ğíœ
            If InStrRev(FileName, ".") > 0 Then
                CleanSheetName = Left(FileName, InStrRev(FileName, ".") - 1)
            Else
                CleanSheetName = FileName
            End If
            
            ' Replace illegal characters ( : € / ? * [ ] ) with underscore
            ' ‹Ö~•¶š ( : € / ? * [ ] ) ‚ğƒAƒ“ƒ_[ƒXƒRƒA‚É’uŠ·
            Dim illegalChars As Variant, charItem As Variant
            illegalChars = Array(":", "€", "/", "?", "*", "[", "]")
            For Each charItem In illegalChars
                CleanSheetName = Replace(CleanSheetName, charItem, "_")
            Next charItem
            
            ' Trim to the last 25 characters to stay within Excel's limits
            ' Excel‚Ì§ŒÀ“à‚Éû‚ß‚é‚½‚ßAŒã‚ë‚©‚ç25•¶š‚ğØ‚èo‚·
            If Len(CleanSheetName) > 25 Then
                CleanSheetName = Right(CleanSheetName, 25)
            End If
            
            ' Handle duplicate sheet names by adding a prefix
            ' d•¡ƒ`ƒFƒbƒN‚Æu“ªv‚Ö‚Ì•¶š•t—^‚É‚æ‚é–¼‘O‚ÌÕ“Ë‰ñ”ğ
            Dim FinalSheetName As String
            Dim suffixIdx As Long
            Dim charList As String
            ' Prefix sequence: A-Z, then 0-9
            ' •t—^‚·‚é•¶šƒŠƒXƒg: A-Z, 0-9 ‚Ì‡
            charList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            
            FinalSheetName = CleanSheetName
            suffixIdx = 0
            
            ' Loop until a unique sheet name is found
            ' ƒ†ƒj[ƒN‚ÈƒV[ƒg–¼‚ªŒ©‚Â‚©‚é‚Ü‚Åƒ‹[ƒv
            Do While SheetExists(FinalSheetName)
                suffixIdx = suffixIdx + 1
                
                ' Add prefix (1 char + underscore) to the start of the name
                ' ƒV[ƒg–¼‚Ìu“ªv‚É1•¶š{ƒAƒ“ƒ_[ƒXƒRƒA‚ğ’Ç‰Á
                If suffixIdx <= Len(charList) Then
                    FinalSheetName = Mid(charList, suffixIdx, 1) & "_" & CleanSheetName
                Else
                    ' Use 2-digit number if pattern exceeds 36
                    ' 36ƒpƒ^[ƒ“‚ğ’´‚¦‚½ê‡‚Í”š2Œ…{ƒAƒ“ƒ_[ƒXƒRƒA
                    FinalSheetName = Format(suffixIdx, "00") & "_" & CleanSheetName
                End If
                
                ' Ensure total length does not exceed 31 characters
                ' ƒV[ƒg–¼§ŒÀ31•¶š‚ğ’´‚¦‚È‚¢‚æ‚¤’²®
                If Len(FinalSheetName) > 31 Then
                    FinalSheetName = Left(FinalSheetName, 31)
                End If
                
                If suffixIdx > 99 Then Exit Do ' Prevent infinite loop
            Loop
            
            ' --- Create New Sheet ---
            ' --- V‹KƒV[ƒgì¬ ---
            Set NewSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            NewSheet.Name = FinalSheetName
            
            ' --- Import Data via QueryTable ---
            ' --- ƒCƒ“ƒ|[ƒgˆ—iQueryTablej ---
            On Error Resume Next
            With NewSheet.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=NewSheet.Range("A1"))
                .TextFilePlatform = 932 ' Shift-JIS
                .TextFileParseType = xlDelimited
                ' Check if CSV or Tab-delimited
                ' CSV‚©ƒ^ƒu‹æØ‚è‚©‚ğ”»’è
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
    ' --- I—¹Œã‚ÌŒãˆ— ---
    
    ' Restore screen updating and alerts
    ' ‰æ–ÊXV‚ÆŒx‚ğÄŠJ
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Reset error handling
    ' ƒGƒ‰[ƒnƒ“ƒhƒŠƒ“ƒO‚ğƒŠƒZƒbƒg
    On Error GoTo 0
    
    ' Call the data extraction procedure
    ' ƒf[ƒ^’ŠoƒvƒƒV[ƒWƒƒ‚ğŒÄ‚Ño‚·
    Call ExtractZData
    
    ' Return to the main "Top" sheet
    ' ÅŒã‚ÉƒƒCƒ“‚Ì "Top" ƒV[ƒg‚ğƒAƒNƒeƒBƒu‚É‚·‚é
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
' --- w’è‚µ‚½–¼‘O‚ÌƒV[ƒg‚ªƒ[ƒNƒuƒbƒN“à‚É‘¶İ‚·‚é‚©Šm”F‚·‚éŠÖ” ---
Function SheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    
    ' Disable error handling to attempt object assignment
    ' ƒIƒuƒWƒFƒNƒg‚ÌŠ„‚è“–‚Ä‚ğs‚·‚é‚½‚ßAƒGƒ‰[ƒnƒ“ƒhƒŠƒ“ƒO‚ğˆê“I‚É–³Œø‰»
    On Error Resume Next
    
    ' Try to set the worksheet object by name
    ' –¼‘O‚ğw’è‚µ‚Äƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒg‚Ìæ“¾‚ğ‚İ‚é
    Set ws = ThisWorkbook.Sheets(SheetName)
    
    ' Reset error handling
    ' ƒGƒ‰[ƒnƒ“ƒhƒŠƒ“ƒO‚ğƒŠƒZƒbƒg
    On Error GoTo 0
    
    ' If the object 'ws' is not Nothing, the sheet exists
    ' ƒIƒuƒWƒFƒNƒg 'ws' ‚ª Nothing ‚Å‚È‚¯‚ê‚ÎAƒV[ƒg‚Í‘¶İ‚·‚é
    SheetExists = Not ws Is Nothing
    
End Function



' --- Subroutine to select multiple files and write their paths to the active sheet ---
' --- •¡”‚Ìƒtƒ@ƒCƒ‹‚ğ‘I‘ğ‚µA‚»‚ÌƒpƒXî•ñ‚ğƒAƒNƒeƒBƒuƒV[ƒg‚É‘‚«‚ŞƒTƒuƒvƒƒV[ƒWƒƒ ---
Sub SelectFiles()
    Dim openWb As Workbook
    Dim openFileName As Variant, fileVar As Variant
    Dim InfoArray As Variant
    Dim WriteRow As Long ' Counter for writing info / î•ñ‚ğ‘‚«‚Şs‚ÌƒJƒEƒ“ƒ^[
    Dim targetSheet As Worksheet ' Target sheet for writing / ‘‚«‚İ‘ÎÛƒV[ƒg

    ' Branch process for Windows or Mac to select multiple files
    ' Windows”Å‚©Mac”Å‚©‚É‚æ‚Á‚Äˆ—‚ğ•ª‚¯‚ÄAƒtƒ@ƒCƒ‹‚ğ•¡”‘I‘ğ‚·‚é
    openFileName = WINorMAC
    
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' --- Case: Windows ---
        ' --- Windows‚Ìê‡ ---
        If IsEmpty(openFileName) Or openFileName(1) = False Then
            MsgBox "Action canceled by user." ' ƒLƒƒƒ“ƒZƒ‹‚³‚ê‚Ü‚µ‚½
            Exit Sub
        End If
    Else
        ' --- Case: Mac ---
        ' --- Mac‚Ìê‡ ---
        If openFileName = "" Then
            MsgBox "Action canceled by user." ' ƒLƒƒƒ“ƒZƒ‹‚³‚ê‚Ü‚µ‚½
            Exit Sub
        Else
            ' Split the string by commas and store into an array
            ' •¶š—ñ‚ğƒJƒ“ƒ}‚Å‹æ•ª‚¯‚µ‚ÄA”z—ñ‚ÉŠi”[‚·‚é
            openFileName = Split(openFileName, ",")
        End If
    End If
    
    ' --- Start: Writing file information ---
    ' --- ƒtƒ@ƒCƒ‹î•ñ‘‚«‚İˆ— ŠJn ---
    
    ' Set the active sheet as the destination
    ' ƒAƒNƒeƒBƒuƒV[ƒg‚ğ‘‚«‚İ‘ÎÛ‚Æ‚·‚é
    Set targetSheet = ActiveSheet
    
    ' Set the starting row (e.g., Row 2 if there is a header)
    ' ‘‚«‚İŠJns‚ğİ’è (—á: ƒwƒbƒ_[‚ª‚ ‚ê‚Î2s–Ú‚©‚çŠJn)
    WriteRow = 1
    
    ' Loop through each selected file
    ' ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚ğƒ‹[ƒvˆ—
    For Each fileVar In openFileName
        
        ' Path conversion for Mac environment
        ' Mac‚Ìê‡‚ÌƒpƒX•ÏŠ·
        If Application.OperatingSystem Like "*Mac*" Then
            ' Convert MacScript path format to a format recognizable by Workbooks.Open
            ' MacScript‚ÌƒpƒXŒ`®‚©‚çAWorkbooks.Open‚ª”F¯‚Å‚«‚éŒ`®‚Ö•ÏŠ·
            fileVar = Replace(Replace(fileVar, ":", "/"), "Macintosh HD", "")
        End If

        ' Retrieve file path information
        ' InfoArray(0) = Directory, InfoArray(1) = File Name
        ' ƒtƒ@ƒCƒ‹ƒpƒXî•ñ‚ğæ“¾ (0:ƒfƒBƒŒƒNƒgƒŠ, 1:ƒtƒ@ƒCƒ‹–¼)
        InfoArray = GetPathInfo(CStr(fileVar))
        
        ' --- Write information to the sheet ---
        ' --- ƒV[ƒg‚Éî•ñ‚ğ‹L“ü ---
        WriteRow = WriteRow + 1 ' Move to the next row / Ÿ‚Ìs‚ÖˆÚ“®

        ' Column A: Serial Number / ˜A”Ô
        targetSheet.Cells(WriteRow, 1).Value = WriteRow - 1
        
        ' Column B: Directory of the selected file / ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚ÌƒfƒBƒŒƒNƒgƒŠ
        targetSheet.Cells(WriteRow, 2).Value = InfoArray(0)
        
        ' Column C: File name of the selected file / ‘I‘ğ‚µ‚½ƒtƒ@ƒCƒ‹‚Ìƒtƒ@ƒCƒ‹–¼
        targetSheet.Cells(WriteRow, 3).Value = InfoArray(1)
        
        ' Note: Original code for opening/closing files is commented out
        ' ƒtƒ@ƒCƒ‹‚ğŠJ‚­/•Â‚¶‚éˆ—‚ª•K—v‚Èê‡‚ÍˆÈ‰º‚ÌƒRƒƒ“ƒgƒAƒEƒg‚ğ‰ğœ‚µ‚Ä‚­‚¾‚³‚¢
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
    ' --- ƒtƒ@ƒCƒ‹î•ñ‘‚«‚İˆ— I—¹ ---
    
    ' MsgBox "File information has been written to the active sheet."

End Sub



' --- Subroutine to extract frequency and impedance data from raw data sheets ---
' --- ¶ƒf[ƒ^ƒV[ƒg‚©‚çü”g”‚ÆƒCƒ“ƒs[ƒ_ƒ“ƒXƒf[ƒ^‚ğ’ŠoE“]‹L‚·‚éƒTƒuƒvƒƒV[ƒWƒƒ ---
Sub ExtractZData()
    Dim ws As Worksheet, extSheet As Worksheet
    Dim lastRow As Long, dataStartRow As Long, i As Long
    Dim targetName As String
    
    ' Disable screen updates for performance
    ' ƒpƒtƒH[ƒ}ƒ“ƒXŒüã‚Ì‚½‚ß‰æ–ÊXV‚ğ’â~
    Application.ScreenUpdating = False
    
    ' Loop through all worksheets in the workbook
    ' ƒ[ƒNƒuƒbƒN“à‚Ì‘SƒV[ƒg‚ğƒ‹[ƒvˆ—
    For Each ws In ThisWorkbook.Worksheets
        ' Process sheets except those already ending in "ext" or the source list sheet
        ' Šù‚É "ext" ‚ÅI‚í‚éƒV[ƒgA‚Ü‚½‚ÍŒ³ƒŠƒXƒgƒV[ƒgiSheet1‚âTopjˆÈŠO‚ğˆ—
        If Not ws.Name Like "*ext" And ws.Name <> "Sheet1" And ws.Name <> "Top" Then
            
            dataStartRow = 0
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Search for the row containing header termination markers
            ' ƒwƒbƒ_[‚ÌI—¹‚ğ¦‚·ƒ}[ƒJ[iEnd Comments / End Headerj‚ğŒŸõ
            For i = 1 To lastRow
                Dim currentText As String
                currentText = ws.Cells(i, 1).Text
                
                If currentText Like "*End Comments*" Or currentText Like "*End Header*" Then
                    ' Data starts on the next row
                    ' ƒf[ƒ^ŠJns‚Íƒ}[ƒJ[‚ÌŸ‚Ìs
                    dataStartRow = i + 1
                    Exit For
                End If
            Next i
            
            ' Proceed if data start row was found
            ' ƒf[ƒ^ŠJns‚ªŒ©‚Â‚©‚Á‚½ê‡‚Ì‚İ‘±s
            If dataStartRow > 0 And dataStartRow <= lastRow Then
                ' Adjust sheet name length to fit Excel's limit (31 chars)
                ' ƒV[ƒg–¼‚ğExcel‚Ì§ŒÀi31•¶šj‚Éû‚Ü‚é‚æ‚¤’²®
                Dim safeBaseName As String
                safeBaseName = Left(ws.Name, 28)
                targetName = safeBaseName & "ext"
                
                ' Delete existing sheet with the same name if it exists
                ' “¯–¼‚ÌŠù‘¶ƒV[ƒg‚ª‚ ‚éê‡‚Ííœ
                On Error Resume Next
                Application.DisplayAlerts = False
                Sheets(targetName).Delete
                Application.DisplayAlerts = True
                On Error GoTo 0
                
                ' Add a new sheet after the current source sheet
                ' Œ»İ‚ÌQÆŒ³ƒV[ƒg‚Ì’¼Œã‚ÉV‹KƒV[ƒg‚ğ’Ç‰Á
                Set extSheet = ThisWorkbook.Sheets.Add(After:=ws)
                extSheet.Name = targetName
                
                ' Create Headers (Columns A to C)
                ' ƒwƒbƒ_[ì¬ (A-C—ñ)
                extSheet.Range("A1:C1").Value = Array("Freq(Hz)", "Z'", "Z''")
                
                ' Transfer (Copy) data values
                ' ƒf[ƒ^‚Ì“]‹L
                Dim rowCount As Long
                rowCount = lastRow - dataStartRow + 1
                
                ' Copy Freq (Col 1), Z' (Col 5), and Z'' (Col 6)
                ' ü”g”(1—ñ–Ú)AÀ•”(5—ñ–Ú)A‹••”(6—ñ–Ú)‚ğƒRƒs[
                ws.Cells(dataStartRow, 1).Resize(rowCount, 1).Copy extSheet.Range("A2") ' Freq
                ws.Cells(dataStartRow, 5).Resize(rowCount, 1).Copy extSheet.Range("B2") ' Z'
                ws.Cells(dataStartRow, 6).Resize(rowCount, 1).Copy extSheet.Range("C2") ' Z''
                
                ' Auto-fit columns for readability
                ' “Ç‚İ‚â‚·‚³‚Ì‚½‚ß‚É—ñ•‚ğ©“®’²®
                extSheet.Columns("A:C").AutoFit
            End If
            
        End If
    Next ws
    
    ' Restore screen updating
    ' ‰æ–ÊXV‚ğÄŠJ
    Application.ScreenUpdating = True
    
    ' MsgBox "Data extraction completed.", vbInformation
    ' MsgBox "’ŠoE“]‹L‚ªŠ®—¹‚µ‚Ü‚µ‚½B", vbInformation

End Sub



' --- Main Routine: Iterate through all "ext" sheets to analyze and aggregate results ---
' --- 1. ƒƒCƒ“ƒ‹[ƒ`ƒ“F‘SextƒV[ƒg‚ğ„‰ñ‚µ‚Ä‰ğÍ‚ÆW–ñ‚ğs‚¤ ---
Sub ProcessAllExtSheets()
    Dim ws As Worksheet
    Dim summarySheet As Worksheet
    Dim sName As String
    Dim colorIdx As Long
    Dim totalExtSheets As Long
    
    ' Keep ScreenUpdating enabled to reflect drawing progress
    ' ‰æ–ÊXV‚ğu’â~‚³‚¹‚È‚¢v‚±‚Æ‚Å•`‰æ‚ğ”½‰f‚³‚¹‚é
    Application.ScreenUpdating = True
    
    ' --- 1. Initialize the Summary Sheet ---
    ' --- 1. W–ñ—pƒV[ƒg‚Ì‰Šú‰» ---
    sName = "Summary_Plots"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add a new summary sheet at the end
    ' ÅŒã‚ÉƒTƒ}ƒŠ[ƒV[ƒg‚ğV‹K’Ç‰Á
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    summarySheet.Name = sName
    
    ' Create the initial layout (Empty chart frames)
    ' ‰ŠúƒŒƒCƒAƒEƒgi‹ó‚Ì˜g‘g‚İj‚ğì¬
    Call ArrangeSummaryCharts(summarySheet)
    
    ' Count the number of target sheets
    ' ‘ÎÛƒV[ƒg”‚ÌƒJƒEƒ“ƒg
    totalExtSheets = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then totalExtSheets = totalExtSheets + 1
    Next ws
    
    If totalExtSheets = 0 Then
        MsgBox "No target 'ext' sheets found.", vbExclamation ' ‘ÎÛƒV[ƒg‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñB
        Exit Sub
    End If

    ' --- 2. Analysis Loop ---
    ' --- ‰ğÍƒ‹[ƒv ---
    colorIdx = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then
            
            ' Keep the summary sheet visible during analysis
            ' ‰ğÍ’†‚àí‚ÉƒTƒ}ƒŠ[ƒV[ƒg‚ğ•\¦‚µ‚Ä‚¨‚­
            summarySheet.Activate
            
            ' Run DRT analysis on the target sheet
            ' ”wŒi‚Å‘ÎÛƒV[ƒg‚Ì‰ğÍ‚ğÀs
            ws.Activate
            On Error Resume Next
            Call ActiveSheetDRT_all
            On Error GoTo 0
            
            ' Update summary charts with the new data
            ' ƒOƒ‰ƒt‚ğXV
            Call AddToMeasuredNyquist(ws, summarySheet, colorIdx)
            Call AddToCalcNyquist(ws, summarySheet, colorIdx)
            Call AddToDRTSpectrum(ws, summarySheet, colorIdx)
            
            ' Force refresh of the summary view
            ' •`‰æ‚ğ‹­§“I‚É”½‰f‚³‚¹‚é‚½‚ß‚ÉˆêuƒTƒ}ƒŠ[‚ğ•\¦‚µ‚ÄDoEvents‚ğÀs
            summarySheet.Activate
            DoEvents
            
            colorIdx = colorIdx + 1
        End If
    Next ws
    
    ' --- 3. Final display adjustment ---
    ' --- 3. ÅŒã‚É‰ü‚ß‚ÄƒTƒ}ƒŠ[ƒV[ƒg‚ğ•\¦ ---
    summarySheet.Activate
    summarySheet.Range("A1").Select
    
    MsgBox "Analysis and summary creation completed for all sheets.", vbInformation
    ' ‚·‚×‚Ä‚Ì‰ğÍ‚ÆƒTƒ}ƒŠ[ì¬‚ªŠ®—¹‚µ‚Ü‚µ‚½B
    
End Sub



' --- Function to return distinct colors for chart series ---
' --- ƒOƒ‰ƒt‚ÌƒVƒŠ[ƒY—p‚É–¾Šm‚ÉˆÙ‚È‚éF‚ğ•Ô‚·ŠÖ” ---
Function GetRGBColor(idx As Long) As Long
    Dim colors(0 To 19) As Long
    
    ' Define 20 distinct colors for scientific plotting
    ' ‰ÈŠw“I‚Èƒvƒƒbƒg—p‚É20í—Ş‚Ì–¾Šm‚ÉˆÙ‚È‚éF‚ğ’è‹`
    colors(0) = RGB(255, 0, 0)      ' Red / Ô
    colors(1) = RGB(0, 0, 255)      ' Blue / Â
    colors(2) = RGB(0, 128, 0)      ' Dark Green / ”Z‚¢—Î
    colors(3) = RGB(255, 165, 0)    ' Orange / ƒIƒŒƒ“ƒW
    colors(4) = RGB(128, 0, 128)    ' Purple / ‡
    colors(5) = RGB(0, 255, 255)    ' Cyan / ƒVƒAƒ“
    colors(6) = RGB(255, 20, 147)   ' Deep Pink / ƒsƒ“ƒN
    colors(7) = RGB(0, 100, 0)      ' Darker Green / [—Î
    colors(8) = RGB(139, 69, 19)    ' Saddle Brown / ’ƒF
    colors(9) = RGB(0, 0, 128)      ' Navy / ®
    colors(10) = RGB(255, 215, 0)   ' Gold / ƒS[ƒ‹ƒh
    colors(11) = RGB(128, 128, 0)   ' Olive / ƒIƒŠ[ƒu
    colors(12) = RGB(255, 0, 255)   ' Magenta / ƒ}ƒ[ƒ“ƒ^
    colors(13) = RGB(75, 0, 130)    ' Indigo / ƒCƒ“ƒfƒBƒS
    colors(14) = RGB(0, 255, 0)     ' Lime Green / –¾‚é‚¢—Î
    colors(15) = RGB(165, 42, 42)   ' Brown / ƒuƒ‰ƒEƒ“
    colors(16) = RGB(70, 130, 180)  ' Steel Blue / ƒXƒ`[ƒ‹ƒuƒ‹[
    colors(17) = RGB(255, 127, 80)  ' Coral / ƒR[ƒ‰ƒ‹
    colors(18) = RGB(47, 79, 79)    ' Dark Slate Gray / ƒ_[ƒNƒXƒŒ[ƒgƒOƒŒƒC
    colors(19) = RGB(0, 206, 209)   ' Turquoise / ƒ^[ƒRƒCƒY
    
    ' Use Mod to cycle through the colors if idx exceeds 19
    ' ƒCƒ“ƒfƒbƒNƒX‚ª19‚ğ’´‚¦‚½ê‡‚Í Mod ‚ğg—p‚µ‚ÄF‚ğƒ‹[ƒv‚³‚¹‚é
    GetRGBColor = colors(idx Mod 20)
End Function


' --- Subroutine to add measured Nyquist data to the summary chart ---
' --- ‘ª’è‚³‚ê‚½Nyquistƒf[ƒ^‚ğƒTƒ}ƒŠ[ƒOƒ‰ƒt‚É’Ç‰Á‚·‚éƒTƒuƒvƒƒV[ƒWƒƒ ---
Sub AddToMeasuredNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Attempt to find the existing chart named "Chart_Measured"
    ' "Chart_Measured" ‚Æ‚¢‚¤–¼‘O‚ÌŠù‘¶ƒOƒ‰ƒt‚Ìæ“¾‚ğs
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Measured")
    On Error GoTo 0
    
    ' If the chart does not exist, create and initialize it
    ' ƒOƒ‰ƒt‚ª‘¶İ‚µ‚È‚¢ê‡‚ÍAV‹Kì¬‚µ‚Ä‰Šúİ’è‚ğs‚¤
    If chtObj Is Nothing Then
        ' Position and size of the chart
        ' ƒOƒ‰ƒt‚Ì”z’u‚ÆƒTƒCƒY
        Set chtObj = targetSheet.ChartObjects.Add(10, 10, 400, 350)
        chtObj.Name = "Chart_Measured"
        
        With chtObj.Chart
            .ChartType = xlXYScatter
            .HasTitle = True
            .ChartTitle.Text = "Measured Nyquist"
            
            ' X-Axis: Real Impedance (Z')
            ' X²: ƒCƒ“ƒs[ƒ_ƒ“ƒXÀ•” (Z')
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Z' / Ohm"
            
            ' Y-Axis: Negative Imaginary Impedance (-Z'')
            ' Y²: •‰‚ÌƒCƒ“ƒs[ƒ_ƒ“ƒX‹••” (-Z'')
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "-Z'' / Ohm"
            .Axes(xlValue).ReversePlotOrder = True ' Standard EIS inversion / “d‹C‰»Šw‚ÌŠµK‚É]‚¢”½“]
            
            .HasLegend = True
        End With
    End If
    
    ' Add a new data series for the current worksheet
    ' Œ»İ‚ÌƒV[ƒg—p‚ÌV‚µ‚¢ƒf[ƒ^ƒVƒŠ[ƒY‚ğ’Ç‰Á
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name
        .XValues = ws.Range("B2:B" & lastRow) ' Z' data
        .Values = ws.Range("C2:C" & lastRow)  ' Z'' data
        
        ' Set marker style and apply the distinct color
        ' ƒ}[ƒJ[ƒXƒ^ƒCƒ‹‚ğİ’è‚µAˆêˆÓ‚ÌF‚ğ“K—p
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 4
        .Format.Fill.ForeColor.RGB = GetRGBColor(idx) ' Assign color / F‚ÌŠ„‚è“–‚Ä
        .Format.Line.Visible = msoFalse               ' Hide lines between points / “_ŠÔ‚Ìü‚Í”ñ•\¦
    End With
    
End Sub

' --- Subroutine to add Calculated Nyquist (Fit) data to the summary chart ---
' --- ŒvZ‚³‚ê‚½NyquistiƒtƒBƒbƒeƒBƒ“ƒOjƒf[ƒ^‚ğƒTƒ}ƒŠ[ƒOƒ‰ƒt‚É’Ç‰Á‚·‚é ---
Sub AddToCalcNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    Dim xRange As Range, yRange As Range
    
    ' Extract rows where Column G (7th) is marked as "Used"
    ' G—ñi7”Ô–Új‚ª "Used" ‚Ìs‚¾‚¯‚ğ’Šo‚µ‚ÄƒŒƒ“ƒW‚ÉŠi”[
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
    ' —LŒø‚Èƒf[ƒ^‚ªŒ©‚Â‚©‚ç‚È‚¢ê‡‚ÍI—¹
    If xRange Is Nothing Then Exit Sub

    ' Attempt to find the existing chart "Chart_Calc"
    ' Šù‘¶‚Ì "Chart_Calc" ƒOƒ‰ƒt‚Ìæ“¾‚ğs
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Calc")
    On Error GoTo 0
    
    ' Create and initialize the chart if it doesn't exist
    ' ƒOƒ‰ƒt‚ª‘¶İ‚µ‚È‚¢ê‡‚ÍV‹Kì¬‚µ‚Ä‰Šú‰»
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
    ' ˆêˆÓ‚ÌF‚ğg—p‚µ‚ÄV‚µ‚¢ƒVƒŠ[ƒY‚ğ’Ç‰Á
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
' --- DRTƒXƒyƒNƒgƒ‹ƒf[ƒ^‚ğƒTƒ}ƒŠ[ƒOƒ‰ƒt‚É’Ç‰Á‚·‚é ---
Sub AddToDRTSpectrum(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim j As Long, targetCol As Long, endRow As Long
    
    ' Identify the "Optimal" lambda column
    ' uOptimalv‚Æ”»’è‚³‚ê‚½ƒ‰ƒ€ƒ_‚Ì—ñ‚ğ“Á’è
    targetCol = 0
    For j = 2 To ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        If ws.Cells(j, 11).Value = "Optimal" Then targetCol = 15 + (j - 1): Exit For
    Next j
    
    ' Exit if no optimal column is found
    ' Å“K‚È—ñ‚ªŒ©‚Â‚©‚ç‚È‚¢ê‡‚ÍI—¹
    If targetCol = 0 Then Exit Sub
    
    ' Determine the data range for the frequency grid
    ' ü”g”ƒOƒŠƒbƒh‚Ìƒf[ƒ^”ÍˆÍ‚ğŠm’è
    For j = 2 To ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then endRow = j Else Exit For
    Next j
    If endRow > 10 Then endRow = endRow - 3
    
    ' Find or create the DRT Spectrum chart
    ' DRTƒXƒyƒNƒgƒ‹ƒOƒ‰ƒt‚ğæ“¾‚Ü‚½‚ÍV‹Kì¬
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_DRT")
    On Error GoTo 0
    
    If chtObj Is Nothing Then
        Set chtObj = targetSheet.ChartObjects.Add(830, 10, 450, 350)
        chtObj.Name = "Chart_DRT"
        With chtObj.Chart
            .ChartType = xlXYScatterLinesNoMarkers: .HasTitle = True: .ChartTitle.Text = "DRT Spectrum"
            .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Frequency (Hz)"
            .Axes(xlCategory).ScaleType = xlLogarithmic ' Standard Log scale for DRT / DRT‚Ì•W€“I‚È‘Î”²
            .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "g(tau) / Ohm"
            .HasLegend = True
        End With
    End If
    
    ' Add the DRT series
    ' DRTƒVƒŠ[ƒY‚ğ’Ç‰Á
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
' --- ‘SƒTƒ}ƒŠ[ƒOƒ‰ƒt‚ğ®—ñ‚³‚¹‚é”z’uŠÖ” ---
Sub ArrangeSummaryCharts(ws As Worksheet)
    Dim i As Long
    Dim names As Variant: names = Array("Chart_Measured", "Chart_Calc", "Chart_DRT")
    
    ' Loop through the three standard summary charts
    ' 3‚Â‚Ì•W€ƒTƒ}ƒŠ[ƒOƒ‰ƒt‚ğ‡‚Éˆ—
    For i = 0 To 2
        On Error Resume Next
        With ws.ChartObjects(names(i))
            ' Apply standard positioning and gridlines
            ' •W€“I‚È”z’u‚Æ–Ú·ü‚Ì“K—p
            .Left = i * 460 + 10: .Top = 10: .Width = 450: .Height = 350
            .Chart.Axes(xlCategory).HasMajorGridlines = True
            .Chart.Axes(xlValue).HasMajorGridlines = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
        On Error GoTo 0
    Next i
End Sub
