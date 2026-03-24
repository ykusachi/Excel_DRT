Attribute VB_Name = "Module1"
Option Explicit

'WindowsとMacで切り分けるプロシージャー
Function WINorMAC() As Variant
    Dim MyFiles As Variant
     ' Test for the operating system.
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' Is Windows.
        MyFiles = Select_File_Or_Files_Windows
    Else
        ' Is a Mac and will test if running Excel 2011 or higher.
        If Val(Application.Version) > 14 Then
            MyFiles = Select_File_Or_Files_Mac
        End If
    End If
    
    ' 選択したファイルを戻り値に設定する
    WINorMAC = MyFiles
End Function
    


' Windowsでファイル選択ダイアログを表示する
Function Select_File_Or_Files_Windows()
    Dim SaveDriveDir As String
    Dim MyPath As String
    Dim Fname As Variant
    Dim n As Long
    Dim FnameInLoop As String
    Dim mybook As Workbook

    ' Save the current directory.
    SaveDriveDir = CurDir

    ' Set the path to the folder that you want to open.
    MyPath = Application.DefaultFilePath

    ' Change drive/directory to MyPath.
    ChDrive MyPath
    ChDir MyPath

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

    ' 選択したファイルを戻り値に設定する
    Select_File_Or_Files_Windows = Fname
End Function



' Macでファイル選択ダイアログを表示する
    
    Function Select_File_Or_Files_Mac() As Variant
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim n As Long
    Dim Fname As String
    Dim mybook As Workbook

    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"

    ' 以下の行で、ファイルタイプを .z 拡張子のみに限定しました。
    ' "org.openxmlformats.spreadsheetml.sheet" の部分を {'z'} に置き換えています。
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFiles to (choose file of type " & _
             " {""z""} " & _
               "with prompt ""Please select a .z file or files"" default location alias """ & _
               MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
               "set applescript's text item delimiters to """" " & vbNewLine & _
               "return theFiles"

    MyFiles = MacScript(MyScript)
    On Error GoTo 0
        
    ' 選択したファイルを戻り値に設定する
    Select_File_Or_Files_Mac = MyFiles
End Function
    
    




' Windowsでファイル選択ダイアログを表示する
Function bIsBookOpen(ByRef szBookName As String) As Boolean
    ' Contributed by Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function

' ファイルのフルパスを受け取り、パスとファイル名に分割する関数
' 戻り値: Array(ディレクトリのパス, ファイル名)
Function GetPathInfo(ByVal FullPath As String) As Variant
    Dim PathSeparator As String
    Dim LastSeparatorPos As Long
    Dim DirPath As String
    Dim FileName As String


    Dim i As Long
    Dim CheckChars As Variant
    
    ' List of separaton characters キャラクターコードで判断
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
        ' ディレクトリのパス: 最後の区切り文字まで
        DirPath = Left(FullPath, LastSeparatorPos)
        ' ファイル名: 最後の区切り文字の次から最後まで
        FileName = Mid(FullPath, LastSeparatorPos + 1)
    Else
        ' 区切り文字が見つからない場合は、フルパス全体をファイル名と見なす
        DirPath = ""
        FileName = FullPath
    End If
    
    GetPathInfo = Array(DirPath, FileName)
End Function

Sub InsertTextCsvFiles()
    
    Dim targetSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim FilePath As String, FileName As String
    Dim NewSheet As Worksheet
    
    Set targetSheet = ActiveSheet
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "リストが見つからないか、データがありません。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For i = 2 To lastRow
        FileName = targetSheet.Cells(i, "C").Value
        FilePath = targetSheet.Cells(i, "B").Value & FileName
        
        If FilePath <> "" And FileName <> "" Then
            
            ' --- 1. シート名の生成ロジック ---
            Dim CleanSheetName As String
            ' 拡張子を削除
            If InStrRev(FileName, ".") > 0 Then
                CleanSheetName = Left(FileName, InStrRev(FileName, ".") - 1)
            Else
                CleanSheetName = FileName
            End If
            
            ' 禁止文字の置換 ( :  / ? * [ ] )
            Dim illegalChars As Variant, charItem As Variant
            illegalChars = Array(":", "", "/", "?", "*", "[", "]")
            For Each charItem In illegalChars
                CleanSheetName = Replace(CleanSheetName, charItem, "_")
            Next charItem
            
            ' 【修正】後ろから25文字を切り出す
            If Len(CleanSheetName) > 25 Then
                CleanSheetName = Right(CleanSheetName, 25)
            End If
            
            ' 重複チェックと「頭」への文字付与
            Dim FinalSheetName As String
            Dim suffixIdx As Long
            Dim charList As String
            ' A-Z, 0-9 の順で付与
            charList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            
            FinalSheetName = CleanSheetName
            suffixIdx = 0
            
            Do While SheetExists(FinalSheetName)
                suffixIdx = suffixIdx + 1
                ' 【修正】シート名の「頭」に1文字＋アンダースコアを追加
                If suffixIdx <= Len(charList) Then
                    FinalSheetName = Mid(charList, suffixIdx, 1) & "_" & CleanSheetName
                Else
                    ' 36パターンを超えた場合は数字2桁＋アンダースコア
                    FinalSheetName = Format(suffixIdx, "00") & "_" & CleanSheetName
                End If
                
                ' シート名制限31文字を超えないよう調整（念のため）
                If Len(FinalSheetName) > 31 Then
                    FinalSheetName = Left(FinalSheetName, 31)
                End If
                
                If suffixIdx > 99 Then Exit Do ' 無限ループ防止
            Loop
            ' --- ロジック終了 ---
            
            ' 新規シート作成
            Set NewSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            NewSheet.Name = FinalSheetName
            
            ' インポート処理（既存のQueryTableロジックを継承）
            On Error Resume Next
            With NewSheet.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=NewSheet.Range("A1"))
                .TextFilePlatform = 932
                .TextFileParseType = xlDelimited
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
    

    'MsgBox "インポート完了。シート名の頭に識別子を付与しました。", vbInformation
    
' --- ここから修正 ---
    
    ' 1. まず画面更新と警告を再開させる
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' 2. エラー無視を一度リセット
    On Error GoTo 0
    
    ' 3. ExtractZDataを呼び出す
    ' ※ExtractZDataの中でシート追加をしているため、そこでも最後に画面更新が True になっている必要があります
    Call ExtractZData
    
    ' 4. 最後に確実に Sheet1 をアクティブにする
    On Error Resume Next
    Dim wsList As Worksheet
    Set wsList = ThisWorkbook.Sheets("Top")
    
    If Not wsList Is Nothing Then
        wsList.Select ' ActivateよりSelectの方が確実にフォーカスが当たることがあります
        wsList.Cells(1, 1).Select ' ついでに左上にカーソルを置く
    Else
        MsgBox "Top が見つかりませんでした。名称を確認してください。"
    End If
    On Error GoTo 0

End Sub

Function SheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function





Sub SelectFiles()
    Dim openWb As Workbook
    Dim openFileName As Variant, fileVar As Variant
    Dim InfoArray As Variant
    Dim WriteRow As Long ' 情報を書き込む行のカウンター
    Dim targetSheet As Worksheet ' 書き込み対象シート

    ' Windows版かMac版かによって処理を分けて、ファイルを複数選択する
    openFileName = WINorMAC
    
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' Windows
        If IsEmpty(openFileName) Or openFileName(1) = False Then
            MsgBox "キャンセルされました"
            Exit Sub
        End If
    Else
        ' Mac
        If openFileName = "" Then
            MsgBox "キャンセルされました"
            Exit Sub
        Else
            ' 文字列をカンマで区分けして、配列に格納する
            openFileName = Split(openFileName, ",")
        End If
    End If
    
    ' --- ここからファイル情報書き込み処理 ---
    
    ' アクティブシートを書き込み対象とする
    Set targetSheet = ActiveSheet
    
    ' 書き込み開始行を設定 (例: ヘッダーの次の行、2行目から開始)
    WriteRow = 1 ' ヘッダーがなければ1から、あれば2から始める
    
    ' 選択したファイルを開く（元のコードのロジック）
    For Each fileVar In openFileName
        
        ' Macの場合のパス変換 (元のコードのロジック)
        If Application.OperatingSystem Like "*Mac*" Then
            ' MacScriptのパス形式から、Workbooks.Openが認識できる形式へ変換
            ' 元のコードではOpenするためにスラッシュ区切りにしていましたが、
            ' GetPathInfoを使う場合でもパスの区切り方を統一するため、ここでは処理を維持します。
            fileVar = Replace(Replace(fileVar, ":", "/"), "Macintosh HD", "")
        End If

        ' ファイルパス情報を取得
        ' InfoArray(0) = ディレクトリ, InfoArray(1) = ファイル名
        InfoArray = GetPathInfo(CStr(fileVar))
        
        ' --- シートに情報を記入 ---
        WriteRow = WriteRow + 1 ' 次の行へ移動

        ' A列: 連番
        targetSheet.Cells(WriteRow, 1).Value = WriteRow - 1
        
        ' B列: 選択したファイルのディレクトリ
        targetSheet.Cells(WriteRow, 2).Value = InfoArray(0)
        
        ' C列: 選択したファイルのファイル名
        targetSheet.Cells(WriteRow, 3).Value = InfoArray(1)
        
        ' 以下は元のコードの「選択したファイルを開く/閉じる」処理
        ' ファイル情報を書き込むだけなら不要なので、必要に応じてコメントアウトを解除してください
        ' On Error Resume Next ' ファイルを開く処理が失敗しても続行
        ' Workbooks.Open fileVar
        ' Set openWb = ActiveWorkbook
        
        ' ' 処理を書く
        
        ' Application.DisplayAlerts = False
        ' If Not openWb Is Nothing Then openWb.Close
        ' Application.DisplayAlerts = True
        ' Set openWb = Nothing
        ' On Error GoTo 0
        
    Next fileVar
    
    ' --- ファイル情報書き込み処理 終了 ---
    
'    MsgBox "ファイルの情報をアクティブシートに書き込みました。"

End Sub



Sub ExtractZData()
    Dim ws As Worksheet, extSheet As Worksheet
    Dim lastRow As Long, dataStartRow As Long, i As Long
    Dim targetName As String
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' 「ext」で終わるシートと元リストシート以外を処理
        If Not ws.Name Like "*ext" And ws.Name <> "Sheet1" Then
            
            dataStartRow = 0
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' 「End Comments」または「End Header」を含む行を検索
            For i = 1 To lastRow
                Dim currentText As String
                currentText = ws.Cells(i, 1).Text
                
                If currentText Like "*End Comments*" Or currentText Like "*End Header*" Then
                    dataStartRow = i + 1
                    Exit For
                End If
            Next i
            
            If dataStartRow > 0 And dataStartRow <= lastRow Then
                ' シート名を安全な長さに調整
                Dim safeBaseName As String
                safeBaseName = Left(ws.Name, 28)
                targetName = safeBaseName & "ext"
                
                ' 同名シートの削除
                On Error Resume Next
                Application.DisplayAlerts = False
                Sheets(targetName).Delete
                Application.DisplayAlerts = True
                On Error GoTo 0
                
                ' シートの追加
                Set extSheet = ThisWorkbook.Sheets.Add(After:=ws)
                extSheet.Name = targetName
                
                ' ヘッダー作成 (A-C列のみ)
                extSheet.Range("A1:C1").Value = Array("Freq(Hz)", "Z'", "Z''")
                
                ' データの転記
                Dim rowCount As Long
                rowCount = lastRow - dataStartRow + 1
                
                ws.Cells(dataStartRow, 1).Resize(rowCount, 1).Copy extSheet.Range("A2") ' Freq
                ws.Cells(dataStartRow, 5).Resize(rowCount, 1).Copy extSheet.Range("B2") ' Z'
                ws.Cells(dataStartRow, 6).Resize(rowCount, 1).Copy extSheet.Range("C2") ' Z''
            End If
            
        End If
    Next ws
    
    Application.ScreenUpdating = True
'    MsgBox "抽出（転記のみ）が完了しました。", vbInformation

End Sub

' 1. メインルーチン：全extシートを巡回して解析と集約を行う
Sub ProcessAllExtSheets()
    Dim ws As Worksheet
    Dim summarySheet As Worksheet
    Dim sName As String
    Dim colorIdx As Long
    Dim totalExtSheets As Long
    
    ' 画面更新を「停止させない」ことで描画を反映させる
    Application.ScreenUpdating = True
    
    ' 1. 集約用シートの初期化
    sName = "Summary_Plots"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    summarySheet.Name = sName
    
    ' 初期レイアウト（空の枠組み）を作成しておく
    Call ArrangeSummaryCharts(summarySheet)
    
    ' 対象シート数のカウント
    totalExtSheets = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then totalExtSheets = totalExtSheets + 1
    Next ws
    
    If totalExtSheets = 0 Then
        MsgBox "対象シートが見つかりません。"
        Exit Sub
    End If

    ' --- 解析ループ ---
    colorIdx = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*ext" Then
            
            ' 解析中も常にサマリーシートを表示しておく
            summarySheet.Activate
            
            ' 背景で対象シートの解析を実行
            ' (ActiveSheetに対して動くマクロの場合はws.Activateが必要なため、一瞬切り替わる場合があります)
            ws.Activate
            On Error Resume Next
            Call ActiveSheetDRT_all
            On Error GoTo 0
            
            ' グラフを更新
            Call AddToMeasuredNyquist(ws, summarySheet, colorIdx)
            Call AddToCalcNyquist(ws, summarySheet, colorIdx)
            Call AddToDRTSpectrum(ws, summarySheet, colorIdx)
            
            ' 描画を強制的に反映させるために一瞬サマリーを表示してDoEvents
            summarySheet.Activate
            DoEvents
            
            colorIdx = colorIdx + 1
        End If
    Next ws
    
    ' 3. 最後に改めてサマリーシートを表示
    summarySheet.Activate
    summarySheet.Range("A1").Select
    
    MsgBox "すべての解析とサマリー作成が完了しました。", vbInformation
End Sub

' --- 【強化版】明確に異なる色を返す関数 ---
Function GetRGBColor(idx As Long) As Long
    Dim colors(0 To 19) As Long
    
    colors(0) = RGB(255, 0, 0)      '赤
    colors(1) = RGB(0, 0, 255)      '青
    colors(2) = RGB(0, 128, 0)      '濃い緑
    colors(3) = RGB(255, 165, 0)    'オレンジ
    colors(4) = RGB(128, 0, 128)    '紫
    colors(5) = RGB(0, 255, 255)    'シアン
    colors(6) = RGB(255, 20, 147)   'ピンク
    colors(7) = RGB(0, 100, 0)      '深緑
    colors(8) = RGB(139, 69, 19)    '茶色
    colors(9) = RGB(0, 0, 128)      '紺
    colors(10) = RGB(255, 215, 0)   'ゴールド
    colors(11) = RGB(128, 128, 0)   'オリーブ
    colors(12) = RGB(255, 0, 255)   'マゼンタ
    colors(13) = RGB(75, 0, 130)    'インディゴ
    colors(14) = RGB(0, 255, 0)     '明るい緑
    colors(15) = RGB(165, 42, 42)   'ブラウン
    colors(16) = RGB(70, 130, 180)  'スチールブルー
    colors(17) = RGB(255, 127, 80)  'コーラル
    colors(18) = RGB(47, 79, 79)    'ダークスレートグレイ
    colors(19) = RGB(0, 206, 209)   'ターコイズ
    
    GetRGBColor = colors(idx Mod 20)
End Function

' --- 各グラフ作成プロシージャ（色取得部分のみ修正） ---

Sub AddToMeasuredNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Measured")
    On Error GoTo 0
    If chtObj Is Nothing Then
        Set chtObj = targetSheet.ChartObjects.Add(10, 10, 400, 350)
        chtObj.Name = "Chart_Measured"
        With chtObj.Chart
            .ChartType = xlXYScatter: .HasTitle = True: .ChartTitle.Text = "Measured Nyquist"
            .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Z' / Ohm"
            .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "-Z'' / Ohm"
            .Axes(xlValue).ReversePlotOrder = True: .HasLegend = True
        End With
    End If
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name: .XValues = ws.Range("B2:B" & lastRow): .Values = ws.Range("C2:C" & lastRow)
        .MarkerStyle = xlMarkerStyleCircle: .MarkerSize = 4
        .Format.Fill.ForeColor.RGB = GetRGBColor(idx) '色指定
        .Format.Line.Visible = msoFalse
    End With
End Sub

' --- 計算Nyquistグラフ (Usedのみ抽出) ---
Sub AddToCalcNyquist(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    Dim xRange As Range, yRange As Range
    
    ' G列(7番目)が "Used" の行だけを抽出
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 7).Value) = "Used" Then
            If xRange Is Nothing Then
                Set xRange = ws.Cells(i, 9)  ' I列 (Z' Calc)
                Set yRange = ws.Cells(i, 10) ' J列 (-Z'' Calc)
            Else
                Set xRange = Union(xRange, ws.Cells(i, 9))
                Set yRange = Union(yRange, ws.Cells(i, 10))
            End If
        End If
    Next i
    
    If xRange Is Nothing Then Exit Sub

    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_Calc")
    On Error GoTo 0
    
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
    
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name
        .XValues = xRange
        .Values = yRange
        .Format.Line.ForeColor.RGB = GetRGBColor(idx)
        .Format.Line.Weight = 1.5
    End With
End Sub

Sub AddToDRTSpectrum(ws As Worksheet, targetSheet As Worksheet, idx As Long)
    Dim chtObj As ChartObject: Dim ser As Series
    Dim j As Long, targetCol As Long, endRow As Long
    targetCol = 0
    For j = 2 To ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        If ws.Cells(j, 11).Value = "Optimal" Then targetCol = 15 + (j - 1): Exit For
    Next j
    If targetCol = 0 Then Exit Sub
    For j = 2 To ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then endRow = j Else Exit For
    Next j
    If endRow > 10 Then endRow = endRow - 3
    On Error Resume Next
    Set chtObj = targetSheet.ChartObjects("Chart_DRT")
    On Error GoTo 0
    If chtObj Is Nothing Then
        Set chtObj = targetSheet.ChartObjects.Add(830, 10, 450, 350)
        chtObj.Name = "Chart_DRT"
        With chtObj.Chart
            .ChartType = xlXYScatterLinesNoMarkers: .HasTitle = True: .ChartTitle.Text = "DRT Spectrum"
            .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Frequency (Hz)"
            .Axes(xlCategory).ScaleType = xlLogarithmic
            .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "g(tau) / Ohm"
            .HasLegend = True
        End With
    End If
    Set ser = chtObj.Chart.SeriesCollection.NewSeries
    With ser
        .Name = ws.Name: .XValues = ws.Range(ws.Cells(2, 15), ws.Cells(endRow, 15)): .Values = ws.Range(ws.Cells(2, targetCol), ws.Cells(endRow, targetCol))
        .Format.Line.ForeColor.RGB = GetRGBColor(idx) '色指定
        .Format.Line.Weight = 2
    End With
End Sub

' --- 配置関数（変更なし） ---
Sub ArrangeSummaryCharts(ws As Worksheet)
    Dim i As Long
    Dim names As Variant: names = Array("Chart_Measured", "Chart_Calc", "Chart_DRT")
    For i = 0 To 2
        On Error Resume Next
        With ws.ChartObjects(names(i))
            .Left = i * 460 + 10: .Top = 10: .Width = 450: .Height = 350
            .Chart.Axes(xlCategory).HasMajorGridlines = True
            .Chart.Axes(xlValue).HasMajorGridlines = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
        On Error GoTo 0
    Next i
End Sub

