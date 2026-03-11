Attribute VB_Name = "Module2"
Option Explicit
' ==========================================
' グローバル設定（ここを変更して調整してください）
' ==========================================
' 1. KKフィルターの閾値 (%)
Public Const KK_THRESHOLD As Double = 3#

' 2. λスキャンの設定
Public Const LAMBDA_START_EXP As Double = 0#   ' 10^-1 から開始
Public Const LAMBDA_END_EXP As Double = 10#    ' 10^-13 まで
Public Const LAMBDA_STEP As Double = 0.2        ' ステップ幅

' 3. DRTスペクトルの端点カット数
Public Const CUT_LOW_FREQ As Integer = 3       ' 低周波側のカット点数3
Public Const CUT_HIGH_FREQ As Integer = 0      ' 高周波側のカット点数 (0でカットなし)

Sub CalculateMagAndPhase()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    
    Set ws = ActiveSheet
    
    ' B列(Z')を基準に最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' データが2行目以降に存在するかチェック
    If lastRow < 2 Then
        MsgBox "計算対象のデータが見つかりません（2行目以降にデータが必要です）。"
        Exit Sub
    End If
    
    rowCount = lastRow - 1
    
    ' ヘッダーの入力（D1, E1）
    ws.Cells(1, 4).Value = "Magnitude"
    ws.Cells(1, 5).Value = "Phase"
    
    ' --- Magnitude の計算 (D列) ---
    ' 式: SQRT(Z'^2 + Z''^2)
    With ws.Range("D2:D" & lastRow)
        .FormulaR1C1 = "=SQRT(RC[-2]^2 + RC[-1]^2)"
    End With
    
    ' --- Phase の計算 (E列) ---
    ' 式: ATAN2(Z', Z'') * 180 / PI
    With ws.Range("E2:E" & lastRow)
        .FormulaR1C1 = "=ATAN2(RC[-3], RC[-2]) * 180 / PI()"
    End With
    
    ' 計算結果を値として確定（数式を消して軽量化）
    With ws.Range("D2:E" & lastRow)
        .Value = .Value
    End With
    
    'MsgBox "Activeシートの Magnitude と Phase を計算しました。", vbInformation
End Sub


' ==========================================
' 1. メインルーチン (R_infinity 統合・アーティファクト抑制版)
' ==========================================
Sub RunDRT()
    Dim ws As Worksheet
    Dim lastRow As Long, nPoints As Long, nValid As Long, nTau As Long
    Dim i As Long, j As Long, k As Long
    Dim freq() As Double, zReal() As Double, zImag() As Double
    Dim vFreq() As Double, vReal() As Double, vImag() As Double, vOmega() As Double
    Dim isValid() As Boolean
    Dim A() As Double, b() As Double, lambda As Double, tauGrid() As Double
    Dim PI As Double: PI = 3.14159265358979
    
    ' NNLS反復制御用
    Dim outerIter As Long, innerIter As Long
    Const MAX_OUTER As Long = 500
    Const MAX_INNER As Long = 200
    
    Set ws = ActiveSheet
    
    ' データの最終行取得
    On Error Resume Next
    lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    nPoints = lastRow - 1
    
    If nPoints < 10 Then
        MsgBox "データ点数が不足しています。"
        Exit Sub
    End If
    
    Application.StatusBar = "データ検証中 (KK Fit)..."
    
    ' --- 1. データ読み込み ---
    ReDim freq(1 To nPoints), zReal(1 To nPoints), zImag(1 To nPoints), isValid(1 To nPoints)
    For i = 1 To nPoints
        freq(i) = ws.Cells(i + 1, 1).Value
        zReal(i) = ws.Cells(i + 1, 2).Value
        zImag(i) = ws.Cells(i + 1, 3).Value
    Next i
    
    ' --- 2. KKフィルタ実行 ---
    Dim wk As Variant
    wk = PerformKKFit(nPoints, freq, zReal, zImag)
    
    nValid = 0
    ws.Cells(1, 6).Value = "KK_Res(%)"
    ws.Cells(1, 7).Value = "Status"
    
    For i = 1 To nPoints
        Dim zImagFit As Double: zImagFit = 0
        Dim omega_i As Double: omega_i = 2 * PI * freq(i)
        
        For j = 1 To nPoints
            Dim f_j_tmp As Double
            If freq(j) <= 0 Then
                f_j_tmp = 1E-10
            Else
                f_j_tmp = freq(j)
            End If
            
            Dim tau_j As Double: tau_j = 1 / (2 * PI * f_j_tmp)
            zImagFit = zImagFit - wk(j, 1) * (omega_i * tau_j) / (1 + (omega_i * tau_j) ^ 2)
        Next j
        
        Dim mag As Double: mag = Sqr(zReal(i) ^ 2 + zImag(i) ^ 2)
        Dim resPerc As Double: resPerc = (Abs(zImag(i) - zImagFit) / (mag + 1E-10)) * 100
        ws.Cells(i + 1, 6).Value = resPerc
        
        If resPerc <= KK_THRESHOLD Then ' 定数を参照
'        If resPerc <= 5 Then
            isValid(i) = True
            nValid = nValid + 1
            ws.Cells(i + 1, 7).Value = "Used"
            ws.Cells(i + 1, 7).Interior.Color = RGB(200, 255, 200)
        Else
            isValid(i) = False
            ws.Cells(i + 1, 7).Value = "Excluded(KK)"
            ws.Cells(i + 1, 7).Interior.Color = RGB(255, 200, 200)
        End If
    Next i
    
    If nValid < 5 Then
        Application.StatusBar = False
        MsgBox "有効データ不足です。"
        Exit Sub
    End If

    ' --- 3. 行列再構築 (Usedのみ) ---
    ReDim vFreq(1 To nValid), vReal(1 To nValid), vImag(1 To nValid), vOmega(1 To nValid)
    Dim activeIdx As Long: activeIdx = 0
    For i = 1 To nPoints
        If isValid(i) = True Then
            activeIdx = activeIdx + 1
            vFreq(activeIdx) = freq(i)
            vReal(activeIdx) = zReal(i)
            vImag(activeIdx) = zImag(i)
            vOmega(activeIdx) = 2 * PI * vFreq(activeIdx)
        End If
    Next i

    ' --- 4. 時定数グリッド作成 (マージンなしでUsed範囲に一致) ---
    nTau = 100
    ReDim tauGrid(1 To nTau)
    Dim minF As Double: minF = GetMin(vFreq)
    Dim maxF As Double: maxF = GetMax(vFreq)
    For j = 1 To nTau
        tauGrid(j) = (1 / (2 * PI * maxF)) * (maxF / minF) ^ ((j - 1) / (nTau - 1))
    Next j

    ' --- 5. NNLS用行列 A, b の構築 (nT+1 列目に R_inf を追加) ---
    ReDim A(1 To 2 * nValid, 1 To nTau + 1), b(1 To 2 * nValid, 1 To 1)
    For i = 1 To nValid
        b(i, 1) = vReal(i)
        b(i + nValid, 1) = -vImag(i)
        For j = 1 To nTau
            Dim wt As Double: wt = vOmega(i) * tauGrid(j)
            A(i, j) = 1 / (1 + wt ^ 2)           ' 実部
            A(i + nValid, j) = wt / (1 + wt ^ 2)  ' 虚部
        Next j
        ' 最後の列: R_infinity成分 (周波数に依存しないオフセット)
        A(i, nTau + 1) = 1
        A(i + nValid, nTau + 1) = 0
    Next i
    
    Dim matAtA As Variant: matAtA = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(A), A)
    Dim matAtb As Variant: matAtb = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(A), b)

    ' --- 6. ヘッダー設定 ---
    ws.Cells(1, 11).Value = "Flag"
    ws.Cells(1, 12).Value = "lambda"
    ws.Cells(1, 13).Value = "Log(ResSum)"
    ws.Cells(1, 14).Value = "Log(SolSum)"
    ws.Cells(1, 15).Value = "Freq_Grid(Hz)"
    
    ' --- 7. λスキャン実行 ---
    ' --- RunDRT 内の λスキャン実行部分 ---
    Dim nSteps As Long
    nSteps = Int((LAMBDA_END_EXP - LAMBDA_START_EXP) / LAMBDA_STEP) + 1
    
    For k = 1 To nSteps
        Dim exponent As Double: exponent = LAMBDA_START_EXP + (k - 1) * LAMBDA_STEP
        lambda = 10 ^ (-exponent)
        
        Application.StatusBar = "[" & ActiveSheet.Name & "] Analysing: λ=10^-" & _
                                Format(exponent, "0.00") & " (" & k & "/" & nSteps & ")"
        
    
'    For k = 1 To 49
'        Dim exponent As Double: exponent = 1# + (k - 1) * 0.25
'        lambda = 10 ^ (-exponent)
'
''        Application.StatusBar = "Analysing: λ=10^-" & Format(exponent, "0.00") & " (" & k & "/49)"
'        ' シート名を先頭に追加
'        Application.StatusBar = "[" & ActiveSheet.Name & "] Analysing: λ=10^-" & _
'                                Format(exponent, "0.00") & " (" & k & "/49)"
        
        DoEvents
        
        Dim x() As Double: ReDim x(1 To nTau + 1, 1 To 1)
        Dim P() As Boolean: ReDim P(1 To nTau + 1)
        Dim w() As Double: ReDim w(1 To nTau + 1)
        Dim isConverged As Boolean: isConverged = False
        
        For outerIter = 1 To MAX_OUTER
            Dim Ax() As Double: Ax = GetAx_Product(matAtA, x, nTau + 1, lambda)
            
            For i = 1 To nTau + 1
                w(i) = matAtb(i, 1) - Ax(i, 1)
            Next i
            
            Dim maxW As Double: maxW = -1E+30
            Dim t As Long: t = 0
            For i = 1 To nTau + 1
                If P(i) = False Then
                    If w(i) > maxW Then
                        maxW = w(i)
                        t = i
                    End If
                End If
            Next i
            
            If maxW <= 1E-09 Then
                isConverged = True
                Exit For
            End If
            
            P(t) = True
            
            For innerIter = 1 To MAX_INNER
                Dim z As Variant: z = SolveReducedSystem(matAtA, matAtb, P, nTau + 1, lambda)
                If IsEmpty(z) Then
                    Exit For
                End If
                
                Dim allPos As Boolean: allPos = True
                Dim minAlpha As Double: minAlpha = 2
                Dim q As Long: q = 0
                
                For i = 1 To nTau + 1
                    If P(i) = True Then
                        If z(i, 1) < -1E-12 Then
                            allPos = False
                            Dim alpha As Double: alpha = x(i, 1) / (x(i, 1) - z(i, 1))
                            If alpha < minAlpha Then
                                minAlpha = alpha
                                q = i
                            End If
                        End If
                    End If
                Next i
                
                If allPos = True Then
                    For i = 1 To nTau + 1
                        x(i, 1) = z(i, 1)
                    Next i
                    Exit For
                Else
                    For i = 1 To nTau + 1
                        x(i, 1) = x(i, 1) + minAlpha * (z(i, 1) - x(i, 1))
                    Next i
                    P(q) = False
                End If
            Next innerIter
        Next outerIter
        
        Dim failStr As String
        If isConverged = True Then
            failStr = ""
        Else
            failStr = " (Fail)"
        End If
        
        Call WriteSpectralResultShifted(ws, k, lambda, exponent, A, x, b, nValid, nTau, tauGrid, failStr)
    Next k
    
    Application.StatusBar = False
'    MsgBox "解析完了。"
End Sub

' ==========================================
' 2. 補助関数
' ==========================================

' 行列求解 (R_inf を正則化から除外)
Function SolveReducedSystem(AtA, Atb, P() As Boolean, n As Long, lam As Double) As Variant
    Dim tempA() As Double: ReDim tempA(1 To n, 1 To n)
    Dim tempB() As Double: ReDim tempB(1 To n, 1 To 1)
    Dim i As Long, j As Long
    Dim stab As Double: stab = 1E-11
    
    For i = 1 To n
        If P(i) = True Then
            tempB(i, 1) = Atb(i, 1)
            For j = 1 To n
                If P(j) = True Then
                    tempA(i, j) = AtA(i, j)
                    If i = j Then
                        If i < n Then
                            tempA(i, j) = tempA(i, j) + lam + stab
                        Else
                            tempA(i, j) = tempA(i, j) + stab
                        End If
                    End If
                Else
                    tempA(i, j) = 0
                End If
            Next j
        Else
            tempA(i, i) = 1
            tempB(i, 1) = 0
        End If
    Next i
    
    On Error Resume Next
    SolveReducedSystem = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(tempA), tempB)
    If Err.Number <> 0 Then
        SolveReducedSystem = Empty
    End If
    On Error GoTo 0
End Function

' 行列ベクトル積
Function GetAx_Product(AtA, x, n, lam) As Double()
    Dim res() As Double: ReDim res(1 To n, 1 To 1)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To n
            res(i, 1) = res(i, 1) + AtA(i, j) * x(j, 1)
        Next j
        If i < n Then
            res(i, 1) = res(i, 1) + lam * x(i, 1)
        End If
    Next i
    GetAx_Product = res
End Function

' 結果書き出し
Sub WriteSpectralResultShifted(ws, k, lam, expo, A, x, b, nV, nT, tG, msg As String)
    Dim Ag As Variant: Ag = Application.WorksheetFunction.MMult(A, x)
    Dim rS As Double: rS = 0
    Dim sS As Double: sS = 0
    Dim i As Long
    
    For i = 1 To 2 * nV
        rS = rS + (Ag(i, 1) - b(i, 1)) ^ 2
    Next i
    For i = 1 To nT
        sS = sS + x(i, 1) ^ 2
    Next i
    
    ws.Cells(k + 1, 12).Value = lam
    ws.Cells(k + 1, 13).Value = WorksheetFunction.Log10(rS + 1E-20)
    ws.Cells(k + 1, 14).Value = WorksheetFunction.Log10(sS + 1E-20)
    ws.Cells(1, 15 + k).Value = "λ:10^-" & Format(expo, "0.00") & msg
    
    For i = 1 To nT
        If k = 1 Then
            ws.Cells(i + 1, 15).Value = 1 / (2 * 3.14159265358979 * tG(i))
        End If
        ws.Cells(i + 1, 15 + k).Value = x(i, 1)
    Next i
    
    ' R_inf の出力
    If k = 1 Then
        ws.Cells(nT + 2, 15).Value = "R_inf(Ohm)"
    End If
    ws.Cells(nT + 2, 15 + k).Value = x(nT + 1, 1)
End Sub

' KK適合度計算
Function PerformKKFit(n As Long, f() As Double, zR() As Double, zI() As Double) As Variant
    Dim i As Long, j As Long, Ak() As Double, Bk() As Double
    Dim PI As Double: PI = 3.14159265358979
    ReDim Ak(1 To n, 1 To n), Bk(1 To n, 1 To 1)
    For i = 1 To n
        Bk(i, 1) = zR(i)
        For j = 1 To n
            Dim f_j As Double
            If f(j) <= 0 Then
                f_j = 1E-10
            Else
                f_j = f(j)
            End If
            Dim om As Double: om = 2 * PI * f(i)
            Dim ta As Double: ta = 1 / (2 * PI * f_j)
            Ak(i, j) = 1 / (1 + (om * ta) ^ 2)
        Next j
    Next i
    Dim AtA As Variant: AtA = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(Ak), Ak)
    For i = 1 To n
        AtA(i, i) = AtA(i, i) + 1E-06
    Next i
    On Error Resume Next
    PerformKKFit = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(AtA), _
                  Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(Ak), Bk))
    On Error GoTo 0
End Function

Function GetMax(arr() As Double) As Double
    Dim i As Long, m As Double: m = -1E+30
    For i = LBound(arr) To UBound(arr)
        If arr(i) > m Then
            m = arr(i)
        End If
    Next i
    GetMax = m
End Function

Function GetMin(arr() As Double) As Double
    Dim i As Long, m As Double: m = 1E+30
    For i = LBound(arr) To UBound(arr)
        If arr(i) < m Then
            m = arr(i)
        End If
    Next i
    GetMin = m
End Function


Sub FindOptimalLambda_Normalized()

Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim logRes() As Double, logSol() As Double
    Dim normRes() As Double, normSol() As Double
    Dim minRes As Double, maxRes As Double
    Dim minSol As Double, maxSol As Double
    Dim minDistance As Double, optIdx As Long
    Dim targetCol As Long
    
    Set ws = ActiveSheet
    
    ' --- 1. データの範囲取得 ---
    ' L列(12): lambdaのリスト, M列(13): LogRes, N列(14): LogSol
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "L-Curveデータが見つかりません。"
        Exit Sub
    End If
    
    ' --- 2. 既存の色のリセット ---
    ' K列(Flag), L列(lambda), および P列(16)以降の色をすべて消す
    ws.Columns(11).ClearContents
    ws.Columns(11).Interior.ColorIndex = xlNone
    ws.Columns(12).Interior.ColorIndex = xlNone
    If lastCol >= 16 Then
        ws.Range(ws.Columns(16), ws.Columns(lastCol)).Interior.ColorIndex = xlNone
    End If
    
    ws.Cells(1, 11).Value = "Flag"
    
    ' --- 3. データの読み込み ---
    ReDim logRes(2 To lastRow), logSol(2 To lastRow)
    minRes = 1E+30: maxRes = -1E+30
    minSol = 1E+30: maxSol = -1E+30
    
    For i = 2 To lastRow
        logRes(i) = ws.Cells(i, 13).Value ' M列: Log(ResSum)
        logSol(i) = ws.Cells(i, 14).Value ' N列: Log(SolSum)
        
        If logRes(i) < minRes Then minRes = logRes(i)
        If logRes(i) > maxRes Then maxRes = logRes(i)
        If logSol(i) < minSol Then minSol = logSol(i)
        If logSol(i) > maxSol Then maxSol = logSol(i)
    Next i
    
    ' --- 4. 正規化 [0, 1] ---
    ReDim normRes(2 To lastRow), normSol(2 To lastRow)
    For i = 2 To lastRow
        normRes(i) = (logRes(i) - minRes) / (maxRes - minRes + 1E-20)
        normSol(i) = (logSol(i) - minSol) / (maxSol - minSol + 1E-20)
    Next i
    
    ' --- 5. 最小距離法（原点からの最短距離） ---
    minDistance = 1E+30
    optIdx = 2 ' 初期値(2行目)
    
    For i = 2 To lastRow
        Dim dist As Double
        dist = Sqr(normRes(i) ^ 2 + normSol(i) ^ 2)
        If dist < minDistance Then
            minDistance = dist
            optIdx = i ' ここで「何番目のλか」が確定する
        End If
    Next i
    
    ' --- 6. 結果の着色 ---
    ' 1. L-Curveリスト側のOptimal行（K, L列）を着色
    ws.Cells(optIdx, 11).Value = "Optimal"
    ws.Cells(optIdx, 11).Interior.Color = vbYellow
    ws.Cells(optIdx, 12).Interior.Color = vbYellow
    
    ' 2. P列以降の「対応する列」を特定して着色
    ' RunDRT側のロジック「targetCol = 15 + k」に合わせる
    ' optIdx=2(1つ目のλ)なら15+1=16列目(P列)
    targetCol = 15 + (optIdx - 1)
    
    If targetCol <= lastCol Then
        ' その列全体（あるいはデータ範囲）を塗りつぶす
        ws.Columns(targetCol).Interior.Color = vbYellow
        ' タイトル部分（1行目）を太字にする
        ws.Cells(1, targetCol).Font.Bold = True
        
        ' 該当する列が表示されるようにスクロール
        Application.Goto ws.Cells(1, targetCol), True
    End If
    
'    MsgBox "判定完了。" & vbCrLf & _
'           "Optimal λ: " & ws.Cells(optIdx, 12).Value & vbCrLf & _
'           "着色列: " & Split(ws.Cells(1, targetCol).Address, "$")(1)
End Sub


Sub CalculateImpedanceFromOptimalDRT()
    Dim ws As Worksheet
    Dim lastRowFreq As Long, nPoints As Long
    Dim lastRowStep As Long, nTau As Long
    Dim i As Long, j As Long
    Dim optimalStep As Long
    Dim targetCol As Long
    Dim rInf As Double
    Dim PI As Double: PI = 3.14159265358979
    
    Dim freq() As Double, omega() As Double
    Dim tauGrid() As Double, gValue() As Double
    Dim zRealCalc() As Double, zImagCalc() As Double
    
    Set ws = ActiveSheet
    
    ' --- 1. Optimalなステップを特定 ---
    lastRowStep = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    optimalStep = 0
    For i = 2 To lastRowStep
        ' K列(11列目)のフラグを確認
        If ws.Cells(i, 11).Value = "Optimal" Then
            optimalStep = i - 1
            Exit For
        End If
    Next i
    
    If optimalStep = 0 Then
        MsgBox "K列に 'Optimal' フラグが見つかりません。先に判定マクロを実行してください。"
        Exit Sub
    End If
    
    ' --- 2. データの準備 ---
    lastRowFreq = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    nPoints = lastRowFreq - 1
    ReDim freq(1 To nPoints), omega(1 To nPoints)
    ReDim zRealCalc(1 To nPoints), zImagCalc(1 To nPoints)
    
    For i = 1 To nPoints
        freq(i) = ws.Cells(i + 1, 1).Value
        omega(i) = 2 * PI * freq(i)
    Next i
    
    ' --- 3. DRTデータの読み込み ---
    ' O列(15列目)の数値データ(周波数グリッド)をカウント
    nTau = 0
    For j = 2 To ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then
            nTau = nTau + 1
        Else
            ' 文字列（R_infなど）が出現したら終了
            Exit For
        End If
    Next j
    
    If nTau = 0 Then
        MsgBox "O列に有効な周波数グリッドが見つかりません。"
        Exit Sub
    End If
    
    ReDim tauGrid(1 To nTau), gValue(1 To nTau)
    targetCol = 15 + optimalStep ' P列以降の該当列
    
    For j = 1 To nTau
        ' O列の周波数から時定数tauへ変換
        tauGrid(j) = 1 / (2 * PI * ws.Cells(j + 1, 15).Value)
        ' 対応する列から強度gを取得
        gValue(j) = ws.Cells(j + 1, targetCol).Value
    Next j
    
    ' --- 4. R_inf（直列抵抗）の取得 ---
    ' DRTデータ末尾の次(nTau + 2行目)に書き込まれた値を取得
    rInf = 0
    If IsNumeric(ws.Cells(nTau + 2, targetCol).Value) Then
        rInf = ws.Cells(nTau + 2, targetCol).Value
    End If

    ' --- 5. インピーダンスの再計算 ---
    For i = 1 To nPoints
        zRealCalc(i) = rInf ' 実部のオフセット
        zImagCalc(i) = 0
        
        For j = 1 To nTau
            Dim wt As Double: wt = omega(i) * tauGrid(j)
            ' Tikhonov正規化で用いたカーネル関数に基づく再構成
            zRealCalc(i) = zRealCalc(i) + gValue(j) / (1 + wt ^ 2)
            zImagCalc(i) = zImagCalc(i) - (gValue(j) * wt) / (1 + wt ^ 2)
        Next j
    Next i
    
    ' --- 6. 結果の出力 (H, I, J列) ---
    ws.Cells(1, 8).Value = "Freq(Hz)"
    ws.Cells(1, 9).Value = "calc-Z'"
    ws.Cells(1, 10).Value = "calc-Z''"
    
    For i = 1 To nPoints
        ws.Cells(i + 1, 8).Value = freq(i)
        ws.Cells(i + 1, 9).Value = zRealCalc(i)
        ws.Cells(i + 1, 10).Value = zImagCalc(i)
    Next i
    
'    MsgBox "OptimalなDRT（" & ws.Cells(optimalStep + 1, 12).Value & "）から再計算を完了しました。"
End Sub














Sub CreateColeColeComparisonChart_Styled_A7()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long, i As Long
    Dim startRow As Long, endRow As Long
    
    Set ws = ActiveSheet
    
    ' --- 1. "Used" 範囲の特定 ---
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    startRow = 0: endRow = 0
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value = "Used" Then
            If startRow = 0 Then startRow = i
            endRow = i
        End If
    Next i

    ' --- 2. 既存グラフの削除 ---
    On Error Resume Next
    ws.ChartObjects("ColeCole_Comparison").Delete
    On Error GoTo 0
    
    ' --- 3. グラフの作成 ---
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("A7").Left, _
        Top:=ws.Range("A7").Top, _
        Width:=400, _
        Height:=420)
    chartObj.Name = "ColeCole_Comparison"
    
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        
        ' タイトル抹消（前回の修正を反映）
        .HasTitle = True
        DoEvents ' ここで描画を待機
        .ChartTitle.Text = ""
        .HasTitle = False
        
        ' シリーズ1: Measured Data
        Dim ser1 As Series
        Set ser1 = .SeriesCollection.NewSeries
        With ser1
            .Name = "Measured Data"
            .XValues = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2))
            .Values = ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 3))
            .Format.Line.ForeColor.RGB = RGB(160, 160, 160)
            .Format.Line.DashStyle = msoLineDash
            .MarkerStyle = xlMarkerStyleSquare
            .MarkerSize = 4
        End With
        
        ' シリーズ2: Optimal Fit
        If startRow > 0 Then
            Dim ser2 As Series
            Set ser2 = .SeriesCollection.NewSeries
            With ser2
                .Name = "Optimal Fit (DRT)"
                .XValues = ws.Range(ws.Cells(startRow, 9), ws.Cells(endRow, 9))
                .Values = ws.Range(ws.Cells(startRow, 10), ws.Cells(endRow, 10))
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                .Format.Line.Weight = 2.5
                .MarkerStyle = xlMarkerStyleNone
            End With
        End If

        ' --- 4. 軸と外観の設定 ---
        ' (軸設定の後に PlotArea を操作するのが安全)
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Z' / Ohm"
            .TickLabelPosition = xlTickLabelPositionHigh
            .MajorTickMark = xlInside
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "-Z'' / Ohm"
            .ReversePlotOrder = True
            .MajorTickMark = xlInside
            .CrossesAt = 0
        End With

        ' --- 5. エラー対策：プロットエリアの操作 ---
        ' グラフの自動計算を完了させるための「おまじない」
        DoEvents
        Dim tempTop As Double
        On Error Resume Next
        tempTop = .PlotArea.InsideTop ' 一度値を読み取ってオブジェクトを認識させる
        On Error GoTo 0
        
        ' ここでサイズ指定
        With .PlotArea
            .Top = 20
            .Left = 60
            .Width = 320
            .Height = 280
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        ' 凡例とタイトルの位置
        .HasLegend = True
        With .Legend
            .Position = xlLegendPositionBottom
            .Top = chartObj.Height - 30
        End With
        
        .Axes(xlCategory).AxisTitle.Top = chartObj.Height - 75
        .ChartArea.Format.Line.Visible = msoFalse
        
        ' 1:1 スケール調整
        Call RescaleAxesToSquare(chartObj.Chart)
    End With
    
   ' MsgBox "Cole-Coleプロットを作成しました。", vbInformation
End Sub


' ------------------------------------------------------------
' サブ関数：軸の最大値を同期させて 1:1 スケールにする
' ------------------------------------------------------------
Private Sub RescaleAxesToSquare(cht As Chart)
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim maxRange As Double
    
    ' 自動設定を一度リセット
    cht.Axes(xlCategory).MinimumScaleIsAuto = True
    cht.Axes(xlCategory).MaximumScaleIsAuto = True
    cht.Axes(xlValue).MinimumScaleIsAuto = True
    cht.Axes(xlValue).MaximumScaleIsAuto = True
    
    xMin = cht.Axes(xlCategory).MinimumScale
    xMax = cht.Axes(xlCategory).MaximumScale
    yMin = cht.Axes(xlValue).MinimumScale
    yMax = cht.Axes(xlValue).MaximumScale
    
    ' 大きい方のレンジを採用
    maxRange = IIf((xMax - xMin) > (yMax - yMin), (xMax - xMin), (yMax - yMin))
    
    cht.Axes(xlCategory).MaximumScale = xMin + maxRange
    cht.Axes(xlValue).MaximumScale = yMin + maxRange
End Sub

Sub CreateDRTSpectrumChart_Styled_G7()
Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim lastRowStep As Long, nTauRow As Long
    Dim i As Long, j As Long
    Dim optimalStep As Long
    Dim targetCol As Long
    Dim startTauRow As Long, endTauRow As Long
    Dim refCht As ChartObject
    
    Set ws = ActiveSheet
    
    ' --- 1. Optimalなステップを特定 ---
    lastRowStep = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    optimalStep = 0
    For i = 2 To lastRowStep
        If ws.Cells(i, 11).Value = "Optimal" Then
            optimalStep = i - 1
            Exit For
        End If
    Next i
    
    If optimalStep = 0 Then
        MsgBox "先に判定マクロを実行してください。"
        Exit Sub
    End If
    
' --- 2. プロット範囲の特定 (数値のみを抽出し、端点のスパイクをカット) ---
    nTauRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    targetCol = 15 + optimalStep
    
    startTauRow = 0
    endTauRow = 0
    
'    ' 数値データの有効範囲を探す
'    For j = 2 To nTauRow
'        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then
'            If startTauRow = 0 Then startTauRow = j
'            endTauRow = j
'        Else
'            Exit For
'        End If
'    Next j
'
'    ' 【重要】低周波側（末尾）のスパイクを除外する処理
'    ' データが十分にある場合、末尾の3点（最も低周波な部分）をプロットから外す
'    If endTauRow > startTauRow + 10 Then
'        endTauRow = endTauRow - 3
'    End If
'
    ' 【オプション】高周波側（開始点）も不安定なら1～2点飛ばす
    ' If startTauRow > 0 Then startTauRow = startTauRow + 1
    ' --- CreateDRTSpectrumChart 内の範囲特定部分 ---
' 数値データの有効範囲を探す
    For j = 2 To nTauRow
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then
            If startTauRow = 0 Then startTauRow = j
            endTauRow = j
        Else
            Exit For
        End If
    Next j
    
    ' 高周波側（開始点）のカット
    If startTauRow > 0 Then
        startTauRow = startTauRow + CUT_HIGH_FREQ
    End If
    
    ' 低周波側（末尾）のカット
    If endTauRow > startTauRow + 10 Then
        endTauRow = endTauRow - CUT_LOW_FREQ
    End If
    
    
    
    ' --- 3. サイズリファレンス (Cole-Coleプロットに合わせる) ---
    ' 1つ目のグラフ（Nyquist図）のサイズを取得。なければデフォルト値。
    On Error Resume Next
    Set refCht = ws.ChartObjects(1)
    On Error GoTo 0
    
    ' --- 4. グラフの作成/更新 ---
    Dim chartName As String: chartName = "DRTSpectrumChart"
    On Error Resume Next
    Set chtObj = ws.ChartObjects(chartName)
    On Error GoTo 0
    
    If chtObj Is Nothing Then
        ' 新規作成（サイズは後で設定）
        Set chtObj = ws.ChartObjects.Add(Left:=ws.Range("G7").Left, Top:=ws.Range("G7").Top, _
                                         Width:=400, Height:=300)
        chtObj.Name = chartName
    End If
    
    ' 配置とサイズを強制指定
    With chtObj
        .Left = ws.Range("G7").Left
        .Top = ws.Range("G7").Top
        If Not refCht Is Nothing Then
            .Width = refCht.Width
            .Height = refCht.Height
        End If
    End With
    
    With chtObj.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        
        ' タイトルを完全に削除
        .HasTitle = False
        
        ' 既存シリーズのクリアと追加
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "DRT Spectrum"
            .XValues = ws.Range(ws.Cells(startTauRow, 15), ws.Cells(endTauRow, 15))
            .Values = ws.Range(ws.Cells(startTauRow, targetCol), ws.Cells(endTauRow, targetCol))
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' 赤色
            .Format.Line.Weight = 2
        End With
        
        ' --- 軸とフォーマットの設定 ---
        
        ' X軸 (Frequency)
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Frequency (Hz)"
            .ScaleType = xlLogarithmic
            ' ラベルをグラフの下（枠外）に移動
            .TickLabelPosition = xlTickLabelPositionLow
            ' 枠線
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        End With
        
        ' Y軸 (g(tau))
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "g(tau) / Ohm"
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        End With
        
        ' プロットエリアの枠線
        .PlotArea.Format.Line.Visible = msoTrue
        .PlotArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .PlotArea.Format.Line.Weight = 1
        
        ' 凡例非表示
        .HasLegend = False
    End With
    
'    MsgBox "DRTスペクトルをG7セルに、Cole-Coleプロットと同じサイズで配置しました。"
End Sub



Sub ActiveSheetDRT_all()

Call CalculateMagAndPhase
 Call RunDRT
 Call FindOptimalLambda_Normalized
 Call CalculateImpedanceFromOptimalDRT
 Call CreateColeColeComparisonChart_Styled_A7
 Call CreateDRTSpectrumChart_Styled_G7

End Sub
