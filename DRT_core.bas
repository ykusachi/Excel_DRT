Attribute VB_Name = "Module21"
Option Explicit
' ============================================================
' GLOBAL SETTINGS (Modify these values to adjust analysis)
' ============================================================

' 1. KK-Filter Threshold (%)
' Points with residuals exceeding this percentage will be excluded.
Public Const KK_THRESHOLD As Double = 3#

' 2. Lambda (É…) Scan Settings
' Scans regularization parameter from 10^(-START) to 10^(-END).
Public Const LAMBDA_START_EXP As Double = 0#   ' Start exponent (e.g., 0 for 10^0)
Public Const LAMBDA_END_EXP As Double = 10#    ' End exponent (e.g., 10 for 10^-10)
Public Const LAMBDA_STEP As Double = 0.2       ' Increment step for the exponent

' 3. DRT Spectrum Endpoint Trimming
' Number of data points to remove from the edges to eliminate spikes.
Public Const CUT_LOW_FREQ As Integer = 3       ' Points to cut at Low-Frequency end
Public Const CUT_HIGH_FREQ As Integer = 0      ' Points to cut at High-Frequency end (0 = No cut)



' ======================================================================================
' DESCRIPTION:
' This preprocessing procedure calculates the Impedance Magnitude (|Z|) and
' the Phase Angle (theta) from the raw Complex Impedance data (Z' and Z'').
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. Magnitude calculation: Based on the Pythagorean theorem:
'    |Z| = Sqr( (Z')^2 + (Z'')^2 )
' 2. Phase Angle calculation: Based on the inverse tangent of the ratio
'    between the imaginary and real parts: theta = ArcTan( Z'' / Z' )
' 3. Degree Conversion: Converts the result from radians to degrees
'    (theta_deg = theta_rad * 180 / PI).
'
' WHY:
' These values are essential for Bode plots and for the initial assessment
' of the electrochemical system's response across the frequency spectrum.
'
' INPUTS:
' - Real Impedance (Z') in Column B.
' - Imaginary Impedance (Z'') in Column C.
'
' OUTPUTS:
' - Populates Impedance Magnitude in Column D.
' - Populates Phase Angle (Degrees) in Column E.
' ======================================================================================

Sub CalculateMagAndPhase()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    
    Set ws = ActiveSheet
    
    ' Get the last row based on Column B (Z')
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Check if data exists from row 2 onwards
    If lastRow < 2 Then
        MsgBox "Target data not found (Data required from row 2 onwards)."
        Exit Sub
    End If
    
    rowCount = lastRow - 1
    
    ' Input headers (D1, E1)
    ws.Cells(1, 4).Value = "Magnitude"
    ws.Cells(1, 5).Value = "Phase"
    
    ' --- Calculate Magnitude (Column D) ---
    ' Formula: SQRT(Z'^2 + Z''^2)
    With ws.Range("D2:D" & lastRow)
        .FormulaR1C1 = "=SQRT(RC[-2]^2 + RC[-1]^2)"
    End With
    
    ' --- Calculate Phase (Column E) ---
    ' Formula: ATAN2(Z', Z'') * 180 / PI
    With ws.Range("E2:E" & lastRow)
        .FormulaR1C1 = "=ATAN2(RC[-3], RC[-2]) * 180 / PI()"
    End With
    
    ' Convert formulas to values to reduce file size and improve performance
    With ws.Range("D2:E" & lastRow)
        .Value = .Value
    End With
    
    'MsgBox "Magnitude and Phase have been calculated for the active sheet.", vbInformation
End Sub



' ==========================================
' ============================================================
' 1. Main Routine (R_infinity Integrated & Artifact-Suppressed)
' ============================================================
' ======================================================================================
' DESCRIPTION:
' This is the Core Engine of the DRT analysis. It performs a "Lambda Scan" to
' calculate the Distribution of Relaxation Times (g(tau)) across multiple
' regularization strengths.
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. Discretization: Constructs a Frequency Grid (Tau Grid) in Column O and
'    initializes the Kernel Matrix (A) based on Debye elements.
' 2. Tikhonov Inversion: Solves the ill-posed inverse problem:
'    Minimize ||A*g - Z||^2 + lambda * ||g||^2.
' 3. Non-Negativity Constraint: Employs an iterative Active-Set (NNLS-like) algorithm
'    to ensure all calculated intensities (g) are physically valid (>= 0).
' 4. Lambda Iteration: Loops from 10^0 down to 10^-8 to generate a series of
'    potential solutions (the L-curve data).
' 5. Ohmic Resistance (R_inf): Simultaneously estimates the series resistance
'    as a constant offset in the Real part of the impedance.
'
' INPUTS:
' - Experimental Z' and Z'' data (Columns B and C).
' - Measurement frequencies (Column A).
'
' OUTPUTS:
' - Populates the Frequency Grid in Column O.
' - Generates DRT spectra for each lambda step in Columns P, Q, R, etc.
' - Calculates Residuals and Solution Norms for L-curve analysis in Columns M and N.
' ======================================================================================
Sub RunDRT()
    Dim ws As Worksheet
    Dim lastRow As Long, nPoints As Long, nValid As Long, nTau As Long
    Dim i As Long, j As Long, k As Long
    Dim freq() As Double, zReal() As Double, zImag() As Double
    Dim vFreq() As Double, vReal() As Double, vImag() As Double, vOmega() As Double
    Dim isValid() As Boolean
    Dim A() As Double, b() As Double, lambda As Double, tauGrid() As Double
    Dim PI As Double: PI = 3.14159265358979
    
    ' NNLS (Non-Negative Least Squares) Iteration Control
    Dim outerIter As Long, innerIter As Long
    Const MAX_OUTER As Long = 500
    Const MAX_INNER As Long = 200
    
    Set ws = ActiveSheet
    
    ' Get the last row of data
    On Error Resume Next
    lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    nPoints = lastRow - 1
    
    If nPoints < 10 Then
        MsgBox "Insufficient data points (Minimum 10 required).", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "Validating data (KK-Consistency Fit)..."
    
    ' --- 1. Load Data ---
    ReDim freq(1 To nPoints), zReal(1 To nPoints), zImag(1 To nPoints), isValid(1 To nPoints)
    For i = 1 To nPoints
        freq(i) = ws.Cells(i + 1, 1).Value
        zReal(i) = ws.Cells(i + 1, 2).Value
        zImag(i) = ws.Cells(i + 1, 3).Value
    Next i
    
    ' --- 2. Execute KK-Filter (Kramers-Kronig) ---
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
        
        ' Filter based on the global KK_THRESHOLD constant
        If resPerc <= KK_THRESHOLD Then
            isValid(i) = True
            nValid = nValid + 1
            ws.Cells(i + 1, 7).Value = "Used"
            ws.Cells(i + 1, 7).Interior.Color = RGB(200, 255, 200) ' Light Green
        Else
            isValid(i) = False
            ws.Cells(i + 1, 7).Value = "Excluded(KK)"
            ws.Cells(i + 1, 7).Interior.Color = RGB(255, 200, 200) ' Light Red
        End If
    Next i
    
    If nValid < 5 Then
        Application.StatusBar = False
        MsgBox "Insufficient valid data points within the KK threshold.", vbCritical
        Exit Sub
    End If

    ' --- 3. Reconstruct Matrices (Used points only) ---
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

    ' --- 4. Create Time-Constant Grid (Matches 'Used' frequency range) ---
    nTau = 100
    ReDim tauGrid(1 To nTau)
    Dim minF As Double: minF = GetMin(vFreq)
    Dim maxF As Double: maxF = GetMax(vFreq)
    For j = 1 To nTau
        tauGrid(j) = (1 / (2 * PI * maxF)) * (maxF / minF) ^ ((j - 1) / (nTau - 1))
    Next j

    ' --- 5. Build A and b Matrices for NNLS (R_inf added as nTau+1 column) ---
    ReDim A(1 To 2 * nValid, 1 To nTau + 1), b(1 To 2 * nValid, 1 To 1)
    For i = 1 To nValid
        b(i, 1) = vReal(i)
        b(i + nValid, 1) = -vImag(i)
        For j = 1 To nTau
            Dim wt As Double: wt = vOmega(i) * tauGrid(j)
            A(i, j) = 1 / (1 + wt ^ 2)              ' Real part contribution
            A(i + nValid, j) = wt / (1 + wt ^ 2)    ' Imaginary part contribution
        Next j
        ' Final column: R_infinity component (Frequency-independent offset)
        A(i, nTau + 1) = 1
        A(i + nValid, nTau + 1) = 0
    Next i
    
    Dim matAtA As Variant: matAtA = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(A), A)
    Dim matAtb As Variant: matAtb = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(A), b)

    ' --- 6. Set Summary Headers ---
    ws.Cells(1, 11).Value = "Flag"
    ws.Cells(1, 12).Value = "Lambda (É…)"
    ws.Cells(1, 13).Value = "Log(ResidualSum)"
    ws.Cells(1, 14).Value = "Log(SolutionSum)"
    ws.Cells(1, 15).Value = "Freq_Grid(Hz)"
    
    ' --- 7. Execute Lambda (É…) Scan ---
    Dim nSteps As Long
    nSteps = Int((LAMBDA_END_EXP - LAMBDA_START_EXP) / LAMBDA_STEP) + 1
    
    For k = 1 To nSteps
        Dim exponent As Double: exponent = LAMBDA_START_EXP + (k - 1) * LAMBDA_STEP
        lambda = 10 ^ (-exponent)
        
        Application.StatusBar = "[" & ActiveSheet.Name & "] Analyzing: É… = 10^-" & _
                                Format(exponent, "0.00") & " (" & k & "/" & nSteps & ")"
        
        DoEvents
        
        Dim x() As Double: ReDim x(1 To nTau + 1, 1 To 1)
        Dim P() As Boolean: ReDim P(1 To nTau + 1)
        Dim w() As Double: ReDim w(1 To nTau + 1)
        Dim isConverged As Boolean: isConverged = False
        
        ' NNLS Core Algorithm (Active Set Method)
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
        
        ' Write results to worksheet
        Call WriteSpectralResultShifted(ws, k, lambda, exponent, A, x, b, nValid, nTau, tauGrid, failStr)
    Next k
    
    Application.StatusBar = False
    ' MsgBox "Analysis completed."
End Sub

'' ======================================================================================
' DESCRIPTION:
' This function solves a "Reduced" Linear System as part of an Active-Set
' (NNLS-like) algorithm to ensure non-negative DRT results (g(tau) >= 0).
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. Masking (Boolean P): Uses a boolean array 'P' to identify which variables
'    are "Active" (Positive) and which are constrained to Zero.
' 2. Matrix Reduction: Creates a smaller, temporary subsystem consisting only
'    of the Active variables from the Regularized Normal Equation (A^T * A + lambda * I).
' 3. Solving: Performs a full Matrix Inversion (via Gaussian elimination or LU)
'    on the reduced k x k matrix to find the optimal amplitudes for the active set.
'
' WHY:
' Direct solvers can return negative intensities, which are physically impossible
' in DRT. By solving only for the "Active" subset, the algorithm iteratively
' converges to a physically valid, non-negative distribution.
'
' INPUTS:
' - AtA: The full Gramian matrix (A transposed * A).
' - Atb: The full projection vector (A transposed * b).
' - P(): A Boolean array where 'True' indicates an active (non-zero) variable.
' - n: The total number of points in the tau grid.
' - lam: The regularization parameter (lambda).
'
' RETURNS:
' - A Variant array (Double) containing the solution for the active variables,
'   with constrained variables set to zero.
' ======================================================================================
Function SolveReducedSystem(AtA, Atb, P() As Boolean, n As Long, lam As Double) As Variant
    Dim tempA() As Double: ReDim tempA(1 To n, 1 To n)
    Dim tempB() As Double: ReDim tempB(1 To n, 1 To 1)
    Dim i As Long, j As Long
    Dim stab As Double: stab = 1E-11 ' Tikhonov-like stabilization factor
    
    For i = 1 To n
        If P(i) = True Then
            tempB(i, 1) = Atb(i, 1)
            For j = 1 To n
                If P(j) = True Then
                    tempA(i, j) = AtA(i, j)
                    If i = j Then
                        ' Apply regularization to all elements except the last one (R_infinity)
                        If i < n Then
                            tempA(i, j) = tempA(i, j) + lam + stab
                        Else
                            ' Only add minimal stabilization for R_infinity
                            tempA(i, j) = tempA(i, j) + stab
                        End If
                    End If
                Else
                    tempA(i, j) = 0
                End If
            Next j
        Else
            ' If the variable is not in the active set, set diagonal to 1 and RHS to 0
            tempA(i, i) = 1
            tempB(i, 1) = 0
        End If
    Next i
    
    On Error Resume Next
    ' Solve via matrix inversion: x = (A^T * A)^-1 * (A^T * b)
    SolveReducedSystem = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(tempA), tempB)
    
    ' Handle cases where the matrix might be singular
    If Err.Number <> 0 Then
        SolveReducedSystem = Empty
    End If
    On Error GoTo 0
End Function

'' ======================================================================================
' DESCRIPTION:
' This function calculates the Matrix-Vector product of the Regularized Normal
' Equation: (A^T * A + lambda * I) * x.
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. Normal Equation Product: Computes the standard matrix-vector product between
'    the Gramian matrix (A^T * A) and the current solution vector (x).
' 2. Ridge (Tikhonov) Term: Adds the regularization term (lambda * x) directly
'    to the result. This is equivalent to adding lambda to the diagonal
'    elements of the A^T * A matrix.
'
' WHY:
' This product is a critical component of iterative solvers (like Conjugate Gradient)
' used to find the Distribution of Relaxation Times (DRT) without performing
' a full matrix inversion, which is computationally expensive.
'
' INPUTS:
' - AtA: The Gramian matrix (A transposed * A).
' - x: The current solution vector (g intensities).
' - n: The dimension of the vector/matrix (number of tau grid points).
' - lam: The regularization parameter (lambda).
'
' RETURNS:
' - A Double array representing the resulting product vector.
' ======================================================================================

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

'' ======================================================================================
' DESCRIPTION:
' This procedure outputs the results of a single DRT (Distribution of Relaxation Times)
' inversion for a specific regularization parameter (Lambda) to the worksheet.
'
' LOGIC & DATA OUTPUT:
' 1. Column Positioning: Shifts the target column (starting from Col P) based on the
'    index 'k' of the lambda scan.
' 2. Metadata: Writes the Lambda value (10^-expo) and the current residual (b - Ax)
'    to the top of the column.
' 3. Spectral Data: Writes the calculated DRT intensity 'x' (amplitudes) corresponding
'    to the frequency grid.
' 4. Ohmic Offset: Appends the R_inf (series resistance) at the end of the spectral data.
' 5. Formatting: Applies bold font to the headers for improved readability.
'
' INPUTS:
' - ws: The destination Worksheet.
' - k: The current iteration index for the lambda loop.
' - lam: The actual Lambda value used.
' - expo: The exponent of Lambda (e.g., 2 for 10^-2).
' - A, x, b: The Kernel matrix, Solution vector (g), and Measurement vector (Z).
' - nV, nT: Number of measurement points and number of tau grid points.
' - tG(): The frequency/time-constant grid array.
' - msg: A status message (e.g., "Optimal").
' ======================================================================================

Sub WriteSpectralResultShifted(ws, k, lam, expo, A, x, b, nV, nT, tG, msg As String)
    Dim Ag As Variant: Ag = Application.WorksheetFunction.MMult(A, x)
    Dim resSumSq As Double: resSumSq = 0 ' Residual Sum of Squares (RSS)
    Dim solSumSq As Double: solSumSq = 0 ' Solution Norm (Regularization term)
    Dim i As Long
    
    ' Calculate Residual Sum of Squares (Data fidelity term)
    For i = 1 To 2 * nV
        resSumSq = resSumSq + (Ag(i, 1) - b(i, 1)) ^ 2
    Next i
    
    ' Calculate Solution Norm (Model complexity term)
    ' Excluding the last element (R_inf) from the norm calculation
    For i = 1 To nT
        solSumSq = solSumSq + x(i, 1) ^ 2
    Next i
    
    ' Write L-curve parameters to Columns L, M, and N
    ws.Cells(k + 1, 12).Value = lam
    ws.Cells(k + 1, 13).Value = WorksheetFunction.Log10(resSumSq + 1E-20)
    ws.Cells(k + 1, 14).Value = WorksheetFunction.Log10(solSumSq + 1E-20)
    
    ' Set Column Header for the current Lambda step
    ws.Cells(1, 15 + k).Value = "É…:10^-" & Format(expo, "0.00") & msg
    
    ' Output Frequency Grid and DRT Spectrum g(tau)
    For i = 1 To nT
        If k = 1 Then
            ' Calculate and write Frequency (Hz) based on Tau Grid in Column O
            ws.Cells(i + 1, 15).Value = 1 / (2 * 3.14159265358979 * tG(i))
        End If
        ' Write the DRT amplitude for each frequency point
        ws.Cells(i + 1, 15 + k).Value = x(i, 1)
    Next i
    
    ' Output R_infinity (Ohmic Resistance offset)
    If k = 1 Then
        ws.Cells(nT + 2, 15).Value = "R_inf(Ohm)"
    End If
    ws.Cells(nT + 2, 15 + k).Value = x(nT + 1, 1)
End Sub

'' ======================================================================================
' DESCRIPTION:
' This function performs a Kramers-Kronig (KK) Consistency Fit using a series
' of Debye elements (RC circuits). It is used to verify the physical validity
' of Electrochemical Impedance Spectroscopy (EIS) data.
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. Kernel Matrix (Ak): Constructs a matrix based on the Real part of the Debye
'    transfer function: 1 / (1 + (omega * tau)^2).
' 2. Target Vector (Bk): Uses the measured Real Impedance (zR) as the target.
' 3. Solving (Least Squares): Solves the linear system using the Normal Equation:
'    x = (A^T * A + lambda*I)^-1 * (A^T * b).
' 4. Regularization: Adds a small Ridge (Tikhonov) term (1E-06) to the diagonal
'    of the A^T * A matrix to ensure numerical stability during inversion.
'
' INPUTS:
' - n: Number of frequency points.
' - f(): Array of frequencies (Hz).
' - zR(): Array of measured Real Impedance.
' - zI(): Array of measured Imaginary Impedance (used for cross-validation).
'
' RETURNS:
' - A Variant array containing the calculated fitting weights (amplitudes).
' ======================================================================================

Function PerformKKFit(n As Long, f() As Double, zR() As Double, zI() As Double) As Variant
    Dim i As Long, j As Long, Ak() As Double, Bk() As Double
    Dim PI As Double: PI = 3.14159265358979
    
    ReDim Ak(1 To n, 1 To n), Bk(1 To n, 1 To 1)
    
    ' --- 1. Construct the Kernel Matrix for KK-Fit ---
    For i = 1 To n
        ' Use Real part as the target (b vector) for fitting
        Bk(i, 1) = zR(i)
        
        For j = 1 To n
            Dim f_j As Double
            ' Avoid division by zero for frequency
            If f(j) <= 0 Then
                f_j = 1E-10
            Else
                f_j = f(j)
            End If
            
            Dim omega_i As Double: omega_i = 2 * PI * f(i)
            Dim tau_j As Double: tau_j = 1 / (2 * PI * f_j)
            
            ' Debye element kernel for the real part: 1 / (1 + (omega * tau)^2)
            Ak(i, j) = 1 / (1 + (omega_i * tau_j) ^ 2)
        Next j
    Next i
    
    ' --- 2. Solve the Linear System (Normal Equation: A^T * A * x = A^T * b) ---
    Dim AtA As Variant: AtA = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(Ak), Ak)
    
    ' Add a small Ridge regularization term (Tikhonov) to stabilize matrix inversion
    For i = 1 To n
        AtA(i, i) = AtA(i, i) + 1E-06
    Next i
    
    On Error Resume Next
    ' Calculate fitting weights x = (A^T * A)^-1 * (A^T * b)
    PerformKKFit = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(AtA), _
                   Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(Ak), Bk))
    
    If Err.Number <> 0 Then
        PerformKKFit = Empty
    End If
    On Error GoTo 0
End Function


' ======================================================================================
' DESCRIPTION:
' This utility function retrieves the Maximum value from a given array of Double
' precision numbers.
'
' LOGIC:
' 1. Initializes the maximum tracker (m) with a very small value (-1E+30).
' 2. Iterates through the entire array from LBound to UBound.
' 3. Compares each element and updates the tracker if a larger value is found.
'
' INPUT:
' - arr(): An array of Double numbers (e.g., frequencies, residuals).
'
' RETURNS:
' - The largest numerical value found within the array.
' ======================================================================================

Function GetMax(arr() As Double) As Double
    Dim i As Long, m As Double
    m = -1E+30 ' Initialize with a very small number
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) > m Then
            m = arr(i)
        End If
    Next i
    GetMax = m
End Function

'======================================================================================
' DESCRIPTION:
' This utility function retrieves the Minimum value from a given array of Double
' precision numbers.
'
' LOGIC:
' 1. Initializes the minimum tracker (m) with a very large value (1E+30).
' 2. Iterates through the entire array from LBound to UBound.
' 3. Compares each element and updates the tracker if a smaller value is found.
'
' INPUT:
' - arr(): An array of Double numbers (e.g., frequencies, residuals).
'
' RETURNS:
' - The smallest numerical value found within the array.
' ======================================================================================
Function GetMin(arr() As Double) As Double
    Dim i As Long, m As Double
    m = 1E+30 ' Initialize with a very large number
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) < m Then
            m = arr(i)
        End If
    Next i
    GetMin = m
End Function

' ======================================================================================
' DESCRIPTION:
' This procedure automatically identifies the "Optimal" Regularization Parameter (Lambda)
' using the Normalized L-Curve Method.
'
' LOGIC & MATHEMATICAL APPROACH:
' 1. L-Curve Construction: Uses the log-log coordinates of the Residual Sum of Squares
'    (RSS) and the Solution Norm (model complexity).
' 2. Normalization: Scales both the log(Residual) and log(Solution Norm) into a [0, 1]
'    range to treat both axes with equal weight.
' 3. Minimum Distance Method: Calculates the Euclidean distance from the origin (0,0)
'    to each point on the normalized L-curve. The point with the "Minimum Distance"
'    is identified as the "elbow," representing the optimal balance between
'    data fitting and over-smoothing.
'
' OUTPUTS:
' - Marks the identified row in Column K with the "Optimal" flag and Yellow highlighting.
' - Highlights the corresponding DRT result column (Column P onwards) in Yellow.
' - Automatically scrolls the window to the selected optimal result column.
' ======================================================================================
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
    
    ' --- 1. Get Data Range ---
    ' Col L(12): Lambda list, Col M(13): LogRes, Col N(14): LogSol
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "L-Curve data not found.", vbExclamation
        Exit Sub
    End If
    
    ' --- 2. Reset Existing Highlighting ---
    ' Clear content and formatting in Col K (Flag), Col L (Lambda), and results from Col P (16) onwards
    ws.Columns(11).ClearContents
    ws.Columns(11).Interior.ColorIndex = xlNone
    ws.Columns(12).Interior.ColorIndex = xlNone
    If lastCol >= 16 Then
        ws.Range(ws.Columns(16), ws.Columns(lastCol)).Interior.ColorIndex = xlNone
    End If
    
    ws.Cells(1, 11).Value = "Flag"
    
    ' --- 3. Load and Find Bounds for Normalization ---
    ReDim logRes(2 To lastRow), logSol(2 To lastRow)
    minRes = 1E+30: maxRes = -1E+30
    minSol = 1E+30: maxSol = -1E+30
    
    For i = 2 To lastRow
        logRes(i) = ws.Cells(i, 13).Value ' Col M: Log(ResSum)
        logSol(i) = ws.Cells(i, 14).Value ' Col N: Log(SolSum)
        
        If logRes(i) < minRes Then minRes = logRes(i)
        If logRes(i) > maxRes Then maxRes = logRes(i)
        If logSol(i) < minSol Then minSol = logSol(i)
        If logSol(i) > maxSol Then maxSol = logSol(i)
    Next i
    
    ' --- 4. Normalize to [0, 1] Range ---
    ReDim normRes(2 To lastRow), normSol(2 To lastRow)
    For i = 2 To lastRow
        normRes(i) = (logRes(i) - minRes) / (maxRes - minRes + 1E-20)
        normSol(i) = (logSol(i) - minSol) / (maxSol - minSol + 1E-20)
    Next i
    
    ' --- 5. Minimum Distance Method (Shortest distance from origin) ---
    ' Identifies the "elbow" of the L-curve in normalized space
    minDistance = 1E+30
    optIdx = 2 ' Initial value (Row 2)
    
    For i = 2 To lastRow
        Dim dist As Double
        dist = Sqr(normRes(i) ^ 2 + normSol(i) ^ 2)
        If dist < minDistance Then
            minDistance = dist
            optIdx = i ' Row index for the optimal Lambda
        End If
    Next i
    
    ' --- 6. Highlight Results ---
    ' 1. Highlight the Optimal row in the L-Curve list (Cols K and L)
    ws.Cells(optIdx, 11).Value = "Optimal"
    ws.Cells(optIdx, 11).Interior.Color = vbYellow
    ws.Cells(optIdx, 12).Interior.Color = vbYellow
    
    ' 2. Identify and highlight the corresponding result column starting from Col P
    ' Matches RunDRT logic where targetCol = 15 + k
    ' If optIdx = 2 (1st Lambda), targetCol = 15 + (2-1) = 16 (Col P)
    targetCol = 15 + (optIdx - 1)
    
    If targetCol <= lastCol Then
        ' Fill the entire result column
        ws.Columns(targetCol).Interior.Color = vbYellow
        ' Set header (Row 1) to Bold
        ws.Cells(1, targetCol).Font.Bold = True
        
        ' Scroll to the identified optimal column
        Application.GoTo ws.Cells(1, targetCol), True
    End If
    
    ' MsgBox "Optimal Lambda identified: " & ws.Cells(optIdx, 12).Value
End Sub

' ======================================================================================
' DESCRIPTION:
' This procedure reconstructs the complex impedance (Z' and Z'') from the
' Distribution of Relaxation Times (DRT) spectrum and the series resistance (R_inf).
'
' PROCESS FLOW:
' 1. Locates the "Optimal" lambda results identified by the optimization macro.
' 2. Extracts the frequency grid, DRT amplitudes (g values), and Ohmic resistance (R_inf).
' 3. Applies the Debye kernel transformation to calculate the theoretical impedance
'    at each measurement frequency.
' 4. Outputs the reconstructed Z' and Z'' values to Columns H, I, and J.
'
' INPUTS:
' - Frequency data in Column A.
' - DRT results and Frequency Grid in Columns O, P, and beyond.
' - "Optimal" flag in Column K.
'
' OUTPUTS:
' - Reconstructed Frequency, Z', and Z'' in Columns H, I, and J.
' ======================================================================================
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
    
    ' --- 1. Identify the Optimal Lambda Step ---
    ' Search for the "Optimal" flag in Column K (Column 11)
    lastRowStep = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    optimalStep = 0
    For i = 2 To lastRowStep
        If ws.Cells(i, 11).Value = "Optimal" Then
            optimalStep = i - 1
            Exit For
        End If
    Next i
    
    If optimalStep = 0 Then
        MsgBox "The 'Optimal' flag was not found in Column K. Please run the Optimization macro first.", vbExclamation
        Exit Sub
    End If
    
    ' --- 2. Prepare Frequency Data ---
    ' Load measurement frequencies from Column A
    lastRowFreq = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    nPoints = lastRowFreq - 1
    ReDim freq(1 To nPoints), omega(1 To nPoints)
    ReDim zRealCalc(1 To nPoints), zImagCalc(1 To nPoints)
    
    For i = 1 To nPoints
        freq(i) = ws.Cells(i + 1, 1).Value
        omega(i) = 2 * PI * freq(i)
    Next i
    
    ' --- 3. Load DRT Results (Frequency Grid and g values) ---
    ' Count numeric data points in the Frequency Grid (Column O)
    nTau = 0
    For j = 2 To ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then
            nTau = nTau + 1
        Else
            ' Exit loop if a label (like "R_inf") or empty cell is encountered
            Exit For
        End If
    Next j
    
    If nTau = 0 Then
        MsgBox "No valid frequency grid found in Column O.", vbExclamation
        Exit Sub
    End If
    
    ReDim tauGrid(1 To nTau), gValue(1 To nTau)
    targetCol = 15 + optimalStep ' Optimal data column (P, Q, R...)
    
    For j = 1 To nTau
        ' Convert grid frequency from Column O to relaxation time (tau)
        tauGrid(j) = 1 / (2 * PI * ws.Cells(j + 1, 15).Value)
        ' Get the DRT amplitude (g) from the identified optimal column
        gValue(j) = ws.Cells(j + 1, targetCol).Value
    Next j
    
    ' --- 4. Retrieve R_inf (Series Resistance) ---
    ' Get the value stored at Row (nTau + 2) in the optimal results column
    rInf = 0
    If IsNumeric(ws.Cells(nTau + 2, targetCol).Value) Then
        rInf = ws.Cells(nTau + 2, targetCol).Value
    End If

    ' --- 5. Reconstruct Impedance based on DRT Model ---
    ' Z_calc = R_inf + Integral[ g(tau) / (1 + j*omega*tau) d(ln tau) ]
    For i = 1 To nPoints
        zRealCalc(i) = rInf ' Real part offset
        zImagCalc(i) = 0
        
        For j = 1 To nTau
            Dim wt As Double: wt = omega(i) * tauGrid(j)
            ' Standard Debye kernel summation
            zRealCalc(i) = zRealCalc(i) + gValue(j) / (1 + wt ^ 2)
            zImagCalc(i) = zImagCalc(i) - (gValue(j) * wt) / (1 + wt ^ 2)
        Next j
    Next i
    
    ' --- 6. Output Reconstructed Data (Columns H, I, J) ---
    ws.Cells(1, 8).Value = "Freq(Hz)"
    ws.Cells(1, 9).Value = "calc-Z' (Ohm)"
    ws.Cells(1, 10).Value = "calc-Z'' (Ohm)"
    
    For i = 1 To nPoints
        ws.Cells(i + 1, 8).Value = freq(i)
        ws.Cells(i + 1, 9).Value = zRealCalc(i)
        ws.Cells(i + 1, 10).Value = zImagCalc(i)
    Next i
    
    'MsgBox "Impedance reconstruction completed.", vbInformation
End Sub



' ======================================================================================
' PURPOSE:
' Standardizes the Nyquist plot by forcing a 1:1 aspect ratio between the
' Real (X) and Imaginary (Y) axes.
'
' LOGIC:
' 1. Calculates the maximum data range (Max - Min) for both axes.
' 2. Determines the larger of the two ranges.
' 3. Resets the Minimum and Maximum scales of both axes to the same interval width.
'
' WHY:
' In electrochemical impedance, a semicircle only looks like a semicircle if
' 1 Ohm on the X-axis is visually equal to 1 Ohm on the Y-axis.
' ======================================================================================
Sub RescaleAxesToSquare(cht As Chart)
    Dim xMin As Double, xMax As Double
    Dim yMin As Double, yMax As Double
    Dim xRange As Double, yRange As Double, maxRange As Double
    
    With cht
        ' Temporarily set to auto to find natural bounds
        .Axes(xlCategory).MinimumScaleIsAuto = True
        .Axes(xlCategory).MaximumScaleIsAuto = True
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
        
        xMin = .Axes(xlCategory).MinimumScale
        xMax = .Axes(xlCategory).MaximumScale
        yMin = .Axes(xlValue).MinimumScale
        yMax = .Axes(xlValue).MaximumScale
        
        xRange = xMax - xMin
        yRange = yMax - yMin
        
        ' Determine the global maximum range to create a square scale
        If xRange > yRange Then
            maxRange = xRange
            .Axes(xlValue).MinimumScale = yMin
            .Axes(xlValue).MaximumScale = yMin + maxRange
        Else
            maxRange = yRange
            .Axes(xlCategory).MinimumScale = xMin
            .Axes(xlCategory).MaximumScale = xMin + maxRange
        End If
        
        ' Force major units to be consistent
        .Axes(xlCategory).MajorUnit = .Axes(xlValue).MajorUnit
    End With
End Sub


' ======================================================================================
' DESCRIPTION:
' Generates a professionally styled Cole-Cole (Nyquist) plot comparing:
'   1. Measured Data: Experimental Z' vs Z'' values (Columns B and C).
'   2. Optimal Fit: Reconstructed Z' vs Z'' from the DRT model (Columns I and J).
'
' LOGIC & DATA RANGE:
' - Scans Column G for the "Used" status to align the fit with valid data points.
' - Inverts the Vertical Axis (-Z'') as per Electrochemical Impedance (EIS) standards.
' - Enforces a 1:1 Aspect Ratio (Square Scaling) by calling "RescaleAxesToSquare"
'   to ensure that semi-circles appear geometrically accurate.
' - Positioned at Cell A7 with a specific layout optimized for scientific reports.
' ======================================================================================
Sub CreateColeColeComparisonChart_Styled_A7()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long, i As Long
    Dim startRow As Long, endRow As Long
    
    Set ws = ActiveSheet
    
    ' --- 1. Identify "Used" data range for plotting ---
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    startRow = 0: endRow = 0
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value = "Used" Then
            If startRow = 0 Then startRow = i
            endRow = i
        End If
    Next i

    ' --- 2. Remove existing chart named "ColeCole_Comparison" ---
    On Error Resume Next
    ws.ChartObjects("ColeCole_Comparison").Delete
    On Error GoTo 0
    
    ' --- 3. Create Chart Object at Cell A7 ---
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("A7").Left, _
        Top:=ws.Range("A7").Top, _
        Width:=400, _
        Height:=420)
    chartObj.Name = "ColeCole_Comparison"
    
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        
        ' Remove Chart Title for clean scientific presentation
        .HasTitle = True
        DoEvents ' Allow UI to catch up
        .ChartTitle.Text = ""
        .HasTitle = False
        
        ' --- Series 1: Measured Data (Reference) ---
        Dim ser1 As Series
        Set ser1 = .SeriesCollection.NewSeries
        With ser1
            .Name = "Measured Data"
            .XValues = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2))
            .Values = ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 3))
            .Format.Line.ForeColor.RGB = RGB(160, 160, 160) ' Grey
            .Format.Line.DashStyle = msoLineDash
            .MarkerStyle = xlMarkerStyleSquare
            .MarkerSize = 4
        End With
        
        ' --- Series 2: Optimal Fit (DRT Result) ---
        If startRow > 0 Then
            Dim ser2 As Series
            Set ser2 = .SeriesCollection.NewSeries
            With ser2
                .Name = "Optimal Fit (DRT)"
                .XValues = ws.Range(ws.Cells(startRow, 9), ws.Cells(endRow, 9))
                .Values = ws.Range(ws.Cells(startRow, 10), ws.Cells(endRow, 10))
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Solid Red
                .Format.Line.Weight = 2.5
                .MarkerStyle = xlMarkerStyleNone
            End With
        End If

        ' --- 4. Axis and Aesthetic Configuration (EIS Standards) ---
        ' X-Axis: Real Impedance
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Z' / Ohm"
            .TickLabelPosition = xlTickLabelPositionHigh
            .MajorTickMark = xlInside
        End With
        
        ' Y-Axis: Negative Imaginary Impedance
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "-Z'' / Ohm"
            .ReversePlotOrder = True ' Conventional inversion for Nyquist plots
            .MajorTickMark = xlInside
            .CrossesAt = 0
        End With

        ' --- 5. Plot Area and Layout Adjustments ---
        DoEvents
        Dim tempTop As Double
        On Error Resume Next
        tempTop = .PlotArea.InsideTop ' Ensure PlotArea object is initialized
        On Error GoTo 0
        
        With .PlotArea
            .Top = 20
            .Left = 60
            .Width = 320
            .Height = 280
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        ' Legend Configuration
        .HasLegend = True
        With .Legend
            .Position = xlLegendPositionBottom
            .Top = chartObj.Height - 30
        End With
        
        ' Final cleanup of labels and area borders
        .Axes(xlCategory).AxisTitle.Top = chartObj.Height - 75
        .ChartArea.Format.Line.Visible = msoFalse
        
        ' Enforce 1:1 Aspect Ratio for geometric accuracy
        Call RescaleAxesToSquare(chartObj.Chart)
    End With
    
End Sub

' ======================================================================================
' PURPOSE:
' Generates a professionally styled DRT Spectrum chart at cell G7.
'
' LOGIC & DATA RANGE:
' 1. Identifies the "Optimal" lambda column (starting from Col P) based on Flag in Col K.
' 2. Maps the frequency grid from Column O to the X-axis (logarithmic).
' 3. Trims high and low frequency edge artifacts using CUT_HIGH_FREQ / CUT_LOW_FREQ.
' 4. Matches the chart dimensions to the existing Nyquist plot for layout consistency.
'
' OUTPUT:
' - An XY Scatter chart (smooth lines) named "DRTSpectrumChart" placed at Cell G7.
' ======================================================================================
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
    
    ' --- 1. Identify the Optimal Lambda Step ---
    ' Search for the "Optimal" flag in Column K (Col 11)
    lastRowStep = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    optimalStep = 0
    For i = 2 To lastRowStep
        If ws.Cells(i, 11).Value = "Optimal" Then
            optimalStep = i - 1
            Exit For
        End If
    Next i
    
    If optimalStep = 0 Then
        MsgBox "Please run the Optimization/Selection macro first.", vbExclamation
        Exit Sub
    End If
    
    ' --- 2. Define Plotting Range (Extracting numbers and cutting edge artifacts) ---
    nTauRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    targetCol = 15 + optimalStep
    
    startTauRow = 0
    endTauRow = 0
    
    ' Find the valid range of numerical frequency data in Column O
    For j = 2 To nTauRow
        If IsNumeric(ws.Cells(j, 15).Value) And ws.Cells(j, 15).Value > 0 Then
            If startTauRow = 0 Then startTauRow = j
            endTauRow = j
        Else
            ' Exit loop if non-numeric labels (like R_inf) are reached
            Exit For
        End If
    Next j
    
    ' Apply Edge Cutting to remove numerical instabilities (Spikes)
    ' High-frequency side (start of data)
    If startTauRow > 0 Then
        startTauRow = startTauRow + CUT_HIGH_FREQ
    End If
    
    ' Low-frequency side (end of data)
    If endTauRow > startTauRow + 10 Then
        endTauRow = endTauRow - CUT_LOW_FREQ
    End If
    
    ' --- 3. Size Reference (Match with Cole-Cole/Nyquist Plot) ---
    ' Attempt to get the size of the first chart on the sheet for uniform layout
    On Error Resume Next
    Set refCht = ws.ChartObjects(1)
    On Error GoTo 0
    
    ' --- 4. Chart Creation / Update ---
    Dim chartName As String: chartName = "DRTSpectrumChart"
    On Error Resume Next
    Set chtObj = ws.ChartObjects(chartName)
    On Error GoTo 0
    
    If chtObj Is Nothing Then
        ' Create new chart if it doesn't exist
        Set chtObj = ws.ChartObjects.Add(Left:=ws.Range("G7").Left, Top:=ws.Range("G7").Top, _
                                         Width:=400, Height:=300)
        chtObj.Name = chartName
    End If
    
    ' Enforce Position and Size
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
        
        ' Remove Chart Title for professional scientific look
        .HasTitle = False
        
        ' Clear existing series and add the DRT Spectrum series
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "DRT Spectrum"
            .XValues = ws.Range(ws.Cells(startTauRow, 15), ws.Cells(endTauRow, 15))
            .Values = ws.Range(ws.Cells(startTauRow, targetCol), ws.Cells(endTauRow, targetCol))
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Solid Red
            .Format.Line.Weight = 2
        End With
        
        ' --- Axis and Formatting Configuration ---
        
        ' X-Axis (Frequency)
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Frequency (Hz)"
            .ScaleType = xlLogarithmic
            ' Move labels to the bottom of the chart
            .TickLabelPosition = xlTickLabelPositionLow
            ' Borders and Gridlines
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        End With
        
        ' Y-Axis (g(tau))
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "g(tau) / Ohm"
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
            .MinimumScale = 0
        End With
        
        ' Plot Area Border
        .PlotArea.Format.Line.Visible = msoTrue
        .PlotArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .PlotArea.Format.Line.Weight = 1
        
        ' Remove Legend for a cleaner look
        .HasLegend = False
    End With
    
    ' MsgBox "DRT Spectrum chart placed at G7, matching the size of the reference plot."
End Sub



' ======================================================================================
' PURPOSE:
' The Main Controller Macro for the DRT Analysis Suite.
'
' EXECUTION STEPS:
' 1. CalculateMagAndPhase: Prepares magnitude and phase data from Z' and Z''.
' 2. RunDRT: Performs the Tikhonov Inversion across multiple regularization parameters.
' 3. FindOptimalLambda_Normalized: Identifies the best lambda using the L-curve "elbow."
' 4. CalculateImpedanceFromOptimalDRT: Reconstructs Z and -Z'' from the resulting model.
' 5. CreateColeColeComparisonChart_Styled_A7: Draws the Nyquist comparison plot.
' 6. CreateDRTSpectrumChart_Styled_G7: Draws the final DRT spectrum g(tau).
' ======================================================================================
Sub ActiveSheetDRT_all()

    ' --- 1. Data Preparation ---
    Call CalculateMagAndPhase
    
    ' --- 2. Numerical Inversion (The heavy lifting) ---
    Call RunDRT
    
    ' --- 3. Parameter Optimization ---
    Call FindOptimalLambda_Normalized
    
    ' --- 4. Model Reconstruction & Verification ---
    Call CalculateImpedanceFromOptimalDRT
    
    ' --- 5. Professional Visualization ---
    Call CreateColeColeComparisonChart_Styled_A7
    Call CreateDRTSpectrumChart_Styled_G7

    ' Final Notification (Optional)
    'MsgBox "DRT Analysis Workflow Completed Successfully.", vbInformation, "Process Finished"

End Sub
