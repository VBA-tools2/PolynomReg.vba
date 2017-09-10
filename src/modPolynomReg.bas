Attribute VB_Name = "modPolynomReg"

Option Explicit

'module originally from Gerhard Krucker which was created 17.08.1995, 22.09.1996
'and 2004 for VBA in EXCEL7. It was highly modified by Stefan Pinnow in 2016
'and 2017.

'==============================================================================
'requires the 'modArraySupport' module from Chip Pearson available at
'<http://www.cpearson.com/excel/VBAArrays.htm>
'and the function 'modUsefulFunctions.VariableType'
'==============================================================================

Sub AddUDFToCustomCategory()
    
    '==========================================================================
    'how should the category be named?
'    Const vCategory As String = "Math. & Trigonom."
    Const vCategory As Integer = 3   '"Math. & Trigonom."
    '==========================================================================
    
    With Application
        .MacroOptions _
            Category:=vCategory, _
            Macro:="Polynom", _
            Description:="Calculates polynomial expression " & _
                "f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n", _
            ArgumentDescriptions:=Array( _
                "Coefficients (a0, a1, a2, ...)", _
                "Independent variable (x)", _
                "(Optional) TRUE = interpret #NA's as 0's" _
            )
        .MacroOptions _
            Category:=vCategory, _
            Macro:="PolynomReg", _
            Description:="Calculates polynomial coefficients (a0,...,an)", _
            ArgumentDescriptions:=Array( _
                "Array of 'x' values", _
                "Array of 'y' values", _
                "Polynomial degree", _
                "(Optional) TRUE = return coefficients vertically", _
                "(Optional) TRUE = ignore #NA entries" _
            )
        .MacroOptions _
            Category:=vCategory, _
            Macro:="PolynomRegRel", _
            Description:="Calculates polynomial coefficients (a0,...,an)", _
            ArgumentDescriptions:=Array( _
                "Array of 'x' values", _
                "Array of 'y' values", _
                "Polynomial degree", _
                "(Optional) TRUE = return coefficients vertically", _
                "(Optional) TRUE = ignore #NA entries" _
            )
    End With
    
End Sub


Public Function Polynom( _
    Coefficients As Variant, _
    x As Double, _
    Optional NA As Variant _
        ) As Variant
Attribute Polynom.VB_Description = "Calculates polynomial expression f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n"
    
    Dim i As Integer
    Dim sum As Double
    Dim arrCoeffs() As Variant
    
    
    'convert possible range to array
    Coefficients = Coefficients
    
    If Not ExtractVector(Coefficients, arrCoeffs) Then GoTo errHandler
    
    'if 'NA' is present and its value is 'TRUE' then remove all trailing 'NAs' lines
    If Not IsMissing(NA) Then
        'this only makes sense if more than one coefficient is given
        If NA = True Then
            If Not RemoveNALines(arrCoeffs) Then GoTo errHandler
        End If
    End If
    
    'check, if the coefficients are a scalar or a vector
    'if it is a scalar use the simple form
    For i = LBound(arrCoeffs) To UBound(arrCoeffs)
        sum = sum + arrCoeffs(i) * x ^ (i - LBound(arrCoeffs))
    Next
    
    'return the result
    Polynom = sum
    Exit Function
    
    
errHandler:
    Polynom = CVErr(xlErrNA)
    
End Function


'==============================================================================
'calculate the polynomial coefficients a0,...,an for the polynomial trend
'function n-th degree for m data points using the method of least squares.
'Parameter:
'- x                = array of x values (number of points: m, any)
'- y                = array of y values (number of points: m, any)
'- PolynomialDegree = degree of to generate polynomial trend function
'- VerticalOutput   = optional argument to allow a vertical output of the
'                     polynomial coefficients
'- IgnoreNAs        = optional argument to ignore "NA" data points
'The result will be returned as array (vector)
Function PolynomReg( _
    x As Variant, _
    y As Variant, _
    PolynomialDegree As Integer, _
    Optional VerticalOutput As Variant, _
    Optional IgnoreNAs As Variant _
        ) As Variant
Attribute PolynomReg.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
    
    '---
    ''VerticalOutput' must be a boolean
    If IsMissing(VerticalOutput) Or IsEmpty(VerticalOutput) Then
        VerticalOutput = False
    ElseIf Not VariableType(VerticalOutput) = "Boolean" Then
        GoTo errHandler
    End If
    
    ''IgnoreNAs' must be a boolean
    If IsMissing(IgnoreNAs) Or IsEmpty(IgnoreNAs) Then
        IgnoreNAs = False
    ElseIf Not VariableType(IgnoreNAs) = "Boolean" Then
        GoTo errHandler
    End If
    '---
    
    
    PolynomReg = MasterPolynomReg( _
            x, y, _
            PolynomialDegree, _
            CBool(VerticalOutput), _
            CBool(IgnoreNAs), _
            False _
    )
    Exit Function
    
    
errHandler:
    PolynomReg = CVErr(xlErrNA)
    
End Function


'calculate the polynomial coefficients a0,...,an for the polynomial trend
'function n-th degree for m data points using the method of least relative
'squares.
'Parameter:
'- x                = array of x values (number of points: m, any)
'- y                = array of y values (number of points: m, any)
'- PolynomialDegree = degree of to generate polynomial trend function
'- VerticalOutput   = optional argument to allow a vertical output of the
'                     polynomial coefficients
'- IgnoreNAs        = optional argument to ignore "NA" data points
'The result will be returned as array (vector)
Function PolynomRegRel( _
    x As Variant, _
    y As Variant, _
    PolynomialDegree As Integer, _
    Optional VerticalOutput As Variant, _
    Optional IgnoreNAs As Variant _
        ) As Variant
Attribute PolynomRegRel.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
Attribute PolynomRegRel.VB_ProcData.VB_Invoke_Func = " \n3"
    
    '---
    ''VerticalOutput' must be a boolean
    If IsMissing(VerticalOutput) Or IsEmpty(VerticalOutput) Then
        VerticalOutput = False
    ElseIf Not VariableType(VerticalOutput) = "Boolean" Then
        GoTo errHandler
    End If
    
    ''IgnoreNAs' must be a boolean
    If IsMissing(IgnoreNAs) Or IsEmpty(IgnoreNAs) Then
        IgnoreNAs = False
    ElseIf Not VariableType(IgnoreNAs) = "Boolean" Then
        GoTo errHandler
    End If
    '---
    
    
    PolynomRegRel = MasterPolynomReg( _
            x, y, _
            PolynomialDegree, _
            CBool(VerticalOutput), _
            CBool(IgnoreNAs), _
            True _
    )
    Exit Function
    
    
errHandler:
    PolynomRegRel = CVErr(xlErrNA)
    
End Function


Private Function MasterPolynomReg( _
    x As Variant, _
    y As Variant, _
    PolynomialDegree As Integer, _
    VerticalOutput As Boolean, _
    IgnoreNAs As Boolean, _
    UseRelativeVersion As Boolean _
        ) As Variant
    
    'amount of 'x' and 'y' values
    Dim CountX As Integer, CountY As Integer
    Dim NoOfNonNADataPoints As Integer
    Dim CoefficientMatrix() As Double
    'variable for the inverse of 'CoefficientMatrix'
    '(it has to be a variant to be able to use 'WorksheetFunction.MInverse')
    Dim InverseCoefficientMatrix As Variant
    Dim VectorOfConstants() As Double
    'dynamic array for the polynomial coefficients a0,...,an
    '(it has to be of type Variant' because of the special handler for
    ' 'PolynomialDegree = 0')
    Dim a() As Variant
    
    'dynamic arrays to store given 'x' and 'y' data as vectors (instead of arrays)
    Dim xAsVector() As Variant, yAsVector() As Variant
    'dynamic arrays to store given data revised by 'NA' data
    Dim xWithoutNAs() As Double, yWithoutNAs() As Double
    
    
    '---
    ''PolynomialDegree' has to be an integer >= 0
    If PolynomialDegree < 0 Then
        GoTo errHandler
    End If
    '---
    
    
    'convert 'x' and 'y' to arrays (in case they are ranges)
    x = x
    y = y
    
    'count number of data points in given arrays
    CountX = UBound(x) - LBound(x) + 1
    CountY = UBound(y) - LBound(y) + 1
    
    'the number of points has to be identical for 'x' and 'y'
    If CountX <> CountY Then GoTo errHandler
    
    'the polynomial coefficient must be smaller than the number of given points
    If CountX <= PolynomialDegree Then GoTo errHandler
    
    
    'prepare vectors 'xWithoutNAs' and 'yWithoutNAs'
    '(which are then used to calculate the polynomial coefficients)
    If IgnoreNAs = False Then
        If Not ExtractVector(x, xWithoutNAs) Then GoTo errHandler
        If Not ExtractVector(y, yWithoutNAs) Then GoTo errHandler
    Else
        'else copy 'x' to 'xAsVector' and 'y' to 'yAsVector'
        If Not ExtractVector(x, xAsVector) Then GoTo errHandler
        If Not ExtractVector(y, yAsVector) Then GoTo errHandler
        
        If Not CopyOnlyNonNALines( _
                xAsVector, yAsVector, _
                xWithoutNAs, yWithoutNAs, _
                PolynomialDegree _
        ) Then GoTo errHandler
    End If
    
    '--------------------------------------------------------------------------
    
    'transfer (new) number of 'x' elements to 'NoOfNonNADataPoints'
    NoOfNonNADataPoints = UBound(xWithoutNAs) - LBound(xWithoutNAs) + 1
    'check again, if number of (real) data points is smaller than the given
    'polynomial degree
    If NoOfNonNADataPoints <= PolynomialDegree Then GoTo errHandler
    
    CoefficientMatrix = Calculate_CoefficientMatrix( _
            xWithoutNAs, yWithoutNAs, _
            PolynomialDegree, _
            UseRelativeVersion _
    )
    VectorOfConstants = Calculate_VectorOfConstants( _
            xWithoutNAs, yWithoutNAs, _
            PolynomialDegree _
    )
    
    'invert coefficient matrix 'CoefficientMatrix'
    '(MINVERSE can't write back to 'CoefficientMatrix')
    '(please note that the resulting array starts with index 1)
    InverseCoefficientMatrix = Application.WorksheetFunction.MInverse(CoefficientMatrix)
    
    a = Calculate_PolynomialCoefficients( _
            InverseCoefficientMatrix, _
            VectorOfConstants, _
            PolynomialDegree _
    )
    
    If PolynomialDegree = 0 Then
        Call HandleSpecialCaseForPolynomialDegreeEqualsZero(a)
    End If
    
    'return coefficient vector a_0,...,a_n
    If VerticalOutput = True Then
        MasterPolynomReg = Application.WorksheetFunction.Transpose(a)
    Else
        MasterPolynomReg = a
    End If
    
    Exit Function
    
    
errHandler:
    MasterPolynomReg = CVErr(xlErrNA)
    
End Function


'==============================================================================
Private Function Calculate_CoefficientMatrix( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Integer, _
    ByVal UseRelativeVersion As Boolean _
        ) As Double()
    
    Dim i As Integer
    Dim j As Integer
    Dim SumOfPowersXK() As Double
    Dim CoefficientMatrix() As Double
    
    
    SumOfPowersXK = Calculate_SumOfPowersXK( _
            x, y, _
            PolynomialDegree, _
            UseRelativeVersion _
    )
    
    ReDim CoefficientMatrix(0 To PolynomialDegree, 0 To PolynomialDegree)
    For i = LBound(CoefficientMatrix, 1) To UBound(CoefficientMatrix, 1)
        For j = LBound(CoefficientMatrix, 2) To i
            CoefficientMatrix(i, j) = SumOfPowersXK(i + j)
            CoefficientMatrix(j, i) = SumOfPowersXK(i + j)
        Next
    Next
    
    Calculate_CoefficientMatrix = CoefficientMatrix
    
End Function


'calculate sum of powers 'xk' and store it in a corresponding array
Private Function Calculate_SumOfPowersXK( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Integer, _
    ByVal UseRelativeVersion As Boolean _
        ) As Double()
    
    Dim i As Integer
    Dim k As Integer
    Dim SumOfPowersXK() As Double
    
    
    ReDim SumOfPowersXK(0 To 2 * PolynomialDegree)
    If UseRelativeVersion = True Then
        For i = LBound(SumOfPowersXK) To UBound(SumOfPowersXK)
            SumOfPowersXK(i) = 0
            For k = LBound(x) To UBound(x)
                SumOfPowersXK(i) = SumOfPowersXK(i) + x(k) ^ i / y(k) ^ 2
            Next
        Next
    Else
        For i = LBound(SumOfPowersXK) To UBound(SumOfPowersXK)
            SumOfPowersXK(i) = 0
            For k = LBound(x) To UBound(x)
                SumOfPowersXK(i) = SumOfPowersXK(i) + x(k) ^ i
            Next
        Next
    End If
    
    Calculate_SumOfPowersXK = SumOfPowersXK
    
End Function


Private Function Calculate_VectorOfConstants( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Integer _
        ) As Double()
    
    Dim i As Integer
    Dim k As Integer
    'dynamic array for the sum of powers for 'xk*yk'
    Dim SumOfPowersXKYK() As Double
    
    
    'calculate sum of powers 'xk*yk' and store it in a corresponding array
    ReDim SumOfPowersXKYK(0 To PolynomialDegree)
    For i = LBound(SumOfPowersXKYK) To UBound(SumOfPowersXKYK)
        SumOfPowersXKYK(i) = 0
        For k = LBound(x) To UBound(x)
            SumOfPowersXKYK(i) = SumOfPowersXKYK(i) + x(k) ^ i * y(k)
        Next
    Next
    
    'the sum of powers 'xk*yk' is the vector of constants
    Calculate_VectorOfConstants = SumOfPowersXKYK
    
End Function


'solve system of equations 'CoefficientMatrix * a = c' with matrix inversion
Private Function Calculate_PolynomialCoefficients( _
    ByRef InverseCoefficientMatrix As Variant, _
    ByRef VectorOfConstants() As Double, _
    ByVal PolynomialDegree As Integer _
        ) As Variant
    
    Dim i As Integer
    Dim j As Integer
    Dim a() As Variant
    
    
    'polynomial coefficients a0,...,an (a(0) = a0)
    ReDim a(0 To PolynomialDegree)
    
    'matrix multiplication 'a = G_inverse * VectorOfConstants'
'---
'<https://stackoverflow.com/a/7307992>
'   a = WorksheetFunction.MMult(InverseCoefficientMatrix, VectorOfConstants)
'---
    'as a reminder: 'InverseCoefficientMatrix' is a 1-based array
    'which is coming from 'WorksheetFunction.MInverse'
    'it is also needed a special handler for 'PolynomialDegree = 0' because of
    'an anomaly (bug?) in Excel, where a 1D array (z(1 to ..., 1 to 1)) is
    'returned as vector (z(1 to ...)) after inversion with 'WorksheetFunction.MInverse'
    '(see <https://stackoverflow.com/a/28800474/5776000>)
    If PolynomialDegree = 0 Then
        a(0) = InverseCoefficientMatrix(1) * VectorOfConstants(0)
    Else
        For i = LBound(a) To UBound(a)
            a(i) = 0
            For j = LBound(a) To UBound(a)
                a(i) = a(i) + InverseCoefficientMatrix(i + 1, j + 1) * VectorOfConstants(j)
            Next
        Next
    End If
    
    Calculate_PolynomialCoefficients = a
    
End Function


'needed because otherwise the returned array will consist of the a(0) value in
'*all* cells (and not '#NA' values for the "unused coefficients)
Private Sub HandleSpecialCaseForPolynomialDegreeEqualsZero( _
    a() As Variant _
)
    
    ReDim Preserve a(0 To 1)
    a(1) = CVErr(xlErrNA)
    
End Sub

'function to make vectors of the ranges/arrays and optionally only transfer
'non-NA values
Private Function ExtractVector( _
    Source As Variant, _
    DestVector As Variant _
        ) As Boolean
    
    Dim N As Integer
    
    
    Select Case NumberOfArrayDimensions(Source)
        Case 2
            If UBound(Source, 1) > 1 And UBound(Source, 2) = 1 Then
                If Not GetColumn(Source, DestVector, 1) Then Exit Function
            ElseIf UBound(Source, 1) = 1 And UBound(Source, 2) > 1 Then
                If Not GetRow(Source, DestVector, 1) Then Exit Function
            Else
                Exit Function
            End If
        Case 1
            If Not CopyArray(Source, DestVector, True) Then Exit Function
            N = UBound(DestVector) - LBound(DestVector) + 1
            If Not ChangeBoundsOfArray(DestVector, 1, N) Then Exit Function
        Case 0
            ReDim DestVector(0 To 0)
            DestVector(0) = Source
        Case Else
    End Select
    
    ExtractVector = True
    
End Function


Private Function CopyOnlyNonNALines( _
    ByRef xSource As Variant, _
    ByRef ySource As Variant, _
    ByRef xDest As Variant, _
    ByRef yDest As Variant, _
    ByVal PolynomialDegree As Integer _
        ) As Boolean
    
    Dim i As Long
    Dim j As Long
    
    
    'instantiate 'xDest' and 'yDest'
    ReDim xDest(1 To UBound(xSource) - LBound(xSource) + 1)
    ReDim yDest(1 To UBound(xSource) - LBound(xSource) + 1)
    
    'cycle through each entry
    For i = LBound(xSource) To UBound(xSource)
        'if both values are of numeric type then transfer them to 'xDest' and 'yDest'
        If IsNumeric(xSource(i)) And IsNumeric(ySource(i)) Then
            j = j + 1
            xDest(j) = xSource(i)
            yDest(j) = ySource(i)
        'if not, it is allowed that the values are of the error type 'NA'
        ElseIf Application.WorksheetFunction.IsNA(xSource(i)) Or _
                Application.WorksheetFunction.IsNA(ySource(i)) Then
        'else at least one of the 'xSource' or 'ySource' points contains a
        'not allowed value
        Else
            CopyOnlyNonNALines = False
            Exit Function
        End If
    Next
    
    'redim 'xDest' and 'yDest' to only the populated values
    ReDim Preserve xDest(1 To j)
    ReDim Preserve yDest(1 To j)
    
    
    'check again, if the polynomial coefficient is smaller than the number of
    'given points
    If j > PolynomialDegree Then
        CopyOnlyNonNALines = True
    End If
    
End Function


Private Function RemoveNALines(Arr As Variant) As Boolean
    
    Dim i As Integer
    
    
    For i = UBound(Arr) To LBound(Arr) Step -1
        'if the actual coefficient is not a number ...
        If Not IsNumeric(Arr(i)) Then
            '... and not the error value 'NA' then exit the function
            If Arr(i) <> CVErr(xlErrNA) Then
                RemoveNALines = False
                Exit Function
            End If
        'if it is a number stop cycling
        '(because then only numbers will follow)
        Else
            Exit For
        End If
    Next
    
    'redim 'Arr' to the numeric values only
    ReDim Preserve Arr(LBound(Arr) To i)
    
    RemoveNALines = True
    
End Function
