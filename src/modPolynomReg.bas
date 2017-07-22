Attribute VB_Name = "modPolynomReg"

Option Explicit
Option Base 0

'==============================================================================
'requires the 'modArraySupport' module from Chip Pearson available at
'<http://www.cpearson.com/excel/VBAArrays.htm>
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
                "(Optional) TRUE = interpret #NV's as 0's" _
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
                "(Optional) TRUE = ignore #NV entries" _
            )
        .MacroOptions _
            Category:=vCategory, _
            Macro:="PolynomRegRel", _
            Description:="Calculates polynomial coefficients (a0,...,an)", _
            ArgumentDescriptions:=Array( _
                "Array of 'x' values", _
                "Array of 'y' values", _
                "Polynomial order" _
            )
    End With
    
End Sub


Public Function Polynom( _
    Coefficients As Variant, _
    x As Double, _
    Optional NV As Variant _
        ) As Variant
Attribute Polynom.VB_Description = "Calculates polynomial expression f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n"
    
    Dim i As Integer
    Dim sum As Double
    Dim bCoeffRange As Boolean
    
    
    If TypeName(Coefficients) = "Range" Then
        bCoeffRange = True
        'check dimensions of the coefficient matrix
        If Coefficients.Rows.Count >= 1 And Coefficients.Columns.Count = 1 Then
            'everything fine
            Coefficients = Coefficients.Value2
        ElseIf Coefficients.Rows.Count = 1 And Coefficients.Columns.Count >= 1 Then
            'transpose the vector
            Coefficients = Application.WorksheetFunction.Transpose(Coefficients.Value2)
        Else
            'here you haven't specified a scalar or vector but a matrix
            '--> exit the function
            Polynom = "wrong"
            Exit Function
        End If
    End If
    
    'if 'NV' is present and its value is 'TRUE' then replace all trailing 'NVs' with 0s
    On Error Resume Next
    If Not IsMissing(NV) Then
        'this only makes sense if more than one coefficient is given
        If NV = True And IsArray(Coefficients) = True Then
            For i = UBound(Coefficients) To LBound(Coefficients) Step -1
                'if the coefficients are given as "range" than 'Coefficients'
                'is a two-dimensional array
                If bCoeffRange = True Then
                    'if the actual coefficient is the error value 'NA' than
                    'replace it with a zero
                    If Not IsNumeric(Coefficients(i, 1)) Then
                        If Coefficients(i, 1) = CVErr(xlErrNA) Then
                            Coefficients(i, 1) = 0
                        End If
                    Else
                        Exit For
                    End If
                'if function is called from another function then 'Coefficients'
                'is a vector and so just has one dimension
                Else
                    If Not IsNumeric(Coefficients(i)) Then
                        'if the actual coefficient is the error value 'NA' than
                        'replace it with a zero
                        If Coefficients(i) = CVErr(xlErrNA) Then
                            Coefficients(i) = 0
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    On Error GoTo 0
    
    'check, if the coefficients are a scalar or a vector
    'if it is a scalar use the simple form
    If IsArray(Coefficients) = False Then
        sum = Coefficients
    'else the long form
    Else
        If bCoeffRange = True Then
            For i = 1 To UBound(Coefficients)
                If IsNumeric(Coefficients(i, 1)) Then
                    sum = sum + Coefficients(i, 1) * x ^ (i - 1)
                Else
                    Polynom = CVErr(xlErrNum)
                    Exit Function
                End If
            Next
        'also there should be no shift in indexes to exponents so the
        'formula is a bit different than for the "Range" version
        Else
            For i = LBound(Coefficients) To UBound(Coefficients)
                If IsNumeric(Coefficients(i)) Then
                    sum = sum + Coefficients(i) * x ^ (i - LBound(Coefficients))
                Else
                    Polynom = CVErr(xlErrNum)
                    Exit Function
                End If
            Next
        End If
    End If
    
    Polynom = sum
    
End Function


'==============================================================================
' Berechnen der Polynomkoeffizienten a0,..,an für eine polynomiale
' Ausgleichsfunktion n-ten Grades für m Datenpunkte.
' Parameter: x = Array mit x-Werten (Anzahl: m, beliebig)
'            y = Array mit y-Werten (Anzahl: m, beliebig)
'            n = Grad des zu erzeugenden Ausgleichspolynoms
'
' Das Resultat wird als Funktionswert (Arrayfunktion) retourniert
' Autor: Gerhard Krucker
' Datum: 17.08.1995, 22.09.1996
' Sprache: VBA for EXCEL7
' -----
' Stefan Pinnow     20.12.2016
' - revised "Parameterkontrolle"
' - added optional parameter 'IgnoreNVs' so lines with 'NV' values are skipped
' - added optional parameter 'VerticalOutput' which returns the polynomial
'   coefficients vertically if this parameter is 'True'
'
Function PolynomReg( _
    x As Variant, _
    y As Variant, _
    PolynomialDegree As Double, _
    Optional VerticalOutput As Variant, _
    Optional IgnoreNVs As Variant _
        ) As Variant
Attribute PolynomReg.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
    
    'Anzahl x- und y-Werte
    Dim CountX As Integer, CountY As Integer
    'Dynamisches Array fuer die Summe der Potenzen von xk
    Dim Sxk() As Double
    'Dynamisches Array fuer die Summe der Potenzen von xk * yk
    Dim Sxkyk() As Double
    'Dynamisches Array fuer die Koeffizientenmatrix G
    Dim G() As Double, g1 As Variant
    'Dynamisches Array fuer den Konstantenvektor c
    Dim C() As Double
    'Dynamisches Array fuer die Polynomkoeffizienten a0,..,an
    Dim a() As Double
    'running variables
    Dim i As Integer, j As Integer
    
    'dynamic arrays to store given 'x' and 'y' data as vectors (instead of arrays)
    Dim xx() As Variant, yy() As Variant
    'dynamic arrays to store given data revised by 'NV' data
    Dim xxx() As Double, yyy() As Double
    
    
    '---
    'Parameterkontrolle
    '---
    ''Polynomgrad' muss eine ganze Zahl >= 1 sein
    If Not IsNumeric(PolynomialDegree) Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    ElseIf CInt(PolynomialDegree) <> PolynomialDegree Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    ElseIf PolynomialDegree < 1 Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    End If
    
    ''VerticalOutput' must be an integer (or boolean)
    If IsMissing(VerticalOutput) Then
        VerticalOutput = False
    ElseIf Not IsNumeric(VerticalOutput) Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    ElseIf CInt(VerticalOutput) <> VerticalOutput Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    End If
    
    ''IgnoreNVs' must be an integer (or boolean)
    If IsMissing(IgnoreNVs) Then
        IgnoreNVs = False
    ElseIf Not IsNumeric(IgnoreNVs) Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    ElseIf CInt(IgnoreNVs) <> IgnoreNVs Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    End If
    '---
    
    'count number of data points in given arrays
    If TypeName(x) = "Range" Then
        CountX = x.Count
    Else
        CountX = UBound(x) - LBound(x) + 1
    End If
    If TypeName(y) = "Range" Then
        CountY = y.Count
    Else
        CountY = UBound(y) - LBound(y) + 1
    End If
    
    'the polynomial coefficient must be smaller than the number of given (x) points
    If CountX <= PolynomialDegree Then
        PolynomReg = CVErr(xlErrValue)
        Exit Function
    End If
    '---
    
    'convert 'x' and 'y' to arrays (in case they are ranges)
    x = x
    y = y
    
    
    'if 'NV' is False copy 'x' to 'xxx', y to 'yyy' and set 'j' to 'CountX'
    If IgnoreNVs = False Then
        If Not ExtractVectors(x, xxx) Then
            PolynomReg = CVErr(xlErrValue)
        End If
        If Not ExtractVectors(y, yyy) Then
            PolynomReg = CVErr(xlErrValue)
        End If
        j = CountX
    Else
        'else copy 'x' to 'xx' and 'y' to 'yy'
        If Not ExtractVectors(x, xx) Then
            PolynomReg = CVErr(xlErrValue)
        End If
        If Not ExtractVectors(y, yy) Then
            PolynomReg = CVErr(xlErrValue)
        End If
        
        'instantiate 'xx' and 'yy'
        ReDim xxx(CountX)
        ReDim yyy(CountY)
        
        'cycle through each entry
        For i = LBound(xx) To UBound(xx)
            'if both values are of numeric type then transfer them to 'xxx' and 'yyy'
            If IsNumeric(xx(i)) And IsNumeric(yy(i)) Then
                j = j + 1
                xxx(j) = xx(i)
                yyy(j) = yy(i)
            'if not, it is allowed that the values are of the error type 'NA'
            ElseIf Application.WorksheetFunction.IsNA(xx(i)) Or _
                    Application.WorksheetFunction.IsNA(yy(i)) Then
            'else at least one of the 'xx' or 'yy' points contains a not allowed value
            Else
                PolynomReg = CVErr(xlErrValue)
                Exit Function
            End If
        Next
        'check again, if the polynomial coefficient is smaller than the number of
        'given (x) points
        If j <= PolynomialDegree Then
            PolynomReg = CVErr(xlErrValue)
            Exit Function
        End If
    End If
    
    'transfer (new) number of x elements to 'CountX'
    CountX = j
    
    'Summe der Potenzen xk und xk*yk berechnen und den entsprechende Arrays abspeichern
    ReDim Sxk(PolynomialDegree * 2)     'Arrays auf passende Groesse dimensionieren
    ReDim Sxkyk(PolynomialDegree)       'Die Arrayindizes laufen von 0..PolynomialDegree, resp 0..2*PolynomialDegree
    For i = 0 To 2 * PolynomialDegree
        Sxk(i) = 0
        For j = 1 To CountX             'Für jeden Datenpunkt
            Sxk(i) = Sxk(i) + xxx(j) ^ i
        Next j
    Next i
    For i = 0 To PolynomialDegree
        Sxkyk(i) = 0
        For j = 1 To CountX
            Sxkyk(i) = Sxkyk(i) + xxx(j) ^ i * yyy(j)
        Next j
    Next i
    
    'Koeffizientenmatrix G und Konstantenvektor c erzeugen
    ReDim G(1 To PolynomialDegree + 1, 1 To PolynomialDegree + 1)       'Matrix mit Indizes 0..PolynomialDegree,0..PolynomialDegree dimensionieren
    ReDim g1(1 To PolynomialDegree + 1, 1 To PolynomialDegree + 1)      'Matrix für die Inverse von G (MINV kann nicht in G zurueckschreiben)
    ReDim C(1 To PolynomialDegree + 1)
    ReDim a(0 To PolynomialDegree)                      'Polynomkoeffizienten a0,..,an (a(0) = a0)
    
    'Koeffizientenmatrix G und Konstantenvektor c aufbauen
    For i = 0 To PolynomialDegree
        For j = 0 To i
            G(i + 1, j + 1) = Sxk(i + j)
            G(j + 1, i + 1) = Sxk(i + j)
        Next j
        C(i + 1) = Sxkyk(i)
    Next i
    
    'Gleichungssystem G * a = c lösen mit Matrixinversion
    g1 = Application.WorksheetFunction.MInverse(G)      'Koeffizientenmatrix G invertieren
    For i = 1 To PolynomialDegree + 1                   'Matrixmultiplikation a = G1 * c
        a(i - 1) = 0
        For j = 1 To PolynomialDegree + 1
            a(i - 1) = a(i - 1) + g1(i, j) * C(j)
        Next j
    Next i
    
    'return coefficient vector a_0,...,a_n
    If VerticalOutput = True Then
        PolynomReg = Application.WorksheetFunction.Transpose(a)
    Else
        PolynomReg = a
    End If
    
End Function


'Berechnen der Polynomkoeffizienten a0,..,an für eine polynomiale Ausgleichsfunktion
'n-ten Grades für m Datenpunkte nach der Methode der kleinsten Summe der relativen Fehlerquadrate.
'Parameter: x = Array mit x-Werten (Anzahl: m, beliebig)
' y = Array mit y-Werten (Anzahl: m, beliebig)
' n = Grad des zu erzeugenden Ausgleichspolynoms
'
' Das Resultat wird als Funktionswert (Arrayfunktion) retourniert
' Autor: Gerhard Krucker
' Datum: 17.8.1995, 22. 9. 1996, 24.9.2004
' Sprache: VBA for EXCEL7, EXCEL XP
'
Function PolynomRegRel(x As Variant, y As Variant, Polynomgrad As Integer) As Variant
Attribute PolynomRegRel.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
    Dim AnzX, AnzY      'Anzahl x- und y-Werte
    Dim m               'Anzahl auszugleichender Datenpunkte
    Dim Sxk()           'Dynamisches Array für die Summe der Potenzen von xk
    Dim Sxkyk()         'Dynamisches Array für die Summe der Potenzen von xk * yk
    Dim G(), g1         'Dynamisches Array für die Koeffizientenmatrix G
    Dim C()             'Dynamisches Array für den Konstantenvektor c
    Dim a()             'Dynamisches Array für die Polynomkoeffizienten a0,..,an
    Dim i, j, k
    
    ' Parameterkontrollen
    If (Polynomgrad < 1) And Polynomgrad <> "Integer" Then
        MsgBox "Polynomgrad muss eine Ganzzahl >= 1 sein!"
        PolynomRegRel = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Anzahl Datenpunkte in den Arrays bestimmen
    If TypeName(x) = "Range" Then
        AnzX = x.Count
    Else
        AnzX = UBound(x) - LBound(x) + 1
    End If
    If TypeName(y) = "Range" Then
        AnzY = y.Count
    Else
        AnzY = UBound(y) - LBound(y) + 1
    End If
    
    If (AnzX <> AnzY) Then
        MsgBox "Anzahl x-Werte und Anzahl y-Werte muss gleich gross sein"
        PolynomRegRel = CVErr(xlErrValue)
        Exit Function
    End If
    If AnzX <= Polynomgrad Then
        MsgBox "Anzahl Datenpunkte muss > Polynomgrad sein!"
        PolynomRegRel = CVErr(xlErrValue)
        Exit Function
    End If
    m = AnzX
    
    'Summe der Potenzen xk/yk^2 und xk*yk berechnen und in den entsprechende Arrays abspeichern
    ReDim Sxk(Polynomgrad * 2)      'Arrays auf passende Größe dimensionieren
    ReDim Sxkyk(Polynomgrad)        'Die Arrayindizes laufen von 0..Polynomgrad, resp 0..2*Polynomgrad
    For i = 0 To 2 * Polynomgrad
        Sxk(i) = 0
        For k = 1 To m              'für jeden Datenpunkt
            Sxk(i) = Sxk(i) + x(k) ^ i / y(k) ^ 2
        Next k
    Next i
    For i = 0 To Polynomgrad
        Sxkyk(i) = 0
        For k = 1 To m
            Sxkyk(i) = Sxkyk(i) + x(k) ^ i / y(k)
        Next k
    Next i
    
    'Koeffizientenmatrix G und Konstantenvektor c erzeugen
    ReDim G(1 To Polynomgrad + 1, 1 To Polynomgrad + 1)     'Matrix mit Indizes 0..Polynomgrad,0..Polynomgrad dimensionieren
    ReDim g1(1 To Polynomgrad + 1, 1 To Polynomgrad + 1)    'Matrix für die Inverse von G (MINV kann nicht in G zurückschreiben)
    ReDim C(1 To Polynomgrad + 1)
    ReDim a(0 To Polynomgrad)       'Polynomkoeffizienten a0,..,an (a(0) = a0)
    
    For i = 0 To Polynomgrad        'Koeffizientenmatrix G und Konstantenvektor c aufbauen
        For j = 0 To i
            G(i + 1, j + 1) = Sxk(i + j)
            G(j + 1, i + 1) = Sxk(i + j)
        Next j
        C(i + 1) = Sxkyk(i)
    Next i
    
    ' Gleichungssystem G * a = c lösen mit Matrixinversion
    g1 = Application.MInverse(G)    'Koeffizientenmatrix G invertieren
    For i = 1 To Polynomgrad + 1    'Matrixmultiplikation a = G1 * c
        a(i - 1) = 0
        For j = 1 To Polynomgrad + 1
            a(i - 1) = a(i - 1) + g1(i, j) * C(j)
        Next j
    Next i
    
    PolynomRegRel = a               'Koeffizientenvektor a0,..,an retournieren
    
End Function


'==============================================================================
'function to make vectors of the ranges/arrays and optionally only transfer
'non-NV values
Private Function ExtractVectors( _
    a As Variant, _
    b As Variant _
        ) As Boolean
        
    If NumberOfArrayDimensions(a) = 1 Then
        If Not CopyArray(b, a, True) Then Exit Function
        
'check if right
        Dim N As Integer
        N = UBound(b) - LBound(b) + 1
        
        If Not ChangeBoundsOfArray(b, 1, N) Then Exit Function
    Else
        If UBound(a, 1) > 1 And UBound(a, 2) = 1 Then
            If Not GetColumn(a, b, 1) Then Exit Function
        ElseIf UBound(a, 1) = 1 And UBound(a, 2) > 1 Then
            If Not GetRow(a, b, 1) Then Exit Function
        Else
            Exit Function
        End If
    End If
    
    ExtractVectors = True
    
End Function
