Attribute VB_Name = "modPolynomRegTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As New Rubberduck.AssertClass
Private vXData As Variant
Private vYData As Variant
Private vCoeffs As Variant
Private aExpected As Variant
Private aActual As Variant

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'==============================================================================
'unit tests for `PolynomReg'
'==============================================================================

'@TestMethod
Public Sub PolynomReg_OnlyNAValues_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 1)
    vYData = Array(CVErr(xlErrNA), CVErr(xlErrNA))
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 0)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_OnlyInvalidValues_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 1)
    vYData = Array("s", CVErr(xlErrNum))
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 0)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_OnlyInvalidValuesWithNA_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 1)
    vYData = Array("s", CVErr(xlErrNum))
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 0, , True)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_Order0_ReturnsValue()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 1)
    vYData = Array(2, 3)
    aExpected = Array(2.5, CVErr(xlErrNA))
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 0)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected(0)), CDbl(aActual(0))
    Assert.AreEqual CDbl(aExpected(1)), CDbl(aActual(1))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_Order1_ReturnsArrayOfValues()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 1)
    vYData = Array(2, 3)
    aExpected = Array(2, 1)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 1)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected(0)), CDbl(aActual(0))
    Assert.AreEqual CDbl(aExpected(1)), CDbl(aActual(1))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_Order1WithNA_ReturnsArrayOfValues()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(0, 0.5, 1)
    vYData = Array(2, CVErr(xlErrNA), 3)
    aExpected = Array(2, 1)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 1, , True)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected(0)), CDbl(aActual(0))
    Assert.AreEqual CDbl(aExpected(1)), CDbl(aActual(1))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub PolynomReg_Order2_ReturnsArrayOfValues()
    On Error GoTo TestFail
    
    'Arrange:
    vXData = Array(-1, 1, 2)
    vYData = Array(-2, 2, 1)
    aExpected = Array(1, 2, -1)
    
    'Act:
    aActual = modPolynomReg.PolynomReg(vXData, vYData, 2)
    
    'Assert:
    'TODO: for whatever reason the following results 'FALSE'
'    Assert.AreEqual CDbl(aExpected(0)), CDbl(aActual(0))
    Assert.IsTrue Abs(aExpected(0) - aActual(0)) < 0.000000001
    Assert.AreEqual CDbl(aExpected(1)), CDbl(aActual(1))
    Assert.AreEqual CDbl(aExpected(2)), CDbl(aActual(2))
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'==============================================================================
'unit tests for `Polynom'
'==============================================================================

'@TestMethod
Public Sub Polynom_OnlyNACoeff_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = CVErr(xlErrNA)
    vXData = 5
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_OnlyNACoeffWithNA_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = CVErr(xlErrNA)
    vXData = 5
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData, True)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_OnlyNumError_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = CVErr(xlErrNum)
    vXData = 5
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData)
    
    'Assert:
    Assert.AreEqual aExpected, aActual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_1ValidValue_ReturnsValue()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = 15
    vXData = 5
    aExpected = vCoeffs
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected), CDbl(aActual)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_2ValidValues_ReturnsValue()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = Array(-5, 0.5)
    vXData = 5
    aExpected = -2.5
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected), CDbl(aActual)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_2ValidValuesPlusNA_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = Array(-5, 0.5, CVErr(xlErrNA))
    vXData = 5
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected), CDbl(aActual)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_2ValidValuesPlusNAWithNA_ReturnsValue()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = Array(-5, 0.5, CVErr(xlErrNA))
    vXData = 5
    aExpected = -2.5
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData, True)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected), CDbl(aActual)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Polynom_2ValidValuesPlusNAInBetweenWithNA_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    vCoeffs = Array(-5, CVErr(xlErrNA), 0.5)
    vXData = 5
    aExpected = CVErr(xlErrNA)
    
    'Act:
    aActual = modPolynomReg.Polynom(vCoeffs, vXData, True)
    
    'Assert:
    Assert.AreEqual CDbl(aExpected), CDbl(aActual)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub