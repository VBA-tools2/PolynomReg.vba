Attribute VB_Name = "modUsefulFunctions"

Option Explicit
Option Base 1


'==============================================================================
'Returns the variable type of the given parameter
'if it is a range, it will check the upper left cell in that range
'(inspired by ...)
'<http://spreadsheetpage.com/index.php/tip/determining_the_data_type_of_a_cell/>
'<https://stackoverflow.com/a/1994169>
Function VariableType(c As Variant) As Variant
'    Application.Volatile
    
    If TypeName(c) = "Range" Then
        Set c = c.Range("A1")
    End If
    
    Select Case True
        Case IsEmpty(c)
            VariableType = "Empty"   'vbEmpty
        Case Application.WorksheetFunction.IsText(c)
            VariableType = "String"  'vbString
        Case Application.WorksheetFunction.IsLogical(c)
            VariableType = "Boolean" 'vbBoolean
        Case Application.WorksheetFunction.IsError(c)
            VariableType = "Error"   'vbError
        Case IsDate(c)
            VariableType = "Date"    'vbDate
'        Case InStr(1, c.text, ":") <> 0
'            VariableType = "Time"
        Case IsNumeric(c)
            If c = Int(c) Then
                VariableType = "Integer"
            Else
                VariableType = "Double"
            End If
        Case IsObject(c)
            VariableType = "Object"
        Case IsArray(c)
            VariableType = "Array"
        Case Else
            Select Case VarType(c)
                Case vbCurrency
                Case vbObject
                    VariableType = "Object"
                Case vbVariant
                Case vbDataObject
                Case vbUserDefinedType
                Case vbArray
                    VariableType = "Array"
                Case Else
            End Select
    End Select
End Function
