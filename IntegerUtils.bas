Attribute VB_Name = "IntegerUtils"
'---------------------------------------------------------------------------------------
' Module    : IntegerUtils
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Integer related functions
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : IntTryParse
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Parse a string representation of an integer as an integer datatype
'---------------------------------------------------------------------------------------
'
Public Function IntTryParse(ByVal sValue As String, ByRef iOutput As Integer, Optional ByVal iDefault As Integer = 0) As Boolean
    iOutput = iDefault ' Set the output to the default zero

    Dim bIsInt As Boolean
    bIsInt = StrIsInt(sValue)
    If bIsInt Then
        iOutput = CInt(sValue) ' Parse as an integer
    End If
    IntTryParse = bIsInt
End Function
