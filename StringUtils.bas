Attribute VB_Name = "StringUtils"
'---------------------------------------------------------------------------------------
' Module    : StringUtils
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : String related functions such as if a string contains digits or is an integer
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : StrIsAlNum
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Checks if a string contains only alphanumeric characters ( 0-9 and A-Z ).
'---------------------------------------------------------------------------------------
'
Public Function StrIsAlNum(ByVal sValue As String) As Boolean
    StrIsAlNum = Internal_StrRegExp(sValue, "^[\dA-Fa-f]+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsAlpha
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Checks if a string contains only alphabetic characters ( A-Z )
'---------------------------------------------------------------------------------------
'
Public Function StrIsAlpha(ByVal sValue As String) As Boolean
    StrIsAlpha = Internal_StrRegExp(sValue, "^[A-Fa-f]+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsASCII
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Checks if a string contains only ASCII characters ( 0-127 )
'---------------------------------------------------------------------------------------
'
Public Function StrIsASCII(ByVal sValue As String) As Boolean
    StrIsASCII = Internal_StrRegExp(sValue, "^[\x00-\x7F]+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsBool
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Checks if a string is representing a boolean value
'---------------------------------------------------------------------------------------
'
Public Function StrIsBool(ByVal sValue As String) As Boolean
    StrIsBool = Internal_StrRegExp(sValue, "^(?:false|true)$", True)
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsDigits
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Checks if a string is digits only i.e. 0-9
'---------------------------------------------------------------------------------------
'
Public Function StrIsDigits(ByVal sValue As String) As Boolean
    StrIsDigits = Internal_StrRegExp(sValue, "^\d+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsHex
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Checks if a string is representing a hexadecimal value
'---------------------------------------------------------------------------------------
'
Public Function StrIsHex(ByVal sValue As String) As Boolean
    StrIsHex = Internal_StrRegExp(sValue, "^0[xX][\dA-Fa-f]+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrInsert
' Author    : SoftwareSpot
' Date      : 11/03/2015
' Purpose   : Insert a string at a specified position. Positive to insert from the left, negative to insert from the right
'---------------------------------------------------------------------------------------
'
Public Function StrInsert(ByVal sString As String, ByVal sInsert As String, ByVal iPosition As Long) As String
    Dim iLength As Integer
    iLength = Len(sString)
    If iPosition < 0 Then
        iPosition = iLength + iPosition
    End If
    If iLength < iPosition Or iPosition < 0 Then
        StrInsert = sString
        Exit Function
    End If
    StrInsert = Left(sString, iPosition) & sInsert & Right(sString, iLength - iPosition)
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrIsInt
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Checks if a string is an integer
'---------------------------------------------------------------------------------------
'
Public Function StrIsInt(ByVal sValue As String) As Boolean
    StrIsInt = Internal_StrRegExp(sValue, "^-?\d+$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrSanitizeFileName
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Replace invalid characters inc. whitespace in a filename
'---------------------------------------------------------------------------------------
'
Public Function StrSanitizeFileName(ByVal sValue As String) As String
    Dim oRegExp As New RegExp
    With oRegExp
        .Pattern = "[:\\/*?""<>|" & Chr$(32) & "]"
    End With
    StrSanitizeFileName = oRegExp.Replace(sValue, "_")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrTrimLeft
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Trims a number of characters from the left hand side of a string
'---------------------------------------------------------------------------------------
'
Public Function StrTrimLeft(ByVal sValue As String, ByVal iCount As Long) As String
    Dim iLength As Long
    iLength = Len(sValue)
    If iLength = 0 Or iCount <= 0 Then
        StrTrimLeft = sValue
        Exit Function
    End If
    If iCount < iLength Then
        StrTrimLeft = Right(sValue, iLength - iCount)
        Exit Function
    End If
    StrTrimLeft = vbNullString
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrTrimRight
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : Trims a number of characters from the right hand side of a string
'---------------------------------------------------------------------------------------
'
Public Function StrTrimRight(ByVal sValue As String, ByVal iCount As Long) As String
    Dim iLength As Long
    iLength = Len(sValue)
    If iLength = 0 Or iCount <= 0 Then
        StrTrimRight = sValue
        Exit Function
    End If
    If iCount < iLength Then
        StrTrimRight = Left(sValue, iLength - iCount)
        Exit Function
    End If
    StrTrimRight = vbNullString
End Function

'---------------------------------------------------------------------------------------
' Procedure : Internal_StrRegExp
' Author    : SoftwareSpot
' Date      : 10/03/2015
' Purpose   : A wrapper for regular expression testing
'---------------------------------------------------------------------------------------
'
Public Function Internal_StrRegExp(ByVal sString, ByVal sPattern As String, Optional ByVal bIsIgnoreCase As Boolean = False)
    Dim oRegExp As New RegExp
    With oRegExp
        .IgnoreCase = bIsIgnoreCase
        .Pattern = sPattern
    End With
    Internal_StrRegExp = oRegExp.Test(sString)
End Function
