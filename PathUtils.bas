Attribute VB_Name = "PathUtils"
'---------------------------------------------------------------------------------------
' Module    : PathUtils
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Filepath and filename related functions
'---------------------------------------------------------------------------------------

Public Const BACKSLASH_LEN As Integer = 1
Public Const READYSTATE_FINISHED As Integer = 4 ' Unable to find a constant for .readyState

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : FileExists
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Determines if a file or directory exists
'---------------------------------------------------------------------------------------
'
Public Function FileExists(ByVal sFilePath As String) As Boolean ' Checking for an extension = StrRegExp(sFilePath, "\.\w+")
    If IsFile(sFilePath) Then ' First assume it's a file
        FileExists = True
        Exit Function
    End If
    FileExists = IsDir(sFilePath) ' If all else fails then check as though it was a folder.
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileSize
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Get the size of a file (wrapper for FileLen)
'---------------------------------------------------------------------------------------
'
Public Function FileSize(ByVal sFilePath As String) As Long
    FileSize = FileLen(sFilePath)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDir
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Get the directory from a string filepath
'---------------------------------------------------------------------------------------
'
Public Function GetDir(ByVal sFilePath As String) As String
    Dim iIndex As Integer
    If StrRegExp(sFilePath, "\.\w+$") Then ' If there is an extension at the end of the filepath
        iIndex = InStrRev(sFilePath, "\")
        If iIndex > 0 Then
            GetDir = Left(sFilePath, iIndex - 1)
            Exit Function
        End If
    End If
    GetDir = IIf(Right(sFilePath, BACKSLASH_LEN) = "\", StrTrimRight(sFilePath, BACKSLASH_LEN), sFilePath) ' Remove appended backslash
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileName
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Get the filename from a string filepath
'---------------------------------------------------------------------------------------
'
Public Function GetFileName(ByVal sFilePath As String) As String
    If Right(sFilePath, BACKSLASH_LEN) <> "\" Then
        Dim iIndex As Integer
        iIndex = InStrRev(sFilePath, "\")
        If iIndex > 0 Then
            GetFileName = StrTrimLeft(sFilePath, iIndex)
            Exit Function
        End If
    End If
    GetFileName = vbNullString
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTempPath
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Retrieve the system temporary path
'---------------------------------------------------------------------------------------
'
Public Function GetTempPath() As String
    GetTempPath = Environ$("TEMP")
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsDir
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Checks if a filepath exists and is a directory
'---------------------------------------------------------------------------------------
'
Public Function IsDir(ByVal sFilePath As String) As Boolean
    On Error Resume Next
    IsDir = ((GetAttr(sFilePath) And vbDirectory) = vbDirectory)
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsDowloadedFile
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Download a url to a specified filepath
'---------------------------------------------------------------------------------------
'
Public Function IsDowloadedFile(ByVal sUrl As String, ByVal sFilePath As String, Optional ByVal bForceDelete As Boolean = False) As Boolean ' http://stackoverflow.com/questions/17877389/how-do-i-download-a-file-using-vba-without-internet-explorer
    On Error GoTo IsDowloadedFile_Error:
    
    Dim oFSO As New FileSystemObject ' As this will also be used in the error handler
    
    If Not bForceDelete And oFSO.FileExists(sFilePath) Then
        IsDowloadedFile = True
        Exit Function
    End If
    
    Dim oXmlHttpRequest As New XMLHTTP
    oXmlHttpRequest.Open "GET", sUrl, False
    oXmlHttpRequest.send

    Do While oXmlHttpRequest.readyState <> Constants.READYSTATE_FINISHED ' Wait for the download to finish
        DoEvents
    Loop
    
    Dim adResponse() As Byte
    adResponse = oXmlHttpRequest.responseBody ' Returns a byte array

    ' Delete the outdated file
    If bForceDelete And oFSO.FileExists(sFilePath) Then
        oFSO.DeleteFile sFilePath, True ' Could use Kill() as well
    End If
    Dim sDir As String
    sDir = GetDir(sFilePath)
    If Not oFSO.FolderExists(sDir) Then
        MkDir (sDir)
    End If
    
    ' Create a local file and save the data
    Dim hFileOpen As Long
    hFileOpen = FreeFile
    Open sFilePath For Binary Access Write As #hFileOpen
    Put #hFileOpen, , adResponse
    Close #hFileOpen
    
    IsDowloadedFile = True
    Exit Function
    
IsDowloadedFile_Error:
    If Err.Number <> 0 Then
        Debug.Print "A random error occurred = Code: " & CStr(Err.Number) & ", Description: " & Err.Description
        Err.Clear
    End If
    
    IsDowloadedFile = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsFile
' Author    : SoftwareSpot
' Date      : 25/03/2015
' Purpose   : Checks if a filepath exists and is a file
'---------------------------------------------------------------------------------------
'
Public Function IsFile(ByVal sFilePath As String) As Boolean
    On Error Resume Next
    IsFile = ((GetAttr(sFilePath) And vbDirectory) <> vbDirectory)
End Function
