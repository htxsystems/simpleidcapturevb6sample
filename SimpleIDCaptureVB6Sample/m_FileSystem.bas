Attribute VB_Name = "m_FileSystem"
Option Explicit

' Custom File System Utility Module

' Replaces the functionality to the
' Scripting.FileSystemObject to some extent

' 2000 Nathan Moschkin (nmosch@tampabay.rr.com)

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const MAX_PATH = 260
Private Const WM_USER = &H400

Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal bLen As Long, ByVal lpszBuffer As String) As Long

Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilename As String) As Long

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFilename As String) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Private Declare Function SHBrowseForFolder Lib "Shell32" (bInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Const SHGFI_DISPLAYNAME = &H200
Private Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1   ' For finding a folder to start document searching
Private Const BIF_DONTGOBELOWDOMAIN = &H2 ' For starting the Find Computer
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20 ' insist on valid result (or CANCEL)

Private Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Private Const BIF_BROWSEFORPRINTER = &H2000 ' Browsing for Printers
Private Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything

' message from browser

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILED = 3  ' lParam:szPath ret:1(cont),0(EndDialog)
' Private Const BFFM_VALIDATEFAILEDW = 4  ' lParam:wzPath ret:1(cont),0(EndDialog)

' messages to browser

Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)

Private Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long
    DisplayName As String
    lpszTitle As String
    ulFlags As Integer
    lpfn As Long
    lParam As Long
    iImage As Integer
End Type

Private Type FILETIME
        Hi As Long
        Lo As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private wfind As WIN32_FIND_DATA
Private mBrowseCurrent As String

Public BrowserLastFolder As Long

' SHBrowseForFolder callback function

' Since we need the absolute path, we take the path of the most recently
' selected item.  The only way to do that is through processing of the
' BFFM_SELCHANGED event generated by SHBrowseForFolder, and retrieve the
' full path using the SHGetPathFromIDList and wParam

Private Function BrowseCallback(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim u As Long

    If uMsg = BFFM_SELCHANGED Then
        mBrowseCurrent = String(260, 0)
        SHGetPathFromIDList wParam, mBrowseCurrent
        BrowserLastFolder = wParam
    End If
        
End Function

' Cannot get the address of a function into a variable directly,
' so use this function to return the numeric value of the pointer

Private Function CallbackAddress(ByVal Address As Long) As Long
    CallbackAddress = Address
End Function

' Opens the folder browser

Public Function FolderBrowse(ByVal hwnd As Long, Optional ByVal vTitleText As String = "Browse for Folder") As String
    Dim bInfo As BROWSEINFO, i As Long
    bInfo.ulFlags = &H30 + 13
    
    bInfo.hWndOwner = hwnd
    bInfo.lpszTitle = vTitleText
    bInfo.DisplayName = String(260, 0)
    
    ' Get the address of the callback function
    ' must be passed indirectly, see notes above (CallbackAddress() function)
    
    i = CallbackAddress(AddressOf BrowseCallback)
    
    bInfo.lpfn = i
    
    i = SHBrowseForFolder(bInfo)
    
    i = InStr(1, mBrowseCurrent, Chr(0))
    FolderBrowse = GetAbsolutePathName(Mid(mBrowseCurrent, 1, i - 1))
    If Len(FolderBrowse) Then
        If Mid(FolderBrowse, Len(FolderBrowse), 1) <> "\" Then
            FolderBrowse = FolderBrowse + "\"
        End If
    End If
    
End Function

' Add a backslash if there is none, or remove double backslash

Public Function CleanDir(varStr As String) As String

    If Mid(varStr, Len(varStr) - 1) = "\\" Then
        CleanDir = Mid(varStr, 1, Len(varStr) - 1)
    ElseIf Mid(varStr, Len(varStr)) <> "\" Then
        CleanDir = varStr + "\"
    Else
        CleanDir = varStr
    End If

End Function

' Return the current working directory for the current process

Public Function CurrentDirectory() As String
    Dim Str As String * 260
    Dim i As Integer
    
    GetCurrentDirectory 260, Str
    i = InStr(1, Str, Chr(0))
    CurrentDirectory = Mid(Str, 1, i - 1)
    
End Function

' Return all the files in a directory in a string array that match a search expression

Public Function EnumFiles(ByVal vSearchExpression As String) As String()

    Dim b() As String, z As Long, i As Long
    
    Dim l As Long
    
    l = FindFirstFile(vSearchExpression, wfind)
    If l = -1& Then Exit Function
    
    i = 1
    Do While i <> 0
        
        ReDim Preserve b(z)
        b(z) = Mid(wfind.cFileName, 1, InStr(1, wfind.cFileName, Chr(0)) - 1)
        z = z + 1
        i = FindNextFile(l, wfind)
    Loop
    FindClose l
    
    EnumFiles = b
    
End Function

' Get the absolute file name of a file

Public Function GetAbsolutePathName(ByVal FileName As String) As String

    Dim s As String * 261, T As String * 261
    Dim i As Integer
    
    GetFullPathName FileName, 260, s, T
    
    i = InStr(1, s, Chr(0))
    GetAbsolutePathName = Mid(s, 1, i - 1)
    
End Function

' extract the path of the filename somehow

Public Function ExtractPath(ByVal FileName As String)
    
    Dim s As String * 261, T As String * 261
    Dim u As String, sfile As String
    Dim i As Integer
    Dim sh As SHFILEINFO
    Dim ff As WIN32_FIND_DATA, l As Long
    
    If FileName = "" Then Exit Function
    
    sfile = FileName
    i = Len(sfile)
    If Mid(sfile, i, 1) = "\" Then
        sfile = Mid(sfile, 1, i - 1)
    End If
    
    l = FindFirstFile(sfile, ff)
    
    If l <> -1& Then
        FindClose l
    
        ' if the string that the calling function has passed is a path in and of
        ' itself, we assume that this is the path that they want anyway
        
        If ff.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            ExtractPath = FileName
            i = Len(FileName)
        End If
        
        Exit Function
        
    End If
    
    
    ' otherwise just get the filename of the file ...

    GetFullPathName FileName, 260, s, T
    
    ' ... get the displayname ...
    
    SHGetFileInfo FileName, 0, sh, Len(s), SHGFI_DISPLAYNAME
    u = Mid(sh.szDisplayName, 1, InStr(1, sh.szDisplayName, Chr(0)) - 1)
    
    ' ... compute the difference and return it!
    
    i = InStr(1, UCase(s), UCase(u))
    ExtractPath = Mid(s, 1, i - 1)

End Function

' Recursively makes a directory into existance.

' Creates every directory that does not exist, descending from root
' to the path argument passed, as needed.

Public Function MkRecurse(ByVal vDir As String)
    
    Dim j As Integer, i As Integer, s As String
    
    i = 4
    j = 1
    Do While j <> 0
        j = InStr(i, vDir, "\")
        If j = 0 Then j = Len(vDir)
        
        s = Mid(vDir, 1, j)
        If (FolderExists(s) = False) Then
            MkDir s
        End If
        If j >= Len(vDir) Then Exit Do
        i = j + 1
    Loop

End Function

' Extract the display name for an existing file

Public Function GetFileName(ByVal FileName As String) As String

    Dim s As SHFILEINFO
    Dim i As Integer
    
    
    SHGetFileInfo FileName, 0, s, Len(s), SHGFI_DISPLAYNAME
    GetFileName = Mid(s.szDisplayName, 1, InStr(1, s.szDisplayName, Chr(0)) - 1)
    
End Function

' Returns a boolean indicating weather the file exists or not

Public Function FileExists(ByVal FileName As String) As Boolean
    Dim wf As WIN32_FIND_DATA
    Dim i As Long
    
    i = FindFirstFile(FileName, wf)

    If i <> -1& Then
        FileExists = True
        FindClose i
    End If

End Function

' Returns a boolean indicating weather the folder exists or not

Public Function FolderExists(ByVal FolderName As String) As Boolean

    Dim i As Long, s As String
    Dim wf As WIN32_FIND_DATA
    
    s = FolderName
    i = Len(s)
    If Len(s) = 0 Then Exit Function
    
    If Mid(s, i, 1) = "\" Then
        s = Mid(s, 1, i - 1)
    End If
    
    wf.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY
    
    i = FindFirstFile(s, wf)

    If i <> -1& Then
        If wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        End If
        FindClose i
    End If
    

End Function

'' Get the barebones name of a file, optionally hold on to the path.

Public Function Barename(vData As String, Optional ByVal KeepPath As Boolean = True) As String

    Dim i As Long, s As String, j As Long
    
    i = InStrRev(vData, "\")
    
    If i = Len(vData) Then
        s = Mid(vData, 1, i - 1)
    Else
        s = vData
    End If
    
    i = InStrRev(vData, ".")
    j = InStrRev(vData, "\")
    
    If Not j Then
        
        If i Then
            Barename = Mid(s, 1, i - 1)
        Else
            Barename = s
        End If
    
    Else
    
        If i Then
        
            If i > j Then
                
                If Not KeepPath Then
                    Barename = Mid(s, j + 1, i - 1)
                Else
                    Barename = Mid(s, j + 1)
                End If
                
            Else
                
                If Not KeepPath Then
                    Barename = Mid(s, j + 1)
                Else
                    Barename = s
                End If
            
            End If
            
        Else
            
            Barename = Mid(s, j + 1)
        End If
        
    End If
    
    Exit Function
        
End Function

'' Write a string to a binary stream

Public Function WriteString(ByVal FileNo As Integer, StrData As String)

    Dim x As Long
    
    x = Len(StrData)
    
    Put FileNo, , x&
    Put FileNo, , StrData
    
End Function

'' Read a string from a binary stream

Public Function ReadString(ByVal FileNo As Integer) As String

    Dim x As Long
    
    Get FileNo, , x&
    
    ReadString = String(x, Chr(0))
    
    Get FileNo, , ReadString
    
End Function

'' Find a string in a collection of strings (optional case-insensitive)
'' For use with matching files.

Public Function StrFind(vStrs() As String, vMatch As String, Optional ByVal vIgnoreCase As Boolean = False) As Boolean

    Dim i As Long, j As Long
    Dim x As Long
    
    StrFind = False
    
    On Error Resume Next
    
    i = -1
    j = -1
    
    i = LBound(vStrs)
    j = UBound(vStrs)
    
    If i < 0 Or j < 0 Then Exit Function
    
    If vMatch = "" Then Exit Function
    
    For x = i To j
    
        If vIgnoreCase = True Then
            If LCase(vStrs(x)) = LCase(vMatch) Then
                StrFind = True
                Exit Function
            End If
        Else
            If vStrs(x) = vMatch Then
                StrFind = True
                Exit Function
            End If
        End If
    Next x
    
End Function

Public Function FileErr(ByVal ForWrite As Boolean, ByVal ForRead As Boolean) As Boolean


    

End Function

'' Generate a unique file name in the specified directory with the provided prefix
'' Optionally reset internal counter.

Public Function GetUnique(varPrefix As String, Optional varPath As String, Optional varReset As Boolean) As String

    Static c As Long

    Dim s As String, j As String, strs() As String
    Dim x As Boolean, y As Long, T As Date
    
    On Error Resume Next
    
    x = -1
    y = -1
    
    If varPath = "" Then varPath = CurrentDirectory
    If (c = 0) Or (varReset = True) Then c = Now() And &H3FFF&
    
    s = CleanDir(varPath)
    j = s + Mid(varPrefix, 1, 4) + "*.T$$"
        
    strs = EnumFiles(j)
    
    j = Mid(varPrefix, 1, 4) + LCase(Hex(c)) + ".T$$"
    
    x = StrFind(strs, j, True)
    
    Do While (x = True) And (c <= &HFFFF)
    
        c = c + 1
        
        j = Mid(varPrefix, 1, 4) + LCase(Hex(c)) + ".T$$"
    
        x = StrFind(strs, j, True)
        
    Loop

    If (c > &HFFFF&) Then Exit Function
    
    GetUnique = s + j
    
End Function

'' Binary search in a file (byte aligned).  Returns the position where the value begins in the file.
'' Set seekstart to -1 to search from current position

Public Function BinaryFind(ByVal varFile As Integer, varValue As Variant, Optional ByVal SeekStart As Long = 0) As Long
    Dim varOffset As Long
    
    Dim varBinInteger As Integer
    Dim varBinByte As Byte
    Dim varBinLong As Long
    Dim varLOF As Long
    
    Dim b As Byte, c As Integer
    Dim e As Integer, f As Long
    
    If SeekStart <> -1& Then
        Seek varFile, SeekStart
    End If
    
    varLOF = LOF(varFile)
    
    Select Case VarType(varValue)
    
        Case vbByte:
        
            varBinByte = varValue
            varOffset = Seek(varFile)
            
            Do While varOffset <= varLOF
                    
                varOffset = Seek(varFile)
                
                Get varFile, , b
                
                If b = varBinByte Then
                    BinaryFind = varOffset
                    Exit Function
                End If
                                    
                varOffset = Seek(varFile)
                    
            Loop
        
        Case vbInteger:
        
            varBinInteger = varValue
            varOffset = Seek(varFile)
            
            Do While varOffset <= varLOF
                    
                Get varFile, , b
                e = ((e And &HFF) * &H100) Or b
                
                If e = varBinInteger Then
                    varOffset = Seek(varFile) - 2
                    Seek varFile, varOffset
                    BinaryFind = varOffset
                    Exit Function
                End If
                                    
                varOffset = Seek(varFile)
                    
            Loop
        
        Case vbLong:
        
            '' Long value searches are word aligned
        
            varBinLong = varValue
            varOffset = Seek(varFile)
            
            Do While varOffset <= varLOF
                                    
                Get varFile, , c
                f = ((f And &HFFFF&) * &H10000) Or c
                
                If f = varBinLong Then
                    varOffset = Seek(varFile) - 4
                    Seek varFile, varOffset
                    BinaryFind = varOffset
                    Exit Function
                End If
                                    
                varOffset = Seek(varFile)
                    
            Loop
        
        
    End Select

    BinaryFind = -1&

End Function

Public Function GetArrayLen(varArray As Variant) As Long

    Dim lb As Long, ub As Long
    
    On Error Resume Next
    
    lb = -1
    ub = -1
    
    lb = LBound(varArray)
    ub = UBound(varArray)
    
    If lb = -1 Then
        GetArrayLen = 0&
        Exit Function
    End If
    
    GetArrayLen = (ub - lb) + 1

End Function

' Delete a portion of a file

Public Function Sectate(varFilename As String, varStart As Long, varLen As Long) As Boolean

    Dim f As Integer, e As Integer
    Dim b() As Byte
    Dim varStr As String
    
    Dim lPos1 As Long, lPos2 As Long, l As Long
    Dim lSizeMax As Long
    
    On Error GoTo SectateErr
    
    e = FreeFile
    f = FreeFile
    
    lSizeMax = 65535
    
    varStr = GetUnique(varFilename, , True)
    
    Open varFilename For Binary Access Read Write Lock Read Write As e
    Open varStr For Binary Access Read Write Lock Read Write As f
        
    lPos1 = Seek(e)
    
    Do
            
        l = varStart - lPos1
                
        If l <= 0 Then Exit Do
        If l > lSizeMax Then l = lSizeMax
        
        ReDim b(1 To l)
        
        Get e, , b
        Put f, , b
        
        lPos1 = Seek(e)
    
    Loop
        
    Seek e, (varStart + varLen)
    lPos1 = Seek(e)
    
    Do
    
        l = LOF(e) - lPos1
                
        If l <= 0 Then Exit Do
        If l > lSizeMax Then l = lSizeMax
        
        ReDim b(1 To l)
        
        Get e, , b
        Put f, , b
        
        lPos1 = Seek(e)
    
    Loop
        
    Close e
    Close f
    
    DeleteFile varFilename
    MoveFile varStr, varFilename
    
    Sectate = True
        
SectateErr:

    Close e
    Close f
    
    DeleteFile varStr
    Sectate = False
    Exit Function

End Function

Public Function DisplayErr(ByVal varECode As Long, varEText As String, Optional ByVal varMsgBoxStyle As VbMsgBoxStyle = (vbApplicationModal Or vbOKCancel Or vbCritical)) As VbMsgBoxResult

    If IsEmpty(varMsgBoxStyle) Then
    
        DisplayErr = MsgBox("Encountered Error #" + Format(varECode) + ": " + varEText + ".  Press 'Cancel' to end the program", varMsgBoxStyle, "Error Encountered")
        If DisplayErr = vbCancel Then Exit Function
        
    Else
        
        DisplayErr = MsgBox("Encountered Error #" + Format(varECode) + ": " + varEText, varMsgBoxStyle, "Error Encountered")
    
    End If
    
End Function


