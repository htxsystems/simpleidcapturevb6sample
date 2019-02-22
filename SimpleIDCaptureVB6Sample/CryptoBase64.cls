VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CryptoBase64
'
'This is a class wrapper for the Base64 conversion functions in the Windows
'CryptoAPI.  It is defined such that its presence in a VB6 project results
'in a global predeclared instance, i.e. there is no need to explicitly
'create an instance of this class.
'
'This class cannot be used on Win9x operating systems.
'
'Properties:
'
'   Base64Format As Base64FormatEnum (R/W)
'
'           Specifies the format of line breaks in the Base64 text.  Some
'           values are only available in later OSs.
'
'           Default is bfmtCrLf.
'
'   OsVersion As OSVersionEnum (RO)
'
'           Can be used to detect what OS version the program is running
'           under.  Meant for detection of the features this class can
'           support.
'
'Methods:
'
'   Decode(ByRef Base64Buf As String) As Byte()
'
'           Converts a Unicode VB String Base64 value back into the original
'           binary image as a Byte array.  Result should be assigned to a
'           dynamic Byte array in the calling code.
'
'           Conversion is done without respect to Base64Format but the
'           CryptoAPI (and thus Decode) is forgiving of all line-break
'           formats.
'
'   Encode(ByRef BinaryBuf() As Byte) As String
'
'           Converts a binary value in a Byte array into a Base64-encoded
'           representation as a Unicode VB String.
'
'           Formatting is as requested via Base64Format.
'

Private Const CLASS_EXCEPT_BASE As Long = &H8004E300

Public Enum CryptoBase64ExceptEnum
    cbxGetOSVersFailed = CLASS_EXCEPT_BASE
    cbxNotNT
    cbxWinXPOrLaterReqd
    cbxWinVistaOrLaterReqd
    cbxStringToBinaryFailed
    cbxBinaryToStringFailed
End Enum

Public Enum Base64FormatEnum
    bfmtCrLF = 0
    bfmtLfOnly
    bfmtNone
End Enum

Public Enum OSVersionEnum
    osvWinXP = 501
    osvWinVista = 600
End Enum

Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
        'Operating System     Value
        'Windows 3.1            3
        'Windows 95             4
        'Windows 98             4
        'Windows Me             4
        'Windows NT 3.51        3
        'Windows NT 4.0         4
        'Windows 2000           5
        'Windows XP             5
        'Windows .Net Server    5
        'Windows 2003 Server    5
        'Windows 2003 R2 Server 5
        'Windows Vista          6
        'Windows 2008 Server    6
    dwMinorVersion As Long
        'Operating System     Value
        'Windows 3.1            1
        'Windows 95             0
        'Windows 98             10
        'Windows Me             90
        'Windows NT 3.51        51
        'Windows NT 4.0         0
        'Windows 2000           0
        'Windows XP             1
        'Windows .Net Server    1
        'Windows 2003 Server    2
        'Windows 2003 R2 Server 2
        'Windows Vista          0
        'Windows 2008 Server    0
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    
    'Extended information (optional), i.e. OSVERSIONINFOEX:
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
        'Operating System     Value
        'NT Workstation         1
        'NT Domain Controller   2
        'NT Server              3
    wReserved As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFO) As Long

Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPT_STRING_NOCR As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000

Private Declare Function CryptBinaryToString Lib "Crypt32" _
    Alias "CryptBinaryToStringW" ( _
    ByRef pbBinary As Byte, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As Long, _
    ByRef pcchString As Long) As Long

Private Declare Function CryptStringToBinary Lib "Crypt32" _
    Alias "CryptStringToBinaryW" ( _
    ByVal pszString As Long, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As Long, _
    ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, _
    ByRef pdwFlags As Long) As Long

Private m_OSVersion As OSVersionEnum
Private m_lngBase64Format As Long

Public Property Get Base64Format() As Base64FormatEnum
    If m_lngBase64Format = 0 Then
        Base64Format = bfmtCrLF
    ElseIf m_lngBase64Format = CRYPT_STRING_NOCR Then
        Base64Format = bfmtLfOnly
    Else
        Base64Format = bfmtNone
    End If
End Property

Public Property Let Base64Format(ByVal Format As Base64FormatEnum)
    If Format = bfmtLfOnly Then
        If m_OSVersion < osvWinXP Then
            Err.Raise cbxWinXPOrLaterReqd, _
                      "CryptoBase64.Base64Format", _
                      "This format is only supported in Windows XP/2003 and later"
        Else
            m_lngBase64Format = CRYPT_STRING_NOCR
        End If
    ElseIf Format = bfmtNone Then
        'If m_OSVersion < osvWinVista Then
            'Err.Raise cbxWinVistaOrLaterReqd, _
             '         "CryptoBase64.Base64Format", _
              '        "This format is only supported in Windows Vista/2008 and later"
        'Else
            m_lngBase64Format = CRYPT_STRING_NOCRLF
        'End If
    Else
        m_lngBase64Format = 0
    End If
End Property

Public Function Decode(ByRef Base64Buf As String) As Byte()
    Dim lngOutLen As Long
    Dim dwActualUsed As Long
    Dim bytBuf() As Byte
    
    'Determine output buffer length required. Note:
    '  StrPtr(vbNullString) is just a way to get a null pointer,
    '  even though we're really talking about a Byte array here.
    CryptStringToBinary StrPtr(Base64Buf), _
                        Len(Base64Buf), _
                        CRYPT_STRING_BASE64, _
                        StrPtr(vbNullString), _
                        lngOutLen, _
                        0&, _
                        dwActualUsed
    'Convert Base64 to binary.
    ReDim bytBuf(lngOutLen - 1)
    If CryptStringToBinary(StrPtr(Base64Buf), _
                           Len(Base64Buf), _
                           CRYPT_STRING_BASE64, _
                           VarPtr(bytBuf(0)), _
                           lngOutLen, _
                           0&, _
                           dwActualUsed) = 0 Then
        Err.Raise cbxStringToBinaryFailed, _
                  "CryptoBase64.Decode", _
                  "CryptStringToBinary failed, error " & CStr(Err.LastDllError)
    Else
        Decode = bytBuf
    End If
End Function

Public Function Encode(ByRef BinaryBuf() As Byte) As String
    Dim bytBuf() As Byte
    Dim lngOutLen As Long
    Dim strBase64 As String
    
    'Determine Base64 output String length required.
    CryptBinaryToString BinaryBuf(0), _
                        UBound(BinaryBuf) + 1, _
                        CRYPT_STRING_BASE64 Or m_lngBase64Format, _
                        StrPtr(vbNullString), _
                        lngOutLen
    'Convert binary to Base64.
    Encode = String(lngOutLen, 0)
    If CryptBinaryToString(BinaryBuf(0), _
                           UBound(BinaryBuf) + 1, _
                           CRYPT_STRING_BASE64 Or m_lngBase64Format, _
                           StrPtr(Encode), _
                           lngOutLen) = 0 Then
        Err.Raise cbxBinaryToStringFailed, _
                  "CryptoBase64.Encode", _
                  "CryptBinaryToString failed, error " & CStr(Err.LastDllError)
    End If
End Function

Public Property Get OSVersion() As OSVersionEnum
    OSVersion = m_OSVersion
End Property

Private Sub Class_Initialize()
    Dim osvinfData As OSVERSIONINFO
    
    With osvinfData
        .dwOSVersionInfoSize = Len(osvinfData)
        .szCSDVersion = ""
        If GetVersionEx(osvinfData) = 0 Then
            Err.Raise cbxGetOSVersFailed, _
                      "CryptoBase64 Initialize", _
                      "GetVersionEx failed, error " & CStr(Err.LastDllError)
        End If
        If .dwPlatformId <> VER_PLATFORM_WIN32_NT Then
            Err.Raise cbxNotNT, _
                      "CryptoBase64 Initialize", _
                      "CryptoAPI is only available on NT-based OSs"
        End If
        m_OSVersion = .dwMajorVersion * 100 + .dwMinorVersion
    End With
    Base64Format = bfmtCrLF
End Sub


