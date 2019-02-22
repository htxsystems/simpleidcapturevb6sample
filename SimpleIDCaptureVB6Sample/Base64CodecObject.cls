VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base64CodecObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DataTypeConstants
    dtBinary = &H0
    dtBase64 = &H1
End Enum

Private B64 As BASE64STRUCT
Private m_ShowProgress As Boolean

Private m_DataType As Base64Codec.DataTypeConstants

Public Property Get DataType() As Base64Codec.DataTypeConstants
    DataType = m_DataType
End Property

Public Property Get Length() As Long
    Length = B64.Length
End Property

Public Property Get Data() As Byte()
    
    Data = B64.Data
    
End Property

Public Sub Clear()
    Erase B64.Data
    B64.Length = 0
    
End Sub

Public Function Encode(Bytes() As Byte, ByVal Length As Long) As Long

    Encode = Encode64(Bytes, Length, B64, m_ShowProgress)
    m_DataType = dtBase64
    
End Function

Public Function Decode(Bytes() As Byte, ByVal Length As Long) As Long
    On Error Resume Next
    
    Erase B64.Data
    
    ReDim B64.Data(0 To (Length - 1))
    B64.Length = Decode64(Bytes, Length, B64.Data, m_ShowProgress)
    m_DataType = dtBinary

End Function

Public Property Let Unicode_String(ByVal vData As String)
    Dim vNarrow As String, _
        vBytes() As Byte, _
        i As Long
        
    i = Len(vData)
    vNarrow = StrConv(vData, vbFromUnicode)
    ReDim vBytes(0 To (i - 1))
    
    CopyMemory vBytes(0), ByVal StrPtr(vNarrow), i
    Decode vBytes, i
    
End Property

Public Property Get UnicodeString() As String
    UnicodeString = StrConv(B64.Data, vbUnicode)
End Property

Public Property Let ShowProgressDialog(ByVal vData As Boolean)
    m_ShowProgress = vData
End Property

Public Property Get ShowProgressDialog() As Boolean
    ShowProgressDialog = m_ShowProgress
End Property

Private Sub Class_Terminate()
    
    Erase B64.Data
    
End Sub

Public Function EncodeFile(ByVal lpFilename As String, Optional ByVal lpOutfile As String, Optional ByVal PromptOverwrite As Boolean = True) As Boolean

    Dim lpData() As Byte, _
        i As Integer, _
        n As Long
        
    If Not FileExists(lpFilename) Then Exit Function
    
    If (UCase(lpFilename) = UCase(lpOutfile)) Then
        lpOutfile = lpOutfile + ".base64"
    End If
    
    If (FileExists(lpOutfile) And PromptOverwrite) Then
        i = MsgBox(lpOutfile & " exists.  Overwrite?", vbYesNo + vbApplicationModal)
        If (i = vbNo) Then
            Exit Function
        Else
            DeleteFile lpOutfile
        End If
    End If
        
    i = FreeFile

    Open lpFilename For Binary Access Read Lock Write As i
    
    n = LOF(i) - 1
    
    ReDim lpData(0 To n)
    Get i, , lpData
    Close i

    Me.Encode lpData, n + 1

    If lpOutfile <> "" Then
    
        i = FreeFile
        Open lpOutfile For Binary Access Write As i
        
        Put i, , B64.Data
        Close i
    End If

    Erase lpData

End Function

Public Function DecodeFile(ByVal lpFilename As String, Optional ByVal lpOutfile As String, Optional ByVal PromptOverwrite As Boolean = True) As Boolean

    Dim lpData() As Byte, _
        i As Integer, _
        n As Long
        
    If Not FileExists(lpFilename) Then Exit Function
        
    If (UCase(lpFilename) = UCase(lpOutfile)) Then
        lpOutfile = lpOutfile + ".bin"
    End If
    
    If (FileExists(lpOutfile) And PromptOverwrite) Then
        i = MsgBox(lpOutfile & " exists.  Overwrite?", vbYesNo + vbApplicationModal)
        If (i = vbNo) Then
            Exit Function
        Else
            DeleteFile lpOutfile
        End If
    End If
        
    i = FreeFile

    Open lpFilename For Binary Access Read Lock Write As i
    
    n = LOF(i) - 1

    ReDim lpData(0 To n)
    Get i, , lpData
    Close i
    
    Me.Decode lpData, n + 1

    If lpOutfile <> "" Then
    
        i = FreeFile
        Open lpOutfile For Binary Access Write As i
        
        Put i, , B64.Data
        Close i
    End If

    Erase lpData

End Function





