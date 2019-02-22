Attribute VB_Name = "m_MathExt"
Option Explicit

'' Extended mathematic process library for Visual Basic
''
''
Public Const ProgressDiv = 17

Public Type BASE64STRUCT
    Length As Long
    Data() As Byte
End Type

Public Enum EncodingSchemeConstants
    esRaw = 0
    esBinPack = 1
End Enum

Public Type FRACTION
    fDividend As Long
    fDivisor As Long
End Type

Public Type DIVLEN
    lActual  As Long
    lProcess As Long
    lReturn  As Long
End Type

Public Const BitPackIdConstant = &H7FE432AD

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const BASE64TABLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Public Const BASE64PAD = 61
Public Const BASE64PADRETURN = 254

Public B64CodeOut(0 To 63) As Byte
Public B64CodeReturn(0 To 255) As Byte

Public B64TableCreated As Boolean

Public LShiftTab(0 To 255, 0 To 7) As Byte
Public RShiftTab(0 To 255, 0 To 7) As Byte

Public Sub Main()
    InitMath
End Sub

Public Function InitMath()
    
    GenBase64Tables
    GenShiftTables
    
End Function

Public Sub GenShiftTables()

    Dim i As Integer, _
        j As Integer
        
    Dim z As Integer, _
        n As Integer
        
    For j = 0 To 255
        
        For i = 0 To 7
                                                        
            n = j
            For z = 1 To i
                n = (n And &HFE) / 2
            Next z
            
            RShiftTab(j, i) = n
        
        Next i
        
    Next j
        
    For j = 0 To 255
        
        For i = 0 To 7
                                                        
            n = j
            For z = 1 To i
                n = (n And &H7F) * 2
            Next z
            
            LShiftTab(j, i) = n
        
        Next i
        
    Next j
        
    
End Sub

Public Function RShift(ByVal Value As Byte, ByVal Shift As Long) As Byte

    Select Case Shift
    
        Case Is >= 8
            RShift = 0
            
        Case Is < 0
            RShift = LShift(Value, Abs(Shift))
            
        Case Else
            RShift = RShiftTab(Value, Shift)
    
    End Select

End Function

Public Function LShift(ByVal Value As Byte, ByVal Shift As Byte) As Long

    Select Case Shift
    
        Case Is >= 8
            LShift = 0
            
        Case Is < 0
            LShift = RShift(Value, Abs(Shift))
            
        Case Else
            LShift = LShiftTab(Value, Shift)
                
    End Select

End Function

Public Function GenBase64Tables()
    Dim i As Integer, d As Integer
    
    For i = 0 To 255
        B64CodeReturn(i) = &H7F
    Next i
    
    For i = 0 To 63
        d = Asc(Mid(BASE64TABLE, i + 1, 1))
        B64CodeOut(i) = d And 255
        B64CodeReturn(d) = i And &H3F
    Next i
    
    B64CodeReturn(BASE64PAD) = BASE64PADRETURN
        
    B64TableCreated = True
    
End Function

Public Function BitFraction(ByVal nBits As Integer, bFract As FRACTION)

    Dim i As Long, _
        b As Long
        
    Dim s As Single
    
    s = nBits / 8
    
    i = 1
    Do While (s * i) <> Round(s * i)
        i = i + 1
    Loop

    bFract.fDividend = (s * i)
    bFract.fDivisor = i
    
End Function

Public Function GetDivLen(ByVal nBits As Integer, ByVal lLength As Long, dLen As DIVLEN, Optional ByVal Reverse As Boolean)
    
    Dim f As FRACTION, _
        x As Long, _
        y As Long
    
    BitFraction nBits, f
    
    x = lLength
    
    dLen.lActual = lLength
    
    If Not Reverse Then
        If (x Mod f.fDividend) > 0 Then
            x = (x + (f.fDividend - (x Mod f.fDividend)))
        End If
        
        dLen.lProcess = x
        dLen.lReturn = (x / f.fDividend) * f.fDivisor
    
    Else
        If (x Mod f.fDivisor) > 0 Then
            x = (x + (f.fDivisor - (x Mod f.fDivisor)))
        End If
        
        dLen.lProcess = x
        dLen.lReturn = (x / f.fDivisor) * f.fDividend
    
    End If
    
End Function

Public Function Decode64(DataIn() As Byte, ByVal Length As Long, DataOut() As Byte, Optional ByVal ShowProgress As Boolean = True) As Long
    On Error Resume Next
    
    Dim dLen As DIVLEN, _
        l As Long, _
        n As Long, _
        j As Long, _
        v As Long
    
    Dim Quartet(0 To 3) As Byte
    
    Dim Progress As frmProgress
    
    If (ShowProgress = True) Then
        Set Progress = New frmProgress
        Progress.Show
        DoEvents
    End If
    
    l = (Length / 3) * 4
    ReDim DataOut(0 To l)
    
    v = 0
    n = 0
    For l = 0 To (Length - 1) Step 0
                
        j = 0
        Do While (j < 4)
        
            If (B64CodeReturn(DataIn(l)) <> &H7F) Then
                
                Quartet(j) = B64CodeReturn(DataIn(l))
                j = j + 1
                    
                
            End If
            
            l = l + 1
            
            If (l >= Length) Then
            
                Do While j < 4
                    Quartet(j) = BASE64PADRETURN
                    j = j + 1
                Loop
                
                Exit Do
            End If
                
        Loop

        If (Quartet(0) = BASE64PADRETURN) Or (Quartet(1) = BASE64PADRETURN) Then
            l = -1&
            Exit For
        End If
        
        DataOut(v) = LShift(Quartet(0), 2) Or RShift(Quartet(1), 4)
        v = v + 1

        If (Quartet(2) = BASE64PADRETURN) Then
            l = -1&
            Exit For
        End If

        DataOut(v) = LShift(Quartet(1), 4) Or RShift(Quartet(2), 2)
        v = v + 1
        
        If (Quartet(3) = BASE64PADRETURN) Then
            l = -1&
            Exit For
        End If
        
        DataOut(v) = LShift(Quartet(2), 6) Or Quartet(3)
        v = v + 1
        
        If ShowProgress = True Then
            If n >= (Length / ProgressDiv) Then
                n = 0
                Progress.ProgressBar1.Value = (l / (Length - 1)) * 100
                DoEvents
            Else
                n = n + 1
            End If
        End If
        
    Next l
        
    ReDim Preserve DataOut(0 To (v - 1))
    
    Decode64 = v
    
    If ShowProgress = True Then
        Unload Progress
        Set Progress = Nothing
    End If

End Function

Public Function Encode64(DataIn() As Byte, ByVal Length As Long, b64Out As BASE64STRUCT, Optional ByVal ShowProgress As Boolean = True)

    Dim dLen As DIVLEN, _
        l As Long, _
        x As Long, _
        j As Long, _
        v As Long
    
    Dim lC As Long, _
        n As Long
    
    Dim Progress As frmProgress
    
    If (ShowProgress = True) Then
        Set Progress = New frmProgress
    End If
    
    GetDivLen 6, Length, dLen
    
    ReDim Preserve DataIn(0 To (dLen.lProcess - 1))
    If Round(dLen.lReturn / 76) <> (dLen.lReturn / 76) Then
        v = Round(dLen.lReturn / 76) + 1
    Else
        v = dLen.lReturn / 76
    End If
    
    v = v * 2
    
    l = (dLen.lReturn - 1) + v
    
    ReDim b64Out.Data(0 To l)
    b64Out.Length = l + 1
    
    l = dLen.lProcess - 1
    v = 0
    
    If ShowProgress = True Then
        Progress.Show
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    n = 0
    For x = 0 To l Step 3
        
        b64Out.Data(v) = B64CodeOut(RShift(DataIn(x), 2))
        b64Out.Data(v + 1) = B64CodeOut((LShift(DataIn(x), 4) And &H30) Or _
                             (RShift(DataIn(x + 1), 4)))
        b64Out.Data(v + 2) = B64CodeOut((LShift(DataIn(x + 1), 2) And &H3C) Or _
                              (RShift(DataIn(x + 2), 6)))
                              
        b64Out.Data(v + 3) = B64CodeOut(DataIn(x + 2) And &H3F)
        
        v = v + 4
        lC = lC + 4
        If (lC >= 76) Then
            
            lC = 0
            If (l - x) > 3 Then
                b64Out.Data(v) = 13
                b64Out.Data(v + 1) = 10
                v = v + 2
            End If
        End If
        
        If ShowProgress = True Then
            If n >= (l / ProgressDiv) Then
                n = 0
                Progress.ProgressBar1.Value = (x / l) * 100
                DoEvents
            Else
                n = n + 1
            End If
        End If
        
    Next x
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    Select Case (dLen.lProcess - dLen.lActual)
    
        Case 1:
            b64Out.Data(v - 1) = BASE64PAD
        
        Case 2:
            b64Out.Data(v - 1) = BASE64PAD
            b64Out.Data(v - 2) = BASE64PAD
        
    End Select
            
    If lC <> 0 Then
        b64Out.Data(v) = 13
        b64Out.Data(v + 1) = 10
        ReDim Preserve b64Out.Data(0 To v + 1)
        
    Else
        ReDim Preserve b64Out.Data(0 To v - 1)
    
    End If
            
    If (ShowProgress = True) Then
        Unload Progress
        Set Progress = Nothing
    End If

End Function

Public Function EncodeData(DataIn() As Byte, _
                           DataOut() As Byte, _
                           ByVal DataInLength As Long, _
                           ByVal nBitsOut As Long, _
                           Optional ByVal EncodingScheme As EncodingSchemeConstants = esBinPack, _
                           Optional ByVal ShowProgress As Boolean = False)

    Dim a As Long, _
        b As Long, _
        c As Long, _
        d As Long, _
        e As Long, _
        f As Long, _
        g As Long
        
    Dim lpIn As Long, _
        lpOut As Long
        
    Dim lCountOut As Long
    
    
    
    
    Select Case EncodingScheme
    
    
    
    End Select
    
    
    
End Function


Public Function DecodeData(DataIn() As Byte, _
                           DataOut() As Byte, _
                           ByVal DataInLength As Long, _
                           ByVal nBitsIn As Long, _
                           Optional ByVal EncodingScheme As EncodingSchemeConstants = esBinPack, _
                           Optional ByVal ShowProgress As Boolean = False)

    Dim a As Long, _
        b As Long, _
        c As Long, _
        d As Long, _
        e As Long, _
        f As Long, _
        g As Long
        
    Dim lpIn As Long, _
        lpOut As Long
        
    Dim lCountOut As Long
    
    
    
    Select Case EncodingScheme
    
    
    
    End Select




End Function




