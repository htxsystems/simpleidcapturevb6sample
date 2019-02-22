VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modImageFromStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type GUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type

Private Type PicBmp
    size        As Long
    Type        As Long
    hBmp        As Long
    hpal        As Long
    Reserved    As Long
End Type

Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IUnknown, image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Sub IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any)

Private Const GMEM_MOVEABLE As Long = &H2

Public Function LoadStreamImage(stm As ADODB.stream) As stdPicture
    Dim token   As Long
    Dim gpInput As GdiplusStartupInput
    Dim img     As Long
    Dim dat()   As Byte
    
    gpInput.GdiplusVersion = 1
    
    If GdiplusStartup(token, gpInput) Then Exit Function
    ' Get binary data
    dat = stm.Read()
    Dim hMem    As Long
    ' Allocate memory block
    hMem = GlobalAlloc(GMEM_MOVEABLE, stm.size)
    
    If hMem Then
        Dim ptr     As Long
        Dim istm    As IUnknown
        ' Copy data to allocated block
        ptr = GlobalLock(hMem)
        CopyMemory ByVal ptr, dat(0), stm.size
        GlobalUnlock hMem
        ' Create IStream interface object
        If CreateStreamOnHGlobal(hMem, 1&, istm) = 0 Then
            ' Load from stream
            If GdipLoadImageFromStream(istm, img) = 0 Then
                Dim hBmp    As Long
                ' Create a new gdi bitmap from gdi+ bitmap
                If GdipCreateHBITMAPFromBitmap(img, hBmp, vbBlack) = 0 Then
                    Dim pic As PicBmp
                    Dim iid As GUID
                    ' Set picture description
                    With pic
                        .hBmp = hBmp
                        .size = Len(pic)
                        .Type = vbPicTypeBitmap
                    End With
                    ' Get iid (IPicture)
                    IIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), iid
                    ' Create COM picture object (IPicture)
                    OleCreatePictureIndirect pic, iid, 1, LoadStreamImage
                    
                End If
                
                GdipDisposeImage img
                
            End If
            
        End If
        
        GlobalFree hMem
        
    End If
    
    GdiplusShutdown token
    
End Function



