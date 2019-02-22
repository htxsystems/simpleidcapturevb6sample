VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form Main 
   Caption         =   "SimpleID Vb6 Sample"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   10920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   7320
      Width           =   11415
   End
   Begin TabDlg.SSTab tabMenu 
      Height          =   5775
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Register"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbTitle01"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbCustom1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbCustom2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbCustom3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fFingers"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btClearRegister"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "pbRegisterFace"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCustom2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCustom3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pbRegisterFingerprint"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btLoadImageRegisterFace"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btStartRegister"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCustom1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "btRegisterFaceOnly"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "btCancelRegister"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "btSubmitRegister"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Search Fingerprint"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btCancelSearchFingerprint"
      Tab(1).Control(1)=   "btSearchFingerprint"
      Tab(1).Control(2)=   "pbSearchFingerprint"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Search Face"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btSearchFace"
      Tab(2).Control(1)=   "btLoadImageSearchFace"
      Tab(2).Control(2)=   "pbSearchFace"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Delete"
      TabPicture(3)   =   "Main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "btDeletePerson"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton btCancelSearchFingerprint 
         Caption         =   "Cancel Search"
         Height          =   615
         Left            =   -72240
         TabIndex        =   40
         Top             =   4680
         Width           =   4815
      End
      Begin VB.CommandButton btSearchFingerprint 
         Caption         =   "Search Fingerprint"
         Height          =   615
         Left            =   -72240
         TabIndex        =   39
         Top             =   3960
         Width           =   4815
      End
      Begin VB.PictureBox pbSearchFingerprint 
         Height          =   3255
         Left            =   -72240
         ScaleHeight     =   3195
         ScaleWidth      =   4635
         TabIndex        =   38
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton btSearchFace 
         Caption         =   "Search Face"
         Height          =   615
         Left            =   -66000
         TabIndex        =   37
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton btLoadImageSearchFace 
         Caption         =   "Load Image"
         Height          =   615
         Left            =   -72240
         TabIndex        =   36
         Top             =   4080
         Width           =   4695
      End
      Begin VB.PictureBox pbSearchFace 
         AutoRedraw      =   -1  'True
         Height          =   3255
         Left            =   -72240
         ScaleHeight     =   3195
         ScaleWidth      =   4635
         TabIndex        =   35
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton btDeletePerson 
         Caption         =   "Delete Person"
         Height          =   975
         Left            =   -71760
         TabIndex        =   34
         Top             =   1800
         Width           =   3975
      End
      Begin VB.CommandButton btSubmitRegister 
         Caption         =   "Submit"
         Height          =   615
         Left            =   9720
         TabIndex        =   33
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton btCancelRegister 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   7920
         TabIndex        =   32
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton btRegisterFaceOnly 
         Caption         =   "Register Face Only"
         Height          =   615
         Left            =   5400
         TabIndex        =   31
         Top             =   4920
         Width           =   2415
      End
      Begin VB.TextBox txtCustom1 
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton btStartRegister 
         Caption         =   "Start Register"
         Height          =   615
         Left            =   2880
         TabIndex        =   16
         Top             =   4920
         Width           =   2415
      End
      Begin VB.CommandButton btLoadImageRegisterFace 
         Caption         =   "Load Image"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   2415
      End
      Begin VB.PictureBox pbRegisterFingerprint 
         Height          =   2175
         Left            =   9000
         ScaleHeight     =   2115
         ScaleWidth      =   2355
         TabIndex        =   13
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtCustom3 
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtCustom2 
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.PictureBox pbRegisterFace 
         Height          =   2175
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton btClearRegister 
         Caption         =   "Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame fFingers 
         Caption         =   "Fingers"
         Enabled         =   0   'False
         Height          =   1455
         Left            =   2880
         TabIndex        =   14
         Top             =   3360
         Width           =   8535
         Begin VB.CommandButton btLeftLittle 
            Caption         =   "Left Little"
            Height          =   495
            Left            =   6840
            TabIndex        =   30
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btLeftRing 
            Caption         =   "Left Ring"
            Height          =   495
            Left            =   5160
            TabIndex        =   29
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btLeftMiddle 
            Caption         =   "Left Middle"
            Height          =   495
            Left            =   3480
            TabIndex        =   28
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btLeftIndex 
            Caption         =   "Left Index"
            Height          =   495
            Left            =   1800
            TabIndex        =   27
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btLeftThumb 
            Caption         =   "Left Thumb"
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btRightLittle 
            Caption         =   "Right Little"
            Height          =   495
            Left            =   6840
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btRightRing 
            Caption         =   "Right Ring"
            Height          =   495
            Left            =   5160
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btRightMiddle 
            Caption         =   "Right Middle"
            Height          =   495
            Left            =   3480
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btRightIndex 
            Caption         =   "Right Index"
            Height          =   495
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btRightThumb 
            Caption         =   "Right Thumb"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Label lbCustom3 
         Caption         =   "Custom 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbCustom2 
         Caption         =   "Custom 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbCustom1 
         Caption         =   "Custom 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lbTitle01 
         Alignment       =   2  'Center
         Caption         =   "Face image must be loaded BEFORE starting registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   8415
      End
   End
   Begin VB.Frame fAccount 
      Caption         =   "Account"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtPersonId 
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtAccountId 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbPersonId 
         Caption         =   "Person Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Account Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lbLog 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents simpleID As simpleID
Attribute simpleID.VB_VarHelpID = -1
Private currentOperation As OperationType
Private wsUrl As String
Private apiKey As String

Dim WithEvents Bas64 As base64
Attribute Bas64.VB_VarHelpID = -1

Dim FileName As String
Dim byteBuffer() As Byte
Dim strBuffer As String
Dim flgFile As Integer
Dim encBuffer As String
Dim decBuffer As String

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
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Sub IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any)

Private Const GMEM_MOVEABLE As Long = &H2

Function OpenFile(FileName$, Mode%, RLock%, RecordLen%) As Integer
  Const REPLACEFILE = 1, READAFILE = 2, ADDTOFILE = 3
  Const RANDOMFILE = 4, BINARYFILE = 5
  Const NOLOCK = 0, RDLOCK = 1, WRLOCK = 2, RWLOCK = 3
  Dim FileNum%
  Dim Action%
  FileNum% = FreeFile
  On Error GoTo OpenErrors
  Select Case Mode
    Case REPLACEFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Output Shared As FileNum%
            Case RDLOCK
                Open FileName For Output Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Output Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Output Lock Read Write As FileNum%
        End Select
    Case READAFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Input Shared As FileNum%
            Case RDLOCK
                Open FileName For Input Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Input Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Input Lock Read Write As FileNum%
        End Select
    Case ADDTOFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Append Shared As FileNum%
            Case RDLOCK
                Open FileName For Append Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Append Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Append Lock Read Write As FileNum%
        End Select
    Case RANDOMFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Random Shared As FileNum% Len = RecordLen%
            Case RDLOCK
                Open FileName For Random Lock Read As FileNum% Len = RecordLen%
            Case WRLOCK
                Open FileName For Random Lock Write As FileNum% Len = RecordLen%
            Case RWLOCK
                Open FileName For Random Lock Read Write As FileNum% Len = RecordLen%
        End Select
    Case BINARYFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Binary Shared As FileNum%
            Case RDLOCK
                Open FileName For Binary Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Binary Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Binary Lock Read Write As FileNum%
        End Select
    Case Else
      Exit Function
  End Select
  OpenFile = FileNum%
Exit Function
OpenErrors:
  Action% = FileErrors(Err)
  Select Case Action%
    Case 0
      Resume            'Resumes at line where ERROR occured
    Case 1
        Resume Next     'Resumes at line after ERROR
    Case 2
        OpenFile = 0     'Unrecoverable ERROR-reports error, exits function with error code
        Exit Function
    Case Else
        MsgBox Error$(Err) + vbCrLf + "After line " + Str$(Erl) + vbCrLf + "Program will TERMINATE!"
        'Unrecognized ERROR-reports error and terminates.
        'End
  End Select
End Function

Private Function FileErrors(errVal As Integer) As Integer
'Return Value 0=Resume,              1=Resume Next,
'             2=Unrecoverable Error, 3=Unrecognized Error
Dim msgType%
Dim Msg$
Dim Response%
msgType% = 48
Select Case errVal
    Case 68
      Msg$ = "That device appears Unavailable."
      msgType% = msgType% + 4
    Case 71
      Msg$ = "Insert a Disk in the Drive"
    Case 53
      Msg$ = "Cannot Find File"
      msgType% = msgType% + 5
   Case 57
      Msg$ = "Internal Disk Error."
      msgType% = msgType% + 4
    Case 61
      Msg$ = "Disk is Full.  Continue?"
      msgType% = 35
    Case 64, 52
      Msg$ = "That Filename is Illegal!"
    Case 70
      Msg$ = "File in use by another user!"
      msgType% = msgType% + 5
    Case 76
      Msg$ = "Path does not Exist!"
      msgType% = msgType% + 2
    Case 54
      Msg$ = "Bad File Mode!"
    Case 55
      Msg$ = "File is Already Open."
    Case 62
      Msg$ = "Read Attempt Past End of File."
    Case Else
      FileErrors = 3
      Exit Function
  End Select
  Response% = MsgBox(Msg$, msgType%, "Disk Error")
  Select Case Response%
    Case 1, 4
      FileErrors = 0
    Case 5
      FileErrors = 1
    Case 2, 3
      FileErrors = 2
    Case Else
      FileErrors = 3
  End Select
End Function


Private Sub btCancelRegister_Click()
   simpleID.CancelRegister
   ClearRegisterForm
End Sub

Private Sub btCancelSearchFingerprint_Click()
   simpleID.CancelFingerprintSearch
   pbSearchFingerprint.Picture = Nothing
End Sub

Private Sub btClearRegister_Click()
simpleID.CancelRegister
ClearRegisterForm
End Sub

Private Sub btLeftIndex_Click()
   StartFingerCapture FINGERID_LEFT_INDEX
End Sub

Private Sub btLeftLittle_Click()
   StartFingerCapture FINGERID_LEFT_LITTLE
End Sub

Private Sub btLeftMiddle_Click()
   StartFingerCapture FINGERID_LEFT_MIDDLE
End Sub

Private Sub btLeftRing_Click()
   StartFingerCapture FINGERID_LEFT_RING
End Sub

Private Sub btLeftThumb_Click()
   StartFingerCapture FINGERID_LEFT_THUMB
End Sub

Private Sub btLoadImageRegisterFace_Click()
   CommonDialog.Filter = "All Images (*.jpg,*.jpeg)|*.jpg;*.jpeg;"
   CommonDialog.DialogTitle = "Open Image"
   CommonDialog.ShowOpen

   If CommonDialog.FileName = "" Then Exit Sub

   FileName = CommonDialog.FileName

   pbRegisterFace.Picture = LoadPicture(CommonDialog.FileName)

   pbRegisterFace.ScaleMode = 3
   pbRegisterFace.AutoRedraw = True
   pbRegisterFace.PaintPicture pbRegisterFace.Picture, _
   0, 0, pbRegisterFace.ScaleWidth, pbRegisterFace.ScaleHeight, _
   0, 0, pbRegisterFace.Picture.Width / 26.46, _
   pbRegisterFace.Picture.Height / 26.46
    
   pbRegisterFace.Picture = pbRegisterFace.image

End Sub

Private Sub btLoadImageSearchFace_Click()
   CommonDialog.Filter = "All Images (*.jpg,*.jpeg)|*.jpg;*.jpeg;"
   CommonDialog.DialogTitle = "Open Image"
   CommonDialog.ShowOpen

   If CommonDialog.FileName = "" Then Exit Sub

   FileName = CommonDialog.FileName

   pbSearchFace.Picture = LoadPicture(CommonDialog.FileName)

   pbSearchFace.ScaleMode = 3
   pbSearchFace.AutoRedraw = True
   pbSearchFace.PaintPicture pbSearchFace.Picture, _
   0, 0, pbSearchFace.ScaleWidth, pbSearchFace.ScaleHeight, _
   0, 0, pbSearchFace.Picture.Width / 26.46, _
   pbSearchFace.Picture.Height / 26.46
    
   pbSearchFace.Picture = pbSearchFace.image

End Sub


Private Sub btRegisterFaceOnly_Click()
   Dim FileNum As Integer
   Dim img As String
   Dim crypto As New CryptoBase64
   currentOperation = OperationType_REGISTER
   
   If (pbRegisterFace.Picture = 0) Then
      MsgBox "Please, load an image first"
      Exit Sub
   End If
   
   txtLog.Text = ""

   FileNum = OpenFile(FileName, 5, 0, 80)

   If FileNum = 0 Then
      MsgBox "Could not open File!" & vbCrLf & FileName$
   End If

   ReDim byteBuffer(LOF(FileNum) - 1)

   Get #FileNum, , byteBuffer

   Close #FileNum
   
   crypto.Base64Format = bfmtNone
  
   encBuffer = crypto.Encode(byteBuffer)
   
   encBuffer = Replace(encBuffer, vbNullChar, "")
   
   simpleID.StartRegister txtAccountId.Text, txtPersonId.Text, txtCustom1.Text, txtCustom2.Text, txtCustom3.Text, encBuffer, True
   
End Sub

Private Sub btRightIndex_Click()
   StartFingerCapture FINGERID_RIGHT_INDEX
End Sub

Private Sub btRightLittle_Click()
   StartFingerCapture FINGERID_RIGHT_LITTLE
End Sub

Private Sub btRightMiddle_Click()
   StartFingerCapture FINGERID_RIGHT_MIDDLE
End Sub

Private Sub btRightMiddle1_Click()
   StartFingerCapture FINGERID_RIGHT_MIDDLE
End Sub

Private Sub btRightRing_Click()
   StartFingerCapture FINGERID_RIGHT_RING
End Sub

Private Sub btRightThumb_Click()
   StartFingerCapture FINGERID_RIGHT_THUMB
End Sub

Private Sub btSearchFace_Click()
   Dim FileNum As Integer
   Dim img As String
   Dim crypto As New CryptoBase64
   
   currentOperation = OperationType_SEARCH_FACE
   
   If (pbSearchFace.Picture = 0) Then
      MsgBox "Please, load an image first"
      Exit Sub
   End If
   
   txtLog.Text = ""

   FileNum = OpenFile(FileName, 5, 0, 80)

   If FileNum = 0 Then
      MsgBox "Could not open File!" & vbCrLf & FileName$
   End If

   ReDim byteBuffer(LOF(FileNum) - 1)

   Get #FileNum, , byteBuffer

   Close #FileNum
   
   crypto.Base64Format = bfmtNone
  
   encBuffer = crypto.Encode(byteBuffer)
   
   encBuffer = Replace(encBuffer, vbNullChar, "")
   
   simpleID.SearchFace txtAccountId.Text, txtPersonId.Text, encBuffer
   
End Sub

Private Sub btSearchFingerprint_Click()
   currentOperation = OperationType_SEARCH
   
   simpleID.SearhFingerprint txtAccountId.Text, txtPersonId.Text
   
   pbSearchFingerprint.Picture = Nothing
End Sub

Private Sub btStartRegister_Click()
   Dim FileNum As Integer
   Dim img As String
   Dim crypto As New CryptoBase64
   
   currentOperation = OperationType_REGISTER
   
   txtLog.Text = ""
   
   If Not (pbRegisterFace.Picture = 0) Then
   
      FileNum = OpenFile(FileName, 5, 0, 80)
   
      If FileNum = 0 Then
         MsgBox "Could not open File!" & vbCrLf & FileName$
      End If
   
      ReDim byteBuffer(LOF(FileNum) - 1)
   
      Get #FileNum, , byteBuffer
   
      Close #FileNum
      
      crypto.Base64Format = bfmtNone
     
      encBuffer = crypto.Encode(byteBuffer)
      
      encBuffer = Replace(encBuffer, vbNullChar, "")
   
   Else
      encBuffer = ""
   End If
   
   simpleID.StartRegister txtAccountId.Text, txtPersonId.Text, txtCustom1.Text, txtCustom2.Text, txtCustom3.Text, encBuffer, False

End Sub

Private Sub btSubmitRegister_Click()
   simpleID.SubmitRegister
End Sub

Private Sub Form_Load()
   wsUrl = "URL_SIMPEID"
   apiKey = "API_KEY"
   Set simpleID = New simpleID
   simpleID.Configure wsUrl, apiKey
   
   Set Bas64 = New base64
   
   simpleID.Connect
End Sub

Private Sub btDeletePerson_Click()

simpleID.Delete txtAccountId.Text, txtPersonId.Text

End Sub

Private Sub simpleId_onSimpleIDEvent(ByVal sender As Variant, ByVal e As SimpleIDResponseEventArgs)
   ProcessSimpleIDResponse e.SimpleIDResponse
End Sub

Private Sub StartFingerCapture(ByVal finger As FINGERID)
   simpleID.StartFingerCapture (finger)
End Sub


Private Sub ProcessSimpleIDResponse(ByVal Response As SimpleIDResponse)
   WriteLog Response
   
   Dim returnCode As returnCode
   Dim responseType As responseType
   Dim base64Image As String
   Dim crypto As New CryptoBase64
   Dim image As image
   Dim stdPicture As New stdPicture
   Dim stm As ADODB.stream
   
   returnCode = Response.TransactionInformation.returnCode
   responseType = Response.TransactionInformation.responseType
   
   If (responseType = ResponseType_FINGER_CAPTURE_STATUS) Then
      base64Image = Response.TransactionInformation.Person.FingerprintImage.base64Image
      
      byteBuffer = crypto.Decode(base64Image)
     
      Set stdPicture = ArrayToPicture(byteBuffer(), 0, UBound(byteBuffer) + 1)
      
      If currentOperation = OperationType_SEARCH Then
         
         pbSearchFingerprint.Picture = stdPicture
         
         pbSearchFingerprint.ScaleMode = 3
         pbSearchFingerprint.AutoRedraw = True
         pbSearchFingerprint.PaintPicture pbSearchFingerprint.Picture, _
         0, 0, pbSearchFingerprint.ScaleWidth, pbSearchFingerprint.ScaleHeight, _
         0, 0, pbSearchFingerprint.Picture.Width / 26.46, _
         pbSearchFingerprint.Picture.Height / 26.46
    
         pbSearchFingerprint.Picture = pbSearchFingerprint.image
   
      Else
         If currentOperation = OperationType_REGISTER Then
            pbRegisterFingerprint.Picture = stdPicture
            
            pbRegisterFingerprint.ScaleMode = 3
            pbRegisterFingerprint.AutoRedraw = True
            pbRegisterFingerprint.PaintPicture pbRegisterFingerprint.Picture, _
            0, 0, pbRegisterFingerprint.ScaleWidth, pbRegisterFingerprint.ScaleHeight, _
            0, 0, pbRegisterFingerprint.Picture.Width / 26.46, _
            pbRegisterFingerprint.Picture.Height / 26.46
      
            pbRegisterFingerprint.Picture = pbRegisterFingerprint.image
            
            If Response.TransactionInformation.Person.FingerprintImage.Finished Then
               btSubmitRegister.Enabled = True
            End If
         End If
      End If
   End If
   
   If (responseType = ResponseType_PROCESSING_STATUS) Then
      If Response.TransactionInformation.ProcessingStatus = ProcessingStatus_REGISTRATION_READY Then
         fFingers.Enabled = True
         btCancelRegister.Enabled = True
      End If
   End If
   
End Sub

Private Sub WriteLog(ByVal Response As SimpleIDResponse)

txtLog.Text = txtLog.Text & Response.GetResponse & vbCrLf

End Sub

Private Sub ClearRegisterForm()
   pbRegisterFingerprint.Picture = Nothing
   fFingers.Enabled = False
   btCancelRegister.Enabled = False
   pbRegisterFace.Picture = Nothing
   btRegisterFaceOnly.Enabled = False

End Sub


Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function


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

Public Function ArrayToPicture(inArray() As Byte, Offset As Long, size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: Validate array is not empty before sending here & if array contains
    ' invalid data, the APIs used can lock up your app.
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
            End If
        End If
    End If

End Function
