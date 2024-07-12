VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "VBTRIM"
   ClientHeight    =   5670
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   378
   ScaleMode       =   3  'Ëß¸¾Ù
   ScaleWidth      =   627
   Begin VB.TextBox txtSuffix 
      Appearance      =   0  'Ì×¯Ä
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6420
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "_"
      Top             =   600
      Width           =   1005
   End
   Begin VB.TextBox txtCutSizeY 
      Appearance      =   0  'Ì×¯Ä
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4260
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "100"
      Top             =   600
      Width           =   1005
   End
   Begin VB.TextBox txtCutSizeX 
      Appearance      =   0  'Ì×¯Ä
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "100"
      Top             =   600
      Width           =   1005
   End
   Begin VB.PictureBox picRightSide 
      Align           =   4  '‰E‘µ‚¦
      Height          =   5250
      Left            =   9030
      ScaleHeight     =   5190
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  '‰º‘µ‚¦
      Height          =   420
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   623
      TabIndex        =   23
      Top             =   5250
      Width           =   9405
      Begin VB.TextBox txtStatusRight 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '‚È‚µ
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8910
         TabIndex        =   25
         Text            =   "XXXXXXXXXXXX"
         Top             =   90
         Width           =   2085
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '‚È‚µ
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   30
         TabIndex        =   24
         Text            =   "XXXXXXXXXXXXXX"
         Top             =   90
         Width           =   4485
      End
   End
   Begin VB.Frame frmExp 
      Height          =   915
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   975
      Begin VB.OptionButton optExpBmp 
         Caption         =   "BMP"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Width           =   750
      End
      Begin VB.OptionButton optExpJpg 
         Caption         =   "JPG"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   165
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin VB.PictureBox picSavePicture 
      Appearance      =   0  'Ì×¯Ä
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5655
      ScaleHeight     =   33
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   81
      TabIndex        =   22
      Top             =   2355
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "•Û‘¶ <„£"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   1005
   End
   Begin VB.PictureBox picResize 
      Appearance      =   0  'Ì×¯Ä
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5655
      ScaleHeight     =   33
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   81
      TabIndex        =   21
      Top             =   1755
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmResize 
      Height          =   570
      Left            =   5370
      TabIndex        =   7
      Top             =   -30
      Width           =   3510
      Begin VB.OptionButton optResize3 
         Caption         =   "1/3"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   165
         Width           =   750
      End
      Begin VB.OptionButton optResize4 
         Caption         =   "1/4"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   165
         Width           =   750
      End
      Begin VB.OptionButton optResize2 
         Caption         =   "1/2"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   165
         Width           =   750
      End
      Begin VB.OptionButton optResize 
         Caption         =   "Œ´¡"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin VB.PictureBox picSrcPicture 
      Appearance      =   0  'Ì×¯Ä
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5655
      ScaleHeight     =   33
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   81
      TabIndex        =   19
      Top             =   1155
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Shape shpSrcCut 
         Height          =   360
         Left            =   360
         Top             =   75
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog dlgFileOpen 
      Left            =   6270
      Top             =   2985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "ŠJ‚­"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   1005
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "„@SP"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4260
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   45
      Width           =   1005
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "ƒ@BS"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   1005
   End
   Begin VB.PictureBox picViewPicture 
      Appearance      =   0  'Ì×¯Ä
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   60
      ScaleHeight     =   138
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   266
      TabIndex        =   20
      Top             =   945
      Width           =   3990
      Begin VB.Shape shpViewCutBottom 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   1785
         Top             =   1485
         Width           =   120
      End
      Begin VB.Shape shpViewCutLeftBottom 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   615
         Top             =   1440
         Width           =   120
      End
      Begin VB.Shape shpViewCutRight 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   2925
         Top             =   870
         Width           =   120
      End
      Begin VB.Shape shpViewCutLeft 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   600
         Top             =   915
         Width           =   120
      End
      Begin VB.Shape shpViewCutRightTop 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   2895
         Top             =   285
         Width           =   120
      End
      Begin VB.Shape shpViewCutTop 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   1665
         Top             =   255
         Width           =   120
      End
      Begin VB.Shape shpViewCutLeftTop 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   615
         Top             =   270
         Width           =   120
      End
      Begin VB.Shape shpViewCutRightBottom 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '•s“§–¾
         BorderColor     =   &H00FFFFFF&
         Height          =   120
         Left            =   2925
         Top             =   1455
         Width           =   120
      End
      Begin VB.Shape shpViewCut 
         Height          =   1020
         Left            =   780
         Top             =   420
         Width           =   2115
      End
   End
   Begin VB.FileListBox filFile 
      Height          =   3150
      Left            =   7020
      Pattern         =   "*.jpg"
      TabIndex        =   18
      Top             =   1155
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lblSuffix 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "Ú”öŽ«(&Z)"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5370
      TabIndex        =   16
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label lblCutSizeY 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "ØŽæ‚‚³(&Y)"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3210
      TabIndex        =   14
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label lblCutSizeX 
      Alignment       =   2  '’†‰›‘µ‚¦
      Caption         =   "ØŽæ•(&X)"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1095
      TabIndex        =   12
      Top             =   660
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intZoom As Integer
Dim strPathName As String
Dim sngStartX As Single
Dim sngStartY As Single
Dim blnMove As Boolean
Dim ctrSizeChg(8) As Control
Dim blnSizeChg(8) As Boolean

Private Sub btnNext_Click()
    If filFile.ListIndex < filFile.ListCount - 1 Then
        filFile.ListIndex = filFile.ListIndex + 1
        Call filFile_Click
    End If
End Sub

Private Sub btnOpen_Click()
    Dim strExt As String
    Dim objFso As Object
    Dim objFile As Object
    Dim strFileName As String
    Dim intCount As Integer
    Dim i As Integer
    '
    strExt = ""
    If optExpJpg.Value Then
        strExt = "*.jpg"
    End If
    If optExpBmp.Value Then
        strExt = "*.bmp"
    End If
    dlgFileOpen.FileName = strExt
    filFile.Pattern = strExt
    dlgFileOpen.ShowOpen
    strFileName = Trim(dlgFileOpen.FileName)
    If Len(strFileName) > 0 And strFileName <> strExt Then
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFso.GetFile(strFileName)
        strPathName = objFile.ParentFolder
        Set objFso = Nothing
        Set objFile = Nothing
        filFile.Path = strPathName
        If Right(strPathName, 1) <> "\" Then
            strPathName = strPathName & "\"
        End If
        strFileName = Mid(strFileName, Len(strPathName) + 1)
        intCount = filFile.ListCount
        For i = 0 To intCount
            If filFile.List(i) = strFileName Then
                filFile.ListIndex = i
                Exit For
            End If
        Next
        Call filFile_Click
    End If
End Sub

Private Sub btnPrev_Click()
    If filFile.ListIndex > 0 Then
        filFile.ListIndex = filFile.ListIndex - 1
        Call filFile_Click
    End If
End Sub

Private Sub btnSave_Click()
    Dim strCaption As String
    Dim strFileName As String
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    '
On Error Resume Next
    MkDir strPathName & "SAVE"
On Error GoTo 0
    If Len(strPathName) > 0 Then
        Call picViewPicture_MouseUp(0, 0, 0, 0)
        shpSrcCut.Left = shpViewCut.Left * intZoom
        shpSrcCut.Top = shpViewCut.Top * intZoom
        shpSrcCut.Width = shpViewCut.Width * intZoom
        shpSrcCut.Height = shpViewCut.Height * intZoom
        sngLeft = shpSrcCut.Left
        sngTop = shpSrcCut.Top
        sngWidth = shpSrcCut.Width
        sngHeight = shpSrcCut.Height
        picSavePicture.Cls
        picSavePicture.Width = sngWidth
        picSavePicture.Height = sngHeight
        picSavePicture.ScaleWidth = sngWidth
        picSavePicture.ScaleHeight = sngHeight
        picSavePicture.ScaleMode = vbPixels
        picSavePicture.PaintPicture picSrcPicture.Picture, 0, 0, sngWidth, sngHeight, sngLeft, sngTop, sngWidth, sngHeight
        strFileName = filFile.List(filFile.ListIndex)
        Select Case LCase(Right(strFileName, 4))
        Case ".jpg", ".bmp"
            strFileName = Mid(strFileName, 1, Len(strFileName) - 4)
        End Select
        SavePicture picSavePicture.Image, strPathName & "SAVE\" & strFileName & txtSuffix.Text & ".bmp"
    End If
End Sub

Private Sub filFile_Click()
    picSrcPicture.Picture = LoadPicture(filFile.FileName)
    txtStatus.Text = filFile.FileName
    txtStatusRight.Text = "X=" & picSrcPicture.ScaleWidth & _
                         ",Y=" & picSrcPicture.ScaleHeight
    txtStatus.Refresh
    txtStatusRight.Refresh
    If optResize.Value Then
        picResize.Width = picSrcPicture.Width
        picResize.Height = picSrcPicture.Height
        picResize.ScaleWidth = picSrcPicture.ScaleWidth
        picResize.ScaleHeight = picSrcPicture.ScaleHeight
    Else
        Call frmResize_Click
        picResize.ScaleWidth = Fix(picSrcPicture.ScaleWidth / intZoom)
        picResize.ScaleHeight = Fix(picSrcPicture.ScaleHeight / intZoom)
        picResize.Width = Fix(picSrcPicture.Width / intZoom)
        picResize.Height = Fix(picSrcPicture.Height / intZoom)
        picResize.ScaleMode = vbPixels
    End If
On Error Resume Next
    picResize.Cls
    picResize.PaintPicture picSrcPicture, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight
    picViewPicture.Cls
    picViewPicture.PaintPicture picResize.Image, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight
On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    '
    sngLeft = shpViewCut.Left
    sngTop = shpViewCut.Top
    sngWidth = shpViewCut.Width
    sngHeight = shpViewCut.Height
    '
    Select Case KeyCode
    Case vbKeySpace
        Call btnNext_Click
    Case vbKeyBack
        Call btnPrev_Click
    Case vbKeyReturn
        Call btnSave_Click
    Case vbKeyLeft
        If Shift = 0 Then
            If shpViewCut.Left > 0 Then
                shpViewCut.Left = shpViewCut.Left - 1
            End If
        Else
            If shpViewCut.Width > 24 Then
                shpViewCut.Width = shpViewCut.Width - 1
            End If
        End If
    Case vbKeyRight
        If Shift = 0 Then
            If shpViewCut.Left < picViewPicture.Width - shpViewCut.Width Then
                shpViewCut.Left = shpViewCut.Left + 1
            End If
        Else
            shpViewCut.Width = shpViewCut.Width + 1
        End If
    Case vbKeyUp
        If Shift = 0 Then
            If shpViewCut.Top > 0 Then
                shpViewCut.Top = shpViewCut.Top - 1
            End If
        Else
            If shpViewCut.Height > 24 Then
                shpViewCut.Height = shpViewCut.Height - 1
            End If
        End If
    Case vbKeyDown
        If Shift = 0 Then
            If shpViewCut.Top < picViewPicture.Height - shpViewCut.Height Then
                shpViewCut.Top = shpViewCut.Top + 1
            End If
        Else
            shpViewCut.Height = shpViewCut.Height + 1
        End If
    End Select
    If sngLeft <> shpViewCut.Left Or _
       sngTop <> shpViewCut.Top Or _
       sngWidth <> shpViewCut.Width Or _
       sngHeight <> shpViewCut.Height Then
        Call frmResize_Click
        txtCutSizeX.Text = shpViewCut.Width * intZoom
        txtCutSizeY.Text = shpViewCut.Height * intZoom
        Call picViewPicture_MouseUp(0, 0, 0, 0)
    End If
    Select Case KeyCode
    Case vbKeySpace, vbKeyBack, vbKeyReturn, vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
        KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    '
    strPathName = ""
    sngStartX = 0
    sngStartY = 0
    blnMove = False
    Set ctrSizeChg(0) = shpViewCutLeftTop
    Set ctrSizeChg(1) = shpViewCutTop
    Set ctrSizeChg(2) = shpViewCutRightTop
    Set ctrSizeChg(3) = shpViewCutLeft
    Set ctrSizeChg(4) = shpViewCutRight
    Set ctrSizeChg(5) = shpViewCutLeftBottom
    Set ctrSizeChg(6) = shpViewCutBottom
    Set ctrSizeChg(7) = shpViewCutRightBottom
    For i = 0 To 7
        blnSizeChg(i) = False
    Next
    Call picViewPicture_MouseUp(0, 0, 0, 0)
    txtStatus.Text = ""
    txtStatusRight.Text = ""
    picRightSide.Width = 0
    Call frmResize_Click
    shpViewCut.Left = 30
    shpViewCut.Top = 30
    shpViewCut.Width = Fix(Val(txtCutSizeX.Text) / intZoom)
    shpViewCut.Height = Fix(Val(txtCutSizeY.Text) / intZoom)
    Call picViewPicture_MouseUp(0, 0, 0, 0)
End Sub

Private Sub Form_Resize()
    txtStatus.Width = picStatus.ScaleWidth - txtStatusRight.Width - txtStatus.Left * 3
    txtStatusRight.Left = picStatus.ScaleWidth - txtStatusRight.Width - txtStatus.Left
    picViewPicture.Width = picRightSide.Left - picViewPicture.Left
    picViewPicture.Height = picStatus.Top - picViewPicture.Top
End Sub

Private Sub frmResize_Click()
    intZoom = 1
    If optResize2.Value Then
        intZoom = 2
    End If
    If optResize3.Value Then
        intZoom = 3
    End If
    If optResize4.Value Then
        intZoom = 4
    End If
End Sub

Private Sub optResize_Click()
    Call filFile_Click
    Call frmResize_Click
    shpViewCut.Left = Fix(shpSrcCut.Left / intZoom)
    shpViewCut.Top = Fix(shpSrcCut.Top / intZoom)
    shpViewCut.Width = Fix(shpSrcCut.Width / intZoom)
    shpViewCut.Height = Fix(shpSrcCut.Height / intZoom)
    Call picViewPicture_MouseUp(0, 0, 0, 0)
End Sub

Private Sub optResize2_Click()
    Call optResize_Click
End Sub

Private Sub optResize3_Click()
    Call optResize_Click
End Sub

Private Sub optResize4_Click()
    Call optResize_Click
End Sub

Private Sub picViewPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnInner As Boolean
    Dim i As Integer
    Dim j As Integer
    '
    If Button = vbLeftButton Then
        j = 0
        For i = 0 To 7
            blnInner = (X >= ctrSizeChg(i).Left)
            blnInner = blnInner And (Y >= ctrSizeChg(i).Top)
            blnInner = blnInner And (X <= ctrSizeChg(i).Left + ctrSizeChg(i).Width)
            blnInner = blnInner And (Y <= ctrSizeChg(i).Top + ctrSizeChg(i).Height)
            If blnInner Then
                sngStartX = X - ctrSizeChg(i).Left
                sngStartY = Y - ctrSizeChg(i).Top
                blnSizeChg(i) = True
                For j = 0 To 7
                    ctrSizeChg(j).Visible = False
                Next
                Exit For
            End If
        Next
        If j = 0 Then
            blnInner = (X >= shpViewCut.Left)
            blnInner = blnInner And (Y >= shpViewCut.Top)
            blnInner = blnInner And (X <= shpViewCut.Left + shpViewCut.Width)
            blnInner = blnInner And (Y <= shpViewCut.Top + shpViewCut.Height)
            If blnInner Then
                sngStartX = X - shpViewCut.Left
                sngStartY = Y - shpViewCut.Top
                blnMove = True
                For j = 0 To 7
                    ctrSizeChg(j).Visible = False
                Next
            End If
        End If
    End If
End Sub

Private Sub picViewPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnInner As Boolean
    Dim intMousePointer(8) As Integer
    Dim sngX As Single
    Dim sngY As Single
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim i As Integer
    '
On Error Resume Next
    If Button = vbLeftButton Then
        sngLeft = shpViewCut.Left
        sngTop = shpViewCut.Top
        sngWidth = shpViewCut.Width
        sngHeight = shpViewCut.Height
        sngX = X - sngStartX
        sngY = Y - sngStartY
        If blnSizeChg(0) Then
            shpViewCut.Width = shpViewCut.Width + shpViewCut.Left - sngX
            shpViewCut.Left = sngX
            shpViewCut.Height = shpViewCut.Height + shpViewCut.Top - sngY
            shpViewCut.Top = sngY
        End If
        If blnSizeChg(1) Then
            shpViewCut.Height = shpViewCut.Height + shpViewCut.Top - sngY
            shpViewCut.Top = sngY
        End If
        If blnSizeChg(2) Then
            shpViewCut.Width = sngX - shpViewCut.Left
            shpViewCut.Height = shpViewCut.Height + shpViewCut.Top - sngY
            shpViewCut.Top = sngY
        End If
        If blnSizeChg(3) Then
            shpViewCut.Width = shpViewCut.Width + shpViewCut.Left - sngX
            shpViewCut.Left = sngX
        End If
        If blnSizeChg(4) Then
            shpViewCut.Width = sngX - shpViewCut.Left
        End If
        If blnSizeChg(5) Then
            shpViewCut.Width = shpViewCut.Width + shpViewCut.Left - sngX
            shpViewCut.Left = sngX
            shpViewCut.Height = sngY - shpViewCut.Top
        End If
        If blnSizeChg(6) Then
            shpViewCut.Height = sngY - shpViewCut.Top
        End If
        If blnSizeChg(7) Then
            shpViewCut.Width = sngX - shpViewCut.Left
            shpViewCut.Height = sngY - shpViewCut.Top
        End If
        If shpViewCut.Width < 24 Or _
           shpViewCut.Height < 24 Then
            shpViewCut.Left = sngLeft
            shpViewCut.Top = sngTop
            shpViewCut.Width = sngWidth
            shpViewCut.Height = sngHeight
        End If
        If shpViewCut.Width <> sngWidth Or _
           shpViewCut.Height <> sngHeight Then
            Call frmResize_Click
            txtCutSizeX.Text = shpViewCut.Width * intZoom
            txtCutSizeY.Text = shpViewCut.Height * intZoom
        End If
        If blnMove Then
            shpViewCut.Left = X - sngStartX
            shpViewCut.Top = Y - sngStartY
        End If
    End If
    picViewPicture.MousePointer = vbArrow
    intMousePointer(0) = vbSizeNWSE
    intMousePointer(1) = vbSizeNS
    intMousePointer(2) = vbSizeNESW
    intMousePointer(3) = vbSizeWE
    intMousePointer(4) = vbSizeWE
    intMousePointer(5) = vbSizeNESW
    intMousePointer(6) = vbSizeNS
    intMousePointer(7) = vbSizeNWSE
    For i = 0 To 7
        blnInner = (X >= ctrSizeChg(i).Left)
        blnInner = blnInner And (Y >= ctrSizeChg(i).Top)
        blnInner = blnInner And (X <= ctrSizeChg(i).Left + ctrSizeChg(i).Width)
        blnInner = blnInner And (Y <= ctrSizeChg(i).Top + ctrSizeChg(i).Height)
        If blnInner Then
            picViewPicture.MousePointer = intMousePointer(i)
        End If
    Next
On Error GoTo 0
End Sub

Private Sub picViewPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngWidth As Single
    Dim sngMid As Single
    Dim sngBot As Single
    Dim i As Integer
    '
    blnMove = False
    For i = 0 To 7
        blnSizeChg(i) = False
    Next
    shpViewCutLeftTop.Visible = True
    shpViewCutTop.Visible = True
    shpViewCutRightTop.Visible = True
    shpViewCutLeft.Visible = True
    shpViewCutRight.Visible = True
    shpViewCutLeftBottom.Visible = True
    shpViewCutBottom.Visible = True
    shpViewCutRightBottom.Visible = True
    '
    sngWidth = Fix(shpViewCutLeftTop.Width / 2)
    shpViewCutLeftTop.Top = shpViewCut.Top - sngWidth
    shpViewCutTop.Top = shpViewCut.Top - sngWidth
    shpViewCutRightTop.Top = shpViewCut.Top - sngWidth
    shpViewCutLeftTop.Left = shpViewCut.Left - sngWidth
    shpViewCutLeft.Left = shpViewCut.Left - sngWidth
    shpViewCutLeftBottom.Left = shpViewCut.Left - sngWidth
    sngMid = shpViewCut.Left + Fix(shpViewCut.Width / 2)
    shpViewCutTop.Left = sngMid - sngWidth
    shpViewCutBottom.Left = sngMid - sngWidth
    sngMid = shpViewCut.Top + Fix(shpViewCut.Height / 2)
    shpViewCutLeft.Top = sngMid - sngWidth
    shpViewCutRight.Top = sngMid - sngWidth
    sngBot = shpViewCut.Left + shpViewCut.Width
    shpViewCutRightTop.Left = sngBot - sngWidth
    shpViewCutRight.Left = sngBot - sngWidth
    shpViewCutRightBottom.Left = sngBot - sngWidth
    sngBot = shpViewCut.Top + shpViewCut.Height
    shpViewCutLeftBottom.Top = sngBot - sngWidth
    shpViewCutBottom.Top = sngBot - sngWidth
    shpViewCutRightBottom.Top = sngBot - sngWidth
    Call frmResize_Click
    shpSrcCut.Left = shpViewCut.Left * intZoom
    shpSrcCut.Top = shpViewCut.Top * intZoom
    shpSrcCut.Width = shpViewCut.Width * intZoom
    shpSrcCut.Height = shpViewCut.Height * intZoom
End Sub

Private Sub txtCutSizeX_GotFocus()
    txtCutSizeX.Text = shpSrcCut.Width
    txtCutSizeY.Text = shpSrcCut.Height
End Sub

Private Sub txtCutSizeX_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtCutSizeX_LostFocus
    End If
End Sub

Private Sub txtCutSizeX_LostFocus()
    shpSrcCut.Width = Val(txtCutSizeX.Text)
    shpSrcCut.Height = Val(txtCutSizeY.Text)
    Call frmResize_Click
    shpViewCut.Width = Fix(shpSrcCut.Width / intZoom)
    shpViewCut.Height = Fix(shpSrcCut.Height / intZoom)
    Call picViewPicture_MouseUp(0, 0, 0, 0)
End Sub

Private Sub txtCutSizeY_GotFocus()
    txtCutSizeX_GotFocus
End Sub

Private Sub txtCutSizeY_KeyPress(KeyAscii As Integer)
    Call txtCutSizeX_KeyPress(KeyAscii)
End Sub

Private Sub txtCutSizeY_LostFocus()
    Call txtCutSizeX_LostFocus
End Sub
