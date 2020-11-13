VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9945
   ClientLeft      =   1785
   ClientTop       =   2100
   ClientWidth     =   22350
   LinkTopic       =   "Form1"
   ScaleHeight     =   663
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   1490
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5295
      Left            =   7200
      ScaleHeight     =   349
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
      ScaleWidth      =   381
      TabIndex        =   10
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton Commandok 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5535
      Left            =   16200
      ScaleHeight     =   365
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
      ScaleWidth      =   405
      TabIndex        =   11
      Top             =   840
      Width           =   6135
   End
   Begin VB.CommandButton ComCancel 
      Caption         =   "閉める"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "解像度を設定"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox ComPix 
         Height          =   300
         ItemData        =   "frmMain.frx":0000
         Left            =   2160
         List            =   "frmMain.frx":0019
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmMain.frx":0062
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "解像度を設定："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label labGx 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   3960
         TabIndex        =   3
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmMain.frx":0076
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label labPCPix 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Label labaft 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   16200
      TabIndex        =   12
      Top             =   240
      Width           =   105
   End
   Begin VB.Label labwh 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   9
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":008C
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComCancel_Click() 'アンインストール
   '''エラートラップ
    On Error Resume Next
    Dim strPc As String
    Dim strGx As String
    '''開始ログ
    Output.Log App.EXEName, "frmMain", "ComCancel_Click", eLg08MethodStart, "×ボタン"
    strPc = frmMain.labPCPix.Caption
    strGx = frmMain.labGx.Caption
'    If strPc <> strGx Then
'        'comMsg.Box "PCの解像度と現在の解像度不一致です。確認してください。", vbCritical, "エラー"
'        'comMsg.Box "検査項目が存在しません。マスタ(Order_ETreat)を確認してください。", , "エラー"
'        comMsg.Box "想定外のタグプロパティが設定されています。検査オーダのテンプレートの設定を確認してください。", vbCritical, "エラー"
'    Else
'
'    End If
    frmMain.Hide
    Output.Log App.EXEName, "frmMain", "ComPix_Click", eLg09MethodEnd, "フォーム解放"
   
ErrTrap:
    '''エラーログ
    Output.Log App.EXEName, "frmSub", "ComCancel_Click", eLg07Trace, "ComCancel_Clickエラー", Err.Number, Err.Description
    'Resume Next
End Sub

Private Sub Commandok_Click()
    Dim HEIPIX As Variant
    Dim WDIPIX As Variant
    Dim strPix As String
    Dim iPos As Integer
    Dim i As Integer
    Dim j As Integer
    Dim varBit As Variant
    Dim strTemp As String
    Dim str As String
    Dim curDire As String
    Dim curFile As String
    Dim pixelColor As Variant
    Dim newColor As Variant
    
    Dim R As Variant
    Dim G As Variant
    Dim B As Variant
    Picture2.Picture = LoadPicture("E:\v2-fd813aa825c6a503410533d674c01581_540x450 (1).jpeg")
    strPix = frmMain.ComPix.Text
    iPos = InStr(1, strPix, "*", vbTextCompare)
    
    
    
    HEIPIX = Trim(Mid(strPix, 1, iPos - 1))
    WDIPIX = Trim(Mid(strPix, iPos + 1, Len(frmMain.ComPix.Text)))
    
    
    
    
    Picture2.ScaleHeight = HEIPIX
    Picture2.ScaleWidth = WDIPIX
'    Picture2.Height = HEIPIX / 96 * 1440
'    Picture2.Width = WDIPIX / 96 * 1440
'    For i = 0 To HEIPIX - 1
'        For j = 0 To WDIPIX - 1
'            Debug.Print Picture2.Point(i, j)
'        Next j
'    Next i
'    For i = 0 To Picture2.Width - 1
'        For j = 0 To Picture2.Height - 1
'                R = (Picture2.Point(i, j) And &HFF) Mod 256
'                G = (Picture2.Point(i, j) And &HFF00) Mod 256
'                B = (Picture2.Point(i, j) And &HFF0000) Mod 256
'                Picture2.PSet (X, Y), RGB(R * 0.5, G * 0.5, B * 0.5)
'        Next j
'    Next i
 
        ' Set the PictureBox to display the image.
    curDire = "C:\egmain-ex\Data\Xtml_S\ExTSchema\pix"
    curFile = curDire + "\pix.JPG"
    Kill curFile
    If Len(Dir$(curDire, vbDirectory)) = 0 Then MkDir curDire
    
    If Len(Dir$(curFile)) > 0 Then Kill curFile
    SavePicture Picture2.Image, curFile
    
   'varBit = Int(FileLen("C:\egmain-ex\Data\Xtml_S\ExTSchema\pix\pix.JPG") / 1024)
    varBit = Int(HEIPIX * WDIPIX * 24 / 1024 / 8)
    
    
    frmMain.labaft.Caption = Picture2.ScaleHeight & "*" & Picture2.ScaleWidth & vbCrLf & "メモリー量:" & varBit & "KB"
End Sub

Private Sub ComPix_Click()
  '''エラートラップ
    On Error Resume Next
    '''開始ログ
    Output.Log App.EXEName, "frmMain", "ComPix_Click", eLg08MethodStart, "×ボタン"
    frmMain.labGx.Caption = frmMain.ComPix.Text
    Output.Log App.EXEName, "frmMain", "ComPix_Click", eLg09MethodEnd, "フォーム解放"
ErrTrap:
    '''エラーログ
    Output.Log App.EXEName, "frmSub", "ComPix_Click", eLg07Trace, "ComPix_Clickエラー", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
'''エラートラップ
    On Error Resume Next
    Dim strPixX As String
    Dim strPixY As String
    Dim varHei As Variant
    Dim varWdi As Variant
    Dim varBit As Variant

    '''開始ログ
    Output.Log App.EXEName, "frmMain", "ComPix_Click", eLg08MethodStart, "×ボタン"
    Picture1.Picture = LoadPicture("E:\v2-fd813aa825c6a503410533d674c01581_540x450 (1).jpeg")
    strPixX = Screen.Width / Screen.TwipsPerPixelX
    strPixY = Screen.Height / Screen.TwipsPerPixelY
    varHei = Picture1.ScaleHeight
    varWdi = Picture1.ScaleWidth
    varBit = Int(varHei * varWdi * 24 / 1024 / 8)
    labwh.Caption = varHei & "*" & varWdi & vbCrLf & "メモリー量:" _
           & Int(varHei * varWdi * 24 / 1024 / 8) & "KB"
    frmMain.labPCPix.Caption = strPixX & "*" & strPixY
    
    Output.Log App.EXEName, "frmMain", "ComPix_Click", eLg09MethodEnd, "フォーム解放"
ErrTrap:
    '''エラーログ
    Output.Log App.EXEName, "frmSub", "ComPix_Click", eLg07Trace, "ComPix_Clickエラー", Err.Number, Err.Description
End Sub

