VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Bitmap: New"
   ClientHeight    =   4980
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlTools 
      Left            =   4320
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":025C
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":036E
            Key             =   "SelColor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0480
            Key             =   "Erase"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0592
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06A4
            Key             =   "FCirc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07B6
            Key             =   "Circ"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08C8
            Key             =   "Flood"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DA
            Key             =   "FRect"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AEC
            Key             =   "Rect"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BFE
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D10
            Key             =   "Capture"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBMP 
      Caption         =   "Bitmap----------Toolbox"
      Height          =   615
      Left            =   5640
      TabIndex        =   13
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox picTB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C8D0D4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1200
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   15
         Top             =   240
         Width           =   240
      End
      Begin VB.PictureBox picBMP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C8D0D4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   14
         Top             =   240
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgMnu 
      Left            =   3600
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A46
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CA0
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EFA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2154
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23AE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2608
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2862
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ABC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D16
            Key             =   "PasteAll"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2040
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picWork 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame fraCurr 
      Caption         =   "Color Selection"
      Height          =   3945
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      Begin VB.OptionButton optPal 
         Caption         =   "16 Colors"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   17
         Top             =   3660
         Width           =   1215
      End
      Begin VB.OptionButton optPal 
         Caption         =   "Many Colors"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   3660
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.PictureBox picPal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   240
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   12
         ToolTipText     =   "Right or Left click to select color, Double click for custom color"
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label1 
         Caption         =   "Transparent Color"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblTran 
         BackColor       =   &H00C8D0D4&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblPal 
         Alignment       =   2  'Center
         Caption         =   "R0,G0,B0"
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblRInfo 
         Caption         =   "Right: R0,G0,B255"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblLInfo 
         Caption         =   "Left:   R255,G0,B0"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblRight 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblLeft 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C8D0D4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   480
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C8D0D4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   480
         Begin VB.Shape shSel 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Shape shRect 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Left            =   600
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   16
         X2              =   16
         Y1              =   88
         Y2              =   152
      End
      Begin VB.Shape shCirc 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   615
         Left            =   1800
         Shape           =   2  'Oval
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar TBTools 
      Height          =   3690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   6509
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select"
            Object.ToolTipText     =   "Selection Rectangle"
            ImageKey        =   "Select"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pencil"
            Object.ToolTipText     =   "Pencil"
            ImageKey        =   "Pencil"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Line"
            Object.ToolTipText     =   "Line"
            ImageKey        =   "Line"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rect"
            Object.ToolTipText     =   "Rectangle"
            ImageKey        =   "Rect"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FRect"
            Object.ToolTipText     =   "Filled Rectange"
            ImageKey        =   "FRect"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Circ"
            Object.ToolTipText     =   "Circle"
            ImageKey        =   "Circ"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FCirc"
            Object.ToolTipText     =   "Filled Circle"
            ImageKey        =   "FCirc"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SelColor"
            Object.ToolTipText     =   "Color Selection"
            ImageKey        =   "SelColor"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flood"
            Object.ToolTipText     =   "Flood Fill"
            ImageKey        =   "Flood"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Text"
            Object.ToolTipText     =   "Text"
            ImageKey        =   "Text"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Erase"
            Object.ToolTipText     =   "Erase"
            ImageKey        =   "Erase"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4605
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9340
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFArr 
         Caption         =   "New"
         Index           =   0
      End
      Begin VB.Menu mnuFArr 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu mnuFArr 
         Caption         =   "Save"
         Index           =   2
      End
      Begin VB.Menu mnuFArr 
         Caption         =   "SaveAs"
         Index           =   3
      End
      Begin VB.Menu mnuFArr 
         Caption         =   "Paste Clipboard"
         Index           =   4
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu MRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEArr 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEArr 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEArr 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEArr 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEArr 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuEArr 
         Caption         =   "Redo"
         Enabled         =   0   'False
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Included a dummy UserControl, so that after
' saving & quitting the app you can assign a
' TB Bitmap to it, close its designer window
' and code window and see how it looks in the IDE

Option Explicit
Private Const Pix As Long = 20
Private Const PixH As Long = 10
Private i As Long, j As Long, k As Long
Private Gx As Long, Gy As Long
Private Ix As Long, Iy As Long
Private SGx As Long, SGy As Long
Private PSx As Long, PSy As Long
Private SIx As Long, SIy As Long
Private PalX As Long, PalY As Long, PalB As Integer
Private Pasted As Boolean
Private CurrTool As Long
Private CurrColor As Long
Private CurrFileName As String
Private KeyVal As Long
Private MruArr(1 To 5) As String 'for MRU
Private MruCnt As Long
Private cUndo As New Collection
Private cRedo As New Collection
Private Btn As MSComctlLib.Button
Private oIni As cIni
Private Frm As Form

Private Sub Form_Load()
 Init
 Show
 DoEvents
End Sub
Private Sub Form_Unload(Cancel As Integer)
 SaveSettings
 Set cUndo = Nothing
 Set cRedo = Nothing
 Set Frm = Nothing
 Set Btn = Nothing
End Sub

Private Sub lblTran_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then
  lblRight.BackColor = lblTran.BackColor
 Else
  lblLeft.BackColor = lblTran.BackColor
 End If
End Sub

Private Sub mnuFArr_Click(Index As Integer)
 Select Case Index
  Case 0 'new
   Set picBMP.Picture = LoadPicture
   Pic2Grid
   CurrFileName = ""
  Case 1 'open
   CurrFileName = OpenFileName()
   'in case user cancelled
   If Len(CurrFileName) = 0 Then Exit Sub
   DoLoad True
   ShowTBBitMap
   UpdateUndo
   UpdateMRU
  Case 2 'save
   If Len(CurrFileName) = 0 Then
    CurrFileName = SaveFileName()
   End If
   If Len(CurrFileName) = 0 Then Exit Sub
   SavePicture picBMP.Image, CurrFileName
   UpdateMRU
  Case 3 'save as
   'curious, never seen an icon or bmp for this
   CurrFileName = SaveFileName()
   'in case user cancelled
   If Len(CurrFileName) = 0 Then Exit Sub
   SavePicture picBMP.Image, CurrFileName
   UpdateMRU
  Case 4 'paste clipboard
   DoLoad False
   CurrFileName = ""
 End Select
 If Len(CurrFileName) Then
  Caption = "Menu Bitmap: " & CurrFileName
 Else
  Caption = "Menu Bitmap: New"
 End If
End Sub

Private Sub optPal_Click(Index As Integer)
 If Index = 0 Then
  DrawPalette
 Else
  DrawPalette True
 End If
End Sub

Private Sub picPal_DblClick()
'to add from the color dialog
 Dim idx As Long, oc As Long
 oc = GetPixel(picPal.hdc, PalX, PalY)
 With CD
  .CancelError = True
  .Flags = cdlCCFullOpen Or cdlCCRGBInit
  .Color = oc
  On Error GoTo Canx
  .ShowColor
  idx = 16 * (PalY \ 10) + PalX \ 10
  Pal(idx) = .Color
  'user has added a new color
  'so change the pic to reflect it
  'otherwise the pic won't be in sync with the palette
  For PalY = 0 To 12
   For PalX = 0 To 12
    If GetPixel(picBMP.hdc, PalX, PalY) = oc Then
     SetPixelV picBMP.hdc, PalX, PalY, .Color
    End If
   Next
  Next
  DrawPalette
  If PalB = vbLeftButton Then
   lblLeft.BackColor = .Color
   lblLInfo.Caption = "Left:   R" & RedV(.Color) & ",G" & GrnV(.Color) & ",B" & BluV(.Color)
  Else
   lblRight.BackColor = .Color
   lblRInfo.Caption = "Right: R" & RedV(.Color) & ",G" & GrnV(.Color) & ",B" & BluV(.Color)
  End If
 End With
Canx:
End Sub

Private Sub picPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim MC As Long
 'for the double click
 PalX = X: PalY = Y: PalB = Button
 MC = GetPixel(picPal.hdc, X, Y)
 If Button = vbLeftButton Then
  lblLeft.BackColor = MC
  lblLInfo.Caption = "Left:   R" & RedV(MC) & ",G" & GrnV(MC) & ",B" & BluV(MC)
 Else
  lblRight.BackColor = MC
  lblRInfo.Caption = "Right: R" & RedV(MC) & ",G" & GrnV(MC) & ",B" & BluV(MC)
 End If
End Sub

Private Sub picPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'display color info
 Dim MC As Long
 MC = GetPixel(picPal.hdc, X, Y)
 lblPal.Caption = "R" & RedV(MC) & ",G" & GrnV(MC) & ",B" & BluV(MC)
End Sub

Private Sub TBTools_ButtonClick(ByVal Button As MSComctlLib.Button)
 For i = 1 To TBTools.Buttons.Count
  TBTools.Buttons(i).Value = tbrUnpressed
 Next
 TBTools.Buttons(Button.Index).Value = tbrPressed
 TBTools.Refresh
 CurrTool = Button.Index
 'show user some help for the tools
 With SB.Panels(3)
  Select Case Button.Index
   Case TPencil
    .Text = "Free draw"
   Case TRect, TFRect
    .Text = "Hold a shift key for a square"
   Case TText
    .Text = "Click the grid to position the text"
   Case TLine
    .Text = "Hold a shift key for Hor/Vert Line"
   Case TCirc, TFCirc
    .Text = "Hold a shift key for Circle"
   Case TErase
    .Text = "Free draw in white"
   Case TSelect
    .Text = "Selection tool for Cut, Copy, Paste"
   Case TFlood
    .Text = "Flood an area with selected color"
   Case TSelColor
    .Text = "Click the grid to get desired color"
  End Select
 End With
End Sub
Private Sub mnuEArr_Click(Index As Integer)
 Select Case Index
  Case 0 'cut
   PasteXY picSel.Left \ Pix, picSel.Top \ Pix, True
   picSel.Visible = False
   mnuEArr(0).Enabled = False
   mnuEArr(1).Enabled = False
   mnuEArr(2).Enabled = True
   UpdateUndo
  Case 1 'copy
   picSel.Visible = False
   mnuEArr(0).Enabled = False
   mnuEArr(1).Enabled = False
   mnuEArr(2).Enabled = True
  Case 2 'paste
   picSel.Move 0, 0
   picSel.Visible = True
   Pasted = True
  Case 4 'undo
   DoUnDo
  Case 5 'redo
   DoReDo
 End Select
End Sub

Private Sub MRU_Click(Index As Integer)
 If FileExists(MRU(Index).Caption) Then
  CurrFileName = MRU(Index).Caption
  DoLoad True
 End If
End Sub

Private Sub picSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 PSx = X: PSy = Y
 If Pasted = False Then
  PasteXY picSel.Left \ Pix, picSel.Top \ Pix, True
 End If
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim nx As Long, ny As Long
 If Button Then
  With picSel
   nx = ((.Left + (X - PSx)) \ Pix) * Pix
   ny = ((.Top + (Y - PSy)) \ Pix) * Pix
   .Move nx, ny
  End With
 End If
End Sub

Private Sub picSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 PasteXY picSel.Left \ Pix, picSel.Top \ Pix, False
 picSel.Visible = False
 mnuEArr(0).Enabled = False
 mnuEArr(1).Enabled = False
 mnuEArr(2).Enabled = True
 ShowTBBitMap
 UpdateUndo
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Gx = (X \ Pix) * Pix: Gy = (Y \ Pix) * Pix
 Ix = X \ Pix: Iy = Y \ Pix
 SGx = Gx: SGy = Gy 'save these for mousemove & mouseup
 SIx = Ix: SIy = Iy
 'get the drawing color
 If Button = vbRightButton Then
  CurrColor = lblRight.BackColor
 Else
  CurrColor = lblLeft.BackColor
 End If
 Select Case CurrTool
  Case TPencil
   Call picGrid_MouseMove(Button, Shift, X, Y)
  Case TErase
   Call picGrid_MouseMove(Button, Shift, X, Y)
  Case TLine 'use the Line control to delineate the line
   'make it start in the center of the box
   Lin.X1 = Gx + PixH: Lin.X2 = Gx + PixH
   Lin.Y1 = Gy + PixH: Lin.Y2 = Gy + PixH
   Lin.Visible = True
  Case TRect, TFRect, TSelect
   Pasted = False 'in case we're selecting
   shRect.Move Gx + PixH, Gy + PixH, 0, 0
   shRect.Visible = True
  Case TCirc, TFCirc
   shCirc.Move Gx + PixH, Gy + PixH, 0, 0
   shCirc.Visible = True
  Case TSelColor
   If Button = vbRightButton Then
    lblRight.BackColor = GetPixel(picBMP.hdc, Ix, Iy)
    lblRInfo.Caption = "Right: R" & RedV(lblRight.BackColor) & ",G" & GrnV(lblRight.BackColor) & ",B" & BluV(lblRight.BackColor)
   Else
    lblLeft.BackColor = GetPixel(picBMP.hdc, Ix, Iy)
    lblLInfo.Caption = "Left:   R" & RedV(lblLeft.BackColor) & ",G" & GrnV(lblLeft.BackColor) & ",B" & BluV(lblLeft.BackColor)
   End If
  Case TFlood
   picBMP.FillStyle = vbFSSolid
   picBMP.FillColor = CurrColor 'color to fill with
   ExtFloodFill picBMP.hdc, Ix, Iy, GetPixel(picBMP.hdc, Ix, Iy), 1
   Pic2Grid
 End Select
 picBMP.Refresh
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim MC As Long
 Gx = (X \ Pix) * Pix: Gy = (Y \ Pix) * Pix
 Ix = X \ Pix: Iy = Y \ Pix
 MC = GetPixel(picBMP.hdc, Ix, Iy)
 SB.Panels(1).Text = Left$("X: " & Right$(" " & Ix, 2) & "   ", 5) & _
   " Y: " & Right$(" " & Iy, 2)
 SB.Panels(2).Text = "R" & RedV(MC) & ",G" & GrnV(MC) & ",B" & BluV(MC)
 If Button Then 'dragging the shape or freedrawing
  Select Case CurrTool
   Case TPencil
    picGrid.Line (Gx + 1, Gy + 1)-(Gx + Pix - 1, Gy + Pix - 1), CurrColor, BF
    SetPixelV picBMP.hdc, Ix, Iy, CurrColor
   Case TErase
    picGrid.Line (Gx + 1, Gy + 1)-(Gx + Pix - 1, Gy + Pix - 1), picGrid.BackColor, BF
    SetPixelV picBMP.hdc, Ix, Iy, picGrid.BackColor
   Case TLine
    If Shift Then 'horizontal or vertical line
     If Abs(Gx - SGx) > Abs(Gy - SGy) Then
      Gy = SGy
     Else
      Gx = SGx
     End If
    End If
    With Lin
     'size the line control
     .X1 = SGx + PixH
     .X2 = Gx + PixH
     .Y1 = SGy + PixH
     .Y2 = Gy + PixH
    End With
   Case TRect, TFRect, TSelect
    With shRect
     'a little math here to
     'allow the rect to be drawn left to right or vice versa
     ' or top to bottom or vice versa
     If Gx - SGx < 0 And Gy - SGy < 0 Then
      .Left = Gx + PixH
      .Top = Gy + PixH
     ElseIf Gx - SGx < 0 Then
      .Left = Gx + PixH
     ElseIf Gy - SGy < 0 Then
      .Top = Gy + PixH
     Else
      .Left = SGx + PixH
      .Top = SGy + PixH
     End If
     If Shift Then 'for a square
      .Width = Abs(Gx - SGx)
      .Height = Abs(Gx - SGx)
     Else
      .Width = Abs(Gx - SGx)
      .Height = Abs(Gy - SGy)
     End If
    End With

   Case TCirc, TFCirc
    With shCirc
     If Gx - SGx < 0 And Gy - SGy < 0 Then
      .Left = Gx + PixH
      .Top = Gy + PixH
     ElseIf Gx - SGx < 0 Then
      .Left = Gx + PixH
     ElseIf Gy - SGy < 0 Then
      .Top = Gy + PixH
     Else
      .Left = SGx + PixH
      .Top = SGy + PixH
     End If
     If Shift Then 'for a circle
      .Width = Abs(Gx - SGx)
      .Height = Abs(Gx - SGx)
     Else
      .Width = Abs(Gx - SGx)
      .Height = Abs(Gy - SGy)
     End If
    End With

  End Select
 End If
 picBMP.Refresh
End Sub

Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Gx = (X \ Pix) * Pix: Gy = (Y \ Pix) * Pix
 Ix = X \ Pix: Iy = Y \ Pix
 Select Case CurrTool
  Case TSelect
   shRect.Visible = False
   If shRect.Width < 2 Or shRect.Height < 2 Then
    picSel.Cls
    picSel.Visible = False
    Exit Sub
   End If
   picSel.Visible = True
   'size picSel to the shape size
   picSel.Move shRect.Left - PixH, shRect.Top - PixH, shRect.Width, shRect.Height
   shSel.Move 1, 1, picSel.Width - 2, picSel.Height - 2
   picSel.Cls
   'and copy the selected part of the grid to it
   BitBlt picSel.hdc, 0, 0, picSel.ScaleWidth, picSel.ScaleHeight, _
     picGrid.hdc, SGx, SGy, vbSrcCopy
   mnuEArr(0).Enabled = True
   mnuEArr(1).Enabled = True
  Case TLine
   With Lin
    .Visible = False
    picBMP.Line (.X1 \ Pix, .Y1 \ Pix)-(.X2 \ Pix, .Y2 \ Pix), CurrColor
    'line does not get the last x,y pixel
    SetPixelV picBMP.hdc, .X2 \ Pix, .Y2 \ Pix, CurrColor
   End With
   Pic2Grid
  Case TRect, TFRect
   With shRect
    .Visible = False
    If CurrTool = TRect Then
     picBMP.Line (.Left \ Pix, .Top \ Pix)-((.Left + .Width) \ Pix, (.Top + .Height) \ Pix), CurrColor, B
    Else
     picBMP.Line (.Left \ Pix, .Top \ Pix)-((.Left + .Width) \ Pix, (.Top + .Height) \ Pix), CurrColor, BF
    End If
   End With
   Pic2Grid
  Case TCirc, TFCirc
   With shCirc
    .Visible = False
    'for the ellipse call below
    picBMP.ForeColor = CurrColor
    If CurrTool = TCirc Then
     picBMP.FillStyle = vbFSTransparent
    Else
     picBMP.FillStyle = vbFSSolid
     picBMP.FillColor = CurrColor
    End If
    Ellipse picBMP.hdc, .Left \ Pix, .Top \ Pix, (.Left + .Width) \ Pix, (.Top + .Height) \ Pix
   End With
   Pic2Grid
  Case TSelColor
   SetPencil
  Case TText
   Set Frm = New frmText
   Frm.Move Left + picGrid.Left + picGrid.Width, Top + picGrid.Top + picGrid.Height \ 2
   Frm.Show vbModal, Me
   'retrieve the selected font items
   picBMP.FontName = gFontName
   picBMP.FontBold = gFontBold
   picBMP.FontItalic = gFontItalic
   picBMP.FontSize = gFontSize
   picBMP.ForeColor = CurrColor
   picBMP.CurrentX = SIx 'saved from mousedown
   picBMP.CurrentY = SIy '  "       "
   picBMP.Print gText
   Pic2Grid
 End Select
 picBMP.Refresh
 ShowT
 ShowTBBitMap
 UpdateUndo 'mouseup so save for undo
End Sub
Private Sub ShowTBBitMap()
'shows the bitmap as viewed in the IDE
 Dim TC As Long, C As Long
 Dim X As Long, Y As Long
 TC = GetPixel(picBMP.hdc, 0, 14) 'the transparent color
 lblTran.BackColor = TC
 For Y = 0 To 14
  For X = 0 To 15
   C = GetPixel(picBMP.hdc, X, Y)
   If C = TC Then
    SetPixelV picTB.hdc, X, Y, picTB.BackColor
   Else
    SetPixelV picTB.hdc, X, Y, C
   End If
  Next
 Next
 picTB.Refresh
End Sub
Private Function Invert(ByVal clr As Long) As Long
 Dim RedV&, GrnV&, BluV&
 RedV = clr And &HFF&
 GrnV = (clr And &HFF00&) \ 256
 BluV = (clr And &HFF0000) \ 65536
 Invert = RGB(255 - RedV, 255 - GrnV, 255 - BluV)
End Function

Private Sub DeleteCollections()
 Set cUndo = New Collection
 Set cRedo = New Collection
 KeyVal = 0
 mnuEArr(4).Enabled = False 'undo
 mnuEArr(5).Enabled = False 'redo
End Sub
Private Sub UpdateUndo()
'save the current pic in the undo coll
 KeyVal = KeyVal + 1 'just a unique no for coll
 picBMP.Picture = picBMP.Image
 cUndo.Add picBMP.Picture, CStr(KeyVal)
 mnuEArr(4).Enabled = cUndo.Count > 1
 mnuEArr(5).Enabled = cRedo.Count > 0
End Sub
Private Sub DoUnDo()
 cRedo.Add cUndo.Item(cUndo.Count)
 cUndo.Remove cUndo.Count
 picBMP.Picture = cUndo.Item(cUndo.Count)
 picBMP.Refresh
 mnuEArr(4).Enabled = cUndo.Count > 1
 mnuEArr(5).Enabled = cRedo.Count > 0
 Pic2Grid
End Sub
Private Sub DoReDo()
 cUndo.Add cRedo.Item(cRedo.Count)
 cRedo.Remove cRedo.Count
 picBMP.Picture = cUndo.Item(cUndo.Count)
 picBMP.Refresh
 mnuEArr(4).Enabled = cUndo.Count > 1
 mnuEArr(5).Enabled = cRedo.Count > 0
 Pic2Grid
End Sub
Private Sub FixColors()
 Dim LP As LOGPALETTE
 Dim X As Long
 Dim Y As Long
 Dim C As Long
 Dim n As Long
 Dim hPal As Long
 Dim bChanged As Boolean
 With LP
  CopyMemory .palPalEntry(0), Pal(0), 1024
  .palNumEntries = 256
  .palVersion = &H300
 End With
 hPal = CreatePalette(LP)
 For Y = 0 To 12
  For X = 0 To 12
   C = GetPixel(picBMP.hdc, X, Y)
   If InPal(C) = False Then
    'color is not in our palette
    'so get the nearest color index
    n = GetNearestPaletteIndex(hPal, C)
    '    Debug.Print n, Hex$(Pal(n)), Hex$(c)
    'and put it in our palette
    Pal(n) = C
    bChanged = True
   End If
  Next
 Next
 DeleteObject hPal
 If bChanged Then
  DrawPalette
  optPal(0).Value = True
 End If
End Sub
' search palette for given color
Private Function InPal(ByVal clr As Long) As Boolean
 For i = 0 To 255
  If clr = Pal(i) Then
   InPal = True: Exit Function
  End If
 Next
End Function

Private Sub Init()
'put bitmaps in the menus
 SetMenuItemBMP Me.hwnd, 0, 0, imgMnu.ListImages("New").Picture
 SetMenuItemBMP Me.hwnd, 0, 1, imgMnu.ListImages("Open").Picture
 SetMenuItemBMP Me.hwnd, 0, 2, imgMnu.ListImages("Save").Picture
 SetMenuItemBMP Me.hwnd, 0, 4, imgMnu.ListImages("PasteAll").Picture

 SetMenuItemBMP Me.hwnd, 1, 0, imgMnu.ListImages("Cut").Picture
 SetMenuItemBMP Me.hwnd, 1, 1, imgMnu.ListImages("Copy").Picture
 SetMenuItemBMP Me.hwnd, 1, 2, imgMnu.ListImages("Paste").Picture
 SetMenuItemBMP Me.hwnd, 1, 4, imgMnu.ListImages("Undo").Picture
 SetMenuItemBMP Me.hwnd, 1, 5, imgMnu.ListImages("Redo").Picture
 Set oIni = New cIni
 oIni.Path = App.Path & "\TB.ini"
 oIni.Section = "MRU"
 LoadSettings 'get MRU list
 GetPal 'load & draw the color palette
 DrawPalette 'default user colors
 DrawGrid
 SetPencil
End Sub

Private Sub DrawPalette(Optional ByVal Basic As Boolean)
 Dim X As Long, Y As Long, k As Long
 With picPal
  If Basic Then
   For Y = 0 To .ScaleHeight - 1 Step 40
    For X = 0 To .ScaleWidth - 1 Step 40
     picPal.Line (X, Y)-(X + 40, Y + 40), QBColor(k), BF
     k = k + 1
    Next
   Next
  Else
   For Y = 0 To .ScaleHeight - 1 Step 10
    For X = 0 To .ScaleWidth - 1 Step 10
     picPal.Line (X, Y)-(X + 10, Y + 10), Pal(k), BF
     k = k + 1
    Next
   Next
  End If
 End With
End Sub

Private Sub DrawGrid()
 With picGrid
  For i = 0 To .ScaleWidth Step Pix
   picGrid.Line (0, i)-(.ScaleWidth, i)
   picGrid.Line (i, 0)-(i, .ScaleHeight)
  Next
  ShowT
 End With
End Sub
Private Sub ShowT()
 With picGrid
  .CurrentX = 5
  .CurrentY = .ScaleHeight - .TextHeight("T")
  .ForeColor = Invert(picBMP.Point(0, 14))
  picGrid.Print "T"
  .ForeColor = vbBlack
 End With
End Sub
Private Sub Pic2Grid()
'expand the bitmap pic to grid size
 picGrid.PaintPicture picBMP.Image, 0, 0, picGrid.ScaleWidth, picGrid.ScaleHeight
 'then draw lines on it
 DrawGrid
End Sub
Private Sub DoLoad(ByVal Pic As Boolean)
 Dim SP As StdPicture
 Dim H As Long, W As Long
 Dim Msg As String
 If Pic Then
  Set SP = LoadPicture(CurrFileName)
 Else
  Set SP = Clipboard.GetData(vbCFBitmap)
  If SP.Handle = 0 Then
   MsgBox "No picture on clipboard"
   Exit Sub
  End If
 End If
 'check the size
 W = CLng(ScaleX(SP.Width, vbHimetric, vbPixels))
 H = CLng(ScaleX(SP.Height, vbHimetric, vbPixels))
 If W > 16 Or H > 15 Then
  Msg = "This image is not 16x15" & vbNewLine & _
    "Select Yes to Crop Portion" & vbNewLine & _
    "Or No to Size the Image to 16x15"
  If MsgBox(Msg, vbYesNo) = vbYes Then
   Set Frm = New frmCrop
   Set Frm.picSrc.Picture = SP
   'try to size the form to fit the picture

   Frm.Show vbModal
  Else
   'here just stretchblt the pic to fit
   Set picWork.Picture = SP 'picWork has AutoSize = True
   'allegedly produces better quality stretches
   SetStretchBltMode picBMP.hdc, HALFTONE
   StretchBlt picBMP.hdc, 0, 0, 16, 15, _
     picWork.hdc, 0, 0, picWork.ScaleWidth, picWork.ScaleHeight, vbSrcCopy
  End If
 Else
  'pic is ok
  Set picBMP.Picture = SP
 End If
 GetPal 'reload the default palette
 FixColors 'change any colors that don't match
 Pic2Grid
 DeleteCollections 'reset undo/redo
 UpdateUndo 'in case user wants to undo this
 ShowTBBitMap
End Sub
Private Function OpenFileName() As String
 With CD
  .CancelError = True
  .Filter = "Picture Files|*.bmp;*.jpg;*.ico;*.gif"
  On Error GoTo Canx
  .ShowOpen
  OpenFileName = .FileName
 End With
Canx:
End Function
Private Function SaveFileName() As String
 With CD
  .CancelError = True
  .Filter = "Bitmap Files|*.bmp"
  .DefaultExt = "bmp"
  On Error GoTo Canx
  .ShowSave
  SaveFileName = .FileName
 End With
Canx:
End Function
'The MRU business is much easier
' if you have a fixed number
' of MRUs-here I'm using 5
Private Sub UpdateMRU()
 If Len(CurrFileName) = 0 Then Exit Sub
 'check exists
 For i = 1 To 5
  If CurrFileName = MruArr(i) Then
   Exit Sub 'could move it to top
  End If
 Next
 'move all down 1 slot
 For i = 5 To 2 Step -1
  MruArr(i) = MruArr(i - 1)
 Next
 If MruCnt < 5 Then MruCnt = MruCnt + 1
 MruArr(1) = CurrFileName 'put new at top
 FillMnuMRU
End Sub
Private Sub FillMnuMRU()
 For i = 1 To MruCnt
  If i > MRU.UBound Then Load MRU(i)
  MRU(i).Visible = True
  MRU(i).Caption = MruArr(i)
 Next
End Sub
Private Sub SaveSettings()
 With oIni
  .Key = "Count"
  .Value = MruCnt
  For i = 1 To MruCnt
   .Key = "File" & i
   .Value = MruArr(i)
  Next
 End With
End Sub
Private Sub LoadSettings()
 Dim Pth As String, Cnt As String
 With oIni
  .Key = "Count"
  Cnt = .Value
  For i = 1 To Cnt
   .Key = "File" & i
   Pth = .Value
   '  'in case the file went away
   If FileExists(Pth) Then
    MruCnt = MruCnt + 1
    MruArr(MruCnt) = Pth
   End If
  Next
 End With
 FillMnuMRU
End Sub

Private Sub SetPencil()
'for certain tools, return the
' drawing tool to pencil
 For i = 1 To TBTools.Buttons.Count
  TBTools.Buttons(i).Value = tbrUnpressed
 Next
 TBTools.Refresh
 CurrTool = TPencil
 TBTools.Buttons(CurrTool).Value = tbrPressed
End Sub
Private Sub PasteXY(ByVal X As Long, ByVal Y As Long, ByVal Clear As Boolean)
'picSel will hold a picture
'of the selected part of the grid
'this routine just puts the
'colors in picBMP at right position
'or clears it for the cut operation
 Dim mx As Long, my As Long, C As Long
 With picSel
  For my = 0 To .ScaleHeight - 1 Step 20
   For mx = 0 To .ScaleWidth - 1 Step 20
    If Clear Then
     C = picBMP.BackColor
    Else
     C = GetPixel(.hdc, mx + PixH, my + PixH)
    End If
    SetPixelV picBMP.hdc, X + mx \ Pix, Y + my \ Pix, C
   Next
  Next
 End With
 Pic2Grid
End Sub

