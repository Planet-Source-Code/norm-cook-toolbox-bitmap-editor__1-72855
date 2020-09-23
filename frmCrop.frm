VERSION 5.00
Begin VB.Form frmCrop 
   Caption         =   "Crop Picture to 16x15"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "frmCrop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   6120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdOK 
         Caption         =   "Use this Selection"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCanx 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.PictureBox picCrop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1920
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   2
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   0
      MouseIcon       =   "frmCrop.frx":014A
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   0
      Top             =   0
      Width           =   3030
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -15
         Picture         =   "frmCrop.frx":0A14
         Top             =   -15
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCrop.frx":0D1E
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   5055
   End
End
Attribute VB_Name = "frmCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SX As Long, SY As Long
Private Sub Form_Paint()
 BitBlt picCrop.hdc, 0, 0, 16, 15, picSrc.hdc, Image1.Left + 1, Image1.Top + 1, vbSrcCopy
 picCrop.Refresh
End Sub

Private Sub Form_Resize()
'Frame1.Move 0, ScaleHeight - Frame1.Height
End Sub
Private Sub cmdCanx_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 Set frmMain.picBMP.Picture = picCrop.Image
 Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim nl As Long, nt As Long
 nl = Image1.Left: nt = Image1.Top
 Select Case KeyCode
  Case 40
   nt = nt + 1
  Case 38
   nt = nt - 1
  Case 39
   nl = nl + 1
  Case 37
   nl = nl - 1
 End Select
 If nl < -1 Then nl = -1
 If nl > picSrc.ScaleWidth - 17 Then nl = picSrc.ScaleWidth - 17
 If nt < -1 Then nt = -1
 If nt > picSrc.ScaleHeight - 16 Then nt = picSrc.ScaleHeight - 16
 Image1.Move nl, nt
 BitBlt picCrop.hdc, 0, 0, 16, 15, picSrc.hdc, Image1.Left + 1, Image1.Top + 1, vbSrcCopy
 picCrop.Refresh
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SX = X \ 15: SY = Y \ 15
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim nl As Long, nt As Long
 If Button Then
  nl = Image1.Left + X \ 15 - SX
  nt = Image1.Top + Y \ 15 - SY
  If nl < -1 Then nl = -1
  If nl > picSrc.ScaleWidth - 17 Then nl = picSrc.ScaleWidth - 17
  If nt < -1 Then nt = -1
  If nt > picSrc.ScaleHeight - 16 Then nt = picSrc.ScaleHeight - 16
  Image1.Move nl, nt
  BitBlt picCrop.hdc, 0, 0, 16, 15, picSrc.hdc, Image1.Left + 1, Image1.Top + 1, vbSrcCopy
  picCrop.Refresh
 End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 BitBlt picCrop.hdc, 0, 0, 16, 15, picSrc.hdc, Image1.Left + 1, Image1.Top + 1, vbSrcCopy
 picCrop.Refresh
End Sub

