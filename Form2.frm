VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Pallete"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMouse 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox Light 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   6360
      MousePointer    =   2  'Cross
      ScaleHeight     =   2775
      ScaleWidth      =   255
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chosen Color"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2775
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   6000
      MousePointer    =   2  'Cross
      ScaleHeight     =   2775
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      Picture         =   "Form2.frx":110C
      ScaleHeight     =   255
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   3000
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   3240
      MousePointer    =   2  'Cross
      Picture         =   "Form2.frx":3926
      ScaleHeight     =   2790
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   120
      Width           =   2625
   End
   Begin VB.Frame Frame2 
      Caption         =   "RGB"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   3015
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "255"
         Top             =   150
         Width           =   495
      End
      Begin VB.TextBox txtG 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "255"
         Top             =   150
         Width           =   495
      End
      Begin VB.TextBox txtB 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "255"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Decimal"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   3015
      Begin VB.TextBox txtLong 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "16777215"
         Top             =   150
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "HEX"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   3015
      Begin VB.TextBox txtHex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "#FFFFFF"
         Top             =   150
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTransfer As Double
Public HexTransfer As String
Public RTransfer As Double
Public BTransfer As Double
Public GTransfer As Double

Dim RGBValues(4) As Long
Dim NowColor As Double
Dim XTOP As Double, XLEFT As Double

Dim IsDown As Boolean

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Function HexColor()
    HexRed = Hex$(txtR.Text)
    If Len(HexRed) = 1 Then HexRed = "0" & HexRed
    HexGreen = Hex$(txtG.Text)
    If Len(HexGreen) = 1 Then HexGreen = "0" & HexGreen
    HexBlue = Hex$(txtB.Text)
    If Len(HexBlue) = 1 Then HexBlue = "0" & HexBlue
    txtHex.Text = "#" & HexRed & HexGreen & HexBlue
End Function

Public Function RGBPicker()
    RGBValues(3) = CLng(NowColor)
    RGBValues(0) = RGBValues(3) And 255
    RGBValues(1) = (RGBValues(3) And 65280) \ 256&
    RGBValues(2) = (RGBValues(3) And 16711680) \ 65535

    txtR.Text = RGBValues(0)
    txtG.Text = RGBValues(1)
    txtB.Text = RGBValues(2)

    Light.DrawWidth = 2
    P = 0
    For I = 1 To 254
    P = P + 13
    Light.ForeColor = RGB(RGBValues(0), RGBValues(1), I)
    Light.Line (0, P)-(245, P)
    Next I
End Function

Private Sub Command1_Click()
    Form1.Text9.Text = Picture4.BackColor
    Form1.Picture1.BackColor = Picture4.BackColor
    Form1.HScroll1.Value = txtR.Text
    Form1.HScroll2.Value = txtG.Text
    Form1.HScroll3.Value = txtB.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    IsDown = False
End Sub

Private Sub light_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim TEMP_COLOR As Long
    IsDown = True
    TEMP_COLOR = GetPixel(Light.hdc, X / 15, Y / 15)
    Picture4.BackColor = TEMP_COLOR
End Sub

Private Sub light_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If IsDown = False Then
        Picture3.BackColor = GetPixel(Light.hdc, X / 15, Y / 15)
    ElseIf IsDown = True Then
        Dim TEMP_COLOR2 As Long
        TEMP_COLOR2 = GetPixel(Light.hdc, X / 15, Y / 15)
        Picture4.BackColor = TEMP_COLOR2
        Picture3.BackColor = TEMP_COLOR2
    End If
End Sub

Private Sub Light_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    IsDown = False
    NowColor = Picture4.BackColor
    txtLong.Text = NowColor
    RGBPicker
    HexColor
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim TEMP_COLOR As Long
    IsDown = True
    TEMP_COLOR = GetPixel(Picture1.hdc, X / 15, Y / 15)
    Picture4.BackColor = TEMP_COLOR
    NowColor = TEMP_COLOR
    txtLong.Text = NowColor
    RGBPicker
    HexColor
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If IsDown = False Then
        Picture3.BackColor = GetPixel(Picture1.hdc, X / 15, Y / 15)
    ElseIf IsDown = True Then
        Dim TEMP_COLOR2 As Long
        TEMP_COLOR2 = GetPixel(Picture1.hdc, X / 15, Y / 15)
        Picture4.BackColor = TEMP_COLOR2
        Picture3.BackColor = TEMP_COLOR2
        NowColor = TEMP_COLOR2
        txtLong.Text = NowColor
        RGBPicker
        HexColor
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    IsDown = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim TEMP_COLOR As Long
    IsDown = True
    TEMP_COLOR = GetPixel(Picture2.hdc, X / 15, Y / 15)
    Picture4.BackColor = TEMP_COLOR
    NowColor = TEMP_COLOR
    txtLong.Text = NowColor
    RGBPicker
    HexColor
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If IsDown = False Then
        Picture3.BackColor = GetPixel(Picture2.hdc, X / 15, Y / 15)
    ElseIf IsDown = True Then
        Dim TEMP_COLOR2 As Long
        TEMP_COLOR2 = GetPixel(Picture2.hdc, X / 15, Y / 15)
        Picture4.BackColor = TEMP_COLOR2
        Picture3.BackColor = TEMP_COLOR2
        NowColor = TEMP_COLOR
        txtLong.Text = NowColor
        RGBPicker
        HexColor
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    IsDown = False
End Sub

Private Sub Picture3_Click()
    Picture4.BackColor = Picture3.BackColor
    NowColor = Picture3.BackColor
    txtLong.Text = NowColor
    RGBPicker
    HexColor
End Sub
