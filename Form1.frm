VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demian Net Color Picker v2.0"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8490
   Begin VB.Frame Frame1 
      Caption         =   "Demian Net Color Picker"
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame9 
         Caption         =   "URL"
         Height          =   855
         Left            =   4920
         TabIndex        =   31
         Top             =   3240
         Width           =   3375
         Begin VB.CommandButton Command5 
            Caption         =   "Copy"
            Height          =   285
            Left            =   2640
            TabIndex        =   33
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "http://www.DemianNet.com/"
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Decimal"
         Height          =   855
         Left            =   1440
         TabIndex        =   28
         Top             =   3240
         Width           =   3375
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Copy"
            Height          =   285
            Left            =   2640
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Misc"
         Height          =   1695
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
         Begin VB.CommandButton Command8 
            Caption         =   "Pallete"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "White"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Black"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Hex/HTML"
         Height          =   855
         Left            =   4920
         TabIndex        =   18
         Top             =   2400
         Width           =   3375
         Begin VB.CommandButton Command2 
            Caption         =   "Copy"
            Height          =   285
            Left            =   2640
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "#000000"
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "RGB"
         Height          =   855
         Left            =   1440
         TabIndex        =   17
         Top             =   2400
         Width           =   3375
         Begin VB.CommandButton Command1 
            Caption         =   "Copy"
            Height          =   285
            Left            =   2640
            TabIndex        =   21
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0,0,0"
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Color"
         Height          =   2055
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            Height          =   1695
            Left            =   120
            ScaleHeight     =   1635
            ScaleWidth      =   3075
            TabIndex        =   16
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Red, Green, Blue"
         Height          =   2055
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   3375
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   27
            Top             =   1680
            Width           =   255
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   26
            Top             =   1080
            Width           =   255
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   25
            Top             =   480
            Width           =   255
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   14
            Top             =   1440
            Width           =   3135
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   360
            ScaleHeight     =   195
            ScaleWidth      =   2835
            TabIndex        =   13
            Top             =   1680
            Width           =   2895
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   12
            Top             =   840
            Width           =   3135
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   360
            ScaleHeight     =   195
            ScaleWidth      =   2835
            TabIndex        =   11
            Top             =   1080
            Width           =   2895
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   10
            Top             =   240
            Width           =   3135
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   360
            ScaleHeight     =   195
            ScaleWidth      =   2835
            TabIndex        =   9
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Dec    Hex   "
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "00"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "00"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "00"
            Top             =   240
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function LeftPad(Value, Size As Long, Optional PadCharacter As String = " ") As String
    LeftPad = "" & Value
    While Len(LeftPad) < Size
        LeftPad = PadCharacter & LeftPad
    Wend
End Function

Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText Text7.Text
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text8.Text
End Sub

Private Sub Command3_Click()
    Randomize
    Random = Int(Rnd * 255)
    HScroll1.Value = Random
    Randomize
    Random = Int(Rnd * 255)
    HScroll2.Value = Random
    Randomize
    Random = Int(Rnd * 255)
    HScroll3.Value = Random
End Sub

Private Sub Command4_Click()
    Clipboard.Clear
    Clipboard.SetText Text9.Text
End Sub

Private Sub Command5_Click()
    Clipboard.Clear
    Clipboard.SetText Text10.Text
End Sub

Private Sub Command6_Click()
    HScroll1.Value = 0
    HScroll2.Value = 0
    HScroll3.Value = 0
End Sub

Private Sub Command7_Click()
    HScroll1.Value = 255
    HScroll2.Value = 255
    HScroll3.Value = 255
End Sub

Private Sub Command8_Click()
    Form2.Show 1
End Sub

Private Sub Form_Load()
    On Error Resume Next
    HScroll1.Value = GetSetting("Demian Net Color Picker", "Colors", "RED")
    HScroll2.Value = GetSetting("Demian Net Color Picker", "Colors", "GREEN")
    HScroll3.Value = GetSetting("Demian Net Color Picker", "Colors", "BLUE")
    Form1.Top = GetSetting("Demian Net Color Picker", "Pos", "TOP")
    Form1.Left = GetSetting("Demian Net Color Picker", "Pos", "LEFT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "Demian Net Color Picker", "Colors", "RED", HScroll1.Value
    SaveSetting "Demian Net Color Picker", "Colors", "GREEN", HScroll2.Value
    SaveSetting "Demian Net Color Picker", "Colors", "BLUE", HScroll3.Value
    SaveSetting "Demian Net Color Picker", "Pos", "TOP", Form1.Top
    SaveSetting "Demian Net Color Picker", "Pos", "LEFT", Form1.Left
End Sub

Private Sub HScroll1_Change()
    Picture2.BackColor = RGB(HScroll1.Value, 0, 0)
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text2.Text = HScroll1.Value
    Text1.Text = LeftPad(Hex(HScroll1.Value), 2, 0)
    Text7.Text = HScroll1.Value & "," & HScroll2.Value & "," & HScroll3.Value
    Text8.Text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
    Text9.Text = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    Picture3.BackColor = RGB(0, HScroll2.Value, 0)
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text3.Text = HScroll2.Value
    Text4.Text = LeftPad(Hex(HScroll2.Value), 2, 0)
    Text7.Text = HScroll1.Value & "," & HScroll2.Value & "," & HScroll3.Value
    Text8.Text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
    Text9.Text = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Scroll()
    HScroll2_Change
End Sub

Private Sub HScroll3_Change()
    Picture4.BackColor = RGB(0, 0, HScroll3.Value)
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text5.Text = HScroll3.Value
    Text6.Text = LeftPad(Hex(HScroll3.Value), 2, 0)
    Text7.Text = HScroll1.Value & "," & HScroll2.Value & "," & HScroll3.Value
    Text8.Text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
    Text9.Text = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Scroll()
    HScroll3_Change
End Sub

