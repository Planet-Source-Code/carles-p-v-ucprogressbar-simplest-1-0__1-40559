VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Test 
   Caption         =   "Owner drawn progress bar"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "!"
      Height          =   270
      Left            =   1440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2265
      Width           =   285
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Test.frx":0000
      Left            =   1440
      List            =   "Test.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1815
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Smooth"
      Height          =   255
      Left            =   270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3855
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1245
      MaxLength       =   5
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "5000"
      Top             =   255
      Width           =   630
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test"
      Height          =   375
      Left            =   6510
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3375
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   6525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   285
      TabIndex        =   11
      Top             =   3375
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Progress.ucProgress ucProgress1 
      Height          =   375
      Left            =   285
      Top             =   1365
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
   End
   Begin VB.Label Label5 
      Caption         =   "Random colors"
      Height          =   210
      Left            =   285
      TabIndex        =   6
      Top             =   2295
      Width           =   1350
   End
   Begin VB.Label Label4 
      Caption         =   "Border style:"
      Height          =   210
      Left            =   285
      TabIndex        =   4
      Top             =   1875
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Max value:"
      Height          =   225
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "VB progress bar"
      Height          =   225
      Left            =   285
      TabIndex        =   8
      Top             =   3120
      Width           =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Owner drawn progress bar"
      Height          =   225
      Left            =   285
      TabIndex        =   2
      Top             =   1095
      Width           =   2280
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    Combo1.ListIndex = 2
End Sub



Private Sub Combo1_Click()
    ucProgress1.BorderStyle = Combo1.ListIndex
End Sub

Private Sub Command3_Click()

  Dim R As Byte, G As Byte, B As Byte
    
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
    
    ucProgress1.BackColor = RGB(R, G, B)
    ucProgress1.ForeColor = RGB(Not R, Not G, Not B)
End Sub

Private Sub Check1_Click()
    ProgressBar1.Scrolling = Check1
End Sub



Private Sub Command1_Click()

  Dim i As Long
  
    ucProgress1.Max = Text1
    For i = 0 To Text1
        ucProgress1 = i
    Next i
    For i = Text1 To 0 Step -1
        ucProgress1 = i
    Next i
End Sub

Private Sub Command2_Click()

  Dim i As Long
  
    ProgressBar1.Max = Text1
    For i = 0 To Text1
        ProgressBar1 = i
    Next i
    For i = Text1 To 0 Step -1
        ProgressBar1 = i
    Next i
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text1_Change()
    If (Val(Text1) < 1) Then
        Text1 = 1
        Text1.SelStart = 1
    End If
End Sub
