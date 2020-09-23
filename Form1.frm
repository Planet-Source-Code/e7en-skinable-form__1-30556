VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Image RS 
      Height          =   90
      Left            =   0
      Top             =   2400
      Width           =   210
   End
   Begin VB.Image UR 
      Height          =   210
      Left            =   240
      Top             =   240
      Width           =   90
   End
   Begin VB.Image UM 
      Height          =   135
      Left            =   480
      Top             =   600
      Width           =   210
   End
   Begin VB.Image BR 
      Height          =   135
      Left            =   240
      Top             =   1320
      Width           =   210
   End
   Begin VB.Image BM 
      Height          =   120
      Left            =   120
      Top             =   2520
      Width           =   105
   End
   Begin VB.Image BL 
      Height          =   90
      Left            =   240
      Top             =   3960
      Width           =   105
   End
   Begin VB.Image LS 
      Height          =   225
      Left            =   120
      Top             =   1680
      Width           =   90
   End
   Begin VB.Image UL 
      Height          =   120
      Left            =   120
      Top             =   840
      Width           =   195
   End
   Begin VB.Image bak 
      Height          =   135
      Left            =   240
      Top             =   3600
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
UnStretchIt
LoadGFX (Dir1.Path)
StretchIt
GFX
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path & "\skinz\Default"
UnStretchIt
LoadGFX (Dir1.Path)
StretchIt
GFX
End Sub

Sub LoadGFX(FileDir As String)

bak.Picture = LoadPicture(FileDir & "\Back_Pic.bmp")

UL.Picture = LoadPicture(FileDir & "\UpLeft_pic.bmp")
UR.Picture = LoadPicture(FileDir & "\UpRight_pic.bmp")
UM.Picture = LoadPicture(FileDir & "\UpMiddle_pic.bmp")

LS.Picture = LoadPicture(FileDir & "\LeftSide_pic.bmp")
RS.Picture = LoadPicture(FileDir & "\RightSide_pic.bmp")

BL.Picture = LoadPicture(FileDir & "\BottomLeft_pic.bmp")
BM.Picture = LoadPicture(FileDir & "\BottomMiddle_pic.bmp")
BR.Picture = LoadPicture(FileDir & "\BottomRight_pic.bmp")

End Sub

Sub UnStretchIt()
bak.Stretch = False

UL.Stretch = False
UR.Stretch = False
UM.Stretch = False

LS.Stretch = False
RS.Stretch = False

BL.Stretch = False
BM.Stretch = False
BR.Stretch = False
End Sub

Sub StretchIt()
bak.Stretch = True

UL.Stretch = True
UR.Stretch = True
UM.Stretch = True

LS.Stretch = True
RS.Stretch = True

BL.Stretch = True
BM.Stretch = True
BR.Stretch = True
End Sub

Sub GFX()
bak.Left = 0
bak.Top = 0
bak.Width = Me.Width
bak.Height = Me.Height

UL.Top = 0
UL.Left = 0

BL.Top = Me.Height - BL.Height
BL.Left = 0

LS.Top = UL.Height
LS.Height = Me.Height - UL.Height - BL.Height
LS.Left = 0

UM.Width = Me.Width - UL.Width - UR.Width
UM.Left = UL.Width
UM.Top = 0

UR.Top = 0
UR.Left = Me.Width - UR.Width

BR.Left = Me.Width - BR.Width
BR.Top = Me.Height - BR.Height

RS.Left = Me.Width - RS.Width
RS.Height = Me.Height - UR.Height - BR.Height
RS.Top = UR.Height

BM.Top = Me.Height - BM.Height
BM.Left = BL.Width
BM.Width = Me.Width - BL.Width - BR.Width

End Sub

Private Sub Form_Resize()
GFX
End Sub
