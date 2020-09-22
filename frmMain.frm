VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Fecal Emmiting Dragon"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrShitFall 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1560
      Top             =   4320
   End
   Begin VB.PictureBox picApple 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   7080
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   -1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Timer tmrMove 
      Interval        =   25
      Left            =   840
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   360
      Top             =   4320
   End
   Begin PicClip.PictureClip pcpDragon 
      Index           =   0
      Left            =   1920
      Top             =   960
      _ExtentX        =   8334
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   2
      Cols            =   3
      Picture         =   "frmMain.frx":1000
   End
   Begin PicClip.PictureClip pcpDragon 
      Index           =   1
      Left            =   1920
      Top             =   2520
      _ExtentX        =   8334
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   2
      Cols            =   3
      Picture         =   "frmMain.frx":19526
   End
   Begin VB.PictureBox PicDragon 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   1575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox picMan 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   4560
      ScaleHeight     =   3135
      ScaleWidth      =   1935
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1935
   End
   Begin PicClip.PictureClip pcpMan 
      Left            =   -840
      Top             =   6360
      _ExtentX        =   14049
      _ExtentY        =   10583
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      Picture         =   "frmMain.frx":31A4C
   End
   Begin VB.Shape shpShit 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   6960
      Shape           =   2  'Oval
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblAmmo 
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Ammo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape shpGround 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   525
      Left            =   0
      Top             =   5760
      Width           =   10455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurrCell As Byte
Private SHPix As Integer
Private SWPix As Integer
Private CellsWide As Integer
Private CellsHigh As Integer
Private MoveDir As Byte
Private MoveLen As Integer
Private Ammo As Integer
Private FacingLeft As Boolean
Private Score As Integer
Private ShitFalling As Boolean

Private Sub Form_Load()

Randomize Timer
SHPix = Screen.Height / Screen.TwipsPerPixelY
SWPix = Screen.Width / Screen.TwipsPerPixelX
CellsWide = Int(SWPix / 32)
CellsHigh = Int(SHPix / 32)
frmMain.Top = 0
frmMain.Left = 0
frmMain.Width = Screen.Width
frmMain.Height = Screen.Height


shpGround.Top = SHPix - 50
shpGround.Width = SWPix
picMan.Height = pcpMan.CellHeight
picMan.Width = pcpMan.CellWidth
picMan.Top = shpGround.Top - picMan.Height

Dim x
For x = 1 To 5
    Load picApple(x)
    picApple(x).Top = Int(Rnd * (SHPix - 210))
    picApple(x).Left = Int(Rnd * SWPix)
    picApple(x).Visible = True
Next x

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case 32
        If ShitFalling = True Then Exit Sub
        If Ammo <= 0 Then Exit Sub
        Ammo = Ammo - 1
        lblAmmo.Caption = Ammo
        ShitFalling = True
        tmrShitFall.Enabled = True
        shpShit.Top = PicDragon.Top + PicDragon.Height
        If FacingLeft = True Then shpShit.Left = PicDragon.Left + 100
        If FacingLeft = False Then shpShit.Left = PicDragon.Left
        shpShit.Visible = True
    
    Case 37
        If MoveDir <> 0 Then Exit Sub
        MoveDir = 1
    
    Case 38
        If MoveDir <> 0 Then Exit Sub
        MoveDir = 2
            
    Case 39
        If MoveDir <> 0 Then Exit Sub
        MoveDir = 3
    
    Case 40
        If MoveDir <> 0 Then Exit Sub
        MoveDir = 4

End Select

End Sub

Private Sub Timer1_Timer()

CurrCell = CurrCell + 1
If CurrCell > 5 Then CurrCell = 0
If FacingLeft = True Then PicDragon.Picture = pcpDragon(0).GraphicCell(CurrCell)
If FacingLeft = False Then PicDragon.Picture = pcpDragon(1).GraphicCell(CurrCell)
picMan.Picture = pcpMan.GraphicCell(CurrCell)
picMan.Left = picMan.Left + 12
If picMan.Left >= SWPix Then
    picMan.Left = -4
End If

End Sub

Private Sub tmrMove_Timer()

If MoveDir = 0 Then Exit Sub
If MoveLen > 2 Then
    MoveDir = 0
    MoveLen = 0
End If

Select Case MoveDir

    Case 1
        FacingLeft = True
        MoveLen = MoveLen + 1
        If PicDragon.Left <= 32 Then
            PicDragon.Left = SWPix - 15
        Else
            PicDragon.Left = PicDragon.Left - 16
        End If
        
    Case 2
        MoveLen = MoveLen + 1
        If PicDragon.Top <= 16 Then
            PicDragon.Top = 0
        Else
            PicDragon.Top = PicDragon.Top - 16
        End If

    Case 3
        FacingLeft = False
        MoveLen = MoveLen + 1
        If PicDragon.Left >= SWPix - 15 Then
            PicDragon.Left = -16
        Else
            PicDragon.Left = PicDragon.Left + 16
        End If
        
    Case 4
        MoveLen = MoveLen + 1
        If PicDragon.Top >= SHPix - 200 Then
            PicDragon.Top = SHPix - 190
        Else
            PicDragon.Top = PicDragon.Top + 16
        End If
        
End Select

Dim x

For x = 1 To 5
    If picApple(x).Left + picApple(x).Width >= PicDragon.Left And picApple(x).Left <= PicDragon.Left + PicDragon.Width And picApple(x).Top + picApple(x).Height >= PicDragon.Top And picApple(x).Top <= PicDragon.Top + PicDragon.Height Then
        picApple(x).Visible = False
        picApple(x).Top = Int(Rnd * (SHPix - 210))
        picApple(x).Left = Int(Rnd * SWPix)
        Ammo = Ammo + 1
        lblAmmo.Caption = Ammo
        picApple(x).Visible = True
    End If
Next x


End Sub

Private Sub tmrShitFall_Timer()

shpShit.Top = shpShit.Top + 20

If shpShit.Top > shpGround.Top - picMan.Height Then
    If shpShit.Left - shpShit.Width >= picMan.Left And shpShit.Left <= picMan.Left + picMan.Width Then
        Score = Score + 1
        lblScore.Caption = Score
        ShitFalling = False
        shpShit.Visible = False
        Exit Sub
    End If
End If

If shpShit.Top + shpShit.Height >= shpGround.Top Then
    ShitFalling = False
    shpShit.Visible = False
End If

End Sub
