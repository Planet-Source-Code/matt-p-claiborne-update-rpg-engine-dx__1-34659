VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar Lifebar 
      Height          =   180
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   240
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   840
      Width           =   9600
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6840
         Top             =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   210
      ScaleHeight     =   7230
      ScaleWidth      =   9630
      TabIndex        =   3
      Top             =   810
      Width           =   9660
   End
   Begin MSComctlLib.ProgressBar ExpBar 
      Height          =   180
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Class: Warrior"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "EXP:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Level: 0"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Health:"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'directx shiit
Dim Directx As New DirectX7
Dim DDraw As DirectDraw7

Dim DDsd1 As DDSURFACEDESC2
Dim DDsd2 As DDSURFACEDESC2
Dim DDsd3 As DDSURFACEDESC2
Dim DDsd4 As DDSURFACEDESC2

Dim suOff1 As DirectDrawSurface7
Dim suPrim As DirectDrawSurface7
Dim CharSprite As DirectDrawSurface7
Dim Tiles As DirectDrawSurface7

Dim Clipper As DirectDrawClipper
Dim ColorKey As DDCOLORKEY

Dim bInit As Boolean
'game shiit
Const DELAY_TIME = 20
Dim MAP_WIDTH As Long
Dim MAP_HEIGHT As Long
Const MOVE_RATE = 0.18

Dim Char As MoB
Dim Running As Boolean
Dim CharDC As Long
Dim CharSpriteDC As Long
Dim TempDC As Long
Dim bmpProperties As BITMAP
Dim FPS As Long
Dim LastCheck As Long
Dim AnimCount As Single
Dim Map() As MapTile
Dim PicDC As Long
Dim StageDC As Long
Dim BmpOld As Long
Dim TileSet As Long
Dim Retval As Long
Dim Xt As Integer
Dim Yt As Integer
Dim Xz As Integer
Dim Yz As Integer
Dim PosX As Long
Dim PosY As Long
Dim Xr As Long
Dim Yr As Long
Dim Elapsed As Long
 

Private Sub Command1_Click()
Char.Life = Char.Life - 10
UpdateChar
End Sub

Private Sub Command2_Click()
Char.Exp = Char.Exp + 10
UpdateExp
End Sub

Private Sub Form_Load()
Init

End Sub

Private Sub ResetChar()
'Char.Name = InputBox("Enter You name", "Name?")
Label3.Caption = Char.Name
Char.x = 320
Char.y = 240
Char.OldX = 320
Char.OldY = 240
Char.Direction = 1
Char.Movement = 3

Char.Level = 0
Char.Exp = 0
Char.LevelEXP = 100

Char.MaxLife = 100
Char.Life = Char.MaxLife
Label1.Caption = "Health: " & Char.Life & "/" & Char.MaxLife

Lifebar.Max = Char.MaxLife
Lifebar.Value = Char.Life

UpdateExp
UpdateChar
UpdateLevel
End Sub
Private Sub UpdateLevel()
Label2.Caption = "Level: " & Char.Level
End Sub
Private Sub UpdateExp()
If Char.Exp >= Char.LevelEXP Then
Char.Exp = Char.LevelEXP
ExpBar.Value = Char.Exp
Label4.Caption = "Exp: " & Char.Exp & "/" & Char.LevelEXP
MsgBox "Level Up", vbCritical, "level!"
Char.Level = Char.Level + 1
UpdateLevel
Char.Exp = 0
Char.LevelEXP = Char.LevelEXP * 2
End If

Label4.Caption = "Exp: " & Char.Exp & "/" & Char.LevelEXP
ExpBar.Max = Char.LevelEXP
ExpBar.Value = Char.Exp
End Sub
Private Sub UpdateChar()
If Char.Life <= 0 Then
Lifebar.Value = 0
Label1.Caption = "Health: " & Char.Life & "/" & Char.MaxLife

MsgBox "You have been SLAYED!", vbCritical, "DEAD"
ResetChar
End If

Label1.Caption = "Health: " & Char.Life & "/" & Char.MaxLife

Lifebar.Value = Char.Life
End Sub

Private Sub LoadMap(MapPath As String)
Dim T As String
Dim TileInfo() As String


Open App.Path & "/" & MapPath For Input As 1

Input #1, T
MAP_WIDTH = T
Input #1, T
MAP_HEIGHT = T
ReDim Map(MAP_WIDTH / 32, MAP_HEIGHT / 32) As MapTile

For Yt = 0 To (MAP_HEIGHT / 32) - 1
Input #1, T
TileInfo = Split(T, "|")
    For Xt = 0 To (MAP_WIDTH / 32) - 1
        Map(Xt, Yt).x = Mid(TileInfo(Xt), 1, 1) * 32
        Map(Xt, Yt).y = Mid(TileInfo(Xt), 3, 1) * 32
        Map(Xt, Yt).Walkable = Mid(TileInfo(Xt), 5, 1)
    Next
Next


End Sub

Private Sub Init()

Set DDraw = Directx.DirectDrawCreate("")
Call DDraw.SetCooperativeLevel(Form1.hwnd, DDSCL_NORMAL)
Me.Show
DDsd1.lFlags = DDSD_CAPS
DDsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
Set suPrim = DDraw.CreateSurface(DDsd1)
Set Clipper = DDraw.CreateClipper(0)
Clipper.SetHWnd Picture1.hwnd
suPrim.SetClipper Clipper

ColorKey.high = 0
ColorKey.low = 0

suPrim.SetColorKey DDCKEY_SRCBLT, ColorKey


DDsd2.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
DDsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
DDsd2.lWidth = 672
DDsd2.lHeight = 512
Set suOff1 = DDraw.CreateSurface(DDsd2)


DDsd3.lFlags = DDSD_CAPS
DDsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set CharSprite = DDraw.CreateSurfaceFromFile(App.Path & "/cecil.bmp", DDsd3)
CharSprite.SetColorKey DDCKEY_SRCBLT, ColorKey

DDsd3.lFlags = DDSD_CAPS
DDsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set Tiles = DDraw.CreateSurfaceFromFile(App.Path & "/tiles.bmp", DDsd4)

bInit = True

ResetChar
LoadMap "map1.txt"

If bInit = True Then Running = True

Do While Running = True


Elapsed = GetTickCount() - LastCheck
LastCheck = GetTickCount

MoveChar
DrawStage

'copy buffer to screen

FPS = FPS + 1
Nada:



DoEvents
Loop


End Sub



Private Sub Form_Resize()
'If Form1.WindowState = 0 Then
'Picture1.Width = Form1.ScaleWidth - 33
'Picture1.Height = Form1.ScaleHeight - 68
'Picture2.Width = Form1.ScaleWidth - 29
'Picture2.Height = Form1.ScaleHeight - 64
'End If
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Timer1_Timer()
Form1.Caption = FPS & " - FPS    " & FPS * 20 * 15 & " - Sprites Per Second"
FPS = 0

End Sub

Private Sub MoveChar()
Dim Xd As Long
Dim Yd As Long

'Time based modeling below, hey! that was easy
Char.Moved = False
Char.Movement = Elapsed * MOVE_RATE

If GetAsyncKeyState(vbKeyA) Then
    Char.Movement = 2.5 * Char.Movement
End If

If GetAsyncKeyState(vbKeyS) Then
    Char.Movement = 0.3 * Char.Movement
End If

If GetAsyncKeyState(GM_ESCAPE) Then
    Running = False
    Unload Me
    End
End If

Char.OldY = Char.y
Char.OldX = Char.x

If GetAsyncKeyState(GM_LEFT) Then
    Char.x = Char.x - Char.Movement
    Char.Direction = 3
    Char.Moved = True
    GoTo Moved
End If

If GetAsyncKeyState(GM_UP) Then
    Char.y = Char.y - Char.Movement
    Char.Direction = 0
    Char.Moved = True
    GoTo Moved
End If

If GetAsyncKeyState(GM_RIGHT) Then
    Char.x = Char.x + Char.Movement
    Char.Direction = 1
    Char.Moved = True
    GoTo Moved
End If

If GetAsyncKeyState(GM_DOWN) Then
    Char.y = Char.y + Char.Movement
    Char.Direction = 2
    Char.Moved = True
    GoTo Moved
End If


Moved:

If Char.x < 0 Then
    Char.x = 0
ElseIf Char.x + 32 > MAP_WIDTH Then
    Char.x = MAP_WIDTH - 32
End If

If Char.y < 0 Then
    Char.y = 0
ElseIf Char.y + 32 > MAP_HEIGHT Then
    Char.y = MAP_HEIGHT - 32
End If


Xd = Int(Char.x / 32)
Yd = Int(Char.y / 32)

If Char.x Mod 32 = 0 Then
    Xz = 0
Else
    Xz = 1
End If

If Char.y Mod 32 = 0 Then
    Yz = 0
Else
    Yz = 1
End If

For Xr = 0 To Xz
For Yr = 0 To Yz

If Map(Xd + Xr, Yd + Yr).Walkable = 1 Then

    If Char.Direction = 1 Then
        Char.x = (Xd + Xr - 1) * 32
        Char.y = Char.OldY
    End If
    
    If Char.Direction = 3 Then
        Char.x = (Xd + Xr + 1) * 32
        Char.y = Char.OldY
    End If
    
    If Char.Direction = 2 Then
        Char.x = Char.OldX
        Char.y = (Yd + Yr - 1) * 32
    End If
    
    If Char.Direction = 0 Then
        Char.x = Char.OldX
        Char.y = (Yd + Yr + 1) * 32
    End If

End If

Next
Next


If Char.Moved = True Then
AnimCount = AnimCount + Char.Movement / 20
If AnimCount > 2 Then AnimCount = 0
Else
AnimCount = 0
End If



End Sub

Private Sub DrawStage()
Dim ddrval As Long
Dim rTemp As RECT
Dim r1 As RECT
Dim r2 As RECT

Dim B1 As RECT
Dim B2 As RECT

Dim Rdum As RECT
Rdum.Left = 0
Rdum.Top = 0
Rdum.Right = 640
Rdum.Bottom = 480

Call Directx.GetWindowRect(Picture1.hwnd, rTemp)

Xr = (Char.x - 320)
Yr = (Char.y - 240)

If Xr < 0 Then Xr = 0
If Yr < 0 Then Yr = 0

If Xr + 640 > MAP_WIDTH Then Xr = MAP_WIDTH - 640
If Yr + 480 > MAP_HEIGHT Then Yr = MAP_HEIGHT - 480

Xz = (Xr) Mod 32
Yz = (Yr) Mod 32

For Xt = 0 To 640 Step 32
For Yt = 0 To 480 Step 32

B2.Left = Xt - (Xz)
B2.Top = Yt - (Yz)
B2.Right = B2.Left + 32
B2.Bottom = B2.Top + 32

B1.Left = Map(Int((Xr + Xt) / 32), Int((Yr + Yt) / 32)).x
B1.Right = B1.Left + 32
B1.Top = Map(Int((Xr + Xt) / 32), Int((Yr + Yt) / 32)).y
B1.Bottom = B1.Top + 32

If B2.Left < 0 Then
    B1.Left = B1.Left + -B2.Left
    B2.Left = 0
End If

If B2.Top < 0 Then
    B1.Top = B1.Top + -B2.Top
    B2.Top = 0
End If



'copy a 32x32map tile to the stage
suOff1.Blt B2, Tiles, B1, DDBLTFAST_DONOTWAIT

Next
Next

If Char.x < 320 Then
        PosX = -(320 - Char.x)
    ElseIf Char.x > MAP_WIDTH - 320 Then
        PosX = (Char.x - MAP_WIDTH) + 320
    Else
        PosX = 0
End If


If Char.y < 240 Then
        PosY = -(240 - Char.y)
    ElseIf Char.y > MAP_HEIGHT - 240 Then
        PosY = (Char.y - MAP_HEIGHT) + 240
    Else
        PosY = 0
End If







r2.Top = Char.Direction * 32
r2.Left = CInt(AnimCount) * 32
r2.Right = r2.Left + 32
r2.Bottom = r2.Top + 32

r1.Left = PosX + 320
r1.Top = PosY + 240
r1.Right = r1.Left + 32
r1.Bottom = r1.Top + 32

'copy charactor to stage
suOff1.Blt r1, CharSprite, r2, DDBLTFAST_DONOTWAIT Or DDBLT_KEYSRC

'copy stage to screen
suPrim.Blt rTemp, suOff1, Rdum, DDBLT_WAIT


End Sub


