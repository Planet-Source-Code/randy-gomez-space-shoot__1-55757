VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bottom Ship - Arrow keys, CTRL to fire   Top Ship - ESDF, A to fire"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGameArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2190
      Left            =   1200
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   0
      Top             =   510
      Width           =   2565
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Type StarObject
    x As Single
    y As Single
    Move As Single
    Color As Long
    Diam As Single
End Type

Private Type AmmoObject
    picno As Integer
    x As Integer
    y As Integer
    Move As Integer
    ImgNo As Integer
    ImgCount As Integer
    TickCounter As Integer
    xSpan As Single
    Power As Integer
    Fired As Boolean
    SideY As Single
    SideXL As Single
    SideXR As Single
    SubCounter As Integer
    FireTime As Integer
End Type

Private Type TrailObject
    x As Single
    y As Single
    XMov As Single
    YMov As Single
    LifeTime As Integer
    picno As Integer
End Type

Private Type ShipObject
    picHDC As Long
    maskHDC As Long
    Left As Single
    Top As Single
    Width As Integer
    Height As Integer
    MaxY As Single
    MinY As Single
    HitWidth As Integer
    Hit As Integer
    BankCount As Integer
    ImgX As Integer
    Firing As Boolean
    FireTicker As Integer
    RocketsOn As Boolean
    PowerUp As Integer
    Bombing As Boolean
    SpeedLR As Integer
    SpeedUD As Integer
    shipAmmo(15) As AmmoObject
    shipTrail(25) As TrailObject
    shotTrail(6) As TrailObject
End Type

Private Type EnemyObject
    picHDC As Long
    maskHDC As Long
    Left As Single
    Top As Single
    Width As Integer
    Height As Integer
    Hit As Integer
    ImgX As Integer
    Firing As Boolean
    FireTicker As Integer
    RocketsOn As Boolean
    SpeedLR As Integer
    shipAmmo(15) As AmmoObject
    shipTrail(25) As TrailObject
    Counter As Integer
    OnScreen As Boolean
End Type

Dim RValue As Integer, GValue As Integer, BVAlue As Integer

Dim TextRec As RECT
Dim BRec(20) As RECT
Dim CtrLt As Integer
Dim LtTimer As Long

Dim SlowStar(100) As StarObject
Dim FasterStar(20) As StarObject
Dim btmship As ShipObject
Dim topship As ShipObject
Dim Enemy1 As EnemyObject
Dim Enemy1Ammo(20) As AmmoObject
Dim RocketsOn As Boolean
Dim RocketCount As Integer, Trail(25) As TrailObject
Dim BackMove As Single
Dim btmShotTrail(6) As TrailObject
Dim blnGameOn As Boolean


Private Function IsKeyDown(AsciiKeyCode As Byte) As Boolean
    
    If GetKeyState(AsciiKeyCode) < -125 Then IsKeyDown = True
    
End Function


Private Sub Form_Activate()

    blnGameOn = True
    GameLoop

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        blnGameOn = False
        Unload Me
        Unload frmSprites
    ElseIf KeyCode = vbKeyZ Then
        blnGameOn = True
        GameLoop
    ElseIf KeyCode = vbKeyControl Then
        btmship.Firing = True
    ElseIf KeyCode = vbKeyL Then
        btmship.PowerUp = 1
    ElseIf KeyCode = vbKeyK Then
        btmship.PowerUp = 0
    ElseIf KeyCode = vbKeyA Then
        topship.Firing = True
    ElseIf KeyCode = vbKeyB Then
        topship.PowerUp = 1
    ElseIf KeyCode = vbKeyN Then
        topship.PowerUp = 0
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyControl Then
        btmship.Firing = False
    ElseIf KeyCode = vbKeyA Then
        topship.Firing = False
    ElseIf KeyCode = vbKeyUp Then
        btmship.RocketsOn = False
    ElseIf KeyCode = vbKeyD Then
        topship.RocketsOn = False
    End If

End Sub

Private Sub Form_Load()
Dim i As Integer

'    Load frmBackgrd
    Load frmSprites
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = 650 * Screen.TwipsPerPixelX
    Me.Height = 700 * Screen.TwipsPerPixelY
    Me.picGameArea.Left = 0
    Me.picGameArea.Top = 0
    Me.picGameArea.Width = 650
    Me.picGameArea.Height = 700
    picGameArea.BackColor = RGB(0, 0, 0)

    Randomize
    
    'set up position, speed and colour (white) of faster moving stars
    For i = 1 To 20
        FasterStar(i).x = Rnd * picGameArea.ScaleWidth
        FasterStar(i).y = Rnd * picGameArea.ScaleHeight + 30
        FasterStar(i).Move = Rnd * 2 + 1
        RValue = Int(FasterStar(i).Move * 80)
        GValue = Int(FasterStar(i).Move * 80)
        BVAlue = Int(FasterStar(i).Move * 80)
        FasterStar(i).Color = RGB(RValue, GValue, BVAlue)
        FasterStar(i).Diam = FasterStar(i).Move + 1
    Next i
        
    'set up position, speed and color(shades of grey) of slow stars
    For i = 1 To 100
        SlowStar(i).x = Rnd * picGameArea.ScaleWidth
        SlowStar(i).y = Rnd * picGameArea.ScaleHeight + 30
        SlowStar(i).Move = Rnd + 1
        SlowStar(i).Move = 1
        RValue = Int(Rnd * 200) + 100
        GValue = Int(Rnd * 200) + 100
        BVAlue = Int(Rnd * 200) + 100
        SlowStar(i).Color = RGB(RValue, GValue, BVAlue)
    Next i

    With btmship
        .picHDC = frmSprites.bottomship.hdc
        .maskHDC = frmSprites.bottomshipmask.hdc
        .Left = picGameArea.Width / 2 - 30
        .Top = picGameArea.Height / 2 + 30
        .ImgX = 0
        .Width = 60
        .Height = 60
        .MinY = picGameArea.ScaleHeight / 2 + 25
        .MaxY = picGameArea.ScaleHeight - 100
    End With
    
    With topship
        .picHDC = frmSprites.topship.hdc
        .maskHDC = frmSprites.topshipmask.hdc
        .Left = picGameArea.Width / 2 - 30
        .Top = picGameArea.Height / 2 - 200
        .ImgX = 0
        .Width = 60
        .Height = 60
        .MinY = 10
        .MaxY = picGameArea.ScaleHeight / 2 - 100
    End With

    With Enemy1
        .picHDC = frmSprites.enemy1pic.hdc
        .maskHDC = frmSprites.enemy1mask.hdc
        .Width = 120
        .Height = 60
        .Left = -120
        .Top = picGameArea.ScaleWidth / 2 - 10
        .SpeedLR = 2
    End With
        
    For i = 0 To 2
        Enemy1Ammo(i).y = picGameArea.ScaleHeight + 1
        Enemy1Ammo(i).ImgCount = 0
        Enemy1Ammo(i).ImgNo = 0
        Enemy1Ammo(i).Move = 5
    Next i
    
    
    For i = 0 To 20
            BRec(i).Left = i * picGameArea.ScaleWidth / 20
            BRec(i).Right = BRec(i).Left + picGameArea.ScaleWidth / 25
            BRec(i).Top = picGameArea.ScaleHeight / 2 - 10
            BRec(i).Bottom = picGameArea.ScaleHeight / 2
    Next i

End Sub

Private Sub ShowStars()
Dim i As Integer

        'sets positions of stars (pixel dots) and displays them using PSet function
        'there is a slow moving set (200) in varying shades of grey
        'and some faster moving stars (50) in white
        
        For i = 1 To 100
            If i < 21 Then
                FasterStar(i).x = FasterStar(i).x + FasterStar(i).Move
                picGameArea.ForeColor = FasterStar(i).Color
                picGameArea.FillColor = FasterStar(i).Color
                picGameArea.FillStyle = 0
                Ellipse picGameArea.hdc, FasterStar(i).x, FasterStar(i).y, FasterStar(i).x + FasterStar(i).Diam, FasterStar(i).y + FasterStar(i).Diam
                If FasterStar(i).x > picGameArea.ScaleWidth Then
                    FasterStar(i).x = 0
                    FasterStar(i).y = Rnd * (picGameArea.ScaleHeight + 128) - 64
                    FasterStar(i).Move = Rnd * 2 + 1
                    RValue = Int(FasterStar(i).Move * 80)
                    GValue = RValue
                    BVAlue = RValue
                    FasterStar(i).Color = RGB(RValue, GValue, BVAlue)
                    FasterStar(i).Diam = FasterStar(i).Move + 1
                End If
            End If
            SlowStar(i).x = SlowStar(i).x - SlowStar(i).Move
            SetPixelV picGameArea.hdc, SlowStar(i).x, SlowStar(i).y, SlowStar(i).Color
            
            If SlowStar(i).x < 0 Then
                SlowStar(i).x = picGameArea.ScaleWidth
                SlowStar(i).y = Rnd * (picGameArea.ScaleHeight + 128) - 64
                RValue = Int(Rnd * 200) + 100
                GValue = RValue
                BVAlue = RValue
                SlowStar(i).Color = RGB(RValue, GValue, BVAlue)
            End If
        Next i


End Sub

Private Sub ShowShip(shp As ShipObject, Direction As Integer)
Dim i As Integer

        BitBlt picGameArea.hdc, shp.Left, shp.Top, shp.Width, shp.Height, shp.maskHDC, shp.ImgX, 0, SRCAND
        BitBlt picGameArea.hdc, shp.Left, shp.Top, shp.Width, shp.Height, shp.picHDC, shp.ImgX, 0, SRCPAINT

        
        Select Case shp.BankCount
            Case 0
                shp.ImgX = 0
                shp.HitWidth = 60
            Case 1
                shp.ImgX = shp.Width
                shp.HitWidth = 55
            Case 2
                shp.ImgX = 2 * shp.Width
               shp.HitWidth = 50
            Case 3
               shp.ImgX = 3 * shp.Width
                shp.HitWidth = 40
            Case -1
                shp.ImgX = 4 * shp.Width
                shp.HitWidth = 55
            Case -2
                shp.ImgX = 5 * shp.Width
                shp.HitWidth = 50
            Case -3
                shp.ImgX = 6 * shp.Width
                shp.HitWidth = 40
        End Select

        If shp.Firing Then
            shp.FireTicker = shp.FireTicker + 1
            If shp.FireTicker = 6 Then
                For i = 1 To 15
                    If shp.shipAmmo(i).Fired = False Then
                        shp.shipAmmo(i).Fired = True
                        shp.shipAmmo(i).Power = shp.PowerUp
                        shp.shipAmmo(i).xSpan = 5
                        shp.shipAmmo(i).x = shp.Left + (shp.Width / 2 - 3)
                        If Direction = 1 Then
                            shp.shipAmmo(i).y = shp.Top + shp.Height - 25
                            shp.shipAmmo(i).SideY = shp.Top + shp.Height - 20
                            shp.shipAmmo(i).picno = 3
                        ElseIf Direction = -1 Then
                            shp.shipAmmo(i).y = shp.Top
                            shp.shipAmmo(i).SideY = shp.Top - 5
                            shp.shipAmmo(i).picno = 0
                        End If
                        shp.shipAmmo(i).Move = 25
                        shp.shipAmmo(i).SideXL = shp.Left + shp.Width / 2 - 12
                        shp.shipAmmo(i).SideXR = shp.Left + shp.Width / 2
                        Exit For
                    End If
                Next i
                shp.FireTicker = 0
            End If
        End If

End Sub

Private Sub SetShipPosition(shp As ShipObject, Direction As Integer, FrwdKey As Byte, BkwdKey As Byte, LftKey As Byte, RghtKey As Byte)

        'sets ship position with calls to IsKeyDown sub in 'Functions' public module
        'makes use of the Windows API GetKeyState function
        
        If shp.Hit < 4 Then
            If IsKeyDown(LftKey) Then
                If shp.Left > 10 Then
                    If shp.SpeedLR > -8 Then shp.SpeedLR = shp.SpeedLR - 1
                    shp.Left = shp.Left + shp.SpeedLR
                End If
                If shp.BankCount < 3 Then
                    shp.BankCount = shp.BankCount + 1
                End If
            ElseIf IsKeyDown(RghtKey) Then
                If shp.Left < picGameArea.ScaleWidth - (shp.Width + 10) Then
                    If shp.SpeedLR < 8 Then shp.SpeedLR = shp.SpeedLR + 1
                    shp.Left = shp.Left + shp.SpeedLR
                End If
                If shp.BankCount > -3 Then
                    shp.BankCount = shp.BankCount - 1
                End If
            Else
                If shp.BankCount > 0 Then
                    shp.BankCount = shp.BankCount - 1
                ElseIf shp.BankCount < 0 Then
                    shp.BankCount = shp.BankCount + 1
                End If
                If shp.SpeedLR > 3 Then shp.SpeedLR = shp.SpeedLR - 1
                If shp.SpeedLR < -3 Then shp.SpeedLR = shp.SpeedLR + 1
                If shp.Left > 10 And shp.Left < picGameArea.ScaleWidth - (shp.Width + 10) Then
                    shp.Left = shp.Left + shp.SpeedLR
                End If
            End If
            
            If IsKeyDown(FrwdKey) Then
                If shp.Top > shp.MinY Then
                    If shp.SpeedUD < 8 Then shp.SpeedUD = shp.SpeedUD + 1
                    shp.Top = shp.Top + shp.SpeedUD * Direction
                Else
                    shp.SpeedUD = 0
                End If
                If shp.MinY = picGameArea.ScaleHeight / 2 + 25 Then shp.RocketsOn = True
            ElseIf IsKeyDown(BkwdKey) Then
                If shp.Top < shp.MaxY Then
                    If shp.SpeedUD > -8 Then shp.SpeedUD = shp.SpeedUD - 1
                    shp.Top = shp.Top + shp.SpeedUD * Direction
                Else
                    shp.SpeedUD = 0
                End If
                If shp.MinY = 10 Then shp.RocketsOn = True
            Else
                If shp.SpeedUD > 3 Then shp.SpeedUD = shp.SpeedUD - 1
                If shp.SpeedUD < -3 Then shp.SpeedUD = shp.SpeedUD + 1
                If shp.Top > shp.MinY And shp.Top < shp.MaxY Then
                    shp.Top = shp.Top + shp.SpeedUD * Direction
                End If
            End If
        End If

'        If IsKeyDown(FrwdKey) = False Or IsKeyDown(BkwdKey) = False Then
'            shp.RocketsOn = False
'        End If

End Sub

Private Sub ShowFiring(shp As ShipObject, Direction As Integer)
Dim i As Integer

        For i = 1 To 15
            If shp.shipAmmo(i).Fired = True Then
                shp.shipAmmo(i).y = shp.shipAmmo(i).y + shp.shipAmmo(i).Move * Direction
                If shp.shipAmmo(i).Power = 0 Then
                    BitBlt picGameArea.hdc, shp.shipAmmo(i).x, shp.shipAmmo(i).y, 6, 24, frmSprites.picShot(shp.shipAmmo(i).picno).hdc, 0, 0, SRCPAINT
                ElseIf shp.shipAmmo(i).Power = 1 Then
                    shp.shipAmmo(i).SideXR = shp.shipAmmo(i).SideXR + shp.shipAmmo(i).Move * Cos(1.05)
                    shp.shipAmmo(i).SideXL = shp.shipAmmo(i).SideXL - shp.shipAmmo(i).Move * Cos(1.05)
                    shp.shipAmmo(i).SideY = shp.shipAmmo(i).SideY + shp.shipAmmo(i).Move * Sin(1.05) * Direction
                    BitBlt picGameArea.hdc, shp.shipAmmo(i).x, shp.shipAmmo(i).y, 6, 24, frmSprites.picShot(shp.shipAmmo(i).picno).hdc, 0, 0, SRCPAINT
                    BitBlt picGameArea.hdc, shp.shipAmmo(i).SideXL, shp.shipAmmo(i).SideY, 17, 24, frmSprites.picShot(shp.shipAmmo(i).picno + 1).hdc, 0, 0, SRCPAINT
                    BitBlt picGameArea.hdc, shp.shipAmmo(i).SideXR, shp.shipAmmo(i).SideY, 17, 24, frmSprites.picShot(shp.shipAmmo(i).picno + 2).hdc, 0, 0, SRCPAINT
                End If
                
                If shp.shipAmmo(i).y < -32 Or shp.shipAmmo(i).y > picGameArea.ScaleHeight Then
                    shp.shipAmmo(i).Fired = False
                End If
            End If
        Next i

End Sub

Private Sub ShowShotTrails(shp As ShipObject, Direction As Integer)
Dim k As Integer

        For k = 0 To 6
            If shp.shotTrail(k).LifeTime > 0 Then
                shp.shotTrail(k).LifeTime = shp.shotTrail(k).LifeTime - 1
                shp.shotTrail(k).x = shp.shotTrail(k).x + shp.shotTrail(k).XMov
                shp.shotTrail(k).y = shp.shotTrail(k).y + shp.shotTrail(k).YMov * Direction
                BitBlt picGameArea.hdc, shp.shotTrail(k).x, shp.shotTrail(k).y, 6, 8, frmSprites.rocket.hdc, shp.shotTrail(k).picno, 0, SRCPAINT
                
                If shp.shotTrail(k).LifeTime / 6 - Int(shp.shotTrail(k).LifeTime / 6) = 0 Then
                    shp.shotTrail(k).picno = shp.shotTrail(k).picno + 6
                End If
            
            Else    'particle died and is reseted to the orign
                If shp.Firing Then
                    shp.shotTrail(k).x = shp.Left + Int(Rnd * 4) + (shp.Width / 2 - 4)
                    shp.shotTrail(k).y = shp.Top + shp.Height / 2 + ((shp.Height / 2) * Direction)
                    shp.shotTrail(k).LifeTime = Int(Rnd * 30)
                    shp.shotTrail(k).XMov = Round(Rnd * 3, 2) - 1
                    shp.shotTrail(k).YMov = Round((Rnd * 8), 2) + 4
                    shp.shotTrail(k).picno = 0
                End If
            End If
        Next k

End Sub

Private Sub ShowRocketTrails(shp As ShipObject, Direction As Integer)
Dim k As Integer

        For k = 0 To 25
            If shp.shipTrail(k).LifeTime > 0 Then
                shp.shipTrail(k).LifeTime = shp.shipTrail(k).LifeTime - 1
                shp.shipTrail(k).x = shp.shipTrail(k).x + shp.shipTrail(k).XMov
                shp.shipTrail(k).y = shp.shipTrail(k).y + shp.shipTrail(k).YMov * Direction
                
                BitBlt picGameArea.hdc, shp.shipTrail(k).x + 5, shp.shipTrail(k).y, 6, 8, frmSprites.rocket.hdc, shp.shipTrail(k).picno, 0, SRCPAINT
                BitBlt picGameArea.hdc, shp.shipTrail(k).x - 10, shp.shipTrail(k).y + Int(Rnd * 3), 6, 8, frmSprites.rocket.hdc, shp.shipTrail(k).picno, 0, SRCPAINT
                
                If shp.shipTrail(k).LifeTime / 3 - Int(shp.shipTrail(k).LifeTime / 3) = 0 Then
                    shp.shipTrail(k).picno = shp.shipTrail(k).picno + 6
                End If
            
            Else    'particle died and is reseted to the orign
                If shp.RocketsOn Then
                    shp.shipTrail(k).x = shp.Left + (Int(Rnd * 2) + shp.Width / 2)
                    shp.shipTrail(k).y = shp.Top + 27 + 27 * Direction * -1
                    shp.shipTrail(k).LifeTime = Int(Rnd * 20)
                    shp.shipTrail(k).XMov = Round(Rnd - 0.5, 2)
                    shp.shipTrail(k).YMov = Round((Rnd * 2) - 1, 2) * Direction
                    shp.shipTrail(k).picno = 0
                End If
            End If
        Next k

End Sub

Private Sub ShowBarrier()
Dim j As Integer
Dim Mod1 As Single

    LtTimer = LtTimer + 1
    If LtTimer / 2 - Int(LtTimer / 2) = 0 Then
        CtrLt = CtrLt + 1
        If CtrLt = 30 Then CtrLt = -10
        LtTimer = 0
    End If
    
    For j = 0 To 20
        Mod1 = IIf(Abs(CtrLt - j) > 10, 10, Abs(CtrLt - j))
        picGameArea.ForeColor = 0
        picGameArea.FillColor = RGB(0, 0, 255 - 20 * Mod1)
'        Rectangle picGameArea.hdc, BRec(j).Left - Mod1, BRec(j).Top, BRec(j).Right - Mod1 / 2, BRec(j).Bottom
        RoundRect picGameArea.hdc, BRec(j).Left - Mod1, BRec(j).Top, BRec(j).Right - Mod1 / 2, BRec(j).Bottom, 8, 8
    Next j

End Sub

Private Sub ShowEnemy1()

    Enemy1.Counter = Enemy1.Counter + 1
    If Enemy1.Counter = 100 Then
        Enemy1.OnScreen = True
    End If
    
    If Enemy1.OnScreen Then
        BitBlt picGameArea.hdc, Enemy1.Left, Enemy1.Top, Enemy1.Width, Enemy1.Height, Enemy1.maskHDC, 0, 0, SRCAND
        BitBlt picGameArea.hdc, Enemy1.Left, Enemy1.Top, Enemy1.Width, Enemy1.Height, Enemy1.picHDC, 0, 0, SRCPAINT
        Enemy1.Left = Enemy1.Left + Enemy1.SpeedLR
        If Enemy1.Left >= picGameArea.ScaleWidth + 50 Then
            Enemy1.OnScreen = False
            Enemy1.Counter = 0
            Enemy1.Left = -120
        End If
        
    Dim k As Integer

        For k = 0 To 25
            If Enemy1.shipTrail(k).LifeTime > 0 Then
                Enemy1.shipTrail(k).LifeTime = Enemy1.shipTrail(k).LifeTime - 1
                Enemy1.shipTrail(k).x = Enemy1.shipTrail(k).x + Enemy1.shipTrail(k).XMov
                Enemy1.shipTrail(k).y = Enemy1.shipTrail(k).y + Enemy1.shipTrail(k).YMov
                
                BitBlt picGameArea.hdc, Enemy1.shipTrail(k).x - 63, Enemy1.shipTrail(k).y - 20, 6, 8, frmSprites.rocket.hdc, Enemy1.shipTrail(k).picno, 0, SRCPAINT
                BitBlt picGameArea.hdc, Enemy1.shipTrail(k).x - 63, Enemy1.shipTrail(k).y + 16, 6, 8, frmSprites.rocket.hdc, Enemy1.shipTrail(k).picno, 0, SRCPAINT
                
                If Enemy1.shipTrail(k).LifeTime / 3 - Int(Enemy1.shipTrail(k).LifeTime / 3) = 0 Then
                    Enemy1.shipTrail(k).picno = Enemy1.shipTrail(k).picno + 6
                End If
            
            Else    'particle died and is reseted to the orign
                Enemy1.shipTrail(k).x = Enemy1.Left + (Int(Rnd * 2) + Enemy1.Width / 2)
                Enemy1.shipTrail(k).y = Enemy1.Top + 27
                Enemy1.shipTrail(k).LifeTime = Int(Rnd * 20)
                Enemy1.shipTrail(k).XMov = Round(Rnd - 0.5, 2)
                Enemy1.shipTrail(k).YMov = Round((Rnd * 2) - 1, 2)
                Enemy1.shipTrail(k).picno = 0
            End If
        Next k
    
    End If

End Sub

Private Sub ShowEnemy1Ammo()
Dim i As Integer

    If Enemy1.Counter > 100 Then
        If Enemy1.OnScreen Then
            For i = 0 To 5
                Enemy1Ammo(i).x = Enemy1Ammo(i).x + 1
                Enemy1Ammo(i).TickCounter = Enemy1Ammo(i).TickCounter + 1
                If i < 3 Then
                    Enemy1Ammo(i).y = Enemy1Ammo(i).y + Enemy1Ammo(i).Move
                Else
                    Enemy1Ammo(i).y = Enemy1Ammo(i).y - Enemy1Ammo(i).Move
                End If
                If Enemy1Ammo(i).TickCounter < 80 Then
                    Enemy1Ammo(i).ImgCount = Enemy1Ammo(i).ImgCount + 1
                    If Enemy1Ammo(i).ImgCount = 10 Then
                        Enemy1Ammo(i).ImgNo = Enemy1Ammo(i).ImgNo + 1
                        Enemy1Ammo(i).ImgCount = 0
                    End If
                    If i < 3 Then
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y, 6, 8, frmSprites.rocket2.hdc, Enemy1Ammo(i).ImgNo * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y - 4, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 1) * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y - 8, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 2) * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y - 12, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 3) * 6, 0, SRCPAINT
                    Else
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y, 6, 8, frmSprites.rocket2.hdc, Enemy1Ammo(i).ImgNo * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y + 4, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 1) * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y + 8, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 2) * 6, 0, SRCPAINT
                        BitBlt picGameArea.hdc, Enemy1Ammo(i).x, Enemy1Ammo(i).y + 12, 6, 8, frmSprites.rocket2.hdc, (Enemy1Ammo(i).ImgNo + 3) * 6, 0, SRCPAINT
                    End If
                Else
                    Enemy1Ammo(i).SubCounter = Enemy1Ammo(i).SubCounter + 1
                    If Enemy1Ammo(i).SubCounter < Enemy1Ammo(i).FireTime Then Exit For
                    If i < 3 Then
                        Enemy1Ammo(i).y = Enemy1.Top + Enemy1.Height
                        Enemy1Ammo(i).x = Enemy1.Left + 44 + i * 20
                    Else
                        Enemy1Ammo(i).y = Enemy1.Top
                        Enemy1Ammo(i).x = Enemy1.Left - 17 + i * 20
                    End If
                    Enemy1Ammo(i).ImgCount = 0
                    Enemy1Ammo(i).ImgNo = 0
                    Enemy1Ammo(i).Move = 6
                    Enemy1Ammo(i).TickCounter = 0
                    Enemy1Ammo(i).SubCounter = 0
                    Enemy1Ammo(i).FireTime = Int(Rnd * 10)
                End If
            Next i
        End If
    End If
                

End Sub

Private Sub GameLoop()
Dim k As Integer

    Do While blnGameOn
        picGameArea.Cls
        ShowStars
        ShowBarrier
        SetShipPosition btmship, -1, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        SetShipPosition topship, -1, vbKeyE, vbKeyD, vbKeyS, vbKeyF
        
        ShowShip btmship, -1
        ShowShip topship, 1
        
        ShowRocketTrails btmship, -1
        ShowRocketTrails topship, 1
        
        ShowFiring btmship, -1
        ShowFiring topship, 1
        
        ShowShotTrails btmship, -1
        ShowShotTrails topship, 1
        
        ShowEnemy1
        ShowEnemy1Ammo
        
        DoEvents
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmSprites
    Unload Me
    End
    
End Sub
