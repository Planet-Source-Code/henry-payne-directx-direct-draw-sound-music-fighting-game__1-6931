VERSION 5.00
Begin VB.Form Game 
   BorderStyle     =   0  'None
   Caption         =   "Blood Match"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Fight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer ready 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   2640
   End
   Begin VB.Timer match 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4200
      Top             =   2640
   End
   Begin VB.Timer fireanim 
      Interval        =   25
      Left            =   2760
      Top             =   2640
   End
   Begin VB.Timer manreset 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3240
      Top             =   2640
   End
   Begin VB.Timer man2reset 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3720
      Top             =   2640
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim isplayingcheck As Boolean
Dim musictrack As String
Dim ground As String

Dim dxsound As DirectSound
Dim optionshigh As Boolean
Dim practicehigh As Boolean
Dim joinhigh As Boolean
Dim hosthigh As Boolean
Dim quithigh As Boolean


Dim hit1 As DirectSoundBuffer
Dim hit2 As DirectSoundBuffer
Dim death1 As DirectSoundBuffer
Dim death2 As DirectSoundBuffer
Dim punch1 As DirectSoundBuffer
Dim punch2 As DirectSoundBuffer
Dim practice As DirectSoundBuffer
Dim quitbuff As DirectSoundBuffer
Dim joinbuff As DirectSoundBuffer
Dim hostbuff As DirectSoundBuffer
Dim beback As DirectSoundBuffer
Dim astalavista As DirectSoundBuffer
Dim makeday As DirectSoundBuffer
Dim optionsbuff As DirectSoundBuffer
Dim crowd As DirectSoundBuffer
Dim ticktock As DirectSoundBuffer
Dim invincible As DirectSoundBuffer
Dim death3 As DirectSoundBuffer
Dim wah As DirectSoundBuffer
Dim wah2 As DirectSoundBuffer




Dim bufferDesc As DSBUFFERDESC
Dim waveformat  As WAVEFORMATEX


Dim name1 As String
Dim name2 As String

Dim MusicPlay As Boolean
Dim playfx As Boolean
Dim showframe As Boolean
'holds the options

Dim MatchTime As Integer
'for storing the match time

'for pausing the game.
Dim once As String

'for the fire
Dim i As Integer

'hold the damage on the men...
Dim dmgonman1 As Integer
Dim dmgonman2 As Integer

'stuff below is for getting frame rate...
Dim o As Integer
Dim TLast As Single
Dim Fps As Single

'this is for making the menu
Dim screen As String
Dim highlight As String
                                                
Dim winner As String
                                                
'this is for storing the actions...
Dim manaction As String
Dim man2action As String


'this is the positions of the men
Dim man1X As Long
Dim man1Y As Long
Dim man2X As Long
Dim man2Y As Long


Dim DirectX As New DirectX7
'everything is created from this object

Dim ddraw As DirectDraw7
'this is created from the directx object

Dim MainSurf As DirectDrawSurface7
'this surface will hold the pic

Dim Primary As DirectDrawSurface7
Dim Backbuffer As DirectDrawSurface7
'set primary and back buffer

Dim ddsd1 As DDSURFACEDESC2
'describes the primary surface

Dim ddsd2 As DDSURFACEDESC2
'describes our picture

Dim ddsd4 As DDSURFACEDESC2
'holds the screen information


Dim brunning As Boolean
'see if program is still running

Dim CurModeActiveStatus As Boolean
'used to see if users computer can support the display mode

Dim bRestore As Boolean
'an error flag

Dim binit As Boolean
' used to see if everything has been started

Dim file3 As String
'file

Dim man2 As DirectDrawSurface7
'make man2

Dim man1 As DirectDrawSurface7
'make man 1


Dim man2des As DDSURFACEDESC2
'descripotion of man 2

Dim mandescription As DDSURFACEDESC2
'description of man 1

'*******menu*************

Dim menu As DirectDrawSurface7
'make menu

Dim menudes As DDSURFACEDESC2
'description of menu

Dim fire As DirectDrawSurface7
'make fire

Dim firedes As DDSURFACEDESC2
'description of fire


'**************MUSIC******************
Dim FileName As String
'file name of the midi

Dim perf As DirectMusicPerformance
'the performance
Dim seg As DirectMusicSegment
'the segment
Dim segstate As DirectMusicSegmentState
'the state of the segment
Dim loader As DirectMusicLoader
'the loader for the midi file

'~~~~~~~~~~MOUSE~~~~~~~~~~~~~~~~
Dim MouseSurf As DirectDrawSurface7 'This is the variable that holds everything. It wont do anything until it has been
 '"created" though
Dim ddsd5 As DDSURFACEDESC2  'This is a descriptionof the surface. You should already know about these.
Dim MouseX As Integer
Dim MouseY As Integer 'These two values hold the X & Y values
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long 'This API call will be used to stop the normal windows cursor from appearing.

Private Sub fireanim_Timer()
i = i + 1
If i > 5 Then i = 0
'If mainscreen = False Then fireanim.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If once <> "" Then Exit Sub                                         'exit if the ready sign is being displayed
'If screen <> "game" Then Exit Sub

If KeyCode = vbKeyLeft Then
If manaction = "jump" Then manaction = "jumpingleftu": Exit Sub     'if man is in plain jump then make him jump sideways
If Left(manaction, 4) = "jump" Then Exit Sub                        'if man is in air then dont make him walk
If man2X < man1X Then If man1X - 40 <= man2X Then Exit Sub          'if man is at face of other man dont make him walk
man1X = man1X - 10                                                  'make man walk 20 pixels
manaction = "walkL"                                                 'set the mans image to walk
manreset.Enabled = True                                             'enable the timer that resets the man after walking
If man1X < 5 Then man1X = 5                                         'dont let the man walk off the left side of the screen
End If

If KeyCode = vbKeyRight Then
If manaction = "jump" Then manaction = "jumpingrightu": Exit Sub    'if man is in plain jump then make him jump sideways
If Left(manaction, 4) = "jump" Then Exit Sub                        'if man is in air then dont make him walk
If man2X > man1X Then If man1X + 40 >= man2X Then Exit Sub                                'if man is at face of other man dont make him walk
man1X = man1X + 10                                                  'make man walk 20 pixels
manaction = "walkR"                                                 'set the mans image to walk
manreset.Enabled = True                                             'enable the timer that resets the man after walking
If man1X > 575 Then man1X = 575                                     'dont let man walk off the right side of the screen
End If

If KeyCode = vbKeyD Then
If man2action = "jump" Then man2action = "jumpingleftu": Exit Sub
If Left(man2action, 4) = "jump" Then Exit Sub
If man1X < man2X Then If man2X - 40 <= man1X Then Exit Sub
man2X = man2X - 10
man2action = "walkL"
man2reset.Enabled = True
If man2X > 575 Then man2X = 575                                     'dont let man walk off the right side of the screen
End If

If KeyCode = vbKeyG Then
If man2action = "jump" Then man2action = "jumpingrightu": Exit Sub
If Left(man2action, 4) = "jump" Then Exit Sub
If man1X > man2X Then If man2X + 40 >= man1X Then Exit Sub
man2X = man2X + 10
man2action = "walkR"
man2reset.Enabled = True
If man2X < 5 Then man2X = 5
End If
'move the men back and forward


If KeyCode = vbKeyUp Then
If manaction = "walkL" Then manaction = "jumpingleftu": Exit Sub   'if man is walking then make him run-jump
If manaction = "walkR" Then manaction = "jumpingrightu": Exit Sub  'if man is walking right then make him run-jump
If Left(manaction, 4) = "jump" Then Exit Sub                       'if man is in air then dont make him jump again
manaction = "jump"                                                 'set the mans graphic to jump
End If

If KeyCode = vbKeyR Then
If man2action = "walkL" Then man2action = "jumpingleftu": Exit Sub   'if man is walking then make him run-jump
If man2action = "walkR" Then man2action = "jumpingrightu": Exit Sub  'if man is walking right then make him run-jump
If Left(man2action, 4) = "jump" Then Exit Sub                       'if man is in air then dont make him jump again
man2action = "jump"                                                 'set the mans graphic to jump
End If


'code below is for punching man1 from man2
If KeyCode = vbKeyQ Then
'If man2action = "kick" Or man2action = "punch" Then Exit Sub
  If Left(man2action, 4) = "walk" Or man2action = "normal" Then
    man2action = "punch"
    punch2.play DSBPLAY_DEFAULT
    If man1X > man2X Then
      If man2X + 40 >= man1X And man1Y > man2Y - 40 Then
        manaction = "gethit"
        manreset = True
      End If
    ElseIf man1X < man2X Then
      If man2X - 40 <= man1X And man1Y > man2Y - 40 Then
        manaction = "gethit"
        manreset = True
      End If
    End If
  Else
    Exit Sub
  End If
man2reset.Enabled = True
End If


'code below is for kicking man1 from man2
If KeyCode = vbKeyW Then
If Left(man2action, 11) = "jumpingleft" Then man2action = "jumpkickl": Exit Sub
If Left(man2action, 12) = "jumpingright" Then man2action = "jumpkickr": Exit Sub
If Left(man2action, 4) = "jump" Then man2action = "jumpkick": Exit Sub
'If man2action = "kick" Or man2action = "punch" Then Exit Sub
  If Left(man2action, 4) = "walk" Or man2action = "normal" Then
    man2action = "kick"
    If man1X > man2X Then
      If man2X + 60 >= man1X And man1Y > man2Y - 20 Then
        manaction = "getkick"
        manreset = True
      End If
    ElseIf man1X < man2X Then
      If man2X - 60 <= man1X And man1Y > man2Y - 20 Then
        manaction = "getkick"
        manreset = True
      End If
    End If
  Else
    Exit Sub
  End If
man2reset.Enabled = True
End If


'code below is for punching man2 from man1
If KeyCode = vbKeyInsert Then
'If manaction = "kick" Or manaction = "punch" Then Exit Sub
  If Left(manaction, 4) = "walk" Or manaction = "normal" Then
    manaction = "punch"
    punch1.play DSBPLAY_DEFAULT
    
    If man2X < man1X Then
      If man1X - 40 <= man2X And man2Y > man1Y - 40 Then
        man2action = "gethit"
        man2reset = True
      End If
    ElseIf man2X >= man1X Then
      If man1X + 40 >= man2X And man2Y > man1Y - 40 Then
        man2action = "gethit"
        man2reset = True
      End If
    End If
  Else
    Exit Sub
  End If
manreset.Enabled = True
End If

'code below is for kicking man2 from man1
If KeyCode = vbKeyHome Then
If Left(manaction, 11) = "jumpingleft" Then manaction = "jumpkickl": Exit Sub
If Left(manaction, 12) = "jumpingright" Then manaction = "jumpkickr": Exit Sub
If Left(manaction, 4) = "jump" Then manaction = "jumpkick": Exit Sub
'If manaction = "kick" Or manaction = "punch" Then Exit Sub
  If Left(manaction, 4) = "walk" Or manaction = "normal" Then
    manaction = "kick"
    If man2X < man1X Then
      If man1X - 60 <= man2X And man2Y > man1Y - 20 Then
        man2action = "getkick"
        man2reset = True
      End If
    ElseIf man2X > man1X Then
      If man1X + 60 >= man2X And man2Y > man1Y - 20 Then
        man2action = "getkick"
        man2reset = True
      End If
    End If
  Else
    Exit Sub
  End If
manreset.Enabled = True
End If

If KeyCode = vbKeyEscape Then
EndIt
End If


If KeyCode = vbKeySpace Then
If screen = "win" Then
screen = "menu"
bRestore = True

man1X = 500
man1Y = 250

man2X = 100
man2Y = 250

manaction = "normal"
man2action = "normal"

dmgonman2 = 100
dmgonman1 = 100

MatchTime = 60

End If
End If

'bRestore = True



End Sub

Private Sub Form_Load()



ChDir App.Path

ShowCursor False
     

Set Backbuffer = Nothing
Set Primary = Nothing
Set MouseSurf = Nothing
Set man1 = Nothing
Set man2 = Nothing
Set fire = Nothing
Set menu = Nothing

playfx = False
MusicPlay = False
showframe = False

name1 = "Player 1"
name2 = "Player 2"

man1X = 500
man1Y = 250

man2X = 100
man2Y = 250

manaction = "normal"
man2action = "normal"

bRestore = False
screen = "menu"

dmgonman2 = 100
dmgonman1 = 100

MatchTime = 60

Open "options.ini" For Input As #1
Dim tempvar As String
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then MusicPlay = True
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then playfx = True
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then showframe = True
Line Input #1, tempvar
name1 = tempvar
Line Input #1, tempvar
name2 = tempvar
Line Input #1, tempvar
If Right(tempvar, 10) = "Sunset    " Then
  ground = "Sunset"
Else
  ground = "Lighthouse"
End If
Line Input #1, tempvar
If Right(tempvar, 13) = "Mortal Kombat" Then
  musictrack = "Mortal Kombat"
Else
  musictrack = "Terminator 2"
End If


Close #1

Call soundinit

Call playmusic

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if screen <> "menu" Then Exit Sub
If X > 175 And X < 475 And Y > 225 And Y < 275 Then
  screen = "game"
  bRestore = True
  Backbuffer.SetForeColor RGB(4, 67, 1)
  once = "getready"
ElseIf X > 175 And X < 425 And Y > 375 And Y < 425 Then
  Call ddraw.RestoreDisplayMode 'sets the screen resolution back to what it was before the program was started.
  Call ddraw.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL) 'tells DirectX that the application is no longer 'EXCLUSIVE. If we were to just drop the program it would crash, as DirectX will still be expecting an exclusive application 'to be running
  WindowState = 1
  Options.Show
  Options.SetFocus
  ShowCursor True
ElseIf X > 175 And X < 425 And Y > 425 And Y < 475 Then
  Call EndIt
End If

End Sub

Private Sub Form_Paint()
    BLT    'Blt when windows tells your window to refresh itself
End Sub

 
Sub Init()
On Error GoTo errOut
' goes to sub on a error

Set ddraw = DirectX.DirectDrawCreate("")
'this creates a object in the ddraw thing

Me.Show
'show the screen

Call ddraw.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
'change all the display shit to stuff.

Call ddraw.SetDisplayMode(640, 480, 16, 0, DDSDM_DEFAULT)
'set to 640 by 480

ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
'get the screen surface and create backbuffer

ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
'describes the surface

ddsd1.lBackBufferCount = 1
'this shows that there will be 1 back buffer

Set Primary = ddraw.CreateSurface(ddsd1)
'creates the surface from the info in ddsd1

Dim caps As DDSCAPS2

caps.lCaps = DDSCAPS_BACKBUFFER
'get the primary surface to create a backbuffer from the caps description

Set Backbuffer = Primary.GetAttachedSurface(caps)
'defines it as a backbuffer

Backbuffer.GetSurfaceDesc ddsd4
'makes the ddsd4 hold all the info on the current settings, ie hight and width

Backbuffer.SetFontTransparency True
'makes any writing appear transparent

Backbuffer.SetForeColor RGB(4, 67, 1)
'this makes teh text green

InitSurfaces
'init the surfaces

binit = True
brunning = True


Restart:

Do While brunning
BLT
DoEvents
If ready.Enabled = True Then GoTo Ending
Loop
'that is the main program loop


errOut:
EndIt
'the error handler

Ending:
GoTo Restart
End Sub
 
 
Sub InitSurfaces()

Dim FileX As String
'hold the bitmaps filename

Set MainSurf = Nothing
'clears the surface


ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
'these thing load the bitmap

ddsd2.lWidth = ddsd4.lWidth
'makes the surface the same as the screen

ddsd2.lHeight = ddsd4.lHeight
'same as above

'The height & width can be specified in pixels, and can be bigger than the screen. The bitmap will be 'stretched/squashed to fit these settings. NOTE: a big surface will require lots of memory, which means low performance and 'potential crashes on low-spec machines
If screen = "menu" Then
FileX = App.Path & "\Graphics\Menu Background.bmp"
End If

If screen = "game" Then
  If ground = "Sunset" Then
    FileX = App.Path & "\Graphics\Ground1.bmp"
  ElseIf ground = "Lighthouse" Then
    FileX = App.Path & "\Graphics\Ground2.bmp"
  End If
End If

If screen = "win" Then
FileX = App.Path & "\Graphics\Menu Background.bmp"
End If


'set filex

Set MainSurf = ddraw.CreateSurfaceFromFile(FileX, ddsd2)
'create a new surface from the specified file, based on the ddsd2 description

Call makemenu
Call makeman2
Call makeman
Call makemouse
Call makefire

Dim key As DDCOLORKEY 'A transparency description, similiar to a surface description
key.low = 0 'the transparencies use a range, anything between the low value and the high value is transparent. In this case 'only value 0(black) will be transparent
key.high = 0
man1.SetColorKey DDCKEY_SRCBLT, key 'Applies the colour key to the mouse surface.
man2.SetColorKey DDCKEY_SRCBLT, key
MouseSurf.SetColorKey DDCKEY_SRCBLT, key 'Applies the colour key to the mouse surface.
menu.SetColorKey DDCKEY_SRCBLT, key
fire.SetColorKey DDCKEY_SRCBLT, key
'transpercy for man1 & 2 & mouse & menu & fire

End Sub


Sub EndIt()
Call ddraw.RestoreDisplayMode 'sets the screen resolution back to what it was before the program was started.
Call ddraw.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL) 'tells DirectX that the application is no longer 'EXCLUSIVE. If we were to just drop the program it would crash, as DirectX will still be expecting an exclusive application 'to be running
End 'end the program

End Sub

Sub BLT()
On Local Error GoTo errOut
'error handler

'loop music
'If perf.IsPlaying = False Then Call playmusic
isplayingcheck = perf.IsPlaying(seg, segstate)
If isplayingcheck = False Then Call perf.PlaySegment(seg, 0, 0)

'setting up the font
myfont.Name = "Comic Sans MS"   'a system font.
myfont.Bold = True
myfont.Size = 10
myfont.Bold = False

'sets the font to the backbuffer
Backbuffer.SetFont myfont  'Apply the font to the surface.


If binit = False Then Exit Sub
'if not initilized then exit

Dim ddrval As Long
'holds the number that directdraw returns when it is blting

Dim rMain As RECT
'this will keep us from trying to blt in case we lose the surfaces (alt-tab)

Do Until ExModeActive
'way of saying "do until it returns true"

DoEvents
'lets windows do its things

Loop
'if something goes wrong then we need to
'get back the surfaces...

DoEvents
If bRestore Then
bRestore = False
ddraw.RestoreAllSurfaces
InitSurfaces
End If
'must init the surfaces again if they
'we're lost. When this happens the
'first line of initsurfaces is
'important



rMain.Bottom = ddsd2.lHeight
rMain.Right = ddsd2.lWidth
'this gets the area of the bitmap we want to
'blt (all of it)


'blt to the backbuffer from our surface
'to the screen surface such that our bitmap
'appears over the windows

ddrval = Backbuffer.BltFast(0, 0, MainSurf, rMain, DDBLTFAST_WAIT)
'this shoves the rmain area onto the back buffer
'Dim r As Long
'Dim g As Long
'Dim b As Long
'Dim v As Long

If screen = "game" Then    'do if not in the menu
Call game
End If

If screen = "menu" Then     'do while in the menu
Call showmenu
Call showfire
Call Do_Mouse
End If

If screen = "win" Then
Call showfire
Call Do_Mouse
If winner = "man1" Then Call Backbuffer.DrawText(260, 240, name1 + " Was The Winner", False)
If winner = "man2" Then Call Backbuffer.DrawText(260, 240, name2 + " Was The Winner", False)
Call Backbuffer.DrawText(260, 300, "Press Space To Continue", False)
End If


If o = 30 Then
If TLast <> 0 Then Fps = 30 / (Timer - TLast)
TLast = Timer
o = 0
End If
o = o + 1
If showframe = True Then Call Backbuffer.DrawText(250, 400, "Frames per Second " + Format$(Fps, "#.0"), False)



If once = "getready" Then
  Call Backbuffer.DrawText(305, 210, "Get Ready", False)
  ready.Enabled = True
ElseIf once = "fight" Then
  Call Backbuffer.DrawText(307, 210, "FIGHT!", False)
  ready.Enabled = True
End If

Primary.Flip Nothing, DDFLIP_WAIT
'this flips the backbuffer the the screen
'this is what the user sees

errOut:
'skips redrwawing the scene

End Sub


Function ExModeActive() As Boolean
    Dim TestCoopRes As Long 'holds the return value of the test.
    TestCoopRes = ddraw.TestCooperativeLevel 'Tells DDraw to do the test

    If (TestCoopRes = DD_OK) Then
        ExModeActive = True 'everything is fine
    Else
        ExModeActive = False 'this computer doesn't support this mode
    End If
End Function




Sub showman()
Dim manval As Long
Dim rman As RECT

If manaction = "normal" Then rman.Left = 1
If manaction = "punch" Then rman.Left = 121
If manaction = "kick" Then rman.Left = 181
If manaction = "getjk" Then rman.Left = 241
If manaction = "gethit" Then rman.Left = 241
If manaction = "hit" Then rman.Left = 301
If manaction = "getkick" Then rman.Left = 361
If manaction = "kicked" Then rman.Left = 421
If Left(manaction, 4) = "jump" Then rman.Left = 481
If Left(manaction, 4) = "walk" Then rman.Left = 61
If Left(manaction, 8) = "jumpkick" Then rman.Left = 541: wah.play DSBPLAY_DEFAULT


rman.Bottom = 100
rman.Right = rman.Left + 59

manval = Backbuffer.BltFast(man1X, man1Y, man1, rman, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)

End Sub




Sub showman2()
Dim man2val As Long
Dim rman2 As RECT

If man2action = "normal" Then rman2.Left = 1
If man2action = "punch" Then rman2.Left = 121
If man2action = "kick" Then rman2.Left = 181
If man2action = "getjk" Then rman2.Left = 241
If man2action = "gethit" Then rman2.Left = 241
If man2action = "hit" Then rman2.Left = 301
If man2action = "getkick" Then rman2.Left = 361
If man2action = "kicked" Then rman2.Left = 421
If Left(man2action, 4) = "jump" Then rman2.Left = 481
If Left(man2action, 4) = "walk" Then rman2.Left = 61
If Left(man2action, 8) = "jumpkick" Then rman2.Left = 541: wah2.play DSBPLAY_DEFAULT


rman2.Top = 0
rman2.Bottom = 100
rman2.Right = rman2.Left + 59

man2val = Backbuffer.BltFast(man2X, man2Y, man2, rman2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
End Sub




Sub makeman2()
Dim FileX As String


man2des.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
man2des.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
'describe man 2

man2des.lWidth = 600
man2des.lHeight = 200
'size of man 2

FileX = App.Path & "\Graphics\Blue Man.bmp"
'bmp of man2

Set man2 = ddraw.CreateSurfaceFromFile(FileX, man2des)
'make man in memory

End Sub



Sub makeman()
Dim FileX As String


mandescription.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
mandescription.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
'describe the man

mandescription.lWidth = 600
mandescription.lHeight = 200
'man's size

FileX = App.Path & "\Graphics\Yellow Man.bmp"
'mans bmp

Set man1 = ddraw.CreateSurfaceFromFile(FileX, mandescription)
'make man
End Sub



Private Sub man2reset_Timer()
If Left(man2action, 3) = "get" Then hit2.play DSBPLAY_DEFAULT
If Left(man2action, 4) = "jump" Then GoTo skip

If man2action = "gethit" Then
  man2action = "hit"
  dmgonman2 = dmgonman2 - 3
  Exit Sub
End If

If man2action = "getkick" Then
  man2action = "kicked"
  dmgonman2 = dmgonman2 - 2
  Exit Sub
End If

If man2action = "getjk" Then
  man2action = "hit"
  dmgonman2 = dmgonman2 - 4
  Exit Sub
End If


If man2action <> "normal" Then man2action = "normal"
If man2Y <> 250 Then man2Y = 250
man2reset.Enabled = False

skip:
End Sub

Private Sub manreset_Timer()
If Left(manaction, 3) = "get" Then hit1.play DSBPLAY_DEFAULT
If Left(manaction, 4) = "jump" Then GoTo skip

If manaction = "gethit" Then
  manaction = "hit"
  dmgonman1 = dmgonman1 - 3
  Exit Sub
End If

If manaction = "getkick" Then
  manaction = "kicked"
  dmgonman1 = dmgonman1 - 2
  Exit Sub
End If

If manaction = "getjk" Then
  manaction = "hit"
  dmgonman1 = dmgonman1 - 4
  Exit Sub
End If



If manaction <> "normal" Then manaction = "normal"
If man1Y <> 250 Then man1Y = 250
manreset.Enabled = False

skip:
End Sub
Sub playmusic()
'''''***************This is the music crap***************
On Local Error GoTo error

Set perf = DirectX.DirectMusicPerformanceCreate()   'creates the buffer called perf
Call perf.Init(Nothing, 0)   'clears the buffer
perf.SetPort -1, 80 ' set the port

Call perf.SetMasterAutoDownload(True)   'fuck knows, but it doesnt work without it
perf.SetMasterVolume (50) 'set the volume

Set loader = Nothing        'clear the loader
Set loader = DirectX.DirectMusicLoaderCreate  'create the loader

If musictrack = "Mortal Kombat" Then
  FileName = App.Path & "\Music\Sound.mid"  'set the file to load
ElseIf musictrack = "Terminator 2" Then
  FileName = App.Path & "\Music\Sound2.mid"
End If


Set seg = loader.LoadSegment(FileName)  ' load the file into seg using the loader
        
seg.SetStandardMidiFile  'set seg as a standard midi file

Call play

Exit Sub
error:
MsgBox Err.Description & Err.Source

End Sub
Sub play()

If MusicPlay = True Then Call perf.PlaySegment(seg, 0, 0)  'play the segment
Init  'Calls the initialize sequence when the form is loaded
End Sub

Sub makemouse()

Dim fileY As String
ddsd5.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN 'These were explained in the beginner tutorials
ddsd5.lWidth = 32
ddsd5.lHeight = 32

'^^ Creates a 32x32 surface, this will hold a single frame cursor that is 32x32.
fileY = App.Path & "\Graphics\Transparent Mouse.bmp"   'You should already have a variable defined (If you followed the beginner tutorials)
                                        ' If you haven't, define FileX as a string: Dim FileX as string
Set MouseSurf = ddraw.CreateSurfaceFromFile(fileY, ddsd5) 'Create "MouseSurf" from the mouse.bmp file.
End Sub


Sub Do_Mouse()
Dim Rval2 As Long    'to store the return value
Dim rMouse As RECT  'a RECT to define where to copy from
rMouse.Top = 0
rMouse.Left = 0
rMouse.Right = 32
rMouse.Bottom = 32
Rval2 = Backbuffer.BltFast(MouseX, MouseY, MouseSurf, rMouse, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)  'IMPORTANT: ALL THIS NEEDS TO BE ON THE SAME LINE. When you paste it into VB,
'make sure it all stays on the same line.
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = Int(X)
MouseY = Int(Y)

If X > 175 And X < 475 And Y > 225 And Y < 275 Then
highlight = "prac"
ElseIf X > 175 And X < 475 And Y > 275 And Y < 325 Then
highlight = "join"
ElseIf X > 175 And X < 475 And Y > 325 And Y < 375 Then
highlight = "host"
ElseIf X > 175 And X < 475 And Y > 375 And Y < 425 Then
highlight = "opt"
ElseIf X > 175 And X < 475 And Y > 425 And Y < 475 Then
highlight = "quit"
Else
highlight = ""
End If



End Sub



Sub makemenu()
Dim FileX As String
menudes.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
menudes.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

menudes.lWidth = 300
menudes.lHeight = 500


FileX = App.Path & "\Graphics\Menu.bmp"

Set menu = ddraw.CreateSurfaceFromFile(FileX, menudes)
End Sub
Sub showmenu()
Dim showprac As Long
Dim showhost As Long
Dim showjoin As Long
Dim showoptions As Long
Dim showquit As Long

Dim rprac As RECT
Dim rhost As RECT
Dim rjoin As RECT
Dim roptions As RECT
Dim rquit As RECT

rprac.Top = 1
rhost.Top = 100
rjoin.Top = 200
roptions.Top = 300
rquit.Top = 400




If highlight = "prac" Then
rprac.Top = 50
If practicehigh <> True Then practice.play DSBPLAY_DEFAULT: practicehigh = True: optionshigh = False: joinhigh = False: hosthigh = False: quithigh = False


ElseIf highlight = "host" Then
rhost.Top = 150
If hosthigh <> True Then hostbuff.play DSBPLAY_DEFAULT: practicehigh = False: optionshigh = False: joinhigh = False: hosthigh = True: quithigh = False

ElseIf highlight = "join" Then
rjoin.Top = 250
If joinhigh <> True Then joinbuff.play DSBPLAY_DEFAULT: practicehigh = False: optionshigh = False: joinhigh = True: hosthigh = False: quithigh = False

ElseIf highlight = "opt" Then
roptions.Top = 350
If optionshigh <> True Then optionsbuff.play DSBPLAY_DEFAULT: practicehigh = False: optionshigh = True: joinhigh = False: hosthigh = False: quithigh = False

ElseIf highlight = "quit" Then
rquit.Top = 450
If quithigh <> True Then quitbuff.play DSBPLAY_DEFAULT: practicehigh = False: optionshigh = False: joinhigh = False: hosthigh = False: quithigh = True

End If

'setting other limits
rhost.Bottom = rhost.Top + 50
rhost.Right = 300

rprac.Bottom = rprac.Top + 50
rprac.Right = 300

rjoin.Bottom = rjoin.Top + 50
rjoin.Right = 300

roptions.Bottom = roptions.Top + 50
roptions.Right = 300

rquit.Bottom = rquit.Top + 50
rquit.Right = 300

'making buttons
showjoin = Backbuffer.BltFast(180, 275, menu, rjoin, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
showhost = Backbuffer.BltFast(175, 325, menu, rhost, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
showprac = Backbuffer.BltFast(180, 225, menu, rprac, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
showoptions = Backbuffer.BltFast(175, 375, menu, roptions, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
showquit = Backbuffer.BltFast(175, 425, menu, rquit, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub
Sub showfire()
Dim firevalleft As Long
Dim firevalright As Long
Dim rc As RECT
'Dim i As Integer
 
' See what we need to blt
If i = 0 Then rc.Left = 1
If i = 1 Then rc.Left = 86
If i = 2 Then rc.Left = 171
If i = 3 Then rc.Left = 1
If i = 4 Then rc.Left = 86
If i = 5 Then rc.Left = 171
        
rc.Right = rc.Left + 84
        
If i > 2 Then
rc.Top = 90
Else
rc.Top = 1
End If
       
rc.Bottom = rc.Top + 88
        
Dim l As Long
Dim number As Long
number = 20

For l = 1 To 13
Backbuffer.SetForeColor RGB(number, number, number)
Call Backbuffer.DrawBox(540 + l, 178, 566 - l, 480)
Call Backbuffer.DrawBox(80 + l, 178, 106 - l, 480)
number = number + 10
Next l

' Increment and see the i value
firevalright = Backbuffer.BltFast(510, 100, fire, rc, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
firevalleft = Backbuffer.BltFast(50, 100, fire, rc, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


End Sub


Sub makefire()

Dim FileX As String
menudes.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
menudes.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

menudes.lWidth = 320
menudes.lHeight = 200

FileX = App.Path & "\Graphics\Fire.bmp"

Set fire = ddraw.CreateSurfaceFromFile(FileX, firedes)
End Sub


Sub man2jump()
If man2action = "jump" Then
 man2Y = man2Y - 4                                   'if man is jumping then make him rise
 If man2Y <= 125 Then man2action = "jumpun"           'if man has reached the max height then make him fall
End If

If man2action = "jumpun" Then
  man2Y = man2Y + 4                                  'if man is falling then make him fall
  If man2Y >= 250 Then man2action = "normal": man2Y = 250        'make man stand when he hits the ground
End If

If man2action = "jumpingleftu" Then
  man2X = man2X - 2                                 'if man is jumping left then make him move left
  man2Y = man2Y - 4                                 'if man is jumping left then make him move up
  If man2Y <= 125 Then man2action = "jumpingleftd"   'if man reaches max height then make him fall
  If man2X <= 10 Then man2action = "jumpun"
End If

If man2action = "jumpingleftd" Then
  man2X = man2X - 2                                 'if man is falling then make him move left
  man2Y = man2Y + 4                                 'if man is falling then make him move down
  If man2Y >= 250 Then man2action = "normal"         'if man hits ground then make him normal stance
  If man2X <= 10 Then man2action = "jumpun"
End If

If man2action = "jumpingrightu" Then
  man2X = man2X + 2                                 'if man is jumping left then make him move right
  man2Y = man2Y - 4                                 'if man is jumping left then make him move up
  If man2Y <= 125 Then man2action = "jumpingrightd"  'if man reaches max height then make him fall
  If man2X > 575 Then man2action = "jumpun"
End If

If man2action = "jumpingrightd" Then
  man2X = man2X + 2                                 'if man is falling then make him move right
  man2Y = man2Y + 4                                 'if man is falling then make him move down
  If man2Y >= 250 Then man2action = "normal"         'if man hits ground then make him normal stance
  If man2X >= 575 Then man2action = "jumpun"
End If


If man2action = "jumpkick" Then
  man2Y = man2Y + 4
  If man2Y = 250 Then man2action = "normal"
  If man2X <= 10 Then man2action = "jumpun"
End If

If man2action = "jumpkickl" Then
  man2X = man2X - 4
  man2Y = man2Y + 4
  If man2Y = 250 Then man2action = "normal"
  If man2X <= 10 Then man2action = "jumpun"
End If

If man2action = "jumpkickr" Then
  man2X = man2X + 4
  man2Y = man2Y + 4
  If man2Y = 250 Then man2action = "normal"
  If man2X >= 575 Then man2action = "jumpun"
End If

If man2X > man1X Then
  If man1X + 60 > man2X And man2Y < man1Y And man2action = "jumpkick" Then
    manaction = "getjk"
    manreset = True
  End If
  If man1X + 60 > man2X And man2Y < man1Y And man2action = "jumpkickl" Then
    manaction = "getjk"
    manreset = True
  End If
ElseIf man2X < man1X Then
  If man1X - 60 < man2X And man2Y < man1Y And man2action = "jumpkick" Then
    manaction = "getjk"
    manreset = True
  End If
  If man1X - 60 < man2X And man2Y < man1Y And man2action = "jumpkickr" Then
    manaction = "getjk"
    manreset = True
  End If
End If


End Sub

Sub manjump()
If manaction = "jump" Then
 man1Y = man1Y - 4                                   'if man is jumping then make him rise
 If man1Y <= 125 Then manaction = "jumpun"            'if man has reached the max height then make him fall
End If

If manaction = "jumpun" Then
  man1Y = man1Y + 4                                  'if man is falling then make him fall
  If man1Y >= 250 Then manaction = "normal": man2Y = 250          'make man stand when he hits the ground
End If

If manaction = "jumpingleftu" Then
  man1X = man1X - 2                                 'if man is jumping left then make him move left
  man1Y = man1Y - 4                                 'if man is jumping left then make him move up
  If man1Y <= 125 Then manaction = "jumpingleftd"    'if man reaches max height then make him fall
  If man1X <= 10 Then manaction = "jumpun"
End If

If manaction = "jumpingleftd" Then
  man1X = man1X - 2                                  'if man is falling then make him move left
  man1Y = man1Y + 4                                 'if man is falling then make him move down
  If man1Y >= 250 Then manaction = "normal"          'if man hits ground then make him normal stance
  If man1X <= 10 Then manaction = "jumpun"
End If

If manaction = "jumpingrightu" Then
  man1X = man1X + 2                                 'if man is jumping left then make him move right
  man1Y = man1Y - 4                                 'if man is jumping left then make him move up
  If man1Y <= 125 Then manaction = "jumpingrightd"   'if man reaches max height then make him fall
  If man1X >= 575 Then manaction = "jumpun"
End If

If manaction = "jumpingrightd" Then
  man1X = man1X + 2                                 'if man is falling then make him move right
  man1Y = man1Y + 4                                 'if man is falling then make him move down
  If man1Y >= 250 Then manaction = "normal"          'if man hits ground then make him normal stance
  If man1X >= 575 Then manaction = "jumpun"
End If

If manaction = "jumpkick" Then
  man1Y = man1Y + 4
  If man1Y >= 250 Then manaction = "normal"
  If man1X <= 10 Then manaction = "jumpun"
End If

If manaction = "jumpkickl" Then
  man1X = man1X - 4
  man1Y = man1Y + 4
  If man1Y >= 250 Then manaction = "normal"
  If man1X <= 10 Then manaction = "jumpun"
End If

If manaction = "jumpkickr" Then
  man1X = man1X + 4
  man1Y = man1Y + 4
  If man1Y >= 250 Then manaction = "normal"
  If man1X >= 575 Then manaction = "jumpun"
End If

If man1X > man2X Then
  If man2X + 60 > man1X And man1Y < man2Y And manaction = "jumpkick" Then
    man2action = "getjk"
    man2reset = True
  End If
  If man2X + 60 > man1X And man1Y < man2Y And manaction = "jumpkickl" Then
    man2action = "getjk"
    man2reset = True
  End If
ElseIf man1X < man2X Then
  If man2X - 60 < man1X And man1Y < man2Y And manaction = "jumpkick" Then
    man2action = "getjk"
    man2reset = True
  End If
  If man2X - 60 < man1X And man1Y < man2Y And manaction = "jumpkickr" Then
    man2action = "getjk"
    man2reset = True
  End If
End If


End Sub

Private Sub match_Timer()
MatchTime = MatchTime - 1
If MatchTime = 0 Then Call DisplayWin
End Sub

Private Sub ready_Timer()
Randomize Timer
Dim mang As Integer
mang = Int(Rnd(1) * 6)


If once = "fight" Then once = "": match.Enabled = True: If mang = 3 Then makeday.play DSBPLAY_DEFAULT

If once = "getready" Then once = "fight"
ready.Enabled = False
End Sub


Sub showmanother()
Dim manval As Long
Dim rman As RECT

If manaction = "normal" Then rman.Left = 541
If manaction = "punch" Then rman.Left = 421
If manaction = "kick" Then rman.Left = 361
If manaction = "getjk" Then rman.Left = 301
If manaction = "gethit" Then rman.Left = 301
If manaction = "hit" Then rman.Left = 241
If manaction = "getkick" Then rman.Left = 181
If manaction = "kicked" Then rman.Left = 121
If Left(manaction, 4) = "jump" Then rman.Left = 61
If Left(manaction, 4) = "walk" Then rman.Left = 481
If Left(manaction, 8) = "jumpkick" Then rman.Left = 1

rman.Top = 100
rman.Bottom = 200
rman.Right = rman.Left + 59

manval = Backbuffer.BltFast(man1X, man1Y, man1, rman, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)

End Sub
Sub showman2other()
Dim man2val As Long
Dim rman2 As RECT

If man2action = "normal" Then rman2.Left = 541
If man2action = "punch" Then rman2.Left = 421
If man2action = "kick" Then rman2.Left = 361
If man2action = "getjk" Then rman2.Left = 301
If man2action = "gethit" Then rman2.Left = 301
If man2action = "hit" Then rman2.Left = 241
If man2action = "getkick" Then rman2.Left = 181
If man2action = "kicked" Then rman2.Left = 121
If Left(man2action, 4) = "jump" Then rman2.Left = 61
If Left(man2action, 4) = "walk" Then rman2.Left = 481
If Left(man2action, 8) = "jumpkick" Then rman2.Left = 1


rman2.Top = 100
rman2.Bottom = 200
rman2.Right = rman2.Left + 59

man2val = Backbuffer.BltFast(man2X, man2Y, man2, rman2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
End Sub


Sub DisplayWin()
screen = "win"
bRestore = True

match.Enabled = False
fireanim.Enabled = True
man2reset.Enabled = False
manreset.Enabled = False
ready.Enabled = False

Randomize Timer
Dim mang As Integer
mang = Int(Rnd(1) * 4)


If dmgonman1 < dmgonman2 Then winner = "man1": death2.play DSBPLAY_DEFAULT: If mang = 1 Then beback.play DSBPLAY_DEFAULT: If mang = 3 Then invincible.play DSBPLAY_DEFAULT

If dmgonman1 > dmgonman2 Then winner = "man2": death3.play DSBPLAY_DEFAULT: If mang = 2 Then astalavista.play DSBPLAY_DEFAULT



End Sub


Sub game()
'below is some stuff to stop drawing off the screen
If man1X > man2X Then
  If man1X < (man2X + 30) And man1Y = man2Y Then
    man1X = man1X + 20
    man2X = man2X - 20
  End If
ElseIf man1X < man2X Then
  If (man1X + 30) > man2X And man1Y = man2Y Then
    man1X = man1X - 20
    man2X = man2X + 20
  End If
End If

If dmgonman1 <= 0 Then DisplayWin
If dmgonman2 <= 0 Then DisplayWin

Call manjump                  'update the mans jumping status, if any
Call man2jump                 'update man2's jumping status, if any

If man1X > man2X Then
  Call showman2other                 'show man2
  Call showman                  'show man1
ElseIf man1X < man2X Then
  Call showman2
  Call showmanother
End If
     
Dim r As Long
Dim g As Long
Dim b As Long
Dim v As Long
     
r = 4                         'set the R,G & B values for the forecolor
g = 67
b = 1

For v = 1 To 10                                                  'this lot of code draws
Backbuffer.SetForeColor RGB(r, g, b)                             'the linear gradient
r = r + 2                                                        'for each mans life
g = g + 30
b = b + 1
Call Backbuffer.DrawBox(10, 20 + v, ((3 * dmgonman2) + 10), 40 - v)
Call Backbuffer.DrawBox((630 - (3 * dmgonman1)), 20 + v, 630, 40 - v)
Next v


Call Backbuffer.DrawText(313, 50, Int(MatchTime), False)
Call Backbuffer.DrawText(20, 45, name1, False)    'draw each players life under
Call Backbuffer.DrawText(560, 45, name2, False)   'his lifebar
End Sub




Sub soundinit()

Set dxsound = DirectX.DirectSoundCreate("")

dxsound.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY

bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
waveformat.nFormatTag = WAVE_FORMAT_PCM
waveformat.nChannels = 2
waveformat.lSamplesPerSec = 44100
waveformat.nBitsPerSample = 16
waveformat.nBlockAlign = waveformat.nBitsPerSample / 8 * waveformat.nChannels
waveformat.lAvgBytesPerSec = waveformat.lSamplesPerSec * waveformat.nBlockAlign
  
  
If playfx = False Then GoTo 564

Set hit1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\hit1.wav", bufferDesc, waveformat)
Set hit2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\hit2.wav", bufferDesc, waveformat)
Set death1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\death1.wav", bufferDesc, waveformat)
Set death2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\death2.wav", bufferDesc, waveformat)
Set punch1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\punch1.wav", bufferDesc, waveformat)
Set punch2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\punch2.wav", bufferDesc, waveformat)
Set practice = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\practice.wav", bufferDesc, waveformat)
Set quitbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\quit.wav", bufferDesc, waveformat)
Set joinbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\join.wav", bufferDesc, waveformat)
Set hostbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\host.wav", bufferDesc, waveformat)
Set beback = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\beback.wav", bufferDesc, waveformat)
Set astalavista = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\astalavista.wav", bufferDesc, waveformat)
Set makeday = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\makeday.wav", bufferDesc, waveformat)
Set optionsbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\options.wav", bufferDesc, waveformat)
Set crowd = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\crowdah.wav", bufferDesc, waveformat)
Set ticktock = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\ticktock.wav", bufferDesc, waveformat)
Set invincible = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\invincible.wav", bufferDesc, waveformat)
Set death3 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\death3.wav", bufferDesc, waveformat)
Set wah = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\wah.wav", bufferDesc, waveformat)
Set wah2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\wah2.wav", bufferDesc, waveformat)
GoTo 743

564
Set hit1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set hit2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set death1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set death2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set punch1 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set punch2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set practice = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set quitbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set joinbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set hostbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set beback = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set astalavista = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set makeday = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set optionsbuff = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set crowd = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set ticktock = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set invincible = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set death3 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set wah = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
Set wah2 = dxsound.CreateSoundBufferFromFile(App.Path & "\sounds\nosound.wav", bufferDesc, waveformat)
743


End Sub

