Attribute VB_Name = "MainCode"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'API for sleep command
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'API for playing wav-file
   
Const SND_SYNC = &H0         '  play synchronously (default)
Const SND_NODEFAULT = &H2    '  silence not default, if sound not found

Public Type XDX
    DX As DirectX7
    DDraw As DirectDraw7
    MainSurface As DirectDrawSurface7
    SurfaceDesc As DDSURFACEDESC2
    ReminderFont As IFont
End Type
Public Type XMessage
    StartTime As String
    StartHour As Byte
    StartMinute As Byte
    TimeValue As Integer
    Message As String
End Type
Public DX As XDX 'DirectX objects
Public Message() As XMessage 'Message objects
Public MessageCount As Integer 'Total message count
Public MonitoredMessage As Integer 'Current message number being monitored

Sub Main()
    
    'Add preprocessors here if any
    Form_main.Show
    DoEvents
End Sub
Sub InitDirectX()

Dim i As Integer
Set DX.ReminderFont = New StdFont 'Create the font we are using to draw text

Set DX.DX = New DirectX7 'Create main directX -object
Set DX.DDraw = DX.DX.DirectDrawCreate("") 'Create main DirectDraw -object
DX.DDraw.SetCooperativeLevel Form_main.hwnd, DDSCL_NORMAL 'Set co-operative level

'Create target surface, where to blit our graphics (direct screen!)
With DX.SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
Set DX.MainSurface = DX.DDraw.CreateSurface(DX.SurfaceDesc)

End Sub
Sub DrawText(text As String, X As Integer, Y As Integer, outerColor As Long, innerColor As Long)

'We are blitting the text 4 times. Each to a bit different location to produce font edge effect
'After that we blit the text to the center of all other blits to produce 'inner' color
DX.MainSurface.SetForeColor outerColor
DX.MainSurface.DrawText X, Y, text, False
DX.MainSurface.DrawText X + 2, Y, text, False
DX.MainSurface.DrawText X, Y + 2, text, False
DX.MainSurface.DrawText X + 2, Y + 2, text, False

DX.MainSurface.SetForeColor innerColor
DX.MainSurface.DrawText X + 1, Y + 1, text, False
End Sub

Sub DisplayMessage(MsgString As String)
Dim i As Long
Dim i2 As Long

' Initialize fixed message font
With DX.ReminderFont
    .Size = 14
    .Name = "Arial"
    .Bold = True
    .Underline = False
End With

DX.MainSurface.SetFont DX.ReminderFont

sndPlaySound "msgsound.wav", SND_SYNC Or SND_NODEFAULT 'Play wave

For i = 0 To 3
    DrawText "Reminder!", 20, 20, RGB(255, 255, 255), RGB(0, 0, 0) 'Draw reminder text
    Sleep 400
    DrawText "Reminder!", 20, 20, RGB(0, 0, 0), RGB(255, 255, 255) 'Invert text to blank it
    Sleep 400
Next

' Initialize message font (We are using fixed-width font to create animation more easily)
With DX.ReminderFont
    .Size = 14
    .Name = "FixedSys"
    .Bold = True
    .Underline = False
End With

DX.MainSurface.SetFont DX.ReminderFont

'Display animation
For i = 1 To Len(MsgString) - 1
    DrawText Mid$(MsgString, i, 1), 11 + (i * 9), 45, RGB(255, 255, 255), RGB(0, 0, 0)
    DrawText Mid$(MsgString, i + 1, 1), 11 + ((i + 1) * 9), 45, RGB(0, 0, 0), RGB(255, 255, 255)
    Sleep 50
Next
'Blink last character
For i2 = 0 To 4
    DrawText Mid$(MsgString, i, 1), 11 + (i * 9), 45, RGB(255, 255, 255), RGB(0, 0, 0)
    Sleep 400
    DrawText Mid$(MsgString, i, 1), 11 + (i * 9), 45, RGB(0, 0, 0), RGB(255, 255, 255)
    Sleep 400
Next

'Monitor next message (If there isn't next message, start in the beginning)
If MessageCount - 1 = MonitoredMessage Then
    MonitoredMessage = 0
Else:
    MonitoredMessage = MonitoredMessage + 1
End If
End Sub
Sub ExitApplication()

'Free memory
Set DX.MainSurface = Nothing
Set DX.DDraw = Nothing
Set DX.DX = Nothing
End Sub
