VERSION 5.00
Begin VB.Form Form_main 
   BackColor       =   &H8000000A&
   Caption         =   "Visual Reminder"
   ClientHeight    =   5040
   ClientLeft      =   5640
   ClientTop       =   1965
   ClientWidth     =   3885
   Icon            =   "Form_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3885
   Begin VB.CommandButton Command3 
      Caption         =   "&Remove item"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Current reminders"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   60
      TabIndex        =   6
      Top             =   1620
      Width           =   3735
      Begin VB.ListBox List_output 
         BackColor       =   &H00C0FFC0&
         Height          =   2400
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Add reminder"
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3735
      Begin VB.TextBox Text_input 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   180
         MaxLength       =   80
         TabIndex        =   4
         Text            =   "Message"
         Top             =   420
         Width           =   3375
      End
      Begin VB.TextBox Text_time 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "00:00"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Time:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1020
         Width           =   495
      End
   End
   Begin VB.Timer Timer_monitor 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   1740
      Top             =   4500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Start remainder"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   4620
      Width           =   1455
   End
End
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Add message

'Check if the time is valid (users puts all kind of crap in textboxes so we do this)
If Len(Text_time.text) >= 3 Then
    If Not ((Mid$(Text_time.text, 3, 1) = ":") And IsNumeric(Left$(Text_time.text, 2)) _
    And IsNumeric(Right$(Text_time.text, 2)) And Left(Text_time.text, 2) < 24 _
    And Right(Text_time, 2) < 60) Then
        MsgBox "Invalid time!", vbExclamation, "Error!"
        Exit Sub
    End If
Else:
    MsgBox "Invalid time!", vbExclamation, "Error!"
    Exit Sub
End If
'Check if the message is valid (message should contain at least 3 characters)
If Len(Text_input.text) <= 2 Then
    MsgBox "Too short message / No message!", vbExclamation, "Error!"
    Exit Sub
End If

'If the textboxes are OK, continue

Dim i As Integer
Dim tmpMessage As XMessage

With tmpMessage
    .Message = Text_input.text
    .StartTime = Text_time.text
    .StartHour = Left(Text_time.text, 2)
    .StartMinute = Right(Text_time.text, 2)
    .TimeValue = .StartHour * 60 + .StartMinute
End With

'Also check if a message with the same time already exists
If MessageCount >= 1 Then
    For i = 0 To MessageCount - 1
        If Message(i).TimeValue = tmpMessage.TimeValue Then
            MsgBox "Same time already exists!", vbExclamation, "Error!"
            Exit Sub
        End If
    Next
End If

MessageCount = MessageCount + 1
ReDim Preserve Message(MessageCount - 1) 'Reallocate space to hold our new message

'If theres already at least one other message, organise current message
'to its queue position (in .TimeValue order)
'Since this is quite messy opeation, I added few debugging lines

If MessageCount >= 2 Then
    Debug.Print "messages to comp.: " & MessageCount
    For i = MessageCount - 1 To 1 Step -1
        Debug.Print "compare " & tmpMessage.TimeValue; " to " & Message(i - 1).TimeValue
        If tmpMessage.TimeValue > Message(i - 1).TimeValue Then
            Debug.Print "OK, replacing"
            Message(i) = tmpMessage
            Exit For
        Else:
            If i = 1 Then
                Debug.Print "End of list, replacing"
                Message(1) = Message(0)
                Message(0) = tmpMessage
            Else:
                Debug.Print "FAIL, searching"
                Message(i) = Message(i - 1)
            End If
        End If
    Next
Else:
    Message(MessageCount - 1) = tmpMessage 'If theres no other messages, just copy tmpMessage to array
End If

'Display list of all messages
List_output.Clear 'Clear previous list

For i = 0 To MessageCount - 1 'Browse through messages
    List_output.AddItem Message(i).StartTime & "  " & Message(i).Message
Next

Command2.Enabled = True
End Sub

Private Sub Command2_Click()

InitDirectX 'Initialize directX
Timer_monitor.Enabled = True 'Enable monitoring
Form_main.Hide 'Hide the form
AddToTray 'Sub to add a tray icon
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Command3_Click()

'Remove message and redefine message array
For i = List_output.ListIndex To MessageCount - 2
    Message(i) = Message(i + 1)
Next
List_output.RemoveItem (List_output.ListIndex) 'Remove this item from array
MessageCount = MessageCount - 1

'Since arrays starts at 0, we are using a bit different method to remove items
If MessageCount >= 1 Then
    ReDim Preserve Message(MessageCount - 1) 'Reduce array size by one (MessageCount - 1)
Else:
    MessageCount = 0
    ReDim Message(0)
End If

If MessageCount = 0 Then Command2.Enabled = False 'Enable start reminder button
Command3.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'This controls the tray icon
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX

If TrayIconOn = True Then
    Select Case lMsg
    Case WM_LBUTTONUP
        Call Shell_NotifyIcon(NIM_DELETE, nfIconData) 'delete tray icon
        TrayIconOn = False
        Timer_monitor.Enabled = False 'Disable monitoring
        Form_main.Show 'Show main form
    'We could use the rest to do something else:
    'Case WM_RBUTTONUP
    'Case WM_MOUSEMOVE
    'Case WM_LBUTTONDOWN
    'Case WM_LBUTTONDBLCLK
    'Case WM_RBUTTONDOWN
    'Case WM_RBUTTONDBLCLK
    'Case Else
    End Select
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
ExitApplication
End 'Make sure we are leaving this application
End Sub

Private Sub List_output_Click()

If List_output.SelCount = 0 Then
    Command3.Enabled = False
Else:
    Command3.Enabled = True
End If
End Sub

Private Sub Timer_monitor_Timer()

'The timer is monitoring when to display next message to the screen
'The interval is set to 20 sec. (This doesn't use CPU almost at all)
If Message(MonitoredMessage).StartHour = Hour(Time) And _
    Message(MonitoredMessage).StartMinute = Minute(Time) Then
    DisplayMessage Message(MonitoredMessage).Message
End If
End Sub
