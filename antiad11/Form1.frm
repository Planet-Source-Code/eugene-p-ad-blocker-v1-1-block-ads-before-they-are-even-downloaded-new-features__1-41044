VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AD Blocker 1.1 - By Eugene - http://www.eugenius.tk"
   ClientHeight    =   5955
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Send Me TO Tray !"
      Height          =   330
      Left            =   4500
      TabIndex        =   20
      Top             =   3765
      Width           =   1620
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   2385
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "Form1.frx":1242
      Top             =   2235
      Width           =   3900
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Statistics"
      Height          =   2745
      Left            =   105
      TabIndex        =   10
      Top             =   3045
      Width           =   6165
      Begin VB.TextBox txtFake 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "Form1.frx":12AC
         Top             =   1485
         Width           =   5865
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enable Stats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   1305
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5595
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   80
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   135
         X2              =   5985
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display In Place of Blocked ADs: (can be html, see my example)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   1230
         Width           =   5460
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*starts a small web server, must keep app running*"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1620
         TabIndex        =   14
         Top             =   285
         Width           =   3555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         Height          =   195
         Left            =   1185
         TabIndex        =   12
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AD's Blocked:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   585
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Block !"
      Height          =   405
      Left            =   90
      TabIndex        =   7
      Top             =   2175
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   105
      TabIndex        =   6
      Top             =   1740
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "Form1.frx":1335
      Left            =   3240
      List            =   "Form1.frx":1D83
      TabIndex        =   4
      Top             =   240
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operating System Detection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   3000
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   75
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1125
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows XP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   2820
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows 95/98/ME"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   2805
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows NT/2000"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   750
         Width           =   2820
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   105
      TabIndex        =   18
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VOTE FOR ME !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3731D&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.eugenius.tk"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   90
      X2              =   6240
      Y1              =   2100
      Y2              =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AD Server Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   30
      Width           =   1815
   End
   Begin VB.Menu mnuPop 
      Caption         =   "wHATEvER !"
      Begin VB.Menu mnuRestore 
         Caption         =   "Display &AD Blocker"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStat 
         Caption         =   "00000000"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As Integer
X = 0
ProgressBar1.Max = List1.ListCount

If Text2 <> "" Then  'IF TEXTBOX2 doesnt have nothing THEN
    Open Text2 & "hosts" For Append As #1 'OPEN the HOSTS File For Appending Write
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# AD Blocker List By - Eugene"
    
    Do Until X = List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
        ProgressBar1.Value = X 'increase the Progressbar value by 1
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing
    
    MsgBox "AD Blocker ! Injected into the system, process complete ! Thank you for using AD Blocker !  !", vbInformation, "Complete !"
End If

End Sub

Private Sub Command2_Click()
    Winsock1.Close 'Close Winsock or RESET
    Winsock1.Listen 'Listen for new Connections
End Sub

Private Sub Command3_Click()
    Me.Visible = False 'hide me
    Call AddToTray(Me, "[eugenius] AD Block 1.1", Me.Icon) 'add to system tray
End Sub

Private Sub Form_Load()
If FExist(App.Path & "\database.txt") = True Then 'if the database.txt FIle Exists then
    Call Loadlistbox(App.Path & "\database.txt", List1) 'Load the contents of the file to the listbox1
Else
    Call SaveListBox(App.Path & "\database.txt", List1) 'Save the default contents to the database.txt
End If

    Label1.Caption = "AD Server Database: " & List1.ListCount & " blocked" 'Display the current # ad servers !
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    mnuPop.Visible = False
    
'///// FIND the HOST FILE
If FExist("c:\windows\system32\drivers\etc\hosts") = True Then
    Option1.ForeColor = &HC000&
    Option1.FontBold = True
    Option1.Caption = Option1.Caption & " *DETECTED*"
    Option1.Value = True
    Text2 = "c:\windows\system32\drivers\etc\"
ElseIf FExist("c:\windows\hosts") = True Then
    Option2.ForeColor = &HC000&
    Option2.FontBold = True
    Option2.Caption = Option2.Caption & " *DETECTED*"
    Option2.Value = True
    Text2 = "c:\windows\"
ElseIf FExist("c:\winnt\system32\drivers\etc\hosts") = True Then
    Option3.ForeColor = &HC000&
    Option3.FontBold = True
    Option3.Caption = Option3.Caption & " *DETECTED*"
    Option3.Value = True
    Text2 = "c:\winnt\system32\drivers\etc\"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    
        Case WM_RBUTTONUP    'When RIGHT mouse button is UP
            PopupMenu mnuPop 'Pop UP the hidden MenuPop
        Case WM_RBUTTONDOWN

    End Select
End Sub

Private Sub Form_Terminate()
    Call RemoveFromTray  'Remove the tray icon if its still alive
End Sub

Private Sub Form_Unload(Cancel As Integer)
X = MsgBox("Are you sure you want to EXIT ? Pressing NO will send to Tray !, you might want to send to tray if your running the Statistics Server !", vbQuestion + vbYesNo, "Are you sure you want to Exit")
If X = vbNo Then  'if users presses no then
    Cancel = 1     'abort the unloading
    Command3_Click 'goto Command3_Click sub (send to tray)
    Exit Sub       'Exit sub
End If

    Call RemoveFromTray 'Remove the tray icon if its still alive
End Sub

Private Sub mnuRestore_Click()
    Call RemoveFromTray   'remove tray icon
    Me.Show               'show me !
End Sub

Private Sub Option1_Click()
    Text2 = "c:\windows\system32\drivers\etc\"  'set the FOLDER where the HOST file EXISTS
End Sub

Private Sub Option2_Click()
    Text2 = "c:\windows\" 'set the FOLDER where the HOST file EXISTS
End Sub

Private Sub Option3_Click()
    Text2 = "c:\winnt\system32\drivers\etc\" 'set the FOLDER where the HOST file EXISTS
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID   'ACCEPT the CONNECTION
    Winsock1.SendData txtFake   'SEND THE Replacement AD to Browser
    Label6 = Label6 + 1         'ADD 1 to the Statisitics Ticker
    mnuStat.Caption = "ADs Blocked: " & Label6    'set the System Tray Menu to display the curreny AD's Blocked
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.Close 'Close the COnnection
    Winsock1.Listen '(LISTEN) Wait for another connection
End Sub


