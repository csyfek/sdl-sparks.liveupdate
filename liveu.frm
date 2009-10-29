VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Blake's Live Update"
   ClientHeight    =   1965
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "liveu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1590
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Live Update Control Panel"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Timer Timer2 
         Left            =   5640
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   6120
         Top             =   240
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6000
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Begin Live Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4920
         Picture         =   "liveu.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Begin to Start Live Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim TransferSuccess As Boolean
    UpdateTime = 0
    Timer2.Interval = 1000
    Command1.Enabled = False
    ProgressBar1.Value = 1
    status$ = "Checking for updated version."
    TransferSuccess = GetInternetFile(Inet1, "http://php.indiana.edu/~bpell/liveupdate.html", "c:")

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        Exit Sub
    End If
       
    ProgressBar1.Value = 2
    
    status$ = "Version check success."
    
    Open "c:\liveupdate.html" For Input As #1
        Input #1, updatever$
    Close #1
      
    If updatever$ > myVer Then
        Label1.Caption = "There is an update available to version " + updatever
    Else
        Label1.Caption = "There is no update available."
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If

    status$ = "Getting updated file."

    MsgBox ("This is where you put the updated .exe file")
    TransferSuccess = GetInternetFile(Inet1, "http://ThisIsAFakeURL.com/~myaccount/myfile.exe", "c:\temp")

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If
    
    ProgressBar1.Value = 3
    Timer2.Interval = 0
    
    X = MsgBox("Live Update Complete!", vbInformation)
    Command1.Enabled = True

    frmAbout.Show

End Sub

Private Sub Form_Load()

On Local Error GoTo 200

' myVer = App.Major & "." & App.Minor & "." & App.Revision


' this is where the updated program needs to write it's current version
' number to.  The above commented out line puts the version number in
' the correct format.

status$ = "Idle"
UpdateTime = 0


Open "c:\ver.dat" For Input As #1
    Input #1, myVer
Close #1

Exit Sub

200 myVer = "1.0.0"
X = MsgBox("Version information has not been found, Live Update will assume it's Version 1.0.0")

Resume 205

205 End Sub

Private Sub mnufile_Click()

End Sub

Private Sub Timer1_Timer()
If Inet1.StillExecuting = False Then
    StatusBar1.Panels(1).Text = "Status: Idle"
Else
    StatusBar1.Panels(1).Text = "Status: " & status$
End If

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub
