VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   4560
      Left            =   105
      ScaleHeight     =   4530
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   90
      Width           =   5520
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   30
         ScaleHeight     =   2475
         ScaleWidth      =   5400
         TabIndex        =   15
         Top             =   900
         Width           =   5430
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   30
            TabIndex        =   16
            Top             =   30
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   4260
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            PictureAlignment=   5
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   16777215
            BackColor       =   8421504
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3555
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":030A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   345
         Left            =   60
         TabIndex        =   14
         Top             =   3525
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         Caption         =   "?"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Form1.frx":0624
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2940
         Top             =   165
      End
      Begin Project1.lvButtons_H lvButtons_H1 
         Height          =   270
         Left            =   5160
         TabIndex        =   12
         Top             =   3540
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "Form1.frx":062A
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H SetBut 
         Height          =   345
         Left            =   60
         TabIndex        =   11
         Top             =   3915
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         Caption         =   "Settings ..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H TrayBut 
         Height          =   345
         Left            =   1230
         TabIndex        =   10
         Top             =   3930
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         Caption         =   "&Hide"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H CloseBut 
         Height          =   345
         Left            =   3930
         TabIndex        =   9
         Top             =   3930
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Command2 
         Height          =   345
         Left            =   2580
         TabIndex        =   8
         Top             =   3930
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         Caption         =   "&Update"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox SearchText 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3435
         TabIndex        =   6
         Top             =   3570
         Width           =   1665
      End
      Begin Project1.KnightRider KRi 
         Height          =   180
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4320
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   318
         ForeColor       =   65535
         Effect          =   1
         Speed           =   10
         Tail            =   10
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -45
         Picture         =   "Form1.frx":0983
         ScaleHeight     =   855
         ScaleWidth      =   5700
         TabIndex        =   3
         Top             =   0
         Width           =   5700
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   4380
            Top             =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Planet-source-code.com Latest Codes Sniffer Pro"
            Height          =   195
            Left            =   750
            TabIndex        =   7
            Top             =   600
            Width           =   3480
         End
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Index           =   1
         Left            =   -30
         TabIndex        =   2
         Top             =   3435
         Width           =   5565
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Index           =   0
         Left            =   -75
         TabIndex        =   1
         Top             =   855
         Width           =   5595
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   270
         Left            =   3345
         Shape           =   4  'Rounded Rectangle
         Top             =   3540
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Search on PSC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   5
         Top             =   3585
         Width           =   1305
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   615
      Picture         =   "Form1.frx":10799
      Top             =   4905
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   90
      Picture         =   "Form1.frx":10AA3
      Top             =   4905
      Width           =   480
   End
   Begin VB.Menu MyMenu 
      Caption         =   "MyMenu"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Updater 
         Caption         =   "Update"
      End
      Begin VB.Menu Ayar 
         Caption         =   "Settings..."
      End
      Begin VB.Menu About 
         Caption         =   "About..."
      End
      Begin VB.Menu sprt 
         Caption         =   "-"
      End
      Begin VB.Menu MyExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim nid As NOTIFYICONDATA

Private Sub About_Click()
AboutForm.Show vbModal, Me
End Sub

Private Sub Ayar_Click()
Settings.Show vbModal, Me
End Sub

Private Sub CloseBut_Click()
MyExit_Click
End Sub








Private Sub Command2_Click()
ATABeni
End Sub

Private Sub Form_Activate()
ReadMySettings
TrayBut.Refresh
CloseBut.Refresh
Command2.Refresh
SetBut.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Timer2.Tag = 0
End Sub

Private Sub Form_Load()
'Me.Move Screen.Width - Me.Width, Screen.Height - (Me.Height + 500)
Me.Move Screen.Width - Me.Width, Screen.Height - (Me.Height + 500)
TrayBut_Click

End Sub



Public Sub ATABeni()
KRi.Enabled = True

Dim Oku As String
'Close #1: Open "C:\PSC\gelen.txt" For Input As #1

'a = 0
'Do Until EOF(1)
'    Line Input #1, Oku
    'If InStr(1, Oku, "<marq") > 0 Then a = 1
'    'If a > 0 Then BenimAlan = BenimAlan & Oku
'    BenimAlan = BenimAlan & Oku
'Loop
BenimAlan = GetUrlSource("http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=" & Settings.Combo1.ListIndex + 1)

If InStr(1, BenimAlan, "<marq") < 1 Then
    KRi.Enabled = False
Exit Sub
End If

Dim Met As String, Py As Long, X As String, r As String

   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "1", 350
   ListView1.ColumnHeaders.Add , , "2", 5000
   ListView1.View = lvwReport

Dim itmX As ListItem
t = 0
Py = 1
On Error GoTo Hata
For i = 1 To 9
        Met = GetMyString(BenimAlan, Py, "ShowCode.asp?", "/", "</")
            Py = Val(StringDondur(Met, "Last", ";"))
            r = Left(Met, InStr(1, Met, Chr(34)) - 1)
            X = RemoveDoubleSpace(RightToWalk(Format(Met), ">", 0, "Last"))
         
            'If previous data = New data ;)
            If i = 1 And X <> MyControl Then
                NewAlert
            End If
            If i = 1 Then MyControl = X
         
         Set itmX = ListView1.ListItems.Add(, , " ")
         ListView1.ListItems(1).ForeColor = QBColor(13)
         itmX.SubItems(1) = X
         itmX.Tag = r
         itmX.SmallIcon = 1
Next i

KRi.Enabled = False
Exit Sub
Hata:
    KRi.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
        PopupMenu MyMenu
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
       Case WM_RBUTTONDOWN

          PopupMenu MyMenu
          
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select



End Sub

Private Sub ListView1_Click()
Timer2.Tag = 0
End Sub

Private Sub ListView1_DblClick()
Timer2.Tag = 0
If ListView1.ListItems.Count < 1 Then Exit Sub
    URLOpen "http://www.planet-source-code.com/vb/scripts/" & ListView1.SelectedItem.Tag
End Sub

Private Sub lvButtons_H1_Click()
Timer2.Tag = 0
URLOpen "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=" & SearchText.Text & "&lngWId=" & Settings.Combo1.ListIndex + 1
End Sub

Private Sub lvButtons_H2_Click()
Timer2.Tag = 0
AboutForm.Show vbModal, Me
End Sub

Private Sub MyExit_Click()
Timer2.Tag = 0
If MsgBox(ReadINI(App.Path & "\" & Settings.Combo2.Text, "Main", "00074"), vbInformation + vbYesNo) = vbYes Then
    If Settings.Check5.Value = 1 Then SaveMySettings
    TrayBut_Click
    Shell_NotifyIcon NIM_DELETE, nid
    End
End If
    
End Sub

Private Sub SearchText_KeyPress(KeyAscii As Integer)
Timer2.Tag = 0
End Sub

Private Sub SetBut_Click()
Timer2.Tag = 0
Settings.Show vbModal, Me
End Sub

Private Sub Show_Click()
Timer2.Tag = 0
Me.Width = 5745
Me.Height = 4800
Me.Show

Me.Move Screen.Width - Me.Width, Screen.Height - (Me.Height + 500)

End Sub

Private Sub Timer1_Timer()
'Set inter = New CWinInetConnection
'If inter.IsConnected Then
'    ATABeni
'End If
Timer1.Tag = Val(Timer1.Tag) + 1
If Val(Timer1.Tag) >= Val(Settings.Text1.Text) Then
    ATABeni
    Timer1.Tag = 0
End If
End Sub

Private Sub Timer2_Timer()
If Val(Timer2.Tag) = 30 And Settings.Check4.Value = 1 Then
    TrayBut_Click
End If
Timer2.Tag = Val(Timer2.Tag) + 1
End Sub

Private Sub TrayBut_Click()
Timer2.Tag = 0
For i = Me.Height To 0 Step -50
Me.Top = Me.Top + 50
Me.Height = i
DoEvents
Next

SendTRAY
End Sub
Sub SendTRAY()

   'Set the individual values of the NOTIFYICONDATA data type.
   nid.cbSize = Len(nid)
   nid.hwnd = Form1.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Form1.Icon
   nid.szTip = "PSCNews Pro 2.0" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Updater_Click()
ATABeni
End Sub
