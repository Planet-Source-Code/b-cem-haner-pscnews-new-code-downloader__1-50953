VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   PSCNews Settings"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3960
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   4665
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         Caption         =   "Save settings on exit"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Index           =   0
         Left            =   -225
         TabIndex        =   14
         Top             =   3285
         Width           =   4815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2325
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   2745
         Width           =   2130
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Auto hide if form IDLE"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   11
         Top             =   1245
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2325
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2355
         Width           =   2100
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "New code message alert"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   8
         Top             =   945
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "New code sound alert"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   7
         Top             =   630
         Width           =   3750
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Autostart with Windows loading"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2340
         TabIndex        =   5
         Text            =   "120"
         Top             =   1965
         Width           =   405
      End
      Begin Project1.lvButtons_H SetBut 
         Height          =   345
         Left            =   165
         TabIndex        =   1
         Top             =   3435
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         Caption         =   "&Default"
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
      Begin Project1.lvButtons_H CloseBut 
         Height          =   345
         Left            =   2985
         TabIndex        =   2
         Top             =   3435
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
         Left            =   1350
         TabIndex        =   3
         Top             =   3435
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   609
         Caption         =   "&Save"
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
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3975
         Picture         =   "Settings.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multilanguage support"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   2805
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default PSC language"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   2415
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh second"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   2040
         Width           =   1125
      End
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseBut_Click()
Me.Hide
End Sub

Private Sub Combo2_Click()
Language App.Path & "\" & Combo2.Text
End Sub

Private Sub Command2_Click()
Form1.Timer1.Tag = 0
SaveMySettings
MsgBox ReadINI(App.Path & "\" & Settings.Combo2.Text, "Main", "00073")

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Visual Basic", 0
Combo1.AddItem "Java/JavaScript", 1
Combo1.AddItem "C/C++", 2
Combo1.AddItem "ASP", 3
Combo1.AddItem "SQL", 4
Combo1.AddItem "Perl", 5
Combo1.AddItem "Delphi", 6
Combo1.AddItem "PHP", 7
Combo1.AddItem "Cold Fusion", 8
Combo1.AddItem ".NET", 9


GetLanguageFiles
End Sub
Sub GetLanguageFiles()
MyPath = App.Path & "\"
myname = Dir(MyPath, vbNormal)
Do While myname <> ""
    If LCase(Right(myname, 4)) = ".lng" Then
    Combo2.AddItem myname
    End If
   myname = Dir
Loop
Combo2.ListIndex = 0
End Sub

Private Sub SetBut_Click()
If MsgBox(ReadINI(App.Path & "\" & Settings.Combo2.Text, "Main", "00072"), vbInformation + vbYesNo) = vbYes Then
    SetDefaultData
    Command2_Click
End If

End Sub
