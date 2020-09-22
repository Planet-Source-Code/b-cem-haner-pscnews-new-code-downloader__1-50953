VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   About ..."
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -60
      ScaleHeight     =   855
      ScaleWidth      =   5700
      TabIndex        =   0
      Top             =   0
      Width           =   5700
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Planet-source-code.com Latest Codes Sniffer Pro"
         Height          =   195
         Left            =   750
         TabIndex        =   1
         Top             =   600
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   0
      TabIndex        =   2
      Top             =   765
      Width           =   5640
      Begin Project1.lvButtons_H CloseBut 
         Height          =   345
         Left            =   2055
         TabIndex        =   9
         Top             =   3255
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         Caption         =   "&Close"
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
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   1875
         Picture         =   "AboutForm.frx":0000
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIAL THANKS FOR YOUR VOTES !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1125
         TabIndex        =   8
         Top             =   2535
         Width           =   3420
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "©"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2520
         TabIndex        =   7
         Top             =   2175
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email : technical@cemhaner.com"
         Height          =   195
         Index           =   3
         Left            =   1650
         TabIndex        =   6
         Top             =   1935
         Width           =   2370
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2004"
         Height          =   195
         Index           =   2
         Left            =   2715
         TabIndex        =   5
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B.Cem HANER"
         Height          =   195
         Index           =   1
         Left            =   2295
         TabIndex        =   4
         Top             =   1740
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.planet-source-code.com"
         Height          =   195
         Index           =   0
         Left            =   1515
         TabIndex        =   3
         Top             =   2820
         Width           =   2580
      End
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseBut_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Me.Picture2.Picture = Form1.Picture2.Picture
End Sub

