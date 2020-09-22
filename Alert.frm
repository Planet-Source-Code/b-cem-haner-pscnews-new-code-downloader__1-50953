VERSION 5.00
Begin VB.Form Alert 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   ScaleHeight     =   1290
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   105
      Width           =   4065
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   15
         Top             =   15
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PSC CODE !"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   555
         TabIndex        =   1
         Top             =   270
         Width           =   3045
      End
   End
End
Attribute VB_Name = "Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
OnTop Me.hwnd
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
