Attribute VB_Name = "Module1"
Global inter As CWinInetConnection
Global Const MySound = "notify.wav"
Global MyControl As String
Global BenimAlan As String
Global myStr As String


Public Function StringDondur(GelenStr As String, Keyword As String, Ayrac As String) As String
Dim KeyWordBul As Integer
Dim DonenSon As Integer


Keyword = Keyword + "="
StringDondur = ""
KeyWordBul = InStr(1, GelenStr, Keyword)
If KeyWordBul > 0 Then
    DonenSon = InStr(KeyWordBul, GelenStr, Ayrac)
    If DonenSon <= 0 Then
        StringDondur = Mid(GelenStr, KeyWordBul + Len(Keyword))
    Else
        StringDondur = Mid(GelenStr, KeyWordBul + Len(Keyword), DonenSon - KeyWordBul - Len(Keyword))
    End If
End If

End Function

Public Function GetMyString(Txt As String, Position As Long, Keyword As String, BackwardSTR As String, ForwardSTR As String) As String
Dim X As Long, Y As Long, t As Long
        X = InStr(Position, Txt, Keyword)
        Y = InStr(X + Len(Keyword), Txt, ForwardSTR)
        t = InStrRev(Txt, BackwardSTR, X) + Len(BackwardSTR)
        GetMyString = Mid(Txt, t, Y - t) & Space(5) & "Last=" & Y
End Function

Public Function RightToWalk(Txt As String, Keyword As String, ToADD As Long, SeeTOFin As String) As String
Dim X As Long, Y As Long, t As Long
        X = InStr(1, Txt, Keyword) + Len(Keyword)
        Y = InStr(X + ToADD, Txt, SeeTOFin)
        RightToWalk = Mid(Txt, X + ToADD, Y - (X + ToADD))
End Function
Public Function RemoveDoubleSpace(Txt As String)
Dim X As Long, Ark As String
Ark = Txt
      Do
        X = InStr(1, Ark, "  ")
        If X > 0 Then
            Ark = Left(Ark, X) & Right(Ark, Len(Ark) - (X + 1))
        Else
            RemoveDoubleSpace = Ark
            Exit Function
        End If
      Loop
      
End Function

Public Function Language(MyFile As String, Optional param As String) As Boolean

If MyFile = "" Then
    MyFile = App.Path & "\English.LNG"
End If

Form1.Label1.Caption = ReadINI(MyFile, "Main", "00047")
Form1.Label2.Caption = ReadINI(MyFile, "Main", "00001")
AboutForm.Label2.Caption = ReadINI(MyFile, "Main", "00001")
Form1.SetBut.Caption = ReadINI(MyFile, "Main", "00051")
Form1.TrayBut.Caption = ReadINI(MyFile, "Main", "00052")
Form1.Command2.Caption = ReadINI(MyFile, "Main", "00053")
Form1.CloseBut.Caption = ReadINI(MyFile, "Main", "00054")
AboutForm.Label1(1).Caption = ReadINI(MyFile, "Main", "00039") & " B.Cem HANER"
AboutForm.Label1(5).Caption = ReadINI(MyFile, "Main", "00040")
AboutForm.CloseBut.Caption = ReadINI(MyFile, "Main", "00054")
AboutForm.Caption = "   " & ReadINI(MyFile, "Main", "00057")

Settings.Check1.Caption = ReadINI(MyFile, "Main", "00041")
Settings.Check2.Caption = ReadINI(MyFile, "Main", "00042")
Settings.Check3.Caption = ReadINI(MyFile, "Main", "00043")
Settings.Check4.Caption = ReadINI(MyFile, "Main", "00044")
Settings.Label1(0).Caption = ReadINI(MyFile, "Main", "00045")
Settings.Label1(1).Caption = ReadINI(MyFile, "Main", "00046")
Settings.Label1(2).Caption = ReadINI(MyFile, "Main", "00048")


Settings.Caption = "   PSCNews " & ReadINI(MyFile, "Main", "00050")
Settings.SetBut.Caption = ReadINI(MyFile, "Main", "00055")
Settings.Command2.Caption = ReadINI(MyFile, "Main", "00056")
Settings.CloseBut.Caption = ReadINI(MyFile, "Main", "00054")

End Function

Public Sub SaveMySettings()
With Settings
    SaveINI App.Path & "\PSCNews.INI", "Main", "AutoStart", .Check1.Value
    
   If .Check1.Value = 1 Then
' IF YOUR PROJECT CONVERT EXE AND IF USE .... REMOVE REM ;)
'        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\", "Run\", App.Title, App.Path & "\" & App.EXEName & ".exe"
    End If
    
    SaveINI App.Path & "\PSCNews.INI", "Main", "SoundAlert", .Check2.Value
    SaveINI App.Path & "\PSCNews.INI", "Main", "MessageAlert", .Check3.Value
    SaveINI App.Path & "\PSCNews.INI", "Main", "AutoHide", .Check4.Value
    SaveINI App.Path & "\PSCNews.INI", "Main", "SaveOnExit", .Check5.Value
    SaveINI App.Path & "\PSCNews.INI", "Main", "DefaultLanguage", .Combo2.Text
    SaveINI App.Path & "\PSCNews.INI", "Main", "RefreshSecond", .Text1.Text
    SaveINI App.Path & "\PSCNews.INI", "Main", "Language", .Combo1.ListIndex
    SaveINI App.Path & "\PSCNews.INI", "Main", "DefaultLanguage", .Combo2.Text
End With
End Sub
Public Sub ReadMySettings()
On Error GoTo MyError
With Settings
    .Check1.Value = ReadINI(App.Path & "\PSCNews.INI", "Main", "AutoStart")
    .Check2.Value = ReadINI(App.Path & "\PSCNews.INI", "Main", "SoundAlert")
    .Check3.Value = ReadINI(App.Path & "\PSCNews.INI", "Main", "MessageAlert")
    .Check4.Value = ReadINI(App.Path & "\PSCNews.INI", "Main", "AutoHide")
    .Check5.Value = ReadINI(App.Path & "\PSCNews.INI", "Main", "SaveOnExit")
    .Text1.Text = ReadINI(App.Path & "\PSCNews.INI", "Main", "RefreshSecond")
    .Combo1.ListIndex = ReadINI(App.Path & "\PSCNews.INI", "Main", "Language")
    .Combo2.Text = ReadINI(App.Path & "\PSCNews.INI", "Main", "DefaultLanguage")
    Language App.Path & "\" & .Combo2.Text
End With
Exit Sub
MyError:
    SetDefaultData
    SaveMySettings

End Sub
Public Sub PlaySound()
    If Dir(App.Path & "\" & MySound) <> "" Then
        sndPlaySound App.Path & "\" & MySound, SND_ASYNC
    End If
End Sub
Public Sub URLOpen(Gelen As String)
ShellExecute Form1.hwnd, "Open", Gelen, "", "", 1
End Sub
Public Sub NewAlert()
    'If New message sound checked
    If Settings.Check2.Value = 1 Then
        PlaySound
    End If
    
    'If New message sound checked
    If Settings.Check2.Value = 1 Then
        Alert.Show vbModal, Form1
    End If
    
    
End Sub
Public Sub SetDefaultData()
With Settings
    .Check1.Value = 1
    .Check2.Value = 1
    .Check3.Value = 1
    .Check4.Value = 0
    .Check5.Value = 1
    .Text1.Text = 20
    .Combo1.ListIndex = 0
    .Combo2.Text = "English.LNG"
    Language App.Path & .Combo2.Text
End With
End Sub
