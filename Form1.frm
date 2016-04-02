VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Words V3.0 单词狂背 负一的平方根制作"
   ClientHeight    =   3075
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   12585
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "音标(F2)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11280
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "看答案(F1)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   120
      TabIndex        =   1
      Text            =   "test"
      Top             =   720
      Width           =   12375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   9855
   End
   Begin VB.Label Label1 
      Caption         =   "test"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
   End
   Begin VB.Menu MenuFile 
      Caption         =   "文件"
      Begin VB.Menu FileExit 
         Caption         =   "退出"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuDict 
      Caption         =   "词库"
      Begin VB.Menu DictOpen 
         Caption         =   "打开词库..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu DictReload 
         Caption         =   "重新加载词库"
         Shortcut        =   {F4}
      End
      Begin VB.Menu DictView 
         Caption         =   "查看词库..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MenuTool 
      Caption         =   "工具"
      Begin VB.Menu ToolStat 
         Caption         =   "统计..."
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助"
      Begin VB.Menu HelpAbout 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R As String, S() As String, i As Long
Dim P As Long, C As Long, LineCount As Long
Dim Word(10000) As String
Dim Pron(10000) As String
Dim Expl(10000) As String
Dim N As Long
Dim FN As String
Dim Rt As Long, Wr As Long

Function LoadWords() As Boolean
    On Error GoTo Errp
    
    Open App.Path & "\" & FN & ".txt" For Input As #1
    P = 0
    LineCount = 0
    While Not EOF(1)
        P = P + 1
        Line Input #1, R
        Word(P) = R
        Line Input #1, R
        Pron(P) = R
        Line Input #1, R
        Expl(P) = R
        LineCount = LineCount + 1
    Wend
    Close
    C = P
    For i = 1 To C
        Expl(i) = Replace(Expl(i), "&", "&&")
    Next
    
    
'    For i = 1 To c
'        If InStr(1, word(i), " ") > 0 Then MsgBox word(i)
'    Next
    'MsgBox "词库单词数=" & C
    
    Randomize
    Rt = 0
    Wr = 0
    NewWord
    Label2.Caption = FN & ".txt 词库单词数=" & C
    LoadWords = True
    
    Exit Function
    
Errp:
    MsgBox "找不到词库文件，请将词库文件放入本程序所在的文件夹中"
    LoadWords = False
End Function

Private Sub Command1_Click()
    Label2.Caption = "上一个答案： " & Word(N) & " " & Expl(N)
    Wr = Wr + 1
    NewWord
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    MsgBox Pron(N)
End Sub

Private Sub DictOpen_Click()
    Dim FN1 As String, OldFN As String
    FN1 = InputBox("请输入词库文件名(不含.txt)", , "Words")
    If FN1 <> "" Then
        OldFN = FN
        FN = FN1
        If Not LoadWords Then FN = OldFN
    End If
End Sub

Private Sub DictReload_Click()
    LoadWords
End Sub

Private Sub DictView_Click()
    MsgBox "如果修改词库，修改后请重新加载词库才可生效"
    Shell "explorer " & """" & App.Path & "\" & FN & ".txt" & """"
End Sub

Private Sub FileExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FN = "vocabulary"
    If Not LoadWords Then Unload Me
End Sub

Function IsLetter(x As String) As Boolean
    IsLetter = (x >= "a" And x <= "z") Or (x >= "A" And x <= "Z")
End Function

Function IsWord(x As String) As Boolean
    IsWord = True
    Dim i As Integer, M As String
    For i = 1 To Len(x)
        M = Mid(x, i, 1)
        If Not (IsLetter(M) Or M = "'" Or M = "-") Then
            IsWord = False
            Exit Function
        End If
    Next
End Function

Sub NewWord()
    N = Int(Rnd * C) + 1
    Label1.Caption = Expl(N)
    Text1.Text = ""
End Sub

Private Sub HelpAbout_Click()
    MsgBox "负一的平方根 制作 2014年10月 QQ:903806024"
End Sub

Private Sub Label1_Click()
    MsgBox Label1.Caption & " " & N & "/" & C
End Sub

Private Sub Label2_Click()
    If Label2.Caption <> "" Then MsgBox Label2.Caption
End Sub

Private Sub Text1_Change()
    If LCase(Text1.Text) = LCase(Word(N)) Then
        Label2.Caption = "正确！"
        Rt = Rt + 1
        NewWord
    ElseIf Text1.Text <> "" Then
        Label2.Caption = ""
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then Command1_Click
    If KeyCode = 113 Then Command2_Click
End Sub

Private Sub ToolStat_Click()
    If Rt + Wr > 0 Then
        MsgBox "已背单词数=" & Wr + Rt & vbCrLf & "正确=" & Rt & vbCrLf & "错误=" & Wr & vbCrLf & "正确率=" & Int(Rt / (Rt + Wr) * 100000) / 1000 & "%"
    Else
        MsgBox "你还没有背单词"
    End If
End Sub
