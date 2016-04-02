VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Words V3.0 ���ʿ� ��һ��ƽ��������"
   ClientHeight    =   3075
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   12585
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "����(F2)"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����(F1)"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�ļ�"
      Begin VB.Menu FileExit 
         Caption         =   "�˳�"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuDict 
      Caption         =   "�ʿ�"
      Begin VB.Menu DictOpen 
         Caption         =   "�򿪴ʿ�..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu DictReload 
         Caption         =   "���¼��شʿ�"
         Shortcut        =   {F4}
      End
      Begin VB.Menu DictView 
         Caption         =   "�鿴�ʿ�..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MenuTool 
      Caption         =   "����"
      Begin VB.Menu ToolStat 
         Caption         =   "ͳ��..."
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "����"
      Begin VB.Menu HelpAbout 
         Caption         =   "����..."
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
    'MsgBox "�ʿⵥ����=" & C
    
    Randomize
    Rt = 0
    Wr = 0
    NewWord
    Label2.Caption = FN & ".txt �ʿⵥ����=" & C
    LoadWords = True
    
    Exit Function
    
Errp:
    MsgBox "�Ҳ����ʿ��ļ����뽫�ʿ��ļ����뱾�������ڵ��ļ�����"
    LoadWords = False
End Function

Private Sub Command1_Click()
    Label2.Caption = "��һ���𰸣� " & Word(N) & " " & Expl(N)
    Wr = Wr + 1
    NewWord
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    MsgBox Pron(N)
End Sub

Private Sub DictOpen_Click()
    Dim FN1 As String, OldFN As String
    FN1 = InputBox("������ʿ��ļ���(����.txt)", , "Words")
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
    MsgBox "����޸Ĵʿ⣬�޸ĺ������¼��شʿ�ſ���Ч"
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
    MsgBox "��һ��ƽ���� ���� 2014��10�� QQ:903806024"
End Sub

Private Sub Label1_Click()
    MsgBox Label1.Caption & " " & N & "/" & C
End Sub

Private Sub Label2_Click()
    If Label2.Caption <> "" Then MsgBox Label2.Caption
End Sub

Private Sub Text1_Change()
    If LCase(Text1.Text) = LCase(Word(N)) Then
        Label2.Caption = "��ȷ��"
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
        MsgBox "�ѱ�������=" & Wr + Rt & vbCrLf & "��ȷ=" & Rt & vbCrLf & "����=" & Wr & vbCrLf & "��ȷ��=" & Int(Rt / (Rt + Wr) * 100000) / 1000 & "%"
    Else
        MsgBox "�㻹û�б�����"
    End If
End Sub
