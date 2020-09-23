VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ARABIC TEST"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "IF A=1 THEN MSGBOX ""DO YOU WANT TO QUIT"" IN ARABIC LANGUAGE"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CHECK TEXT1 FOR PASSWORD IN ARABIC LANGUAGE"
      Height          =   1095
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ARABIC LANGUAGE"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENGLISH LANGUAGE"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "TEST3"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "TEST2"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "TEST1"
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Option Compare Text  ' FRO NOT CASE SENSETIVE IN PASSWORD
' begine
'this code example help arabic programmer to write code and call function in arabic language
' end
Dim ãÊÛíÑ_äÕí As String
Const ss_english As String = "DO YOU WANT TO QUIT ?"
Const ss_arabic As String = "?åá ÊÑÛÈ ÈÇáÎÑæÌ "
Private Sub Command1_Click()
' call the func QUIT in english language

QUIT "DO YOU WANT TO QUIT"

End Sub
Private Sub Command2_Click()
' call the func QUIT in ARABIC language
' QUIT IN ARABIC MEANS ÇÎÑÌ
ÇÎÑÌ "åá ÊÑÛÈ ÈÇáÎÑæÌ"


End Sub


Function QUIT(Optional ByVal action As String = ss_english) As String()
If MsgBox(ss_english, vbYesNo) = vbYes Then
Unload Me
Set Form1 = Nothing
End
End If
End Function

Function ÇÎÑÌ(Optional ByVal action As String = ss_arabic) As String
If MsgBox(ss_arabic, vbYesNo + vbMsgBoxRight) = vbYes Then
Unload Me
Set Form1 = Nothing
End
End If

End Function

Private Sub Command3_Click()
' THIS ARTABIC STETEMNT MEANS
' CALL SUB CHECK MY PASSWORD

ÊÍÞÞ_ãä_ßáãÉ_ÇáÓÑ

End Sub

Sub ÊÍÞÞ_ãä_ßáãÉ_ÇáÓÑ() ' SUB CHECK_PASSWORD_KEYWORD
If Text1.Text = "MUAD" Then ÇÎÑÌ
End Sub





' REMARK
'TO USE IF STATEMENT
' OR USE DO WHILE
' OR USE ANY OTHER
'IT IS VERY EASY JUST CREATE FUNCTION AND CALL IT WITH ARABIC NAME


' FINISH
' I THINK VB IS THE MOST FLEXABLE LANGUAGE IN THE WORLD







Private Sub Command4_Click()
'DIM A AS STRING
'A=1
'IF A="1" THEN END
ÚÑÝ ãÊÛíÑ_äÕí
ÚæÖ_ÞíãÉ "1"
Ã = "1"

ÇÐÇ "Ã", "1"
If A = 1 Then QUIT

End Sub

'DIM A AS STRING
Function ÚÑÝ(ÇáãÊÛíÑ As String)
If ÇáãÊÛíÑ = "ãÊÛÑ_äÕí" Then
Dim Ã  As String
End If
End Function

' ASSIGN "1" TOI A VARIABLE
Function ÚæÖ_ÞíãÉ(ÇáÞíãÉ As String)

Ã = ÇáÞíãÉ

End Function

'FUNCTION IF THEN DO


'FUNCTION IF THEN DO

Function ÇÐÇ(ÇáãÊÛíÑ As String, ÇáÞíãÉ As String) As String
' IF THE VARIBLE NAME IS A AND IT IS VALUE="1" THEN THE FUNCTION=VALUE
Print ÇáãÊÛíÑ
Print ÇáÞíãÉ
If ÇáãÊÛíÑ = "Ã" And ÇáÞíãÉ = "1" Then ÇÎÑÌ
End Function







