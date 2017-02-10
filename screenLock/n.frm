VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "重设"
   ClientHeight    =   1395
   ClientLeft      =   5460
   ClientTop       =   2655
   ClientWidth     =   5055
   Icon            =   "n.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MouseIcon       =   "n.frx":6912
   Picture         =   "n.frx":71DC
   ScaleHeight     =   1395
   ScaleWidth      =   5055
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "请输入高级密码："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "13963858419" Then
e = MsgBox("高级密码正确。是否继续？", , "确认")



Form2.Show 1
Form3.Visible = False
Form3.Visible = False
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

If Text1.Text = "13963858419" Then
e = MsgBox("高级密码正确。是否继续？", , "确认")


Form2.Show 1
Form3.Visible = False

Else
MsgBox ("密码错误，请重新输入")
Text1.Text = ""

End If
End If

End Sub
