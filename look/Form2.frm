VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "验证"
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4080
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4080
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Default         =   -1  'True
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   1560
      MouseIcon       =   "Form2.frx":219D2
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":2229C
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "输入密码："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text4.Text Then
gs = 1

Open "c:\sgxt\looktext\gs.dat" For Output As #3
Print #3, gs
Close #3
Open "c:\sgxt\looktext\sj.dat" For Output As #3
Print #3, ""

Close #3
X = MsgBox("保存成功！", 64, 保存)
End



Else
msg = MsgBox("密码错误", vbOKOnly, 错误)
Text1.Text = ""
End If
End Sub

Private Sub Form_Load()
Dim a As String
Open App.Path & "\lookpassword.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, a
Loop
Close #1
Text3.Text = a

l = Split(Text3.Text, ",")
For i = 1 To UBound(l)
Text4.Text = Text4.Text & Chr(l(i) / 66)
Next i


End Sub

Private Sub Image1_Click()
If Text1.Text = Text4.Text Then
gs = 1

Open "c:\sgxt\looktext\gs.dat" For Output As #3
Print #3, gs
Close #3
Open "c:\sgxt\looktext\sj.dat" For Output As #3
Print #3, ""

Close #3
X = MsgBox("保存成功！", 64, 保存)
End



Else
msg = MsgBox("密码错误", vbOKOnly, 错误)
Text1.Text = ""
End If
End Sub
