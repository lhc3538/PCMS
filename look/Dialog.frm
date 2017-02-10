VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3870
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Textoldpass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   5160
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "原密码："
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "确认输入："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "密码："
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Dim a As String
Open App.Path & "\lookpassword.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, a
Loop
Close #1
Text3.Text = a
Dim l
l = Split(Text3.Text, ",")
Dim i As Integer
If Text3.Text <> "" Then
For i = 1 To UBound(l)
Text4.Text = Text4.Text & Chr(l(i) / 66)
Next i
End If
End Sub

Private Sub OKButton_Click()
If Textoldpass.Text = Text4.Text Then
If Text1.Text = Text2.Text Then
Dim b As String
b = Len(Text2.Text)
Dim i As Integer
For i = 1 To b
Dim a As String
a = Asc(Mid(Text2.Text, i, 1))
b = b & "," & a * 66
Next i
Text5.Text = b

Open App.Path & "\lookpassword.dat" For Output As #2
Print #2, Text5.Text
Close #2
MsgBox ("成功")
Else
MsgBox ("输入不一致")
End If
Else
MsgBox ("密码输入不正确")
End If
End Sub
