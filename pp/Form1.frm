VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = Format(Now, "ss")
Open "c:\SGxt\MMtext\pp.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, a
Loop
Close #2
Text3.Text = a

Name App.Path & "\" & "¡¦" & Text3.Text & "_pcms.exe" As App.Path & "\" & "¡¦" & Label1.Caption & "_pcms.exe"
Open "c:\SGxt\MMtext\pp.dat" For Output As #2
Print #2, Label1.Caption
Close #2
Text1.Text = "¡¦" & Label1.Caption & "_pcms.exe"
l = Shell("c:\sgxt\" & Text1.Text, vbNormalFocus)

End
End Sub
