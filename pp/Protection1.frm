VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Protection1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "c:\SGxt\MMtext\pp.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, a
Loop
Close #2
Text1.Text = a
Open "c:\SGxt\MMtext\pp1.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, a
Loop
Close #2
Text2.Text = a

Name App.Path & "\" & "¡¦" & Text1.Text & "_pcms.exe" As App.Path & "\" & "¡¦" & Text2.Text & "_pcms.exe"
Kill ("c:\SGxt\MMtext\pp1.dat")
Open "c:\SGxt\MMtext\pp.dat" For Output As #2
Print #2, Text2.Text
Close #2
l = Shell(App.Path & "\Protection2.exe")
End

End Sub
