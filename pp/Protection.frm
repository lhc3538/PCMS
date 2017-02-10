VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4365
   Icon            =   "Protection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4365
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = Format(Now, "ss")
Open "c:\SGxt\MMtext\pp1.dat" For Append As #2
Close #2

Open "c:\SGxt\MMtext\pp1.dat" For Output As #2
Print #2, Label1.Caption
Close #2
a = Shell(App.Path & "\Protection1.exe")
End

End Sub
