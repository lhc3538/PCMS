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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
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
'获取用户名
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub Form_Load()

'获取用户名

 Dim st     As String * 100
          Dim pln     As Long
          pln = 99
          GetUserName st, pln
          Dim user As String
        user = Left$(st, pln)
Label3.Caption = ":" & Format(Date, "yyyy年mm月dd日") & Format(Now, "hh点mm分ss秒")

gs = gs + 1
Open "c:\SGxt\looktext\gs.dat" For Output As #4
Print #4, gs
Close #4


Dim sj As String

Open "c:\SGxt\looktext\sj.dat" For Input As #2
Do While Not EOF(2)
Line Input #2, sj
Loop
Close #2
sj = sj & "'" & user & Label3.Caption
Open "c:\SGxt\looktext\sj.dat" For Output As #4
Print #4, sj
Close #4
End

End Sub
