VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日志查看器"
   ClientHeight    =   6255
   ClientLeft      =   7080
   ClientTop       =   5235
   ClientWidth     =   5115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":319A
   ScaleHeight     =   6255
   ScaleWidth      =   5115
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   4230
         ItemData        =   "Form1.frx":34B8E
         Left            =   240
         List            =   "Form1.frx":34B90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "备份"
         Height          =   255
         Left            =   3840
         MouseIcon       =   "Form1.frx":34B92
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   1680
         MouseIcon       =   "Form1.frx":34CE4
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":355AE
         Top             =   4680
         Width           =   1440
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   2040
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设置保存密码"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3480
      MouseIcon       =   "Form1.frx":3608A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":36954
      Top             =   5640
      Width           =   1140
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   600
      MouseIcon       =   "Form1.frx":370C9
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":37993
      Top             =   5640
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command2_Click()
Dim i
i = 0
Do While i < List1.ListCount
List1.Selected(i) = True
i = i + 1

Loop
End Sub



Private Sub Form_Load()
WebBrowser1.Navigate "http://www.2345.com/?3872"

Text1.Text = List1.ListCount
Open "C:\sgxt\looktext\gs.dat" For Input As #4
Do While Not EOF(4)
Line Input #4, gs
Loop
Close #4
Text1.Text = gs

Dim sj As String

Open "C:\sgxt\looktext\sj.dat" For Input As #1
Do While Not EOF(1)
   Input #1, sj
   Loop
Close #1
Text3.Text = sj

l = Split(Text3.Text, "'")
For i = 1 To UBound(l)
Text4.Text = l(i)
List1.List(i - 1) = Text4.Text
Next i
End Sub



Private Sub Image1_Click()
Dim i
i = 0
Do While i < List1.ListCount
List1.Selected(i) = True
i = i + 1

Loop
Dim a
a = 0
Do While a < List1.ListCount
If List1.Selected(a) = True Then
List1.RemoveItem a
Else
a = a + 1
End If
Loop


End Sub

Private Sub Image2_Click()
Form2.Show 1
End Sub

Private Sub Image3_Click()
End
End Sub

Private Sub Label1_Click()
If Dir(App.Path & "\lookpassword.dat") = "" Then
Open App.Path & "\lookpassword.dat" For Append As #1
Print , ""
Close #1
Label1.Visible = True

Else
Dialog.Show 1
End If


End Sub

Private Sub Label2_Click()
Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Set xlApp = CreateObject("Excel.Application")

'′使用模板
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\开机时间记录.xls")
    On Error GoTo 0
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = False
    xlSheet.Activate
    
    
    
    '′处理数据 , 填充Excel表
    For i = 1 To Text1.Text
    l = Split(List1.List(i - 1), ":")

     xlSheet.Cells(i + 1, 1) = l(0) '写入用户
     xlSheet.Cells(i + 1, 2) = l(1) '写入时间
    Next i
  '  ……

    
   ' ′//////////////
    

    
    xlApp.Visible = True '′显示表格
    Set xlApp = Nothing '′交还控制给Excel
    Set xlBoook = Nothing
    Set xlSheet = Nothing
End Sub

