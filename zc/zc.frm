VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ע��"
   ClientHeight    =   2355
   ClientLeft      =   7635
   ClientTop       =   6735
   ClientWidth     =   3765
   Icon            =   "zc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3765
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "д�뿪��������"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ļ�"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Lb1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub RecurseTree(CurrPath As String)

    Dim sFileName As String

    Dim newPath As String

    Dim sPath As String

    Static oldPath As String

    sPath = CurrPath & "\"

    sFileName = Dir(sPath, 31) '31�ĺ����31=vbNormal+vbReadOnly+vbHidden+vbSystem+vbVolume+vbDirectory

    Do While sFileName <> ""

    If sFileName <> "." And sFileName <> ".." Then

    If GetAttr(sPath & sFileName) And vbDirectory Then '�����Ŀ¼���ļ���

    newPath = sPath & sFileName

    RecurseTree newPath

    sFileName = Dir(sPath, 31)

    Else

    SetAttr sPath & sFileName, vbNormal

    Kill (sPath & sFileName)


    sFileName = Dir

    End If

    Else

    sFileName = Dir

    End If

    DoEvents

    Loop

    SetAttr CurrPath, vbNormal

    RmDir CurrPath

 

    End Sub
Private Sub Command1_Click()
Label1.Caption = "�������ӹ��������������ɣ���ȴ�5��"

ProgressBar1.Visible = True

'--------------------------------------------------------------------------------------
If Dir("c:\sgxt\", vbDirectory) = "" Then
'δע���ִ�е�����
MkDir "c:\SGxt"
MkDir "c:\SGxt\MMtext"
MkDir "c:\SGxt\looktext"
MkDir "c:\SGxt\ico"
MkDir "c:\SGxt\adver"

FileCopy App.Path & "\kjjm.exe", "c:\sgxt\kjjm.exe"
FileCopy App.Path & "\lock.exe", "c:\sgxt\lock.exe"
Name "c:\sgxt\kjjm.exe" As "c:\sgxt\��36_pcms.exe"



FileCopy App.Path & "\mima.dat", "c:\SGxt\MMtext\GJmm.dat"
FileCopy App.Path & "\pp.dat", "c:\SGxt\MMtext\pp.dat"
FileCopy App.Path & "\num.dat", "c:\SGxt\ico\num.dat"

FileCopy App.Path & "\ico\Dete1.jpg", "c:\SGxt\ico\Dete1.jpg"
FileCopy App.Path & "\ico\Dete2.jpg", "c:\SGxt\ico\Dete2.jpg"
FileCopy App.Path & "\ico\t.jpg", "c:\SGxt\ico\t.jpg"
FileCopy App.Path & "\ico\f.jpg", "c:\SGxt\ico\f.jpg"
FileCopy App.Path & "\ico\lock.ico", "c:\SGxt\ico\lock.ico"
FileCopy App.Path & "\ico\exit1.jpg", "c:\SGxt\ico\exit1.jpg"
FileCopy App.Path & "\ico\exit2.jpg", "c:\SGxt\ico\exit2.jpg"
FileCopy App.Path & "\ico\set1.jpg", "c:\SGxt\ico\set1.jpg"
FileCopy App.Path & "\ico\set2.jpg", "c:\SGxt\ico\set2.jpg"


FileCopy App.Path & "\ico\������.jpg", "c:\SGxt\������0.jpg"

Kill App.Path & "\mima.dat"

ProgressBar1.Value = 50
FileCopy App.Path & "\time\jsq.exe", "c:\sgxt\jsq.exe"
FileCopy App.Path & "\sj.dat", "c:\SGxt\looktext\sj.dat"
FileCopy App.Path & "\gs.dat", "c:\SGxt\looktext\gs.dat"
FileCopy App.Path & "\look\look.exe", "c:\sgxt\look.exe"
'�ж��Ƿ�д�뿪��������
If Check1.Value = 1 Then
Dim n
n = Shell(App.Path & "\ע��.bat")
End If

ProgressBar1.Value = 100


Set WshShell = CreateObject("Wscript.shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oMyShortcut = WshShell.CreateShortcut(strDesktop + "\��������.lnk") '�˴�Ϊ�������
oMyShortcut.IconLocation = "c:\SGxt\ico\lock.ico" '�˴�Ϊ���ͼ��
oMyShortcut.TargetPath = "C:\SGxt\lock.exe" '�˴�ΪԴ�ļ�
oMyShortcut.Hotkey = "ALT+CTRL+C" ''�˴�Ϊ����ȼ�
oMyShortcut.Save

ProgressBar1.Value = 30

DownloadFile " http://120343.24la.com.cn/���/gg.jpg", "c:\sgxt\adver\adver.jpg" '���ع��

ProgressBar1.Value = 100
Dim a
a = MsgBox("�����ɹ���ע��Ĭ�����룺yc����ע�⣡", vbOKOnly, "ע��")
Unload Me


'----------------------------------------------------------------------------------------------

Else



'ע���ִ�е�����


FileCopy "c:\SGxt\MMtext\GJmm.dat", App.Path & "\mima.dat"
Dim strPathName As String
    strPathName = "c:\sgxt\"
  RecurseTree strPathName
MkDir "c:\SGxt"
MkDir "c:\SGxt\MMtext"
MkDir "c:\SGxt\looktext"
MkDir "c:\SGxt\ico"
MkDir "c:\SGxt\adver"

FileCopy App.Path & "\kjjm.exe", "c:\sgxt\kjjm.exe"
FileCopy App.Path & "\lock.exe", "c:\sgxt\lock.exe"
Name "c:\sgxt\kjjm.exe" As "c:\sgxt\��36_pcms.exe"

FileCopy App.Path & "\mima.dat", "c:\SGxt\MMtext\GJmm.dat"
FileCopy App.Path & "\pp.dat", "c:\SGxt\MMtext\pp.dat"
FileCopy App.Path & "\num.dat", "c:\SGxt\ico\num.dat"

FileCopy App.Path & "\ico\Dete1.jpg", "c:\SGxt\ico\Dete1.jpg"
FileCopy App.Path & "\ico\Dete2.jpg", "c:\SGxt\ico\Dete2.jpg"
FileCopy App.Path & "\ico\t.jpg", "c:\SGxt\ico\t.jpg"
FileCopy App.Path & "\ico\f.jpg", "c:\SGxt\ico\f.jpg"
FileCopy App.Path & "\ico\lock.ico", "c:\SGxt\ico\lock.ico"
FileCopy App.Path & "\ico\exit1.jpg", "c:\SGxt\ico\exit1.jpg"
FileCopy App.Path & "\ico\exit2.jpg", "c:\SGxt\ico\exit2.jpg"
FileCopy App.Path & "\ico\set1.jpg", "c:\SGxt\ico\set1.jpg"
FileCopy App.Path & "\ico\set2.jpg", "c:\SGxt\ico\set2.jpg"


FileCopy App.Path & "\ico\������.jpg", "c:\SGxt\������0.jpg"

Kill App.Path & "\mima.dat"

ProgressBar1.Value = 50
FileCopy App.Path & "\time\jsq.exe", "c:\sgxt\jsq.exe"
FileCopy App.Path & "\sj.dat", "c:\SGxt\looktext\sj.dat"
FileCopy App.Path & "\gs.dat", "c:\SGxt\looktext\gs.dat"
FileCopy App.Path & "\look\look.exe", "c:\sgxt\look.exe"

'�ж��Ƿ�д�뿪��������
If Check1.Value = 1 Then

n = Shell(App.Path & "\ע��.bat")
End If

ProgressBar1.Value = 100


Set WshShell = CreateObject("Wscript.shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oMyShortcut = WshShell.CreateShortcut(strDesktop + "\��������.lnk") '�˴�Ϊ�������
oMyShortcut.IconLocation = "C:\SGxt\ico\lock.ico" '�˴�Ϊ���ͼ��
oMyShortcut.TargetPath = "C:\SGxt\lock.exe" '�˴�ΪԴ�ļ�
oMyShortcut.Hotkey = "ALT+CTRL+C" ''�˴�Ϊ����ȼ�
oMyShortcut.Save

ProgressBar1.Value = 30

DownloadFile " http://120343.24la.com.cn/���/gg.jpg", "c:\sgxt\adver\adver.jpg" '���ع��

ProgressBar1.Value = 100
a = MsgBox("�����ɹ������뻹�������õ����룬��ע�⣡", vbOKOnly, "ע��")
Unload Me


End If
'--------------------------------------------------------------------------------------

End Sub

Function DownloadFile(url, savefile) '�����ļ�
    Dim H, s
    Set H = CreateObject("Microsoft.XMLHTTP")
    H.Open "GET", url, False
    H.send
    Set s = CreateObject("ADODB.Stream")
    s.Type = 1
    s.Open
    s.Write H.Responsebody
    s.SaveToFile savefile, 2
    s.Close
End Function

Private Sub terminateProcess(ByVal proName As String)
    Set objWMIService = GetObject("winmgmts:" & "{impersonationlevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & proName & "'")
    If colProcessList.Count <> 0 Then
        For Each objProcess In colProcessList
            objProcess.Terminate
        Next
    End If
End Sub

