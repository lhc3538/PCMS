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
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
'���ϵͳ���Ƿ�ȱ�ٿؼ�
If Dir("C:\Windows\System32\TABCTL32.OCX") = "" Then
'ȱ��
FileCopy App.Path & "\TABCTL32.OCX", "C:\Windows\System32\TABCTL32.OCX"


End If
'-----------------------------------------------------------------------
If Dir("C:\Windows\System32\COMCTL32.OCX") = "" Then
'ȱ��
FileCopy App.Path & "\COMCTL32.OCX", "C:\Windows\System32\COMCTL32.OCX"

End If
'---------------------------------------------------------------------------
If Dir("C:\Windows\System32\COMDLG32.OCX") = "" Then
'ȱ��
FileCopy App.Path & "\COMDLG32.OCX", "C:\Windows\System32\COMDLG32.OCX"

End If

'-------------------------------------------------------------------


s = Shell(App.Path & "\zc.exe", vbNormalFocus)
End
End Sub
