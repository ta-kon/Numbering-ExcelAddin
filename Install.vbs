' �A�ԑ��Y�C���X�g�[���X�N���v�g
'
' �Q�l�T�C�g
' ���� SE �̂Ԃ₫
' VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'�A�h�C������ݒ�
addInName = "�A�ԑ��Y"
addInFileName = "�A�ԑ��Y.xlam"

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

If (Not(objFileSys.FileExists(addInFileName))) Then
   MsgBox "�C���X�g�[���t�@�C����������܂���ł����BZip�t�@�C����W�J���Ď��s���Ă��������B", vbExclamation, addInName
   WScript.Quit
End If

'�C���X�g�[����p�X�̍쐬
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
strPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
installPath = strPath  & addInFileName

If MsgBox(addInName & " ���C���X�g�[�����܂����H", vbYesNo + vbQuestion, addInName) = vbNo Then
  WScript.Quit
End If

'�t�@�C���R�s�[(�㏑��)
objFileSys.CopyFile  addInFileName ,installPath , True
Set objFileSys = Nothing

'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'�A�h�C���o�^
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

'Excel �I��
objExcel.Quit
Set objAddin = Nothing
Set objExcel = Nothing

If (Err.Number = 0) Then
  MsgBox "�A�h�C���̃C���X�g�[�����I�����܂����B", vbInformation, addInName
Else
  MsgBox "�G���[���������܂����B" & vbCrLF & "Excel���N�����Ă���ꍇ�͏I�����Ă��������B", vbExclamation, addInName
  WScript.Quit
End If

Set objWshShell = Nothing
