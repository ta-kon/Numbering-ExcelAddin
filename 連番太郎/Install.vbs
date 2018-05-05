' 連番太郎インストールスクリプト
'
' 参考サイト
' ある SE のつぶやき
' VBScript で Excel にアドインを自動でインストール/アンインストールする方法
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'アドイン情報を設定
addInName = "連番太郎"
addInFileName = "連番太郎.xlam"

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

If (Not(objFileSys.FileExists(addInFileName))) Then
   MsgBox "インストールファイルが見つかりませんでした。Zipファイルを展開して実行してください。", vbExclamation, addInName
   WScript.Quit
End If

'インストール先パスの作成
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
strPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
installPath = strPath  & addInFileName

If MsgBox(addInName & " をインストールしますか？", vbYesNo + vbQuestion, addInName) = vbNo Then
  WScript.Quit
End If

'ファイルコピー(上書き)
objFileSys.CopyFile  addInFileName ,installPath , True
Set objFileSys = Nothing

'Excel インスタンス化
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'アドイン登録
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

'Excel 終了
objExcel.Quit
Set objAddin = Nothing
Set objExcel = Nothing

If (Err.Number = 0) Then
  MsgBox "アドインのインストールが終了しました。", vbInformation, addInName
Else
  MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName
  WScript.Quit
End If

Set objWshShell = Nothing
