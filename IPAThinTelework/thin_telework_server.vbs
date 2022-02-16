'=======================================
'メイン処理
'=======================================
Dim objShell
Dim procId

Set objShell = WScript.CreateObject("WScript.Shell")

'インストーラ起動し起動完了を待つ
Set objExec = objShell.Exec("installer.exe")
procId = objExec.ProcessID

'言語選択
inputKye("{ENTER}")

''別exeが起動するので、プロセスIDを変更
procId = changeTarget("thinsetup.exe")

'はじめに
inputKye("{ENTER}")

'選択
inputKye("{ENTER}")

'使用条件
inputKye("{TAB}")
inputKye("{TAB}")
inputKye(" ")
inputKye("{ENTER}")

'重要なお知らせ
inputKye("{ENTER}")

'インストール先ディレクトリ
inputKye("{ENTER}")

' インストール準備の完了
inputKye("{ENTER}")

' 完了
inputKye("{TAB}")
inputKye(" ")
inputKye("{TAB}")
inputKye("{ENTER}")

'ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
'=======================================
'キーボードの入力を送信する
'=======================================
Sub inputKye(strKey)
	focusTarget()
	objShell.SendKeys (strKey)
	WScript.Sleep 1000
End Sub

'=======================================
'対象のウィンドウにフォーカスを当てる
'=======================================
Sub focusTarget()
    Do Until objShell.AppActivate(procId)
        WScript.Sleep 1000
		
		If not processExists() Then
			WScript.Quit
		End if
		
    Loop
End Sub

'=======================================
'派生した別のexeのプロセスを検索する。
'=======================================
Function changeTarget(pname)
	Dim svcs
	Dim procList
	
	Set svcs = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
	Set procList = svcs.ExecQuery("Select * From Win32_Process  Where (Caption = '"& pname &"'  and   ParentProcessId = '" &  procId & "') " )
	
	Do Until procList.Count > 0
		Set svcs = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
		Set procList = svcs.ExecQuery("Select * From Win32_Process  Where (Caption = '"& pname &"'  and   ParentProcessId = '" &  procId & "') " )
		WScript.Sleep 1000
	Loop
	
	changeTarget = ""
	For Each proc In procList
		changeTarget = proc.ProcessId 
	Next
	
	Set svcs = Nothing
	Set procList = Nothing
End Function

'=======================================
'プロセスの死活チェック
'=======================================
Function processExists()
	Dim svcs
	Dim procList
	
	Set svcs = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
	Set procList = svcs.ExecQuery("Select * From Win32_Process  Where (ProcessId = '" &  procId & "') " )
	
	processExists = procList.Count > 0
	
	Set svcs = Nothing
	Set procList = Nothing
End Function






