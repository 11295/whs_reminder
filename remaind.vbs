Option Explicit

Private Const DATA_ITEM_CNT = 3
Private Const DATA_IDX_FLG = 0
Private Const DATA_IDX_DATE = 1
Private Const DATA_IDX_TIME = 2
Private Const DATA_IDX_MSG = 3

Private Const EFFECT_FLG_ON = "1"
Private Const EFFECT_FLG_OFF = "0"

Private Const DATE_FORMAT = "yyyy/mm/dd"
Private Const TIME_FORMAT = "hh:nn"

Private Const NEW_LINE = "@"

' list.iniを開くソフトのフルパス
Private Const APP_PATH = """C:\Program Files\sakura\sakura.exe"""

Call Checker

Sub Checker() 
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim path
	path = fso.getParentFolderName(WScript.ScriptFullName) & "\list.ini"

	' このファイルと同フォルダ内の「list.ini」をオープン
	' 予定データは昇順ソート済前提
	Dim file
	Set file = fso.OpenTextFile(path, 1)
	If Err.Number = 0 Then
		Dim lineData		
    Dim lineArray
		Dim flg
		Dim dt
		Dim tm
		Do While file.AtEndOfStream <> True
			' TAB区切りCSV形式のiniファイル
			' 行データ構造
			' 有効/無効フラグ	日付	時刻	表示内容
			lineData = file.ReadLine
			lineArray = Split(lineData, vbTab)
			If Len(Trim(lineData)) = 0 Or Ubound(lineArray) <> DATA_ITEM_CNT Then
				' トリムして空行または列数が定義と異なる場合、処理終了
				Exit Do
			End If

			If len(lineArray(DATA_IDX_FLG)) = 1 _
				And lineArray(DATA_IDX_FLG) = EFFECT_FLG_ON Then
				' フラグが有効の場合
				dt = DateValue(lineArray(DATA_IDX_DATE))
				If dt = Date Then
					' 今日の場合
					tm = TimeValue(lineArray(DATA_IDX_TIME))
					If Time() <= tm And tm =< DateAdd("n", 3, Time()) Then
						' まだ時間になっていない、かつ3分後以内の予定を表示
						Dim msg
						msg = tm & " " & Replace(lineArray(DATA_IDX_MSG), NEW_LINE, vbNewLine) & vbNewLine & vbNewLine & "(データファイルを開く「はい」、終了「いいえ」)"
						Dim btnRet
						btnRet = MsgBox(msg, vbYesNo + vbInformation + VbMsgBoxSetForeground, "予定")
						
						' ボタンに応じてデータファイルをsakuraで開く
						If btnRet = vbYes Then
							Dim objWShell
							Set objWShell = CreateObject("WScript.Shell")
							objWShell.Run APP_PATH & " -- " & path
							Set objWShell = Nothing
						End If
						' 該当する予定をループ終端まで全部表示する
						' Exit Do
					End If
				End If
			End If
		Loop
		file.Close
	Else
		Exit Sub
	End If

	Set file = Nothing
	Set fso = Nothing
End Sub


