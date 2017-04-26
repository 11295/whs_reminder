Option Explicit

Private Const DATA_ITEM_CNT = 3
Private Const DATA_IDX_FLG = 0
Private Const DATA_IDX_DATE = 1
Private Const DATA_IDX_TIME = 2
Private Const DATA_IDX_MSG = 3

Private Const EFFECT_FLG_ON = "1"
Private Const EFFECT_FLG_OFF = "0"

Private Const DAILY_TASK = "-"

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
		Dim tm
		Do While file.AtEndOfStream <> True
			' TAB区切りCSV形式のiniファイル
			' 行データ構造
			' 有効/無効フラグ	日付	時刻	表示内容
			lineData = file.ReadLine
			lineArray = GetLineArray(lineData)
			If Len(Trim(lineData)) = 0 Then
				' トリムして空行の場合、処理終了
				Exit Do
			Else
				If Ubound(lineArray) <> DATA_ITEM_CNT Then
					' 列数が定義と異なる場合、処理終了
					Exit Do
				End If
			End If

			If IsEffect(lineArray(DATA_IDX_FLG)) Then
				' フラグが有効の場合
				If IsTargetDay(lineArray(DATA_IDX_DATE)) Then
					' 今日または毎日の場合
					tm = IsAlertTime(lineArray(DATA_IDX_TIME))
					If len(tm) > 0 Then
						' まだ時間になっていない、かつ1分前、2分後以内の予定を表示
						Dim msg
						msg = tm & " " & Replace(lineArray(DATA_IDX_MSG), NEW_LINE, vbNewLine) & _
								vbNewLine & vbNewLine & "(データファイルを開く「はい」、終了「いいえ」)"

						If MsgBox(msg, vbYesNo + vbInformation + VbMsgBoxSetForeground, "予定") = vbYes Then
							' ボタンに応じてデータファイルをsakuraで開く
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

' 行データ取得
Function GetLineArray(ByVal lineData)
	GetLineArray = Split(lineData, vbTab)

	Dim col
	For Each col In GetLineArray
		col = Trim(col)
	Next
End Function

' 有効チェック
Function IsEffect(ByVal flg)
	IsEffect = (Len(flg) = 1 And flg = EFFECT_FLG_ON)
End Function

' 日付チェック
Function IsTargetDay(ByVal dt)
	IsTargetDay = True

	If dt = DAILY_TASK Then
		Exit Function
	End If

	If DateValue(dt) = Date Then
		Exit Function
	End If

	IsTargetDay = False
End Function

' 時刻チェック
Function IsAlertTime(ByVal tm)
	IsAlertTime = ""
	If Len(tm) = 0 Then
		Exit Function
	End If

	tm = TimeValue(tm)
	If Not(DateAdd("n", -1, Time()) <= tm _
			And tm =< DateAdd("n", 2, Time())) Then
		' 1分前、2分後以外の場合
		Exit Function
	End If

	IsAlertTime = tm
End Function
