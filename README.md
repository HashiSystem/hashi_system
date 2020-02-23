# kintai-slack-gas

勤怠管理をslackから行います。
gas(Google Apps Script)によってGoogleドライブに勤怠記録が登録されます。

(README作成中・・）

# Features

SlackからSlash Commandを入力し、ダイアログに出勤退勤時刻などを入力することができます。
また、xlsx形式で月報の出力を行うことができます。


# Requirement

管理用のGoogleアカウントがひとつ必要になります。


# Installation


# Usage

・ユーザーの管理
	勤怠スプレッドシートの「user」シートに登録
		(項目)
		#					番号です。PGで使用してません。
		名前				ユーザーの名前です。勤怠シートがこの名前で作成されます。
		slackユーザー名		slackに登録したユーザー名です。
		休憩時間			勤務時間（勤務終了時刻 - 勤務開始時刻）から引かれます。
		グループ			ユーザーの所属するグループです。空でもOK
		参照権限			月報の参照権限です。ALL, グループ名, selfを設定。

・Slashコマンドの入力
	「/kintai」コマンドを入力すると、登録してあるユーザーに
	「出退勤の登録」「月報の出力」ボタンが表示されます。


# Note
 
	・単純な勤務時間計算しかしてません（勤務終了時刻 - 勤務開始時刻 - 休憩時間）
	・入力値のvalidationだったり整合性チェックだったりは未実装
	・slackのタイムアウト
		最初に勤怠登録すると出がち。処理自体はうまくいってるはず・・
		勤務データが増えると心配。
	・ところどころ訳のわからない処理してます。（適当）
 
# Author

* はしシステム
* slack.hashiken@gmail.com
 
# License
 
"kintai-slack-gas" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).


"kintai-slack-gas" is Confidential.
