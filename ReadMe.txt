入力した日付を曜日付きの年月日に変換する PowerShell スクリプト

■ これは何?
入力した日時を曜日付きの年月日にしてクリップボードに送ります。

アポを取る時とかで曜日を間違って恥ずかしい思いをしなくて済みます。

Web サービスで同様のものがあったのですが、年がセットされないので PowerShell で書いてみました。

■ 使い方
・年月日を曜日付きにする
PowerShell プロンプトで「apo 日付」を入力します。
年を省略すると今年になります。

    PS C:\> apo 10/13
    2019年10月13日(日)

この状態で曜日付き日付がクリップボードにセットされているので、そのままペーストします。

・年月日 + 時間を曜日付きにする
PowerShell プロンプトで「apo 日付 時刻」を入力します

    PS C:\> apo 2020/5/18 13:00
    2020年5月18日(月) 13:00 ～

時刻の 「:」は省略可能です

    PS C:\> apo 2020/5/18 1300
    2020年5月18日(月) 13:00 ～


こちらもクリップボードにセットされているので、そのままペーストします。

年月を省略すると、今月と判断します。

    PS C:\> apo
    2019年7月21日(日)

    PS C:\> apo 10:00
    2019年7月21日(日) 10:00 ～

年月日を省略すると、今日と判断します。

■ スクリプトインストール方法
--- 以下を PowerShell プロンプトにコピペ ---

$ModuleName = "AppointDate"
$GitHubName = "MuraAtVwnet"
$URI = "https://raw.githubusercontent.com/$GitHubName/$ModuleName/refs/heads/main/OnlineInstall.ps1"
$OutFile = "~/OnlineInstall.ps1"
Invoke-WebRequest -Uri $URI -OutFile $OutFile
& $OutFile


■ オプション

apo の後に、ハイフンを入力して TAB を叩くと、オプションが補完されるのて、必要オプションを選択してください

apo -[TAB]

-VertionCheck

最新版のスクリプトがあるか確認します
最新版があれば、自動ダウンロード & 更新します


■ スクリプトアンインストール方法
--- 以下を PowerShell プロンプトにコピペ ---

~/UninstallAppointDate.ps1


■ 動作確認環境
Windows PowerShell 5.1
PowerShell 7.4.5 (Windows)

■ Web サイト

入力した日付を曜日付きの年月日に変換する PowerShell スクリプト
http://www.vwnet.jp/Windows/PowerShell/2019021501/AppointDate.htm

■ リポジトリ
GitHub でも公開しているので、最新版が欲しい方はこちらから Clone してください。

https://github.com/MuraAtVwnet/AppointDate
git@github.com:MuraAtVwnet/AppointDate.git

■ おまけ
now と入力すると現在の日時をクリップボードにセットします

