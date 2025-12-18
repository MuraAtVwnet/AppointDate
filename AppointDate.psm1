############################################################
# 指定日を曜日付きにする
############################################################
function Apo([string]$Date, [string]$time, [string]$ToTime, [switch]$VertionCheck){

	if( $VertionCheck ){
		$ModuleName = "AppointDate"
		$GitHubName = "MuraAtVwnet"

		$HomeDirectory = "~/"
		$Module = $ModuleName + ".psm1"
		$Installer = "Install" + $ModuleName + ".ps1"
		$Uninstaller = "Uninstall" + $ModuleName + ".ps1"
		$Vertion = "Vertion" + $ModuleName + ".txt"
		$GithubCommonURI = "https://raw.githubusercontent.com/$GitHubName/$ModuleName/refs/heads/master/"
		$VertionTemp = "VertionTemp" + $ModuleName + ".tmp"
		$VertionFilePath = Join-Path "~/" $Vertion
		$VertionTempFilePath = Join-Path "~/" $VertionTemp
		$VertionFileURI = $GithubCommonURI + "Vertion.txt"


		$Update = $False

		if( -not (Test-Path $VertionFilePath)){
			$Update = $True
		}
		else{
			$LocalVertion = Get-Content -Path $VertionFilePath

			$URI = $VertionFileURI
			$OutFile = $VertionTempFilePath
			Invoke-WebRequest -Uri $URI -OutFile $OutFile
			$NowVertion = Get-Content -Path $VertionTempFilePath
			Remove-Item $VertionTempFilePath

			if( $LocalVertion -ne $NowVertion ){
				$Update = $True
			}
		}

		if( $Update ){
			Write-Output "最新版に更新します"
			Write-Output "更新完了後、PowerShell プロンプトを開きなおしてください"

			$URI = $GithubCommonURI + $Module
			$ModuleFile = $HomeDirectory + $Module
			Invoke-WebRequest -Uri $URI -OutFile $ModuleFile

			$URI = $GithubCommonURI + "Install.ps1"
			$InstallerFile = $HomeDirectory + $Installer
			Invoke-WebRequest -Uri $URI -OutFile $InstallerFile

			$URI = $GithubCommonURI + "Uninstall.ps1"
			$OutFile = $HomeDirectory + $Uninstaller
			Invoke-WebRequest -Uri $URI -OutFile $OutFile

			$URI = $GithubCommonURI + "Vertion.txt"
			$OutFile = $HomeDirectory + $Vertion
			Invoke-WebRequest -Uri $URI -OutFile $OutFile

			& $InstallerFile

			Remove-Item $ModuleFile
			Remove-Item $InstallerFile

			Write-Output "更新完了"
			Write-Output "PowerShell プロンプトを開きなおしてください"
		}
		else{
			Write-Output "更新の必要はありません"
		}
		return
	}

	# 以下本来のコード

	$PointDate = $Date
	$PointTime = $time
	$PointToTime = $ToTime

	# 引数が無い
	if($Date -eq [string]$null){
		# 日付が省略されていので今日とする
		$PointDate = (Get-Date).ToString("yyyy/M/d ")
	}

	# 日付がセットされている
	if($Date.Contains("/")){
		$PointDate = $Date
		$PointTime = $time
		$PointToTime = $ToTime
	}

	# 時刻がセットされていたら今日とする
	elseif($Date.Contains(":") -or ($Date.Length -gt 2)){
		$PointDate = (Get-Date).ToString("yyyy/M/d ")
		$PointTime = $Date
		$PointToTime = $time
	}

	# 時刻の ":" が省略されている場合
	# From Time
	if( $PointTime -ne $null ){
		if( -not $PointTime.Contains(":") ){
			if( $PointTime.Length -eq 4 ){
				$PointTime = $PointTime.Substring(0, 2) + ":" + $PointTime.Substring(2, 2)
			}
			elseif( $PointTime.Length -eq 3 ){
				$PointTime = $PointTime.Substring(0, 1) + ":" + $PointTime.Substring(1, 2)
			}
		}
	}

	# To Time
	if( $PointToTime -ne $null ){
		if( -not $PointToTime.Contains(":") ){
			if( $PointToTime.Length -eq 4 ){
				$PointToTime = $PointToTime.Substring(0, 2) + ":" + $PointToTime.Substring(2, 2)
			}
			elseif( $PointToTime.Length -eq 3 ){
				$PointToTime = $PointToTime.Substring(0, 1) + ":" + $PointToTime.Substring(1, 2)
			}
		}
	}

	# 今日より前の日付がセットされている場合は来月にする
	if([String]$PointDate -Match "^[0-9]{1,2}$"){
		 # 今日より前の日付の場合来月にする
		$NowDay = (get-date).Day
		if( [int]$PointDate -lt $NowDay ){
			$PointDate =  ((Get-Date).AddMonths(1)).ToString("yyyy/M/") + $PointDate
		}
	}

	# 月日時刻(5/10 17:00) にするとエラーになる対策
	try{
		$DateTime = Get-Date $PointDate
	}
	catch{
		$NowMonth = ((Get-Date).Month).ToString()
		$NewDate = $NowMonth + "/" + $PointDate
		try{
			$DateTime = Get-Date $NewDate
		}
		catch{
			return "$Date $time $ToTime は日付として認識できません"
		}
	}

	# 年月だけ指定されていた場合、今日より前の月日の場合は来年にする
	$SlashCount = [regex]::Matches($PointDate, "/").Count
	if( $SlashCount -eq 1 ){
			$NowDateTime = Get-Date
			if( $NowDateTime -gt $DateTime ){
				$DateTime = $DateTime.AddYears(1)
			}
	}

	$strDateTime = ($DateTime).ToString("yyyy/M/d ") + $PointTime

	try{
		$DateTime = Get-Date $strDateTime
	}
	catch{
		return "$Date $time $ToTime は日付として認識できません"
	}

	# 終了時間チェック
	if( $PointToTime -ne [string]$null){
		try{
			$ToDateTime = Get-Date $PointToTime
		}
		catch{
			return "$Date $time $ToTime は日付として認識できません"
		}
	}

	if( $PointTime -eq [string]$null ){
		$TergetDay = $DateTime.ToString("yyyy年MM月dd日(ddd)")
	}
	elseif( $PointToTime -eq [string]$null ){
		$TergetDay = $DateTime.ToString("yyyy年MM月dd日(ddd) HH:mm ～")
	}
	else{
		$TergetDay = $DateTime.ToString("yyyy年MM月dd日(ddd) HH:mm ～ ") + $ToDateTime.ToString("HH:mm")
	}

	# クリップボードにコピー
	$TergetDay | Set-Clipboard

	return $TergetDay
}

################################################
# 現在時刻をクリップボードにセットする
################################################
function now(){
	$NowDateTime = (Get-Date).ToString("yyyy年MM月dd日(ddd) HH:mm")
	echo $NowDateTime
	$NowDateTime | Set-Clipboard
}

