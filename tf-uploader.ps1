# Для основной загрузки создаём задание в планировщике UploadToNAS, Every 15 min, from 0:05 for 1435 minutes every day
# powershell.exe -ExecutionPolicy Bypass -File D:\Work\TK700\Scripts\UploadToNAS\UploadToNAS.ps1
# 												свой ТК	^
# Для загрузки AED создаем шаг в Job-е, запускаем после успешной упаковки.
param
(
	[Parameter(Mandatory = $false)]
    $transferMode,
	[Parameter(Mandatory = $false)]
    $dbPrefix = ".*",
	[Parameter(Mandatory = $false)]
	$fileExt = "7z"
)

$version = "1.18"
[System.Console]::Title = "UploadToNAS [v. $version]"

# Указываем свой ТК.
$tk = 724

# Указываем свой ТК.
$email = ""

# Пропустить проверки рабочего времени для основного цикла.
$skipScheldule = 0
# Начало рабочего дня.
$workTimeStart = Get-Date -Hour 7 -Minute 30 -Second 0
# Конец рабочего дня.
$workTimeEnd = Get-Date -Hour 17 -Minute 30 -Second 0
# Конец рабочего дня (выходной).
$workTimeEndSaturday = Get-Date -Hour 16 -Minute 30 -Second 0
# Путь с рабочими переносами.
$workingDaysFilePath = "\WorkingDays.txt"
$workingDays = @()
# Таймер для основного цикла (миллисекунды).
$loopSleepTimer = 36000

# День недели для загрузки отложенной группы.
$delayedStartDay = 7
# Время начала загрузки отложенной группы.
$delayedStartTimeStart = Get-Date -Hour 0 -Minute 0 -Second 0
# Отложенная группа.
$schelduleGroupForDelayedStart = @(702, 703, 705, 706, 708, 709, 712, 713, 718, 719, 724)

$7zipPath = "D:\Work\Utils\7-Zip\7z.exe"
$logDir = "E:\Backup\Logs\UploadToNAS"
$workTime

# Исходный каталог для обычной передачи.
$sourceDir  = ""
# Исходный каталог для передачи AED.
$sourceDirForAed  = ""
# Каталог NAS для обычной передачи.
$destDirBase = ""
# Каталог NAS для передачи AED.
$destDirForAed = "$destDirBase\AED"
$destDir = "$destDirBase"

$currentDateStringForLog = Get-Date -Format "MM-yyyy"
$copyLogPath = "$logDir\UploadToNAS_$currentDateStringForLog.log"
$errorLogPath = "$logDir\UploadToNAS_Error_$currentDateStringForLog.log"
$mainAedLogPath = "$logDir\UploadToNASAedMain_$currentDateStringForLog.log"
$copyLogAedPath = "$logDir\UploadToNASAed_$currentDateStringForLog.log"
$errorLogAedPath = "$logDir\UploadToNASAed_Error_$currentDateStringForLog.log"

# Доступная ширина канала.
$bandwidthAvailableForLimitedAedTransfer = 5120
# Необходимая ширина канала.
$bandwidthDesiredForLimitedAedTransfer = 2048
$ipgNetworkGap = [math]::Round(($bandwidthAvailableForLimitedAedTransfer - $bandwidthDesiredForLimitedAedTransfer) / ($bandwidthAvailableForLimitedAedTransfer * $bandwidthDesiredForLimitedAedTransfer) * 512 * 1000)

$maxRetries = 100
$maxRetriesForAED = 432000
$waitTimeInSec = 1
$robocopyOptions = @("/MOV", "/COPY:DAT", "/IS", "/Z", "/B", "/R:$maxRetries", "/W:$waitTimeInSec", "/NP", "/NC", "/REG", "/TEE", "/LOG+:$copyLogPath", "/TS")
$robocopyOptionsForAED = @("/E", "/COPY:DAT", "/Z", "/R:$maxRetriesForAED", "/W:$waitTimeInSec", "/NC", "/REG", "/TEE", "/LOG+:$copyLogAedPath", "/TS")
$robocopyOptionsForAEDLimited = @("/E", "/COPY:DAT", "/Z", "/R:$maxRetriesForAED", "/W:$waitTimeInSec", "/NC", "/REG", "/TEE", "/LOG+:$copyLogAedPath", "/TS", "/IPG:$ipgNetworkGap")

$regexPattern = "^.*?($dbPrefix)(_db_|_tlog_).*(\.$fileExt)$"
$regexPatternForFirstArchive = "^.*?($dbPrefix)(_db_|_tlog_).*(\.BAK.$fileExt.001)$"
$regexMatchOptions = [System.Text.RegularExpressions.RegexOptions] 'IgnoreCase'

$smtpServer = ""
$smtpPort = 25
$mailFrom = $env:computername + ""

Function SendToEmail([string] $from, [string] $to, [string] $subject, [string] $body)
{
	Write-Host "SendToEmail -> $to"
	
    $message = new-object Net.Mail.MailMessage;
    $message.From = $from;
    $message.To.Add($to);
    $message.Subject = $subject;
    $message.Body = $body;
    $smtp = new-object Net.Mail.SmtpClient($smtpServer, $smtpPort);

    $smtp.send($message);
}
Function WriteError([string] $errorMsg, [string] $errorLogPath = "$logDir\UploadToNAS__Error_$currentDateStringForLog.log")
{
	Write-Host $errorMsg
	WriteLog $errorMsg $errorLogPath
	SendToEmail $mailFrom $email "Error!" "$errorMsg"
}
Function WriteLog([string] $logString, [string] $logPath, [bool] $sendToEmail = $false, [string] $mailSubject = "")
{
	$stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
	$logMessage = "$stamp $logString"
	Write-Host $logString
	Add-content $logPath -value $logMessage
	
	if($sendToEmail -eq $true)
	{
		SendToEmail $mailFrom $email $mailSubject $logString
	}
}
Function CheckDirExistence([string] $dir)
{
	If(!(test-path -PathType container $dir))
	{
		Write-Host "Creating dir -> $dir"
		
		try
		{
			New-Item -ItemType Directory -Path $dir
		}
		catch
		{
			WriteError ($dir + " | "  + $_.Exception.message)
			throw ($dir + " | "  + $_.Exception.message)
		}
	}
}
Function CheckFileExistence([string] $filePath)
{
	if(Test-Path -Path $filePath -PathType Leaf)
	{
		Write-Host "File $filePath is ok."
	}
	else
	{
		WriteError "Ошибка доступа к файлу -> $filePath"
		throw "Ошибка доступа к файлу -> $filePath"
	}
}
Function CheckDirAccess([string] $directoryPath)
{
    try
	{
		Write-Host "Check dir -> $directoryPath"
		
        $testPath = Join-Path $directoryPath ([IO.Path]::GetRandomFileName())
        [IO.File]::Create($testPath, 1, 'DeleteOnClose') > $null

        return $true
    } 
	catch
	{
		WriteError "Ошибка доступа к директории -> $directoryPath"
		throw "Ошибка доступа к директории -> $directoryPath"
		
        return $false
    } 
	finally
	{
        Remove-Item $testPath -ErrorAction SilentlyContinue
    }
}
Function GetDbNameFromFileName([string] $fileName)
{
	try 
    {
		$results = [Regex]::Matches($fileName, $regexPattern, $regexMatchOptions)
	}
	catch
	{
		WriteError ($file.Name + " | "  + $_.Exception.message)
	}
	
	if($results[0])
	{
		return $results[0].Groups[1].Value
	}
}
# Load non-working dates from file.
Function GetWorkingDays([string] $workingDaysFilePath)
{
	$parsedDates = @()
	
	try
	{
		$rawDates = Get-Content $workingDaysFilePath
		
		if($rawDates.count -gt 0)
		{
			foreach($rawDate in $rawDates)
			{
				$parsedDate = [datetime]::ParseExact($rawDate, 'dd.MM.yyyy', $null)
				$parsedDates += $parsedDate.ToString("dd.MM.yyyy")
			}
			
			WriteLog ("Загружено рабочих дат: " + $parsedDates.count) $mainAedLogPath
		}
		else
		{
			WriteError "Список с рабочими датами пуст." $errorLogPath
		}
	}
	catch
	{
		WriteError ("Ошибка разбора списка рабочих дат."  + $_.Exception.message) $errorLogPath 
	}
	
	return $parsedDates
}
# Check dates for working and dl time.
Function IsDlTime([string[]] $workingDays)
{
	$res = $false

	$dateTimeNow = Get-Date
	$dateNow = $dateTimeNow.ToString("dd.MM.yyyy")
	$currentDayOfWeek = (Get-Date).DayOfWeek.value__

	if($currentDayOfWeek -eq 6 -or $currentDayOfWeek -eq 7)
	{
		$isCurrentDayInWorkingDaysArr = $workingDays -contains $dateNow
		
		if($isCurrentDayInWorkingDaysArr)
		{
			Write-Host "Match!"
			
			if($dateTimeNow.TimeOfDay -gt $workTimeEndSaturday.TimeOfDay)
			{
				Write-Host "Dl time! (weekends) -> " $dateTimeNow.TimeOfDay
				$res = $true
			}
			else
			{
				Write-Host "Work time! (weekends) -> " $dateTimeNow.TimeOfDay
			}
		}
		else
		{
			Write-Host "Dl time! (weekends) -> " $dateTimeNow.TimeOfDay
			$res = $true
		}
	}
	else
	{
		if($dateTimeNow.TimeOfDay -gt $workTimeStart.TimeOfDay -and $dateTimeNow.TimeOfDay -lt $workTimeEnd.TimeOfDay)
		{
			Write-Host "Work time! -> " $dateTimeNow.TimeOfDay 
		} 
		else 
		{
			Write-Host "Dl time! -> " $dateTimeNow.TimeOfDay
			$res = $true
		}
	}	
	
	return $res
}
Function GetSizeOnDisk([string] $fullPathToFile)
{
    process 
	{
        $absolutePath = ($fullPathToFile | Resolve-Path).Path
        $item = Get-Item $absolutePath
        $volume = Get-Volume $item.PSDrive.Name

        $d = [int][System.Math]::Ceiling($item.Length / $volume.AllocationUnitSize)
        return $d * $volume.AllocationUnitSize
    }
}
Function GetFirstArchive([string] $sourceDir)
{
	$fileName

	try
	{
		$fileName = Get-ChildItem -Filter *.001 -Path $sourceDir -Force | Select-Object -First 1
	}
	catch
	{
		WriteError ("Ошибка получения первого архива в директории $sourceDir. | "  + $_.Exception.message) 
	}

	return $fileName
}
Function CheckArchive([string] $archiveFullPath)
{
	$isArchiveOk = $false
	
	if(Test-Path $7zipPath -PathType Leaf)
	{
		if(Test-Path $archiveFullPath -PathType Leaf)
		{
			Write-Host "Check archive: " $archiveFullPath
		
			& $7zipPath l -bse0 -bso0 $archiveFullPath 
		
			if($?)
			{
				$isArchiveOk = $true
			}
		}
		else
		{
			Write-Host "File " $archiveFullPath " not exist."
		} 
	}
	else
	{
		WriteError "7z path '$7zipPath' not found" $errorLogAedPath
		throw "7z path '$7zipPath' not found"
	}
	
	Write-Host "Archive status: " $isArchiveOk
	
	return $isArchiveOk
}
Function StartRcJob([string] $sourceDirFullPath, [string] $destDirFullPath, [bool] $isBandwidthLimited)
{
	$robocopyScriptBlock= 
	{
		param($sourceDirFullPath, $destDirFullPath, $rcOpions)
		$cmdLine = @($sourceDirFullPath, $destDirFullPath, "*.0*") + $rcOpions
		& 'robocopy.exe' $cmdLine
	}
	
	$rcRunOptions = $robocopyOptionsForAED
	
	if($isBandwidthLimited -eq $true)
	{
		$rcRunOptions = $robocopyOptionsForAEDLimited
	}
	
	Write-Host "Rc job params: " $rcRunOptions
	
	$jobId = Start-Job -ScriptBlock $robocopyScriptBlock -ArgumentList $sourceDirFullPath, $destDirFullPath, $rcRunOptions -Name 'AED_Copy_Job'
	
	return $jobId.Id
}
Function StopRcJob([int] $robocopyJobId)
{
	if($robocopyJobId -gt 0)
	{
		Invoke-Command -ScriptBlock { param($rcId) Stop-job -Id $rcId } -Arg $robocopyJobId
		Write-Host "Stopped: $($robocopyJobId)"
	}
}
Function FormatElapsedTime($ts) 
{
    $elapsedTime = ""

    if ( $ts.Minutes -gt 0 )
    {
        $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    }
    else
    {
        $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
    }

    if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0)
    {
        $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
    }

    if ($ts.Milliseconds -eq 0)
    {
        $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
    }

    return $elapsedTime
}
Function CopyFiles([string] $sourceDirFullPath, [string] $destDirFullPath, [string] $fileName)
{
	try 
    {
		$CmdLine = @($sourceDirFullPath, $destDirFullPath, $fileName) + $robocopyOptions
		& 'robocopy.exe' $CmdLine
	}
	catch
	{
		WriteError ($sourceFileFullPath + " | "  + $_.Exception.message)
	}
}
Function DeleteFiles([string] $sourceDir)
{
	try
	{
		Write-Host "Deleting files in " $sourceDir
		Get-ChildItem -Path $sourceDir *.* | foreach { Remove-Item -Path $_.FullName }
	}
	catch
	{
		WriteError ($sourceDir + " | "  + $_.Exception.message) $errorLogAedPath
	}
}
Function StartLoopForUpload([string] $sD, [string] $dD, [string] $firstArcFile)
{
	$prevBandwidthStatus
	$startDateAndTime = (Get-Date).ToString("dd.MM.yyyy HH:mm:ss")
	
	WriteLog ("AED transfer start -> " + $startDateAndTime) $mainAedLogPath
	
	$sw = [Diagnostics.Stopwatch]::StartNew()

	try 
    {
		$robocopyJobId = 0
		$continue = $true
		
		while($continue)
		{
			$isBandwidthLimited = $true
		
			if($skipScheldule -eq 1)
			{
				$prevBandwidthStatus = $false
				$isBandwidthLimited = $false
			}
			else
			{
				$isBandwidthLimited = !(IsDlTime $workingDays)
			}
		
			Write-Host "Is bandwithLimited: " $isBandwidthLimited " prev status: " $prevBandwidthStatus
			
			if($prevBandwidthStatus -ne $isBandwidthLimited)
			{
				WriteLog "Switching bandwidth limit." $mainAedLogPath
				StopRcJob $robocopyJobId
				$robocopyJobId = 0
				Start-Sleep -Milliseconds $loopSleepTimer
			}
			
			$firstArchiveFullPath = $dD.trim() + "\" + $firstArcFile.trim()
			Write-Host "Check " $firstArchiveFullPath
			$isFileCopied = CheckArchive $firstArchiveFullPath
			Write-Host "RC id: " $robocopyJobId " - Dir status -> " $isFileCopied
				
			if($isFileCopied -eq $true)
			{
				WriteLog ("Dir " + $sD + " successfully copied!") $mainAedLogPath
				StopRcJob $robocopyJobId
				$continue = $false
				DeleteFiles $sD
			}
			elseif($robocopyJobId -eq 0)
			{
				$robocopyJobId = StartRcJob $sD $dD $isBandwidthLimited
				$prevBandwidthStatus = $isBandwidthLimited
				Write-Host "Started rc job with id: " $robocopyJobId
			}
			else
			{
				Start-Sleep -Milliseconds $loopSleepTimer
			}
		}
		
		$sw.Stop()
		$elapsedTime = $sw.Elapsed.Duration()
		
		Write-Host "Sending email."
		WriteLog "AED transfer end. $sD dl completed! Elapsed time: $elapsedTime" $mainAedLogPath $true "Backup ""AED"" Ok!"
	}
	catch
	{
		WriteError ($sourceFileFullPath + " | "  + $_.Exception.message) $errorLogAedPath
	}
	finally
	{
		Write-Host "Rc exit code: " + $lastexitcode
		$robocopyJobId = 0
	}
}

Write-Host "Start"

$errorActionPreference = 'Stop'

CheckDirExistence $logDir
CheckDirAccess $logDir
CheckDirExistence $destDir

if($transferMode -eq 1)
{
	CheckFileExistence $7zipPath
	CheckDirExistence $destDirForAed
	
	Write-Host "Scheldule enabled"
	Write-Host "Db prefix: " $dbPrefix $regexPattern " Db ext: " $fileExt
	Write-Host "Ipg network gap: " $ipgNetworkGap 
	
	$workingDays = GetWorkingDays $workingDaysFilePath
	
	if(CheckDirAccess $destDirForAed)
	{
		Write-Host "$destDirForAed -> OK"
		Write-Host "Copy ..."
			
		$firstArchive = GetFirstArchive $sourceDirForAed
		Write-Host "First archive:" $firstArchive	
		
		$isTkInDelayGroup = $schelduleGroupForDelayedStart -contains $tk
	
		Write-Host "Is tk № $tk in delay group? -> $isTkInDelayGroup"
		
		if($isTkInDelayGroup -eq $true)
		{
			Write-Host "Tk № $tk is delayed."
			$continue = $true
	
			while($continue)
			{
				$dateTimeNow = Get-Date
				$currentDayOfWeek = (Get-Date).DayOfWeek.value__
				$dateNow = $dateTimeNow.ToString("dd.MM.yyyy")
				
				if($currentDayOfWeek -eq $delayedStartDay -and $dateTimeNow -gt $delayedStartTimeStart)
				{
					WriteLog ("Delayed group loop time! $currentDayOfWeek $dateNow") $mainAedLogPath
					$continue = $false
				}
				else
				{
					Write-Host "Awaiting for delayed start ..."
					Start-Sleep -Milliseconds $loopSleepTimer
				}
			}
		}
		else
		{
			Write-Host "Tk № $tk in not delayed."
		}
		
		Write-Host "Nas upload - start."
		StartLoopForUpload $sourceDirForAed $destDirForAed $firstArchive
	}
	else
	{
		WriteError "Директория $destDirForAed недоступна." $errorLogAedPath
	}
}
else
{
	if(CheckDirAccess $destDir)
	{
		Write-Host "$destDir -> OK"
		
		$files = Get-ChildItem -path $sourceDir | Select-Object DirectoryName, Name, FullName

		Write-Host "Copy ..."

		foreach ($file in $files) 
		{
			$dbName = GetDbNameFromFileName $file.Name
		
			if($dbName)
			{
				$dbDestFullPath = $destDir + "\" + $dbName

				CopyFiles $file.DirectoryName $dbDestFullPath $file.Name
			}
		}
	}
	else
	{
		WriteError "Директория $destDir недоступна."
	}
}

Write-Host "End"
