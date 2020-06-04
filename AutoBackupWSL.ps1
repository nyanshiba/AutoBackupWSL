#requires -version 7

#--------------------ユーザ設定--------------------
#設定1
$Settings =
@{
    #コピー先の世代管理ディレクトリ名 この名前で世代管理を行うので、変更は非推奨
    DateTime = "/$((Get-Date).ToString("yyMMdd_HHmmss"))"
    #ログに関する設定
    Log =
    @{
        #ログ保存ディレクトリ
        Path = 'C:\BackupLog'
        #ログローテの閾値
        CntMax = 60
    }
    #通知に関する設定
    Post =
    @{
        #Webhook Url
        hookUrl = 'https://discordapp.com/api/webhooks/XXXXXXXXXX'
        #icon
        Icon = 'https://cdn.discordapp.com/emojis/709988551466549258.png'
    }
}
#設定2
$Settings +=
@{
    #前処理
    BeginScript =
    ({})
    #ミラーバックアップリスト
    MirList =
    @(
        [PSCustomObject]@{
            SrcPath = "minecraft@example.com:~/Servers"
            SrcClude = "--exclude='Test1/world/' --exclude='Test2/world/'"
            DstPath = "D:\RemoteServer"
            Execute = "-e 'ssh -p 22 -o StrictHostKeyChecking=no -i ~/.ssh/remote_id_ed25519'"
            Begin =
            ({
                #Minecraftサーバのバックアップ前に、screenのPIDを取得し、全てにsave-off, save-all flushを送る
                $ScriptBlock =
                ({
                    /usr/bin/screen -ls | ForEach-Object {
                        switch ([Regex]::Split($_, '\t|\.')[1])
                        {
                            {$Null -ne $_ -And '' -ne $_} {
                                /usr/bin/screen -p 0 -S $_ -X eval 'stuff "save-off"\\015'
                                /usr/bin/screen -p 0 -S $_ -X eval 'stuff "save-all\\040flush"\\015'
                                Write-Output "save-off, save-all flush: $_"
                            }
                        }
                    }
                })
                #↑のスクリプトブロックをリモートで実行する。WSL上とWindows上両方に鍵を置く(またはそれと同等の状態)必要がある。  
                # -o StrictHostKeyChecking=no が出来なさそうなので、The authenticity of host can't be established.でyes/noを聞かれないようにしておくこと https://github.com/PowerShell/PowerShell/issues/6650
                Invoke-Command -ScriptBlock $ScriptBlock -HostName minecraft@example.com -Port 22 -KeyFilePath "$env:USERPROFILE/.ssh/remote_id_ed25519"
            })
            End =
            ({
                #Minecraftサーバのバックアップ後、screenのPIDを取得し、全てにsave-onを送る
                $ScriptBlock =
                ({
                    /usr/bin/screen -ls | ForEach-Object {
                        switch ([Regex]::Split($_, '\t|\.')[1])
                        {
                            {$Null -ne $_ -And '' -ne $_} {
                                /usr/bin/screen -p 0 -S $_ -X eval 'stuff "save-on"\\015'
                                Write-Output "save-on: $_"
                            }
                        }
                    }
                })
                Invoke-Command -ScriptBlock $ScriptBlock -HostName minecraft@example.com -Port 22 -KeyFilePath "$env:USERPROFILE/.ssh/remote_id_ed25519"
                #リモートバックアップが終わったことを通知する例
                Send-Webhook -Text "Remote backup is complete."
            })
        }
        [PSCustomObject]@{
            SrcPath = "F:\"
            DstPath = "G:"
        }
    )
    #世代管理バックアップリスト
    GenList =
    @(
        [PSCustomObject]@{
            SrcPath = "D:\RemoteServer"
            DstParentPath = "D:"
            DstGenExclude = 'RemoteServer','190819_070002'
            DstGenThold = 30
        }
        [PSCustomObject]@{
            SrcPath = "$env:APPDATA\.minecraft"
            SrcClude = "--exclude='assets/'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming"
        }
        [PSCustomObject]@{
            SrcPath = "$env:LOCALAPPDATA\TotalMixFX"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Local"
        }
        [PSCustomObject]@{
            SrcPath = "C:\Windows\System32\Tasks"
            SrcClude = "--exclude='*/*'"
            DstParentPath = "D:"
            DstChildPath = "\System32"
        }
        [PSCustomObject]@{
            SrcPath = "$env:USERPROFILE\Documents"
            SrcClude = "--exclude='My Music' --exclude='My Pictures' --exclude='My Videos'"
            DstParentPath = "D:"
        }
        [PSCustomObject]@{
            SrcPath = "C:\Minecraft"
            SrcClude = "--exclude='Spigot1/plugins/CoreProtect/database.db' --exclude='Spigot2/plugins/CoreProtect/database.db'"
            DstParentPath = "D:"
            Begin =
            ({
                Get-Process | Where-Object {$_.MainWindowTitle -in "Survival1","Spigot1","Spigot2"} | ForEach-Object {
                    # Minecraft Server Launcer https://gist.github.com/nyanshiba/deed14b985acfb203c519746d6cea857
                    Invoke-Process -File "pwsh" -Arg "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-off'"
                    Invoke-Process -File "pwsh" -Arg "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-all flush'"
                }
            })
            End =
            ({
                Get-Process | Where-Object {$_.MainWindowTitle -in "Survival1","Spigot1","Spigot2"} | ForEach-Object {
                    Invoke-Process -File "pwsh" -Arg "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-on'"
                }
            })
        }
        [PSCustomObject]@{
            SrcPath = "C:\Rec"
            SrcClude = "--exclude='ts/' --exclude='mp4/'"
            DstParentPath = "D:"
        }
    )
    #後処理
    EndScript =
    ({
        #WSLにマウントされているWindows上のディスク残量一覧をWebhookでPostする例
        Send-Webhook -Text "``````$(wsl /bin/df /mnt/* -h | Out-String)``````"

        #世代管理ディレクトリ構造をWebhookでPostする例（Discordのシンタックスハイライトを利用している）
        Send-Webhook -Text "``````md`n$(Get-FolderStructure -Dir /mnt/d -Depth 1)`n``````"

        #サマリーをWebhookでPostする例
        Send-Webhook -End

        #トースト通知を行う例
        Send-Toast -Icon "$PSHome\assets\Powershell_black.ico" -Title "$(Split-Path $PSCommandPath -Leaf)" -Text "Backup Finished at $End.`n$ErrorCount Errors."
    })
}

#--------------------関数--------------------
function Send-Webhook
{
    param
    (
        [string]$Text,
        [System.Object]$Payload,
        [switch]$End,
        [string]$WebhookUrl = $Settings.Post.hookurl
    )

    if ($Null -eq $WebhookUrl)
    {
        return
    }

    #Payloadが指定されている場合はそのままInvoke-RestMethod

    #Textが指定されている場合はPayloadに変換
    if ($Text)
    {
        switch -Wildcard ($Settings.Post.hookUrl)
        {
            "*discord*"
            {
                $Payload =
                [PSCustomObject]@{
                    content = "$Text"
                }
            }
            "*slack*"
            {
                $Payload =
                [PSCustomObject]@{
                    text = "$Text"
                }
            }
        }
    }

    #Payloadが指定されず、終了時に実行している場合のみPayloadをこちらで用意する
    if (!$Payload -And $End)
    {
        switch -Wildcard ($Settings.Post.hookUrl)
        {
            "*discord*"
            {
                $Payload =
                [PSCustomObject]@{
                    username = "$(Split-Path $PSCommandPath -Leaf)"
                    embeds =
                    @(
                        @{
                            title = "$(Split-Path $PSCommandPath -Leaf)"
                            description = "Backup Summary $($Settings.DateTime)**"
                            color = 0x274a7c
                            thumbnail =
                            @{
                                url = $Settings.Post.Icon
                            }
                            fields =
                            @(
                                @{
                                    name = "Start"
                                    value = $Start
                                    inline = 'true'
                                },
                                @{
                                    name = "End"
                                    value = $End
                                    inline = 'true'
                                }
                                @{
                                    name = "Errors"
                                    value = $ErrorCount
                                    inline = 'true'
                                }
                            )
                        }
                    )
                }
            }
            "*slack*"
            {
                $Payload =
                [PSCustomObject]@{
                    text = "Backup Summary $($Settings.DateTime)**"
                    blocks =
                    @(
                        @{
                            type = "section"
                            text =
                            @{
                                type = "mrkdwn"
                                text = "Backup Summary $($Settings.DateTime)**"
                            }
                            accessory =
                            @{
                                type = "image"
                                image_url = $Settings.Post.Icon
                                alt_text = "$(Split-Path $PSCommandPath -Leaf)"
                            }
                        }
                        @{
                            type = "section"
                            block_id = "section789"
                            fields =
                            @(
                                @{
                                    type = "mrkdwn"
                                    text = "*Start*`n$Start"
                                }
                                @{
                                    type = "mrkdwn"
                                    text = "*End*`n$End"
                                }
                                @{
                                    type = "mrkdwn"
                                    text = "*Errors*`n$ErrorCount"
                                }
                            )
                        }
                    )
                }
            }
        }
    }

    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Headers @{ "Content-Type" = "application/json" } -Body ([System.Text.Encoding]::UTF8.GetBytes(($Payload | ConvertTo-Json -Depth 5)))
}

function Send-Toast
{
    param
    (
        [String]$Icon = "$PSHome\assets\Powershell_black.ico",
        [String]$Title = "$(Split-Path $PSCommandPath -Leaf)",
        [String]$Text = "通知内容が未指定です"
    )
    #Windows PowershellでないPowershellのAppIDを取得
    $AppId = "$((Get-StartApps | Where-Object {$_.Name -match "PowerShell" -And $_.Name -notmatch "Windows"} | Select-Object -First 1).AppID)"

    #ロード済み一覧:[System.AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name }
    #WinRTAPIを呼び出す:[-Class-,-Namespace-,ContentType=WindowsRuntime]
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]

    #XmlDocumentクラスをインスタンス化
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    #LoadXmlメソッドを呼び出し、変数templateをWinRT型のxmlとして読み込む
    $xml.LoadXml(@"
<toast>
<visual>
    <binding template="ToastImageAndText02">
        <image id="1" src="$Icon" alt="Powershell Core"/>
        <text id="1">$Title</text>
        <text id="2">$Text</text>
    </binding>  
</visual>
</toast>
"@)

    #ToastNotificationクラスのCreateToastNotifierメソッドを呼び出し、変数xmlをトースト
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($xml)
}

function Get-FolderStructure
{
    param
    (
        [string]$Dir,
        [int]$Depth = 2,
        [string]$Prefix = ""
    )
    if ($Prefix -eq "")
    {
        "<$Dir>"
    }
    Get-ChildItem $Dir | ForEach-Object {
        if ($_.LinkType -in "Junction","SymbolicLink","HardLink")
        {
            "$Prefix+-[$($_.Name)]()"
        }
        elseif ($_.Extension -in ".gz",".bz2",".xz",".jar")
        {
            "$Prefix+* $($_.Name) *"
        }
        elseif (!$_.PSIsContainer)
        {
            "$Prefix+--$($_.Name)"
        }
        elseif ($_.PSIsContainer)
        {
            "$Prefix+-<$($_.Name)>"
        }
        if ($_.PSIsContainer -And $Depth -ge 2)
        {
            Get-FolderStructure -Dir $_.FullName -Depth ($Depth - 1) -Prefix ($Prefix + '|  ')
        }
    }
}

function ConvertTo-WslPath
{
    param
    (
        [String]$Path
    )
    #wslpathと同様"D:"に対応できない
    #return "$(($Path.Replace('\','/')) -replace '^([A-Z]):/(.*)',"/mnt/$($Path.Substring(0,1).ToLower())/`$2")"
    #"D:"に対応
    return [Regex]::Replace($Path, "^([A-Z]):(\\.*)?", { "/mnt/" + $args.Groups[1].Value.ToLower() + $args.Groups[2].Value.Replace('\','/')})
}

function Invoke-Process
{
    param
    (
        [String]$File,
        [String]$Arg,
        [String[]]$ArgList
    )

    Write-Output "`nInvoke-Process`nFile: $File`nArg: $Arg`nArgList: $ArgList`n"

    #cf. https://github.com/guitarrapc/PowerShellUtil/blob/master/Invoke-Process/Invoke-Process.ps1 

    # new Process
    $ps = New-Object System.Diagnostics.Process
    $ps.StartInfo.UseShellExecute = $False
    $ps.StartInfo.RedirectStandardInput = $False
    $ps.StartInfo.RedirectStandardOutput = $True
    $ps.StartInfo.RedirectStandardError = $True
    $ps.StartInfo.CreateNoWindow = $True
    $ps.StartInfo.Filename = $File
    if ($Arg)
    {
        #Windows
        $ps.StartInfo.Arguments = $Arg
    } elseif ($ArgList)
    {
        #Linux
        $ArgList | ForEach-Object {
            $ps.StartInfo.ArgumentList.Add("$_")
        }
    }

    # Event Handler for Output
    $stdSb = New-Object -TypeName System.Text.StringBuilder
    $errorSb = New-Object -TypeName System.Text.StringBuilder
    $scripBlock = 
    {
        <#$x = $Event.SourceEventArgs.Data
        if (-not [String]::IsNullOrEmpty($x))
        {
            [System.Console]::WriteLine($x)
            $Event.MessageData.AppendLine($x)
        }#>
        if (-not [String]::IsNullOrEmpty($EventArgs.Data))
        {
                    
            $Event.MessageData.AppendLine($Event.SourceEventArgs.Data)
        }
    }
    $stdEvent = Register-ObjectEvent -InputObject $ps -EventName OutputDataReceived -Action $scripBlock -MessageData $stdSb
    $errorEvent = Register-ObjectEvent -InputObject $ps -EventName ErrorDataReceived -Action $scripBlock -MessageData $errorSb

    # execution
    $Null = $ps.Start()
    $ps.BeginOutputReadLine()
    $ps.BeginErrorReadLine()

    # wait for complete
    $ps.WaitForExit()
    $ps.CancelOutputRead()
    $ps.CancelErrorRead()

    # verbose Event Result
    $stdEvent, $errorEvent | Out-String -Stream | Write-Verbose

    # Unregister Event to recieve Asynchronous Event output (You should call before process.Dispose())
    Unregister-Event -SourceIdentifier $stdEvent.Name
    Unregister-Event -SourceIdentifier $errorEvent.Name

    # verbose Event Result
    $stdEvent, $errorEvent | Out-String -Stream | Write-Verbose

    # Get Process result
    $stdSb.ToString().Trim()
    $errorSb.ToString().Trim()
    Write-Output "ExitCode: $($ps.ExitCode)"
    [Array]$script:ExitCode += $ps.ExitCode

    if ($Null -ne $process)
    {
        $ps.Dispose()
    }
    if ($Null -ne $stdEvent)
    {
        $stdEvent.StopJob()
        $stdEvent.Dispose()
    }
    if ($Null -ne $errorEvent)
    {
        $errorEvent.StopJob()
        $errorEvent.Dispose()
    }
}

function Invoke-DiffBackup
{
    param
    (
        [String]$Execute,
        [String]$Src,
        [String]$Clude,
        [String]$Dst,
        [ScriptBlock]$Begin,
        [ScriptBlock]$End
    )
    if ($Begin)
    {
        Invoke-Command -ScriptBlock $Begin
    }
    if ($IsWindows)
    {
        $Src = ConvertTo-WslPath -Path $Src
        $Dst = ConvertTo-WslPath -Path $Dst
        Invoke-Process -File "wsl" -Arg "/usr/bin/rsync $Execute -av --delete --delete-excluded $Clude `"$Src`" `"$Dst`""
    } elseif ($IsLinux)
    {
        Invoke-Process -File "/bin/sh" -ArgList "-c", "/usr/bin/rsync $Execute -av --delete --delete-excluded $Clude '$Src' '$Dst'"
    }
    if ($End)
    {
        Invoke-Command -ScriptBlock $End
    }
}

function Invoke-IncrBackup
{
    param
    (
        [String]$Link,
        [String]$Src,
        [String]$Clude,
        [String]$Dst,
        [ScriptBlock]$Begin,
        [ScriptBlock]$End
    )
    if ($Begin)
    {
        Invoke-Command -ScriptBlock $Begin
    }
    if ($IsWindows)
    {
        $Link = ConvertTo-WslPath -Path $Link
        $Src = ConvertTo-WslPath -Path $Src
        $Dst = ConvertTo-WslPath -Path $Dst
        Invoke-Process -File "wsl" -Arg "/usr/bin/rsync -av --delete --delete-excluded $Clude --link-dest=`"$Link`" `"$Src`" `"$Dst`""
    } elseif ($IsLinux)
    {
        Invoke-Process -File "/bin/sh" -ArgList "-c", "/usr/bin/rsync -av --delete --delete-excluded $Clude --link-dest='$Link' '$Src' '$Dst'"
    }
    if ($End)
    {
        Invoke-Command -ScriptBlock $End
    }
}

#バックアップ開始時刻
$Start = "$((Get-Date).ToString("yyyy-MM-dd (ddd) HH:mm:ss"))"

#ログ取り開始
Start-Transcript -LiteralPath "$($Settings.Log.Path)$($Settings.DateTime).log.md"

#--------------------ログローテ--------------------
#古いログの削除
Write-Output "`n## ログローテ`n"
Get-ChildItem -LiteralPath "$($Settings.Log.Path)/" -Include *.txt,*.log.md | Sort-Object LastWriteTime -Descending | Select-Object -Skip $Settings.Log.CntMax | ForEach-Object {
    Write-Output "Log: Deleted $_"
    Remove-Item -LiteralPath "$_"
}

#ユーザ設定をログに記述
Write-Output @"
## ユーザ設定

``````json
$($Settings | ConvertTo-Json | Out-String -Width 4096)
``````
"@

#--------------------前処理--------------------
Write-Output "`n## 前処理`n"
&$Settings.BeginScript

#--------------------ミラー--------------------
Write-Output "`n## ミラー`n"
#ミラーリストの中から、最低限の設定項目があるもののみ実行
$Settings.MirList | Where-Object {$_.SrcPath -And $_.DstPath} | ForEach-Object {
    Write-Output "`n### $($_.SrcPath)`n"
    Write-Output "``````$($_ | Format-Table -Property * | Out-String -Width 4096)``````"
    #コピー先が無ければ新しいディレクトリの作成
    if (!(Test-Path "$($_.DstPath)"))
    {
        Write-Output "New-Item $($_.DstPath)"
        $Null = New-Item "$($_.DstPath)" -itemType Directory
    }
    #差分バックアップ
    Invoke-DiffBackup -Execute "$($_.Execute)" -Clude "$($_.SrcClude)" -Src "$($_.SrcPath)" -Dst "$($_.DstPath)" -Begin $_.Begin -End $_.End
}

#--------------------世代管理--------------------
Write-Output "`n## 世代管理`n"
#世代管理リストの中から、最低限の設定項目があるもののみ実行
$Settings.GenList | Where-Object {$_.SrcPath -And $_.DstParentPath} | ForEach-Object {
    Write-Output "`n### $($_.SrcPath)`n"
    Write-Output "``````$($_ | Format-Table -Property * | Out-String -Width 4096)``````"
    #同じコピー先で最初 グローバル設定 世代管理
    if ($_.DstGenThold)
    {
        #Excludeは必ず指定する必要があるため、有無で条件分岐する必要はない
        $AllGen = Get-ChildItem -Directory "$($_.DstParentPath)/*" -Name -Exclude $_.DstGenExclude
        if (!$?)
        {
            #コピー先で例外 今回のループは抜ける
            return "Get-ChildItem $($_.DstParentPath)/* Exception."
        } elseif ($AllGen.Count -ge $_.DstGenThold)
        {
            #世代数が閾値以上なので閾値内に丸めて最も古い世代をリネームしインクリメンタルバックアップ
            foreach ($OldGen in ($AllGen | Sort-Object -Descending | Select-Object -Skip $_.DstGenThold))
            {
                Write-Output "Remove-Item: $($_.DstParentPath)/$OldGen"
                Remove-Item -LiteralPath "$($_.DstParentPath)/$OldGen" -Recurse -Force
            }
            Write-Output "Rename-Item: $($_.DstParentPath)/$($AllGen | Select-Object -Last $_.DstGenThold | Select-Object -Index 0) -> $($_.DstParentPath)$($Settings.DateTime)"
            Rename-Item -LiteralPath "$($_.DstParentPath)/$($AllGen | Select-Object -Last $_.DstGenThold | Select-Object -Index 0)" "$($_.DstParentPath)$($Settings.DateTime)"
        }
        #世代管理の親ディレクトリ構造をログに出力
        "``````markdown`n$(Get-FolderStructure -Dir $_.DstParentPath -Depth 1)`n``````"
    }
    #新しいディレクトリの作成 世代数が閾値未満、DstChildPath、設定の不備への対応
    if (!(Test-Path "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)"))
    {
        Write-Output "New-Item $($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)"
        $Null = New-Item "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -itemType Directory
    }
    #常に行う グローバル設定がなされていれば$AllGen変数が引き継がれるため、ここが正しく実行される
    #$AllGen.Countの値は更新されないので、1世代目が作成されても0 フルかインクリメンタル(リンク先の有無)の判別に使える
    #設定に不備があっても$AllGen.Count -eq 0でフルバックアップに流れてくれるかな…くらいのお気持ち
    if ($AllGen.Count -eq 0)
    {
        #世代数が0なので新しいディレクトリにフルバックアップ
        Invoke-DiffBackup -Clude "$($_.SrcClude)" -Src "$($_.SrcPath)" -Dst "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -Begin $_.Begin -End $_.End
    } else {
        #インクリメンタルバックアップ リンク先は今の世代が作成される前の時点での最も新しい世代なので-Last 1でおk
        Invoke-IncrBackup -Clude "$($_.SrcClude)" -Link "$($_.DstParentPath)/$($AllGen | Select-Object -Last 1)$($_.DstChildPath)" -Src "$($_.SrcPath)" -Dst "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -Begin $_.Begin -End $_.End
    }
}

#バックアップ終了時刻
$End = "$((Get-Date).ToString("yyyy-MM-dd (ddd) HH:mm:ss"))"

#終了コード配列を参照して、エラーがあった数を集計
$ErrorCount = ($ExitCode | Where-Object {$_ -ne 0}).Count

#--------------------後処理--------------------
Write-Output "`n## 後処理`n"
&$Settings.EndScript


#ログ取り停止
Stop-Transcript
