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
        Path = 'C:\logs\AutoBackupWSL'
        #ログローテの閾値
        CntMax = 365
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
        #稼働中のMinecraftサーバのデータを安全にミラーする例
        [PSCustomObject]@{
            SrcPath = "minecraft@example.com:~/Servers"
            rsyncArgument = "-avz --delete --delete-excluded --exclude='Test1/world/' --exclude='Test2/world/' -e 'ssh -p 22 -o StrictHostKeyChecking=no -i ~/.ssh/remote_id_ed25519' --bwlimit=1750"
            DstPath = "D:\Mirror\RemoteServer"
            Begin =
            ({
                #このディレクトリのバックアップ直前に行われる処理

                #AutoBackupWSL.ps1内のInvoke-Command関数を使用してプロセスを実行する例
                #Minecraftサーバのバックアップ前にsave-off, save-all flushを送る例 https://nyanshiba.com/blog/minecraft-java-edition-3-linux-server#discordへ通知できるminecraftサーバ管理スクリプト-msl.ps1
                Invoke-Command -ScriptBlock ${function:Invoke-Process} -ArgumentList 'pwsh', "-File /home/minecraft/Servers/msl.ps1 -Name CBWSurvival -Action `"save-off`"" -HostName minecraft@example.com -Port 22 -KeyFilePath "$env:USERPROFILE/.ssh/remote_id_ed25519"
                Invoke-Command -ScriptBlock ${function:Invoke-Process} -ArgumentList 'pwsh', "-File /home/minecraft/Servers/msl.ps1 -Name CBWSurvival -Action `"save-all flush`"" -HostName minecraft@example.com -Port 22 -KeyFilePath "$env:USERPROFILE/.ssh/remote_id_ed25519"
                #WSL上とWindows上両方に鍵を置く(またはそれと同等の状態)必要がある。  
                # -o StrictHostKeyChecking=no が出来なさそうなので、The authenticity of host can't be established.でyes/noを聞かれないようにしておくこと https://github.com/PowerShell/PowerShell/issues/6650
            })
            End =
            ({
                #このディレクトリのバックアップ直後に行われる処理

                #Minecraftサーバのバックアップ後にsave-onを送る例
                Invoke-Command -ScriptBlock ${function:Invoke-Process} -ArgumentList 'pwsh', "-File /home/minecraft/Servers/msl.ps1 -Name CBWSurvival -Action `"save-on`"" -HostName minecraft@example.com -Port 22 -KeyFilePath "$env:USERPROFILE/.ssh/remote_id_ed25519"
                #リモートバックアップが終わったことを通知する例
                Send-Webhook -Text "minecraft@example.com Remote backup is complete." -WebhookUrl "https://discordapp.com/api/webhooks/ZZZZZZZZZZ"
            })
        }
        #リモートサーバのアーカイブをミラーしつつ、変更があっても削除しない例
        [PSCustomObject]@{
            SrcPath = "example.com:/mnt/backup/Archive" #rsyncは.ssh/configを参照できる
            rsyncArgument = "-avz --bwlimit=1750" #deleteを同期しない
            DstPath = "D:\Mirror\RemoteServer"
        }
        #Xperia 1 II(Android 11端末)のTermuxのsshdに接続してリモートバックアップ
        [PSCustomObject]@{
            SrcPath = "12:/storage/emulated/0/DCIM"
            rsyncArgument = "-avz --copy-links --exclude='.trashed-*'" # Android<=10は https://wiki.termux.com/wiki/Termux-setup-storage
            DstPath = "E:\Xperia1ii"
        }
        [PSCustomObject]@{
            SrcPath = "12:/storage/emulated/0/Pictures"
            rsyncArgument = "-avz --update --copy-links --exclude='.thumbnails/' --exclude='Twitter/' --exclude='Instagram/'" # --updateでEndの処理内容を再度上書きしないように
            DstPath = "E:\Xperia1ii"
            End =
            {
                # Android版LightroomはICCプロファイルを埋め込まないので、バックアップログを見てLightroomで書き出されたjpgにDisplay P3のICCを埋め込む例
                $Process.StandardOutput -split '\r?\n' | Select-String -Pattern "pictures/AdobeLightroom/.+.jpg" | ForEach-Object {
                    Write-Host $_
                    $img = "E:\Xperia1ii\" + $_ -replace '/','\'
                    magick.exe identify -format "%[profiles]\n" $img
                    magick.exe convert $img -profile E:\Xperia1ii\DisplayP3.icc $img
                    magick.exe identify -format "%[profiles]\n" $img
                }
            }
        }
        [PSCustomObject]@{
            SrcPath = "12:/storage/emulated/0/Movies"
            rsyncArgument = "-avz --copy-links"
            DstPath = "E:\Xperia1ii"
        }
        #WSL上のホームディレクトリをミラーする例
        [PSCustomObject]@{
            SrcPath = "/home/sbn/"
            rsyncArgument = "-a --delete --delete-excluded --exclude='.npm/' --include='*/' --include='.bashrc' --include='.gitconfig' --include='.ssh/***' --exclude='*'"
            DstPath = "E:\fedora34-dev"
        }
        #/mnt/c/下のリポジトリルートディレクトリをミラーする例 WSL2はWSL1に比べて数～十数倍遅いので、AutoBackupWSLでは非推奨 https://github.com/microsoft/WSL/issues/4197
        [PSCustomObject]@{
            SrcPath = "C:\repos"
            # FFmpegやOBSのビルド関係を除外
            rsyncArgument = "-a -delete --delete-excluded --exclude='dependencies2019/' --exclude='cef_binary*/' --exclude='obs-studio/' --exclude='nv-codec-headers/' --exclude='node_modules/' --include='media-autobuild_suite/build/media-autobuild_suite.ini' --include='media-autobuild_suite/local64/bin-*/***' --exclude='media-autobuild_suite/***'"
            DstPath = "E:"
        }
        #タスクスケジューラの設定もバックアップ AutoBackupWSLもこれを使うと思うので
        [PSCustomObject]@{
            SrcPath = "C:\Windows\System32\Tasks"
            rsyncArgument = "-av --delete --delete-excluded --exclude='*/'"
            DstPath = "D:\Mirror\System32"
        }
        #手動インストールした実行ファイル群をバックアップ
        [PSCustomObject]@{
            SrcPath = "C:\bin"
            #JDKやJREを除外
            rsyncArgument = "-av --delete --delete-excluded --exclude='*jdk*/' --exclude='*jre*/' --include='*/'"
            DstPath = "E:"
        }
        #HDDの片方向ミラー 冗長化は記憶域プールでやろう https://nyanshiba.com/blog/powershell-storagepool
        [PSCustomObject]@{
            SrcPath = "F:\"
            rsyncArgument = "-av --delete"
            DstPath = "G:"
        }
    )
    #世代管理バックアップリスト
    GenList =
    @(
        #リモートサーバからミラーしたものを世代管理
        [PSCustomObject]@{
            SrcPath = "D:\Mirror\RemoteServer"
            rsyncArgument = "-av --delete --delete-excluded --exclude='Archive/'" #この例では、リモートサーバのアーカイブは--deleteしていないので世代管理から除外できる
            DstParentPath = "D:" #"D:\yyMMdd_HHmmss"に世代管理される
            DstGenExclude = 'Mirror','190819_070002' #この例では"D:\Mirror"を使っているので世代参照・ローテートから除外する
            DstGenThold = 30 #同じDstParentPathの中で最初に書けば良い この例では、30世代溜まるとインクリメンタル時に古い世代を使い回すので(ログが汚くなるが)更に速くなる
        }
        #Minecraft Launcherの設定のみバックアップ
        [PSCustomObject]@{
            SrcPath = "$env:APPDATA\.minecraft"
            rsyncArgument = "-av --delete --delete-excluded --include='*.*' --exclude='*/'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming" #"D:\yyMMdd_HHmmss\AppData\Roaming"に世代管理される
        }
        #VSCodeのユーザ設定とUntitledのみ最低限バックアップ
        [PSCustomObject]@{
            SrcPath = "$env:APPDATA\Code"
            rsyncArgument = "-av --delete --delete-excluded --include='*/' --include='Backups/workspaces.json' --include='Backups/***/untitled/***' --include='User/settings.json' --include='User/keybindings.json' --include='User/workspaceStorage/***/***' --exclude='*'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming"
        }
        #OBS Studioの最低限の設定のみバックアップ
        [PSCustomObject]@{
            SrcPath = "$env:APPDATA\obs-studio"
            rsyncArgument = "-av --delete --delete-excluded --exclude='crashes/' --exclude='logs/' --exclude='plugin_config/'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming"
        }
        #Firefoxの設定(ブックマーク, タブ, コンテナ, Cookie, 拡張機能リスト, ファビコン, パスワード, Cookieの例外, about:config)をバックアップ
        [PSCustomObject]@{
            SrcPath = "$env:APPDATA\Mozilla\Firefox\Profiles"
            rsyncArgument = "-av --delete --delete-excluded --exclude='storage/' --include='*/' --include='*default/bookmarkbackups/***' --include='*default/sessionstore-backups/***' --include='*default/containers.json' --include='*default/cookies.sqlite' --include='*default/extensions.json' --include='*default/favicons.sqlite' --include='*default/key4.db' --include='*default/logins.json' --include='*default/permissions.sqlite'--include='*default/prefs.js' --include='*default/sessionstore.jsonlz4' --exclude='*'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming\Mozilla\Firefox"
        }
        #Logicool Optionsの感度・キー割り当て設定
        <#[PSCustomObject]@{
            SrcPath = "$env:APPDATA\Logishrd\LogiOptions\devices"
            rsyncArgument = "-av --delete --delete-excluded --include='*/' --include='***/***.xml' --exclude='*'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Roaming\Logishrd\LogiOptions"
        }#>
        #Logicool Options+
        [PSCustomObject]@{
            SrcPath = "$env:LOCALAPPDATA\LogiOptionsPlus"
            rsyncArgument = "-av --delete --delete-excluded --include='*/' --include='settings.json' --include='***/***.xml' --exclude='*'"
            DstParentPath = "D:"
            DstChildPath = "\AppData\Local"
        }
        #ドキュメントのバックアップ
        [PSCustomObject]@{
            SrcPath = "$env:USERPROFILE\Documents"
            rsyncArgument = "-av --no-links --delete --delete-excluded --exclude='My Music' --exclude='My Pictures' --exclude='My Videos'"
            DstParentPath = "D:"
        }
        #Minecraftのゲームディレクトリ、ローカルサーバを安全にバックアップ
        [PSCustomObject]@{
            SrcPath = "C:\Minecraft"
            rsyncArgument = "-av --no-links --delete --delete-excluded --exclude='Spigot1/plugins/CoreProtect/database.db' --exclude='Spigot2/plugins/CoreProtect/database.db'"
            DstParentPath = "D:"
            Begin =
            ({
                Get-Process | Where-Object {$_.MainWindowTitle -in "Survival1","Spigot1","Spigot2"} | ForEach-Object {
                    # Minecraft Server Launcer https://gist.github.com/nyanshiba/deed14b985acfb203c519746d6cea857
                    Invoke-Process -FileName "pwsh" -Arguments "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-off'"
                    Invoke-Process -FileName "pwsh" -Arguments "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-all flush'"
                }
            })
            End =
            ({
                Get-Process | Where-Object {$_.MainWindowTitle -in "Survival1","Spigot1","Spigot2"} | ForEach-Object {
                    Invoke-Process -FileName "pwsh" -Arguments "-File 'C:\bin\msl.ps1' -Name $_ -Action 'save-on'"
                }
            })
        }
    )
    #後処理
    EndScript =
    ({
        #WSLにマウントされているWindows上のディスク残量一覧をWebhookでPostする例
        Send-Webhook -Text "``````$(wsl /bin/df /mnt/* -h | Out-String)``````"

        #世代管理ディレクトリ構造をWebhookでPostする例（Discordのシンタックスハイライトを利用している）
        Send-Webhook -Text "``````md`n$(Get-FolderStructure -Dir D: -Depth 1 | Select-Object -First 5 -Last 10 | Out-String -Width 4096)`n``````"

        #サマリーをWebhookでPostする例
        Send-Webhook -EndEmbed

        #トースト通知を行う例
        Send-Toast
    })
}

#--------------------関数--------------------
function Send-Webhook
{
    param
    (
        [string]$Text,
        [System.Object]$Payload,
        [switch]$EndEmbed,
        [string]$WebhookUrl = $Settings.Post.hookurl
    )

    if ($Null -eq $WebhookUrl)
    {
        return "ERROR Send-Webhook Webhook URL not exist"
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
                    username = "$(Split-Path $PSCommandPath -Leaf)"
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

    #EmdEmbedが指定された場合は終了時用のPayloadをつくる
    if ($EndEmbed)
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
                            description = "Backup Summary $($Settings.DateTime)"
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
                    text = "Backup Summary $($Settings.DateTime)"
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
    # Windows以外はお帰り下さい
    if (!$IsWindows)
    {
        return "Send-Toast: This function only supports Windows."
    }
    
    #PowershellのAppIDを取得
    $AppId = "$((Get-StartApps | Where-Object {$_.Name -match "PowerShell" -And $_.Name -notmatch "Windows"} | Select-Object -First 1).AppID)"

    #PowerShell 7.1.0-preview.4~
    #https://github.com/PowerShell/PowerShell/issues/13042#issuecomment-653357546
    #https://github.com/Windos/BurntToast/blob/main/BurntToast/BurntToast.psm1
    #https://github.com/Windos/BurntToast/blob/main/BurntToast/Public/Submit-BTNotification.ps1
    if ($PSVersionTable.PSVersion -ge [System.Management.Automation.SemanticVersion] '7.1.0-preview.4')
    {
        try
        {
            Get-Package -Name Microsoft.Windows.SDK.NET.Ref -ErrorAction Stop
        }
        catch
        {
            # NuGetの登録が必要か確認して、トーストに必要なパッケージをインストールする
            try
            {
                # Get-PackageProviderにはNuGetがあるが、Get-PackageSourceにはない
                Get-PackageSource -ProviderName NuGet -ErrorAction Stop -WarningAction Stop
            }
            catch [System.Management.Automation.ActionPreferenceStopException] # $_.Exception.GetType().FullName
            {
                Write-Host "DEBUG Get-PackageSource: NuGet not found. So run Register-PackageSource."

                # NuGetを登録
                # このスクリプトの実行環境は [Net.ServicePointManager]::SecurityProtocol = Tls12 or SystemDefault なので関係ないし、 Install-Module や Install-PackageProvider はうまくいかない
                # https://github.com/PowerShell/PowerShellGetv2/issues/599
                # https://docs.microsoft.com/en-us/powershell/module/packagemanagement/register-packagesource#example-1--register-a-package-source-for-the-nuget-provider
                Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -ErrorAction Stop
            }
            
            # Install-Package: No match was found for the specified search criteria and package name 'Microsoft.Windows.SDK.NET.Ref'. Try Get-PackageSource to see all available registered package sources.
            Install-Package Microsoft.Windows.SDK.NET.Ref -Scope CurrentUser -Force
        }
        finally
        {
            $Library = (Get-Item (Get-Package -Name Microsoft.Windows.SDK.NET.Ref).Source).DirectoryName
            Add-Type -AssemblyName "$Library\lib\Microsoft.Windows.SDK.NET.dll"
            Add-Type -AssemblyName "$Library\lib\WinRT.Runtime.dll"
        }
    }
    else
    {
        #ロード済み一覧:[System.AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name }
        #WinRTAPIを呼び出す:[-Class-,-Namespace-,ContentType=WindowsRuntime]
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    }

    #XmlDocumentクラスをインスタンス化
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    #LoadXmlメソッドを呼び出し、変数templateをWinRT型のxmlとして読み込む
    #https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/adaptive-interactive-toasts
    $xml.LoadXml(@"
<toast>
<visual>
    <binding template="ToastGeneric">
        <text hint-maxLines="1">$(Split-Path $PSCommandPath -Leaf)</text>
        <text>Backup Finished.</text>
        <text>$End</text>
        <group>
            <subgroup>
                <text hint-style="base">$ErrorCount errors</text>
                <text hint-style="captionSubtle">$($Settings.Log.Path)$($Settings.DateTime).log</text>
            </subgroup>
        </group>
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
    #wslpath -uが対応していないドライブレターのみ(例: "D:" -> "/mnt/d")に対応する正規表現
    return [Regex]::Replace($Path, "^([A-Z]):(\\.*|/.*)?", { "/mnt/" + $args.Groups[1].Value.ToLower() + $args.Groups[2].Value.Replace('\','/')})
    <#
    #相対パス(例: "\AppData\Roaming" -> "/AppData/Roaming")にも対応
    return [Regex]::Replace(
        $Path,
        "^([A-Z]|)(:|)(\\.*|/.*)?", #[0]([1]ドライブレターがあったりなかったり)([2]ドライブレター後のコロンがあったりなかったり)([3]バックスラッシュかスラッシュ後になんやかんや)
        {
            $(
                switch ($args.Groups[1].Value)
                {
                    "" {""} #相対パス
                    default {"/mnt/$($_.ToLower())"} #完全パス
                }
            ) + $args.Groups[3].Value.Replace('\','/') #残りのバックスラッシュをスラッシュへ置き換え
        }
    )
    #>
}

function Invoke-Process
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $FileName,
        [String]
        $Arguments,
        [String[]]
        $ArgumentList
    )
    Write-Host "DEBUG Invoke-Process`nFile: $FileName`nArguments: $Arguments`nArgumentList: $ArgumentList`n"
    
    if ([String]::IsNullOrEmpty($FileName) -And ([String]::IsNullOrEmpty($Arguments) -Or [String]::IsNullOrEmpty($ArgumentList)))
    {
        return [PSCustomObject]@{
            StartTime = Get-Date
            StandardOutput = ""
            ErrorOutput = "Invoke-Process: 引数の指定が間違っています"
            ExitCode = 1
        }
    }

    #cf. https://github.com/guitarrapc/PowerShellUtil/blob/master/Invoke-Process/Invoke-Process.ps1 

    # new Process
    $ps = New-Object System.Diagnostics.Process
    $ps.StartInfo.UseShellExecute = $False
    $ps.StartInfo.RedirectStandardInput = $False
    $ps.StartInfo.RedirectStandardOutput = $True
    $ps.StartInfo.RedirectStandardError = $True
    $ps.StartInfo.CreateNoWindow = $True
    $ps.StartInfo.Filename = $FileName
    if ($IsWindows)
    {
        $ps.StartInfo.Arguments = $Arguments
    }
    elseif ($IsLinux)
    {
        $ArgumentList | ForEach-Object {
            $ps.StartInfo.ArgumentList.Add("$_")
        }
    }

    # Event Handler for Output
    $stdSb = New-Object -TypeName System.Text.StringBuilder
    $errorSb = New-Object -TypeName System.Text.StringBuilder
    $scripBlock = 
    {
        $x = $Event.SourceEventArgs.Data
        if (-not [String]::IsNullOrEmpty($x))
        {
            [System.Console]::WriteLine($x)
            $Event.MessageData.AppendLine($x)
        }
    }
    $stdEvent = Register-ObjectEvent -InputObject $ps -EventName OutputDataReceived -Action $scripBlock -MessageData $stdSb
    $errorEvent = Register-ObjectEvent -InputObject $ps -EventName ErrorDataReceived -Action $scripBlock -MessageData $errorSb

    # execution
    $ps.Start() > $Null
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
    return [PSCustomObject]@{
        StartTime = $ps.StartTime
        StandardOutput = $stdSb.ToString().Trim()
        ErrorOutput = $errorSb.ToString().Trim()
        ExitCode = $ps.ExitCode
    }

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
        [String]$Src,
        [String]$rsyncArgument,
        [String]$Dst,
        [ScriptBlock]$Begin,
        [ScriptBlock]$End
    )

    Write-Host "DEBUG Invoke-DiffBackup"

    if ($Begin)
    {
        Invoke-Command -ScriptBlock $Begin
    }

    if ($IsWindows)
    {
        $Src = ConvertTo-WslPath -Path $Src
        $Dst = ConvertTo-WslPath -Path $Dst
        $Process = Invoke-Process -FileName "wsl" -Arguments "/usr/bin/rsync $rsyncArgument `"$Src`" `"$Dst`""
    }
    elseif ($IsLinux)
    {
        $Process = Invoke-Process -FileName "/bin/sh" -ArgumentList "-c", "/usr/bin/rsync $rsyncArgument '$Src' '$Dst'"
    }

    [Array]$script:ExitCode += $Process.ExitCode
    $Process | Format-List -Property *
    
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
        [String]$rsyncArgument,
        [String]$Dst,
        [ScriptBlock]$Begin,
        [ScriptBlock]$End
    )

    Write-Host "DEBUG Invoke-IncrBackup"

    if ($Begin)
    {
        Invoke-Command -ScriptBlock $Begin
    }

    if ($IsWindows)
    {
        $Link = ConvertTo-WslPath -Path $Link
        $Src = ConvertTo-WslPath -Path $Src
        $Dst = ConvertTo-WslPath -Path $Dst
        $Process = Invoke-Process -FileName "wsl" -Arguments "/usr/bin/rsync $rsyncArgument --link-dest=`"$Link`" `"$Src`" `"$Dst`""
    }
    elseif ($IsLinux)
    {
        $Process = Invoke-Process -FileName "/bin/sh" -ArgumentList "-c", "/usr/bin/rsync $rsyncArgument --link-dest='$Link' '$Src' '$Dst'"
    }

    [Array]$script:ExitCode += $Process.ExitCode
    $Process | Format-List -Property *

    if ($End)
    {
        Invoke-Command -ScriptBlock $End
    }
}

#バックアップ開始時刻
$Start = "$((Get-Date).ToString("yyyy-MM-dd (ddd) HH:mm:ss"))"

#ログ取り開始
Start-Transcript -LiteralPath "$($Settings.Log.Path)$($Settings.DateTime).log"

"#--------------------ログローテ--------------------"
# 古いログの削除
Get-ChildItem -LiteralPath "$($Settings.Log.Path)/" -Include *.txt,*.log | Sort-Object LastWriteTime -Descending | Select-Object -Skip $Settings.Log.CntMax | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName
    "INFO Remove-Item: $_"
}

"#--------------------ユーザ設定--------------------"
#ユーザ設定をログに記述
foreach ($line in (Get-Content -LiteralPath $PSCommandPath) -split "`n")
{
    if ($line -match '#--------------------関数--------------------')
    {
        break
    }
    $line
}

"#--------------------前処理--------------------"
&$Settings.BeginScript

"#--------------------ミラー--------------------"
#ミラーリストの中から、最低限の設定項目があるもののみ実行
$Settings.MirList | Where-Object {$_.rsyncArgument -And $_.SrcPath -And $_.DstPath} | ForEach-Object {

    #ログ
    $_ | Format-Table -Property * | Out-String -Width 4096

    #コピー先が無ければ新しいディレクトリの作成
    if (!(Test-Path "$($_.DstPath)"))
    {
        "INFO New-Item $($_.DstPath)"
        $Null = New-Item "$($_.DstPath)" -itemType Directory
    }

    #差分バックアップ
    Invoke-DiffBackup -rsyncArgument "$($_.rsyncArgument)" -Src "$($_.SrcPath)" -Dst "$($_.DstPath)" -Begin $_.Begin -End $_.End
}

"#--------------------世代管理--------------------"
#設定の不備がないディレクトリのみ実行
$Settings.GenList | Where-Object {$_.rsyncArgument -And $_.SrcPath -And $_.DstParentPath} | ForEach-Object {

    #ログ
    $_ | Format-Table -Property * | Out-String -Width 4096

    #DstGenTholdが設定されたディレクトリを処理する直前に世代管理ローテーションを行う
    #同じDstParentPathを持つディレクトリは、ここで作成されたディレクトリを使って世代管理される
    if ($_.DstGenThold)
    {
        #DstParentPath内の世代管理されたディレクトリを取得する
        try
        {
            #try catchで捕まえるために-ErrorAction Stopが必要
            $AllGen = Get-ChildItem -Directory "$($_.DstParentPath)/*" -Name -Exclude $_.DstGenExclude -ErrorAction Stop
        }
        catch
        {
            #DstParentPath内の世代の取得すらできないのでループを抜ける
            return "ERROR Get-ChildItem DstParentPath: $($_.DstParentPath)/*"
        }

        #世代数をが閾値未満に丸め込む
        if ($AllGen.Count -ge $_.DstGenThold)
        {
            #世代数が余計にある場合は削除 世代数を減らさない限り実行されない かなり時間がかかる
            foreach ($OldGen in ($AllGen | Sort-Object -Descending | Select-Object -Skip $_.DstGenThold))
            {
                Remove-Item -LiteralPath "$($_.DstParentPath)/$OldGen" -Recurse -Force
                "INFO Remove-Item: $($_.DstParentPath)/$OldGen"
            }
            #DstParentPath内で最も古い世代をリネームして使いまわすことで、変更が少ない&サイズが大きい場合は特に速くなる
            #直前の世代とのインクリメンタルではあるが、最古の世代との差分であるので、ログが汚くなる
            Rename-Item -LiteralPath "$($_.DstParentPath)/$($AllGen | Select-Object -Index 0)" "$($_.DstParentPath)$($Settings.DateTime)"
            "INFO Rename-Item: $($_.DstParentPath)/$($AllGen | Select-Object -Index 0) -> $($_.DstParentPath)$($Settings.DateTime)"
        }

        #DstParentPath内のディレクトリ構造をログに出力
        Get-FolderStructure -Dir $_.DstParentPath -Depth 1
    }

    #新しい世代ディレクトリの作成
    #ディレクトリの存在で判別すると DstGenTholdが設定されたディレクトリ, DstChildPathが設定されたディレクトリ, 設定の不備 に実行される
    if (!(Test-Path "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)"))
    {
        $Null = New-Item "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -itemType Directory
        "INFO New-Item $($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)"
    }

    #フルバックアップかインクリメンタルバックアップを実行するかの判別
    #$AllGen.Countは新しい世代ディレクトリの作成前の値なので、最初に世代が無ければ常に0
    if ($AllGen.Count -eq 0)
    {
        #世代数が0なので新しいディレクトリにフルバックアップ
        Invoke-DiffBackup -rsyncArgument "$($_.rsyncArgument)" -Src "$($_.SrcPath)" -Dst "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -Begin $_.Begin -End $_.End
    } else {
        #インクリメンタルバックアップ リンク先は新しい世代ディレクトリ作成前の最後尾
        Invoke-IncrBackup -rsyncArgument "$($_.rsyncArgument)" -Link "$($_.DstParentPath)/$($AllGen | Select-Object -Last 1)$($_.DstChildPath)" -Src "$($_.SrcPath)" -Dst "$($_.DstParentPath)$($Settings.DateTime)$($_.DstChildPath)" -Begin $_.Begin -End $_.End
    }
}

#バックアップ終了時刻
$End = "$((Get-Date).ToString("yyyy-MM-dd (ddd) HH:mm:ss"))"

#終了コード配列を参照して、エラーがあった数を集計
$ErrorCount = ($ExitCode | Where-Object {$_ -ne 0}).Count
"INFO ErrorCount: $ErrorCount"

"#--------------------後処理--------------------"
&$Settings.EndScript


#ログ取り停止
Stop-Transcript
