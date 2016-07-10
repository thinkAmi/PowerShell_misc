$mail = @{
    from = "example+from@gmail.com";
    to = "example+to@gmail.com";
    smtp_server = "smtp.gmail.com";
    smtp_port = 587;
    user = "example+to@gmail.com";
    password = "1234";
}


function Send-Gmail($mail, $msg){
    $password = ConvertTo-SecureString $mail["password"] -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential $mail["user"], $password
    $host_name = [Net.Dns]::GetHostName()
    Send-MailMessage -To $mail["to"] `
                     -From $mail["from"] `
                     -SmtpServer $mail["smtp_server"] `
                     -Credential $credential `
                     -Port $mail["smtp_port"] `
                     -Subject "$host_name ログ" `
                     -Body $msg `
                     -Encoding UTF8 `
                     -UseSsl
}


# オブジェクトのプロパティを示すため、カッコでくくる
# 念のため、呼ばれた前30秒のログを取得する
$start_date = (Get-Date).AddSeconds(-30)
$end_date = Get-Date

# 除外するイベントログ情報
$exclusions = @(
    @{
        Id = 9999;
        LogName = "Application";
    }
)

# ログ名は、引数にカンマ区切りでセットする
# 例) System,Application
$log_name = $args[0] -split ","

# FilterHashTable用のフィルタ
$filter = @{
    LogName = $log_name;
    Level = 1,2,3;
    StartTime = $start_date;
    EndTime = $end_date;
}

# FilterHashTableでフィルタした後のイベントを取得
$events = Get-WinEvent -FilterHashTable $filter

# 除外指定されているイベント情報は送信しない
foreach($e in $exclusions){
    $events = $events | Where-Object { -not (
        $_.Id -eq $e.Id -and
        $_.LogName -eq $e.LogName
    )}
}


$msg = @"
対象イベント件数: $($events.Count)
--------------------------------------------------


"@

foreach($event in $events){
    $log_task_category = if ($event.Task){ "$($event.TaskDisplayName) ($($event.Task))" } else { "なし" }
    $log_edited_user_id = if ($event.UserId){ "($($event.UserId))" } else { "" }
    
    $log_user_name = if ($event.UserId) {
        # http://yomon.hatenablog.com/entry/2015/06/19/183522
        (New-Object System.Security.Principal.SecurityIdentifier($event.UserId)).Translate([System.Security.Principal.NTAccount]).Value
    } else { "" }

    $log_property = ""
    foreach ($p in $event.properties){
        $log_property += $p.value
    }

    $msg += @"
日時: $($event.TimeCreated.ToString("yyyy/MM/dd HH:mm:ss"))
レベル: $($event.LevelDisplayName) ($($event.Level))
ログの名前: $($event.LogName)
ソース: $($event.ProviderName) 
タスクのカテゴリ: $log_task_category
イベントID: $($event.Id)
キーワード: $($event.KeywordsDisplayNames)
ユーザー: $($log_user_name) $log_edited_user_id
コンピューター: $($event.MachineName)
オペコード: $($event.OpcodeDisplayName) ($($event.Opcode))
プロパティ:
$log_property


メッセージ:
$($event.Message)

--------------------------------------------------


"@
}

Send-Gmail $mail $msg
