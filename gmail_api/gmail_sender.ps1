# 参考 https://www.administrator.de/wissen/powershell-googlemail-gmail-nativ-powershell-verwalten-291531.html

function ConvertTo-Base64Url($str){
    # 参考 http://jason.pettys.name/2014/10/27/sending-email-with-the-gmail-api-in-net-c/

    # ToBase64String()で使うため、バイト列にエンコード
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($str)
    # バイトをBase64でデコードした文字列にする
    $b64str = [System.Convert]::ToBase64String($bytes)
    # ToBase64String()ではURLセーフなBase64ではないため、不足している部分を自分で置き換え
    $without_plus = $b64str -replace '\+', '-'
    $without_slash = $without_plus -replace '/', '_'
    $without_equal = $without_slash -replace '=', ''

    return $without_equal
}

function Send-Gmail(){
    # メールの送受信情報
    $mail = @{
        from = "example@gmail.com";
        to = "example+to@gmail.com";
    }

    # 別モジュールからGoogleの認証情報を取得
    # https://msdn.microsoft.com/en-us/library/dd878284.aspx
    $path = Join-Path . "google_credential.psm1"
    Import-Module -Name $path
    $credential = Get-GoogleCredential

    # 念のため、認証情報を表示
    Write-Host $credential

    if (-not $credential){
        # 認証情報がない場合は、処理しない
        Write-Host "Not Authenticated."
        return
    }

    # 外部ライブラリ`AE.NET.Mail`のdllを読み込んで、使用する
    # 事前準備として、AE.NET.Mailライブラリは、.NETのバージョンが一致するものを選び、ブロックの解除を行う
    # ブロックの解除 http://stackoverflow.com/questions/18801440/powershell-load-dll-got-error-add-type-could-not-load-file-or-assembly-webdr
    # PS S:\Sandbox\gmai_api> $PSVersionTable
    # Name                           Value
    # ----                           -----
    # PSVersion                      5.0.10586.122
    # CLRVersion                     4.0.30319.42000
    $dll = Join-Path . "AE.NET.Mail.dll"
    Add-Type -Path $dll

    $msg = New-Object AE.Net.Mail.MailMessage

    # Fromはプロパティにセット
    $from = New-Object System.Net.Mail.MailAddress $mail["from"]
    $msg.From = $from
    # $msg.From.Add($from) #=> null 値の式ではメソッドを呼び出せません。 

    # ToはAdd()を使う
    $to = New-Object System.Net.Mail.MailAddress $mail["to"]
    $msg.To.Add($to)
    # $msg.To = $to #=> "To" は ReadOnly のプロパティです。

    # ReplyToがないと正常に送信されたにも関わらず、Bounceメールが届く
    # http://stackoverflow.com/questions/28122074/gmail-api-emails-bouncing
    # ReplyToもAddを使う
    $msg.ReplyTo.Add($from)
    # $msg.ReplyTo = $from  #=> "ReplyTo" は ReadOnly のプロパティです。 

    $msg.Subject = "gmail api subject"
    $msg.Body = "body: ハロー Gmail API!"

    # https://msdn.microsoft.com/ja-jp/library/system.io.stringwriter.aspx
    $sw = New-Object System.IO.StringWriter
    $msg.Save($sw)
    $raw = ConvertTo-Base64Url $sw.ToString()

    # メールボディはJSONにする
    $body = @{ "raw" = $raw; } | ConvertTo-Json

    # ユーザIDはURL encodeした値、もしくは`me`で認証したユーザとなる
    # https://developers.google.com/gmail/api/v1/reference/users/messages/send#http-request
    # $user_id = [System.Net.WebUtility]::UrlEncode($mail["from"])
    $user_id = "me"
    $uri = "https://www.googleapis.com/gmail/v1/users/$($user_id)/messages/send?access_token=$($credential.access_token)"

    try {
        $result = Invoke-RestMethod $uri -Method POST -ErrorAction Stop -Body $body -ContentType "application/json"
    }
    catch [System.Exception] {
        Write-Host $Error
        return
    }
    Write-Host $result
}


# エントリポイント
Send-Gmail