$mail = @{
    from = "example@gmail.com";
    to = "example+to@gmail.com";
    smtp_server = "smtp.gmail.com";
    smtp_port = 587;
    user = "example@gmail.com";
    password = "1234";
}

$client = New-Object Net.Mail.SmtpClient($mail["smtp_server"], $mail["smtp_port"])

# GmailはSMTP + SSLで送信する
$client.EnableSsl = $true

# SMTP Authのため、認証情報を設定する
$client.Credentials = New-Object Net.NetworkCredential($mail["user"], $mail["password"])

$msg = New-Object Net.Mail.MailMessage($mail["from"], $mail["to"], "subject", "body: Hello Net class")

$client.Send($msg)