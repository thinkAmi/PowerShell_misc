$mail = @{
    from = "example@gmail.com";
    to = "example+to@gmail.com";
    smtp_server = "smtp.gmail.com";
    smtp_port = 587;
    user = "example@gmail.com";
    password = "1234";
}

$password = ConvertTo-SecureString $mail["password"] -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential $mail["user"], $password

Send-MailMessage -To $mail["to"] `
                 -From $mail["from"] `
                 -SmtpServer $mail["smtp_server"] `
                 -Credential $credential `
                 -Port $mail["smtp_port"] `
                 -Subject "subject" `
                 -Body "body: Hello Cmdlet" `
                 -Encoding UTF8 `
                 -UseSsl