set CREDENTIAL_FILE (Join-Path . "credential.json")
set SECRET_FILE (Join-Path . "client_id.json")
set DATE_FORMAT "yyyy/MM/dd HH:mm:ss"
set GMAIL_SCOPE "https://www.googleapis.com/auth/gmail.send"



function Save-GoogleCredential($credential){
    # 有効期限の確認をするため、作成日をセットしておく(上書きする場合に備え、-Forceとする)
    $credential | Add-Member created_at (Get-Date).ToString($DATE_FORMAT) -Force

    # 扱いやすいJSON形式でトークンを保存しておく
    $credential_file = Join-Path . $CREDENTIAL_FILE
    $credential | ConvertTo-Json | Out-File $CREDENTIAL_FILE -Encoding utf8
}

function Get-GoogleCredential(){

    # https://gallery.technet.microsoft.com/scriptcenter/9e5733bc-8d72-41d7-a6c2-2936cb2cb7fc
    if (-not(Test-Path $SECRET_FILE)) {
        Write-Host "Not found client_id.json file"
        return $null
    }

    # client_id.jsonファイルから、認証情報を取得する
    $json = Get-Content $SECRET_FILE -Encoding UTF8 -Raw | ConvertFrom-Json
    $auth = $json.installed

    
    if (Test-Path $CREDENTIAL_FILE) {
        $current_credential = Get-Content $CREDENTIAL_FILE -Encoding UTF8 -Raw | ConvertFrom-Json
        if (-not ($current_credential.access_token -and $current_credential.token_type -and $current_credential.expires_in `
                  -and $current_credential.refresh_token -and $current_credential.created_at))
        {
            # アクセストークンファイルが正しい形式になっていない場合、処理できない
            Write-Host "No credential file: $($CREDENTIAL_FILE)"
            return $null
        }

        $elapsed_seconds = ((Get-Date) - [DateTime]::ParseExact($current_credential.created_at, $DATE_FORMAT, $null)).TotalSeconds
        if ($elapsed_seconds -lt $current_credential.expires_in ) {
            # expireしてない場合は、アクセストークンを再利用
            Write-Host "Reuse access token..."
            return $current_credential
        }
        else{
            # expireしてたら、リフレッシュする
            Write-Host "Refresh access token..."

            $refresh_body = @{
                "refresh_token" = $current_credential.refresh_token;
                "client_id" = $auth.client_id;
                "client_secret" = $auth.client_secret;
                "grant_type" = "refresh_token";
            }

            try {
                $refreshed_credential = Invoke-RestMethod -Method Post -Uri $auth.token_uri -Body $refresh_body
            }
            catch [System.Exception] {
                Write-Host $Error
                return $null
            }
            
            Save-GoogleCredential $refreshed_credential
            return $refreshed_credential
        }
    }

    # 上記以外の場合は、初回実行と考える
    Write-Host "New access token..."

    $gmail_scope = "https://www.googleapis.com/auth/gmail.send"

    $auth_url = "$($auth.auth_uri)?scope=$($GMAIL_SCOPE)"
    $auth_url += "&redirect_uri=$($auth.redirect_uris[0])"
    $auth_url += "&client_id=$($auth.client_id)"
    $auth_url += "&response_type=code&approval_prompt=force&access_type=offline"

    # IEを起動して、タイトルにcodeが入ったときにcodeを取得して閉じる
    # http://qiita.com/fujimohige/items/5aafe5604a943f74f6f0
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Navigate($auth_url)
    $ie.Visible = $true 

    $code = ""
    while($true){
        # http://canal22.org/advance/ie/ie-documen/
        # http://d.hatena.ne.jp/ritou/20110414/1302711014
        $title = $ie.Document.title
        if (($title) -and ($title.Contains("Success"))) {
            $code = $title.Replace("Success code=", "")
            break
        }
        Start-Sleep -s 1
    }

    # https://technet.microsoft.com/en-us/library/ff730962.aspx
    $ie.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie)
    Remove-Variable ie

    try {
        $new_body = @{
            "client_id" = $auth.client_id;
            "client_secret" = $auth.client_secret;
            "redirect_uri" = $auth.redirect_uris[0];
            "grant_type" = "authorization_code";
            "code" = $code;
        }
        # https://technet.microsoft.com/ja-jp/library/hh849971.aspx
        $new_credential = Invoke-RestMethod -Method Post -Uri $auth.token_uri -Body $new_body
    }
    catch [System.Exception] {
        Write-Host $Error
    }
    
    Save-GoogleCredential $new_credential
    return $new_credential
}

Export-ModuleMember –Function Get-GoogleCredential