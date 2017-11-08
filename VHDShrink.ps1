
################################################
# シュリンク実行
################################################
function ShrinkVHD( $VHD ){
    # VHD 存在確認
    if( -not (Test-Path $VHD)){
        echo "$VHD not found."
        return
    }

    # VHD を RW マウントしてディスク情報を得る
    $DiskInfo = Mount-VHD -Path $VHD -PassThru | Get-Disk

    # マウントしたボリューム情報
    $VolumeInfo = $DiskInfo | Get-Partition | Get-Volume

    # ドライブレターがあるボリュームを defrag
    [array]$DriveLetters = $VolumeInfo | ? DriveLetter -ne $null | % { $_.DriveLetter }
    foreach( $DriveLetter in $DriveLetters ){
        $DriveName = $DriveLetter + ":"
        defrag $DriveName /d /u /v
    }

    # VHD ディスマウント
    Dismount-VHD -DiskNumber $DiskInfo.Number

    # VHD サイズ切りつめ
    Mount-VHD -Path $VHD -ReadOnly -Passthru | Optimize-VHD -Mode Full -Passthru | Dismount-VHD
}

################################################
# メール送信
################################################
function SendMail(
        $MSA,               # メールサーバー
        $MailFrom,          # 送信元
        $RcpTos,            # 宛先
        $Subject,           # タイトル
        $Body,              # 本文
        $AuthUser,          # SMTP Auth アカウント
        $Password           # SMTP Auth パスワード
    ){

    #  .NET Framework 4.0 以降の場合
    if($PSVersionTable.CLRVersion.Major -ge 4){
        # FQDN を求める
        $Data = nslookup $env:computername | Select-String "名前:"
        if( $Data -ne $null ){
            $Part = -split $Data
            $FQDN = $Part[1]
        }

        # FQDN が DNS に登録されているホストの場合
        if( $FQDN -ne $null ){
            # app.config を作成する
            if( $PSVersionTable.PSVersion.Major -ge 3 ){
                $ScriptDir = $PSScriptRoot
            }
            # for PS v2
            else{
                $ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
            }
            $AppConfig = Join-Path $ScriptDir "app.config"
@"
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
    <system.net>
        <mailSettings>
            <smtp>
                <network
                    clientDomain="$FQDN"
                />
            </smtp>
        </mailSettings>
    </system.net>
</configuration>
"@ | set-content $AppConfig -Encoding UTF8

            # HELO 申告ホスト名に FQDN をセットする
            [AppDomain]::CurrentDomain.SetData("APP_CONFIG_FILE", $AppConfig)
            Add-Type -AssemblyName System.Configuration
        }
    }

    # メールデーター
    $Mail = New-Object Net.Mail.MailMessage

    # 送信元
    $Mail.From = $MailFrom
    echo "[INFO] From: $MailFrom"

    # 宛先
    foreach( $RcpTo in $RcpTos ){
        $Mail.To.Add($RcpTo)
        echo "[INFO] To: $RcpTo"
    }

    # タイトル
    $Mail.Subject = $Subject
    echo "[INFO] Subject: $Subject"

    # 本文
    $Mail.Body = $Body
    echo "[INFO] Body: $Body"

    # メール作成
    echo "[INFO] MSA: $MSA"

    # Submission
    if( $AuthUser -ne $null ){
        if( $Password -ne $null ){
            echo "[INFO] SMTP Auth Account : $AuthUser"
            $SmtpClient = New-Object Net.Mail.SmtpClient($MSA, 587)
            $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($AuthUser, $Password)
        }
        else{
            echo "[ERROR] SMTP Auth パスワードがセットされていない"
            return
        }
    }
    # SMTP
    else{
        $SmtpClient = New-Object Net.Mail.SmtpClient($MSA)
    }

    # メール送信
    echo "[INFO] Mail send"
    try{
        $SmtpClient.Send($Mail)
    }
    catch{
        $Now = Get-Date
        $ExecTime = "{0:0000}-{1:00}-{2:00} " -f $Now.Year, $Now.Month, $Now.Day
        $ExecTime += "{0:00}:{1:00}:{2:00}.{3:000} " -f $Now.Hour, $Now.Minute, $Now.Second, $Now.Millisecond
        echo "[ERROR] $ExecTime メール送信に失敗しました"
    }

    $Mail.Dispose()

    return
}

################################################
# main
################################################
$StartTime = (Get-Date).DateTime
[int]$BeforeFreeDiskSize = ((Get-PSDrive d).Free)/1GB

# メール送信
$MSA = "172.24.2.61"
$MailFrom = "VhdShrink@win.monitor.clayapp.jp"
$RcpTos = @("s.murashima@gloops.com", "j.matsumoto@gloops.com")
$Subject = "【dev-rogue-ap】D03428-VMA VHD Shrink start"
$Body = @"
dev-rogue-ap / D03428-VMA の VHD Shrink を開始しました
Start : $StartTime
Shrink 前 Free Disk Size : $BeforeFreeDiskSize
"@
SendMail $MSA $MailFrom $RcpTos $Subject $Body

# 圧縮
$VM = Get-VM
Stop-VM -VMName $VM.Name -Force

[array]$VHDs = dir D:\Disk1\*.vhdx
foreach( $VHD in $VHDs ){
    ShrinkVHD $VHD.FullName
}

Start-VM -VMName $VM.Name

$EndTime = (Get-Date).DateTime
[int]$AfterFreeDiskSize = ((Get-PSDrive d).Free)/1GB

# メール送信
$Subject = "【dev-rogue-ap】D03428-VMA VHD Shrink end"
$Body = @"
dev-rogue-ap / D03428-VMA の VHD Shrink が終了しました
Start : $StartTime
End   : $EndTime
Shrink 前 Free Disk Size : $BeforeFreeDiskSize
Shrink 後 Free Disk Size : $AfterFreeDiskSize
"@
SendMail $MSA $MailFrom $RcpTos $Subject $Body


